/**
 * SharePoint site provisioning tools.
 *
 * Wraps the three Graph endpoints that create SharePoint sites:
 *   • `POST /beta/sites`  (template `sitepagepublishing` → communication site,
 *                          template `sts` → modern team site without M365 group)
 *   • `POST /v1.0/groups` (creates a Microsoft 365 group, which provisions a
 *                          group-connected SharePoint team site asynchronously)
 *   • `GET  /beta/sites/getOperationStatus(operationId='…')` — polls long-
 *                          running site creation operations from `POST /sites`.
 *
 * Both `POST /sites` flavors return 202 Accepted with a `Location` header that
 * points at `getOperationStatus`; we poll there until completion when the
 * caller opts in with `waitForCompletion`.
 *
 * `POST /groups` returns 201 synchronously with the group id, but the
 * SharePoint site behind it provisions in the background; we poll
 * `/groups/{id}/sites/root` until it stops 404'ing.
 *
 * Note on API stability: `POST /sites` is currently beta-only. Microsoft has
 * not committed a v1.0 timeline as of 2026-05. The fork pins beta usage to
 * these tools only, and v1.0 stays the default for every other request.
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getGraphClient, GraphClient } from "../../graph/client.js";
import { Site } from "../../graph/models.js";
import { jsonTextResponse, toolErrorResponse } from "../../graph/contracts.js";
import { getAuthInstance } from "../../auth/microsoft-graph-auth.js";
import {
  pollUntil,
  PollUntilOptions,
  PollingTimeoutError,
} from "../../graph/polling.js";
import { resolveTenantSharePointHostname } from "../../sharepoint/site-resolver.js";

// SharePoint URL leaves: lowercase letters, digits, hyphens. Must start with a
// letter or digit. No length policy enforced — Graph rejects oversize values
// with a clear error so let Graph be the source of truth there.
const SITE_ALIAS_PATTERN = /^[a-z0-9][a-z0-9-]*$/;

// M365 group mail nicknames: letters, digits, dots, hyphens, underscores.
// Cannot start or end with a period (Graph rejects this with a confusing
// error, so we validate up-front).
const MAIL_NICKNAME_PATTERN = /^[A-Za-z0-9_-][A-Za-z0-9._-]*[A-Za-z0-9_-]$|^[A-Za-z0-9_-]$/;

// Test seam: the polling loop's sleep is configurable so unit tests can fast-
// forward through retries. Production code never sets this; tests call
// `__setPollingSleepForTests(async () => {})` to make polling instantaneous.
let pollingSleep: ((ms: number) => Promise<void>) | undefined;

export function __setPollingSleepForTests(
  fn: ((ms: number) => Promise<void>) | undefined,
): void {
  pollingSleep = fn;
}

interface SiteCreationCommonOptions {
  intervalSeconds?: number;
  timeoutSeconds?: number;
}

interface CreateSiteArgs {
  displayName?: unknown;
  alias?: unknown;
  webUrl?: unknown;
  tenantHostname?: unknown;
  description?: unknown;
  locale?: unknown;
  ownerEmail?: unknown;
  shareByEmailEnabled?: unknown;
  waitForCompletion?: unknown;
  timeoutSeconds?: unknown;
  intervalSeconds?: unknown;
}

interface CreateTeamSiteArgs {
  displayName?: unknown;
  mailNickname?: unknown;
  description?: unknown;
  visibility?: unknown;
  ownerEmails?: unknown;
  memberEmails?: unknown;
  hideFromOutlook?: unknown;
  welcomeEmailDisabled?: unknown;
  provisionSiteOnDemand?: unknown;
  waitForSite?: unknown;
  timeoutSeconds?: unknown;
  intervalSeconds?: unknown;
}

interface GetSiteCreationStatusArgs {
  operationId?: unknown;
}

function asString(v: unknown, label: string): string {
  if (typeof v !== "string") {
    throw new Error(`${label} must be a string`);
  }
  return v;
}

function asOptionalString(v: unknown, label: string): string | undefined {
  if (v === undefined || v === null) return undefined;
  if (typeof v !== "string") {
    throw new Error(`${label} must be a string`);
  }
  return v;
}

function asOptionalBoolean(v: unknown, label: string): boolean | undefined {
  if (v === undefined || v === null) return undefined;
  if (typeof v !== "boolean") {
    throw new Error(`${label} must be a boolean`);
  }
  return v;
}

function asOptionalNumber(v: unknown, label: string): number | undefined {
  if (v === undefined || v === null) return undefined;
  if (typeof v !== "number") {
    throw new Error(`${label} must be a number`);
  }
  return v;
}

function asStringArray(v: unknown, label: string): string[] {
  if (v === undefined || v === null) return [];
  if (!Array.isArray(v) || v.some((item) => typeof item !== "string")) {
    throw new Error(`${label} must be an array of strings`);
  }
  return v as string[];
}

interface OperationStatusResponse {
  id: string;
  createdDateTime?: string;
  resourceId?: string;
  resourceLocation?: string;
  status: "notStarted" | "running" | "succeeded" | "failed" | string;
  percentageComplete?: number;
  error?: { code?: string; message?: string };
}

/**
 * Parse the operationId from a `Location` header returned by `POST /beta/sites`.
 * Graph emits something like:
 *   `https://graph.microsoft.com/beta/sites/getOperationStatus(operationId='JXMnaHR0cHMlM0…')`
 */
function extractOperationIdFromLocation(location: string | undefined): string | null {
  if (!location) return null;
  const match = location.match(/operationId='([^']+)'/);
  return match?.[1] ?? null;
}

function toPollOptions(common: SiteCreationCommonOptions, defaultTimeoutSeconds: number): PollUntilOptions {
  return {
    intervalMs: (common.intervalSeconds ?? 3) * 1000,
    timeoutMs: (common.timeoutSeconds ?? defaultTimeoutSeconds) * 1000,
    sleep: pollingSleep,
  };
}

async function fetchOperationStatus(
  client: GraphClient,
  operationId: string,
): Promise<OperationStatusResponse> {
  // Single-quoted operation IDs travel as-is in the URL path; Graph rejects
  // percent-encoded quotes, so we splice the id in directly and trust that
  // Graph-issued ids never contain a literal single quote.
  if (operationId.includes("'")) {
    throw new Error("Invalid operationId — must not contain a single quote.");
  }
  const endpoint = `/sites/getOperationStatus(operationId='${operationId}')`;
  const response = await client.get<OperationStatusResponse>(endpoint, undefined, {
    apiVersion: "beta",
  });
  if (!response.success || !response.data) {
    throw new Error("Failed to fetch site creation operation status");
  }
  return response.data;
}

/**
 * Common flow shared by communication-site and classic-team-site creation:
 * both call `POST /beta/sites` and differ only in the `template` field.
 */
async function createSiteViaPostSites(args: {
  displayName: string;
  alias?: string;
  webUrl?: string;
  description?: string;
  locale?: string;
  ownerEmail?: string;
  shareByEmailEnabled?: boolean;
  template: "sitepagepublishing" | "sts";
  tenantHostname?: string;
  waitForCompletion?: boolean;
  intervalSeconds?: number;
  timeoutSeconds?: number;
}) {
  const client = getGraphClient();

  // 1. Build target webUrl. Explicit `webUrl` always wins; otherwise compose
  //    from the tenant host + alias. We require one of the two.
  let webUrl = args.webUrl;
  let alias = args.alias;
  if (!webUrl) {
    if (!alias) {
      throw new Error("Either `alias` or `webUrl` is required.");
    }
    if (!SITE_ALIAS_PATTERN.test(alias)) {
      throw new Error(
        `Invalid site alias "${alias}". Must match ${SITE_ALIAS_PATTERN.source} (lowercase letters, digits, hyphens; must start with letter or digit).`,
      );
    }
    const host = args.tenantHostname ?? (await resolveTenantSharePointHostname(client));
    webUrl = `https://${host}/sites/${alias}`;
  } else if (!alias) {
    try {
      // Derive alias for response payload only — Graph treats webUrl as truth.
      const parsed = new URL(webUrl);
      alias = parsed.pathname.replace(/^\/sites\//, "").replace(/\/$/, "");
    } catch {
      // Ignore — caller-supplied webUrl will be validated by Graph.
    }
  }

  // 2. App-only callers must specify an owner. Without it, Graph either fails
  //    or creates an orphaned site. Detect the auth mode and refuse early.
  const authMode = getAuthInstance().getAuthMode();
  if (authMode === "client_credentials" && !args.ownerEmail) {
    throw new Error(
      "ownerEmail is required when running in client-credentials (app-only) mode. Pass the UPN of the user who should own the new site.",
    );
  }

  // 3. Build the POST body. Keys are exactly what `POST /beta/sites` accepts;
  //    extra keys are silently ignored by Graph, so don't smuggle anything in.
  const body: Record<string, unknown> = {
    displayName: args.displayName,
    name: args.displayName,
    webUrl,
    template: args.template,
  };
  if (args.description !== undefined) body.description = args.description;
  if (args.locale !== undefined) body.locale = args.locale;
  if (args.shareByEmailEnabled !== undefined) {
    body.shareByEmailEnabled = args.shareByEmailEnabled;
  }
  if (args.ownerEmail) {
    body.ownerIdentityToResolve = { email: args.ownerEmail };
  }

  // 4. Submit. 202 with a Location header is the expected happy path; a 200/201
  //    with a populated site object can also happen on very small tenants.
  const submitResponse = await client.post<Site & Partial<OperationStatusResponse>>(
    "/sites",
    body,
    { apiVersion: "beta" },
  );

  const status = submitResponse.metadata?.status;
  const locationHeader = submitResponse.metadata?.headers?.["location"];
  const operationId = extractOperationIdFromLocation(locationHeader);

  // 5a. Synchronous completion (rare but possible). Return immediately.
  if (status && status !== 202 && submitResponse.data?.id) {
    return jsonTextResponse({
      success: true,
      operation: {
        status: "succeeded",
        operationId,
        synchronous: true,
      },
      site: {
        id: submitResponse.data.id,
        webUrl: submitResponse.data.webUrl ?? webUrl,
        displayName: submitResponse.data.displayName ?? args.displayName,
        description: submitResponse.data.description,
      },
    });
  }

  // 5b. Async creation, caller opted out of waiting — return what we know.
  if (args.waitForCompletion === false) {
    return jsonTextResponse({
      success: true,
      pending: true,
      message:
        "Site creation submitted. Use get_site_creation_status with the returned operationId to track progress.",
      operation: { status: "running", operationId, synchronous: false },
      expectedWebUrl: webUrl,
    });
  }

  // 5c. Poll until succeeded / failed / timeout.
  if (!operationId) {
    throw new Error(
      "POST /beta/sites returned no operationId in the Location header and no site body. Cannot determine completion.",
    );
  }

  const pollOptions = toPollOptions(args, 180);
  let lastStatus: OperationStatusResponse | undefined;
  try {
    const result = await pollUntil<OperationStatusResponse>(
      "site creation",
      async () => {
        const op = await fetchOperationStatus(client, operationId);
        lastStatus = op;
        if (op.status === "succeeded") return { done: true, value: op };
        if (op.status === "failed") {
          throw new Error(
            `Site creation failed: ${op.error?.message ?? "unknown error"}`,
          );
        }
        return { done: false, value: op };
      },
      pollOptions,
    );

    // Try to fetch the resulting site for richer response data. Graph caches
    // can lag here, so a 404 isn't fatal — we just return what the operation
    // gave us.
    let site: Site | undefined;
    if (result.value.resourceLocation) {
      try {
        const lookup = await client.get<Site>(result.value.resourceLocation);
        if (lookup.success && lookup.data) site = lookup.data;
      } catch {
        // Ignore — fall back to the operation payload.
      }
    }

    return jsonTextResponse({
      success: true,
      operation: {
        status: "succeeded",
        operationId,
        attempts: result.attempts,
        elapsedMs: result.elapsedMs,
      },
      site: {
        id: site?.id ?? result.value.resourceId,
        webUrl: site?.webUrl ?? webUrl,
        displayName: site?.displayName ?? args.displayName,
        description: site?.description ?? args.description,
      },
    });
  } catch (error) {
    if (error instanceof PollingTimeoutError) {
      return jsonTextResponse({
        success: false,
        pending: true,
        message: `Polling timed out after ${error.elapsedMs}ms; site may still be provisioning. Use get_site_creation_status with the operationId to check.`,
        operation: {
          status: lastStatus?.status ?? "running",
          operationId,
          attempts: error.attempts,
          elapsedMs: error.elapsedMs,
        },
        expectedWebUrl: webUrl,
      });
    }
    throw error;
  }
}

// Tool 1: Create a communication site
export const createCommunicationSite: Tool = {
  name: "create_communication_site",
  description:
    "Create a new SharePoint Communication site via POST /beta/sites (template=sitepagepublishing). Returns the site id and webUrl once provisioning completes. Beta API; not yet on v1.0.",
  inputSchema: {
    type: "object",
    properties: {
      displayName: {
        type: "string",
        description: "Display name of the new site",
      },
      alias: {
        type: "string",
        description:
          "URL leaf for the site (lowercase letters/digits/hyphens). Combined with the tenant SharePoint host to form the site URL, e.g. alias=marketing-2026 → https://{tenant}.sharepoint.com/sites/marketing-2026",
      },
      webUrl: {
        type: "string",
        description:
          "Full target webUrl. Overrides `alias` + tenant resolution. Use only when you need a non-default hostname or path.",
      },
      tenantHostname: {
        type: "string",
        description:
          "Override the auto-resolved tenant SharePoint hostname (e.g. 'contoso.sharepoint.com'). Rarely needed.",
      },
      description: {
        type: "string",
        description: "Optional site description",
      },
      locale: {
        type: "string",
        description:
          "BCP-47 locale tag (e.g. 'es-MX', 'en-US'). Defaults to es-MX for this fork.",
        default: "es-MX",
      },
      ownerEmail: {
        type: "string",
        description:
          "UPN of the site owner. Optional in delegated mode (defaults to caller); REQUIRED in client-credentials mode.",
      },
      shareByEmailEnabled: {
        type: "boolean",
        description: "Allow guests to be invited by email",
        default: false,
      },
      waitForCompletion: {
        type: "boolean",
        description:
          "If true, poll the long-running operation until it succeeds or times out. If false, return the operationId immediately for asynchronous polling.",
        default: true,
      },
      timeoutSeconds: {
        type: "number",
        description: "Total polling budget when waitForCompletion=true.",
        default: 180,
      },
      intervalSeconds: {
        type: "number",
        description: "Base polling interval (subject to backoff).",
        default: 3,
      },
    },
    required: ["displayName"],
  },
};

export async function handleCreateCommunicationSite(args: CreateSiteArgs) {
  try {
    return await createSiteViaPostSites({
      displayName: asString(args.displayName, "displayName"),
      alias: asOptionalString(args.alias, "alias"),
      webUrl: asOptionalString(args.webUrl, "webUrl"),
      description: asOptionalString(args.description, "description"),
      locale: asOptionalString(args.locale, "locale") ?? "es-MX",
      ownerEmail: asOptionalString(args.ownerEmail, "ownerEmail"),
      shareByEmailEnabled: asOptionalBoolean(
        args.shareByEmailEnabled,
        "shareByEmailEnabled",
      ),
      template: "sitepagepublishing",
      tenantHostname: asOptionalString(args.tenantHostname, "tenantHostname"),
      waitForCompletion: asOptionalBoolean(
        args.waitForCompletion,
        "waitForCompletion",
      ),
      intervalSeconds: asOptionalNumber(args.intervalSeconds, "intervalSeconds"),
      timeoutSeconds: asOptionalNumber(args.timeoutSeconds, "timeoutSeconds"),
    });
  } catch (error) {
    return toolErrorResponse("create_communication_site", error);
  }
}

// Tool 2: Create a modern team site WITHOUT a backing M365 group
export const createTeamSiteClassic: Tool = {
  name: "create_team_site_classic",
  description:
    "Create a SharePoint modern team site WITHOUT a backing Microsoft 365 group via POST /beta/sites (template=sts). Use create_team_site instead when you want a Teams-/group-connected site. Beta API; not yet on v1.0.",
  inputSchema: {
    type: "object",
    properties: {
      displayName: { type: "string", description: "Display name of the new site" },
      alias: {
        type: "string",
        description:
          "URL leaf for the site (lowercase letters/digits/hyphens). Combined with the tenant SharePoint host.",
      },
      webUrl: {
        type: "string",
        description:
          "Full target webUrl. Overrides `alias` + tenant resolution.",
      },
      tenantHostname: {
        type: "string",
        description:
          "Override the auto-resolved tenant SharePoint hostname.",
      },
      description: { type: "string", description: "Optional site description" },
      locale: {
        type: "string",
        description: "BCP-47 locale tag. Defaults to es-MX.",
        default: "es-MX",
      },
      ownerEmail: {
        type: "string",
        description:
          "UPN of the site owner. Required in client-credentials mode.",
      },
      shareByEmailEnabled: {
        type: "boolean",
        description: "Allow guests to be invited by email",
        default: false,
      },
      waitForCompletion: { type: "boolean", default: true },
      timeoutSeconds: { type: "number", default: 180 },
      intervalSeconds: { type: "number", default: 3 },
    },
    required: ["displayName"],
  },
};

export async function handleCreateTeamSiteClassic(args: CreateSiteArgs) {
  try {
    return await createSiteViaPostSites({
      displayName: asString(args.displayName, "displayName"),
      alias: asOptionalString(args.alias, "alias"),
      webUrl: asOptionalString(args.webUrl, "webUrl"),
      description: asOptionalString(args.description, "description"),
      locale: asOptionalString(args.locale, "locale") ?? "es-MX",
      ownerEmail: asOptionalString(args.ownerEmail, "ownerEmail"),
      shareByEmailEnabled: asOptionalBoolean(
        args.shareByEmailEnabled,
        "shareByEmailEnabled",
      ),
      template: "sts",
      tenantHostname: asOptionalString(args.tenantHostname, "tenantHostname"),
      waitForCompletion: asOptionalBoolean(
        args.waitForCompletion,
        "waitForCompletion",
      ),
      intervalSeconds: asOptionalNumber(args.intervalSeconds, "intervalSeconds"),
      timeoutSeconds: asOptionalNumber(args.timeoutSeconds, "timeoutSeconds"),
    });
  } catch (error) {
    return toolErrorResponse("create_team_site_classic", error);
  }
}

// Tool 3: Create a Microsoft-365-group-backed team site
export const createTeamSite: Tool = {
  name: "create_team_site",
  description:
    "Create a SharePoint team site backed by a Microsoft 365 group via POST /v1.0/groups (groupTypes=['Unified']). Group creation is synchronous; the SharePoint site provisions in the background and is polled via /groups/{id}/sites/root.",
  inputSchema: {
    type: "object",
    properties: {
      displayName: {
        type: "string",
        description: "Display name of the group and its team site",
      },
      mailNickname: {
        type: "string",
        description:
          "Email alias for the group (letters/digits/._-, length 1-64). Forms the group address mailNickname@tenant.com.",
      },
      description: { type: "string", description: "Optional group description" },
      visibility: {
        type: "string",
        enum: ["Private", "Public", "HiddenMembership"],
        description: "Group visibility. Defaults to Private.",
        default: "Private",
      },
      ownerEmails: {
        type: "array",
        items: { type: "string" },
        description:
          "UPNs of users who should own the group. Required in client-credentials mode; recommended in delegated mode (defaults to caller).",
      },
      memberEmails: {
        type: "array",
        items: { type: "string" },
        description: "Optional additional members (UPNs).",
      },
      hideFromOutlook: {
        type: "boolean",
        description: "Hide the group from Outlook (HideGroupInOutlook).",
        default: false,
      },
      welcomeEmailDisabled: {
        type: "boolean",
        description: "Suppress the welcome email sent to new members.",
        default: false,
      },
      provisionSiteOnDemand: {
        type: "boolean",
        description:
          "Defer SharePoint site provisioning until first access. When true, this tool skips the sites/root poll.",
        default: false,
      },
      waitForSite: {
        type: "boolean",
        description:
          "If true (default), poll until the group's SharePoint site is reachable. If false, return as soon as the group is created.",
        default: true,
      },
      timeoutSeconds: {
        type: "number",
        description: "Polling budget for sites/root readiness.",
        default: 300,
      },
      intervalSeconds: { type: "number", default: 3 },
    },
    required: ["displayName", "mailNickname"],
  },
};

async function resolveOwnerIds(
  client: GraphClient,
  emails: string[],
): Promise<string[]> {
  const ids: string[] = [];
  for (const email of emails) {
    // `GET /users/{upn-or-id}` returns the directory user; we just need its id
    // to populate `owners@odata.bind`.
    const response = await client.get<{ id: string }>(`/users/${encodeURIComponent(email)}`);
    if (!response.success || !response.data?.id) {
      throw new Error(`Failed to resolve owner email '${email}' to a user id`);
    }
    ids.push(response.data.id);
  }
  return ids;
}

export async function handleCreateTeamSite(args: CreateTeamSiteArgs) {
  try {
    const client = getGraphClient();

    const displayName = asString(args.displayName, "displayName");
    const mailNickname = asString(args.mailNickname, "mailNickname");
    if (mailNickname.length > 64) {
      throw new Error("mailNickname must be 64 characters or fewer");
    }
    if (mailNickname.startsWith(".") || mailNickname.endsWith(".")) {
      throw new Error("mailNickname cannot start or end with a period");
    }
    if (!MAIL_NICKNAME_PATTERN.test(mailNickname)) {
      throw new Error(
        `Invalid mailNickname "${mailNickname}". Allowed characters: letters, digits, dot, hyphen, underscore (not leading/trailing dot).`,
      );
    }

    const description = asOptionalString(args.description, "description");
    const visibility = asOptionalString(args.visibility, "visibility") ?? "Private";
    const ownerEmails = asStringArray(args.ownerEmails, "ownerEmails");
    const memberEmails = asStringArray(args.memberEmails, "memberEmails");
    const hideFromOutlook = asOptionalBoolean(args.hideFromOutlook, "hideFromOutlook");
    const welcomeEmailDisabled = asOptionalBoolean(
      args.welcomeEmailDisabled,
      "welcomeEmailDisabled",
    );
    const provisionSiteOnDemand = asOptionalBoolean(
      args.provisionSiteOnDemand,
      "provisionSiteOnDemand",
    );
    const waitForSite = asOptionalBoolean(args.waitForSite, "waitForSite");
    const intervalSeconds = asOptionalNumber(args.intervalSeconds, "intervalSeconds");
    const timeoutSeconds = asOptionalNumber(args.timeoutSeconds, "timeoutSeconds");

    const authMode = getAuthInstance().getAuthMode();
    if (authMode === "client_credentials" && ownerEmails.length === 0) {
      throw new Error(
        "ownerEmails is required when running in client-credentials (app-only) mode. Without an explicit owner, Graph creates an orphaned group whose SharePoint site never provisions until an admin intervenes.",
      );
    }

    const ownerIds = ownerEmails.length > 0
      ? await resolveOwnerIds(client, ownerEmails)
      : [];
    const memberIds = memberEmails.length > 0
      ? await resolveOwnerIds(client, memberEmails)
      : [];

    // Construct the group payload. `resourceProvisioningOptions` is
    // deliberately omitted — Graph's docs say "let the system manage" it; it
    // flips to ["Team"] automatically when a Team is later created via
    // `PUT /teams`. Setting it manually here would race the system.
    const resourceBehaviorOptions: string[] = [];
    if (welcomeEmailDisabled) resourceBehaviorOptions.push("WelcomeEmailDisabled");
    if (hideFromOutlook) resourceBehaviorOptions.push("HideGroupInOutlook");
    if (provisionSiteOnDemand) resourceBehaviorOptions.push("ProvisionSiteOnDemand");

    const body: Record<string, unknown> = {
      displayName,
      mailNickname,
      mailEnabled: true,
      securityEnabled: false,
      groupTypes: ["Unified"],
      visibility,
    };
    if (description !== undefined) body.description = description;
    if (ownerIds.length > 0) {
      body["owners@odata.bind"] = ownerIds.map(
        (id) => `https://graph.microsoft.com/v1.0/users/${id}`,
      );
    }
    if (memberIds.length > 0) {
      body["members@odata.bind"] = memberIds.map(
        (id) => `https://graph.microsoft.com/v1.0/users/${id}`,
      );
    }
    if (resourceBehaviorOptions.length > 0) {
      body.resourceBehaviorOptions = resourceBehaviorOptions;
    }

    const submitResponse = await client.post<{
      id: string;
      displayName?: string;
      mailNickname?: string;
      mail?: string;
    }>("/groups", body);

    if (!submitResponse.success || !submitResponse.data?.id) {
      throw new Error("Group creation returned no group id");
    }
    const groupId = submitResponse.data.id;

    // If the caller opted into deferred provisioning, the site won't exist
    // until first access — skip the poll and report just the group.
    if (provisionSiteOnDemand || waitForSite === false) {
      return jsonTextResponse({
        success: true,
        pending: !provisionSiteOnDemand,
        group: {
          id: groupId,
          displayName: submitResponse.data.displayName ?? displayName,
          mailNickname: submitResponse.data.mailNickname ?? mailNickname,
          mail: submitResponse.data.mail,
        },
        message: provisionSiteOnDemand
          ? "Group created. SharePoint site will be provisioned on first access (ProvisionSiteOnDemand)."
          : "Group created. SharePoint site is provisioning asynchronously; query /groups/{id}/sites/root later.",
      });
    }

    // Poll /groups/{id}/sites/root until it returns 200. Graph returns 404
    // while the site is still provisioning, which we treat as "not yet".
    const pollOptions: PollUntilOptions = {
      intervalMs: (intervalSeconds ?? 3) * 1000,
      timeoutMs: (timeoutSeconds ?? 300) * 1000,
      sleep: pollingSleep,
    };

    try {
      const result = await pollUntil<Site | null>(
        "group site provisioning",
        async () => {
          try {
            const siteResponse = await client.get<Site>(
              `/groups/${groupId}/sites/root`,
            );
            if (siteResponse.success && siteResponse.data?.id) {
              return { done: true, value: siteResponse.data };
            }
            return { done: false, value: null };
          } catch (err) {
            // 404 → still provisioning. Anything else propagates.
            const errorWithStatus = err as { statusCode?: number; status?: number };
            const statusCode = errorWithStatus.statusCode ?? errorWithStatus.status;
            if (statusCode === 404) {
              return { done: false, value: null };
            }
            throw err;
          }
        },
        pollOptions,
      );

      return jsonTextResponse({
        success: true,
        operation: {
          attempts: result.attempts,
          elapsedMs: result.elapsedMs,
        },
        group: {
          id: groupId,
          displayName: submitResponse.data.displayName ?? displayName,
          mailNickname,
          mail: submitResponse.data.mail,
        },
        site: result.value
          ? {
              id: result.value.id,
              webUrl: result.value.webUrl,
              displayName: result.value.displayName,
            }
          : undefined,
      });
    } catch (error) {
      if (error instanceof PollingTimeoutError) {
        return jsonTextResponse({
          success: false,
          pending: true,
          message: `Group created, but its SharePoint site did not provision within ${error.elapsedMs}ms. The site usually appears within a few more minutes — re-check /groups/${groupId}/sites/root.`,
          group: { id: groupId, displayName, mailNickname },
          operation: { attempts: error.attempts, elapsedMs: error.elapsedMs },
        });
      }
      throw error;
    }
  } catch (error) {
    return toolErrorResponse("create_team_site", error);
  }
}

// Tool 4: Get site creation status (for async callers)
export const getSiteCreationStatus: Tool = {
  name: "get_site_creation_status",
  description:
    "Check the status of an in-flight POST /beta/sites operation. Returns the operation's current status (notStarted | running | succeeded | failed) and, when succeeded, the resourceLocation pointing at the new site.",
  inputSchema: {
    type: "object",
    properties: {
      operationId: {
        type: "string",
        description:
          "Operation id returned by create_communication_site / create_team_site_classic when waitForCompletion=false.",
      },
    },
    required: ["operationId"],
  },
};

export async function handleGetSiteCreationStatus(args: GetSiteCreationStatusArgs) {
  try {
    const client = getGraphClient();
    const operationId = asString(args.operationId, "operationId");
    const op = await fetchOperationStatus(client, operationId);
    return jsonTextResponse({
      success: true,
      operation: {
        id: op.id,
        status: op.status,
        createdDateTime: op.createdDateTime,
        percentageComplete: op.percentageComplete,
        resourceId: op.resourceId,
        resourceLocation: op.resourceLocation,
        error: op.error,
      },
    });
  } catch (error) {
    return toolErrorResponse("get_site_creation_status", error);
  }
}

export const siteProvisioningTools = [
  createCommunicationSite,
  createTeamSiteClassic,
  createTeamSite,
  getSiteCreationStatus,
];

export const siteProvisioningHandlers = {
  create_communication_site: handleCreateCommunicationSite,
  create_team_site_classic: handleCreateTeamSiteClassic,
  create_team_site: handleCreateTeamSite,
  get_site_creation_status: handleGetSiteCreationStatus,
};
