import test from "node:test";
import assert from "node:assert/strict";

import {
  __setGraphClientInstanceForTests,
  GraphClient,
} from "../graph/client.js";
import {
  __setAuthInstanceForTests,
  MicrosoftGraphAuth,
} from "../auth/microsoft-graph-auth.js";
import {
  __setKnownSitesForTests,
  __setTenantHostnameForTests,
  __resetKnownSitesForTests,
} from "../sharepoint/site-resolver.js";
import {
  handleCreateCommunicationSite,
  handleCreateTeamSiteClassic,
  handleCreateTeamSite,
  handleGetSiteCreationStatus,
  __setPollingSleepForTests,
} from "../tools/sharepoint/site-provisioning.js";
import { registerGraphClientTestLifecycle } from "./helpers/test-lifecycle.js";
import {
  createMockGraphClient,
  parsePayload,
  type ToolEnvelope,
} from "./helpers/tool-test-helpers.js";

registerGraphClientTestLifecycle();

// Shared per-test cleanup: clear auth, tenant cache, polling sleep, site
// registry so failures in one test don't poison the next.
test.afterEach(() => {
  __setAuthInstanceForTests(null);
  __setPollingSleepForTests(undefined);
  __setTenantHostnameForTests(null);
  __resetKnownSitesForTests();
});

type MockClientShape = ReturnType<typeof createMockGraphClient>["client"];

// Centralize the `as unknown as GraphClient` cast — the mock client only
// implements a subset of GraphClient, but the test injection point requires
// the full type. Keeping the cast in one place keeps the test bodies clean.
function installMockClient(mock: { client: MockClientShape }): void {
  __setGraphClientInstanceForTests(mock.client as unknown as GraphClient);
}

type MockRequestOptions = { apiVersion?: "v1.0" | "beta" } | undefined;

function installDelegatedAuth(): void {
  __setAuthInstanceForTests({
    getAuthMode: () => "device_code",
  } as unknown as MicrosoftGraphAuth);
}

function installAppOnlyAuth(): void {
  __setAuthInstanceForTests({
    getAuthMode: () => "client_credentials",
  } as unknown as MicrosoftGraphAuth);
}

// Tests fast-forward through the polling loop so we don't burn wall-clock time
// on artificial delays.
function installInstantPolling(): void {
  __setPollingSleepForTests(async () => {
    /* no-op — tests don't need real sleeps */
  });
}

function installTenantHost(host = "contoso.sharepoint.com"): void {
  __setTenantHostnameForTests(host);
}

const LOCATION_HEADER = (operationId: string) =>
  `https://graph.microsoft.com/beta/sites/getOperationStatus(operationId='${operationId}')`;

test("create_communication_site posts to /beta/sites and polls until succeeded", async () => {
  installDelegatedAuth();
  installInstantPolling();
  installTenantHost();

  let statusProbeCount = 0;
  const mock = createMockGraphClient({
    post: async (endpoint: string, body: Record<string, unknown>, options: MockRequestOptions) => {
      assert.equal(endpoint, "/sites");
      assert.equal(options?.apiVersion, "beta");
      assert.equal(body.template, "sitepagepublishing");
      assert.equal(body.webUrl, "https://contoso.sharepoint.com/sites/marketing-2026");
      assert.equal(body.locale, "es-MX");
      return {
        success: true,
        data: {},
        metadata: {
          status: 202,
          headers: { location: LOCATION_HEADER("op-123") },
        },
      };
    },
    get: async (endpoint: string, _params: unknown, options: MockRequestOptions) => {
      if (endpoint.startsWith("/sites/getOperationStatus")) {
        assert.equal(options?.apiVersion, "beta");
        statusProbeCount += 1;
        if (statusProbeCount < 2) {
          return { success: true, data: { id: "op-123", status: "running" } };
        }
        return {
          success: true,
          data: {
            id: "op-123",
            status: "succeeded",
            resourceId: "contoso.sharepoint.com,abc,def",
            resourceLocation:
              "https://graph.microsoft.com/beta/sites/contoso.sharepoint.com,abc,def",
          },
        };
      }
      if (endpoint.startsWith("https://graph.microsoft.com/beta/sites/")) {
        return {
          success: true,
          data: {
            id: "contoso.sharepoint.com,abc,def",
            displayName: "Marketing 2026",
            webUrl: "https://contoso.sharepoint.com/sites/marketing-2026",
            description: "Sitio de marketing",
          },
        };
      }
      throw new Error(`Unexpected GET ${endpoint}`);
    },
  });

  installMockClient(mock);

  const response = (await handleCreateCommunicationSite({
    displayName: "Marketing 2026",
    alias: "marketing-2026",
    description: "Sitio de marketing",
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.success, true);
  assert.equal(payload.operation.status, "succeeded");
  assert.equal(payload.operation.operationId, "op-123");
  assert.equal(payload.site.id, "contoso.sharepoint.com,abc,def");
  assert.equal(payload.site.webUrl, "https://contoso.sharepoint.com/sites/marketing-2026");
  assert.ok(statusProbeCount >= 2, "should have polled getOperationStatus at least twice");
});

test("create_communication_site returns operationId immediately when waitForCompletion=false", async () => {
  installDelegatedAuth();
  installTenantHost();

  const mock = createMockGraphClient({
    post: async () => ({
      success: true,
      data: {},
      metadata: { status: 202, headers: { location: LOCATION_HEADER("op-async") } },
    }),
  });

  installMockClient(mock);

  const response = (await handleCreateCommunicationSite({
    displayName: "Async Site",
    alias: "async-site",
    waitForCompletion: false,
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.success, true);
  assert.equal(payload.pending, true);
  assert.equal(payload.operation.operationId, "op-async");
  // No GET calls should have been issued — caller opted out of polling.
  assert.equal(mock.methodCalls("get").length, 0);
});

test("create_communication_site rejects invalid alias before calling Graph", async () => {
  installDelegatedAuth();
  installTenantHost();

  const mock = createMockGraphClient();
  installMockClient(mock);

  const response = (await handleCreateCommunicationSite({
    displayName: "Bad Alias",
    alias: "Has Spaces and CAPS",
  })) as ToolEnvelope;

  assert.equal(response.isError, true);
  // Must not have called any Graph endpoint.
  assert.equal(mock.calls.length, 0);
});

test("create_communication_site requires ownerEmail in client-credentials mode", async () => {
  installAppOnlyAuth();
  installTenantHost();

  const mock = createMockGraphClient();
  installMockClient(mock);

  const response = (await handleCreateCommunicationSite({
    displayName: "No Owner",
    alias: "no-owner",
  })) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.equal(mock.calls.length, 0);
});

test("create_communication_site reports pending status when polling times out", async () => {
  installDelegatedAuth();
  installInstantPolling();
  installTenantHost();

  const mock = createMockGraphClient({
    post: async () => ({
      success: true,
      data: {},
      metadata: { status: 202, headers: { location: LOCATION_HEADER("op-slow") } },
    }),
    get: async (endpoint: string) => {
      if (endpoint.startsWith("/sites/getOperationStatus")) {
        // Never resolves — always "running".
        return { success: true, data: { id: "op-slow", status: "running" } };
      }
      throw new Error(`Unexpected GET ${endpoint}`);
    },
  });

  installMockClient(mock);

  const response = (await handleCreateCommunicationSite({
    displayName: "Slow Site",
    alias: "slow-site",
    timeoutSeconds: 0.05,
    intervalSeconds: 0.01,
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.success, false);
  assert.equal(payload.pending, true);
  assert.equal(payload.operation.operationId, "op-slow");
});

test("create_team_site_classic uses template=sts", async () => {
  installDelegatedAuth();
  installInstantPolling();
  installTenantHost();

  const capturedBodies: Array<Record<string, unknown>> = [];
  const mock = createMockGraphClient({
    post: async (_endpoint: string, body: Record<string, unknown>) => {
      capturedBodies.push(body);
      return {
        success: true,
        data: {},
        metadata: { status: 202, headers: { location: LOCATION_HEADER("op-sts") } },
      };
    },
    get: async (endpoint: string) => {
      if (endpoint.startsWith("/sites/getOperationStatus")) {
        return {
          success: true,
          data: {
            id: "op-sts",
            status: "succeeded",
            resourceId: "contoso.sharepoint.com,team,123",
          },
        };
      }
      throw new Error(`Unexpected GET ${endpoint}`);
    },
  });

  installMockClient(mock);

  const response = (await handleCreateTeamSiteClassic({
    displayName: "Equipo Sin Grupo",
    alias: "equipo-sin-grupo",
  })) as ToolEnvelope;

  assert.equal(response.isError, undefined);
  assert.equal(capturedBodies[0].template, "sts");
});

test("create_team_site posts /groups then polls /groups/{id}/sites/root", async () => {
  installDelegatedAuth();
  installInstantPolling();

  const userIdByEmail: Record<string, string> = {
    "owner@contoso.com": "user-owner-1",
  };
  let siteProbes = 0;

  const mock = createMockGraphClient({
    get: async (endpoint: string) => {
      const userMatch = endpoint.match(/^\/users\/(.+)$/);
      if (userMatch) {
        const email = decodeURIComponent(userMatch[1]);
        const id = userIdByEmail[email];
        if (!id) throw new Error(`unknown user ${email}`);
        return { success: true, data: { id } };
      }
      if (endpoint === "/groups/group-1/sites/root") {
        siteProbes += 1;
        if (siteProbes < 2) {
          // Simulate Graph 404 while the site is provisioning.
          throw Object.assign(new Error("not yet"), { statusCode: 404 });
        }
        return {
          success: true,
          data: {
            id: "contoso.sharepoint.com,team-site-id",
            displayName: "Equipo Marketing",
            webUrl: "https://contoso.sharepoint.com/sites/equipo-marketing",
          },
        };
      }
      throw new Error(`Unexpected GET ${endpoint}`);
    },
    post: async (endpoint: string, body: Record<string, unknown>) => {
      assert.equal(endpoint, "/groups");
      assert.deepEqual(body.groupTypes, ["Unified"]);
      assert.equal(body.mailEnabled, true);
      assert.equal(body.securityEnabled, false);
      assert.equal(body.mailNickname, "equipo-marketing");
      assert.deepEqual(body["owners@odata.bind"], [
        "https://graph.microsoft.com/v1.0/users/user-owner-1",
      ]);
      return { success: true, data: { id: "group-1", displayName: body.displayName } };
    },
  });

  installMockClient(mock);

  const response = (await handleCreateTeamSite({
    displayName: "Equipo Marketing",
    mailNickname: "equipo-marketing",
    ownerEmails: ["owner@contoso.com"],
    timeoutSeconds: 5,
    intervalSeconds: 0.01,
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.success, true);
  assert.equal(payload.group.id, "group-1");
  assert.equal(payload.site.id, "contoso.sharepoint.com,team-site-id");
  assert.ok(siteProbes >= 2, "should retry past initial 404");
});

test("create_team_site requires ownerEmails in client-credentials mode", async () => {
  installAppOnlyAuth();

  const mock = createMockGraphClient();
  installMockClient(mock);

  const response = (await handleCreateTeamSite({
    displayName: "Sin Dueños",
    mailNickname: "sin-duenos",
  })) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.equal(mock.calls.length, 0);
});

test("create_team_site rejects mailNickname starting with a period", async () => {
  installDelegatedAuth();

  const mock = createMockGraphClient();
  installMockClient(mock);

  const response = (await handleCreateTeamSite({
    displayName: "Bad",
    mailNickname: ".leading-dot",
  })) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.equal(mock.calls.length, 0);
});

test("create_team_site skips site polling when waitForSite=false", async () => {
  installDelegatedAuth();

  const mock = createMockGraphClient({
    post: async () => ({ success: true, data: { id: "group-skip" } }),
  });
  installMockClient(mock);

  const response = (await handleCreateTeamSite({
    displayName: "Sin esperar",
    mailNickname: "sin-esperar",
    waitForSite: false,
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.group.id, "group-skip");
  assert.equal(payload.pending, true);
  assert.equal(mock.methodCalls("get").length, 0);
});

test("create_team_site honors provisionSiteOnDemand", async () => {
  installDelegatedAuth();

  const capturedBodies: Array<Record<string, unknown>> = [];
  const mock = createMockGraphClient({
    post: async (_endpoint: string, body: Record<string, unknown>) => {
      capturedBodies.push(body);
      return { success: true, data: { id: "group-pod" } };
    },
  });
  installMockClient(mock);

  const response = (await handleCreateTeamSite({
    displayName: "On demand",
    mailNickname: "on-demand",
    provisionSiteOnDemand: true,
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.ok(
    (capturedBodies[0].resourceBehaviorOptions as string[]).includes(
      "ProvisionSiteOnDemand",
    ),
    "ProvisionSiteOnDemand flag should be sent to Graph",
  );
  assert.equal(payload.pending, false);
  assert.equal(mock.methodCalls("get").length, 0);
});

test("get_site_creation_status returns the polled operation payload", async () => {
  installDelegatedAuth();

  const mock = createMockGraphClient({
    get: async (endpoint: string, _params: unknown, options: MockRequestOptions) => {
      assert.equal(
        endpoint,
        "/sites/getOperationStatus(operationId='op-status-1')",
      );
      assert.equal(options?.apiVersion, "beta");
      return {
        success: true,
        data: {
          id: "op-status-1",
          status: "succeeded",
          percentageComplete: 100,
          resourceId: "contoso.sharepoint.com,xyz",
        },
      };
    },
  });
  installMockClient(mock);

  const response = (await handleGetSiteCreationStatus({
    operationId: "op-status-1",
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.operation.status, "succeeded");
  assert.equal(payload.operation.percentageComplete, 100);
});

test("get_site_creation_status rejects operationId containing a single quote", async () => {
  installDelegatedAuth();

  const mock = createMockGraphClient();
  installMockClient(mock);

  const response = (await handleGetSiteCreationStatus({
    operationId: "op'injection",
  })) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.equal(mock.calls.length, 0);
});

test("create_communication_site resolves tenant host from the site registry", async () => {
  installDelegatedAuth();
  installInstantPolling();
  __setKnownSitesForTests([
    {
      key: "primary",
      name: "Primary",
      siteId: "site-id-1",
      siteUrl: "https://lanpro.sharepoint.com/sites/primary",
      aliases: ["primary"],
    },
  ]);

  const mock = createMockGraphClient({
    post: async (_endpoint: string, body: Record<string, unknown>) => {
      assert.equal(body.webUrl, "https://lanpro.sharepoint.com/sites/new");
      return {
        success: true,
        data: {},
        metadata: { status: 202, headers: { location: LOCATION_HEADER("op-host") } },
      };
    },
    get: async (endpoint: string) => {
      if (endpoint.startsWith("/sites/getOperationStatus")) {
        return {
          success: true,
          data: { id: "op-host", status: "succeeded", resourceId: "host-resource" },
        };
      }
      throw new Error(`Unexpected GET ${endpoint}`);
    },
  });
  installMockClient(mock);

  const response = (await handleCreateCommunicationSite({
    displayName: "Nuevo",
    alias: "new",
  })) as ToolEnvelope;

  assert.equal(response.isError, undefined);
});
