/**
 * SharePoint Site Pages tools.
 *
 * Wraps the GA v1.0 page endpoints under `/sites/{site-id}/pages/...`:
 *
 *   • `GET    /sites/{site-id}/pages/microsoft.graph.sitePage`           — list
 *   • `GET    /sites/{site-id}/pages/{page-id}/microsoft.graph.sitePage` — get
 *   • `POST   /sites/{site-id}/pages`                                    — create
 *   • `PATCH  /sites/{site-id}/pages/{page-id}/microsoft.graph.sitePage` — update
 *   • `POST   /sites/{site-id}/pages/{page-id}/microsoft.graph.sitePage/publish`
 *   • `DELETE /sites/{site-id}/pages/{page-id}`
 *
 * The `/microsoft.graph.sitePage` type-cast is required on list/get/update/
 * publish — without it Graph returns the abstract `baseSitePage` and the
 * sitePage-only properties (titleArea, canvasLayout, etc.) disappear.
 *
 * Two cross-cutting validators run on every create/update body:
 *   1. `webPartType` GUIDs on standardWebParts are checked against the 14
 *      Graph-supported web parts. Anything else (hero, news, fileViewer,
 *      etc.) is rejected up-front, since Graph's error would otherwise be
 *      "Invalid request" with no hint why.
 *   2. `@odata.type` discriminators are injected where missing. Graph parses
 *      bodies without them most of the time but the failure mode when it
 *      doesn't is opaque — easier to always include them.
 *
 * Note on PATCH semantics: per Graph docs, `update_site_page` REPLACES the
 * entire `canvasLayout` — there's no partial-update path. Callers wanting to
 * add a single webpart must `get_site_page` with `expandCanvas: true`, mutate
 * the returned layout, and PATCH it back.
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getGraphClient, GraphClient } from "../../graph/client.js";
import { GraphResponse } from "../../graph/models.js";
import {
  extractPaginatedResult,
  jsonTextResponse,
  toolErrorResponse,
} from "../../graph/contracts.js";
import { resolveRequiredSharePointSite } from "../../sharepoint/site-resolver.js";

// ---------------------------------------------------------------------------
// Domain constants — from the GA v1.0 Graph docs
// ---------------------------------------------------------------------------

// The 14 web parts that Graph's create/update endpoint will accept. Names are
// for human readability; the GUID is the wire value (`webPartType`).
// Source: learn.microsoft.com/en-us/graph/api/sitepage-create (May 2026).
export const SUPPORTED_STANDARD_WEB_PARTS: Record<string, string> = {
  "bing-maps": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
  button: "0f087d7f-520e-42b7-89c0-496aaf979d58",
  "call-to-action": "df8e44e7-edd5-46d5-90da-aca1539313b8",
  divider: "2161a1c6-db61-4731-b97c-3cdb303f7cbb",
  "document-embed": "b7dd04e1-19ce-4b24-9132-b60a1c2b910d",
  image: "d1d91016-032f-456d-98a4-721247c305e8",
  "image-gallery": "af8be689-990e-492a-81f7-ba3e4cd3ed9c",
  "link-preview": "6410b3b6-d440-4663-8744-378976dc041e",
  "org-chart": "e84a8ca2-f63c-4fb9-bc0b-d8eef5ccb22b",
  people: "7f718435-ee4d-431c-bdbf-9c4ff326f46e",
  "quick-links": "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd",
  spacer: "8654b779-4886-46d4-8ffb-b5ed960ee986",
  "youtube-embed": "544dd15b-cf3c-441b-96da-004d5a8cea1d",
  "title-area": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
};

const SUPPORTED_WEB_PART_GUIDS = new Set(
  Object.values(SUPPORTED_STANDARD_WEB_PARTS).map((g) => g.toLowerCase()),
);

const HORIZONTAL_LAYOUTS = new Set([
  "none",
  "oneColumn",
  "twoColumns",
  "threeColumns",
  "oneThirdLeftColumn",
  "oneThirdRightColumn",
  "fullWidth",
  "unknownFutureValue",
]);

const EMPHASIS_VALUES = new Set([
  "none",
  "neutral",
  "soft",
  "strong",
  "unknownFutureValue",
]);

const PAGE_LAYOUTS = new Set([
  "microsoftReserved",
  "article",
  "home",
  "unknownFutureValue",
]);

const PROMOTION_KINDS = new Set(["page", "newsPost"]);

const TITLE_AREA_LAYOUTS = new Set([
  "imageAndTitle",
  "plain",
  "colorBlock",
  "overlap",
  "unknownFutureValue",
]);

// ---------------------------------------------------------------------------
// Types — wire shapes for the sitePage / canvasLayout resources
// ---------------------------------------------------------------------------

type JsonRecord = Record<string, unknown>;

// We keep these open-ended (`unknown` for nested keys) because the Graph
// schema for webpart.data.properties is per-webpart and not documented per
// type. The validators below enforce the bits we care about.
interface WebPart extends JsonRecord {
  "@odata.type"?: string;
  type?: string; // shorthand: "text" | "standard"
  id?: string;
  innerHtml?: string;
  webPartType?: string;
}

interface CanvasColumn extends JsonRecord {
  id?: string;
  width?: number;
  webparts?: WebPart[];
}

interface CanvasSection extends JsonRecord {
  id?: string;
  layout?: string;
  emphasis?: string;
  columns?: CanvasColumn[];
}

interface CanvasLayout extends JsonRecord {
  horizontalSections?: CanvasSection[];
}

interface TitleArea extends JsonRecord {
  layout?: string;
}

interface SitePageWire extends JsonRecord {
  id?: string;
  name?: string;
  title?: string;
  description?: string;
  webUrl?: string;
  pageLayout?: string;
  promotionKind?: string;
  publishingState?: { level?: string; versionId?: string };
  thumbnailWebUrl?: string;
  showComments?: boolean;
  showRecommendedPages?: boolean;
  titleArea?: TitleArea;
  canvasLayout?: CanvasLayout;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  createdBy?: { user?: { displayName?: string } };
  lastModifiedBy?: { user?: { displayName?: string } };
}

interface SiteRefArgs {
  siteId?: string;
  site?: string;
  siteUrl?: string;
}

// ---------------------------------------------------------------------------
// Validation + normalization
// ---------------------------------------------------------------------------

// Names must end in `.aspx` — every Graph doc example has it and the Site
// Pages library treats the name as a filename.
function normalizePageName(name: string): string {
  const trimmed = name.trim();
  if (!trimmed) throw new Error("Page name must not be empty.");
  return trimmed.toLowerCase().endsWith(".aspx") ? trimmed : `${trimmed}.aspx`;
}

function ensureEnum(
  value: unknown,
  allowed: Set<string>,
  label: string,
): string | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== "string" || !allowed.has(value)) {
    throw new Error(
      `${label} must be one of: ${[...allowed].join(", ")}. Got: ${JSON.stringify(value)}`,
    );
  }
  return value;
}

function isPlainObject(v: unknown): v is JsonRecord {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

// In-place: add `@odata.type` and validate web parts. Mutates the passed
// canvasLayout. We mutate rather than return-new because the structure is
// nested and the caller's reference is what we serialize.
function normalizeCanvasLayout(layout: unknown): void {
  if (layout === undefined || layout === null) return;
  if (!isPlainObject(layout)) {
    throw new Error("canvasLayout must be an object");
  }
  const sections = layout.horizontalSections;
  if (sections === undefined) return;
  if (!Array.isArray(sections)) {
    throw new Error("canvasLayout.horizontalSections must be an array");
  }

  sections.forEach((section: unknown, sIdx: number) => {
    if (!isPlainObject(section)) {
      throw new Error(`canvasLayout.horizontalSections[${sIdx}] must be an object`);
    }
    if (section.layout !== undefined) {
      ensureEnum(section.layout, HORIZONTAL_LAYOUTS, `horizontalSections[${sIdx}].layout`);
    }
    if (section.emphasis !== undefined) {
      ensureEnum(section.emphasis, EMPHASIS_VALUES, `horizontalSections[${sIdx}].emphasis`);
    }
    if (section.id === undefined) section.id = String(sIdx + 1);

    if (!Array.isArray(section.columns)) {
      throw new Error(
        `canvasLayout.horizontalSections[${sIdx}].columns must be an array`,
      );
    }
    section.columns.forEach((column: unknown, cIdx: number) => {
      if (!isPlainObject(column)) {
        throw new Error(
          `horizontalSections[${sIdx}].columns[${cIdx}] must be an object`,
        );
      }
      if (column.id === undefined) column.id = String(cIdx + 1);
      if (!Array.isArray(column.webparts)) {
        throw new Error(
          `horizontalSections[${sIdx}].columns[${cIdx}].webparts must be an array`,
        );
      }
      column.webparts.forEach((wp: unknown, wIdx: number) => {
        normalizeWebPart(wp, `horizontalSections[${sIdx}].columns[${cIdx}].webparts[${wIdx}]`);
      });
    });
  });
}

function normalizeWebPart(wp: unknown, path: string): void {
  if (!isPlainObject(wp)) {
    throw new Error(`${path} must be an object`);
  }
  const wpRec = wp as WebPart;
  // If caller passed shorthand `{ type: "text", innerHtml: "..." }`, expand it.
  if (typeof wpRec.type === "string" && !wpRec["@odata.type"]) {
    if (wpRec.type === "text") {
      wpRec["@odata.type"] = "#microsoft.graph.textWebPart";
    } else if (wpRec.type === "standard") {
      wpRec["@odata.type"] = "#microsoft.graph.standardWebPart";
    } else {
      throw new Error(
        `${path}: shorthand "type" must be "text" or "standard" — got ${JSON.stringify(wpRec.type)}`,
      );
    }
    delete wpRec.type;
  }

  const odataType = wpRec["@odata.type"];
  if (odataType === "#microsoft.graph.textWebPart") {
    if (typeof wpRec.innerHtml !== "string") {
      throw new Error(`${path}: textWebPart requires "innerHtml" (string).`);
    }
    return;
  }
  if (odataType === "#microsoft.graph.standardWebPart") {
    if (typeof wpRec.webPartType !== "string") {
      throw new Error(`${path}: standardWebPart requires "webPartType" (GUID).`);
    }
    if (!SUPPORTED_WEB_PART_GUIDS.has(wpRec.webPartType.toLowerCase())) {
      throw new Error(
        `${path}: webPartType "${wpRec.webPartType}" is not supported by Graph create/update. ` +
          `Supported keys: ${Object.keys(SUPPORTED_STANDARD_WEB_PARTS).join(", ")}.`,
      );
    }
    return;
  }
  throw new Error(
    `${path}: missing or invalid "@odata.type". Use "#microsoft.graph.textWebPart" or "#microsoft.graph.standardWebPart".`,
  );
}

function normalizeTitleArea(titleArea: unknown): void {
  if (titleArea === undefined || titleArea === null) return;
  if (!isPlainObject(titleArea)) {
    throw new Error("titleArea must be an object");
  }
  if (titleArea.layout !== undefined) {
    ensureEnum(titleArea.layout, TITLE_AREA_LAYOUTS, "titleArea.layout");
  }
}

// ---------------------------------------------------------------------------
// Shape conversion helpers
// ---------------------------------------------------------------------------

function summarizePage(page: SitePageWire) {
  return {
    id: page.id,
    name: page.name,
    title: page.title,
    description: page.description,
    webUrl: page.webUrl,
    pageLayout: page.pageLayout,
    promotionKind: page.promotionKind,
    publishingState: page.publishingState,
    thumbnailWebUrl: page.thumbnailWebUrl,
    showComments: page.showComments,
    showRecommendedPages: page.showRecommendedPages,
    createdDateTime: page.createdDateTime,
    lastModifiedDateTime: page.lastModifiedDateTime,
    createdBy: page.createdBy?.user?.displayName,
    lastModifiedBy: page.lastModifiedBy?.user?.displayName,
  };
}

// ---------------------------------------------------------------------------
// Tool: list_site_pages
// ---------------------------------------------------------------------------

export const listSitePages: Tool = {
  name: "list_site_pages",
  description:
    "List SharePoint Site Pages on a site (modern pages in the Site Pages library).",
  inputSchema: {
    type: "object",
    properties: {
      siteId: { type: "string", description: "SharePoint site ID" },
      site: { type: "string", description: "Known SharePoint site alias or canonical URL" },
      siteUrl: { type: "string", description: "Canonical SharePoint site URL" },
      filter: {
        type: "string",
        description: "OData $filter expression (e.g., \"promotionKind eq 'newsPost'\")",
      },
      orderBy: {
        type: "string",
        description: "OData $orderby (default: lastModifiedDateTime desc)",
        default: "lastModifiedDateTime desc",
      },
      limit: { type: "number", description: "Maximum number of pages to return", default: 50 },
      pageToken: {
        type: "string",
        description: "Opaque pagination token from a previous response (Graph nextLink)",
      },
    },
    required: [],
  },
};

interface ListSitePagesArgs extends SiteRefArgs {
  filter?: string;
  orderBy?: string;
  limit?: number;
  pageToken?: string;
}

export async function handleListSitePages(args: ListSitePagesArgs) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const {
      filter,
      orderBy = "lastModifiedDateTime desc",
      limit = 50,
      pageToken,
    } = args;

    const endpoint = pageToken || `/sites/${siteId}/pages/microsoft.graph.sitePage`;
    const params: Record<string, string> = {
      $top: String(limit),
      $orderby: orderBy,
    };
    if (filter) params.$filter = filter;

    const response = await client.get<GraphResponse<SitePageWire>>(
      endpoint,
      pageToken ? undefined : params,
    );

    if (!response.success || !response.data) {
      throw new Error("Failed to list site pages");
    }

    const { items, pagination } = extractPaginatedResult<SitePageWire>(response.data, limit);

    return jsonTextResponse({
      siteId,
      site: resolvedSite,
      filter: filter ?? null,
      orderBy,
      pageCount: items.length,
      pagination,
      pages: items.map(summarizePage),
    });
  } catch (error) {
    return toolErrorResponse("list_site_pages", error);
  }
}

// ---------------------------------------------------------------------------
// Tool: get_site_page
// ---------------------------------------------------------------------------

export const getSitePage: Tool = {
  name: "get_site_page",
  description:
    "Get a SharePoint Site Page by ID. Optionally expands canvasLayout to include sections/webparts.",
  inputSchema: {
    type: "object",
    properties: {
      siteId: { type: "string", description: "SharePoint site ID" },
      site: { type: "string", description: "Known SharePoint site alias or canonical URL" },
      siteUrl: { type: "string", description: "Canonical SharePoint site URL" },
      pageId: { type: "string", description: "Site page ID (GUID)" },
      expandCanvas: {
        type: "boolean",
        description:
          "Include canvasLayout (sections + webparts) in the response. Falls back to a no-expand fetch on server error (Graph beta has a known intermittent 'General Exception' on canvas expand).",
        default: false,
      },
    },
    required: ["pageId"],
  },
};

interface GetSitePageArgs extends SiteRefArgs {
  pageId?: string;
  expandCanvas?: boolean;
}

export async function handleGetSitePage(args: GetSitePageArgs) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const { pageId, expandCanvas = false } = args;
    if (!pageId) throw new Error("pageId is required");

    const endpoint = `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage`;
    const fullParams = expandCanvas ? { $expand: "canvasLayout" } : undefined;

    let response = await client.get<SitePageWire>(endpoint, fullParams);

    // The $expand=canvasLayout path has a known intermittent 500 on Graph.
    // If we asked for expansion and got nothing back, try without — at least
    // the caller gets page metadata and can investigate the canvas error.
    let canvasFallback = false;
    if (expandCanvas && (!response.success || !response.data)) {
      response = await client.get<SitePageWire>(endpoint);
      canvasFallback = true;
    }

    if (!response.success || !response.data) {
      throw new Error("Failed to fetch site page");
    }

    return jsonTextResponse({
      siteId,
      site: resolvedSite,
      page: {
        ...summarizePage(response.data),
        titleArea: response.data.titleArea,
        canvasLayout: response.data.canvasLayout,
      },
      canvasFallback,
    });
  } catch (error) {
    return toolErrorResponse("get_site_page", error);
  }
}

// ---------------------------------------------------------------------------
// Tool: create_site_page
// ---------------------------------------------------------------------------

export const createSitePage: Tool = {
  name: "create_site_page",
  description:
    "Create a new SharePoint Site Page (modern page in the Site Pages library). " +
    "The page is created as a draft (publishingState.level = checkout). Call publish_site_page to make it live.",
  inputSchema: {
    type: "object",
    properties: {
      siteId: { type: "string", description: "SharePoint site ID" },
      site: { type: "string", description: "Known SharePoint site alias or canonical URL" },
      siteUrl: { type: "string", description: "Canonical SharePoint site URL" },
      name: {
        type: "string",
        description:
          "Page filename (must be unique in the Site Pages library). `.aspx` is appended automatically if missing.",
      },
      title: { type: "string", description: "Page title (visible to users)" },
      description: { type: "string", description: "Optional page description" },
      pageLayout: {
        type: "string",
        enum: ["article", "home"],
        description:
          "Page layout. Use `article` for content pages (default), `home` for site home pages. Cannot be changed after creation.",
        default: "article",
      },
      promotionKind: {
        type: "string",
        enum: ["page", "newsPost"],
        description:
          "`page` (default) or `newsPost`. News posts appear in news rollups and the SharePoint mobile feed.",
        default: "page",
      },
      titleArea: {
        type: "object",
        description:
          "titleArea resource. Properties: layout (plain|imageAndTitle|colorBlock|overlap), title, textAboveTitle, showAuthor, showPublishedDate, imageWebUrl, alternativeText, enableGradientEffect, textAlignment.",
        additionalProperties: true,
      },
      canvasLayout: {
        type: "object",
        description:
          "canvasLayout resource: { horizontalSections: [{ layout, emphasis?, columns: [{ width, webparts: [...] }] }] }. " +
          "Webparts are either textWebPart ({ '@odata.type': '#microsoft.graph.textWebPart', innerHtml }) " +
          "or standardWebPart ({ '@odata.type': '#microsoft.graph.standardWebPart', webPartType: <GUID>, data? }). " +
          "Shorthand `{ type: 'text', innerHtml }` is expanded automatically.",
        additionalProperties: true,
      },
      showComments: { type: "boolean", description: "Allow comments on the page", default: false },
      showRecommendedPages: {
        type: "boolean",
        description: "Show 'recommended pages' rollup at the bottom",
        default: false,
      },
      publish: {
        type: "boolean",
        description: "If true, immediately publish the page after creation (single call from the caller's perspective).",
        default: false,
      },
    },
    required: ["name", "title"],
  },
};

interface CreateSitePageArgs extends SiteRefArgs {
  name?: unknown;
  title?: unknown;
  description?: unknown;
  pageLayout?: unknown;
  promotionKind?: unknown;
  titleArea?: unknown;
  canvasLayout?: unknown;
  showComments?: unknown;
  showRecommendedPages?: unknown;
  publish?: unknown;
}

export async function handleCreateSitePage(args: CreateSitePageArgs) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;

    if (typeof args.name !== "string") throw new Error("name is required (string)");
    if (typeof args.title !== "string") throw new Error("title is required (string)");

    const pageLayout = ensureEnum(args.pageLayout ?? "article", PAGE_LAYOUTS, "pageLayout") ?? "article";
    const promotionKind = ensureEnum(args.promotionKind, PROMOTION_KINDS, "promotionKind");

    normalizeTitleArea(args.titleArea);
    normalizeCanvasLayout(args.canvasLayout);

    const body: JsonRecord = {
      "@odata.type": "#microsoft.graph.sitePage",
      name: normalizePageName(args.name),
      title: args.title,
      pageLayout,
    };
    if (args.description !== undefined) body.description = args.description;
    if (promotionKind !== undefined) body.promotionKind = promotionKind;
    if (args.titleArea !== undefined) body.titleArea = args.titleArea;
    if (args.canvasLayout !== undefined) body.canvasLayout = args.canvasLayout;
    if (args.showComments !== undefined) body.showComments = args.showComments;
    if (args.showRecommendedPages !== undefined) {
      body.showRecommendedPages = args.showRecommendedPages;
    }

    const createResponse = await client.post<SitePageWire>(`/sites/${siteId}/pages`, body);
    if (!createResponse.success || !createResponse.data) {
      throw new Error("Page creation returned no data");
    }
    const created = createResponse.data;

    let publishResult: { published: boolean; error?: string } | undefined;
    if (args.publish === true && typeof created.id === "string") {
      publishResult = await tryPublish(client, siteId, created.id);
    }

    return jsonTextResponse({
      success: true,
      message: publishResult?.published
        ? "Page created and published"
        : "Page created (draft)",
      siteId,
      site: resolvedSite,
      page: summarizePage(created),
      publish: publishResult,
    });
  } catch (error) {
    return toolErrorResponse("create_site_page", error);
  }
}

async function tryPublish(
  client: GraphClient,
  siteId: string,
  pageId: string,
): Promise<{ published: boolean; error?: string }> {
  try {
    const resp = await client.post<unknown>(
      `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage/publish`,
      {},
    );
    return { published: resp.success === true };
  } catch (err) {
    return {
      published: false,
      error: err instanceof Error ? err.message : String(err),
    };
  }
}

// ---------------------------------------------------------------------------
// Tool: update_site_page
// ---------------------------------------------------------------------------

export const updateSitePage: Tool = {
  name: "update_site_page",
  description:
    "Update a SharePoint Site Page. NOTE: canvasLayout is replaced wholesale — partial updates are not supported by Graph. " +
    "To edit one webpart, get_site_page with expandCanvas=true, mutate the returned layout, then pass it back here.",
  inputSchema: {
    type: "object",
    properties: {
      siteId: { type: "string", description: "SharePoint site ID" },
      site: { type: "string", description: "Known SharePoint site alias or canonical URL" },
      siteUrl: { type: "string", description: "Canonical SharePoint site URL" },
      pageId: { type: "string", description: "Site page ID (GUID)" },
      title: { type: "string" },
      description: { type: "string" },
      titleArea: { type: "object", additionalProperties: true },
      canvasLayout: { type: "object", additionalProperties: true },
      showComments: { type: "boolean" },
      showRecommendedPages: { type: "boolean" },
      promotionKind: {
        type: "string",
        enum: ["page", "newsPost"],
        description: "Demotion (newsPost → page) is not supported by Graph.",
      },
      thumbnailWebUrl: { type: "string" },
    },
    required: ["pageId"],
  },
};

interface UpdateSitePageArgs extends SiteRefArgs {
  pageId?: string;
  title?: unknown;
  description?: unknown;
  titleArea?: unknown;
  canvasLayout?: unknown;
  showComments?: unknown;
  showRecommendedPages?: unknown;
  promotionKind?: unknown;
  thumbnailWebUrl?: unknown;
}

export async function handleUpdateSitePage(args: UpdateSitePageArgs) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const { pageId } = args;
    if (!pageId) throw new Error("pageId is required");

    const promotionKind = ensureEnum(args.promotionKind, PROMOTION_KINDS, "promotionKind");
    normalizeTitleArea(args.titleArea);
    normalizeCanvasLayout(args.canvasLayout);

    const body: JsonRecord = { "@odata.type": "#microsoft.graph.sitePage" };
    for (const key of [
      "title",
      "description",
      "titleArea",
      "canvasLayout",
      "showComments",
      "showRecommendedPages",
      "thumbnailWebUrl",
    ] as const) {
      const value = args[key];
      if (value !== undefined) body[key] = value;
    }
    if (promotionKind !== undefined) body.promotionKind = promotionKind;

    // Detect no-op early; Graph will accept it but the response is confusing.
    const onlyDiscriminator = Object.keys(body).length === 1;
    if (onlyDiscriminator) {
      throw new Error(
        "update_site_page called with no mutable fields. Pass at least one of: title, description, titleArea, canvasLayout, showComments, showRecommendedPages, promotionKind, thumbnailWebUrl.",
      );
    }

    const endpoint = `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage`;
    const response = await client.patch<SitePageWire>(endpoint, body);
    if (!response.success || !response.data) {
      throw new Error("Page update returned no data");
    }

    return jsonTextResponse({
      success: true,
      message: "Page updated (draft — call publish_site_page to make live)",
      siteId,
      site: resolvedSite,
      page: summarizePage(response.data),
    });
  } catch (error) {
    return toolErrorResponse("update_site_page", error);
  }
}

// ---------------------------------------------------------------------------
// Tool: publish_site_page
// ---------------------------------------------------------------------------

export const publishSitePage: Tool = {
  name: "publish_site_page",
  description:
    "Publish a draft/checked-out SharePoint Site Page. " +
    "If a page-approval flow is configured on the Site Pages library, the page won't go live until the approval completes.",
  inputSchema: {
    type: "object",
    properties: {
      siteId: { type: "string", description: "SharePoint site ID" },
      site: { type: "string", description: "Known SharePoint site alias or canonical URL" },
      siteUrl: { type: "string", description: "Canonical SharePoint site URL" },
      pageId: { type: "string", description: "Site page ID (GUID)" },
    },
    required: ["pageId"],
  },
};

interface PublishSitePageArgs extends SiteRefArgs {
  pageId?: string;
}

export async function handlePublishSitePage(args: PublishSitePageArgs) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const { pageId } = args;
    if (!pageId) throw new Error("pageId is required");

    const endpoint = `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage/publish`;
    const response = await client.post<unknown>(endpoint, {});
    if (!response.success) {
      throw new Error("Publish call did not succeed");
    }

    // Publish returns 204 with no body. Refetch so the caller sees the new
    // publishingState (level: published, versionId bumped).
    const fetched = await client.get<SitePageWire>(`/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage`);

    return jsonTextResponse({
      success: true,
      message: "Page published",
      siteId,
      site: resolvedSite,
      pageId,
      page: fetched.success && fetched.data ? summarizePage(fetched.data) : null,
    });
  } catch (error) {
    return toolErrorResponse("publish_site_page", error);
  }
}

// ---------------------------------------------------------------------------
// Tool: delete_site_page
// ---------------------------------------------------------------------------

export const deleteSitePage: Tool = {
  name: "delete_site_page",
  description: "Delete a SharePoint Site Page. The deleted page goes to the site's Recycle Bin.",
  inputSchema: {
    type: "object",
    properties: {
      siteId: { type: "string", description: "SharePoint site ID" },
      site: { type: "string", description: "Known SharePoint site alias or canonical URL" },
      siteUrl: { type: "string", description: "Canonical SharePoint site URL" },
      pageId: { type: "string", description: "Site page ID (GUID)" },
    },
    required: ["pageId"],
  },
};

interface DeleteSitePageArgs extends SiteRefArgs {
  pageId?: string;
}

export async function handleDeleteSitePage(args: DeleteSitePageArgs) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const { pageId } = args;
    if (!pageId) throw new Error("pageId is required");

    const endpoint = `/sites/${siteId}/pages/${pageId}`;
    const response = await client.delete(endpoint);
    if (!response.success) {
      throw new Error("Page delete did not succeed");
    }

    return jsonTextResponse({
      success: true,
      message: "Page deleted (moved to Recycle Bin)",
      siteId,
      site: resolvedSite,
      pageId,
    });
  } catch (error) {
    return toolErrorResponse("delete_site_page", error);
  }
}

// ---------------------------------------------------------------------------
// Bundle for re-export from sharepoint/index.ts
// ---------------------------------------------------------------------------

export const pageTools: Tool[] = [
  listSitePages,
  getSitePage,
  createSitePage,
  updateSitePage,
  publishSitePage,
  deleteSitePage,
];

export const pageHandlers = {
  list_site_pages: handleListSitePages,
  get_site_page: handleGetSitePage,
  create_site_page: handleCreateSitePage,
  update_site_page: handleUpdateSitePage,
  publish_site_page: handlePublishSitePage,
  delete_site_page: handleDeleteSitePage,
};
