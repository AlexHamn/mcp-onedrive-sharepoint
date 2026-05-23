import test from "node:test";
import assert from "node:assert/strict";

import {
  __setGraphClientInstanceForTests,
  GraphClient,
} from "../graph/client.js";
import {
  __setKnownSitesForTests,
  __resetKnownSitesForTests,
} from "../sharepoint/site-resolver.js";
import {
  handleListSitePages,
  handleGetSitePage,
  handleCreateSitePage,
  handleUpdateSitePage,
  handlePublishSitePage,
  handleDeleteSitePage,
  SUPPORTED_STANDARD_WEB_PARTS,
} from "../tools/sharepoint/pages.js";
import { registerGraphClientTestLifecycle } from "./helpers/test-lifecycle.js";
import {
  createMockGraphClient,
  parsePayload,
} from "./helpers/tool-test-helpers.js";

registerGraphClientTestLifecycle();

const SITE_ID = "lanpromx.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222";
const PAGE_ID = "01ABCD0123456789ABCDEF0123456789AB";

test.beforeEach(() => {
  __setKnownSitesForTests([
    {
      key: "test",
      name: "Test site",
      siteId: SITE_ID,
      siteUrl: "https://lanpromx.sharepoint.com/sites/test",
      driveId: "b!testdrive",
      aliases: ["test"],
    },
  ]);
});

test.afterEach(() => {
  __resetKnownSitesForTests();
});

function installMockClient(mock: ReturnType<typeof createMockGraphClient>) {
  __setGraphClientInstanceForTests(mock.client as unknown as GraphClient);
}

// ---------------------------------------------------------------------------
// list_site_pages
// ---------------------------------------------------------------------------

test("list_site_pages hits the typed /pages endpoint and returns summaries", async () => {
  const mock = createMockGraphClient({
    get: async () => ({
      success: true,
      data: {
        value: [
          {
            id: "p1",
            name: "Welcome.aspx",
            title: "Welcome",
            webUrl: "https://x/welcome",
            pageLayout: "article",
            promotionKind: "page",
            publishingState: { level: "published", versionId: "1.0" },
            createdBy: { user: { displayName: "Luis" } },
          },
        ],
      },
    }),
  });
  installMockClient(mock);

  const res = await handleListSitePages({ siteId: SITE_ID });
  const body = parsePayload(res);

  assert.equal(body.pageCount, 1);
  assert.equal(body.pages[0].name, "Welcome.aspx");
  assert.equal(body.pages[0].createdBy, "Luis");

  const call = mock.calls.find((c) => c.method === "get");
  assert.ok(call, "expected a GET");
  assert.equal(
    call!.args[0],
    `/sites/${SITE_ID}/pages/microsoft.graph.sitePage`,
    "list endpoint must include the sitePage type cast",
  );
  assert.match(call!.args[1].$orderby, /lastModifiedDateTime/);
});

test("list_site_pages forwards a custom $filter when provided", async () => {
  const mock = createMockGraphClient({
    get: async () => ({ success: true, data: { value: [] } }),
  });
  installMockClient(mock);

  await handleListSitePages({
    siteId: SITE_ID,
    filter: "promotionKind eq 'newsPost'",
  });

  const call = mock.calls.find((c) => c.method === "get")!;
  assert.equal(call.args[1].$filter, "promotionKind eq 'newsPost'");
});

// ---------------------------------------------------------------------------
// get_site_page
// ---------------------------------------------------------------------------

test("get_site_page without expand omits canvasLayout from the request", async () => {
  const mock = createMockGraphClient({
    get: async () => ({
      success: true,
      data: { id: PAGE_ID, name: "x.aspx", title: "X" },
    }),
  });
  installMockClient(mock);

  await handleGetSitePage({ siteId: SITE_ID, pageId: PAGE_ID });
  const call = mock.calls.find((c) => c.method === "get")!;
  assert.equal(call.args[0], `/sites/${SITE_ID}/pages/${PAGE_ID}/microsoft.graph.sitePage`);
  assert.equal(call.args[1], undefined);
});

test("get_site_page with expandCanvas adds the expand param", async () => {
  const mock = createMockGraphClient({
    get: async () => ({
      success: true,
      data: { id: PAGE_ID, canvasLayout: { horizontalSections: [] } },
    }),
  });
  installMockClient(mock);

  const res = await handleGetSitePage({
    siteId: SITE_ID,
    pageId: PAGE_ID,
    expandCanvas: true,
  });
  const body = parsePayload(res);
  assert.equal(body.canvasFallback, false);
  assert.deepEqual(body.page.canvasLayout, { horizontalSections: [] });

  const call = mock.calls.find((c) => c.method === "get")!;
  assert.equal(call.args[1].$expand, "canvasLayout");
});

test("get_site_page falls back to a no-expand fetch when canvas expand fails", async () => {
  let gets = 0;
  const mock = createMockGraphClient({
    get: async () => {
      gets++;
      if (gets === 1) return { success: false, data: null };
      return { success: true, data: { id: PAGE_ID, name: "x.aspx", title: "X" } };
    },
  });
  installMockClient(mock);

  const res = await handleGetSitePage({
    siteId: SITE_ID,
    pageId: PAGE_ID,
    expandCanvas: true,
  });
  const body = parsePayload(res);
  assert.equal(body.canvasFallback, true);
  assert.equal(gets, 2);
});

// ---------------------------------------------------------------------------
// create_site_page
// ---------------------------------------------------------------------------

test("create_site_page appends .aspx to the name and sets the @odata.type", async () => {
  let captured: Record<string, unknown> = {};
  const mock = createMockGraphClient({
    post: async (_endpoint: string, body: Record<string, unknown>) => {
      captured = body;
      return { success: true, data: { id: PAGE_ID, name: body.name, title: body.title } };
    },
  });
  installMockClient(mock);

  await handleCreateSitePage({
    siteId: SITE_ID,
    name: "bienvenida",
    title: "Bienvenida",
  });

  assert.equal(captured.name, "bienvenida.aspx");
  assert.equal(captured["@odata.type"], "#microsoft.graph.sitePage");
  assert.equal(captured.pageLayout, "article");
});

test("create_site_page leaves an explicit .aspx alone", async () => {
  let captured: Record<string, unknown> = {};
  const mock = createMockGraphClient({
    post: async (_e: string, body: Record<string, unknown>) => {
      captured = body;
      return { success: true, data: { id: PAGE_ID, name: body.name } };
    },
  });
  installMockClient(mock);

  await handleCreateSitePage({
    siteId: SITE_ID,
    name: "MyPage.aspx",
    title: "X",
  });
  assert.equal(captured.name, "MyPage.aspx");
});

test("create_site_page expands text-webpart shorthand", async () => {
  let captured: Record<string, unknown> = {};
  const mock = createMockGraphClient({
    post: async (_e: string, body: Record<string, unknown>) => {
      captured = body;
      return { success: true, data: { id: PAGE_ID } };
    },
  });
  installMockClient(mock);

  await handleCreateSitePage({
    siteId: SITE_ID,
    name: "x",
    title: "X",
    canvasLayout: {
      horizontalSections: [
        {
          layout: "oneColumn",
          columns: [{ width: 12, webparts: [{ type: "text", innerHtml: "<p>hi</p>" }] }],
        },
      ],
    },
  });

  const layout = captured.canvasLayout as {
    horizontalSections: Array<{ columns: Array<{ webparts: Array<Record<string, unknown>> }> }>;
  };
  const wp = layout.horizontalSections[0].columns[0].webparts[0];
  assert.equal(wp["@odata.type"], "#microsoft.graph.textWebPart");
  assert.equal(wp.innerHtml, "<p>hi</p>");
  assert.equal(wp.type, undefined, "shorthand `type` should be stripped");
});

test("create_site_page auto-assigns section + column IDs when missing", async () => {
  let captured: Record<string, unknown> = {};
  const mock = createMockGraphClient({
    post: async (_e: string, body: Record<string, unknown>) => {
      captured = body;
      return { success: true, data: { id: PAGE_ID } };
    },
  });
  installMockClient(mock);

  await handleCreateSitePage({
    siteId: SITE_ID,
    name: "x",
    title: "X",
    canvasLayout: {
      horizontalSections: [
        {
          layout: "twoColumns",
          columns: [
            { width: 6, webparts: [{ type: "text", innerHtml: "<p>a</p>" }] },
            { width: 6, webparts: [{ type: "text", innerHtml: "<p>b</p>" }] },
          ],
        },
      ],
    },
  });

  const layout = captured.canvasLayout as {
    horizontalSections: Array<{ id: string; columns: Array<{ id: string }> }>;
  };
  const section = layout.horizontalSections[0];
  assert.equal(section.id, "1");
  assert.equal(section.columns[0].id, "1");
  assert.equal(section.columns[1].id, "2");
});

test("create_site_page rejects an invalid horizontalSection layout", async () => {
  const mock = createMockGraphClient({ post: async () => ({ success: true, data: {} }) });
  installMockClient(mock);

  const res = await handleCreateSitePage({
    siteId: SITE_ID,
    name: "x",
    title: "X",
    canvasLayout: {
      horizontalSections: [
        {
          layout: "fivesixthcolumn", // bogus
          columns: [{ width: 12, webparts: [] }],
        },
      ],
    },
  });
  assert.ok((res as { isError?: boolean }).isError, "expected isError=true on validation failure");
  assert.match(res.content[0].text, /horizontalSections\[0\]\.layout/);
});

test("create_site_page rejects an unsupported webPartType GUID", async () => {
  const mock = createMockGraphClient({ post: async () => ({ success: true, data: {} }) });
  installMockClient(mock);

  const res = await handleCreateSitePage({
    siteId: SITE_ID,
    name: "x",
    title: "X",
    canvasLayout: {
      horizontalSections: [
        {
          layout: "oneColumn",
          columns: [
            {
              width: 12,
              webparts: [
                {
                  "@odata.type": "#microsoft.graph.standardWebPart",
                  webPartType: "00000000-0000-0000-0000-000000000000", // fake
                },
              ],
            },
          ],
        },
      ],
    },
  });
  assert.ok((res as { isError?: boolean }).isError);
  assert.match(res.content[0].text, /not supported by Graph create\/update/);
});

test("create_site_page accepts every supported standardWebPart GUID", async () => {
  const mock = createMockGraphClient({ post: async () => ({ success: true, data: { id: PAGE_ID } }) });
  installMockClient(mock);

  for (const [name, guid] of Object.entries(SUPPORTED_STANDARD_WEB_PARTS)) {
    const res = await handleCreateSitePage({
      siteId: SITE_ID,
      name: `wp-${name}`,
      title: name,
      canvasLayout: {
        horizontalSections: [
          {
            layout: "oneColumn",
            columns: [
              {
                width: 12,
                webparts: [
                  {
                    "@odata.type": "#microsoft.graph.standardWebPart",
                    webPartType: guid,
                  },
                ],
              },
            ],
          },
        ],
      },
    });
    assert.ok(!(res as { isError?: boolean }).isError, `webpart ${name} should pass validation`);
  }
});

test("create_site_page with publish=true also calls the publish endpoint", async () => {
  const calls: string[] = [];
  const mock = createMockGraphClient({
    post: async (endpoint: string) => {
      calls.push(endpoint);
      return { success: true, data: { id: PAGE_ID } };
    },
  });
  installMockClient(mock);

  const res = await handleCreateSitePage({
    siteId: SITE_ID,
    name: "x",
    title: "X",
    publish: true,
  });
  const body = parsePayload(res);
  assert.equal(body.publish.published, true);
  assert.equal(calls.length, 2);
  assert.ok(calls[1].endsWith("/microsoft.graph.sitePage/publish"));
});

// ---------------------------------------------------------------------------
// update_site_page
// ---------------------------------------------------------------------------

test("update_site_page sends PATCH with @odata.type and mutable fields only", async () => {
  let captured: Record<string, unknown> = {};
  let capturedEndpoint = "";
  const mock = createMockGraphClient({
    patch: async (endpoint: string, body: Record<string, unknown>) => {
      capturedEndpoint = endpoint;
      captured = body;
      return { success: true, data: { id: PAGE_ID, title: body.title } };
    },
  });
  installMockClient(mock);

  await handleUpdateSitePage({
    siteId: SITE_ID,
    pageId: PAGE_ID,
    title: "Nuevo título",
    description: "Nueva descripción",
  });

  assert.equal(
    capturedEndpoint,
    `/sites/${SITE_ID}/pages/${PAGE_ID}/microsoft.graph.sitePage`,
  );
  assert.equal(captured["@odata.type"], "#microsoft.graph.sitePage");
  assert.equal(captured.title, "Nuevo título");
  assert.equal(captured.description, "Nueva descripción");
  // pageLayout is read-only on update; must not appear.
  assert.equal(captured.pageLayout, undefined);
});

test("update_site_page rejects a body with only the discriminator", async () => {
  const mock = createMockGraphClient({ patch: async () => ({ success: true, data: {} }) });
  installMockClient(mock);

  const res = await handleUpdateSitePage({ siteId: SITE_ID, pageId: PAGE_ID });
  assert.ok((res as { isError?: boolean }).isError);
  assert.match(res.content[0].text, /no mutable fields/);
});

// ---------------------------------------------------------------------------
// publish_site_page + delete_site_page
// ---------------------------------------------------------------------------

test("publish_site_page POSTs the publish action then refetches metadata", async () => {
  const calls: { method: string; endpoint: string }[] = [];
  const mock = createMockGraphClient({
    post: async (endpoint: string) => {
      calls.push({ method: "post", endpoint });
      return { success: true };
    },
    get: async (endpoint: string) => {
      calls.push({ method: "get", endpoint });
      return {
        success: true,
        data: {
          id: PAGE_ID,
          name: "x.aspx",
          publishingState: { level: "published", versionId: "1.0" },
        },
      };
    },
  });
  installMockClient(mock);

  const res = await handlePublishSitePage({ siteId: SITE_ID, pageId: PAGE_ID });
  const body = parsePayload(res);
  assert.equal(body.page.publishingState.level, "published");
  assert.equal(calls.length, 2);
  assert.ok(calls[0].endpoint.endsWith("/microsoft.graph.sitePage/publish"));
  assert.ok(calls[1].endpoint.endsWith("/microsoft.graph.sitePage"));
});

test("delete_site_page uses the un-cast endpoint", async () => {
  let captured = "";
  const mock = createMockGraphClient({
    delete: async (endpoint: string) => {
      captured = endpoint;
      return { success: true };
    },
  });
  installMockClient(mock);

  await handleDeleteSitePage({ siteId: SITE_ID, pageId: PAGE_ID });
  // delete is inherited from baseSitePage; the type-cast segment is omitted.
  assert.equal(captured, `/sites/${SITE_ID}/pages/${PAGE_ID}`);
});
