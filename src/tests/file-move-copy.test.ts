import test from "node:test";
import assert from "node:assert/strict";

import { __setGraphClientInstanceForTests } from "../graph/client.js";
import {
  __resetKnownSitesForTests,
  __setKnownSitesForTests,
} from "../sharepoint/site-resolver.js";
import { handleCopyItem, handleMoveItem } from "../tools/files/index.js";
import { registerGraphClientTestLifecycle } from "./helpers/test-lifecycle.js";
import {
  createMockGraphClient,
  parsePayload,
  type ToolEnvelope,
} from "./helpers/tool-test-helpers.js";

registerGraphClientTestLifecycle();

const SOURCE_SITE_ID =
  "example.sharepoint.com,00000000-0000-0000-0000-000000000010,11111111-1111-1111-1111-111111111111";
const SOURCE_DRIVE_ID = "b!TESTSOURCEDRIVE";
const DEST_SITE_ID =
  "example.sharepoint.com,00000000-0000-0000-0000-000000000020,22222222-2222-2222-2222-222222222222";
const DEST_DRIVE_ID = "b!TESTDESTDRIVE";

test.beforeEach(() => {
  __setKnownSitesForTests([
    {
      key: "source",
      name: "Source",
      siteId: SOURCE_SITE_ID,
      siteUrl: "https://example.sharepoint.com/sites/Source",
      driveId: SOURCE_DRIVE_ID,
      aliases: ["source"],
    },
    {
      key: "dest",
      name: "Dest",
      siteId: DEST_SITE_ID,
      siteUrl: "https://example.sharepoint.com/sites/Dest",
      driveId: DEST_DRIVE_ID,
      aliases: ["dest"],
    },
  ]);
});

test.afterEach(() => {
  __resetKnownSitesForTests();
});

test("copy_item resolves destinationSite alias to a real driveId on parentReference", async () => {
  const mock = createMockGraphClient({
    post: async () => ({ success: true, data: { id: "new-item" } }),
  });
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleCopyItem({
    itemId: "item-1",
    site: "source",
    destinationSite: "dest",
    newName: "copy.txt",
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);

  const postCall = mock.methodCalls("post")[0];
  assert.ok(postCall, "expected a /copy POST call");
  assert.equal(
    postCall.args[0],
    `/drives/${SOURCE_DRIVE_ID}/items/item-1/copy`,
  );
  // The critical fix: parentReference.driveId is the destination *drive* id,
  // NOT the destinationSiteId.
  assert.equal(postCall.args[1].parentReference.driveId, DEST_DRIVE_ID);
  assert.notEqual(postCall.args[1].parentReference.driveId, DEST_SITE_ID);
  assert.equal(payload.destinationDriveId, DEST_DRIVE_ID);
});

test("copy_item resolves destinationSiteId via Graph when no registry entry exists", async () => {
  const lookupCalls: string[] = [];
  const mock = createMockGraphClient({
    get: async (endpoint: string) => {
      lookupCalls.push(endpoint);
      if (endpoint === `/sites/unknown-site-id/drive`) {
        return {
          success: true,
          data: { id: "resolved-drive-from-graph" },
        };
      }
      throw new Error(`unexpected GET ${endpoint}`);
    },
    post: async () => ({ success: true, data: { id: "new-item" } }),
  });
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleCopyItem({
    itemId: "item-1",
    driveId: SOURCE_DRIVE_ID,
    destinationSiteId: "unknown-site-id",
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.deepEqual(lookupCalls, ["/sites/unknown-site-id/drive"]);
  const postCall = mock.methodCalls("post")[0];
  assert.equal(
    postCall.args[1].parentReference.driveId,
    "resolved-drive-from-graph",
  );
  assert.equal(payload.destinationDriveId, "resolved-drive-from-graph");
});

test("copy_item resolves destinationFolderPath relative to the destination drive", async () => {
  const seenGetEndpoints: string[] = [];
  const mock = createMockGraphClient({
    get: async (endpoint: string) => {
      seenGetEndpoints.push(endpoint);
      return {
        success: true,
        data: { id: "folder-on-dest-drive" },
      };
    },
    post: async () => ({ success: true, data: { id: "new-item" } }),
  });
  __setGraphClientInstanceForTests(mock.client as any);

  await handleCopyItem({
    itemId: "item-1",
    site: "source",
    destinationSite: "dest",
    destinationFolderPath: "Archive/2026",
  });

  // The folder must be resolved against the destination drive, not the source.
  assert.ok(
    seenGetEndpoints.some(
      (e) => e === `/drives/${DEST_DRIVE_ID}/root:/Archive/2026`,
    ),
    `expected dest-drive folder lookup, got ${JSON.stringify(seenGetEndpoints)}`,
  );

  const postCall = mock.methodCalls("post")[0];
  assert.equal(
    postCall.args[1].parentReference.driveId,
    DEST_DRIVE_ID,
  );
  assert.equal(
    postCall.args[1].parentReference.id,
    "folder-on-dest-drive",
  );
});

test("copy_item errors loudly when the destination folder lookup fails", async () => {
  const mock = createMockGraphClient({
    get: async () => {
      throw new Error("404 itemNotFound");
    },
    post: async () => ({ success: true }),
  });
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleCopyItem({
    itemId: "item-1",
    site: "source",
    destinationFolderPath: "missing/folder",
  })) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.equal(
    mock.methodCalls("post").length,
    0,
    "/copy must not be POSTed when destination folder lookup fails",
  );
});

test("copy_item works in same-drive mode without any destination references", async () => {
  const mock = createMockGraphClient({
    post: async () => ({ success: true, data: { id: "new-item" } }),
  });
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleCopyItem({
    itemId: "item-1",
    site: "source",
    newName: "renamed.txt",
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  const postCall = mock.methodCalls("post")[0];
  // Same-drive copy: no parentReference.driveId, just a new name.
  assert.equal(postCall.args[1].parentReference.driveId, undefined);
  assert.equal(postCall.args[1].name, "renamed.txt");
  assert.equal(payload.destinationDriveId, null);
});

test("move_item errors when parentFolderPath cannot be resolved", async () => {
  const mock = createMockGraphClient({
    get: async () => ({ success: false }),
    patch: async () => ({ success: true, data: { id: "moved" } }),
  });
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleMoveItem({
    itemId: "item-1",
    site: "source",
    parentFolderPath: "nope/missing",
  })) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.match(response.content[0].text, /Parent folder not found/);
  assert.equal(
    mock.methodCalls("patch").length,
    0,
    "PATCH must not run when the destination folder lookup fails",
  );
});

test("move_item rejects no-op calls instead of pretending to move", async () => {
  const mock = createMockGraphClient({
    patch: async () => ({ success: true, data: { id: "moved" } }),
  });
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleMoveItem({
    itemId: "item-1",
    site: "source",
  })) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.match(response.content[0].text, /at least one of: newName/);
  assert.equal(mock.methodCalls("patch").length, 0);
});

test("move_item happy path: rename via newName", async () => {
  const mock = createMockGraphClient({
    patch: async (_endpoint: string, body: any) => ({
      success: true,
      data: {
        id: "item-1",
        name: body.name,
        webUrl: "https://example.invalid",
        parentReference: { path: "/drive/root:/Docs" },
      },
    }),
  });
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleMoveItem({
    itemId: "item-1",
    site: "source",
    newName: "renamed.txt",
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.item.name, "renamed.txt");
  const patchCall = mock.methodCalls("patch")[0];
  assert.equal(patchCall.args[1].name, "renamed.txt");
  assert.equal(patchCall.args[1].parentReference, undefined);
});
