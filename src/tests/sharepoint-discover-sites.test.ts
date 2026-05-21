import test from "node:test";
import assert from "node:assert/strict";

import { __setGraphClientInstanceForTests } from "../graph/client.js";
import { handleDiscoverSites } from "../tools/sharepoint/index.js";
import { registerGraphClientTestLifecycle } from "./helpers/test-lifecycle.js";
import {
  createMockGraphClient,
  parsePayload,
  type ToolEnvelope,
} from "./helpers/tool-test-helpers.js";

registerGraphClientTestLifecycle();

test("discover_sites uses Graph site search endpoint for explicit searches", async () => {
  const mock = createMockGraphClient({
    get: async (endpoint: string) => {
      if (endpoint === "/sites/root") {
        return {
          success: true,
          data: {
            id: "root-site",
            displayName: "Root Site",
            name: "Root Site",
            webUrl: "https://contoso.sharepoint.com",
            root: {},
          },
        };
      }

      return {
        success: true,
        data: {
          value: [
            {
              id: "site-1",
              displayName: "Primary Workspace",
              name: "Primary Workspace",
              webUrl: "https://contoso.sharepoint.com/sites/primary",
            },
          ],
        },
      };
    },
  });

  __setGraphClientInstanceForTests(mock.client as any);

  // Preserve a non-ASCII character in the search term so the test continues
  // to exercise the UTF-8 URL-encoding code path (é → %C3%A9).
  const response = (await handleDiscoverSites({
    search: "  Café   Records  ",
    includePersonalSite: true,
    limit: 5,
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.search, "Café Records");
  assert.equal(payload.siteCount, 2);
  assert.equal(payload.sites[0].id, "root-site");
  assert.equal(
    mock.methodCalls("get")[0]?.args[0],
    "/sites?search=Caf%C3%A9%20Records",
  );
  assert.deepEqual(mock.methodCalls("get")[0]?.args[1], { $top: "5" });
});

test("discover_sites falls back to wildcard Graph search when no search term is provided", async () => {
  const mock = createMockGraphClient({
    get: async () => ({
      success: true,
      data: {
        value: [
          {
            id: "site-1",
            displayName: "Finance",
            name: "Finance",
            webUrl: "https://contoso.sharepoint.com/sites/finance",
          },
        ],
      },
    }),
  });

  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleDiscoverSites({ limit: 10 })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.search, "all sites");
  assert.equal(payload.siteCount, 1);
  assert.equal(mock.methodCalls("get")[0]?.args[0], "/sites?search=*");
  assert.deepEqual(mock.methodCalls("get")[0]?.args[1], { $top: "10" });
});
