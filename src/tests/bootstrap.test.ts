import test from "node:test";
import assert from "node:assert/strict";

import {
  __resetBootstrapStateForTests,
  __setBootstrapDependenciesForTests,
  bootstrap,
} from "../core/bootstrap.js";
import type { ServerConfig } from "../config/index.js";

function buildConfig(overrides: {
  clientSecret?: string;
  tenantId?: string;
}): ServerConfig {
  return {
    auth: {
      clientId: "11111111-1111-1111-1111-111111111111",
      tenantId: overrides.tenantId ?? "common",
      scopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
      clientSecret: overrides.clientSecret,
    },
    graph: {
      baseUrl: "https://graph.microsoft.com/v1.0",
      timeout: 30000,
      maxRetries: 3,
    },
    cache: { enabled: true, ttl: 3600 },
  };
}

test.afterEach(() => {
  __resetBootstrapStateForTests();
});

test("bootstrap completes in client-credentials mode without a cached user", async () => {
  let getCurrentUserCalls = 0;
  let graphClientCreated = false;

  __setBootstrapDependenciesForTests({
    loadConfig: () => buildConfig({ clientSecret: "test-secret", tenantId: "22222222-2222-2222-2222-222222222222" }),
    initializeAuth: () => ({}) as never,
    getAuthInstance: () => ({
      getCurrentUser: async () => {
        getCurrentUserCalls++;
        return null;
      },
      getAuthMode: () => "client_credentials" as const,
      prewarm: () => undefined,
    }),
    getGraphClient: () => {
      graphClientCreated = true;
      return {} as never;
    },
  });

  await assert.doesNotReject(bootstrap());
  assert.equal(
    getCurrentUserCalls,
    0,
    "getCurrentUser must not be called in client-credentials mode",
  );
  assert.equal(graphClientCreated, true);
});

test("bootstrap rejects in device-code mode when no user is cached", async () => {
  __setBootstrapDependenciesForTests({
    loadConfig: () => buildConfig({}),
    initializeAuth: () => ({}) as never,
    getAuthInstance: () => ({
      getCurrentUser: async () => null,
      getAuthMode: () => "device_code" as const,
      prewarm: () => undefined,
    }),
    getGraphClient: () => ({}) as never,
  });

  await assert.rejects(bootstrap(), /Authentication required/);
});

test("bootstrap completes in device-code mode when a cached user is present", async () => {
  __setBootstrapDependenciesForTests({
    loadConfig: () => buildConfig({}),
    initializeAuth: () => ({}) as never,
    getAuthInstance: () => ({
      getCurrentUser: async () => ({ username: "user@example.com" }),
      getAuthMode: () => "device_code" as const,
      prewarm: () => undefined,
    }),
    getGraphClient: () => ({}) as never,
  });

  await assert.doesNotReject(bootstrap());
});
