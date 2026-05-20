import { loadConfig, type ServerConfig } from "../config/index.js";
import {
  getAuthInstance,
  initializeAuth,
  MicrosoftGraphAuth,
} from "../auth/microsoft-graph-auth.js";
import { getGraphClient } from "../graph/client.js";

type BootstrapDependencies = {
  loadConfig: () => ServerConfig;
  initializeAuth: typeof initializeAuth;
  getAuthInstance: () => Pick<
    MicrosoftGraphAuth,
    "getCurrentUser" | "getAuthMode" | "prewarm"
  >;
  getGraphClient: typeof getGraphClient;
};

const defaultDependencies: BootstrapDependencies = {
  loadConfig,
  initializeAuth,
  getAuthInstance,
  getGraphClient,
};

let dependencies: BootstrapDependencies = { ...defaultDependencies };

let initialized = false;
let initPromise: Promise<void> | null = null;

export async function bootstrap(): Promise<void> {
  if (initialized) return;
  if (initPromise) return initPromise;

  initPromise = (async () => {
    const config = dependencies.loadConfig();
    dependencies.initializeAuth(config.auth);

    const auth = dependencies.getAuthInstance();

    // Device-code mode requires a pre-existing cached user (from setup-auth).
    // Client-credentials mode has no user — the token is acquired on first
    // use via getAccessToken(), so we skip the cache check entirely.
    if (!config.auth.clientSecret) {
      const cachedUser = await auth.getCurrentUser();
      if (!cachedUser) {
        throw new Error(
          "Authentication required. Run `npm run setup-auth` before calling tools, or set MICROSOFT_GRAPH_CLIENT_SECRET for client-credentials auth.",
        );
      }
    }

    dependencies.getGraphClient();
    initialized = true;
  })();

  try {
    await initPromise;
  } catch (err) {
    initPromise = null;
    throw err;
  }
}

export function prewarmAuth(): void {
  try {
    const config = dependencies.loadConfig();
    dependencies.initializeAuth(config.auth);
    const auth = dependencies.getAuthInstance() as unknown as {
      prewarm?: () => void;
    };
    if (typeof auth.prewarm === "function") auth.prewarm();
  } catch (err) {
    // Non-fatal: bootstrap() will raise the real failure on first tool call.
    // Still log so missing config / broken keychain does not vanish.
    const message = err instanceof Error ? err.message : String(err);
    console.error(`[prewarmAuth] skipped: ${message}`);
  }
}

export function __setBootstrapDependenciesForTests(
  overrides?: Partial<BootstrapDependencies>,
): void {
  dependencies = { ...defaultDependencies, ...overrides };
}

export function __resetBootstrapStateForTests(): void {
  initialized = false;
  initPromise = null;
  dependencies = { ...defaultDependencies };
}
