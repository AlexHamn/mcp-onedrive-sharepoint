/**
 * Long-running-operation polling helper.
 *
 * Microsoft Graph signals async work in two flavors:
 *   1. A dedicated `richLongRunningOperation` endpoint (e.g. site creation:
 *      `POST /beta/sites` → `Location: …/getOperationStatus(operationId='…')`).
 *   2. Implicit polling of a target resource until it appears (e.g. group
 *      creation: site for the new group is reachable at
 *      `GET /groups/{id}/sites/root` once provisioning finishes).
 *
 * Both fit the same shape: call a probe repeatedly until it reports done.
 * This module provides the generic loop; the site-provisioning code adapts
 * each Graph endpoint to it.
 */

export interface PollUntilOptions {
  /** Delay between probes, in ms. Defaults to 2000. Capped by `maxBackoffMs`. */
  intervalMs?: number;
  /** Backoff multiplier applied to `intervalMs` after each unsuccessful probe. Defaults to 1.5. */
  backoffMultiplier?: number;
  /** Upper bound on the per-probe delay. Defaults to 10000 (10s). */
  maxBackoffMs?: number;
  /** Total budget for the whole polling loop. Defaults to 120000 (2min). */
  timeoutMs?: number;
  /** Delay before the very first probe. Defaults to 0 (probe immediately). */
  initialDelayMs?: number;
  /**
   * Override the sleep implementation. Tests pass `async () => {}` so the
   * loop runs instantly without monkey-patching globals. Real callers leave
   * this undefined and get setTimeout-based sleep.
   */
  sleep?: (ms: number) => Promise<void>;
}

export interface PollResult<T> {
  value: T;
  attempts: number;
  elapsedMs: number;
}

export class PollingTimeoutError extends Error {
  readonly attempts: number;
  readonly elapsedMs: number;
  readonly lastValue?: unknown;

  constructor(
    context: string,
    attempts: number,
    elapsedMs: number,
    lastValue?: unknown,
  ) {
    super(
      `Polling for ${context} timed out after ${elapsedMs}ms (${attempts} attempts)`,
    );
    this.name = "PollingTimeoutError";
    this.attempts = attempts;
    this.elapsedMs = elapsedMs;
    this.lastValue = lastValue;
  }
}

function defaultSleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Repeatedly invoke `probe` until it returns `{ done: true, value }` or the
 * timeout budget is exhausted. The probe is responsible for distinguishing
 * "still working" (return `done: false`) from "failed permanently" (throw).
 *
 * `context` is purely cosmetic — it appears in timeout error messages.
 */
export async function pollUntil<T>(
  context: string,
  probe: () => Promise<{ done: boolean; value: T }>,
  options: PollUntilOptions = {},
): Promise<PollResult<T>> {
  const intervalMs = options.intervalMs ?? 2000;
  const backoffMultiplier = options.backoffMultiplier ?? 1.5;
  const maxBackoffMs = options.maxBackoffMs ?? 10000;
  const timeoutMs = options.timeoutMs ?? 120000;
  const initialDelayMs = options.initialDelayMs ?? 0;
  const sleep = options.sleep ?? defaultSleep;

  const start = Date.now();
  let attempts = 0;
  let lastValue: T | undefined;

  if (initialDelayMs > 0) {
    await sleep(initialDelayMs);
  }

  while (true) {
    attempts++;
    const result = await probe();
    lastValue = result.value;
    if (result.done) {
      return { value: result.value, attempts, elapsedMs: Date.now() - start };
    }

    const elapsed = Date.now() - start;
    if (elapsed >= timeoutMs) {
      throw new PollingTimeoutError(context, attempts, elapsed, lastValue);
    }

    const rawDelay = intervalMs * Math.pow(backoffMultiplier, attempts - 1);
    const remaining = timeoutMs - elapsed;
    const delay = Math.min(rawDelay, maxBackoffMs, Math.max(remaining, 0));
    if (delay <= 0) {
      throw new PollingTimeoutError(context, attempts, Date.now() - start, lastValue);
    }
    await sleep(delay);
  }
}
