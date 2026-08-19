// Worker wiring for the daily health check. The checks themselves live in
// core/health.js, which knows nothing about Workers and is unit-tested from
// the Node harness with stubbed dependencies; this module supplies the real
// ones — the KV store, the KV-backed Graph token, and the forced
// refresh-token rotation through the same exchange every Graph call uses.
import { runHealthCheck, type HealthReport } from "../core/health.js";
import { runWithTokenProvider } from "../core/token.js";
import type { Env } from "./env.js";
import { mailboxTokenProvider, refreshMailboxToken } from "./ms-token.js";
import { kvStateStore } from "./state-kv.js";

/** Run every health check against the real KV, Graph and token machinery. */
export async function runWorkerHealthCheck(env: Env): Promise<HealthReport> {
  return runWithTokenProvider(mailboxTokenProvider(env), () =>
    runHealthCheck({
      store: kvStateStore(env),
      // A real rotation, not a cache read: refreshMailboxToken always performs
      // the exchange and persists the rotated refresh token before returning.
      refreshToken: async () => {
        await refreshMailboxToken(env);
      },
    })
  );
}
