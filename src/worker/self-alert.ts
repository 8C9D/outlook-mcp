// Worker wiring for the opt-in self-alert route and its heartbeat watchdog. The
// logic — the secret gate, body validation, the fixed recipient, the daily cap,
// staleness — lives in core/self-alert.js and is tested offline; this module
// supplies only the real KV namespace and the one Graph call.
//
// The routes are dispatched in index.ts BEFORE OAuthProvider: they are not
// OAuth-protected, and authenticate with SELF_ALERT_SECRET instead.
import { callGraphServer } from "../core/graph.js";
import {
  handleSelfAlertRequest,
  runSelfAlertWatchdog,
  type SelfAlertDeps,
  type SendMailPayload,
  type WatchdogResult,
} from "../core/self-alert.js";
import { runWithTokenProvider } from "../core/token.js";
import type { Env } from "./env.js";
import { mailboxTokenProvider } from "./ms-token.js";

export { isSelfAlertPath } from "../core/self-alert.js";

function selfAlertDeps(env: Env): SelfAlertDeps {
  return {
    secret: env.SELF_ALERT_SECRET,
    recipient: env.ALLOWED_MS_UPN,
    kv: env.OUTLOOK_KV,
    // The payload arrives fully built by buildSelfAlertMail, addressed to
    // ALLOWED_MS_UPN alone; nothing is added to it here. The token comes from
    // the same KV-backed provider the health check uses.
    sendMail: (payload: SendMailPayload) =>
      runWithTokenProvider(mailboxTokenProvider(env), async () => {
        await callGraphServer("/me/sendMail", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
      }),
  };
}

/** POST /self-alert and POST /self-alert/heartbeat. */
export function handleSelfAlert(request: Request, env: Env): Promise<Response> {
  return handleSelfAlertRequest(request, selfAlertDeps(env));
}

/** One watchdog pass, run on every UPKEEP_CRON tick (core/schedule.js). */
export function runWorkerSelfAlertWatchdog(env: Env): Promise<WatchdogResult> {
  return runSelfAlertWatchdog(selfAlertDeps(env));
}
