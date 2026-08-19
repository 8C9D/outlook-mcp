// Worker wiring for the two LLM features. The logic itself lives in
// core/classifier.js and core/digest.js, which know nothing about Workers and
// are unit-tested from the Node harness against a memory store.
//
// Both run with the KV-backed mailbox token in scope, exactly as an MCP request
// does, and both are strictly background work: a failure logs and is dropped
// rather than costing Graph its 202 or the cron its tick.
import { classifyAndFile } from "../core/classifier.js";
import { runDailyDigest, type DigestOutcome } from "../core/digest.js";
import { graphDigestMailbox } from "../core/digest-mailbox.js";
import { graphClassifierMailbox } from "../core/mail-actions.js";
import { readLlmConfig, torontoDateOf } from "../core/auto-filing.js";
import type { ActivityEntry } from "../core/notifications.js";
import { runWithTokenProvider } from "../core/token.js";
import type { Env } from "./env.js";
import { mailboxTokenProvider } from "./ms-token.js";
import { kvStateStore } from "./state-kv.js";

/** At most this many messages from one delivery are classified. */
const CLASSIFY_PER_DELIVERY = 5;

/**
 * Classify the messages one notification delivery brought in. Called from
 * ctx.waitUntil, so it runs after Graph already has its 202.
 *
 * The enabled check happens here as well as inside classifyAndFile so that a
 * disabled server does not even resolve a Graph token for a notification.
 */
export async function classifyNotified(env: Env, entries: ActivityEntry[]): Promise<void> {
  const store = kvStateStore(env);
  const config = await readLlmConfig(store);
  if (!config.filingEnabled) return;

  const messageIds = entries
    .map((entry) => entry.messageId)
    .filter((id): id is string => typeof id === "string" && id !== "")
    .slice(0, CLASSIFY_PER_DELIVERY);
  if (messageIds.length === 0) return;

  if (!env.ANTHROPIC_API_KEY) {
    console.error("Auto-filing is enabled but ANTHROPIC_API_KEY is not set; nothing classified.");
    return;
  }

  const today = torontoDateOf(new Date());
  await runWithTokenProvider(mailboxTokenProvider(env), async () => {
    const mailbox = graphClassifierMailbox();
    for (const messageId of messageIds) {
      const outcome = await classifyAndFile(messageId, {
        store,
        mailbox,
        apiKey: env.ANTHROPIC_API_KEY!,
        today,
      });
      console.log(`Auto-filing ${messageId}: ${outcome.action} — ${outcome.reason}`);
    }
  });
}

/**
 * Assemble and draft the morning brief. Called from the cron, and only after
 * the caller has checked that America/Toronto really is on the digest hour.
 */
export async function draftMorningBrief(env: Env, force = false): Promise<DigestOutcome> {
  const store = kvStateStore(env);
  const config = await readLlmConfig(store);
  if (!config.digestEnabled) {
    return { drafted: false, reason: "the morning digest is disabled" };
  }
  if (!env.ANTHROPIC_API_KEY) {
    console.error("The morning digest is enabled but ANTHROPIC_API_KEY is not set.");
    return { drafted: false, reason: "no ANTHROPIC_API_KEY is configured on this server" };
  }

  return runWithTokenProvider(mailboxTokenProvider(env), () =>
    runDailyDigest({
      store,
      mailbox: graphDigestMailbox(),
      apiKey: env.ANTHROPIC_API_KEY!,
      today: torontoDateOf(new Date()),
      force,
    })
  );
}
