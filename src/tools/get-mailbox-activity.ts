import { z } from "zod";
import { ACTIVITY_CAP, readActivity } from "../core/notifications.js";
import { STATE_SUBSCRIPTION } from "../core/kv-keys.js";
import { readJson, requireStateStore } from "../core/state.js";
import type { SubscriptionRecord } from "../core/subscriptions.js";
import { ToolResult, errorResult, formatLocal, runTool, textResult } from "./common.js";

export const getMailboxActivitySchema = {
  since_hours: z
    .number()
    .min(0.25)
    .max(168)
    .default(12)
    .describe(
      "Only report notifications received within this many hours (default 12, max 168). Use 6-12 for \"since this morning\"."
    ),
  limit: z
    .number()
    .int()
    .min(1)
    .max(ACTIVITY_CAP)
    .default(25)
    .describe(`Maximum number of entries to return (default 25, max ${ACTIVITY_CAP}).`),
};

const getMailboxActivityArgs = z.object(getMailboxActivitySchema);

export const getMailboxActivityDescription =
  "List mail that arrived recently, from change notifications Microsoft Graph pushed to this server as it happened — no mailbox polling. Answers \"what came in since this morning\" cheaply, newest first, with the time each notification was received. Only the last 50 notifications are kept and only the inbox is watched, so use search_mail or check_new_mail for anything older or for other folders. Available only on the hosted (remote) server; the local stdio server is not reachable from Microsoft and returns an error.";

export async function getMailboxActivityHandler(
  input: z.input<typeof getMailboxActivityArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { since_hours, limit } = getMailboxActivityArgs.parse(input);
    const store = requireStateStore();
    if (store.mode !== "remote") {
      return errorResult(
        "get_mailbox_activity works only on the hosted (remote) server. Microsoft Graph pushes " +
          "change notifications to a public URL, which the local stdio server does not have. " +
          "Use check_new_mail for the same question locally — it asks Graph what changed since " +
          "the last call."
      );
    }

    const subscription = await readJson<SubscriptionRecord>(store, STATE_SUBSCRIPTION);
    if (!subscription?.id) {
      return errorResult(
        "No change-notification subscription exists yet, so nothing is being recorded. It is " +
          "created automatically; try again in a few minutes."
      );
    }

    const cutoff = Date.now() - since_hours * 3600_000;
    const entries = (await readActivity(store))
      .filter((entry) => Date.parse(entry.at) >= cutoff)
      .slice(0, limit);

    const window = `the last ${since_hours} hour(s)`;
    if (entries.length === 0) {
      return textResult(
        `No mail notifications in ${window}. (Watching the inbox; subscription expires ${formatLocal(
          subscription.expirationDateTime
        )}.)`
      );
    }

    const lines = entries.map((entry, index) => {
      const parts = [
        `${index + 1}. ${entry.subject ?? "(subject unavailable)"}`,
        `   From: ${entry.from ?? "(unknown)"}  Notified: ${formatLocal(entry.at)}` +
          (entry.receivedDateTime ? `  Received: ${formatLocal(entry.receivedDateTime)}` : ""),
      ];
      if (entry.messageId) parts.push(`   Message id: ${entry.messageId}`);
      return parts.join("\n");
    });

    return textResult(
      `${entries.length} mail notification(s) in ${window} (newest first):\n\n${lines.join("\n\n")}`
    );
  });
}
