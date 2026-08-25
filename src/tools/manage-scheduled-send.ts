// Scheduled (delayed-delivery) mail: what send_draft parked with send_at, and
// the way to stop it before it goes.
//
// Mechanics, verified live on this consumer mailbox: a deferred send leaves the
// message IN THE DRAFTS FOLDER (still isDraft, same id) with the MAPI deferred-
// send-time property on it; at that instant Exchange sends it and moves it to
// Sent Items. Deleting the parked message before the instant cancels the send —
// and, unlike every other mailbox delete, the message VANISHES outright (no
// Deleted Items copy; the submitted state is discarded with it), so cancel here
// is a permanent discard and both the description and the output say so.
import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import {
  ToolResult,
  errorResult,
  fetchPaged,
  formatLocal,
  runTool,
  textResult,
} from "./common.js";
import { DEFERRED_SEND_PROPERTY } from "./send-draft.js";

export const manageScheduledSendSchema = {
  action: z
    .enum(["list", "cancel"])
    .describe(
      "list: every message currently waiting in Drafts for a scheduled send, with its send time; cancel: stop message_id from sending — the message is DISCARDED outright (see the tool description)."
    ),
  message_id: z
    .string()
    .min(1)
    .optional()
    .describe("cancel: the waiting message's id (from list, or from send_draft's answer)."),
};

const manageScheduledSendArgs = z.object(manageScheduledSendSchema);

export const manageScheduledSendDescription =
  "List or cancel scheduled sends (mail parked by send_draft with send_at, waiting in Drafts for its send time). CAUTION on cancel: verified live, deleting a scheduled message discards it COMPLETELY — the send is stopped, but no copy lands in Deleted Items and the content cannot be recovered, so state the message's subject, recipients and scheduled time to the user and get agreement first; to keep the text, read it with read_message and re-create the draft before cancelling. list is read-only.";

const EXPAND = `singleValueExtendedProperties($filter=id eq '${DEFERRED_SEND_PROPERTY}')`;

/** The deferred send time on a message, or undefined when it carries none. */
function deferredTimeOf(message: any): string | undefined {
  const value = (message.singleValueExtendedProperties ?? []).find(
    (p: any) => String(p.id ?? "").toLowerCase() === DEFERRED_SEND_PROPERTY.toLowerCase()
  )?.value;
  return typeof value === "string" && value ? value : undefined;
}

function describeScheduled(message: any, deferredUtc: string): string {
  const to = (message.toRecipients ?? [])
    .map((r: any) => r.emailAddress?.address)
    .filter(Boolean)
    .join(", ");
  return (
    `${message.subject || "(no subject)"}\n` +
    `  To: ${to || "(none)"}\n` +
    `  Sends at: ${formatLocal(deferredUtc)} (America/Toronto)\n` +
    `  Message id: ${message.id}`
  );
}

export async function manageScheduledSendHandler(
  input: z.input<typeof manageScheduledSendArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, message_id } = manageScheduledSendArgs.parse(input);

    if (action === "list") {
      // Drafts is small; expand the property on each and keep the ones that
      // carry it. (A $filter on the property would also need the $expand
      // anyway, and client-side matching keeps the query simple.)
      const drafts = await fetchPaged(
        `/me/mailFolders/drafts/messages?$select=id,subject,toRecipients&$expand=${encodeURIComponent(EXPAND)}&$top=50`,
        200
      );
      const scheduled = drafts
        .map((m) => ({ message: m, deferredUtc: deferredTimeOf(m) }))
        .filter((entry): entry is { message: any; deferredUtc: string } => !!entry.deferredUtc)
        .sort((a, b) => a.deferredUtc.localeCompare(b.deferredUtc));
      if (scheduled.length === 0) {
        return textResult(
          "No scheduled sends are waiting. (send_draft with send_at creates one; it waits in Drafts until its time.)"
        );
      }
      return textResult(
        `${scheduled.length} scheduled send(s) waiting in Drafts:\n\n` +
          scheduled.map((entry) => describeScheduled(entry.message, entry.deferredUtc)).join("\n\n") +
          "\n\nEach sends by itself at its time; cancel stops one (discarding the message — see the tool description)."
      );
    }

    if (!message_id) return errorResult('Action "cancel" requires message_id (from list).');
    const message = await callGraphServer(
      `/me/messages/${encodeURIComponent(message_id)}?$select=id,subject,toRecipients,isDraft&$expand=${encodeURIComponent(EXPAND)}`
    );
    const deferredUtc = deferredTimeOf(message);
    if (!message.isDraft || !deferredUtc) {
      return errorResult(
        `Message ${message_id} is not a waiting scheduled send${
          message.isDraft ? " (it is an ordinary draft with no send time)" : ""
        } — only messages listed by manage_scheduled_send list can be cancelled here.`
      );
    }
    if (Date.parse(deferredUtc) <= Date.now()) {
      return errorResult(
        `Too late — this message's send time (${formatLocal(deferredUtc)}) has already passed, so it is sending (or has sent) and cannot be cancelled.`
      );
    }

    const description = describeScheduled(message, deferredUtc);
    await callGraphServer(`/me/messages/${encodeURIComponent(message_id)}`, { method: "DELETE" });
    return textResult(
      `Scheduled send cancelled — nothing will be sent.\n\n${description}\n\n` +
        "The message itself was discarded with the schedule (no Deleted Items copy — that is how " +
        "Exchange drops a submitted deferred message). To send it later after all, compose it " +
        "again with create_draft."
    );
  });
}
