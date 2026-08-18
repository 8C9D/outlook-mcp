// Junk / blocked senders.
//
// What Microsoft Graph actually offers a consumer (outlook.com) mailbox here is
// far less than the Outlook web UI suggests, and every alternative was probed
// live against this account before this tool was written (see ASSUMPTIONS.md):
//
//   GET  /beta/me/blockedSenders                    404 UnknownError
//   GET  /beta/me/safeSenders                       404 UnknownError
//   GET  /beta/me/outlook/blockedSenders            400 "Resource not found for the segment"
//   GET  /beta/me/mailboxSettings?$select=blockedSenders
//                                                   400 "Could not find a property named
//                                                        'blockedSenders' on type mailboxSettings"
//   GET  /beta/me/mailboxSettings/junkMailRule      400 "Resource not found for the segment"
//   GET  https://outlook.office.com/api/beta/me/blockedsenders
//                                                   401 (a different resource audience — would
//                                                        need consent this connector does not have)
//
// The one surface that does work is per-message and beta-only:
//   POST /beta/me/messages/{id}/markAsJunk    202 Accepted  (blocks that message's sender)
//   POST /beta/me/messages/{id}/markAsNotJunk 202 Accepted  (unblocks it again)
//
// So: blocking and unblocking are message-scoped, and the resulting lists cannot
// be read back through Graph at all. The tool says so rather than pretending.
import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import {
  ToolResult,
  errorResult,
  formatSender,
  isNotFound,
  runTool,
  textResult,
} from "./common.js";

const GRAPH_BETA = "https://graph.microsoft.com/beta";

export const manageSendersSchema = {
  action: z
    .enum(["block_sender", "unblock_sender"])
    .describe(
      "block_sender: add the sender of message_id to the account's blocked-senders list (their future mail goes straight to Junk); unblock_sender: take that sender off the blocked list again."
    ),
  message_id: z
    .string()
    .min(1)
    .describe(
      "A message FROM the sender to block or unblock (from search_mail or read_message). Microsoft Graph exposes no way to block a bare email address, so this always works from one of the sender's messages."
    ),
  move_message: z
    .boolean()
    .default(true)
    .describe(
      "Also move that message: to Junk Email on block_sender, back to the Inbox on unblock_sender (default true). Set false to change the sender's status while leaving the message where it is."
    ),
};

const manageSendersArgs = z.object(manageSendersSchema);

export const manageSendersDescription =
  "Block or unblock the sender of a message — the mailbox's blocked-senders list, the same one Outlook's \"Block sender\" / \"Not junk\" buttons drive. CAUTION: blocking is account-wide and lasting — every future message from that address is filed as junk until it is unblocked — so name the address to the user and get agreement first. PLATFORM LIMITS on this consumer account, verified live against Microsoft Graph: (1) the blocked and safe sender LISTS CANNOT BE READ — Graph exposes no endpoint for them, so this tool cannot tell you who is currently blocked; check Outlook web (Settings → Mail → Junk email) for that; (2) safe senders cannot be managed at all through Graph; (3) blocking is per-message, not per-address: pass a message from the sender you mean. Because nothing can be read back, say plainly which message's sender you acted on so the user can verify it in Outlook.";

export async function manageSendersHandler(
  input: z.input<typeof manageSendersArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, message_id, move_message } = manageSendersArgs.parse(input);

    let message: any;
    try {
      message = await callGraphServer(
        `/me/messages/${encodeURIComponent(message_id)}?$select=id,subject,from,parentFolderId`
      );
    } catch (err) {
      if (isNotFound(err)) {
        return errorResult(
          `No message ${message_id} — pass the id of a message from the sender you want to ${action === "block_sender" ? "block" : "unblock"} (from search_mail or read_message).`
        );
      }
      throw err;
    }

    const sender = formatSender(message.from);
    if (!message.from?.emailAddress?.address) {
      return errorResult(
        "That message has no sender address (drafts do not), so there is nobody to block or unblock. " +
          "Use a message you received."
      );
    }

    const blocking = action === "block_sender";
    const graphAction = blocking ? "markAsJunk" : "markAsNotJunk";
    const body = blocking ? { moveToJunk: move_message } : { moveToInbox: move_message };

    // Beta-only: v1.0 has no markAsJunk/markAsNotJunk (it answers 400).
    await callGraphServer(
      `${GRAPH_BETA}/me/messages/${encodeURIComponent(message_id)}/${graphAction}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      }
    );

    const moved = move_message
      ? blocking
        ? "The message was moved to Junk Email."
        : "The message was moved back to the Inbox."
      : "The message was left where it was.";

    return textResult(
      `${blocking ? "Blocked" : "Unblocked"} sender: ${sender}\n` +
        `From message: ${message.subject || "(no subject)"}\n` +
        `${moved}\n` +
        (blocking
          ? "Future mail from that address is filed as junk until it is unblocked."
          : "Mail from that address is delivered normally again.") +
        "\nMicrosoft Graph cannot read the blocked-senders list back, so this cannot be verified " +
        "here — check Outlook web (Settings → Mail → Junk email) to see the list itself."
    );
  });
}
