import { z } from "zod";
import { callGraphServer } from "../graph.js";
import { ToolResult, errorResult, runTool, textResult, toRecipients } from "./common.js";

// NOTE: Sending is two-step by structure. This tool only composes drafts;
// the sole send path in this codebase is send_draft (POST /messages/{id}/send).
// Nothing here may ever call /me/sendMail.

export const createDraftSchema = {
  reply_to_message_id: z
    .string()
    .optional()
    .describe(
      "Message id to draft a reply to (from search_mail or read_message). REPLY MODE: provide this and leave 'to', 'subject', and 'forward_message_id' unset — the reply's recipients and subject come from the original message."
    ),
  reply_all: z
    .boolean()
    .default(false)
    .describe(
      "Reply to all original recipients instead of just the sender. Only valid in reply mode (with reply_to_message_id)."
    ),
  forward_message_id: z
    .string()
    .optional()
    .describe(
      "Message id to draft a forward of. FORWARD MODE: provide this plus 'to' (the forward's recipients); leave 'subject' and 'reply_to_message_id' unset — the subject comes from the original message."
    ),
  to: z
    .array(z.string().email())
    .optional()
    .describe(
      "Recipient email addresses. Required in NEW-MESSAGE MODE (together with 'subject') and used in FORWARD MODE; not allowed in reply mode. Exactly one mode must be used."
    ),
  subject: z
    .string()
    .optional()
    .describe("Subject line (new-message mode only; required with 'to')."),
  body: z
    .string()
    .min(1)
    .describe(
      "Message body, treated as plain text. In reply and forward modes it is placed above the quoted original."
    ),
  cc: z.array(z.string().email()).optional().describe("Optional CC email addresses."),
};

const createDraftArgs = z.object(createDraftSchema);

export const createDraftDescription =
  "Create an email draft in the Outlook Drafts folder — this tool never sends; sending requires a separate send_draft call. Three mutually exclusive modes: reply mode (reply_to_message_id, optionally reply_all) drafts a reply with your text above the quoted original; forward mode (forward_message_id + to) drafts a forward; new-message mode (to + subject) drafts a fresh message. body is required in every mode.";

export async function createDraftHandler(
  input: z.input<typeof createDraftArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { reply_to_message_id, reply_all, forward_message_id, to, subject, body, cc } =
      createDraftArgs.parse(input);

    const replyMode = reply_to_message_id !== undefined;
    const forwardMode = forward_message_id !== undefined;
    const newFieldsPresent = to !== undefined || subject !== undefined;
    if (replyMode && forwardMode) {
      return errorResult(
        "reply_to_message_id and forward_message_id are mutually exclusive — pick one mode."
      );
    }
    if (replyMode && newFieldsPresent) {
      return errorResult(
        "Reply mode takes its recipients and subject from the original message — leave 'to' and 'subject' unset."
      );
    }
    if (forwardMode && subject !== undefined) {
      return errorResult(
        "Forward mode takes its subject from the original message — leave 'subject' unset."
      );
    }
    if (reply_all && !replyMode) {
      return errorResult("reply_all is only valid in reply mode (with reply_to_message_id).");
    }
    if (!replyMode && !forwardMode && (!to?.length || !subject)) {
      return errorResult(
        "Exactly one mode required: reply (reply_to_message_id), forward (forward_message_id), or new-message (to + subject). New-message mode needs both a non-empty 'to' list and a 'subject'."
      );
    }

    let draft: any;
    let modeLabel = "";
    if (replyMode || forwardMode) {
      modeLabel = replyMode ? (reply_all ? " (reply all)" : " (reply)") : " (forward)";
      const sourceId = replyMode ? reply_to_message_id! : forward_message_id!;
      const action = replyMode ? (reply_all ? "createReplyAll" : "createReply") : "createForward";
      const created = await callGraphServer(
        `/me/messages/${encodeURIComponent(sourceId)}/${action}`,
        { method: "POST" }
      );
      // Fetch the auto-generated quoted body as text, then prepend the new text above it.
      const existing = await callGraphServer(
        `/me/messages/${created.id}?$select=body,subject,toRecipients`,
        { headers: { Prefer: 'outlook.body-content-type="text"' } }
      );
      const quoted = (existing.body?.content ?? "").replace(/\r\n/g, "\n");
      const patch: any = {
        body: { contentType: "Text", content: `${body}\n\n${quoted}`.trimEnd() + "\n" },
      };
      if (forwardMode && to?.length) patch.toRecipients = toRecipients(to);
      if (cc?.length) {
        const existingCc = created.ccRecipients ?? [];
        patch.ccRecipients = [...existingCc, ...toRecipients(cc)];
      }
      draft = await callGraphServer(`/me/messages/${created.id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(patch),
      });
    } else {
      draft = await callGraphServer("/me/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          subject,
          body: { contentType: "Text", content: body },
          toRecipients: toRecipients(to!),
          ...(cc?.length ? { ccRecipients: toRecipients(cc) } : {}),
        }),
      });
    }

    const recipients = (draft.toRecipients ?? [])
      .map((r: any) => r.emailAddress?.address)
      .filter(Boolean)
      .join(", ");
    const ccList = (draft.ccRecipients ?? [])
      .map((r: any) => r.emailAddress?.address)
      .filter(Boolean)
      .join(", ");
    return textResult(
      `Draft created${modeLabel}.\n` +
        `Subject: ${draft.subject || "(no subject)"}\n` +
        `To: ${recipients || "(none)"}\n` +
        (ccList ? `Cc: ${ccList}\n` : "") +
        `Draft id: ${draft.id}\n` +
        "Saved to Drafts — not sent. Use update_draft to revise, or send_draft to send it."
    );
  });
}
