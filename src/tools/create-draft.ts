import { z } from "zod";
import { callGraphServer } from "../graph.js";
import { ToolResult, errorResult, runTool, textResult, toRecipients } from "./common.js";

// NOTE: This server is draft-only by design. Nothing in this codebase may ever
// call /send or /me/sendMail; the Entra registration has no Mail.Send permission.

export const createDraftSchema = {
  reply_to_message_id: z
    .string()
    .optional()
    .describe(
      "Message id to draft a reply to (from search_mail). REPLY MODE: provide this and leave 'to' and 'subject' unset — the reply's recipients and subject come from the original message."
    ),
  to: z
    .array(z.string().email())
    .optional()
    .describe(
      "Recipient email addresses. NEW-MESSAGE MODE: provide this together with 'subject', and leave 'reply_to_message_id' unset. Exactly one mode must be used."
    ),
  subject: z
    .string()
    .optional()
    .describe("Subject line (new-message mode only; required with 'to')."),
  body: z
    .string()
    .min(1)
    .describe("Message body, treated as plain text. In reply mode it is placed above the quoted original."),
  cc: z.array(z.string().email()).optional().describe("Optional CC email addresses."),
};

const createDraftArgs = z.object(createDraftSchema);

export const createDraftDescription =
  "Create an email draft in the Outlook Drafts folder — this tool NEVER sends mail; the user reviews and sends from Outlook. Two mutually exclusive modes: reply mode (set reply_to_message_id only) drafts a reply with your text above the quoted original; new-message mode (set to + subject) drafts a fresh message. body is required in both modes.";

export async function createDraftHandler(
  input: z.input<typeof createDraftArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { reply_to_message_id, to, subject, body, cc } = createDraftArgs.parse(input);

    const replyMode = reply_to_message_id !== undefined;
    const newMode = to !== undefined || subject !== undefined;
    if (replyMode === newMode) {
      return errorResult(
        "Exactly one mode required: either reply_to_message_id (reply mode) or to + subject (new-message mode) — not both, not neither."
      );
    }
    if (!replyMode && (!to?.length || !subject)) {
      return errorResult("New-message mode requires both a non-empty 'to' list and a 'subject'.");
    }

    let draft: any;
    if (replyMode) {
      const created = await callGraphServer(
        `/me/messages/${encodeURIComponent(reply_to_message_id)}/createReply`,
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
      `Draft created${replyMode ? " (reply)" : ""}.\n` +
        `Subject: ${draft.subject || "(no subject)"}\n` +
        `To: ${recipients || "(none)"}\n` +
        (ccList ? `Cc: ${ccList}\n` : "") +
        `Message id: ${draft.id}\n` +
        "Saved to Drafts — review and send from Outlook."
    );
  });
}
