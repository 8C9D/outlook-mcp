import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { getStateStore } from "../core/state.js";
import { STATE_SIGNATURE } from "../core/kv-keys.js";
import { ToolResult, errorResult, runTool, textResult, toRecipients } from "./common.js";

// NOTE: Sending is two-step by structure. This tool only composes drafts;
// the sole send path in this codebase is send_draft (POST /messages/{id}/send).
// Nothing here may ever call /me/sendMail.

/** Escape text for embedding in an HTML body (signature, plain fragments). */
export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** The signature the user stored with mailbox_settings, or undefined. */
export async function readSignature(): Promise<string | undefined> {
  const store = getStateStore();
  if (!store) return undefined;
  const raw = await store.get(STATE_SIGNATURE).catch(() => null);
  if (!raw) return undefined;
  try {
    const parsed = JSON.parse(raw);
    return typeof parsed?.text === "string" && parsed.text ? parsed.text : undefined;
  } catch {
    return undefined;
  }
}

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
      "Message body. Plain text unless body_format is \"html\". In reply and forward modes it is placed above the quoted original."
    ),
  body_format: z
    .enum(["text", "html"])
    .default("text")
    .describe(
      'How to interpret body: "text" (default) or "html" for a formatted message — body is then the HTML fragment for the new content (links, lists, bold, …).'
    ),
  cc: z.array(z.string().email()).optional().describe("Optional CC email addresses."),
  bcc: z
    .array(z.string().email())
    .optional()
    .describe("Optional BCC email addresses — hidden from the other recipients."),
  importance: z
    .enum(["low", "normal", "high"])
    .optional()
    .describe('Message importance flag shown to recipients ("high" = the red exclamation mark).'),
  request_read_receipt: z
    .boolean()
    .optional()
    .describe("Ask recipients' clients for a read receipt (they may decline to send one)."),
  request_delivery_receipt: z
    .boolean()
    .optional()
    .describe("Ask the receiving server for a delivery receipt."),
  omit_signature: z
    .boolean()
    .default(false)
    .describe(
      "Leave the stored email signature off this draft. By default, a signature saved with mailbox_settings set_signature is appended under the new text."
    ),
};

const createDraftArgs = z.object(createDraftSchema);

export const createDraftDescription =
  "Create an email draft in the Outlook Drafts folder — this tool never sends; sending requires a separate send_draft call. Three mutually exclusive modes: reply mode (reply_to_message_id, optionally reply_all) drafts a reply with your text above the quoted original; forward mode (forward_message_id + to) drafts a forward; new-message mode (to + subject) drafts a fresh message. body is required in every mode (HTML with body_format \"html\"). Also takes cc, bcc, importance, and read/delivery receipt requests. If a signature is stored (mailbox_settings set_signature) it is appended automatically unless omit_signature is set.";

export async function createDraftHandler(
  input: z.input<typeof createDraftArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const {
      reply_to_message_id,
      reply_all,
      forward_message_id,
      to,
      subject,
      body,
      body_format,
      cc,
      bcc,
      importance,
      request_read_receipt,
      request_delivery_receipt,
      omit_signature,
    } = createDraftArgs.parse(input);

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

    const html = body_format === "html";
    const signature = omit_signature ? undefined : await readSignature();
    // The signature sits under the new text — above the quoted tail in
    // reply/forward mode, exactly where Outlook's own clients put it.
    const signedBody = signature
      ? html
        ? `${body}<br><br>${escapeHtml(signature).replace(/\n/g, "<br>")}`
        : `${body}\n\n${signature}`
      : body;

    /** Fields shared by every mode; recipients and body are per-mode. */
    const extraFields = {
      ...(importance !== undefined ? { importance } : {}),
      ...(request_read_receipt !== undefined ? { isReadReceiptRequested: request_read_receipt } : {}),
      ...(request_delivery_receipt !== undefined
        ? { isDeliveryReceiptRequested: request_delivery_receipt }
        : {}),
    };

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
      // Fetch the auto-generated quoted body (text or HTML to match body_format),
      // then place the new text — and the signature — above it.
      const existing = await callGraphServer(
        `/me/messages/${created.id}?$select=body,subject,toRecipients`,
        html ? undefined : { headers: { Prefer: 'outlook.body-content-type="text"' } }
      );
      const quoted = (existing.body?.content ?? "").replace(/\r\n/g, "\n");
      const patch: any = {
        body: html
          ? { contentType: "HTML", content: `${signedBody}<br><br>${quoted}` }
          : { contentType: "Text", content: `${signedBody}\n\n${quoted}`.trimEnd() + "\n" },
        ...extraFields,
      };
      if (forwardMode && to?.length) patch.toRecipients = toRecipients(to);
      if (cc?.length) {
        const existingCc = created.ccRecipients ?? [];
        patch.ccRecipients = [...existingCc, ...toRecipients(cc)];
      }
      if (bcc?.length) patch.bccRecipients = toRecipients(bcc);
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
          body: { contentType: html ? "HTML" : "Text", content: signedBody },
          toRecipients: toRecipients(to!),
          ...(cc?.length ? { ccRecipients: toRecipients(cc) } : {}),
          ...(bcc?.length ? { bccRecipients: toRecipients(bcc) } : {}),
          ...extraFields,
        }),
      });
    }

    const addressList = (recipients: any[] | undefined) =>
      (recipients ?? [])
        .map((r: any) => r.emailAddress?.address)
        .filter(Boolean)
        .join(", ");
    const recipients = addressList(draft.toRecipients);
    const ccList = addressList(draft.ccRecipients);
    const bccList = addressList(draft.bccRecipients);
    return textResult(
      `Draft created${modeLabel}.\n` +
        `Subject: ${draft.subject || "(no subject)"}\n` +
        `To: ${recipients || "(none)"}\n` +
        (ccList ? `Cc: ${ccList}\n` : "") +
        (bccList ? `Bcc: ${bccList}\n` : "") +
        (importance && importance !== "normal" ? `Importance: ${importance}\n` : "") +
        (request_read_receipt ? "Read receipt: requested\n" : "") +
        (request_delivery_receipt ? "Delivery receipt: requested\n" : "") +
        (html ? "Body: HTML\n" : "") +
        (signature ? "Signature: appended (omit_signature skips it)\n" : "") +
        `Draft id: ${draft.id}\n` +
        "Saved to Drafts — not sent. Use update_draft to revise, or send_draft to send it."
    );
  });
}
