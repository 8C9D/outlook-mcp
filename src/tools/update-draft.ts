import { z } from "zod";
import { GraphError, callGraphServer } from "../core/graph.js";
import { ToolResult, errorResult, runTool, textResult, toRecipients } from "./common.js";

export const updateDraftSchema = {
  draft_id: z.string().min(1).describe("The id of the draft to update (from create_draft)."),
  body: z
    .string()
    .optional()
    .describe(
      "New message body. Replaces the entire existing body — plain text unless body_format is \"html\"."
    ),
  body_format: z
    .enum(["text", "html"])
    .default("text")
    .describe('How to interpret body: "text" (default) or "html".'),
  subject: z.string().optional().describe("New subject line."),
  to: z
    .array(z.string().email())
    .optional()
    .describe("New To recipients. REPLACES the whole existing To list — it does not append."),
  cc: z
    .array(z.string().email())
    .optional()
    .describe(
      "New CC recipients. REPLACES the whole existing CC list — it does not append. An empty array clears CC."
    ),
  bcc: z
    .array(z.string().email())
    .optional()
    .describe(
      "New BCC recipients. REPLACES the whole existing BCC list. An empty array clears BCC."
    ),
  importance: z
    .enum(["low", "normal", "high"])
    .optional()
    .describe('New importance flag ("normal" clears a high/low mark).'),
  request_read_receipt: z
    .boolean()
    .optional()
    .describe("Ask recipients' clients for a read receipt (false withdraws the request)."),
  request_delivery_receipt: z
    .boolean()
    .optional()
    .describe("Ask the receiving server for a delivery receipt (false withdraws the request)."),
  remove_attachments: z
    .array(z.string().min(1))
    .optional()
    .describe(
      "Attachment ids to remove from the draft (from read_message's attachment inventory). Removal is per-id and reported per-id."
    ),
};

const updateDraftArgs = z.object(updateDraftSchema);

export const updateDraftDescription =
  "Update an existing email draft: body (text or HTML), subject, to/cc/bcc, importance, read/delivery receipt requests — and remove attachments by id (read_message lists them; add_attachment adds them). The to, cc, and bcc arrays REPLACE the draft's current recipient lists entirely (they do not append). Fails if the id is not a draft. Sending still requires a separate send_draft call.";

export async function updateDraftHandler(
  input: z.input<typeof updateDraftArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const {
      draft_id,
      body,
      body_format,
      subject,
      to,
      cc,
      bcc,
      importance,
      request_read_receipt,
      request_delivery_receipt,
      remove_attachments,
    } = updateDraftArgs.parse(input);
    const nothingToPatch =
      body === undefined &&
      subject === undefined &&
      to === undefined &&
      cc === undefined &&
      bcc === undefined &&
      importance === undefined &&
      request_read_receipt === undefined &&
      request_delivery_receipt === undefined;
    if (nothingToPatch && !remove_attachments?.length) {
      return errorResult(
        "Nothing to update — provide at least one of body, subject, to, cc, bcc, importance, " +
          "request_read_receipt, request_delivery_receipt, or remove_attachments."
      );
    }

    const existing = await callGraphServer(
      `/me/messages/${encodeURIComponent(draft_id)}?$select=isDraft,subject`
    );
    if (!existing.isDraft) {
      return errorResult(
        `Message ${draft_id} is not a draft (subject: ${JSON.stringify(existing.subject ?? "")}) — only drafts can be updated.`
      );
    }

    // Attachment removals first: per-id, so one bad id does not block the rest.
    const removalLines: string[] = [];
    for (const attachmentId of remove_attachments ?? []) {
      try {
        await callGraphServer(
          `/me/messages/${encodeURIComponent(draft_id)}/attachments/${encodeURIComponent(attachmentId)}`,
          { method: "DELETE" }
        );
        removalLines.push(`OK      attachment ${attachmentId} removed`);
      } catch (err) {
        const detail =
          err instanceof GraphError ? `HTTP ${err.status}` : err instanceof Error ? err.message : String(err);
        removalLines.push(`FAILED  attachment ${attachmentId}: ${detail} — read_message lists the draft's attachment ids`);
      }
    }

    const patch: any = {};
    if (body !== undefined) {
      patch.body = { contentType: body_format === "html" ? "HTML" : "Text", content: body };
    }
    if (subject !== undefined) patch.subject = subject;
    if (to !== undefined) patch.toRecipients = toRecipients(to);
    if (cc !== undefined) patch.ccRecipients = toRecipients(cc);
    if (bcc !== undefined) patch.bccRecipients = toRecipients(bcc);
    if (importance !== undefined) patch.importance = importance;
    if (request_read_receipt !== undefined) patch.isReadReceiptRequested = request_read_receipt;
    if (request_delivery_receipt !== undefined) {
      patch.isDeliveryReceiptRequested = request_delivery_receipt;
    }

    const draft = Object.keys(patch).length
      ? await callGraphServer(`/me/messages/${encodeURIComponent(draft_id)}`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(patch),
        })
      : await callGraphServer(
          `/me/messages/${encodeURIComponent(draft_id)}?$select=id,subject,toRecipients,ccRecipients,bccRecipients`
        );

    const addressList = (recipients: any[] | undefined) =>
      (recipients ?? [])
        .map((r: any) => r.emailAddress?.address)
        .filter(Boolean)
        .join(", ");
    const recipients = addressList(draft.toRecipients);
    const ccList = addressList(draft.ccRecipients);
    const bccList = addressList(draft.bccRecipients);
    const changed = [
      ...Object.keys(patch),
      ...(remove_attachments?.length ? [`${remove_attachments.length} attachment removal(s)`] : []),
    ];
    return textResult(
      `Draft updated (${changed.join(", ")}).\n` +
        `Subject: ${draft.subject || "(no subject)"}\n` +
        `To: ${recipients || "(none)"}\n` +
        (ccList ? `Cc: ${ccList}\n` : "") +
        (bccList ? `Bcc: ${bccList}\n` : "") +
        (removalLines.length ? `${removalLines.join("\n")}\n` : "") +
        `Draft id: ${draft.id}\n` +
        "Still in Drafts — use send_draft to send it."
    );
  });
}
