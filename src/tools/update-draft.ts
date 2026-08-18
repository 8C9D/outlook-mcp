import { z } from "zod";
import { callGraphServer } from "../graph.js";
import { ToolResult, errorResult, runTool, textResult, toRecipients } from "./common.js";

export const updateDraftSchema = {
  draft_id: z.string().min(1).describe("The id of the draft to update (from create_draft)."),
  body: z
    .string()
    .optional()
    .describe("New message body (plain text). Replaces the entire existing body."),
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
};

const updateDraftArgs = z.object(updateDraftSchema);

export const updateDraftDescription =
  "Update an existing email draft: body, subject, to, and/or cc. The to and cc arrays REPLACE the draft's current recipient lists entirely (they do not append). Fails if the id is not a draft. Sending still requires a separate send_draft call.";

export async function updateDraftHandler(
  input: z.input<typeof updateDraftArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { draft_id, body, subject, to, cc } = updateDraftArgs.parse(input);
    if (body === undefined && subject === undefined && to === undefined && cc === undefined) {
      return errorResult("Nothing to update — provide at least one of body, subject, to, cc.");
    }

    const existing = await callGraphServer(
      `/me/messages/${encodeURIComponent(draft_id)}?$select=isDraft,subject`
    );
    if (!existing.isDraft) {
      return errorResult(
        `Message ${draft_id} is not a draft (subject: ${JSON.stringify(existing.subject ?? "")}) — only drafts can be updated.`
      );
    }

    const patch: any = {};
    if (body !== undefined) patch.body = { contentType: "Text", content: body };
    if (subject !== undefined) patch.subject = subject;
    if (to !== undefined) patch.toRecipients = toRecipients(to);
    if (cc !== undefined) patch.ccRecipients = toRecipients(cc);

    const draft = await callGraphServer(`/me/messages/${encodeURIComponent(draft_id)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(patch),
    });

    const recipients = (draft.toRecipients ?? [])
      .map((r: any) => r.emailAddress?.address)
      .filter(Boolean)
      .join(", ");
    const ccList = (draft.ccRecipients ?? [])
      .map((r: any) => r.emailAddress?.address)
      .filter(Boolean)
      .join(", ");
    return textResult(
      `Draft updated (${Object.keys(patch).join(", ")}).\n` +
        `Subject: ${draft.subject || "(no subject)"}\n` +
        `To: ${recipients || "(none)"}\n` +
        (ccList ? `Cc: ${ccList}\n` : "") +
        `Draft id: ${draft.id}\n` +
        "Still in Drafts — use send_draft to send it."
    );
  });
}
