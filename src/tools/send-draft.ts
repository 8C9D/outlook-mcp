import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";

// This is the ONLY send path in the codebase: an existing draft, by id.
// One-shot compose-and-send (e.g. /me/sendMail) is deliberately not implemented.

export const sendDraftSchema = {
  draft_id: z
    .string()
    .min(1)
    .describe("The id of the draft to send (from create_draft or update_draft)."),
};

const sendDraftArgs = z.object(sendDraftSchema);

export const sendDraftDescription =
  "Send an existing draft email to its recipients. IRREVERSIBLE: the message leaves the account immediately and cannot be recalled. Before calling, state the draft's exact subject and recipients to the user so they know what is being sent. Fails if the id is not a draft. Compose with create_draft/update_draft first — there is no compose-and-send in one step.";

export async function sendDraftHandler(
  input: z.input<typeof sendDraftArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { draft_id } = sendDraftArgs.parse(input);

    const msg = await callGraphServer(
      `/me/messages/${encodeURIComponent(draft_id)}?$select=isDraft,subject,toRecipients,ccRecipients`
    );
    if (!msg.isDraft) {
      return errorResult(
        `Message ${draft_id} is not a draft (subject: ${JSON.stringify(msg.subject ?? "")}) — only drafts can be sent.`
      );
    }
    const recipients = (msg.toRecipients ?? [])
      .map((r: any) => r.emailAddress?.address)
      .filter(Boolean)
      .join(", ");
    const ccList = (msg.ccRecipients ?? [])
      .map((r: any) => r.emailAddress?.address)
      .filter(Boolean)
      .join(", ");
    if (!recipients) {
      return errorResult("The draft has no To recipients — add them with update_draft first.");
    }

    await callGraphServer(`/me/messages/${encodeURIComponent(draft_id)}/send`, {
      method: "POST",
    });

    return textResult(
      `Draft sent.\n` +
        `Subject: ${msg.subject || "(no subject)"}\n` +
        `To: ${recipients}\n` +
        (ccList ? `Cc: ${ccList}\n` : "") +
        "The message is now in Sent Items (its id changed on send)."
    );
  });
}
