import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { ToolResult, errorResult, formatLocal, runTool, textResult, torontoInstantUtc } from "./common.js";

// This is the ONLY send path in the codebase: an existing draft, by id.
// One-shot compose-and-send (e.g. /me/sendMail) is deliberately not implemented.

/**
 * MAPI PR_DEFERRED_SEND_TIME (0x3FEF, SystemTime), the property behind
 * Outlook's delayed delivery. Verified live on this consumer mailbox: PATCH it
 * onto a draft, POST /send, and the message STAYS IN DRAFTS (still a draft,
 * same id) until the deferred instant, then sends and moves to Sent Items.
 * manage_scheduled_send lists and cancels messages in that parked state.
 */
export const DEFERRED_SEND_PROPERTY = "SystemTime 0x3FEF";

/** How far ahead a scheduled send must be — Exchange needs a little margin. */
export const SEND_AT_MIN_LEAD_MS = 2 * 60 * 1000;
/** And how far out it may be. Matches Outlook web's own one-year horizon. */
export const SEND_AT_MAX_AHEAD_MS = 366 * 24 * 60 * 60 * 1000;

const NAIVE_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?$/;
const OFFSET_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d{1,6})?)?(Z|[+-]\d{2}:?\d{2})$/;

/**
 * Parse and validate a send_at input against `now`. Naive datetimes are
 * America/Toronto wall clock; an explicit offset is honored. Pure, so the
 * offline tier can cover the window checks.
 */
export function resolveSendAt(
  sendAt: string,
  now: Date
): { ok: true; utcIso: string } | { ok: false; message: string } {
  let utcIso: string | undefined;
  if (NAIVE_RE.test(sendAt)) {
    utcIso = torontoInstantUtc(sendAt);
  } else if (OFFSET_RE.test(sendAt)) {
    const d = new Date(sendAt);
    if (!Number.isNaN(d.getTime())) utcIso = d.toISOString().replace(/\.\d{3}Z$/, "Z");
  }
  if (!utcIso) {
    return {
      ok: false,
      message:
        `Could not parse send_at ${JSON.stringify(sendAt)} — use "YYYY-MM-DDTHH:MM" ` +
        "(America/Toronto) or an ISO datetime with an explicit UTC offset.",
    };
  }
  const instant = Date.parse(utcIso);
  if (instant < now.getTime() + SEND_AT_MIN_LEAD_MS) {
    return {
      ok: false,
      message:
        `send_at (${sendAt}) must be at least 2 minutes in the future — to send now, ` +
        "call send_draft without send_at.",
    };
  }
  if (instant > now.getTime() + SEND_AT_MAX_AHEAD_MS) {
    return { ok: false, message: `send_at (${sendAt}) is more than a year away — pick a nearer time.` };
  }
  return { ok: true, utcIso };
}

export const sendDraftSchema = {
  draft_id: z
    .string()
    .min(1)
    .describe("The id of the draft to send (from create_draft or update_draft)."),
  send_at: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Schedule the send instead of sending now: an ISO datetime such as "2026-08-26T09:00" ' +
        "(America/Toronto when no UTC offset is given), at least 2 minutes and at most a year " +
        "ahead. The message stays in Drafts until that moment, then sends by itself — " +
        "manage_scheduled_send lists and cancels it while it waits. Omit to send immediately."
    ),
};

const sendDraftArgs = z.object(sendDraftSchema);

export const sendDraftDescription =
  "Send an existing draft email to its recipients — immediately, or at a scheduled future time via send_at (delayed delivery). IRREVERSIBLE once it goes: an immediate send leaves the account at once and cannot be recalled; a scheduled send parks the message in Drafts until send_at and can still be cancelled with manage_scheduled_send before that moment. Before calling, state the draft's exact subject and recipients (and the scheduled time, if any) to the user. Fails if the id is not a draft. Compose with create_draft/update_draft first — there is no compose-and-send in one step.";

export async function sendDraftHandler(
  input: z.input<typeof sendDraftArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { draft_id, send_at } = sendDraftArgs.parse(input);

    let sendAtUtc: string | undefined;
    if (send_at !== undefined) {
      const resolved = resolveSendAt(send_at, new Date());
      if (!resolved.ok) return errorResult(resolved.message);
      sendAtUtc = resolved.utcIso;
    }

    const msg = await callGraphServer(
      `/me/messages/${encodeURIComponent(draft_id)}?$select=isDraft,subject,toRecipients,ccRecipients,bccRecipients`
    );
    if (!msg.isDraft) {
      return errorResult(
        `Message ${draft_id} is not a draft (subject: ${JSON.stringify(msg.subject ?? "")}) — only drafts can be sent.`
      );
    }
    const addressList = (recipients: any[] | undefined) =>
      (recipients ?? [])
        .map((r: any) => r.emailAddress?.address)
        .filter(Boolean)
        .join(", ");
    const recipients = addressList(msg.toRecipients);
    const ccList = addressList(msg.ccRecipients);
    const bccList = addressList(msg.bccRecipients);
    if (!recipients) {
      return errorResult("The draft has no To recipients — add them with update_draft first.");
    }

    if (sendAtUtc) {
      await callGraphServer(`/me/messages/${encodeURIComponent(draft_id)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          singleValueExtendedProperties: [{ id: DEFERRED_SEND_PROPERTY, value: sendAtUtc }],
        }),
      });
    }
    await callGraphServer(`/me/messages/${encodeURIComponent(draft_id)}/send`, {
      method: "POST",
    });

    const recipientLines =
      `Subject: ${msg.subject || "(no subject)"}\n` +
      `To: ${recipients}\n` +
      (ccList ? `Cc: ${ccList}\n` : "") +
      (bccList ? `Bcc: ${bccList}\n` : "");
    if (sendAtUtc) {
      return textResult(
        `Draft scheduled to send at ${formatLocal(sendAtUtc)} (America/Toronto).\n` +
          recipientLines +
          `Message id: ${draft_id}\n` +
          "Until then it stays in the Drafts folder and nothing has left the account. " +
          "manage_scheduled_send lists it and can cancel it before that time; once the time " +
          "arrives it sends by itself and moves to Sent Items."
      );
    }
    return textResult(
      `Draft sent.\n` +
        recipientLines +
        "The message is now in Sent Items (its id changed on send)."
    );
  });
}
