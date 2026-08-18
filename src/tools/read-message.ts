import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import {
  TZ_PREFER,
  ToolResult,
  formatLocal,
  formatSender,
  runTool,
  textResult,
} from "./common.js";

export const readMessageSchema = {
  message_id: z
    .string()
    .min(1)
    .describe("The id of the message to read (from search_mail, read_thread, or list_folders)."),
  include_attachments_list: z
    .boolean()
    .default(true)
    .describe(
      "Include an inventory of the message's attachments (name, size, content type, attachment id) for use with get_attachment (default true)."
    ),
};

const readMessageArgs = z.object(readMessageSchema);

export const readMessageDescription =
  "Read a single email message in full: headers (from/to/cc, date in America/Toronto, subject), the plain-text body, and an attachment inventory whose attachment ids can be passed to get_attachment. Use this over read_thread when you need one specific message or its attachments.";

const BODY_LIMIT = 10000;

function formatAddressList(recipients: any[] | undefined): string {
  const list = (recipients ?? []).map((r) => formatSender(r)).filter(Boolean);
  return list.length ? list.join(", ") : "(none)";
}

export function formatSize(bytes: number | undefined): string {
  const n = bytes ?? 0;
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

export async function readMessageHandler(
  input: z.input<typeof readMessageArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { message_id, include_attachments_list } = readMessageArgs.parse(input);
    const select =
      "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,hasAttachments,isDraft,conversationId";
    const msg = await callGraphServer(
      `/me/messages/${encodeURIComponent(message_id)}?$select=${select}`,
      { headers: { Prefer: `${TZ_PREFER}, outlook.body-content-type="text"` } }
    );

    let body = (msg.body?.content ?? "").replace(/\r\n/g, "\n").trim();
    if (body.length > BODY_LIMIT) body = `${body.slice(0, BODY_LIMIT)}\n[truncated]`;

    let attachmentSection = "";
    if (include_attachments_list && msg.hasAttachments) {
      const atts = await callGraphServer(
        `/me/messages/${encodeURIComponent(message_id)}/attachments?$select=id,name,contentType,size,isInline`
      );
      const lines = (atts?.value ?? []).map(
        (a: any, i: number) =>
          `${i + 1}. ${a.name || "(unnamed)"} — ${a.contentType || "unknown type"}, ${formatSize(a.size)}${a.isInline ? " (inline)" : ""}\n   Attachment id: ${a.id}`
      );
      attachmentSection = lines.length
        ? `\n\nAttachments (${lines.length}):\n${lines.join("\n")}`
        : "";
    } else if (include_attachments_list) {
      attachmentSection = "\n\nAttachments: none";
    }

    return textResult(
      `Subject: ${msg.subject || "(no subject)"}\n` +
        `From: ${formatSender(msg.from)}\n` +
        `To: ${formatAddressList(msg.toRecipients)}\n` +
        (msg.ccRecipients?.length ? `Cc: ${formatAddressList(msg.ccRecipients)}\n` : "") +
        `Date: ${formatLocal(msg.receivedDateTime ?? msg.sentDateTime)}\n` +
        (msg.isDraft ? "Status: draft (not sent)\n" : "") +
        `Message id: ${msg.id}\n` +
        `Conversation id: ${msg.conversationId}\n\n` +
        (body || "(empty body)") +
        attachmentSection
    );
  });
}
