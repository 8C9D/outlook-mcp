import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import {
  DOWNLOAD_MAX_BYTES,
  DOWNLOAD_TTL_DEFAULT_MINUTES,
  DOWNLOAD_TTL_MAX_MINUTES,
  storeDownload,
} from "../core/downloads.js";
import { getStateStore } from "../core/state.js";
import { ToolResult, errorResult, formatLocal, runTool, textResult } from "./common.js";
import { formatSize } from "./read-message.js";
import { saveToDownloads } from "./save-local.js";

export const getAttachmentSchema = {
  message_id: z.string().min(1).describe("The id of the message the attachment belongs to."),
  attachment_id: z
    .string()
    .min(1)
    .describe("The attachment id (from read_message's attachment inventory)."),
  link_ttl_minutes: z
    .number()
    .int()
    .min(1)
    .max(DOWNLOAD_TTL_MAX_MINUTES)
    .default(DOWNLOAD_TTL_DEFAULT_MINUTES)
    .describe(
      `Hosted server only: how long the download link stays valid, in minutes (default ${DOWNLOAD_TTL_DEFAULT_MINUTES}, max ${DOWNLOAD_TTL_MAX_MINUTES}). Ignored by the local server, which saves the file to disk instead.`
    ),
};

const getAttachmentArgs = z.object(getAttachmentSchema);

export const getAttachmentDescription =
  "Download one attachment of a message. Small text attachments (text/* or JSON under 50 KB) are returned inline on both servers. Everything else depends on where this server runs: the local (stdio) server saves the file to ~/Downloads/outlook-mcp-attachments/ and returns the path, while the hosted server has no filesystem and instead returns a single-mailbox, sign-in-required download link that expires within 15 minutes. Get attachment ids from read_message.";

const INLINE_LIMIT = 50 * 1024;

export async function getAttachmentHandler(
  input: z.input<typeof getAttachmentArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { message_id, attachment_id, link_ttl_minutes } = getAttachmentArgs.parse(input);
    const att = await callGraphServer(
      `/me/messages/${encodeURIComponent(message_id)}/attachments/${encodeURIComponent(attachment_id)}`
    );

    const odataType: string = att["@odata.type"] ?? "";
    if (!att.contentBytes) {
      return errorResult(
        `This attachment has no downloadable file content (${odataType || "unknown type"}). ` +
          "Only file attachments can be saved; item attachments (attached emails/events) and " +
          "reference attachments (cloud links) are not supported."
      );
    }

    const buffer = Buffer.from(att.contentBytes, "base64");
    const name: string = att.name || "attachment";
    const contentType: string = att.contentType ?? "application/octet-stream";
    const textLike = /^text\//i.test(contentType) || /^application\/json\b/i.test(contentType);
    const inline =
      textLike && buffer.length < INLINE_LIMIT
        ? `\n\nContent (${contentType}):\n${buffer.toString("utf8")}`
        : "";

    const store = getStateStore();
    if (store?.mode === "remote") {
      const head =
        `Attachment retrieved.\n` +
        `Name: ${name}\n` +
        `Type: ${contentType}\n` +
        `Size: ${formatSize(buffer.length)}`;
      // Text small enough to read is answered in full; a link to it would be
      // strictly worse for the model and would cost a KV write.
      if (inline) return textResult(`${head}\nDelivered inline (this server has no filesystem).${inline}`);

      if (buffer.length > DOWNLOAD_MAX_BYTES) {
        return errorResult(
          `${name} is ${formatSize(buffer.length)}, too large for this server to hand over ` +
            `(limit ${formatSize(DOWNLOAD_MAX_BYTES)}). Open it in Outlook, or fetch it from the ` +
            "local stdio server, which saves attachments straight to disk."
        );
      }

      const parked = await storeDownload(
        store,
        { name, contentType, base64: att.contentBytes, size: buffer.length },
        link_ttl_minutes
      );
      if (!parked.url) {
        return errorResult(
          "This server does not know its own public URL, so it cannot hand out a download link. " +
            "Set PUBLIC_BASE_URL and redeploy."
        );
      }
      return textResult(
        `${head}\n` +
          `Download: ${parked.url}\n` +
          `Link expires: ${formatLocal(parked.expiresAt)} (in ${link_ttl_minutes} min)\n` +
          "The link needs the same sign-in as this connector — it is useless to anyone else, and " +
          "it stops working when it expires. Fetch it again for a fresh link."
      );
    }

    const savePath = await saveToDownloads(name, buffer);

    return textResult(
      `Attachment saved.\n` +
        `Name: ${att.name || "(unnamed)"}\n` +
        `Type: ${contentType}\n` +
        `Size: ${formatSize(buffer.length)}\n` +
        `Saved to: ${savePath}` +
        inline
    );
  });
}
