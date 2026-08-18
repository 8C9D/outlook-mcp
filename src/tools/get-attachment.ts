import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { z } from "zod";
import { callGraphServer } from "../graph.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";
import { formatSize } from "./read-message.js";

export const getAttachmentSchema = {
  message_id: z.string().min(1).describe("The id of the message the attachment belongs to."),
  attachment_id: z
    .string()
    .min(1)
    .describe("The attachment id (from read_message's attachment inventory)."),
};

const getAttachmentArgs = z.object(getAttachmentSchema);

export const getAttachmentDescription =
  "Download one attachment of a message to ~/Downloads/outlook-mcp-attachments/ and return the saved file path and size. For small text attachments (text/* or JSON under 50 KB) the content is also returned inline. Get attachment ids from read_message.";

const SAVE_DIR = path.join(os.homedir(), "Downloads", "outlook-mcp-attachments");
const INLINE_LIMIT = 50 * 1024;

/** Strip path separators/control chars so an attachment name cannot escape SAVE_DIR. */
function sanitizeFilename(name: string): string {
  const cleaned = name.replace(/[/\\:\0-\x1f]/g, "_").replace(/^\.+/, "_").trim();
  return cleaned || "attachment";
}

/** Return a path in SAVE_DIR that does not collide with an existing file. */
async function collisionFreePath(filename: string): Promise<string> {
  const ext = path.extname(filename);
  const stem = filename.slice(0, filename.length - ext.length);
  for (let i = 0; ; i++) {
    const candidate = path.join(SAVE_DIR, i === 0 ? filename : `${stem} (${i})${ext}`);
    try {
      await fs.access(candidate);
    } catch {
      return candidate;
    }
  }
}

export async function getAttachmentHandler(
  input: z.input<typeof getAttachmentArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { message_id, attachment_id } = getAttachmentArgs.parse(input);
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
    await fs.mkdir(SAVE_DIR, { recursive: true });
    const savePath = await collisionFreePath(sanitizeFilename(att.name ?? "attachment"));
    await fs.writeFile(savePath, buffer);

    const contentType: string = att.contentType ?? "application/octet-stream";
    const textLike = /^text\//i.test(contentType) || /^application\/json\b/i.test(contentType);
    let inline = "";
    if (textLike && buffer.length < INLINE_LIMIT) {
      inline = `\n\nContent (${contentType}):\n${buffer.toString("utf8")}`;
    }

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
