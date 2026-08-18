import { promises as fs } from "node:fs";
import path from "node:path";
import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";
import { formatSize } from "./read-message.js";

export const addAttachmentSchema = {
  draft_id: z
    .string()
    .min(1)
    .describe("The id of the draft to attach the file to (from create_draft)."),
  file_path: z
    .string()
    .min(1)
    .describe("Absolute local path of the file to attach (e.g. /Users/me/report.pdf)."),
  attachment_name: z
    .string()
    .min(1)
    .optional()
    .describe("Name the attachment should carry in the email. Defaults to the file's basename."),
};

const addAttachmentArgs = z.object(addAttachmentSchema);

export const addAttachmentDescription =
  "Attach a local file to an existing email draft (max 25 MB). Natural flow: create_draft → add_attachment (once per file) → send_draft. Fails if the id is not a draft or the file is missing/oversized. This tool never sends — the draft stays in Drafts until send_draft is called.";

const SMALL_LIMIT = 3 * 1024 * 1024; // single-POST fileAttachment below this
const MAX_SIZE = 25 * 1024 * 1024; // hard cap for this tool
const CHUNK_SIZE = 4 * 1024 * 1024; // upload-session chunk size

const CONTENT_TYPES: Record<string, string> = {
  ".txt": "text/plain",
  ".md": "text/markdown",
  ".csv": "text/csv",
  ".html": "text/html",
  ".htm": "text/html",
  ".ics": "text/calendar",
  ".json": "application/json",
  ".xml": "application/xml",
  ".pdf": "application/pdf",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".webp": "image/webp",
  ".heic": "image/heic",
  ".svg": "image/svg+xml",
  ".zip": "application/zip",
  ".gz": "application/gzip",
  ".doc": "application/msword",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xls": "application/vnd.ms-excel",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".ppt": "application/vnd.ms-powerpoint",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".mp3": "audio/mpeg",
  ".mp4": "video/mp4",
  ".mov": "video/quicktime",
};

export function contentTypeForFile(filename: string): string {
  return CONTENT_TYPES[path.extname(filename).toLowerCase()] ?? "application/octet-stream";
}

/**
 * Upload a 3–25 MB file via an Outlook attachment upload session: 4 MB PUT
 * chunks against the session's uploadUrl. The uploadUrl carries its own auth
 * token, so chunk PUTs are plain fetches — NOT callGraphServer (whose bearer
 * token is for graph.microsoft.com, a different audience).
 */
async function uploadViaSession(
  draftId: string,
  name: string,
  contentType: string,
  buffer: Buffer
): Promise<void> {
  const session = await callGraphServer(
    `/me/messages/${encodeURIComponent(draftId)}/attachments/createUploadSession`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        AttachmentItem: { attachmentType: "file", name, contentType, size: buffer.length },
      }),
    }
  );
  const uploadUrl: string | undefined = session?.uploadUrl;
  if (!uploadUrl) throw new Error("Graph did not return an uploadUrl for the attachment session.");

  for (let offset = 0; offset < buffer.length; offset += CHUNK_SIZE) {
    const chunk = buffer.subarray(offset, Math.min(offset + CHUNK_SIZE, buffer.length));
    const range = `bytes ${offset}-${offset + chunk.length - 1}/${buffer.length}`;
    const response = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Type": "application/octet-stream",
        "Content-Length": String(chunk.length),
        "Content-Range": range,
      },
      body: new Uint8Array(chunk),
    });
    if (!response.ok) {
      const body = await response.text();
      throw new Error(
        `Attachment upload failed at ${range}: HTTP ${response.status} ${response.statusText}\n${body.slice(0, 500)}`
      );
    }
  }
}

export async function addAttachmentHandler(
  input: z.input<typeof addAttachmentArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { draft_id, file_path, attachment_name } = addAttachmentArgs.parse(input);

    if (!path.isAbsolute(file_path)) {
      return errorResult(`file_path must be an absolute path, got: ${file_path}`);
    }
    let stat;
    try {
      stat = await fs.stat(file_path);
    } catch {
      return errorResult(`File not found or unreadable: ${file_path}`);
    }
    if (!stat.isFile()) {
      return errorResult(`Not a regular file: ${file_path}`);
    }
    try {
      await fs.access(file_path, fs.constants.R_OK);
    } catch {
      return errorResult(`File exists but is not readable: ${file_path}`);
    }
    // Size guard BEFORE reading any bytes — a stat is enough to refuse oversize files.
    if (stat.size > MAX_SIZE) {
      return errorResult(
        `File is ${formatSize(stat.size)}, over this tool's 25 MB attachment cap. ` +
          "Share large files another way (e.g. a cloud link pasted into the draft body)."
      );
    }

    const msg = await callGraphServer(
      `/me/messages/${encodeURIComponent(draft_id)}?$select=isDraft,subject`
    );
    if (!msg.isDraft) {
      return errorResult(
        `Message ${draft_id} is not a draft (subject: ${JSON.stringify(msg.subject ?? "")}) — attachments can only be added to drafts.`
      );
    }

    const name = attachment_name ?? path.basename(file_path);
    const contentType = contentTypeForFile(name);
    const buffer = await fs.readFile(file_path);

    if (buffer.length < SMALL_LIMIT) {
      await callGraphServer(`/me/messages/${encodeURIComponent(draft_id)}/attachments`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name,
          contentType,
          contentBytes: buffer.toString("base64"),
        }),
      });
    } else {
      await uploadViaSession(draft_id, name, contentType, buffer);
    }

    return textResult(
      `Attachment added.\n` +
        `Name: ${name}\n` +
        `Type: ${contentType}\n` +
        `Size: ${formatSize(buffer.length)}\n` +
        `Draft subject: ${msg.subject || "(no subject)"}\n` +
        `Draft id: ${draft_id}\n` +
        "Still in Drafts — use send_draft to send it."
    );
  });
}
