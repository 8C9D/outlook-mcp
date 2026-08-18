// The raw MIME (.eml) of a message, from GET /me/messages/{id}/$value.
//
// Delivery follows the same split as get_attachment, and for the same reason:
// the local stdio server writes the file to disk and hands back a path, while
// the Worker has no filesystem and parks the bytes for the authenticated
// download route instead (core/downloads.ts).
import { z } from "zod";
import { callGraphServer, callGraphServerBytes } from "../core/graph.js";
import {
  DOWNLOAD_MAX_BYTES,
  DOWNLOAD_TTL_DEFAULT_MINUTES,
  DOWNLOAD_TTL_MAX_MINUTES,
  storeDownload,
} from "../core/downloads.js";
import { getStateStore } from "../core/state.js";
import {
  TZ_PREFER,
  ToolResult,
  errorResult,
  formatLocal,
  formatSender,
  isNotFound,
  runTool,
  textResult,
} from "./common.js";
import { formatSize } from "./read-message.js";
import { saveToDownloads } from "./save-local.js";

export const exportMessageSchema = {
  message_id: z
    .string()
    .min(1)
    .describe("The id of the message to export (from search_mail, read_message, or read_thread)."),
  link_ttl_minutes: z
    .number()
    .int()
    .min(1)
    .max(DOWNLOAD_TTL_MAX_MINUTES)
    .default(DOWNLOAD_TTL_DEFAULT_MINUTES)
    .describe(
      `Hosted server only: how long the download link stays valid, in minutes (default ${DOWNLOAD_TTL_DEFAULT_MINUTES}, max ${DOWNLOAD_TTL_MAX_MINUTES}). Ignored by the local server, which saves the .eml to disk instead.`
    ),
};

const exportMessageArgs = z.object(exportMessageSchema);

export const exportMessageDescription =
  "Export one message as a .eml file — its exact MIME source, headers and body and attachments as it arrived. This is the artifact to produce when a message needs inspecting or reporting rather than reading: a phishing sample to hand to an IT/security team or an abuse address, or an evidential copy of a message that must survive being deleted from the mailbox. read_message with include_headers already answers \"is this spoofed?\" without a file, so reach for this when someone needs the message ITSELF. Where the file goes depends on the server: the local (stdio) server saves it to ~/Downloads/outlook-mcp-attachments/ and returns the path, while the hosted server returns a sign-in-required download link that expires within 15 minutes. Attaching the .eml to an outgoing mail is a separate step — create_draft, then add_attachment with the saved path.";

/** A filesystem-friendly stem for the .eml, from the subject (or the id). */
function filenameFor(subject: string | undefined, messageId: string): string {
  const stem = String(subject ?? "")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/[^\w.\- ]+/g, "")
    .trim()
    .slice(0, 80)
    .trim();
  return `${stem || `message-${messageId.slice(0, 12)}`}.eml`;
}

export async function exportMessageHandler(
  input: z.input<typeof exportMessageArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { message_id, link_ttl_minutes } = exportMessageArgs.parse(input);

    let msg: any;
    try {
      msg = await callGraphServer(
        `/me/messages/${encodeURIComponent(message_id)}?$select=id,subject,from,receivedDateTime,sentDateTime`,
        { headers: { Prefer: TZ_PREFER } }
      );
    } catch (err) {
      if (isNotFound(err)) {
        return errorResult(
          `No message ${message_id} — check the id with search_mail or read_thread.`
        );
      }
      throw err;
    }

    const { bytes } = await callGraphServerBytes(
      `/me/messages/${encodeURIComponent(message_id)}/$value`
    );
    const filename = filenameFor(msg.subject, msg.id ?? message_id);
    const head =
      "Message exported as MIME (.eml).\n" +
      `Subject: ${msg.subject || "(no subject)"}\n` +
      `From: ${formatSender(msg.from)}\n` +
      `Date: ${formatLocal(msg.receivedDateTime ?? msg.sentDateTime)}\n` +
      `Size: ${formatSize(bytes.length)}`;

    const store = getStateStore();
    if (store?.mode === "remote") {
      if (bytes.length > DOWNLOAD_MAX_BYTES) {
        return errorResult(
          `${filename} is ${formatSize(bytes.length)}, too large for this server to hand over ` +
            `(limit ${formatSize(DOWNLOAD_MAX_BYTES)}). Export it from the local stdio server, ` +
            "which writes straight to disk."
        );
      }
      const parked = await storeDownload(
        store,
        {
          name: filename,
          contentType: "message/rfc822",
          base64: Buffer.from(bytes).toString("base64"),
          size: bytes.length,
        },
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
          "it stops working when it expires. Export again for a fresh link."
      );
    }

    const savePath = await saveToDownloads(filename, bytes);
    return textResult(`${head}\nSaved to: ${savePath}`);
  });
}
