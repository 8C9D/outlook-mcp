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
  include_headers: z
    .boolean()
    .default(false)
    .describe(
      "Include the message's internet headers, rendered compactly: the SPF/DKIM/DMARC verdict from Authentication-Results, the Received (delivery) chain oldest-first, and a flag when Reply-To or Return-Path disagrees with From. Use it when a message might be phishing or spoofed, or when the user asks where a message really came from (default false)."
    ),
};

const readMessageArgs = z.object(readMessageSchema);

export const readMessageDescription =
  "Read a single email message in full: headers (from/to/cc, date in America/Toronto, subject), the plain-text body, and an attachment inventory whose attachment ids can be passed to get_attachment. Use this over read_thread when you need one specific message or its attachments. Set include_headers to inspect a suspicious message: it adds the SPF/DKIM/DMARC verdict, the delivery chain, and a warning when Reply-To points somewhere other than From — the first thing to check when a message may be phishing. For the raw MIME source of the same message, use export_message.";

const BODY_LIMIT = 10000;
/** Enough of the delivery chain to see the origin without flooding the answer. */
const HOP_LIMIT = 12;

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

/** All values of one header name, in the order Graph returned them. */
function headerValues(headers: any[], name: string): string[] {
  return headers
    .filter((h) => String(h?.name ?? "").toLowerCase() === name.toLowerCase())
    .map((h) => String(h?.value ?? "").replace(/\s+/g, " ").trim());
}

/** "from a.example by b.example — Tue, 18 Aug 2026 21:58:31 +0000" out of a Received header. */
function compactReceived(value: string): string {
  const from = value.match(/\bfrom\s+([^\s(;]+)/i)?.[1];
  const by = value.match(/\bby\s+([^\s(;]+)/i)?.[1];
  const date = value.includes(";") ? value.slice(value.lastIndexOf(";") + 1).trim() : "";
  const route = `from ${from ?? "(unstated)"} by ${by ?? "(unstated)"}`;
  return date ? `${route} — ${date}` : route;
}

/** SPF/DKIM/DMARC verdicts pulled out of an Authentication-Results value. */
function authVerdicts(value: string): string {
  const verdicts = ["spf", "dkim", "dmarc", "compauth"]
    .map((mechanism) => {
      const hit = new RegExp(`\\b${mechanism}=([a-z]+)`, "i").exec(value);
      return hit ? `${mechanism.toUpperCase()} ${hit[1]}` : undefined;
    })
    .filter(Boolean);
  return verdicts.length ? verdicts.join(" · ") : "no SPF/DKIM/DMARC verdict in the header";
}

function addressOf(recipient: any): string {
  return String(recipient?.emailAddress?.address ?? "").toLowerCase();
}

/**
 * The forensic section: what an experienced reader checks first on a suspicious
 * message — who vouched for it, the route it took, and whether a reply would go
 * somewhere other than where the message claims to be from.
 */
function describeHeaders(msg: any): string {
  const headers: any[] = msg.internetMessageHeaders ?? [];
  const lines: string[] = ["", "Internet headers:"];

  if (headers.length === 0) {
    lines.push(
      "  (none — Microsoft Graph carries internet headers only for messages that arrived by " +
        "email; drafts and items created in Outlook have none.)"
    );
    return lines.join("\n");
  }

  const auth = headerValues(headers, "Authentication-Results");
  if (auth.length === 0) {
    lines.push(
      "  Authentication-Results: ABSENT — nothing vouched for this message's origin. Treat any " +
        "claim it makes about who sent it as unverified."
    );
  } else {
    for (const value of auth) {
      lines.push(`  Authentication-Results: ${authVerdicts(value)}`);
      lines.push(`    raw: ${value.length > 300 ? `${value.slice(0, 300)}…` : value}`);
    }
  }

  const from = addressOf(msg.from);
  const replyTo = (msg.replyTo ?? []).map(addressOf).filter(Boolean);
  if (replyTo.length === 0) {
    lines.push(`  Reply-To: none — replies go to From (${from || "unknown"}).`);
  } else if (replyTo.every((address: string) => address === from)) {
    lines.push(`  Reply-To: ${replyTo.join(", ")} — same as From, nothing odd.`);
  } else {
    lines.push(
      `  Reply-To: ${replyTo.join(", ")}  ** MISMATCH: From is ${from || "unknown"} ** — ` +
        "a reply would go to a different address than the message claims to be from, a common " +
        "phishing and business-email-compromise pattern. Check this before replying."
    );
  }

  const returnPath = headerValues(headers, "Return-Path")[0];
  if (returnPath) {
    const bare = returnPath.replace(/^<|>$/g, "").toLowerCase();
    const domainOf = (address: string) => address.slice(address.lastIndexOf("@") + 1);
    const mismatch = from && bare && domainOf(bare) !== domainOf(from);
    lines.push(
      `  Return-Path: ${returnPath}${mismatch ? "  ** bounces go to a different domain than From **" : ""}`
    );
  }

  const received = headerValues(headers, "Received");
  if (received.length === 0) {
    lines.push("  Received: none recorded.");
  } else {
    // Graph lists Received newest-first, the order the hops were prepended;
    // oldest-first reads as the journey the message actually took.
    const chain = [...received].reverse();
    const shown = chain.slice(0, HOP_LIMIT);
    lines.push(
      `  Received chain (${received.length} hop(s), oldest first${received.length > HOP_LIMIT ? `, first ${HOP_LIMIT} shown` : ""}):`
    );
    shown.forEach((value, i) => lines.push(`    ${i + 1}. ${compactReceived(value)}`));
  }

  lines.push(
    `  (${headers.length} header(s) in total — export_message returns the full MIME source.)`
  );
  return lines.join("\n");
}

export async function readMessageHandler(
  input: z.input<typeof readMessageArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { message_id, include_attachments_list, include_headers } = readMessageArgs.parse(input);
    const select =
      "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,body,hasAttachments,isDraft,conversationId" +
      (include_headers ? ",internetMessageHeaders,replyTo" : "");
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
        attachmentSection +
        (include_headers ? `\n${describeHeaders(msg)}` : "")
    );
  });
}
