import { z } from "zod";
import { GraphError } from "../core/graph.js";
import {
  TZ_PREFER,
  ToolResult,
  escapeODataString,
  fetchPagedWithInfo,
  formatLocal,
  formatSender,
  runTool,
  textResult,
} from "./common.js";

export const readThreadSchema = {
  conversation_id: z
    .string()
    .min(1)
    .describe(
      "The conversation id of the thread to read (the 'Conversation id' returned by search_mail, not a message id)."
    ),
  max_messages: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(20)
    .describe("Maximum number of messages to return from the thread (default 20)."),
};

const readThreadArgs = z.object(readThreadSchema);

export const readThreadDescription =
  "Read an email conversation (thread) oldest-to-newest as plain text, given a conversation id from search_mail. Each message shows sender, datetime (America/Toronto), and body text with quoted-reply tails trimmed. Long bodies are truncated at 2000 characters.";

const BODY_LIMIT = 2000;

/** How far an unordered read goes to find a long thread's newest messages. */
const THREAD_SCAN_CAP = 200;

/**
 * Trim quoted-reply tails at well-known delimiters. Conservative: only trims on
 * unambiguous markers, and never trims when the marker leads the message.
 */
function trimQuotedTail(body: string): string {
  const lines = body.split("\n");
  let cut = -1;
  for (let i = 1; i < lines.length; i++) {
    const line = (lines[i] ?? "").trim();
    if (
      /^-{3,}\s*Original Message\s*-{3,}$/i.test(line) ||
      /^On .{5,200} wrote:$/.test(line) ||
      (/^_{10,}$/.test(line) &&
        lines.slice(i + 1, i + 4).some((l) => /^\s*From:\s/i.test(l ?? "")))
    ) {
      cut = i;
      break;
    }
  }
  if (cut === -1) return body;
  const kept = lines.slice(0, cut).join("\n").trimEnd();
  return kept ? `${kept}\n[quoted reply trimmed]` : body;
}

export async function readThreadHandler(
  input: z.input<typeof readThreadArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { conversation_id, max_messages } = readThreadArgs.parse(input);
    const filter = encodeURIComponent(
      `conversationId eq '${escapeODataString(conversation_id)}'`
    );
    const select = "id,subject,from,receivedDateTime,body";
    const headers = {
      Prefer: `${TZ_PREFER}, outlook.body-content-type="text"`,
    };
    const base = `/me/messages?$filter=${filter}&$select=${select}&$top=${Math.min(max_messages, 50)}`;

    // Newest first from Graph, so a cap costs the oldest messages, not the ones
    // the reader came for. The answer is put back in reading order below.
    let page: { items: any[]; truncated: boolean };
    try {
      page = await fetchPagedWithInfo(`${base}&$orderby=receivedDateTime%20desc`, max_messages, headers);
    } catch (err) {
      // Graph often rejects $filter+$orderby on messages (InefficientFilter).
      // Unordered, a cap would keep an arbitrary subset, so read far enough to
      // find the newest and sort locally.
      if (err instanceof GraphError && /InefficientFilter/i.test(err.body)) {
        page = await fetchPagedWithInfo(base, Math.max(max_messages, THREAD_SCAN_CAP), headers);
      } else {
        throw err;
      }
    }
    const newestFirst = [...page.items].sort((a, b) =>
      String(b.receivedDateTime ?? "").localeCompare(String(a.receivedDateTime ?? ""))
    );
    const messages = newestFirst.slice(0, max_messages).reverse();
    const omitted = page.truncated || newestFirst.length > max_messages;

    if (messages.length === 0) {
      return textResult(`No messages found for conversation id ${conversation_id}.`);
    }

    const subject = messages[0]?.subject || "(no subject)";
    const rendered = messages.map((m, i) => {
      let body = (m.body?.content ?? "").replace(/\r\n/g, "\n").trim();
      body = trimQuotedTail(body);
      if (body.length > BODY_LIMIT) body = `${body.slice(0, BODY_LIMIT)}\n[truncated]`;
      return (
        `--- Message ${i + 1} of ${messages.length} ---\n` +
        `From: ${formatSender(m.from)}\n` +
        `Date: ${formatLocal(m.receivedDateTime)}\n\n` +
        (body || "(empty body)")
      );
    });
    const scope = omitted
      ? `newest ${messages.length} message(s), oldest of those first; earlier messages are NOT shown` +
        (max_messages < 100 ? " — raise max_messages to read further back" : "")
      : `${messages.length} message(s), oldest first`;
    return textResult(`Thread: ${subject} (${scope})\n\n` + rendered.join("\n\n"));
  });
}
