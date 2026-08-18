import { z } from "zod";
import { GraphError } from "../core/graph.js";
import {
  TZ_PREFER,
  ToolResult,
  escapeODataString,
  fetchPaged,
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

    let messages: any[];
    let sortedByGraph = false;
    try {
      messages = await fetchPaged(
        `${base}&$orderby=receivedDateTime`,
        max_messages,
        headers
      );
      sortedByGraph = true;
    } catch (err) {
      // Graph often rejects $filter+$orderby on messages (InefficientFilter);
      // fall back to unordered fetch and sort locally.
      if (err instanceof GraphError && /InefficientFilter/i.test(err.body)) {
        messages = await fetchPaged(base, max_messages, headers);
      } else {
        throw err;
      }
    }
    if (!sortedByGraph) {
      messages.sort((a, b) =>
        String(a.receivedDateTime ?? "").localeCompare(String(b.receivedDateTime ?? ""))
      );
    }

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
    return textResult(
      `Thread: ${subject} (${messages.length} message(s), oldest first)\n\n` + rendered.join("\n\n")
    );
  });
}
