import { z } from "zod";
import {
  TZ_PREFER,
  ToolResult,
  fetchPaged,
  formatLocal,
  formatSender,
  runTool,
  textResult,
} from "./common.js";

export const searchMailSchema = {
  query: z
    .string()
    .min(1)
    .describe(
      "Search text matched against message subjects, senders, and bodies (Microsoft Graph $search). Plain words work best; results are relevance-ranked, not newest-first."
    ),
  folder: z
    .string()
    .default("inbox")
    .describe(
      'Mail folder to search. Well-known names: "inbox" (default), "sentitems", "drafts", "archive", "deleteditems", "junkemail". Also accepts a folder id.'
    ),
  max_results: z
    .number()
    .int()
    .min(1)
    .max(25)
    .default(10)
    .describe("Maximum number of results to return (default 10, max 25)."),
  include_body_preview: z
    .boolean()
    .default(true)
    .describe("Include a one-line preview of each message body (default true)."),
};

const searchMailArgs = z.object(searchMailSchema);

export const searchMailDescription =
  "Search the user's Outlook mail for messages matching a text query. Returns for each hit: subject, sender, received datetime (America/Toronto), message id, conversation id, attachment flag, and optionally a body preview. Use the returned conversation id with read_thread to read the full conversation, or the message id with create_draft to draft a reply.";

export async function searchMailHandler(
  input: z.input<typeof searchMailArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { query, folder, max_results, include_body_preview } = searchMailArgs.parse(input);
    // KQL string literal: escape embedded backslashes and double quotes.
    const kql = `"${query.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
    const select = "id,conversationId,subject,from,receivedDateTime,bodyPreview,hasAttachments";
    // $search cannot be combined with $orderby — results are relevance-ranked.
    const path =
      `/me/mailFolders/${encodeURIComponent(folder)}/messages` +
      `?$search=${encodeURIComponent(kql)}&$select=${select}&$top=${max_results}`;
    const messages = await fetchPaged(path, max_results, { Prefer: TZ_PREFER });

    if (messages.length === 0) {
      return textResult(`No messages matching ${JSON.stringify(query)} in ${folder}.`);
    }
    const lines = messages.map((m, i) => {
      const preview = (m.bodyPreview ?? "").replace(/\s+/g, " ").trim().slice(0, 150);
      return [
        `${i + 1}. ${m.subject || "(no subject)"}`,
        `   From: ${formatSender(m.from)}  At: ${formatLocal(m.receivedDateTime)}${m.hasAttachments ? "  [has attachments]" : ""}`,
        `   Message id: ${m.id}`,
        `   Conversation id: ${m.conversationId}`,
        ...(include_body_preview && preview ? [`   Preview: ${preview}`] : []),
      ].join("\n");
    });
    return textResult(
      `${messages.length} result(s) for ${JSON.stringify(query)} in ${folder} (relevance-ranked):\n\n` +
        lines.join("\n\n")
    );
  });
}
