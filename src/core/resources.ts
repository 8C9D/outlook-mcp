// MCP resources: read-only mailbox context a client can attach without the
// model choosing to call a tool. Both are static URIs (no templates) because
// the two things worth attaching wholesale are "how is this mailbox organised"
// and "what is at the top of the inbox right now".
//
// Registered on the shared McpServer, so stdio and the Worker expose the same
// two resources; the SDK implements resources/list and resources/read the same
// way over both transports.
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { listFoldersHandler } from "../tools/list-folders.js";
import { searchMailHandler } from "../tools/search-mail.js";
import type { ToolResult } from "../tools/common.js";

export const FOLDERS_URI = "outlook://mail/folders";
export const RECENT_INBOX_URI = "outlook://mail/inbox/recent";

/** How many messages the recent-inbox resource carries. */
export const RECENT_INBOX_COUNT = 20;

/**
 * Resource reads have no isError channel — a failed read must reject, so the
 * client reports it rather than attaching an error message as context.
 */
function textOrThrow(result: ToolResult, what: string): string {
  const text = result.content.map((part) => part.text).join("\n");
  if (result.isError) throw new Error(`Could not read ${what}: ${text}`);
  return text;
}

export function registerResources(server: McpServer): void {
  server.registerResource(
    "mail_folders",
    FOLDERS_URI,
    {
      title: "Mail folders",
      description:
        "The mailbox folder tree — top-level folders and one level of subfolders, with unread/total counts and the folder ids the mail tools accept.",
      mimeType: "text/plain",
    },
    async (uri) => ({
      contents: [
        {
          uri: uri.href,
          mimeType: "text/plain",
          text: textOrThrow(await listFoldersHandler({}), "the folder list"),
        },
      ],
    })
  );

  server.registerResource(
    "recent_inbox",
    RECENT_INBOX_URI,
    {
      title: "Recent inbox",
      description: `The ${RECENT_INBOX_COUNT} newest inbox messages, newest first: subject, sender, received time (America/Toronto), ids and a body preview.`,
      mimeType: "text/plain",
    },
    async (uri) => ({
      contents: [
        {
          uri: uri.href,
          mimeType: "text/plain",
          text: textOrThrow(
            await searchMailHandler({ max_results: RECENT_INBOX_COUNT }),
            "the recent inbox"
          ),
        },
      ],
    })
  );
}
