// MCP server entry point (stdio transport). stdout carries JSON-RPC only;
// all logging goes to stderr. This process never initiates interactive auth —
// tools fail with a "run `npm run login`" message when the token cache is unusable.
import { readFileSync } from "node:fs";
import path from "node:path";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { searchMailDescription, searchMailHandler, searchMailSchema } from "./tools/search-mail.js";
import { readThreadDescription, readThreadHandler, readThreadSchema } from "./tools/read-thread.js";
import {
  readMessageDescription,
  readMessageHandler,
  readMessageSchema,
} from "./tools/read-message.js";
import {
  getAttachmentDescription,
  getAttachmentHandler,
  getAttachmentSchema,
} from "./tools/get-attachment.js";
import {
  createDraftDescription,
  createDraftHandler,
  createDraftSchema,
} from "./tools/create-draft.js";
import {
  updateDraftDescription,
  updateDraftHandler,
  updateDraftSchema,
} from "./tools/update-draft.js";
import { sendDraftDescription, sendDraftHandler, sendDraftSchema } from "./tools/send-draft.js";
import {
  manageMessageDescription,
  manageMessageHandler,
  manageMessageSchema,
} from "./tools/manage-message.js";
import {
  listFoldersDescription,
  listFoldersHandler,
  listFoldersSchema,
} from "./tools/list-folders.js";
import { listEventsDescription, listEventsHandler, listEventsSchema } from "./tools/list-events.js";
import {
  createEventDescription,
  createEventHandler,
  createEventSchema,
} from "./tools/create-event.js";
import {
  manageEventDescription,
  manageEventHandler,
  manageEventSchema,
} from "./tools/manage-event.js";
import {
  searchContactsDescription,
  searchContactsHandler,
  searchContactsSchema,
} from "./tools/search-contacts.js";
import {
  manageContactDescription,
  manageContactHandler,
  manageContactSchema,
} from "./tools/manage-contact.js";
import { autoReplyDescription, autoReplyHandler, autoReplySchema } from "./tools/auto-reply.js";
import { PROJECT_ROOT } from "./project-root.js";

const { version } = JSON.parse(
  readFileSync(path.join(PROJECT_ROOT, "package.json"), "utf8")
) as { version: string };

const server = new McpServer({ name: "outlook", version });

server.registerTool(
  "search_mail",
  { description: searchMailDescription, inputSchema: searchMailSchema },
  searchMailHandler
);
server.registerTool(
  "read_thread",
  { description: readThreadDescription, inputSchema: readThreadSchema },
  readThreadHandler
);
server.registerTool(
  "read_message",
  { description: readMessageDescription, inputSchema: readMessageSchema },
  readMessageHandler
);
server.registerTool(
  "get_attachment",
  { description: getAttachmentDescription, inputSchema: getAttachmentSchema },
  getAttachmentHandler
);
server.registerTool(
  "create_draft",
  { description: createDraftDescription, inputSchema: createDraftSchema },
  createDraftHandler
);
server.registerTool(
  "update_draft",
  { description: updateDraftDescription, inputSchema: updateDraftSchema },
  updateDraftHandler
);
server.registerTool(
  "send_draft",
  { description: sendDraftDescription, inputSchema: sendDraftSchema },
  sendDraftHandler
);
server.registerTool(
  "manage_message",
  { description: manageMessageDescription, inputSchema: manageMessageSchema },
  manageMessageHandler
);
server.registerTool(
  "list_folders",
  { description: listFoldersDescription, inputSchema: listFoldersSchema },
  listFoldersHandler
);
server.registerTool(
  "list_events",
  { description: listEventsDescription, inputSchema: listEventsSchema },
  listEventsHandler
);
server.registerTool(
  "create_event",
  { description: createEventDescription, inputSchema: createEventSchema },
  createEventHandler
);
server.registerTool(
  "manage_event",
  { description: manageEventDescription, inputSchema: manageEventSchema },
  manageEventHandler
);
server.registerTool(
  "search_contacts",
  { description: searchContactsDescription, inputSchema: searchContactsSchema },
  searchContactsHandler
);
server.registerTool(
  "manage_contact",
  { description: manageContactDescription, inputSchema: manageContactSchema },
  manageContactHandler
);
server.registerTool(
  "auto_reply",
  { description: autoReplyDescription, inputSchema: autoReplySchema },
  autoReplyHandler
);

await server.connect(new StdioServerTransport());
console.error(`outlook MCP server v${version} ready (stdio).`);
