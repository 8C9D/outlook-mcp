// The tool, prompt and resource surface, shared verbatim by both transports.
// The stdio entry point and the Cloudflare Worker each build an McpServer and
// hand it to registerAll, so the two hosts can never drift in what they expose.
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { searchMailDescription, searchMailHandler, searchMailSchema } from "../tools/search-mail.js";
import { readThreadDescription, readThreadHandler, readThreadSchema } from "../tools/read-thread.js";
import {
  readMessageDescription,
  readMessageHandler,
  readMessageSchema,
} from "../tools/read-message.js";
import {
  getAttachmentDescription,
  getAttachmentHandler,
  getAttachmentSchema,
} from "../tools/get-attachment.js";
import {
  createDraftDescription,
  createDraftHandler,
  createDraftSchema,
} from "../tools/create-draft.js";
import {
  updateDraftDescription,
  updateDraftHandler,
  updateDraftSchema,
} from "../tools/update-draft.js";
import { sendDraftDescription, sendDraftHandler, sendDraftSchema } from "../tools/send-draft.js";
import {
  manageMessageDescription,
  manageMessageHandler,
  manageMessageSchema,
} from "../tools/manage-message.js";
import {
  listFoldersDescription,
  listFoldersHandler,
  listFoldersSchema,
} from "../tools/list-folders.js";
import {
  listEventsDescription,
  listEventsHandler,
  listEventsSchema,
} from "../tools/list-events.js";
import {
  listCalendarsDescription,
  listCalendarsHandler,
  listCalendarsSchema,
} from "../tools/list-calendars.js";
import {
  createEventDescription,
  createEventHandler,
  createEventSchema,
} from "../tools/create-event.js";
import {
  manageEventDescription,
  manageEventHandler,
  manageEventSchema,
} from "../tools/manage-event.js";
import {
  searchContactsDescription,
  searchContactsHandler,
  searchContactsSchema,
} from "../tools/search-contacts.js";
import {
  manageContactDescription,
  manageContactHandler,
  manageContactSchema,
} from "../tools/manage-contact.js";
import { autoReplyDescription, autoReplyHandler, autoReplySchema } from "../tools/auto-reply.js";
import {
  addAttachmentDescription,
  addAttachmentHandler,
  addAttachmentSchema,
} from "../tools/add-attachment.js";
import {
  manageRulesDescription,
  manageRulesHandler,
  manageRulesSchema,
} from "../tools/manage-rules.js";
import {
  createFolderDescription,
  createFolderHandler,
  createFolderSchema,
} from "../tools/create-folder.js";
import {
  manageCategoriesDescription,
  manageCategoriesHandler,
  manageCategoriesSchema,
} from "../tools/manage-categories.js";
import { listTasksDescription, listTasksHandler, listTasksSchema } from "../tools/list-tasks.js";
import { manageTaskDescription, manageTaskHandler, manageTaskSchema } from "../tools/manage-task.js";
import {
  checkNewMailDescription,
  checkNewMailHandler,
  checkNewMailSchema,
} from "../tools/check-new-mail.js";
import {
  getMailboxActivityDescription,
  getMailboxActivityHandler,
  getMailboxActivitySchema,
} from "../tools/get-mailbox-activity.js";
import { registerPrompts } from "./prompts.js";
import { registerResources } from "./resources.js";
import type { ZodRawShape } from "zod";
import type { ToolResult } from "../tools/common.js";

type ToolDefinition = {
  name: string;
  description: string;
  inputSchema: ZodRawShape;
  handler: (input: any) => Promise<ToolResult>;
};

/** Every tool this server exposes, in registration order. */
export const TOOLS: ToolDefinition[] = [
  {
    name: "search_mail",
    description: searchMailDescription,
    inputSchema: searchMailSchema,
    handler: searchMailHandler,
  },
  {
    name: "read_thread",
    description: readThreadDescription,
    inputSchema: readThreadSchema,
    handler: readThreadHandler,
  },
  {
    name: "read_message",
    description: readMessageDescription,
    inputSchema: readMessageSchema,
    handler: readMessageHandler,
  },
  {
    name: "get_attachment",
    description: getAttachmentDescription,
    inputSchema: getAttachmentSchema,
    handler: getAttachmentHandler,
  },
  {
    name: "create_draft",
    description: createDraftDescription,
    inputSchema: createDraftSchema,
    handler: createDraftHandler,
  },
  {
    name: "update_draft",
    description: updateDraftDescription,
    inputSchema: updateDraftSchema,
    handler: updateDraftHandler,
  },
  {
    name: "send_draft",
    description: sendDraftDescription,
    inputSchema: sendDraftSchema,
    handler: sendDraftHandler,
  },
  {
    name: "manage_message",
    description: manageMessageDescription,
    inputSchema: manageMessageSchema,
    handler: manageMessageHandler,
  },
  {
    name: "list_folders",
    description: listFoldersDescription,
    inputSchema: listFoldersSchema,
    handler: listFoldersHandler,
  },
  {
    name: "list_events",
    description: listEventsDescription,
    inputSchema: listEventsSchema,
    handler: listEventsHandler,
  },
  {
    name: "list_calendars",
    description: listCalendarsDescription,
    inputSchema: listCalendarsSchema,
    handler: listCalendarsHandler,
  },
  {
    name: "create_event",
    description: createEventDescription,
    inputSchema: createEventSchema,
    handler: createEventHandler,
  },
  {
    name: "manage_event",
    description: manageEventDescription,
    inputSchema: manageEventSchema,
    handler: manageEventHandler,
  },
  {
    name: "search_contacts",
    description: searchContactsDescription,
    inputSchema: searchContactsSchema,
    handler: searchContactsHandler,
  },
  {
    name: "manage_contact",
    description: manageContactDescription,
    inputSchema: manageContactSchema,
    handler: manageContactHandler,
  },
  {
    name: "auto_reply",
    description: autoReplyDescription,
    inputSchema: autoReplySchema,
    handler: autoReplyHandler,
  },
  {
    name: "add_attachment",
    description: addAttachmentDescription,
    inputSchema: addAttachmentSchema,
    handler: addAttachmentHandler,
  },
  {
    name: "manage_rules",
    description: manageRulesDescription,
    inputSchema: manageRulesSchema,
    handler: manageRulesHandler,
  },
  {
    name: "create_folder",
    description: createFolderDescription,
    inputSchema: createFolderSchema,
    handler: createFolderHandler,
  },
  {
    name: "manage_categories",
    description: manageCategoriesDescription,
    inputSchema: manageCategoriesSchema,
    handler: manageCategoriesHandler,
  },
  {
    name: "list_tasks",
    description: listTasksDescription,
    inputSchema: listTasksSchema,
    handler: listTasksHandler,
  },
  {
    name: "manage_task",
    description: manageTaskDescription,
    inputSchema: manageTaskSchema,
    handler: manageTaskHandler,
  },
  {
    name: "check_new_mail",
    description: checkNewMailDescription,
    inputSchema: checkNewMailSchema,
    handler: checkNewMailHandler,
  },
  {
    name: "get_mailbox_activity",
    description: getMailboxActivityDescription,
    inputSchema: getMailboxActivitySchema,
    handler: getMailboxActivityHandler,
  },
];

/** Register every tool and prompt on a freshly constructed McpServer. */
export function registerAll(server: McpServer): void {
  for (const tool of TOOLS) {
    server.registerTool(
      tool.name,
      { description: tool.description, inputSchema: tool.inputSchema },
      tool.handler
    );
  }
  registerPrompts(server);
  registerResources(server);
}

/** Build the fully-populated MCP server for either transport. */
export function createMcpServer(version: string): McpServer {
  const server = new McpServer({ name: "outlook", version });
  registerAll(server);
  return server;
}
