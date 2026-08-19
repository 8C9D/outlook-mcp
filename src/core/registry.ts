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
  mailboxSettingsDescription,
  mailboxSettingsHandler,
  mailboxSettingsSchema,
} from "../tools/mailbox-settings.js";
import {
  manageSendersDescription,
  manageSendersHandler,
  manageSendersSchema,
} from "../tools/manage-senders.js";
import {
  exportMessageDescription,
  exportMessageHandler,
  exportMessageSchema,
} from "../tools/export-message.js";
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
import {
  getAutoFilingLogDescription,
  getAutoFilingLogHandler,
  getAutoFilingLogSchema,
} from "../tools/get-auto-filing-log.js";
import {
  manageAutoFilingDescription,
  manageAutoFilingHandler,
  manageAutoFilingSchema,
} from "../tools/manage-auto-filing.js";
import { getHealthDescription, getHealthHandler, getHealthSchema } from "../tools/get-health.js";
import { registerPrompts } from "./prompts.js";
import { registerResources } from "./resources.js";
import type { ZodRawShape } from "zod";
import type { ToolAnnotations } from "@modelcontextprotocol/sdk/types.js";
import type { ToolResult } from "../tools/common.js";

type ToolDefinition = {
  name: string;
  description: string;
  inputSchema: ZodRawShape;
  /** All four hints, always — see the rules below. */
  annotations: Required<Omit<ToolAnnotations, "title">>;
  handler: (input: any) => Promise<ToolResult>;
};

/**
 * The four MCP annotation hints are applied by one rule each, so that thirty
 * tools cannot drift into thirty different readings of the same word. Each hint
 * is stated on every tool rather than left to the protocol's defaults, which are
 * "destructive and open-world unless told otherwise" and would be wrong here far
 * more often than right. ASSUMPTIONS.md (v10) records the judgment calls.
 *
 *   readOnlyHint     the call changes nothing: not the mailbox, not this
 *                    server's own state, not the local disk.
 *   destructiveHint  the call can remove or overwrite something the user would
 *                    miss, or do something outward that cannot be taken back.
 *                    Soft deletes still count — the mail leaves where it was.
 *   idempotentHint   a repeat with the same arguments leaves the same state
 *                    (set-shaped), rather than creating, appending or sending
 *                    a second time.
 *   openWorldHint    the call, or the setting it establishes, moves data
 *                    between this mailbox and parties outside it. Reaching
 *                    Microsoft Graph is not itself "open world" — every tool
 *                    here does that — or the hint would tell a caller nothing.
 */
const READ_ONLY = {
  readOnlyHint: true,
  destructiveHint: false,
  idempotentHint: true,
  openWorldHint: false,
} as const;

/** Every tool this server exposes, in registration order. */
export const TOOLS: ToolDefinition[] = [
  {
    name: "search_mail",
    description: searchMailDescription,
    inputSchema: searchMailSchema,
    annotations: READ_ONLY,
    handler: searchMailHandler,
  },
  {
    name: "read_thread",
    description: readThreadDescription,
    inputSchema: readThreadSchema,
    annotations: READ_ONLY,
    handler: readThreadHandler,
  },
  {
    name: "read_message",
    description: readMessageDescription,
    inputSchema: readMessageSchema,
    annotations: READ_ONLY,
    handler: readMessageHandler,
  },
  {
    name: "get_attachment",
    description: getAttachmentDescription,
    inputSchema: getAttachmentSchema,
    // Writes: a file in ~/Downloads on stdio, a short-lived download record in
    // KV on the Worker. Collision-safe names mean a repeat leaves a second copy.
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    handler: getAttachmentHandler,
  },
  {
    name: "export_message",
    description: exportMessageDescription,
    inputSchema: exportMessageSchema,
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    handler: exportMessageHandler,
  },
  {
    name: "create_draft",
    description: createDraftDescription,
    inputSchema: createDraftSchema,
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    handler: createDraftHandler,
  },
  {
    name: "update_draft",
    description: updateDraftDescription,
    inputSchema: updateDraftSchema,
    // Replace semantics on every field it touches, so the same call twice
    // leaves the same draft.
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    handler: updateDraftHandler,
  },
  {
    name: "send_draft",
    description: sendDraftDescription,
    inputSchema: sendDraftSchema,
    // The one irreversible outward act on the surface: mail that has left
    // cannot be recalled, and the draft is no longer a draft.
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true },
    handler: sendDraftHandler,
  },
  {
    name: "manage_message",
    description: manageMessageDescription,
    inputSchema: manageMessageSchema,
    // delete is soft, but move and delete both take mail out of where it was
    // and hand back new ids, so the same call cannot be repeated.
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
    handler: manageMessageHandler,
  },
  {
    name: "list_folders",
    description: listFoldersDescription,
    inputSchema: listFoldersSchema,
    annotations: READ_ONLY,
    handler: listFoldersHandler,
  },
  {
    name: "list_events",
    description: listEventsDescription,
    inputSchema: listEventsSchema,
    annotations: READ_ONLY,
    handler: listEventsHandler,
  },
  {
    name: "list_calendars",
    description: listCalendarsDescription,
    inputSchema: listCalendarsSchema,
    annotations: READ_ONLY,
    handler: listCalendarsHandler,
  },
  {
    name: "create_event",
    description: createEventDescription,
    inputSchema: createEventSchema,
    // Attendees are emailed an invitation the moment the event is created.
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    handler: createEventHandler,
  },
  {
    name: "manage_event",
    description: manageEventDescription,
    inputSchema: manageEventSchema,
    // Cancels the event, and notifies attendees or the organizer as it goes.
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true },
    handler: manageEventHandler,
  },
  {
    name: "search_contacts",
    description: searchContactsDescription,
    inputSchema: searchContactsSchema,
    annotations: READ_ONLY,
    handler: searchContactsHandler,
  },
  {
    name: "manage_contact",
    description: manageContactDescription,
    inputSchema: manageContactSchema,
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
    handler: manageContactHandler,
  },
  {
    name: "auto_reply",
    description: autoReplyDescription,
    inputSchema: autoReplySchema,
    // Nothing is destroyed by a setting, and setting the same reply twice is
    // the same state — but the reply it establishes is delivered to everyone
    // who writes to this account.
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    handler: autoReplyHandler,
  },
  {
    name: "mailbox_settings",
    description: mailboxSettingsDescription,
    inputSchema: mailboxSettingsSchema,
    // Working hours are visible to anyone scheduling with this account, but
    // the call itself sends nothing outward.
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    handler: mailboxSettingsHandler,
  },
  {
    name: "add_attachment",
    description: addAttachmentDescription,
    inputSchema: addAttachmentSchema,
    // The url source fetches from an arbitrary https host, and every call
    // adds another attachment rather than replacing one.
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    handler: addAttachmentHandler,
  },
  {
    name: "manage_rules",
    description: manageRulesDescription,
    inputSchema: manageRulesSchema,
    // A rule acts on all future mail with no per-message approval and can
    // soft-delete it; it stays open-world-free only because forwarding actions
    // are deliberately not offered.
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
    handler: manageRulesHandler,
  },
  {
    name: "manage_senders",
    description: manageSendersDescription,
    inputSchema: manageSendersSchema,
    // Blocking is undone by unblocking, and blocking an already-blocked
    // sender is the same state; the junk list never leaves the mailbox.
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    handler: manageSendersHandler,
  },
  {
    name: "create_folder",
    description: createFolderDescription,
    inputSchema: createFolderSchema,
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    handler: createFolderHandler,
  },
  {
    name: "manage_categories",
    description: manageCategoriesDescription,
    inputSchema: manageCategoriesSchema,
    // Deleting a category cannot be undone from here, and leaves its name on
    // every message that already carries it.
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
    handler: manageCategoriesHandler,
  },
  {
    name: "list_tasks",
    description: listTasksDescription,
    inputSchema: listTasksSchema,
    annotations: READ_ONLY,
    handler: listTasksHandler,
  },
  {
    name: "manage_task",
    description: manageTaskDescription,
    inputSchema: manageTaskSchema,
    // delete is PERMANENT on To Do — there is no Deleted Items to recover from.
    annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
    handler: manageTaskHandler,
  },
  {
    name: "check_new_mail",
    description: checkNewMailDescription,
    inputSchema: checkNewMailSchema,
    // Not read-only: every successful call advances the stored delta position,
    // which is exactly why a repeat does not report the same changes twice.
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    handler: checkNewMailHandler,
  },
  {
    name: "get_mailbox_activity",
    description: getMailboxActivityDescription,
    inputSchema: getMailboxActivitySchema,
    annotations: READ_ONLY,
    handler: getMailboxActivityHandler,
  },
  {
    name: "manage_auto_filing",
    description: manageAutoFilingDescription,
    inputSchema: manageAutoFilingSchema,
    // A switch, not a deletion — but the feature it switches on sends the
    // subject and a body excerpt of arriving mail to the Anthropic API, and
    // add_skip_pattern appends rather than sets.
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    handler: manageAutoFilingHandler,
  },
  {
    name: "get_auto_filing_log",
    description: getAutoFilingLogDescription,
    inputSchema: getAutoFilingLogSchema,
    annotations: READ_ONLY,
    handler: getAutoFilingLogHandler,
  },
  {
    name: "get_health",
    description: getHealthDescription,
    inputSchema: getHealthSchema,
    // Reads the stored heartbeat (remote) or performs Graph GETs (local); the
    // KV probe and the alert draft belong to the CRON, not to this tool.
    annotations: READ_ONLY,
    handler: getHealthHandler,
  },
];

/** Register every tool and prompt on a freshly constructed McpServer. */
export function registerAll(server: McpServer): void {
  for (const tool of TOOLS) {
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: tool.inputSchema,
        annotations: tool.annotations,
      },
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
