// MCP server entry point (stdio transport). stdout carries JSON-RPC only;
// all logging goes to stderr. This process never initiates interactive auth —
// tools fail with a "run `npm run login`" message when the token cache is unusable.
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { searchMailDescription, searchMailHandler, searchMailSchema } from "./tools/search-mail.js";
import { readThreadDescription, readThreadHandler, readThreadSchema } from "./tools/read-thread.js";
import {
  createDraftDescription,
  createDraftHandler,
  createDraftSchema,
} from "./tools/create-draft.js";
import { listEventsDescription, listEventsHandler, listEventsSchema } from "./tools/list-events.js";
import {
  createEventDescription,
  createEventHandler,
  createEventSchema,
} from "./tools/create-event.js";

const PROJECT_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
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
  "create_draft",
  { description: createDraftDescription, inputSchema: createDraftSchema },
  createDraftHandler
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

await server.connect(new StdioServerTransport());
console.error(`outlook MCP server v${version} ready (stdio).`);
