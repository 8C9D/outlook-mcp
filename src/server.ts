// MCP server entry point (stdio transport). stdout carries JSON-RPC only;
// all logging goes to stderr. This process never initiates interactive auth —
// tools fail with a "run `npm run login`" message when the token cache is unusable.
//
// The tool and prompt surface lives in core/registry.js and is shared with the
// Cloudflare Worker entry point (worker/index.ts); this file is only the
// stdio-specific wiring plus the MSAL/disk-cache token source.
import { readFileSync } from "node:fs";
import path from "node:path";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { installMsalTokenProvider } from "./auth.js";
import { createMcpServer } from "./core/registry.js";
import { PROJECT_ROOT } from "./project-root.js";
import { installFileStateStore } from "./state-file.js";

const { version } = JSON.parse(
  readFileSync(path.join(PROJECT_ROOT, "package.json"), "utf8")
) as { version: string };

installMsalTokenProvider();
// check_new_mail's delta position lives in a gitignored JSON file here; on the
// Worker the same keys are KV entries.
installFileStateStore();

const server = createMcpServer(version);

await server.connect(new StdioServerTransport());
console.error(`outlook MCP server v${version} ready (stdio).`);
