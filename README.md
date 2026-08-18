# outlook-mcp

An MCP server that connects Claude to a personal Microsoft (outlook.com) account via Microsoft Graph.
It exposes five tools over stdio: mail search, thread reading, draft creation, and calendar listing/creation.
All datetimes are handled in America/Toronto unless a caller supplies an explicit UTC offset.

## Tools

| Tool | What it does |
| --- | --- |
| `search_mail` | Full-text search over a mail folder (default inbox). Returns subject, sender, local datetime, message id, conversation id, attachment flag, and an optional body preview per hit. Results are relevance-ranked (Graph `$search`). |
| `read_thread` | Renders a conversation oldest-to-newest as plain text given a conversation id, with quoted-reply tails conservatively trimmed and long bodies truncated at 2000 chars. |
| `create_draft` | Creates a draft in the Drafts folder — either a fresh message (`to` + `subject`) or a reply (`reply_to_message_id`) with the new text above the quoted original. Never sends. |
| `list_events` | Lists calendar events for a date window (default: next 7 days), grouped by day, all-day events first. |
| `create_event` | Creates a calendar event (default duration 60 minutes; all-day supported). **If attendees are given, Outlook emails them invitations immediately on creation** — the one write in this server that notifies third parties. |

## Draft-only by design (no Mail.Send)

This server can never send email:

- The Entra app registration has **no `Mail.Send` permission**, so any send attempt would be rejected by Graph.
- The codebase contains **no call to `/send` or `/me/sendMail`** at all. `create_draft` saves to the Drafts folder; you review and send from Outlook yourself.

## Login and re-authentication

The MCP server runs headless and **never prompts for sign-in** — it only uses tokens silently refreshed from the local cache (`.token-cache.json`, mode 0600, gitignored).

- First-time setup or after the refresh token expires/revokes: run `npm run login` in a terminal in this directory and complete the device-code sign-in. The script caches tokens and exits.
- When the cache is unusable, every tool call returns: *"Authentication expired. Run `npm run login` in a terminal in ~/dev/outlook-mcp, then retry."*
- To force a fresh sign-in, delete `.token-cache.json` and run `npm run login`.

## Setup

1. `npm install`
2. Create `.env` with `AZURE_CLIENT_ID=<Application (client) ID of the outlook-mcp Entra app registration>` (registration: personal Microsoft accounts, public client flows enabled, delegated permissions Mail.Read, Mail.ReadWrite, Calendars.ReadWrite, User.Read, offline_access — and no Mail.Send).
3. `npm run login` — one-time interactive device-code sign-in.
4. `npm run serve` — start the MCP server on stdio (this is what an MCP client will launch).

## Scripts

- `npm run login` — interactive device-code sign-in; caches tokens and exits.
- `npm run serve` — run the MCP server (stdio; stdout is protocol-only, logs go to stderr).
- `npm run test:tools` — live test harness: exercises all five tools against the real account plus a stdio protocol smoke test, and verifies it leaves no `[MCP TEST]` artifacts behind.
- `npm run verify` — the original auth/Graph foundation check.
- `npm run typecheck` / `npm run build` — type-check / compile to `dist/`.

## Claude Desktop

The server is registered in `~/Library/Application Support/Claude/claude_desktop_config.json` under `mcpServers` (installed 2026-08-18):

```json
"outlook": {
  "command": "/Users/arthurzhang/.nvm/versions/node/v24.15.0/bin/node",
  "args": ["/Users/arthurzhang/dev/outlook-mcp/dist/server.js"]
}
```

It runs the compiled build (`npm run build` → `dist/server.js`) under a plain `node` — no `tsx` needed at runtime. The server resolves its own project root from its module location, so it finds `.env` and `.token-cache.json` regardless of the working directory Claude Desktop launches it with.

> **Node path caveat:** the `command` is the absolute path to the node binary (resolved via `which node` at install time) because Claude Desktop does not inherit the shell `PATH`. This machine uses nvm, so **upgrading or switching the default node version changes this path** — if the server stops launching after a node upgrade, re-run `which node` and update `command` accordingly.

- **Picking up config changes:** Claude Desktop reads the config only at launch. Fully quit it (Cmd+Q — closing the window is not enough) and reopen.
- **Checking server status:** Settings → Developer → MCP servers shows the `outlook` server and whether it started; in a chat, the tools icon lists its five tools when connected.
- **Logs:** `~/Library/Logs/Claude/mcp-server-outlook.log` (this server's stderr) and `~/Library/Logs/Claude/mcp.log` (general MCP lifecycle) — first place to look when the server shows as failed.
- **Auth expired?** Tool calls will return *"Authentication expired. Run `npm run login` …"* — see [Login and re-authentication](#login-and-re-authentication) above. No Claude Desktop restart is needed after re-login; the next tool call picks up the refreshed cache.
- **After changing the code:** run `npm run build` — Claude Desktop runs `dist/`, not `src/`.
