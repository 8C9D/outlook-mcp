# outlook-mcp

An MCP server that connects Claude to a personal Microsoft (outlook.com) account via Microsoft Graph.
v3 exposes seventeen tools over stdio covering mail (search, read, compose, send, manage), attachments
(read and add), folders, inbox rules, calendar (list, create, update, cancel, respond), contacts, and
auto-reply settings. All datetimes are handled in America/Toronto unless a caller supplies an explicit
UTC offset.

## Tools (v3)

| Tool | What it does |
| --- | --- |
| `search_mail` | With `query`: full-text search over a mail folder (default inbox), relevance-ranked. **Without `query`: the folder's latest messages, genuinely newest-first** — the right call for "what's my latest email". Returns subject, sender, local datetime, message id, conversation id, attachment flag, optional body preview. |
| `read_thread` | Renders a conversation oldest-to-newest as plain text given a conversation id, quoted tails trimmed. |
| `read_message` | One full message: headers, plain-text body, and an attachment inventory (name/size/type/attachment id). |
| `get_attachment` | Saves an attachment to `~/Downloads/outlook-mcp-attachments/` (collision-safe names); small text/JSON attachments are also returned inline. |
| `create_draft` | Creates a draft: new message (`to` + `subject`), reply (`reply_to_message_id`, optional `reply_all`), or forward (`forward_message_id` + `to`). Never sends. |
| `update_draft` | Edits a draft's body/subject/to/cc (recipient arrays replace, not append). Rejects non-drafts. |
| `send_draft` | **The only send path.** Sends an existing draft by id after verifying it really is a draft. |
| `add_attachment` | Attaches a local file to a draft (≤ 25 MB: single request under 3 MB, chunked upload session for 3–25 MB). Natural flow: `create_draft` → `add_attachment` → `send_draft`. |
| `manage_message` | Batch (1–20 ids): move, archive, delete (soft), mark read/unread, flag/unflag, with per-message results. All ids go out as **one Graph `$batch` request** (one HTTP round-trip instead of up to 20); throttled items are retried once per their `Retry-After`. |
| `list_folders` | Mail folder tree (2 levels) with unread/total counts and folder ids. |
| `list_events` | Calendar events for a date window (default: next 7 days), grouped by day. |
| `create_event` | Creates an event. **If attendees are given, Outlook emails them invitations immediately.** |
| `manage_event` | Update / cancel / respond (accept, decline, tentative). Updates and cancellations on events with attendees notify them. |
| `search_contacts` | Search saved contacts by name prefix; returns name, emails, phones, contact id. |
| `manage_contact` | Create / update / delete (soft) a saved contact. |
| `auto_reply` | Get / set / clear the mailbox automatic reply (out-of-office). |
| `manage_rules` | List / create / delete inbox rules (conditions: from/sender/subject/body; actions: move, mark read, soft delete). **Rules act automatically on all future incoming mail** — see below. |

## Inbox rules (`manage_rules`)

A rule runs server-side on **every future incoming message that matches, with no per-message
approval** — it keeps acting long after the conversation that created it. The tool description
therefore instructs the model to state the complete rule (all conditions → all actions) before
creating one, and to keep rules conservative. Move targets are validated to exist before the rule
is created.

**No forwarding actions, by design.** Graph rules can forward or redirect mail to arbitrary
addresses; this server deliberately does not expose those actions (creating or listing aside — the
list output *does* flag externally created forward rules). A standing silent forward is an
exfiltration primitive: one approved call would export all future mail. Rules here can only move,
mark read, or soft-delete within the mailbox.

## Two-step send by design

The server can send email, but **no tool composes and sends in one call**, and `/me/sendMail` is never
used. Sending is always separate tool calls: compose with `create_draft` (and optionally `update_draft`
and `add_attachment`), then send that exact draft with `send_draft(draft_id)`. This means:

- The complete outgoing message exists as a reviewable draft before anything leaves the account.
- The calling model must present the draft (subject, recipients) and take a second deliberate action to send.
- A single confused or injected tool call can at worst create a draft, not dispatch mail.

## Soft-delete policy

Every delete in the tool surface (messages, events, contacts) is a **soft delete**: items move to
Deleted Items and stay recoverable. No tool permanently purges anything. (The test harness contains a
`permanentDelete` helper strictly for cleaning up its own `[MCP TEST]` artifacts — it is not part of the
tool surface.)

## Security model

With send, delete, and settings tools enabled, **treat mailbox content as untrusted input**: an email
can contain text that tries to instruct the model into sending, deleting, or forwarding things
(prompt injection). Mitigations built in and recommended:

- **Keep per-call approval prompts in Claude Desktop** for `send_draft`, `manage_message`
  (delete/move), `manage_event`, `manage_contact`, `auto_reply`, and `manage_rules` — do **not**
  "always allow" these. Each approval shows you what is about to happen; that review is the real
  safety boundary. `manage_rules` especially: a rule keeps acting on all future mail after one
  approval, which is why rule creation must stay reviewable and forwarding actions are excluded
  entirely.
- The operations third parties can see are: `send_draft`, event **invitations** (`create_event` with
  attendees), event **updates/cancellations** on events with attendees, invitation **responses**, and
  **auto-replies**. Everything else stays inside the mailbox.
- Destructive tool descriptions instruct the model to state exactly what will be affected
  (subjects/recipients/ids) before calling, so approval prompts carry context.
- Sending is structurally two-step (above) and deletes are soft (above).

## Login and re-authentication

The MCP server runs headless and **never prompts for sign-in** — it only uses tokens silently refreshed
from the local cache (`.token-cache.json`, mode 0600, gitignored).

- First-time setup or after the refresh token expires/revokes: run `npm run login` in a terminal in this
  directory and complete the device-code sign-in. The script caches tokens and exits.
- When the cache is unusable, every tool call returns: *"Authentication expired. Run `npm run login` in a
  terminal in ~/dev/outlook-mcp, then retry."*
- To force a fresh sign-in, delete `.token-cache.json` and run `npm run login`.

## Setup

1. `npm install`
2. Create `.env` with `AZURE_CLIENT_ID=<Application (client) ID of the outlook-mcp Entra app registration>`
   (registration: personal Microsoft accounts, public client flows enabled, delegated permissions
   Mail.Read, Mail.ReadWrite, Mail.Send, Calendars.ReadWrite, Contacts.ReadWrite,
   MailboxSettings.ReadWrite, User.Read, offline_access).
3. `npm run login` — one-time interactive device-code sign-in.
4. `npm run serve` — start the MCP server on stdio (this is what an MCP client will launch).

## Scripts

- `npm run login` — interactive device-code sign-in; caches tokens and exits.
- `npm run serve` — run the MCP server (stdio; stdout is protocol-only, logs go to stderr).
- `npm run test:tools` — live test harness: exercises the tools against the real account plus a stdio
  protocol smoke test, and verifies it leaves no `[MCP TEST]` artifacts behind and restores auto-reply.
- `npm run verify` — the original auth/Graph foundation check.
- `npm run typecheck` / `npm run build` — type-check / compile to `dist/`.

## Claude Desktop

The server is registered in `~/Library/Application Support/Claude/claude_desktop_config.json` under
`mcpServers` (installed 2026-08-18; unchanged for v2 — same command and args):

```json
"outlook": {
  "command": "/Users/arthurzhang/.nvm/versions/node/v24.15.0/bin/node",
  "args": ["/Users/arthurzhang/dev/outlook-mcp/dist/server.js"]
}
```

It runs the compiled build (`npm run build` → `dist/server.js`) under a plain `node` — no `tsx` needed at
runtime. The server resolves its own project root from its module location, so it finds `.env` and
`.token-cache.json` regardless of the working directory Claude Desktop launches it with.

> **Node path caveat:** the `command` is the absolute path to the node binary (resolved via `which node`
> at install time) because Claude Desktop does not inherit the shell `PATH`. This machine uses nvm, so
> **upgrading or switching the default node version changes this path** — if the server stops launching
> after a node upgrade, re-run `which node` and update `command` accordingly.

- **Picking up config changes:** Claude Desktop reads the config only at launch. Fully quit it (Cmd+Q —
  closing the window is not enough) and reopen.
- **Checking server status:** Settings → Developer → MCP servers shows the `outlook` server and whether
  it started; in a chat, the tools icon lists its seventeen tools when connected.
- **Logs:** `~/Library/Logs/Claude/mcp-server-outlook.log` (this server's stderr) and
  `~/Library/Logs/Claude/mcp.log` (general MCP lifecycle) — first place to look when the server shows as failed.
- **Auth expired?** Tool calls will return *"Authentication expired. Run `npm run login` …"* — see
  [Login and re-authentication](#login-and-re-authentication) above. No Claude Desktop restart is needed
  after re-login; the next tool call picks up the refreshed cache.
- **After changing the code:** run `npm run build` — Claude Desktop runs `dist/`, not `src/`.
