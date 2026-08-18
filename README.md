# outlook-mcp

An MCP server that connects Claude to a personal Microsoft (outlook.com) account via Microsoft Graph.
v5 keeps v4's twenty-one tools and two prompts and serves them over **two transports** — the local
stdio server and a remote Cloudflare Worker (see [Remote deployment](#remote-deployment)) — covering mail (search, read, compose, send,
manage, categorize), attachments (read and add), folders (list and create), inbox rules, categories,
calendar (list, create, update, cancel, respond), contacts, Microsoft To Do tasks, and auto-reply
settings. All datetimes are handled in America/Toronto unless a caller supplies an explicit UTC offset.

## Tools (v4)

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
| `manage_message` | Batch (1–20 ids): move, archive, delete (soft), mark read/unread, flag/unflag, categorize, with per-message results. `categorize` **replaces** a message's categories rather than appending, and validates every name against the mailbox's category list first. All ids go out as **one Graph `$batch` request** (one HTTP round-trip instead of up to 20); throttled items are retried once per their `Retry-After`. |
| `list_folders` | Mail folder tree (2 levels) with unread/total counts and folder ids. |
| `create_folder` | Creates a mail folder at the mailbox root or under `parent_folder`. Rejects a duplicate name at the same level, naming the existing folder's id. |
| `list_events` | Calendar events for a date window (default: next 7 days), grouped by day. |
| `create_event` | Creates an event. **If attendees are given, Outlook emails them invitations immediately.** |
| `manage_event` | Update / cancel / respond (accept, decline, tentative). Updates and cancellations on events with attendees notify them. |
| `search_contacts` | Search saved contacts by name prefix; returns name, emails, phones, contact id. |
| `manage_contact` | Create / update / delete (soft) a saved contact. |
| `auto_reply` | Get / set / clear the mailbox automatic reply (out-of-office). |
| `manage_rules` | List / create / **update (in place)** / delete inbox rules (conditions and **exceptions**: from/sender/subject/body; actions: move, mark read, soft delete). **Rules act automatically on all future incoming mail** — see below. |
| `manage_categories` | List / create / delete the mailbox's Outlook categories (Graph's fixed `preset0`–`preset24` palette). Applying them to mail is `manage_message`'s `categorize`. |
| `list_tasks` | Microsoft To Do tasks grouped overdue / today / upcoming / no due date (America/Toronto). Open tasks by default; `include_completed` and `due_within_days` narrow or widen it. |
| `manage_task` | Create / complete / reopen / update / **delete (permanent)** a To Do task. `linked_message_id` on create turns an email into a task, copying its subject, sender, and an Outlook link into the task notes. |

## Inbox rules (`manage_rules`)

A rule runs server-side on **every future incoming message that matches, with no per-message
approval** — it keeps acting long after the conversation that created it. The tool description
therefore instructs the model to state the complete rule (all conditions → all actions) before
creating one, and to keep rules conservative. Move targets are validated to exist before the rule
is created.

**Update in place, and exceptions (v4).** `manage_rules update` PATCHes an existing rule, keeping its id
and its position in the evaluation order — earlier versions could only delete and recreate, which moved
the rule to the end of the sequence and changed its id. `conditions`, `exceptions`, and `actions` are each
replaced wholesale by what a call passes and left untouched by what it omits, so a call that only narrows
the conditions cannot silently drop the actions. `exceptions` are carve-outs with the same fields as
conditions — matching mail the rule must *not* act on, the safe way to keep a broad rule from catching
the one sender it should leave alone; passing `exceptions: {}` clears them, and `enabled: false` parks a
rule without deleting it. Rules created outside this server keep showing their exceptions in `list`.

**No forwarding actions, by design.** Graph rules can forward or redirect mail to arbitrary
addresses; this server deliberately does not expose those actions (creating or listing aside — the
list output *does* flag externally created forward rules). A standing silent forward is an
exfiltration primitive: one approved call would export all future mail. Rules here can only move,
mark read, or soft-delete within the mailbox.

## Microsoft To Do notes

Tasks live in Microsoft To Do (Graph `/me/todo`), reached with the `Tasks.ReadWrite` scope added in v4.

- **Deletion is permanent.** Unlike mail, events, and contacts, a deleted To Do task does **not** land in
  a recoverable folder — Graph has no undelete for it. `manage_task`'s description says so explicitly and
  tells the model to name the task and get agreement first; `complete` is the non-destructive way to
  finish something while keeping the record.
- **Dates are America/Toronto.** `due_date` is an ISO date and `reminder` an ISO local datetime, both sent
  to Graph with an explicit `America/Toronto` zone. Graph stores them normalised to UTC, so reads pass
  `Prefer: outlook.timezone` to get local wall-clock back — that is what the overdue/today/upcoming
  grouping is computed from.
- **Lists.** `task_list` accepts a list name or id; omitted, it resolves to the account's `defaultList`.
  An unknown name fails with the available list names rather than a bare 404. For `complete`, `reopen`,
  `update`, and `delete`, `task_list` must be the list the task actually lives in — task ids are scoped
  to their list.
- **Email → task.** `manage_task(action: "create", linked_message_id: …)` appends the mail's subject,
  sender, received time, and `webLink` to the task notes. It copies a reference, not the message body,
  and never modifies the message.

## Prompts

Two MCP prompts ship with the server (visible in a client's prompt picker; zero-argument):

| Prompt | What it drives |
| --- | --- |
| `triage_inbox` | Read-only inbox triage: `list_folders`, then the **existing** rules via `manage_rules list`, then a `search_mail` get_latest sample. It then **proposes** batched moves and new/updated rules consistent with the rules already there — and is explicitly forbidden from calling any writing tool until the user approves specific proposals. |
| `morning_brief` | Start-of-day briefing from `search_mail` (get_latest, last 24 h), `list_events` (today), and `list_tasks` (`due_within_days: 3`), rendered as calendar / mail / tasks plus an **actions-needed** section. Read-only. |

Both are prompts, not automation: they instruct the calling model, and every write still goes through a
normal tool call with whatever approval the client enforces. `search_mail`'s listing does not carry
read/unread state, so `morning_brief` tells the model to call `read_message` rather than guess when that
distinction matters.

## Two-step send by design

The server can send email, but **no tool composes and sends in one call**, and `/me/sendMail` is never
used. Sending is always separate tool calls: compose with `create_draft` (and optionally `update_draft`
and `add_attachment`), then send that exact draft with `send_draft(draft_id)`. This means:

- The complete outgoing message exists as a reviewable draft before anything leaves the account.
- The calling model must present the draft (subject, recipients) and take a second deliberate action to send.
- A single confused or injected tool call can at worst create a draft, not dispatch mail.

## Soft-delete policy

Every *mailbox* delete in the tool surface (messages, events, contacts) is a **soft delete**: items move
to Deleted Items and stay recoverable, and no tool permanently purges them. The one exception is
`manage_task` delete: Microsoft To Do has no recoverable deleted-items store, so removing a task is
permanent (see [Microsoft To Do notes](#microsoft-to-do-notes)). (The test harness contains a
`permanentDelete` helper strictly for cleaning up its own `[MCP TEST]` artifacts — it is not part of the
tool surface.)

## Security model

With send, delete, and settings tools enabled, **treat mailbox content as untrusted input**: an email
can contain text that tries to instruct the model into sending, deleting, or forwarding things
(prompt injection). Mitigations built in and recommended:

- **Keep per-call approval prompts in Claude Desktop** for `send_draft`, `manage_message`
  (delete/move), `manage_event`, `manage_contact`, `auto_reply`, `manage_rules`, and `manage_task` —
  do **not** "always allow" these. Each approval shows you what is about to happen; that review is the real
  safety boundary. `manage_rules` especially: a rule keeps acting on all future mail after one
  approval, which is why rule creation must stay reviewable and forwarding actions are excluded
  entirely.
- The operations third parties can see are: `send_draft`, event **invitations** (`create_event` with
  attendees), event **updates/cancellations** on events with attendees, invitation **responses**, and
  **auto-replies**. Everything else stays inside the mailbox.
- Destructive tool descriptions instruct the model to state exactly what will be affected
  (subjects/recipients/ids) before calling, so approval prompts carry context.
- **`manage_task` delete is the one irreversible operation in the surface.** To Do has no recoverable
  deleted-items folder, so a deleted task cannot be restored by this server or by Outlook. Its tool
  description flags this and points the model at `complete` for the non-destructive case, but the
  approval prompt is the real backstop — keep it on.
- Sending is structurally two-step (above) and mailbox deletes are soft (above).
- **The remote endpoint is single-user.** Nothing anonymous can reach `/mcp` or any route that
  touches Graph, and only one Microsoft identity — matched on the Graph `/me` id or UPN captured at
  setup — can complete an authorization. A remote connector runs the same tools with the same
  approval expectations; the caveats above apply there too, and claude.ai's own tool-approval
  prompts are the equivalent safety boundary.

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
   MailboxSettings.ReadWrite, Tasks.ReadWrite, User.Read, offline_access).
3. `npm run login` — one-time interactive device-code sign-in.
4. `npm run serve` — start the MCP server on stdio (this is what an MCP client will launch).

## Scripts

- `npm run login` — interactive device-code sign-in; caches tokens and exits.
- `npm run serve` — run the MCP server (stdio; stdout is protocol-only, logs go to stderr).
- `npm run test:tools` — live test harness: exercises the tools against the real account plus a stdio
  protocol smoke test, and verifies it leaves no `[MCP TEST]` artifacts behind and restores auto-reply.
- `npm run verify` — the original auth/Graph foundation check.
- `npm run typecheck` / `npm run build` — type-check (both the Node and Worker configs) / compile to `dist/`.
- `npm run cf-types` — regenerate `worker-configuration.d.ts` after editing `wrangler.jsonc`.
- `npm run seed:kv` — push the current Microsoft refresh token from `.token-cache.json` into Workers KV.
- `npm run deploy` — deploy the Worker to Cloudflare.
- `npm run test:remote` — live tests against the deployed endpoint (discovery, anonymous rejection, a
  full OAuth exchange, an MCP round-trip, refresh-token rotation); cleans up every KV record it creates.

## Remote deployment

The same 21 tools and 2 prompts are also served over MCP Streamable HTTP from a Cloudflare Worker, so
claude.ai can reach the mailbox as a custom connector without this laptop being on.

**Deployed endpoint:** `https://outlook-mcp.arthur-yuhao-zhang.workers.dev/mcp`

### Architecture

Everything transport-independent lives under `src/core/`: `registry.ts` (the tool and prompt table),
`graph.ts` (the Graph transport), `prompts.ts`, and `token.ts`. Both entry points build the *same*
`McpServer` from `createMcpServer()`, so the two hosts cannot drift apart — `src/test-remote.ts`
asserts the deployed tool list equals the local registry.

```
src/core/*            transport-agnostic: registry, Graph calls, prompts, token indirection
src/tools/*           the 21 tool handlers (unchanged by transport)
src/server.ts         stdio entry  -> MSAL + .token-cache.json      (Claude Desktop)
src/worker/index.ts   Worker entry -> OAuth + refresh token in KV   (claude.ai)
```

The tool layer never knows where its Graph token comes from. `core/token.ts` holds a *token provider*
that each host installs: the stdio server installs MSAL silent acquisition; the Worker installs a
KV-backed provider, scoped per request with `AsyncLocalStorage`.

The Worker is **stateless** — no Durable Objects. Each POST builds a fresh `McpServer` and a
`WebStandardStreamableHTTPServerTransport` with `sessionIdGenerator: undefined`, and discards them
when the response is written.

`@cloudflare/workers-oauth-provider` fronts the whole thing. It owns discovery metadata, dynamic
client registration, PKCE, the token endpoint and bearer validation, and routes only authenticated
requests to `/mcp`. **Anonymous access is impossible** — an unauthenticated call to `/mcp` (by any
verb) gets `401` with a `WWW-Authenticate` challenge, which is what makes a client start the OAuth
flow. This is asserted by test `r3`.

### Single-user allowlist

Only one Microsoft identity may authorize a client. Authorization ends in a Graph `/me` call whose
result must match the `ALLOWED_MS_USER_ID` (Graph `/me` `id`) or `ALLOWED_MS_UPN` secret; anything
else is refused with `403`, and no grant is issued. The check lives in one place
(`isAllowedIdentity` in `src/worker/ms-token.ts`) and every authorization path runs through it.

Identity is proved with Microsoft's **device-code flow**, not a redirect-based auth code flow: the
Entra app registration is a public native client with no web redirect URI, and device code needs
none, so nothing about the registration had to change. `/authorize` shows a code to enter at
microsoft.com/devicelogin and polls until sign-in completes. That Microsoft token is used only to
read `/me` and is never stored.

### Token storage

The mailbox credential is a Microsoft refresh token in the `OUTLOOK_KV` namespace under
`ms:refresh_token`. MSAL Node does not run on workerd, so `src/worker/ms-token.ts` performs the
refresh-token grant directly against `https://login.microsoftonline.com/consumers/oauth2/v2.0/token`
with `fetch`, requesting exactly the scopes already consented (so no new consent is ever needed).
**Microsoft rotates the refresh token on every exchange and the new value is written back to KV**;
test `r12` proves this by forcing a refresh and comparing the stored value before and after.
Access tokens are cached under `ms:access_token` with a TTL so most calls skip the exchange.

Local stdio mode is untouched by all of this: it still uses MSAL and `.token-cache.json`. The two
credential chains are independent (Microsoft does not revoke an old refresh token when it issues a
new one), so the Worker rotating its copy does not disturb the local one.

### Setting it up from scratch

```bash
npx wrangler kv namespace create OUTLOOK_KV   # ids go in wrangler.jsonc
npx wrangler kv namespace create OAUTH_KV
npm run deploy                                 # prints the workers.dev URL

# Secrets — never committed. The client id is not really secret, but .env is
# gitignored and it stays out of the repo for consistency.
printf '%s' "<Entra application (client) id>"   | npx wrangler secret put MS_CLIENT_ID
printf '%s' "<Graph /me id>"                    | npx wrangler secret put ALLOWED_MS_USER_ID
printf '%s' "<userPrincipalName / mail>"        | npx wrangler secret put ALLOWED_MS_UPN

npm run seed:kv        # pushes the refresh token into KV; prints the two allowlist values
npm run test:remote    # 14 live checks against the deployed endpoint
```

`npm run seed:kv` reads `.token-cache.json`, so run `npm run login` first if the local cache is
stale. It hands the token to wrangler through a `0600` temp file rather than argv, and prints only a
SHA-256 fingerprint. Re-run it only after a fresh `npm run login` — at any other time it would
overwrite the Worker's rotated token with an older one.

If `resourceMetadata.resource` in `src/worker/index.ts` does not exactly match the URL pasted into
the client (path included), RFC 9728 discovery fails; update it if the Worker is ever renamed.

### Adding it to claude.ai as a custom connector

1. In claude.ai, open **Settings → Connectors** (Team/Enterprise: **Organization settings →
   Connectors**, added by an owner).
2. Click **Add custom connector**.
3. Paste the full URL **including the path**: `https://outlook-mcp.arthur-yuhao-zhang.workers.dev/mcp`
4. Leave the OAuth Client ID and Secret fields **empty** — the server supports dynamic client
   registration, so Claude registers itself.
5. Click **Add**, then **Connect**. A browser tab opens on the Worker's `/authorize` page showing a
   device code.
6. In another tab open the link shown on that page (`microsoft.com/devicelogin`), enter the code, and
   sign in **as arthur.yuhao.zhang@outlook.com**. Any other account is rejected.
7. The `/authorize` page polls, then returns to claude.ai. The connector's 21 tools are now available.

### Rotating and revoking access

- **Revoke one client** (disconnect claude.ai): remove the connector in claude.ai, then delete its
  records from the OAuth store — `npx wrangler kv key list --namespace-id <OAUTH_KV id> --remote`
  and `npx wrangler kv key delete <key> --namespace-id <OAUTH_KV id> --remote`. Deletions take up to
  a minute to propagate because KV caches reads at the edge.
- **Revoke everything at once**: delete `ms:refresh_token` from `OUTLOOK_KV`. Every tool call then
  fails with an auth error while the OAuth grants stay intact; `npm run seed:kv` restores service.
- **Cut off Microsoft entirely**: remove the app at
  <https://account.live.com/consent/Manage>. That kills the local cache and the Worker's KV token
  together; recover with `npm run login` followed by `npm run seed:kv`.
- **Rotate the mailbox credential**: `npm run login` then `npm run seed:kv`.
- **Take the endpoint down**: `npx wrangler delete` removes the Worker; the KV namespaces survive and
  must be deleted separately if you want the tokens gone.

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
  it started; in a chat, the tools icon lists its twenty-one tools when connected, and the prompt
  picker offers `triage_inbox` and `morning_brief`.
- **Logs:** `~/Library/Logs/Claude/mcp-server-outlook.log` (this server's stderr) and
  `~/Library/Logs/Claude/mcp.log` (general MCP lifecycle) — first place to look when the server shows as failed.
- **Auth expired?** Tool calls will return *"Authentication expired. Run `npm run login` …"* — see
  [Login and re-authentication](#login-and-re-authentication) above. No Claude Desktop restart is needed
  after re-login; the next tool call picks up the refreshed cache.
- **After changing the code:** run `npm run build` — Claude Desktop runs `dist/`, not `src/`.
