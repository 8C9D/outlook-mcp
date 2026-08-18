# Assumptions and decisions

Recorded while executing the "Personal-Outlook MCP foundation" task on 2026-08-18.

## Project location

- Chose `~/dev/outlook-mcp` because `~/dev` already exists and is the established personal-code directory (no `~/projects` or `~/code` present).

## Phase A — completed 2026-08-18 (after user signed up for Azure)

The initial run on 2026-08-18 was blocked (details preserved below).
The user then signed up for Azure with the personal account, which created a "Default Directory" tenant, and the portal work was completed there:

- Registration `outlook-mcp` created in the Default Directory tenant (tenant ID `a289df25-1318-4d0c-87ad-82dd2b4efafb`).
- Application (client) ID `1d362aa5-17bb-4731-828f-4e024be88f01`, written to `.env`.
- **Supported account types**: "Personal Microsoft account users" (`signInAudience: PersonalMicrosoftAccount`), verified on Overview and in the blade URL.
- **Allow public client flows**: set to Enabled under Authentication (Preview) → Settings, saved ("Successfully updated outlook-mcp"), and verified still Enabled after re-opening the blade.
- **API permissions** (all Delegated, verified in the portal after save): Calendars.ReadWrite, Mail.Read, Mail.ReadWrite, offline_access, User.Read — exactly the required list; **Mail.Send is not present**. User.Read was present by default; the other four were added.
- No admin consent granted; consent happens at first sign-in (personal accounts don't support admin-consent permissions anyway, per the portal's own notice).

## Phase A (first attempt) — BLOCKED: Entra app registration could not be created

The portal work was attempted with Claude in Chrome and is blocked by Microsoft-side restrictions, not by a missed step:

1. **Account chosen for the portal**: `arthur.yuhao.zhang@outlook.com` (the personal Microsoft account matching the user's identity), selected from the Azure portal account picker.
2. **Personal-account context** (consumer pseudo-tenant `f8cdef31-a31e-4b4a-93e4-5f571e91255a`): App registrations list is empty, and the **New registration** dialog refuses with "The ability to create applications outside of a directory has been deprecated", offering only "joining the M365 Developer Program" or "signing up for Azure". Portal settings → Directories confirms the account has **no directories** and no subscriptions.
3. **University of Waterloo tenant** (`[redacted-second-alias]`, also signed in): the App registrations blade returns **401 "You do not have access"** — the tenant blocks members from viewing/creating app registrations.
4. Both remedies Microsoft offers (M365 Developer Program signup, Azure signup) create new accounts/agreements, which the automation deliberately does not perform on the user's behalf. **The user must do one of these once**, then create the `outlook-mcp` registration (or re-run this task).

Consequences:

- No Application (client) ID exists; `.env` contains an empty `AZURE_CLIENT_ID` with a note.
- The portal-state record requested in Phase A step 6 (account type, public client flows, permission list) cannot be provided; the *target* state remains: personal Microsoft accounts allowed, public client flows = Yes, delegated permissions exactly Mail.Read, Mail.ReadWrite, Calendars.ReadWrite, User.Read, offline_access, and **no Mail.Send**.
- The verify runs (Phase B6) could not be executed against Graph; `npm run verify` currently fails fast with the intended "AZURE_CLIENT_ID is not set" message.

## Phase B6 — sign-in and verification (2026-08-18)

- The device-code sign-in hit Microsoft's "Verify your email" account-security challenge (code sent to the user's recovery Gmail). Per policy the automation does not complete identity-verification challenges, so the user completed the email verification — and, with it, the consent screen — themselves in the browser; the automation did not get to inspect the consent screen before approval. Mitigation: the app's registered permission list was independently verified in the portal (no Mail.Send), and delegated consent can never exceed the registered list, so the draft-only guarantee holds.
- First `npm run verify`: all three checks PASS (profile, inbox top-3, 7-day calendar view — "no events").
- Second `npm run verify`: completed with **no sign-in prompt** (silent acquisition from `.token-cache.json`, mode 0600) — completion criterion met.

## Phase B — implementation decisions

- Node v24.15.0 is installed, satisfying the "Node 20+" target; tsconfig targets ES2022 with NodeNext modules.
- `npm run verify` uses `tsx` (dev dependency) rather than a compiled build, since the project is scaffold-only; `npm run typecheck` runs `tsc --noEmit`.
- The token cache is written with mode 0600 at creation and re-`chmod`ed after every write, since `writeFile`'s `mode` only applies when the file is created.
- `getAccessToken` uses the first cached account if several exist — only one personal account is expected to ever sign in.
- Calendar check window is computed as now → now+7d in UTC ISO strings (URL-encoded), with `Prefer: outlook.timezone="America/Toronto"` controlling the *returned* event times, per the spec.
- Initialized a plain local git repository (no remote), since the spec's gitignore requirements imply version control.
- `.env` is committed to nothing (gitignored) but was still created with an empty value so the fail-fast path and the file's expected format are in place.

## MCP layer

Recorded while executing the "MCP layer and tools" task on 2026-08-18.

- **SDK/zod versions**: installed `@modelcontextprotocol/sdk@1.30.0` and `zod@4.4.3` (current latest); tools are registered via `McpServer.registerTool` with zod raw shapes as `inputSchema`.
- **stdout audit**: the foundation's only stdout writes were `console.log` diagnostics in `auth.ts`, reachable only when a token is actually acquired/refreshed interactively. They were still all moved to `console.error` (including the device-code prompt, which remains visible in an interactive terminal via stderr), so no import or call path in server mode can write to stdout.
- **Server-mode auth**: `getAccessTokenSilent()` throws `AuthRequiredError` carrying the exact message required by the spec plus the underlying MSAL reason in parentheses; `runTool()` converts it (and every other failure) into an `isError` tool result, so tool calls can never crash the server process.
- **Graph error surface**: non-2xx responses raise `GraphError` (status + raw body); tool results show `code: message` parsed from the JSON error body when available, otherwise the raw body.
- **429 handling** lives in the shared Graph layer (applies to every tool): honor `Retry-After` once (defaulting to 2s when the header is absent or unparsable, capped at 60s), then fail with a readable throttling message.
- **search_mail**: the KQL term is wrapped in double quotes with embedded `"` and `\` backslash-escaped, then the whole `$search` value is `encodeURIComponent`-ed. Folder input is passed to `/me/mailFolders/{folder}/messages` unvalidated — Graph accepts well-known names and folder ids alike and returns a clear error otherwise. Body previews are collapsed to one line and capped at 150 chars.
- **read_thread**: tries `$filter` + `$orderby=receivedDateTime` first and falls back to a local sort only when Graph rejects the combination with an `InefficientFilter` error body, per the spec. `max_messages` is schema-capped at 100. Quoted-tail trimming triggers only on three conservative markers — `--- Original Message ---` lines, a lone `On … wrote:` line, or an Outlook `____` divider followed within 3 lines by a `From:` header — and never on the first line, replacing the tail with `[quoted reply trimmed]`.
- **create_draft mode validation**: "exactly one mode" is interpreted as: `reply_to_message_id` XOR (`to` and/or `subject` present); providing both modes, neither, or an incomplete new-message pair is an `isError` result. `cc` is honored in both modes (in reply mode it is appended to any recipients the reply already carries).
- **Reply-draft body**: after `createReply`, the draft's auto-quoted body is re-fetched as text (`Prefer: outlook.body-content-type="text"`) and the draft is PATCHed to contentType Text with the user's text, a blank line, then the quoted original — keeping the whole draft plain text, consistent with "body treated as plain text".
- **list_events**: "today" is computed in America/Toronto (not the machine's zone). The window is [start_date 00:00, start_date+days 00:00) as naive datetimes, which `calendarView` interprets in the `Prefer: outlook.timezone` zone. Events are paged via `@odata.nextLink` with a 500-event safety cap. Empty days are omitted by construction (days are grouped from returned events only).
- **create_event datetimes**: naive inputs are sent with `timeZone: "America/Toronto"` as required; inputs carrying an explicit offset/Z are converted to UTC and sent with `timeZone: "UTC"` (an equivalent instant — the spec only fixes the interpretation of offset-less input). Date-only start is accepted (00:00). All-day events are normalized to midnight-to-midnight Toronto dates with a minimum span of one day, and `end <= start` dates are corrected to start+1 day.
- **Default event duration** (+60 min) is wall-clock arithmetic on the naive datetime, so an event starting at 23:30 ends 00:30 the next day, and DST transitions keep the displayed times intuitive.
- **Harness search term**: the longest alphanumeric token (≥3 chars) of the latest inbox message's subject, falling back to the whole subject; asserts at least one result and conversation ids present.
- **Harness cleanup**: Graph `DELETE` soft-deletes into Deleted Items (and `/me/messages` includes Deleted Items), so after the spec's DELETE + 404 confirmation, the final artifact check first `permanentDelete`s any remaining `[MCP TEST]` messages, then asserts zero `[MCP TEST]` messages across the mailbox and zero `[MCP TEST]` events. Deleted events are not reachable via `/me/events`, so no purge is needed on the calendar side.
- **Stdio smoke test**: speaks newline-delimited JSON-RPC (protocolVersion 2025-06-18) to a spawned `tsx src/server.ts`, asserts the five expected tool names with object input schemas and non-trivial descriptions, and asserts every non-empty stdout line parses as a `jsonrpc: "2.0"` message.
- **Build script**: `tsconfig.json` keeps `noEmit` for typechecking; `npm run build` uses `tsconfig.build.json` (extends base, emits to `dist/`, gitignored).

## Desktop wiring

Recorded while executing the "harden for headless launch and wire into Claude Desktop" task on 2026-08-18.

- **cwd audit findings**: three cwd-sensitive sites existed — (1) `import "dotenv/config"` in `auth.ts`, which reads `.env` from `process.cwd()` (the only genuinely broken one under an arbitrary cwd); (2) `auth.ts`'s token-cache path and (3) `server.ts`'s package.json read, both derived via a fixed `..` hop from the module directory (correct in both src/ and dist/ layouts by construction, but replaced anyway with the mandated walk-up discovery). `test-tools.ts` had the same `..` pattern (dev-only) and was switched to the shared module for consistency. No `process.cwd()` calls or other relative fs paths exist in `src/`.
- **Root discovery**: `src/project-root.ts` walks up from `import.meta.url`'s directory to the first directory containing `package.json`. `tsc` emits a flat `dist/` (rootDir inferred as `src/`), so both `src/x.ts` and `dist/x.js` sit exactly one level below the root; the walk also survives any future change in emit depth. The walk cannot false-positive on a `node_modules` package.json because it only ascends.
- **dotenv quiet flag**: installed dotenv is 16.6.1, whose newer releases can print an "injecting env" tip via `console.log`; `quiet: true` is passed so stdout can never carry non-protocol bytes. (Verified empirically clean, and the harness's clean-stdout assertion guards it.)
- **Proof of cwd-independence**: a throwaway script spawned the server with `cwd=/tmp` in both modes (`tsx src/server.ts`, `node dist/server.js`) and asserted initialize + tools/list + a real `tools/call list_events` (using the cached token) all succeed with JSON-RPC-only stdout. All six assertions passed; `npm run test:tools` still passes 7/7.
- **Node binary**: `which node` → `/Users/arthurzhang/.nvm/versions/node/v24.15.0/bin/node` (nvm-managed), hardcoded in the config as instructed; README documents that an nvm-driven node change invalidates this path.
- **Config edit**: existing `claude_desktop_config.json` (1590 bytes, no `mcpServers` key, mode 0600) was backed up to `claude_desktop_config.json.bak-20260818-084934` (same directory), then the `mcpServers.outlook` block was inserted surgically at the top of the object. Verified: the result parses as JSON, and stripping the six inserted lines reproduces the backup byte-for-byte; file mode remains 0600.
- **dist/ stays gitignored**: the build is reproducible via `npm run build`; committing artifacts wasn't requested. Consequence: after a fresh clone or `git clean`, `npm run build` must be re-run before Claude Desktop can start the server.
- **No `env` block in the config**: the server locates `.env` itself (section 1), so none is needed — matching the task's instruction.

## v2 full parity

Recorded while executing the "v2 full-parity tool surface" task on 2026-08-18.

### Phase A (portal)
- The three new delegated permissions were added in the portal via Claude in Chrome; the saved list was verified to be exactly: Calendars.ReadWrite, Contacts.ReadWrite, Mail.Read, Mail.ReadWrite, Mail.Send, MailboxSettings.ReadWrite, offline_access, User.Read ("Successfully saved permissions for outlook-mcp").
- The portal's own warning ("users will have to consent even if they've already done so previously") confirms Phase B's re-consent requirement.

### Phase B (re-consent)
- The scope list lives only in `src/auth.ts` (`SCOPES`); `src/login.ts` reuses it via `getAccessToken`, so one edit covers "auth.ts and the login script".
- `.token-cache.json` was deleted before login, exactly as specified.
- The device-code sign-in again hit Microsoft's "Verify your email" account-security challenge (code sent to the recovery Gmail), as in v1. Per policy the automation does not complete identity-verification challenges — the user completes that step; the consent screen is verified before approval per the task spec (or, if the user approves it themselves, the granted scopes are verified afterwards from the token).

### Tool-count discrepancy in the spec
- Phase E says "old five + new nine", but Phase D defines ten new tools (read_message, get_attachment, update_draft, send_draft, manage_message, list_folders, manage_event, search_contacts, manage_contact, auto_reply). All ten were implemented; the stdio test asserts the real 15-tool list.

### Design decisions
- **create_draft forward mode**: `to` is allowed (and usually wanted) with `forward_message_id`; `subject` is not (it comes from the original). `reply_all` is rejected outside reply mode. Reply/forward bodies are built exactly like v1 replies: createReply/createReplyAll/createForward, re-fetch the quoted body as text, PATCH the user's text above it.
- **send_draft** verifies `isDraft` via Graph before POST `/me/messages/{id}/send`, and also rejects drafts with no To recipients (Graph would 400 anyway; the tool gives a clearer message). Output notes the message id changes on send.
- **update_draft** verifies `isDraft` before PATCHing; an empty `cc: []` clears the CC list (documented in the schema).
- **manage_message** loops per id with per-id try/catch; the result lists OK/FAILED per message. The overall result is `isError` only when every message failed. Moves report the new message id Graph assigns.
- **list_folders** shows two levels (top + children) and notes deeper subfolders exist without walking them; an `include_hidden` flag exists mostly so the schema is non-empty (MCP clients handle empty schemas inconsistently).
- **manage_event cancel** picks the Graph call by role and attendee count: organizer with attendees → POST `/cancel` (notifies attendees); organizer without attendees → DELETE (soft, to Deleted Items — Graph's `/cancel` is meaningless without attendees); attendee → POST `/decline` with sendResponse. This matches the spec's "attendee's cancel is a decline".
- **search_contacts** uses `startswith` $filter on displayName/givenName/surname (ORed). Graph's contacts `$search` is inconsistently supported on consumer accounts; prefix filtering is reliable and enough for name lookup.
- **manage_contact** stores `phones` as `businessPhones` (Graph's general-purpose phone list); update semantics are replace-not-append, same as draft recipients.
- **auto_reply set** uses `alwaysEnabled` when no window is given and `scheduled` with America/Toronto datetimes when both start and end are given (providing only one is an error). `externalAudience` is set to `all` so the external message actually reaches strangers; get strips HTML tags from stored messages for readable output.
- **get_attachment** only downloads `fileAttachment`s (item/reference attachments have no content bytes → clear isError). Filenames are sanitized (no path separators/control chars) and collision-suffixed " (1)", " (2)", … in `~/Downloads/outlook-mcp-attachments/`.
- **read_message** truncates bodies at 10 000 chars (single-message reads warrant more than read_thread's 2 000-per-message cap).
- **Version bumped to 0.2.0**; package description updated (it claimed "never sends mail").

### Harness decisions (Phase E)
- v1 tests for search_mail/read_thread/list_events are kept (renamed v1a/v1b/v1d); v1's standalone create_draft test is superseded by the draft-lifecycle test (create → update → send), which exercises create_draft's new-message mode end-to-end.
- The received test message is located by exact-subject $filter polling (5 s interval, 60 s cap) rather than $search, which has indexing lag.
- Test c reuses the received copy from test b (per the spec's ordering), and its cleanup soft-deletes both the received and sent copies via manage_message, then permanently deletes the test folder (test-only purge).
- An extra test (b2) asserts send_draft refuses a non-draft id — the guardrail the design leans on.
- auto_reply restore PATCHes back the exact saved `automaticRepliesSetting` object (not just "clear"), and the final sweep re-checks status and that no "[MCP TEST]" text remains in the reply messages.
- Deleted test contacts are soft-deleted (policy: no purge in the tool surface); the sweep asserts they no longer appear under `/me/contacts`. Graph exposes no supported purge for contacts' Deleted Items, so the recoverable copy in trash is accepted as clean.
- purgeTestFolders also checks `deleteditems`' child folders so a soft-deleted test folder can't linger.
