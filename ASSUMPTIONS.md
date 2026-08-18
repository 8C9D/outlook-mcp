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

### Phase B completion (2026-08-18)
- The device-code sign-in again required Microsoft's "Verify your email" challenge; the user completed the verification and the consent screen themselves (the automation entered the device code only). Mitigation for the unseen consent screen: the granted scopes were verified afterwards from the cached token's target — exactly User.Read, Mail.Read, Mail.ReadWrite, Mail.Send, Calendars.ReadWrite, Contacts.ReadWrite, MailboxSettings.ReadWrite (+ implicit openid/profile) — matching Phase A.
- Silent acquisition confirmed: `npm run verify` completed all three checks with no sign-in prompt immediately after login.
- Full harness: 13/13 PASS on the first run, including the final artifact sweep and exact auto-reply restoration.

## v3 batch 1

Recorded while executing the "v3 batch 1 — daily-use sharpening" task on 2026-08-18.

### Scope verification (before building manage_rules)
- Verified live before implementation: a silent-token GET of `/me/mailFolders/inbox/messageRules` succeeded (0 existing rules), confirming MailboxSettings.ReadWrite covers inbox rules with **no re-consent and no portal changes** — exactly as the task asserted. Entra registration, scopes, and the Claude Desktop config were not touched.

### Extension 1 — get_latest in search_mail
- No-query mode uses `$orderby=receivedDateTime desc` + `$top` through the same `fetchPaged` path (so `@odata.nextLink` still works if Graph pages) with the same Prefer-timezone header and output shape; the header line says "latest message(s) … (newest first)" instead of "result(s) … (relevance-ranked)" so the calling model can tell the modes apart.
- An empty folder in no-query mode returns "No messages in {folder}." (not an error), mirroring the no-matches case.

### Extension 2 — $batch in manage_message
- The 20-id schema cap equals Graph's $batch limit, so every call fits in exactly ONE batch request; ids map to batch item ids "0".."19".
- Per-item 429s are retried once in a **second** $batch containing only the throttled items, after waiting the LONGEST of their Retry-After values (case-insensitive header lookup, default 2 s, capped 60 s). Items still 429 after that are reported FAILED with a "wait a minute" message.
- A batch-level 429 (the /$batch POST itself throttled) is still handled by the shared Graph layer's single retry, unchanged.
- Success criteria per item: any 2xx status. Moves read the new message id from the item's response body, as before. An item id missing from Graph's response array (never observed) is reported FAILED rather than assumed OK.
- `graphRequestLog` was added to src/graph.ts (method + path per Graph call) so the harness can assert exactly one /$batch request; it is an append-only in-process array, unread in server mode, and its comment documents that the transport-level 429 retry does not add an entry.

### Extension 3 — add_attachment
- Guard order: absolute-path check → stat (missing/unreadable → isError) → regular-file check → **size stat BEFORE reading any bytes** (> 25 MB → isError naming the cap) → isDraft guard (same pattern as send_draft) → read + upload.
- < 3 MB uses a single POST fileAttachment (base64 contentBytes); 3–25 MB uses createUploadSession + 4 MB PUT chunks. Chunk PUTs go to the session's `uploadUrl` with a **plain fetch, not callGraphServer**: the uploadUrl embeds its own auth token and lives on outlook.office.com, where the graph.microsoft.com bearer token is the wrong audience.
- contentType comes from a ~30-entry extension map (text, images, office, pdf, archives, media) with fallback application/octet-stream; the extension is taken from the effective attachment name (attachment_name if given, else the file basename).
- The 3 MB single-POST threshold is checked against actual bytes read; the 25 MB cap against stat.size. Graph's own single-POST limit is ~3 MB and personal-account message limits are near 25 MB, so the caps are honest to the platform.

### Extension 4 — manage_rules
- The create path's first Graph call is a GET of the rules collection (used for the auto-assigned sequence = max existing sequence + 1), so a scope/permission problem fails loudly before anything is created; list is itself a GET.
- Forwarding/redirect rule actions are excluded per the task (exfiltration primitive). The list output still FLAGS forward/redirect actions on rules created elsewhere ("FORWARD/REDIRECT (created outside this server)") — surfacing them seemed safer than hiding them. Unknown conditions/actions on externally created rules are shown as raw key: JSON rather than dropped.
- A rule with zero conditions is rejected client-side with an explanation (it would match ALL mail) even though Graph would accept it — conservative-by-default per the task.
- move_to_folder is resolved via GET /me/mailFolders/{target} before creation (well-known name or id); 404 → isError; the rule stores the resolved folder id and the confirmation/list show the folder's display name.
- Rule delete is a hard DELETE of the rule object (Graph has no soft delete for messageRules); the rule stops acting but no mail is touched, so the soft-delete policy for content is unaffected.

### Harness (v3 tests)
- v3a asserts the first no-query result id equals a direct `$orderby=receivedDateTime desc&$top=1` call's id, and (when ≥2 results) compares the first two receivedDateTimes fetched by id.
- v3b snapshots graphRequestLog around the mark_read call and asserts the requests issued were exactly one /$batch and nothing else (a clean run has no throttling, so the strict ==1 is safe), plus 3/3 OK lines.
- v3d uses ~4.5 MB (not exactly 4 MB) so the session path uploads TWO chunks (4 MB + 0.5 MB); verification re-downloads the attachment's contentBytes and compares length AND sha256 against the local file — stronger than trusting Graph's `size` property, which can include metadata overhead.
- v3e's oversize file is created with truncate() to 26 MB — a sparse file whose stat.size is truthfully over the cap without writing 26 MB — and the guard rejects it before any read. The non-draft/missing-file guards run against the latest inbox message id.
- v3f cleans up in a finally block (rule delete by id if the tool-path delete didn't run, then permanentDelete of the test folder); the final sweep gained purgeTestRules() plus assertions that zero [MCP TEST] rules remain and that the run's temp dir is gone.
- Temp files live in one fs.mkdtemp dir under os.tmpdir(), removed per-test and swept at the end.

### Docs/versioning
- Version 0.3.0; README tool table now 17 tools, with a dedicated "Inbox rules" section (future-mail warning + no-forwarding rationale) and manage_rules added to the keep-per-call-approval list in the security model.

## v4 batch 2

Recorded while executing the "v4 batch 2" task on 2026-08-18 (subagent A). The Tasks.ReadWrite scope
and its re-consent had already landed (commit 85b4a3d); no auth, Entra, or Claude Desktop config was
touched in this batch.

### Live API shapes verified before writing code
- `/me/todo/lists` returns exactly one list on this account: "Tasks", `wellknownListName: "defaultList"`.
- **Graph normalises To Do dates to UTC on write** (a due date of `2026-08-20T00:00:00` America/Toronto
  comes back as `2026-08-20T04:00:00Z`), which would break naive date grouping. Reads therefore send
  `Prefer: outlook.timezone="America/Toronto"` — verified to work on `/me/todo`, returning local
  wall-clock — so grouping is a plain string comparison against `torontoToday()`, matching the calendar
  tools' existing approach rather than inventing a second date convention.
- `$filter=status ne 'completed'` and `$filter` on `/me/todo/lists` displayName both work, so the
  open-tasks default is server-side and list lookup by name needs no client scan of every list.
- The mailbox's six existing categories are all Graph presets, confirming the `preset0`–`preset24`
  palette is the real (and only) colour vocabulary.

### create_folder
- **Duplicate detection is client-side and case-insensitive.** Graph would answer `ErrorFolderExists`,
  but that error does not carry the existing folder's id, and the task wanted the id named. The handler
  lists the siblings first and reports the clash with its id; the Graph error is still caught as a
  fallback (lost race, or a hidden folder the sibling scan cannot see) with a message pointing at
  `list_folders`. Case-insensitive because Outlook itself treats sibling names that way.
- Uniqueness is per level: the same name is allowed again one level deeper, and the harness asserts it.
- The parent is resolved with its own GET before creating, so a bad `parent_folder` reports as such
  instead of surfacing as a confusing failure on the create call.
- **Deliberately no delete_folder tool.** It was not in the brief, and folder deletion in Graph takes
  the folder's whole message subtree with it — that belongs in its own batch with its own guards. The
  harness deletes its test folders through a test-only `permanentDelete`, as the earlier batches do.

### Shared `isNotFound` helper (bug found by live testing)
- Graph answers an unknown-but-well-formed id with 404 but an **unparseable** id with
  `400 ErrorInvalidIdMalformed`. The existing 404-only checks therefore leaked a raw
  "Microsoft Graph error (HTTP 400 …) ErrorInvalidIdMalformed" to the model for the common case of a
  made-up folder id. `isNotFound(err)` in common.ts now covers both, and `create_folder`,
  `manage_rules` (move target and rule id), and `manage_task` all use it.
- `ToolInputError` was added to common.ts for caller-fixable problems raised from shared helpers
  (`resolveTaskList`); `runTool` surfaces its message verbatim rather than under the "Tool failed:"
  prefix reserved for unexpected faults.

### manage_rules — update and exceptions
- `update` is a real PATCH: the harness asserts the rule keeps both its **id** and its **sequence**
  (evaluation position), which delete-and-recreate could not do.
- `conditions`, `exceptions`, and `actions` are each replaced wholesale when passed and left untouched
  when omitted — a partial update cannot silently drop a rule's actions (asserted in the harness).
- Guards carried over from create apply to update: passing `conditions: {}` is rejected (a rule with no
  conditions matches ALL mail), and passing `actions: {}` is rejected. `exceptions: {}` **is** allowed
  and clears them — an empty exception set is safe, an empty condition set is not.
- `enabled` is update-only in spirit but also honoured on create (`isEnabled: enabled ?? true`), so a
  rule can be created parked.
- Graph **uppercases `senderContains` values** on storage (`boss` → `BOSS`) while preserving
  `subjectContains`; harness assertions on rule summaries are case-insensitive because of this.
- Still no forwarding: `forwardTo` / `forwardAsAttachmentTo` / `redirectTo` appear only in the read-side
  list summary that flags foreign rules, never in a write path.

### Categories
- `manage_categories delete` takes `category_id` (from `list`), not a name, matching how `manage_rules`
  and `manage_contact` take ids. It pre-checks existence so deleting a stale id gives a readable error.
- Deleting a category **does not** strip it from messages already carrying it (Graph behaviour) — those
  keep the name without a colour. Stated in the tool description and the README rather than papered over
  by hunting down and rewriting every affected message.
- `manage_message categorize` **validates names against the master list before writing**: Graph happily
  accepts unknown category names on a message, which would leave colourless orphan labels behind. Names
  are matched case-insensitively and written back with the master list's exact spelling (harness asserts
  this round-trip).
- Replace-not-append is stated in both the action description and the tool description, and the empty
  array is the documented way to uncategorize — so there is no separate "uncategorize" action to keep in
  sync.

### Microsoft To Do
- **`manage_task delete` is the only irreversible operation in the whole tool surface.** To Do has no
  recoverable deleted-items store. The tool description says so explicitly, points at `complete` for the
  non-destructive case, and the handler reads the task's title before deleting so the confirmation names
  what was destroyed. README's soft-delete policy and security-model sections both carry the caveat.
- **`$select` is not supported on a single To Do task** — `GET …/tasks/{id}?$select=title` returns
  `400 invalidRequest` (found by live testing; the delete path originally used it). Single-task reads are
  now unprojected.
- `resolveTaskList` does one GET of all lists and matches id first, then display name case-insensitively.
  For a personal account with a handful of lists this is cheaper than a speculative GET-by-id, and an
  unknown name can report the available lists instead of a bare 404.
- Task ids are scoped to their list, so `task_list` must name the task's own list for
  complete/reopen/update/delete; the description says this, and a mismatch reports "No task … pass
  task_list if it lives in another list" rather than a raw 404.
- `due_within_days` **keeps overdue tasks** (they are more urgent than the window, not less) and **drops
  tasks with no due date** (they have no date to be within the window). Both are stated in the input
  description. `due_within_days: 0` means "due today or earlier".
- Neither `due_date` nor `reminder` can be *cleared* through `update` — only set. Clearing needs a
  distinct sentinel (empty string or null) that muddies the schema for a rare case; deferred rather than
  guessed at.
- `linked_message_id` copies a **reference** (subject, sender, received time, `webLink`) into the task
  notes, never the message body, and never modifies the message. A user-supplied `body` is kept first,
  with the linked block appended after a blank line.
- Setting `reminder` also sets `isReminderOn: true`; a reminder that is stored but switched off would be
  a silent no-op.

### Prompts
- Both prompts are **zero-argument**. Every MCP client renders argument-less prompts identically, and
  anything a prompt might have parameterised (a time window, a folder) is something the user can say in
  the very next turn. `registerPrompt` is used with `title` + `description` so pickers show useful text.
- `triage_inbox` is explicitly propose-only: it lists the tools it may call, names the writing tools it
  must **not** call during triage, and requires the user to name which proposals to apply. It reads the
  existing rules **before** proposing any, so proposals extend the user's rule set instead of duplicating
  or shadowing it (the reason `manage_rules list` comes second in its sequence, not last).
- `morning_brief` is honest about a real gap: `search_mail`'s listing carries no read/unread flag, so the
  prompt tells the model to call `read_message` when read state actually matters instead of asserting
  something it cannot see. Adding an unread filter to `search_mail` was out of this batch's scope.

### Harness (v4 tests)
- 23 tests total (19 before this batch). New: v4a create_folder, v4b manage_rules update, v4c categories
  + categorize, v4d task lifecycle; the stdio test now also exercises `prompts/list` **and**
  `prompts/get` for both prompts, asserting each renders a >200-char message naming the tools it drives —
  registration alone would not catch an empty or truncated prompt body.
- v4d builds its task from `latestMessage` (a genuinely received message) rather than a draft, because a
  draft has no sender and no useful `webLink` — the linked-note assertions would have been vacuous.
- The final sweep gained category and task purges plus assertions, and now also walks **one level of
  subfolders** under every top-level folder, since `create_folder` can nest and the old root-only scan
  would have missed a stray child.
- Test artifacts stay under the existing `[MCP TEST]` prefix so one sweep covers every batch. Full run:
  23/23 PASS.

### Docs/versioning
- Version 0.4.0. README: 21-tool table, a "Microsoft To Do notes" subsection, a "Prompts" section, an
  update/exceptions paragraph under "Inbox rules", `Tasks.ReadWrite` added to the setup scope list, the
  soft-delete policy narrowed to *mailbox* deletes with the To Do exception called out, and `manage_task`
  added to both the keep-per-call-approval list and a dedicated security-model bullet.

## v5 batch 3 — dual-mode (stdio + Cloudflare Worker)

### Transport-agnostic refactor
- The tools were already transport-independent except for one thing: `graph.ts` imported `auth.ts`,
  which pulls in MSAL, `dotenv` and `node:fs`. That single edge would have dragged all of Node into
  the Worker bundle. Broke it with `src/core/token.ts`: a **token provider** indirection that hosts
  install and `core/graph.ts` consumes. Nothing in `src/tools/` changed beyond import paths.
- The provider is stored in an `AsyncLocalStorage` with a process-wide default. The default serves
  single-tenant hosts (stdio, CLI scripts, harness); the ALS scope serves the Worker, where the KV
  binding only exists per request and concurrent requests must not observe each other's provider.
  `node:async_hooks` is available on workerd under `nodejs_compat`.
- `callGraph` (the device-code-capable variant) moved to `src/graph-interactive.ts` so `core/graph.ts`
  has no Node-only imports at all.
- `src/core/registry.ts` is the single tool/prompt table. Both entry points call `createMcpServer()`,
  so the surfaces cannot drift; remote test `r9` asserts the deployed tool list equals `TOOLS`.
- Version is duplicated: `package.json` (read at runtime by the stdio entry) and `core/version.ts`
  (a literal, because the Worker has no filesystem). Kept in sync by hand at release time.

### SDK choice — why not `agents` / `McpAgent`
- Cloudflare's `McpAgent` (from the `agents` package) is **deprecated and feature-frozen** as of the
  July 2026 SDK v2 release, and it requires Durable Objects (one DO instance per session). Free-tier
  accounts *can* use SQLite-backed DOs, so that was not the blocker — the blockers were that it is a
  deprecated path and that its successor, `createMcpHandler`, wants `@modelcontextprotocol/server@2`,
  a different SDK from the `@modelcontextprotocol/sdk@1.30.0` the 21 tools are written against.
- Chose instead **`WebStandardStreamableHTTPServerTransport`**, which ships in the SDK we already
  depend on (`@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js`) and is fetch-native —
  its own JSDoc carries a Cloudflare Workers example. This keeps *one* MCP SDK across both transports
  and means the Worker registers the identical `McpServer` object the stdio server does.
- Ran **stateless**: `sessionIdGenerator: undefined`, `enableJsonResponse: true`, a fresh server and
  transport per POST. Verified in the SDK source that `validateSession()` returns early when
  `sessionIdGenerator` is undefined, so the "Server not initialized" gate never fires for a
  non-`initialize` request on a fresh transport — which is what makes per-request construction safe.
  Consequence: no Durable Objects, no migrations, no DO bindings in `wrangler.jsonc`.

### OAuth design
- `@cloudflare/workers-oauth-provider@0.10.3` owns the authorization-server surface: RFC 8414 and RFC
  9728 metadata, RFC 7591 dynamic client registration, S256 PKCE, the token endpoint, bearer
  validation, and the `401` + `WWW-Authenticate: ... resource_metadata=...` challenge claude.ai needs
  in order to start the flow. `apiRoute: "/mcp"` means nothing anonymous can reach the MCP handler.
- DCR is enabled because claude.ai registers itself; CIMD was left off (it needs the
  `global_fetch_strictly_public` compat flag and DCR alone is sufficient here).
- `resourceMetadata.resource` is the literal deployed URL **including `/mcp`**. RFC 9728 requires it
  to match what the client was given exactly, and the provider is constructed at module scope where
  `env` is not available, so it cannot be read from a binding. Renaming the Worker means editing it.
- **Identity is proved with Microsoft's device-code flow, not an auth-code redirect.** The Entra app
  registration is a public native client registered for device code with no web redirect URI; adding
  one would have meant changing the registration (portal access this task did not have). Device code
  needs no redirect URI, so `/authorize` renders a code, polls Microsoft, and completes. The token it
  obtains is scoped `offline_access User.Read`, used only to call Graph `/me`, and never stored.
- **Allowlist**: `isAllowedIdentity()` matches the Graph `/me` `id` against `ALLOWED_MS_USER_ID`, or
  `userPrincipalName`/`mail` against `ALLOWED_MS_UPN`. Consumer (MSA) accounts return no
  `userPrincipalName` from Graph and their `/me` `id` (`c85e8009ffaa7530`) is *not* the `oid` in the
  token, so both forms are checked. Every authorization path funnels through this one function, and
  `mcp-handler.ts` re-checks `props.userId` against the allowlist as defence in depth (it catches a
  grant minted before the allowlist was narrowed).
- **Deviation worth flagging:** `/authorize` also accepts `POST` with an `ms_access_token` form field,
  which runs the identical Graph `/me` allowlist check and completes the authorization without a
  browser. This exists so the remote harness can drive a *real* OAuth exchange (DCR → authorize →
  PKCE token exchange) non-interactively; a device-code sign-in cannot be automated. It is not a
  bypass: it demands a live Microsoft access token for the allowlisted account, and anyone holding
  one already has the mailbox. It was preferred over a static test bearer secret precisely because it
  reuses the real gate instead of adding a second credential path. Note that MSA access tokens are
  opaque (not JWTs), so the token cannot be additionally checked for `appid` — validation is
  necessarily "does Graph accept it, and for whom".

### Token handling
- MSAL Node does not run on workerd, so `src/worker/ms-token.ts` issues the refresh-token grant
  directly: form POST to `https://login.microsoftonline.com/consumers/oauth2/v2.0/token` with
  `client_id`, `grant_type=refresh_token`, `refresh_token`, `scope`. No client secret (public
  client), no `redirect_uri`, no `Origin` header.
- `scope` repeats **exactly** the consented set with `offline_access` first. `offline_access` must be
  present or the response carries no rotated refresh token; requesting the same scopes guarantees no
  new consent prompt, which was a hard requirement.
- **Rotation**: the new `refresh_token` is written back to `ms:refresh_token` on every exchange,
  before the access token is cached, so a later failure cannot lose it. Microsoft does *not* revoke
  the previous refresh token when it issues a new one, which is why (a) a concurrent double-refresh
  is not fatal, and (b) the Worker rotating its copy leaves `.token-cache.json` working — the two
  credential chains are independent from the moment of seeding.
- Access tokens are cached in KV under `ms:access_token` with a 5-minute expiry skew and a matching
  `expirationTtl`, so a typical tool call costs no token exchange.
- Seeding (`src/scripts/seed-kv.ts`) reads the single `RefreshToken` entry from `.token-cache.json`,
  refuses to guess if the cache holds more than one account, and passes the value to wrangler through
  a `0600` temp file rather than argv (argv is world-readable via `ps`). It prints only a SHA-256
  fingerprint, and reports the two allowlist values.

### Secrets
- `MS_CLIENT_ID`, `ALLOWED_MS_USER_ID`, `ALLOWED_MS_UPN` are wrangler secrets. Only the first is
  arguably sensitive (it is a public client identifier), but all three stay out of the repo.
- `wrangler.jsonc` is committed and contains only the two KV namespace ids, which are not secrets.
  `.gitignore` gained `.dev.vars` and `.wrangler/`.
- `worker-configuration.d.ts` (generated by `npm run cf-types`) **is** committed so `npm run
  typecheck` works from a clean checkout. It contains binding *names* only, no values.

### Workers limitations hit
- Two tsconfigs were unavoidable: `@types/node` and the workerd runtime types both declare `Request`,
  `Response` and `fetch` globally. `tsconfig.json` excludes `src/worker` and covers the Node side;
  `tsconfig.worker.json` covers `src/worker` + `src/core` + `src/tools` against the generated runtime
  types. `npm run typecheck` runs both.
- KV is eventually consistent and caches reads at the edge for up to a minute. This showed up twice:
  the rotation assertion polls for the new value, and the post-revocation `401` assertion polls for
  up to two minutes. Both are genuine platform behaviour, not flakiness to paper over.
- A `/.well-known/oauth-authorization-server` fetch immediately after the first deploy returned
  Cloudflare error 1042; it resolved on its own within a minute and has not recurred. Treated as
  deploy propagation, not a code fault.

### Tests
- Local harness unchanged and still 23/23 — it now calls `installMsalTokenProvider()` explicitly,
  since the token source is no longer implicit in the import graph.
- `src/test-remote.ts` (`npm run test:remote`) adds 14 live checks against the deployed endpoint:
  discovery metadata (r1, r2), anonymous/bogus-bearer rejection across POST, GET and DELETE (r3),
  DCR (r4), allowlist refusal (r5), a full authorize + PKCE token exchange (r6), bad-verifier
  rejection (r7), `initialize` (r8), the 21-tool surface compared against the registry (r9),
  `prompts/list` (r10), a read-only `list_events` round-trip (r11), refresh-token rotation proven by
  forcing an exchange and comparing stored fingerprints (r12), and teardown (r13, r14).
- Self-cleaning extends to the cloud: the run snapshots `OAUTH_KV`'s key list before registering
  anything and deletes every key that appeared, then asserts the namespace is back to its baseline
  and that the issued bearer no longer opens `/mcp`. Comparing against a baseline rather than
  matching on the client id was deliberate — grants and tokens are keyed by user and grant id, so an
  id-based filter silently missed them (which is exactly how r14 caught the incomplete first version).
- The rotation test deliberately mutates real state (it spends the refresh token). That is safe and
  needs no cleanup: rotation is the normal steady-state behaviour, and the replacement token is live.

### Docs/versioning
- Version 0.5.0. README gained a "Remote deployment" section (architecture, allowlist, token storage,
  from-scratch setup, the exact claude.ai connector steps, and rotation/revocation procedures), five
  new script entries, and a security-model bullet for the remote endpoint.
- Deployed URL: `https://outlook-mcp.arthur-yuhao-zhang.workers.dev/mcp`.
