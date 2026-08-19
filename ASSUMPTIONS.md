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
- **Deviation worth flagging (since gated — see "v6.1: production hardening" below):** `/authorize`
  also accepts `POST` with an `ms_access_token` form field, which runs the identical Graph `/me`
  allowlist check and completes the authorization without a browser. This exists so the remote
  harness can drive a *real* OAuth exchange (DCR → authorize → PKCE token exchange)
  non-interactively; a device-code sign-in cannot be automated. It is not a bypass: it demands a
  live Microsoft access token for the allowlisted account, and anyone holding one already has the
  mailbox. It was preferred over a static test bearer secret precisely because it reuses the real
  gate instead of adding a second credential path. Note that MSA access tokens are opaque (not
  JWTs), so the token cannot be additionally checked for `appid` — validation is necessarily "does
  Graph accept it, and for whom". As of v6.1 this path only runs when `ALLOW_DIRECT_AUTHORIZE=true`,
  which production never sets.

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

## v6 batch 4 — delta queries, change notifications, resources

### Live Graph behaviour verified before writing code
- `$deltatoken=latest` (the cheap "baseline without enumerating" trick) is **OneDrive-only**. Graph
  ignored it on `/me/mailFolders/inbox/messages/delta` and returned an ordinary first page with no
  `@odata.deltaLink`. So a first call really does have to walk the folder.
- `Prefer: odata.maxpagesize=N` **is** honoured on the initial delta request, but the resulting
  `@odata.nextLink` does not carry the preference — follow it without the header and pages revert to
  10 items. Sending the header on every request took baselining the real 1006-message inbox from 92
  requests / 10.8 s to 3 requests / 0.85 s. This is why `drainDelta` passes `headers` on each hop
  rather than only the first.
- A delta entry for an *edited* message carries only `@odata.type`, `id` and the changed properties
  (e.g. just `isRead`), not the `$select`ed set. `check_new_mail` therefore re-reads entries that come
  back without a subject, capped at the result limit, and marks them `[changed]`.
- Change notifications on `/me/mailFolders('inbox')/messages` **are** supported for this personal
  (consumer) Microsoft account — confirmed by creating a real subscription against the deployed
  endpoint and deleting it again before wiring anything. Graph granted `expirationDateTime` ~2.9 days
  out, matching the documented 4230-minute ceiling.
- Subscriptions are visible to `GET /subscriptions` from the local MSAL session because the Worker and
  the CLI use the same Entra app registration and the same user — which is what lets remote test `r16`
  cross-check Graph's view against the KV record.

### State storage (`src/core/state.ts`)
- Delta positions and the notification ring buffer need somewhere to live, and the tool layer must not
  know where. Mirrored the existing token-provider pattern exactly: a `StateStore` interface, an
  `AsyncLocalStorage` scope for the Worker (KV binding is per-request), a process-wide default for
  stdio and the harness.
- The store carries a `mode` (`"local" | "remote"`). That is what `get_mailbox_activity` checks to
  refuse on stdio — better than sniffing for a KV binding, and it lets the harness exercise the remote
  path locally with an in-memory store.
- Local file: `.mcp-state.json`, `0600`, gitignored, next to `.token-cache.json`. Writes go through a
  single promise chain because the whole document is rewritten on each put and MCP tool calls overlap.
- KV keys are namespaced (`delta:`, `sub:`, `activity:`) so they can share `OUTLOOK_KV` with the token
  entries without any chance of collision.

### check_new_mail
- The first call (and `reset`) deliberately reports **nothing**. Dumping a thousand messages on a tool
  whose whole point is "only what's new" would be the wrong answer, and the position is what the caller
  actually wanted.
- The position advances on every successful call, so a change is reported exactly once. Test v5a
  asserts precisely that (the probe message appears once, then never again).
- `MAX_PAGES = 200` bounds a first enumeration; hitting it fails with a readable message rather than
  looping. At 500 items per page that is 100 000 messages.

### Change notifications
- The validation handshake is answered **before** any state is read. It has to be: at the moment Graph
  validates a brand-new subscription there is no subscription record to consult, so a handler that
  looked one up first could never create the first subscription. Test v5b asserts the handler writes
  and needs no state.
- `clientState` is generated in the Worker (two `crypto.randomUUID()`s, hyphens stripped) at creation
  time and stored only in KV. It is deliberately **not** a wrangler secret: it must rotate with the
  subscription, and a secret would have to be re-put by hand every time the subscription is rebuilt.
  It never appears in the repo, in `wrangler.jsonc`, or in any log line.
- Deliveries always get `202`, valid or not. A `4xx` for a bad `clientState` would turn the endpoint
  into an oracle for guessing it, and a non-2xx makes Graph retry the delivery for hours.
- Ring buffer: 50 entries, newest first, read-modify-write. For one mailbox at human mail volume the
  lost-update window is not worth a Durable Object; the alternative would have reintroduced exactly
  the DO dependency batch 3 avoided.
- Enrichment (subject/sender lookup) is capped at 5 messages per delivery and is best-effort — Graph
  wants a response within 10 seconds, and a failed lookup must cost detail, not the notification.

### Subscription upkeep and the cron
- `renewalDecision` is a pure function of (record, notification URL, now) and is unit-tested for every
  branch: absent, lapsed, moved endpoint, near expiry, healthy. `ensureMailSubscription` wraps it and
  takes an injectable Graph transport and clock, so v5d exercises the **real** renewal handler
  (create → keep → renew → recreate-after-404) with no live side effects.
- Cron `17 */6 * * *` with a 24-hour renewal window: four chances a day against a ~70-hour lifetime, so
  a single failed run is harmless. Asks for 4200 minutes rather than the 4230 maximum so clock skew can
  never make Graph reject the request.
- Added a backstop the spec did not ask for: every authenticated MCP request calls the same upkeep
  function in `ctx.waitUntil`. Rationale — a lapsed subscription is invisible until someone asks for
  activity, and the cron cannot be triggered on demand in production, so this is also what bootstraps
  the very first subscription after a deploy. When nothing is due it is one KV read and zero Graph
  calls, which is why it is cheap enough to run on every request.
- `OAuthProvider` exports only `fetch`, so `src/worker/index.ts` now wraps it in an object that also
  exports `scheduled`.

### Resources
- The SDK supports resources identically on both transports (`registerResource` on the shared
  `McpServer`), so no limitation had to be recorded — both stdio and the Worker serve
  `outlook://mail/folders` and `outlook://mail/inbox/recent`. Asserted on stdio (test h) and over
  Streamable HTTP (r14).
- The read callbacks reuse the `list_folders` and `search_mail` handlers rather than duplicating
  formatting, and **throw** on failure: `resources/read` has no `isError` channel, and attaching an
  error string as if it were mailbox content would be worse than failing.

### Harness (v6 tests)
- Local: 28/28 (was 23/23). New: v5a delta lifecycle against real mail, v5b handshake, v5c ingest +
  clientState + ring cap + the activity tool's window filter, v5d the renewal handler, v5e the
  remote-only refusal. The harness installs the real file-backed store but pointed at a throwaway file
  in `os.tmpdir()`, so a run can never disturb the server's own delta position; the sweep asserts the
  position was written there and then deletes it.
- Remote: 20/20 (was 14/14). New: r13 live handshake + forged-delivery rejection (unauthenticated),
  r14 resources, r15 the delta position surviving between requests (which on a stateless Worker can
  only mean KV), r16 the subscription's health cross-checked against Graph plus the cron declaration,
  r17 the full end-to-end (send through the connector → Graph → `/notifications` → ring buffer →
  `get_mailbox_activity`) on a bounded 4-minute poll, r18 cleanup.
- r17's failure path was written before its success path: on timeout it re-checks whether the message
  reached the inbox at all and says whether the gap is in sending or in Graph's delivery, and prints
  the last activity output.
- Cleanup boundary, per the spec: the **production** subscription survives a test run (it is the
  feature, not an artifact). Everything else the run caused does not — the probe mail is
  permanently deleted, only the ring-buffer entries whose subject carries `[MCP TEST]` are removed
  (entries that predate the run are preserved, and the count is asserted), and the delta position is
  restored to exactly what it was, or deleted if there was none.

### Docs/versioning
- Version 0.6.0. README: two new tool rows, a "Resources" section, a "Knowing what is new" section
  contrasting the two mechanisms, a "Change notifications" section under Remote deployment (handshake,
  `clientState`, renewal, `PUBLIC_BASE_URL`), and a security-model bullet for the one public route.
- `PUBLIC_BASE_URL` is a `vars` entry rather than a secret: it is a hostname, it must match the
  deployment, and having it in the committed config is what makes the notification URL reviewable.

## v6.1: production hardening — the direct authorize path is now disabled in production

### The change
- `POST /authorize` with an `ms_access_token` form field (the non-interactive path the remote
  harness used) is now gated on a new binding, `ALLOW_DIRECT_AUTHORIZE`. Unless it is exactly the
  string `"true"`, the Worker answers `403` with a "disabled" detail **before parsing anything** —
  no auth-request parsing, no form read, no Graph call. The flag is deliberately absent from the
  deployed Worker (neither a `vars` entry in `wrangler.jsonc` nor a secret), so production only
  authorizes through the interactive device-code flow. Local `wrangler dev` runs enable it via the
  gitignored `.dev.vars`.
- Why, given the path was already allowlist-gated: (1) it let anonymous callers make the Worker
  forward arbitrary attacker-supplied tokens to Graph, making the endpoint a token-validation
  oracle; (2) it would convert a leaked *short-lived* Microsoft access token into a persistent MCP
  grant with no interactive sign-in — the device-code flow requires a live sign-in the token holder
  may not be able to perform; (3) production should not carry test-only affordances at all.
- Presence-plus-value (`=== "true"`) rather than mere presence was chosen so an accidentally set but
  falsy value (`"false"`, `"0"`, empty) still means disabled.

### Consequences for the remote suite
- r5 changed meaning: it used to prove the allowlist refused a bogus token on this path; it now
  asserts the deployed endpoint refuses the path outright (`403` + a detail matching /disabled/i,
  which also proves the gate fires before token validation — a bad token used to earn a `401`).
- r6 could no longer POST a locally held Microsoft token to production, so the bearer for the
  authenticated tests (r6, r8–r12, r14, r15, r17, r20) now comes from driving the *real* device-code
  flow: r6 fetches the production `/authorize` page, parses the user code, verification URL and flow
  id out of the HTML, asks the human running the suite to enter the code at
  microsoft.com/devicelogin, and polls `/authorize/poll` (10-minute bound) exactly as the page's own
  script does. This is strictly more end-to-end than before — the browser flow itself is now
  exercised.
- A device-code sign-in cannot be automated, so in a headless run (no TTY, or
  `MCP_REMOTE_HEADLESS=1`; force interactive with `MCP_REMOTE_INTERACTIVE=1`) the bearer-dependent
  tests are reported as `SKIP` (counted as passed, listed in the summary) and everything
  unauthenticated still runs: r1–r5, r7 (PKCE rejection needs no bearer), r13 (webhook handshake),
  r16 (subscription health via local Graph + KV), r18/r19 (cleanup, both idempotent when nothing
  authenticated ran).
- `.dev.vars` (gitignored, never deployed) now holds `ALLOW_DIRECT_AUTHORIZE=true` plus mirrors of
  the three wrangler secrets, so a local `wrangler dev` Worker can run the full direct-path exchange.
  Verified both ways against `wrangler dev --local` before deploying: with the flag, a registered
  client plus a real MSAL token completed authorization (`status: "ok"` with a code); without
  `.dev.vars`, the same POST got the `403` "disabled" answer.

### Subscription upkeep race (found live, then fixed)
- Re-verifying the webhook after the v6.1 deploy found **four** live Graph subscriptions for the
  same endpoint, all created within one second during the v6 batch-4 remote run: concurrent
  authenticated requests each ran the `ctx.waitUntil` upkeep backstop, each read a stale KV view
  (KV is eventually consistent, reads cached up to a minute), each decided "create", and the last
  KV write won. Harmless in effect — deliveries from the three orphans failed the clientState
  check and were discarded — but an unbounded unlocked check-then-act race.
- Fix: `ensureMailSubscription` now treats **Graph as the source of truth and KV as a cache**
  whenever the KV record alone cannot justify "keep". It lists the unexpired subscriptions for this
  notification URL + resource, renews the one whose id the KV record names (sweeping all others),
  and only creates when the record's subscription is not live in Graph — after deleting whatever
  else is there. The per-request fast path is unchanged: a healthy record still costs one KV read
  and zero Graph calls.
- **Adoption is impossible, by Graph's design:** list/get responses return `clientState: null`
  (verified live), so a subscription created by a racing peer can never be reused — without its
  secret, every delivery it makes would be discarded. That is why the convergence step *replaces*
  a foreign live subscription rather than adopting it, and why the requested "reuse the existing
  one" is implemented as "reuse iff the KV record holds its clientState".
- Residual window: two upkeeps that both list *before* either creates (true bootstrap only) still
  create one each; the next renew/create-path upkeep sweeps the loser. Asserted in v5f alongside
  the stale-KV case, which now converges to exactly one subscription immediately.
- The three orphans from the incident were deleted live via Graph (`DELETE /subscriptions/{id}`),
  keeping `95f07422…` — the one the `sub:mail` KV record holds the clientState for.

## v7 batch A — attachments on both transports, recurring events, reminders, calendars

### Live Graph behaviour verified before writing code
Everything doubtful for a **personal (consumer)** account was probed against the real mailbox first,
with a throwaway script, before any of it was designed in:
- **Recurrence is fully supported.** A weekly `numbered` series (3 occurrences) was created and
  `calendarView` expanded it into exactly three dates.
- **Per-occurrence edits are supported.** `PATCH /me/events/{occurrenceId}` moved one date and Graph
  turned that occurrence into `type: "exception"`, leaving the other two untouched;
  `DELETE /me/events/{occurrenceId}` removed a single date without disturbing the series.
- **Series-wide edits are supported.** `PATCH` on the `seriesMaster` id renamed every date and set a
  reminder.
- **Reminders are supported** (`isReminderOn` + `reminderMinutesBeforeStart`), on the default and on
  secondary calendars.
- **Secondary calendars are supported.** The account already had `Calendar` (default), `United States
  holidays` and `Birthdays`; a new calendar was created, written to, read back through its own
  `calendarView`, and deleted. Crucially `GET /me/events/{id}` resolves an event living in a
  non-default calendar, which is why `manage_event` needs no `calendar` input.
No consumer-account gap was found, so nothing had to ship as "graceful absence" this time.

### add_attachment — three sources
- Exactly one of `file_path` / `url` / `content_base64`; zero or several is a caller error naming what
  it got. Every source shares the same 25 MB cap, draft check and upload strategy (single POST under
  3 MB, chunked upload session above).
- **Order of checks is deliberate and unchanged from v3:** cheap per-source pre-flight (path absolute
  + `stat` + size, or URL parse + scheme, or base64 decode + size) runs *before* the draft lookup, so
  a missing or oversized file is still reported as such even when the target is not a draft — v3e
  asserts exactly that. The expensive part (reading the file, downloading the URL) runs only after the
  draft check passes, so a 25 MB download is never spent on a message that cannot take attachments.
- `url` is **https only**. Plaintext is refused, and so is `file:` — which is the point: without a
  scheme check, `url` would re-introduce the local-filesystem read that `file_path` is explicitly
  denied on the hosted server, and would do it server-side with no path validation at all.
- The download is read chunk by chunk and abandoned the moment it exceeds 25 MB, rather than
  `arrayBuffer()`-ing whatever arrives. `Content-Length` is checked first when present, but it is a
  claim by a remote server, not a guarantee.
- Content type: the response's own `Content-Type` wins when it says anything specific; a generic
  `application/octet-stream` tells us nothing the filename extension does not, so the existing
  extension table handles it. The URL's last path segment names the file unless `attachment_name`
  overrides it.
- `content_base64` is capped at **3 MB decoded** — the same threshold as the single-POST upload path,
  and about as much base64 as is sane to put through a model's context.
- `file_path` on the hosted server is refused via `getStateStore()?.mode`, the same "what kind of host
  is this" check `get_mailbox_activity` uses. It names the two alternatives instead of failing with a
  filesystem error.

### get_attachment — links instead of a filesystem
- Text-like content under 50 KB is returned inline on **both** transports (unchanged locally). Only
  the "there are bytes and nowhere to put them" case differs.
- The hosted server stores the base64 Graph already returned under an unguessable id
  (`crypto.getRandomValues`, 256 bits, hex) in `dl:<id>` and returns a link. **Two expiries, on
  purpose:** the record carries its own `expiresAt`, enforced on every read (exact, and the only one
  that can be trusted), and the KV entry gets an `expirationTtl` purely as garbage collection — KV
  expiry is eventual and its minimum is 60 s, so it cannot be the deadline. An expired record is
  refused *and deleted*.
- Cap: 18 MB. A KV value may not exceed 25 MB and the bytes are stored base64 (4/3 expansion). Over
  that, the tool says so and points at Outlook or the local server rather than failing obscurely.
- `link_ttl_minutes` (1–15, default 15) exists mostly so the remote suite can assert expiry inside a
  test run; 1 minute is the floor because that is KV's own minimum TTL.

### The download route lives under `/mcp/`, and why that is not cosmetic
- First attempt served it at `/download/<id>` and listed that as a second `apiRoute`. Anonymous
  requests were correctly refused — but so were **authenticated** ones:
  `{"error":"invalid_token","error_description":"Token audience does not match resource server"}`,
  while `/mcp` accepted the very same bearer. Found by driving a real `wrangler dev` Worker with a
  real bearer before the interactive gate, not by reading the code.
- Cause: `workers-oauth-provider` binds an access token's audience to the advertised resource
  (`…/mcp`, pinned by RFC 9728 exact-match for claude.ai) and matches it *path-prefixed on path
  boundaries* (`audienceMatches`). `/download/…` is outside `/mcp`, so no legitimate client's token
  could ever open a link.
- Fix: `DOWNLOAD_ROUTE_PREFIX = "/mcp/download/"`. It is inside the audience, and inside the existing
  `/mcp` `apiRoute` (prefix-matched), so the second route entry went away too — one route, one bearer
  check, links that the connector's own token opens. Re-verified end to end against `wrangler dev`:
  anonymous → 401 with `WWW-Authenticate`, bogus bearer → 401, valid bearer → 200 with the exact 2048
  probe bytes, `Content-Disposition: attachment; filename="blob.bin"`, unissued id → 404, malformed
  id → 404, backdated record → 404 and the record dropped from KV.
- The handler re-checks the grant's `userId` against the allowlist exactly as `mcp-handler.ts` does:
  an unguessable id is a defence against enumeration, not a substitute for authentication.

### StateStore gained two things
- `put(key, value, {ttlSeconds})` — honoured by all three implementations (KV `expirationTtl` with the
  60 s floor, in-memory expiry, and the on-disk store, which wraps a TTL'd entry as
  `{value, expiresAt}` and drops expired entries whenever it rewrites the file; plain strings are
  still read back as before).
- `publicBaseUrl` — how a tool learns the URL to hand back without knowing it runs on a Worker. It
  sits beside `mode` because it answers the same question: what kind of host is this. When it is
  absent the tool says it cannot issue a link rather than emitting a broken one.

### Recurrence mapping
- The caller's vocabulary (`frequency`, `interval`, `weekdays`, `day_of_month`, `month`, `until`,
  `count`) is mapped onto Graph's `pattern` + `range` pair in `create-event.ts`, next to
  `toGraphDateTime` — shared by both calendar tools so a series is described identically wherever it
  is built. Monthly and yearly map to Graph's *absolute* patterns; `relativeMonthly` ("third Tuesday")
  is deliberately not exposed yet — it needs a second axis (index + weekday) that no natural request
  in this mailbox has needed.
- Anything omitted is taken from the event's own start date (weekday for weekly, day/month for
  monthly and yearly), so the common case is `{frequency, count}`.
- `until` and `count` are mutually exclusive (Graph has one `range.type`), and `until` before the
  start date is refused rather than sent. Neither given = `noEnd`, stated in the schema description so
  an unbounded series is never a surprise.
- `range.recurrenceTimeZone` is `America/Toronto`, matching every other datetime in this server.

### manage_event — occurrence vs series
- Graph models a repeating event as a `seriesMaster` plus per-date `occurrence`s with their own ids,
  so the tool reads `type` and `seriesMasterId` *before* acting rather than trusting the caller's
  `scope`. Defaults follow the id: an occurrence id changes that date, a series id changes the series.
- `scope: entire_series` from an occurrence walks up to `seriesMasterId` and re-reads the master (its
  attendee list and organizer flag, not the occurrence's, are what the output reports).
- `this_event_only` against a **series** id is refused, with instructions to get the occurrence id
  from `list_events include_ids`. Picking a date on the caller's behalf would be a guess that mails
  attendees.
- A `recurrence` change on a single occurrence is refused for the same reason: a repeat rule belongs
  to the whole series.
- Both descriptions state the notification consequences, since that is the part a user would be
  surprised by: a series edit notifies every attendee about all occurrences (and replacing the rule
  re-issues the series), one occurrence notifies about that date only.
- `reminder_minutes: -1` means "off" on update. A separate boolean would have been a second way to
  say the same thing; the sentinel is documented in the schema description and in the README.

### list_events / list_calendars decisions
- **A separate `list_calendars` tool, not folded into `list_events`.** Recorded as asked: the two
  answer different questions ("which calendars exist" vs "what is on one"), folding it in would have
  meant a mode flag whose output has nothing in common with the normal one, and `list_folders` already
  sets the precedent that "list the containers, with ids" is its own tool. It also gives `calendar` a
  discoverable vocabulary — an unknown name fails with the real names, and the model can find them
  without guessing. Cost: one more tool (24).
- Like `list_folders`, it carries one harmless optional flag (`writable_only`) because MCP clients
  handle an empty input schema poorly.
- **`include_ids` on `list_events` defaults to false.** Event ids are ~150 characters; printing them
  for every event in a 7-day (or 31-day) window would dominate the output. But without them
  `manage_event` was unreachable from `list_events` at all — a pre-existing gap that only became
  load-bearing once single occurrences became addressable. Off by default keeps the daily listing
  readable; the description tells the model exactly when to turn it on. Occurrence lines say they are
  occurrences, so the model knows a `scope` decision exists.
- `calendar` accepts a name or an id, resolved with one GET and matched exactly, then
  case-insensitively, then by prefix — the same shape as `resolveTaskList`.

### Harness (v7 tests)
- Local: 29 → **34 tests**, all against the real mailbox and self-cleaning.
  - `v7a` attaches the deployed Worker's own `/health` endpoint over `url` — a small, public, always-on
    https resource this project owns, which also proves name-from-URL and type-from-response — plus the
    scheme, parse and unreachable-host guards.
  - `v7b` round-trips `content_base64` and asserts the no-source, two-source, malformed and 3 MB
    guards, and that `file_path` is refused under a remote-mode store.
  - `v7c` covers `get_attachment` on **both** transports from one draft: local save to
    `~/Downloads/outlook-mcp-attachments/` (the file is deleted afterwards, and the directory too if
    the test created it), remote inline text, and a remote binary parked behind a link whose id is
    unguessable, whose record round-trips the exact bytes, and which is refused *and dropped* once its
    `expiresAt` is backdated. Unknown and malformed ids are refused too.
  - `v7d` is the full recurring lifecycle against `calendarView`: weekly ×3 → one occurrence moved
    (exception; the other two hold their time) → series-wide rename + reminder *from an occurrence id*
    → `this_event_only` on the series id refused → series cancelled and every date gone.
  - `v7e` creates a real secondary calendar, checks `list_calendars`, creates an event there with a
    reminder, reads it back through `list_events calendar=…` (and asserts it is **absent** from the
    default calendar), turns the reminder off, and checks the unknown-calendar error names the real
    calendars.
  - The sweep now purges and then asserts on **calendars** and on `[MCP TEST]` events in *every*
    calendar — `/me/events` only sees the default one, so without this a leaked event in a secondary
    calendar would have gone unnoticed.
- Remote: 20 → **22 tests**. `r3` now also asserts the download route refuses anonymous and
  bogus-bearer callers (a well-formed but unissued id must be a `401`, not a `404` — that ordering is
  what keeps links private, and it runs headless). `r18` attaches an https URL from the Worker and
  checks Graph's own inventory with the local token, then confirms `file_path` is refused there.
  `r19` takes the whole download path end to end: a 1-minute link, anonymous → 401, bearer → 200 with
  the exact bytes and the right `Content-Disposition`, an unissued id → 404, then a bounded poll until
  the link dies with 404. The old r18–r20 (cleanups, revoked-bearer sweep) are now r20–r22.
- Both new remote tests are `testAuthed`, so a headless run still reports 22/22 with 12 skips.

### Docs/versioning
- `package.json` and `src/core/version.ts` → **0.7.0** (the harness asserts they match).
- README: v7 header, 24-tool table with the new inputs, and three new sections — attachments on both
  transports (including why the download route sits under `/mcp/`), repeating events (with the
  occurrence/series scope table), and calendars/reminders.

## v8 batch B — To Do depth, junk senders, mailbox settings, message forensics

### Live Graph behaviour verified before writing code
Every doubtful capability was probed against the real **consumer** mailbox with a throwaway script
before anything was designed on top of it. Two of the four features came back different from what the
brief assumed, and the code follows what Graph actually does.

**Microsoft To Do**
- Checklist items work exactly as documented: `POST/GET/PATCH/DELETE
  /me/todo/lists/{list}/tasks/{task}/checklistItems` all succeeded (201/200/200/204), and
  `?$expand=checklistItems` returns them with the task.
- Lists: `POST /me/todo/lists` → 201, `PATCH` (rename) → 200, `DELETE` → 204. All supported; the
  delete is the one this server refuses to expose (below).
- **Task recurrence is create-only.** `POST` with `recurrence` + `dueDateTime` → 201 and the rule
  reads back. Without a due date: `400 "The property 'dueDateTime' is required when creating
  recurrence in the task entity"`. And **every** `PATCH` carrying a recurrence is refused:
  `400 invalidRequest — Invalid JSON, Error converting value "2026-08-20" to type
  'Microsoft.OData.Edm.Date'. Path 'recurrence.range.startDate'` — with or without
  `recurrenceTimeZone`, with `startDate` as a date or a full datetime, with `@odata.type`
  annotations, with `dueDateTime` in the same body, and on **beta** as well as v1.0. (`"2026-08-20"`
  is a perfectly valid `Edm.Date`; the error is the API's, not the caller's.) The one PATCH Graph
  accepts is `recurrence: null` → 200, which really does stop the repeat.
  Consequence: `recurrence` is accepted only on `create`; `update` with `recurrence` fails fast with
  the reason and tells the caller to recreate the task; `clear_recurrence` maps to `recurrence: null`.
- `$select=recurrence` on a task is refused (`400 RequestBroker--ParseUri`), so reads that need the
  rule fetch the whole task.

**Junk / safe senders — the consumer gap of this run.** Nine probes, all recorded in
`src/tools/manage-senders.ts` and in the README table:
`GET /beta/me/blockedSenders` → `404 UnknownError`; `GET /beta/me/safeSenders` → `404`;
`GET /beta/me/outlook/blockedSenders` → `400 "Resource not found for the segment 'blockedSenders'"`;
`GET /beta/me/mailboxSettings?$select=blockedSenders,safeSenders` → `400 "Could not find a property
named 'blockedSenders' on type 'microsoft.graph.mailboxSettings'"`;
`GET /beta/me/mailboxSettings/junkMailRule` → `400 "Resource not found for the segment
'junkMailRule'"`; `GET https://outlook.office.com/api/beta/me/blockedsenders` → `401` (a different
resource audience — an outlook.office.com token would need consent this connector does not have, and
adding scopes was out of scope for this run).
What **does** work is beta-only and message-scoped: `POST /beta/me/messages/{id}/markAsJunk` → `202
Accepted` and `POST /beta/me/messages/{id}/markAsNotJunk` → `202`. The v1.0 endpoint rejects the
`moveToJunk` parameter outright (`400 RequestBodyRead`), so the tool calls beta explicitly by full
URL.
**Decision: ship the tool, shrunk to what the platform actually does.** `manage_senders` offers
`block_sender` and `unblock_sender` only. `list`, `safe_sender` and `unsafe_sender` are **not** in the
enum — an action that can only ever fail is worse than an honest absence, because a model will call
it and then report a failure to the user as if the mailbox were broken. The description states all
three limits (no readable lists, no safe-sender management, per-message rather than per-address) and
the output repeats that the result cannot be verified through Graph and points at Outlook web.
- `move_message` (default true) maps to `moveToJunk` / `moveToInbox`. With it false the probe
  confirmed the message stays in its folder, which is what makes the round-trip test safe to run
  against a real correspondent.
- A draft has no sender: the tool refuses before calling Graph rather than sending a request that
  cannot mean anything.

**Focused Inbox — fully supported**, contrary to the doubt in the brief.
`GET/POST/DELETE /me/inferenceClassification/overrides` all worked on v1.0 (200/201/204) on this
consumer account, so no beta and no fallback is needed. Graph refuses a second override for the same
address, so `set_focus_override` PATCHes an existing one instead of colliding.

**Working hours — supported, with one silent normalisation.** `PATCH /me/mailboxSettings` with
`workingHours` → 200 and `daysOfWeek` / `startTime` / `endTime` take effect. The **time zone does
not**: `{"name": "America/Toronto"}` went in and `Eastern Standard Time` came back. So the tool never
changes the zone — it carries the existing `timeZone` object through — and the README says why.

**MIME export.** `GET /me/messages/{id}/$value` → 200 `text/plain` with the full RFC 822 source
(31 KB for a typical newsletter). `internetMessageHeaders` and `replyTo` are both selectable in the
same `$select` as the rest of `read_message`'s fields, so headers cost no extra request.

### manage_task — why there is no delete-list action
The soft-delete policy already carves out a single exception for To Do tasks, because Graph has no
undelete for them. Deleting a *list* is that same irreversibility multiplied: one call destroys every
task in the list, with no recoverable copy anywhere in the account and no way for the model to have
enumerated (let alone shown the user) what it was about to take. A tool surface that can permanently
destroy work in bulk on one call is exactly what the rest of this server is built to avoid, so
`create_list` and `rename_list` ship and delete does not. The tool description says so and names the
To Do app as the place to do it deliberately. The **test harness** deletes its own lists with a raw
Graph `DELETE` — the same test-only escape hatch used to purge soft-deleted mail — and the test
comments say so, so the exclusion cannot be quietly re-introduced through the harness.

### auto_reply stays a separate tool; mailbox_settings is new
The brief allowed either "extend `auto_reply` into `mailbox_settings`" or "keep it separate", and
separate won:
- `auto_reply`'s vocabulary is `get` / `set` / `clear` **of one thing**. In a settings-wide tool,
  `set` would have to mean "set what?" — working hours, an override, or the out-of-office — and the
  disambiguation would live in the parameters rather than the action, which is exactly the shape that
  makes a model pick wrongly.
- Its description carries an outward-facing caution (every future sender sees the reply) that applies
  to none of the other settings. Diluting it into a general tool description weakens the warning where
  it matters most.
- Backward compatibility becomes structural rather than asserted: the tool, its schema and its tests
  are untouched by this batch, so "the existing tests still pass" is a fact about the diff, not a
  claim about an alias shim.
`mailbox_settings get` therefore reports auto-reply status read-only and names `auto_reply` as the
way to change it, so a model reading only one of the two descriptions still finds the other.

### export_message is a tool, not a flag on read_message
`read_message(export_eml: true)` was the alternative. A read tool that writes a megabyte to
`~/Downloads` (or burns a KV write) because a boolean was set is a side effect hiding inside a read,
and the two things are wanted at different moments: `include_headers` answers "is this spoofed?"
in-conversation, while an export exists to hand the message to someone else. A separate tool also
gets its own description, which is where the phishing-sample / evidential-copy framing belongs. It
reuses Batch A's machinery exactly: `saveToDownloads` locally (extracted from `get_attachment` into
`src/tools/save-local.ts` so both tools share one implementation) and `storeDownload` +
`/mcp/download/<id>` on the Worker, with the same 18 MB cap and ≤ 15 minute TTL.
`core/graph.ts` grew `callGraphServerBytes` for this: the request/retry half of `callGraphWithToken`
was split into `graphFetch`, so the byte path inherits the 429 handling and the request log rather
than fetching Graph on its own.

### read_message include_headers — what gets rendered and why
Sixty raw headers would drown the answer, so the section is a verdict, not a dump: the
`Authentication-Results` value reduced to `SPF/DKIM/DMARC/COMPAUTH` verdicts (raw kept, truncated at
300 chars), the `Received` chain **reversed** to oldest-first (Graph returns it newest-first, the
order hops are prepended) as one `from … by … — date` line per hop capped at 12, and an explicit
`** MISMATCH **` line when `Reply-To` — or the `Return-Path` domain — differs from `From`.
Two absences are reported rather than passed over in silence, because in a phishing check the absence
*is* the finding: no `Authentication-Results` header at all ("nothing vouched for this message"), and
a message with no internet headers whatsoever (a draft, or an item created in Outlook).
Off by default: it costs two extra fields on every read and most reads are not investigations.

### Harness (v8 tests)
- Local: 34 → **40 tests**, all self-cleaning.
  - `v8a` walks the whole To Do surface: create list → duplicate refused → repeating task (verified in
    Graph's own view) → the two recurrence guards (no due date, no change after creation) → two
    subtasks added, one completed **by text** and one removed **by id**, with the checklist re-rendered
    and cross-checked against Graph → `list_tasks include_subtasks` shows the repeat, the tally and the
    ids → `clear_recurrence` really clears it → rename keeps the list id → task deleted. The list
    itself is removed in a `finally` with a raw Graph `DELETE`.
  - `v8b` blocks and unblocks the sender of the newest inbox message with `move_message: false`, so a
    real correspondent's mail never moves, asserts the message did not change folder, unblocks in a
    `finally` so a mid-test failure cannot leave anyone blocked, and checks the draft (no sender) case.
    It cannot assert list membership — Graph exposes no read path, which is the finding itself — so it
    asserts the contract and the honesty of the output instead.
  - `v8c` sets a Focused-Inbox override for `mcp-test-focus@example.com`, finds it in
    `mailbox_settings get`, re-sets it to prove the update-not-duplicate path (one record, new value),
    clears it, and checks the second clear fails cleanly.
  - `v8d` saves `workingHours`, sets days/times, verifies Graph's own copy, checks the
    end-before-start and nothing-to-change guards, restores in a `finally` and asserts the restored
    object is **byte-identical** to the saved one — the same pattern as the auto-reply test.
  - `v8e` asserts the header section on real inbox mail (an `Authentication-Results` verdict, a
    compact hop-per-line chain, a reply-to verdict), that it is **absent** without the flag, and that a
    draft's empty header set is explained.
  - `v8f` exports a real message, parses the `.eml` back as RFC 822 (header block, `\r\n\r\n`
    separator, the five headers that must be there, a non-empty body), matches its `Message-ID`
    against Graph's, checks the bad-id error, and deletes the file.
  - Sweep: now also purges and asserts on `[MCP TEST]` **To Do lists** and `mcp-test-*`
    **Focused-Inbox overrides**, checks the local download directory holds no test file, and asserts
    **working hours** are restored exactly alongside auto-reply.
  - `auto_reply`'s own test (`g`) is untouched and still passes — the backward-compatibility proof.
  - The stdio smoke test now expects **27 tools** and the three new names.
- Remote: 22 → **23 tests**. `r23` (`testAuthed`) exports a real message through the Worker: no
  "Saved to", a `…/mcp/download/<64 hex>` link, anonymous → 401, bearer → 200 with
  `content-type: message/rfc822` and a `.eml` `Content-Disposition`, and the served bytes compared
  **byte-for-byte** against Graph's own `$value` fetched with this machine's token. It drops the KV
  record in a `finally` rather than waiting out the TTL. Headless: 23/23 with 13 skips.

### Docs/versioning
- `package.json` and `src/core/version.ts` → **0.8.0** (the harness asserts they match).
- README: v8 header, 27-tool table, three new sections (mailbox settings; junk senders, with the probe
  table above; message forensics), the To Do notes extended with subtasks, repeating tasks and the
  no-delete-list rule, and the soft-delete section extended with the same rule.

## v9 batch C — LLM-classified filing and a drafted morning brief

Two features that call a language model on the mailbox's own content, on the hosted Worker only.
**Both ship disabled.** The whole batch is written on the assumption that some of the mail is trying
to hijack the model reading it, so most of the decisions below are about making that assumption
structurally true rather than merely stated.

### The model
- **`claude-haiku-4-5`**, verified live against the deployed Worker on 2026-08-19. The API accepts the
  bare alias and resolves it to **`claude-haiku-4-5-20251001`**, which is what comes back in the
  response's `model` field and therefore what the audit log records. The brief named the dated id; the
  alias is used in code because current Anthropic guidance is that the published ids are complete as
  written and date suffixes should not be appended by hand. Both name the same snapshot, so this is a
  naming choice with no behavioural difference — and the live check is what settled it rather than an
  assumption.
- Pricing $1 / $5 per million tokens in/out. Measured on this mailbox: **783 input, ~48–63 output
  tokens per classification, ≈ $0.001 per message**; a digest ≈ $0.005/day. The README carries the
  monthly table; the 200/day cap puts the worst case at ≈ $6.35/month.
- **Raw `fetch`, not `@anthropic-ai/sdk`.** One endpoint is all either feature needs and the module is
  bundled into workerd; a dependency would buy nothing. `core/anthropic.ts` mirrors `core/graph.ts`:
  no Node-only imports, one retry on 429/5xx honouring `Retry-After`, and an error type that carries
  the status and Anthropic's own message but **never the key**.
- **No structured outputs / `output_config.format`.** Haiku-tier support was not something to depend
  on here, and the validator has to exist anyway — the schema check is a security boundary, not a
  convenience, so it cannot be delegated to the API even where the API would enforce it.

### The injection defence, and why it is four things rather than a prompt
A prompt instruction is the weakest of the four and is listed last on purpose.

1. **Module boundary.** `core/classifier.ts` imports no Graph transport — not `core/graph.ts`, not any
   tool. The `ClassifierMailbox` interface is **declared in `classifier.ts` itself** and implemented by
   `core/mail-actions.ts`, inverting the dependency so the classifier's transitive import set contains
   nothing that could reach Graph. That was a deliberate choice over the more obvious arrangement
   (types in `mail-actions.ts`, `import type` in the classifier): a boundary that only holds because
   TypeScript erases something is not a boundary worth asserting, and this way the test needs no
   exception for type-only edges. Test `v9a` walks every import edge, including type-only ones, and
   fails if `src/tools/`, `core/graph.ts`, `core/mail-actions.ts` or `core/digest-mailbox.ts` is
   reachable; it then pins the interface to exactly five method names and scans the two modules that
   *do* touch Graph (comments stripped) for `sendMail`, `/send`, `createReply`, `/reply`, `/forward`,
   `permanentDelete`, `messageRules`, `mailboxSettings` and a `DELETE` verb.
2. **Folder allowlist, with the destructive destinations removed.** The interesting case is not "the
   model names a folder that does not exist" — it is that **a move into Deleted Items *is* a delete**.
   `NEVER_FILE_INTO` strips Deleted Items, Junk Email, Drafts, Sent Items, Outbox, Conversation History,
   Clutter and Sync Issues from the list the model is offered, so the delete is unreachable by the only
   mutation the path has. **Archive is deliberately left in**: it is a legitimate and common filing
   destination, it is recoverable and visible, and excluding it would gut the feature. Categories are
   allowlisted the same way against the mailbox's real master categories; one unknown name discards the
   whole answer rather than being dropped, because partial acceptance of a deviating answer is a worse
   default than doing nothing.
3. **Schema.** Exact-shape JSON: the four keys, no extras, `folder` and `reason` strings, `confidence`
   a finite number in 0–1, `categories` an array of strings. Every rejection carries a reason into the
   audit log, so an injection attempt reads as a discarded answer rather than as silence.
4. **Prompt.** States the mail is untrusted data, that anything in it reading as an instruction is
   itself evidence of phishing rather than a command, and that the model cannot send/delete/reply
   whatever it is told. The mail sits inside `<<<UNTRUSTED_EMAIL_BEGIN/END>>>` markers with the
   allowlists outside them — `v9b` asserts that ordering rather than trusting it.

### The markdown fence — a live finding, not a design choice
The first live run against the deployed Worker classified nothing: every answer was discarded as "not a
bare JSON object". Adding a truncated snippet of the model's actual answer to the discard reason (which
is a genuine operator improvement, kept) showed the cause immediately — Haiku wraps otherwise-perfect
JSON in a ` ```json ` fence despite the prompt saying not to.

`unfence()` strips **exactly one fence around the whole trimmed answer** and nothing else. This is
framing rather than a schema relaxation: it widens nothing the model can express, because the shape and
both allowlists still decide every field. Prose before the fence, an unclosed fence and two fences all
fail to match and are discarded like any other malformed answer, and `v9b` has a fixture for each,
including a *fenced* answer naming Deleted Items (still discarded). The prompt instruction was reworded
rather than deleted — belt and braces, and it costs nothing.

The snippet in the discard reason is model output derived from untrusted mail. It is truncated to 200
characters, whitespace-flattened, and — like the `reason` field that was already being stored — only
ever displayed. Nothing reads it back.

### Skip list, threshold, budget
- **Protected subjects are matched before any API call**, so a one-time passcode is never sent
  anywhere, and (asserted in `v9c`) does not even consume a call from the daily budget. The compiled-in
  list covers one-time/single-use/verification/security codes, verify-log-in, two-factor, 2FA and
  password resets. `add_skip_pattern` extends it; the built-in half cannot be removed, and
  `remove_skip_pattern` says so by name when asked to remove one.
- **Threshold 0.8 by default**, floor 0.5, ceiling 1. Below it the classifier logs its reasoning and
  does nothing — which is the correct outcome when the model is unsure, and is why the prompt tells it
  to be conservative rather than decisive.
- **Daily cap 200 calls per America/Toronto day**, across both features, read-modify-write in KV with a
  2-day TTL. Same lost-update reasoning as the activity ring: the cap exists to stop a runaway loop,
  not to bill to the cent. `set_daily_cap 0` stops all API calls without touching the enable flags.
- Bodies truncated to 2,000 characters, subjects to 300, folder list capped at 60, 300 output tokens.

### The digest
- **Draft, never sent.** `DigestMailbox` has no send method — the same structural argument as the
  classifier, and `v9a` scans `core/digest-mailbox.ts` for send verbs *and* for `/move`, since the
  digest has no business moving anything either. `send_draft` remains the only send path in the repo.
- Separate port module from the classifier's on purpose: neither feature can reach the other's
  capabilities. The classifier cannot draft; the digest cannot move or categorize.
- **Overnight = the last 14 hours**, which at 07:00 is 17:00 the previous day. A wall-clock "since
  18:00 yesterday" would need its own DST reasoning for no benefit.
- The brief's body carries a footer naming the model, the counts it was built from, and the fact that
  it was never sent and that mail was treated as untrusted data — so the artifact explains itself to
  someone who finds it in Drafts without context.

### Cron and DST — two schedules plus a guard
Cloudflare crons are UTC only. 07:00 America/Toronto is **11:00 UTC in EDT** and **12:00 UTC in EST**,
so a single schedule drifts by an hour twice a year. Three options were on the table: one schedule and
accept the drift, one schedule edited twice a year, or both schedules with a guard. The third was
chosen: `0 11 * * *` and `0 12 * * *` both fire year-round and the `scheduled` handler computes the
America/Toronto hour from `event.scheduledTime` and drops whichever tick is not 07:00 locally. Nothing
drifts, nothing needs redeploying, and `v9c` asserts both directions (11:00 UTC is 07:00 in August,
12:00 UTC is 07:00 in January, 11:00 UTC is *not* 07:00 in January).

Belt and braces on top: `runDailyDigest` refuses a second brief for a Toronto date it has already
covered (`llm:digest:last`), so even a double fire produces one draft. The 6-hourly subscription upkeep
now runs only on the ticks that are *not* the digest hour, which is a clarity choice — running it on all
five would have been harmless, since it is idempotent.

### Tool surface
- **`manage_auto_filing` controls both features** rather than splitting into two tools. The brief
  allowed a rename; the name was kept because it is what the gate and the docs refer to, and because
  the two features share every rail that matters (the key, the daily budget, the audit log). The
  digest's flag is fully independent of filing's, as required.
- **Both tools are hosted-only**, in the same shape `get_mailbox_activity` established: the features
  run on the Worker off KV state, so a local stdio call would write settings nothing reads. They say
  that in full rather than silently succeeding. `v9c` asserts both refusals.
- `get_auto_filing_log` shows **no-action entries by default** (`actions_only` opts out). That is the
  point of the log: seeing what the model declined to do, and why, is how you decide whether to trust
  it. An action-only default would hide exactly the evidence an injection attempt leaves.
- Neither tool can enable anything by accident: `readLlmConfig` treats a missing, corrupt or partial
  record as **both features off**, asserted in `v9c` with deliberately corrupt JSON.

### Notification wiring
`handleNotificationRequest` gained an `onAccepted` callback, called synchronously after the ring-buffer
write and never awaited — Graph retries anything not answered promptly with 2xx, so the Worker hands
the classification to `ctx.waitUntil` and answers 202 immediately. A throw from the callback is caught
and logged: follow-up work must not cost us the acknowledgement. At most 5 messages per delivery are
classified. `defaultHandler.fetch` gained the `ctx` parameter to carry `waitUntil` down to the route.

With filing disabled — the shipped state — the callback reads one KV key and returns, so the added cost
to the notification path on a default deployment is one KV read.

### Tests
- Local: 40 → **45** (44 run + 1 skip on a default checkout).
  - `v9a` the module boundary, described above.
  - `v9b` fixtures with a canned model answer. The happy path files a receipt first, so "no action
    everywhere" cannot pass by accident. Then: a body ordering the model to forward and delete, a
    captured model naming Deleted Items, prose, prose-before-a-fence, an unclosed fence, an array, a
    missing key, an extra key, a stringly confidence, confidence 42, non-string categories, a
    non-allowlist folder, a folder differing only in case, an invented category, sub-threshold
    confidence, the no-folder sentinel — every one asserted to leave the mailbox untouched *and* to be
    audited with a reason. Plus the marker-ordering assertion and the threshold in both directions.
  - `v9c` the rails: defaults off, corrupt config off, disabled means the model is never called,
    protected subjects skipped before the API and off-budget, custom skip patterns, the daily cap
    counting/resetting/hard-stopping, body truncation, the audit ring capping newest-first, both tools'
    honest refusal on a local store, and the Toronto date/hour helpers.
  - `v9d` drives the digest against the **real mailbox** with a canned brief and asserts the draft
    exists, is a draft, is addressed to this mailbox, is titled `Morning brief — <date>` and carries the
    footer; that a second run the same day is refused; and that the entry is audited with its usage. The
    sentinel date is **2099-01-01** so the sweep can match the subject without ever touching a genuine
    brief — the brief's subject deliberately carries no `[MCP TEST]` prefix, because that is the
    feature's real format.
  - `v9e` runs the **real production prompt against the real API**, gated on the key being in the
    environment or `.dev.vars` and skipped cleanly when it is not. The key is a deployed Worker secret
    and is not on this machine, so a default checkout **skips** it; the harness gained SKIP support and
    a summary line that reports skips, mirroring the remote suite. The live path is instead covered by
    `r25` and by the manual verification recorded below.
  - Sweep: purges and asserts on morning-brief drafts for the sentinel date.
  - stdio smoke test now expects **29 tools** and the two new names.
- Remote: 23 → **25**.
  - `r24` is the gate criterion as a test, and runs **unauthenticated** so it holds headless: it reads
    `llm:config` from the deployed KV and fails if either flag is on, asserts that `readLlmConfig`
    really does default both off (so an absent record cannot mean anything else), and checks both digest
    cron schedules are declared. On this deployment the key is **absent** — the shipped default.
  - `r25` (`testAuthed`) is the live end-to-end: enable through the tool, lower the threshold to 0.5 for
    the run, send a plain shop receipt to self, poll the audit log, then assert the model id, that
    tokens were spent, that the message really left the inbox, that Graph's folder matches the one the
    audit log names, that it is not on the never-file list, and that `get_auto_filing_log` shows it.
    A `finally` turns filing back off whatever happened.
  - The arrival poll gets **300 s**. Self-delivery plus a Graph notification lags well past a minute and
    a slow delivery is not a failure — the operational caveat from this run's batch brief.
  - `r20` restores `llm:config` and `llm:audit` to exactly what predated the run and then re-asserts the
    deployed server is back in the state `r24` requires.
  - Headless: **25/25, 14 skipped**.

### The threshold in `r25`, and a real behavioural finding
The first live classification of a `[MCP TEST]`-prefixed receipt came back at **0.6 confidence**, with
the model's own reason noting that the test marker in the subject looked like spam. That is the model
being appropriately cautious about an odd subject, not a bug — so `r25` lowers the threshold to 0.5 for
its own run rather than pretending a test-marked message is ordinary mail. Real mail with no such marker
classified at 0.75 in the same session.

### Live verification performed during this batch
Against the **deployed** Worker, headlessly, driving a synthetic notification with the real `clientState`
read from KV (which is exactly what Graph does, minus the wait):
1. Auto-filing enabled by writing `llm:config` directly, since the tool path needs an interactive bearer.
2. `POST /notifications` → `202 {"accepted":1,"discarded":0}`.
3. Audit entry appeared in KV: `action: "moved"`, `folder: "Archive"`, `confidence: 0.75`,
   `model: "claude-haiku-4-5-20251001"`, `usage: {input: 783, output: 48}`, reason "Transactional
   receipt email. No action needed. Safe to archive after review."
4. Graph confirmed the message had moved from Inbox to Archive.
5. `llm:config` deleted, probe mail purged, audit log and budget counter restored. `r24` then confirmed
   the deployed state is back to both-features-disabled.

### Docs/versioning
- `package.json` and `src/core/version.ts` → **0.9.0** (the harness asserts they match).
- README: v9 header, 29-tool table, and a new **LLM mail intelligence** section carrying the honest cost
  table, the enable/disable recipe, the four-mechanism injection explanation, the DST design and the
  API-key handling. The security model gains a bullet on why the auto-filing path is fenced in code
  rather than in a prompt — it is the one place where a model reads untrusted mail *and* acts without a
  human approving each call.
- `wrangler.jsonc`: `ANTHROPIC_API_KEY` documented in the secrets comment, and the two digest cron
  schedules documented alongside the upkeep one.

## v10 batch D — annotations, a doctor, and docs a stranger can follow

### The annotation scheme, and the judgment calls in it
Every tool now declares **all four** MCP hints. Stating them all was itself a decision: the protocol's
defaults for an absent hint are "not read-only, destructive, not idempotent, open-world", which is
wrong for the thirteen read-only tools and misleading for most of the rest. One rule defines each hint
(the rules are in `core/registry.ts`, the resulting table in the README):

- **`readOnlyHint`** — changes nothing: not the mailbox, not the server's own state, not local disk.
- **`destructiveHint`** — can remove or overwrite something the user would miss, or do something
  outward that cannot be taken back.
- **`idempotentHint`** — a repeat leaves the same state (set-shaped) rather than creating, appending or
  sending again.
- **`openWorldHint`** — the call, or the setting it establishes, moves data between this mailbox and
  parties outside it.

The calls that could reasonably have gone the other way, and why they went this way:

- **`openWorldHint` is not "talks to a remote API".** Under that reading all twenty-nine tools are
  open-world and the hint tells a caller nothing. The line drawn instead is the *mailbox boundary*,
  which makes the hint discriminating: seven tools have it, twenty-two do not.
- **The hint covers the setting a call establishes, not only the call.** `auto_reply` writes a mailbox
  setting and sends nothing — but the reply it turns on is delivered to everyone who writes to the
  account, so it is open-world. `manage_auto_filing` is open-world by the same rule: the switch commits
  the server to sending subjects and body excerpts to the Anthropic API. The alternative (hint the HTTP
  call only) would have marked the two tools that reach the outside world hardest as closed.
- **`manage_rules` is destructive but NOT open-world**, and that is only true because forwarding
  actions were deliberately left out of the tool in v3. If a rule could forward, this hint would have
  to flip. Worth stating: it is the one annotation that depends on an earlier design decision holding.
- **Soft deletes still count as destructive.** `manage_message` is destructive even though its delete
  is soft: the mail leaves where it was and comes back with a new id.
- **`idempotentHint` is read as PUT-shaped vs POST-shaped, not as "does a repeat error?"** By the
  letter of the spec, `create_folder` is idempotent — a duplicate name is refused and the environment
  is unchanged — and so, more alarmingly, is `send_draft`, whose second call finds no draft to send.
  Both are annotated **not** idempotent, because a caller reading the hint is asking "is retrying
  safe?", and for a create-or-send the honest answer is "it will not do what you meant".
- **Three tools are not read-only although they look it.** `check_new_mail` advances the stored delta
  position on every successful call — which is exactly the feature. `get_attachment` and
  `export_message` write a file to `~/Downloads` on stdio and a short-lived download record to KV on
  the Worker, with collision-safe names, so a repeat leaves a second copy and neither is idempotent.
  The cost is real: a client that auto-approves read-only tools will now prompt for `check_new_mail`.
  Accepted — the alternative is a hint that lies about a state change.
- **`manage_senders` is neither destructive nor open-world, and is idempotent.** Blocking is undone by
  unblocking and blocking twice is the same state; the junk list never leaves the mailbox. Recorded as
  a judgment call because its default `move_message: true` does move the message to Junk — reversibly,
  by the opposite action, which is why it did not tip the destructive hint.

The table is frozen twice on purpose: in `core/registry.ts` and again in the harness (`v10a`). A hint
cannot drift without someone changing it in both places deliberately.

### Getting the hints onto both wires — and what could not be tested headlessly
`registerTool`'s config gained `annotations`, so both transports carry them for free: they build the
same server from `registerAll`. Asserting it took three places:

- `v10a` checks the registry itself, plus the rules (a read-only tool may not claim to be destructive)
  and the two tools the safety story rests on.
- The stdio smoke test now compares the annotations that **actually came over the wire** against the
  same frozen table, so a serialization change cannot silently drop them.
- `r9` gained a per-tool comparison of the deployed `tools/list` against the local registry.

`r9` is `testAuthed`, and it cannot be otherwise: `/mcp` requires a bearer, production mints one only
through an interactive device-code sign-in, and there is deliberately no unauthenticated MCP surface to
read a tool list from. **So a headless run cannot verify the deployed annotations directly.** What it
can do is verify it is looking at the right build: `/health` now reports `version`, and the new
unauthenticated `r26` fails if the deployed Worker is not running this checkout's version. Headless
therefore proves "the deployed build is the one whose registry carries these hints", and the authed
`r9` proves the hints themselves. Serving the version anonymously is not a disclosure decision worth
agonizing over — `/health` still says nothing about the mailbox or whether anyone is authorized.

### `npm run doctor`
Four stages, in the order a broken install fails: environment (no network, no credentials), the
sign-in cached on disk, a live probe, and the deployed Worker. `--env-only` runs the first stage alone
— that is the stage a fresh clone can pass, and the one the batch's fresh-clone dry run exercises.
Exit code 1 if any check FAILs; WARN never fails the run.

- **Graph access tokens for a personal Microsoft account are not JWTs.** The first implementation
  decoded the `scp` claim to list granted scopes and got nothing back, because an MSA Graph token is an
  opaque compact ticket, not a JWT. The scope check now reads MSAL's own
  `AuthenticationResult.scopes`, which is authoritative, free, and works. (It reports `openid` and
  `profile` alongside the eight requested scopes; `missingScopes` therefore checks containment, not
  equality.) `auth.ts` gained `getGrantedScopes()` for this.
- **Two Graph probes, not one.** `/me` needs only `User.Read`, so it cannot tell a fully-consented
  sign-in from one that was granted the profile scope and nothing else. Reading the inbox folder is
  what proves the mail scopes were really consented to.
- **Translations, not error numbers.** `AADSTS70002` → public client flows are off; `AADSTS50020` and
  `AADSTS700016` → the app's audience excludes personal Microsoft accounts; a plain Graph `403` → the
  permission is missing from the registration, and re-running `npm run login` is required because
  consent is granted per scope at sign-in. These are the three failures this project actually hit, and
  an unexplained AADSTS number sends a stranger to a search engine.
- The deployment stage can only WARN: a local-only install has no Worker, and a Worker one version
  behind still works.
- The module is importable — the CLI runs only when `process.argv[1]` is this file — so `v10b` asserts
  the environment stage passes here, that each translation still fires, that a 404 is *not* explained
  as a permission problem, and the scope arithmetic.

### `package.json` and `core/version.ts`
The v9 notes claimed the harness asserted these two stay in sync. It did not. `v10b` now does, and the
stdio smoke test additionally asserts the version the server reports over `initialize` is the same one.

### Docs
- **`SETUP.md`** is new: zero to working, for someone who has never seen the repo. The Entra
  registration gets the detail it needs — personal-accounts audience, `allowPublicClient`, the eight
  delegated permissions, and a table mapping each AADSTS number to the setting that is wrong — then
  install and login, client configuration, and the optional Worker deploy and claude.ai connector as
  clearly-marked optional sections. The README's setup sections shrank to pointers so the two cannot
  drift; what stayed in the README is the reference material (the `seed:kv` caveat, the
  `ALLOW_DIRECT_AUTHORIZE` warning, the `resourceMetadata` exact-match requirement).
- **README restructured for a reader who has not seen it before**: an eight-row feature table by area,
  the security model promoted to the top as five structural claims (each linking to its full section,
  which was renamed "Security model in detail"), an architecture sketch, a cost statement in the
  overview, and the new **Tool annotations** section carrying the rules, the full table and the six
  calls worth explaining. Stale counts fixed throughout (23/27 tools → 29), and the version-labelled
  prose ("new in v8", "v9 adds") de-versioned — a stranger does not know what v8 was.
- `package.json` and `src/core/version.ts` → **0.10.0**.

## Enable-operate-publish run, Batch 1 — the LLM features turned ON (2026-08-19)

- The user explicitly confirmed enabling BOTH auto-filing and the morning digest, accepting the
  API cost. Enabled by writing `llm:config` to the Worker's KV (the same record
  `manage_auto_filing` writes) with everything else at the shipped defaults: threshold 0.8,
  daily call cap 200, no extra skip patterns. `/health` re-checked: still only
  status/service/version, nothing sensitive.
- **Digest smoke mechanism**: no local ANTHROPIC_API_KEY exists (the key is a deployed-Worker
  secret the user put there themselves), so the handler could not be invoked from Node with the
  real model. Instead a one-off token-gated route (`POST /internal/digest-smoke`, 64-hex
  random header token) was deployed UNCOMMITTED, called `draftMorningBrief(env)` exactly as the
  cron does (no force — config and idempotency checks intact), and was reverted and redeployed
  away within two minutes. Nothing entered git history; the route 404s now. The resulting
  "Morning brief — 2026-08-19" draft was verified via Graph (isDraft, addressed to the owner)
  and LEFT in the mailbox as the user's first real digest.
- **Auto-filing smoke**: one `[MCP TEST]` receipt probe sent-to-self at the DEFAULT 0.8
  threshold (not tuned, unlike r25 which drops to 0.5); classified `moved → Archive` at
  confidence 0.85, move verified via Graph, decision in the audit log, then the probe copies
  permanently deleted and the activity + audit ring buffers swept of test entries.
- **The remote suite's disabled-state assumptions were retired**: r24 now captures the live
  config and asserts only that it parses (the shipped-defaults-off assertion stays); r25's
  cleanup restores the exact captured config instead of calling `disable_filing`; r20 restores
  the captured config, asserts byte-equality with it, and sweeps `[MCP TEST]` audit entries
  while KEEPING real decisions made during the run. This exposed a latent bug: the harness's
  `kvGet` returned wrangler's trailing newline, so a captured value could never round-trip
  byte-equal through `kvPut` — fixed by stripping the one trailing newline.
- Live filer during the local suite: its probes were classified but produced no action
  (below threshold / model chose none), so the local tests are undisturbed by the enabled
  filer; the two test-marked audit entries were swept afterwards.

## Batch 2 — Operational maturity (v0.11.0)

Recorded while executing the "Batch 2" task on 2026-08-19: self-monitoring, rules backup, CI, and
their tests.

### Health-check cron: schedule and dispatch
- **`37 13 * * *`** (daily). 13:37 UTC is 09:37 Toronto in EDT and 08:37 in EST — always morning for
  the owner, after the digest ticks (11:00/12:00 UTC), and colliding with none of the upkeep ticks
  (minute 17 of hours 5/11/17/23 UTC). Unlike the digest it has no wall-clock meaning to preserve, so
  the DST drift of a fixed UTC hour is irrelevant and one schedule suffices.
- Dispatched on `event.cron === HEALTH_CRON` rather than the Toronto hour: the hour-based guard exists
  only because the digest needs the same local hour across DST, and adding a third hour-dispatched job
  would have entangled the two. The remaining ticks keep the existing hour dispatch unchanged.

### The health check itself (core/health.ts)
- Five checks, in failure-domain order: KV round-trip (a probe value put and read back, 5-minute TTL),
  a **forced refresh-token rotation** (`refreshMailboxToken` on the Worker — a real exchange, with the
  rotated token persisted first, exactly the steady-state path), the subscription named by `sub:mail`
  alive in Graph with a future expiry, and the two LLM error counters. Every dependency (store, Graph
  transport, rotation, drafting, clock, threshold) is injectable, which is what makes each failure mode
  unit-testable offline (o1–o7) and lets the remote e2e (r27) run the REAL checks from Node.
- **The alert is a DRAFT created directly in the inbox** (`POST /me/mailFolders/inbox/messages`,
  `isDraft: true`, addressed to the owner), never sent — `send_draft` remains the only send path in the
  repo. Subject `outlook-mcp health: <failing checks>`; body carries per-check detail, "since when"
  (carried forward from the previous heartbeat so a persistent failure keeps its original start time),
  and a fix pointer per check (re-seed procedure for token failures; `wrangler tail` /
  `get_auto_filing_log` for the rest). The draft is created BEFORE the heartbeat is written, so a dead
  KV — whose heartbeat write would fail — still produces the one signal it can.
- Healthy runs write only the heartbeat (`health:last`: timestamp, verdict, per-check results). No
  draft dedupe beyond the daily cadence: at most one draft per day while failing was judged the right
  amount of noise for a personal mailbox, and a deleted-but-unfixed alert resurfaces next morning.
- **Error counters** (`err:filing:<toronto-date>`, `err:digest:<toronto-date>`, JSON
  {count, firstAt, lastAt, lastReason}, 2-day TTL — the same per-Toronto-day + TTL design as the API
  budget): incremented by `recordFeatureError` in exactly the paths that swallow failures — the
  classifier's Anthropic-call catch and its outer catch, and the digest's equivalents. The recorder
  itself never throws (counting an error must not cause one), and it lives in `core/auto-filing.ts`
  because the classifier's import boundary (v9a/o15) forbids it from reaching anything that can touch
  Graph — `core/health.ts` imports the Graph transport and is asserted NOT reachable from the
  classifier. Threshold: 5 per feature per Toronto day.
- `get_health` is annotated fully read-only: it reads the heartbeat (remote) or performs two Graph GETs
  (local). The KV probe and the draft belong to the cron. Local mode runs only what is honest locally —
  silent sign-in and inbox access — and names the five remote-only checks instead of faking them.

### Rules backup (manage_rules export / import)
- Portable format `outlook-mcp-rules/1`: Graph's own field shapes (displayName, sequence, isEnabled,
  conditions, exceptions, actions) plus the rule id for matching; volatile/read-only Graph fields are
  dropped. Export always returns the JSON inline; the LOCAL transport also writes
  `~/Downloads/outlook-mcp-attachments/inbox-rules-<toronto-date>.json` via the same `saveToDownloads`
  path as export_message (the Worker has no filesystem and says so).
- Import is **dry-run by default** (`apply: true` applies) and **never deletes**: matching is by rule id
  first, then displayName case-insensitively (covers a deleted-and-recreated rule); live rules absent
  from the backup land in a `liveOnly` list that both the dry-run and apply outputs print under an
  explicit "import never deletes" heading. Updates are field-level PATCHes in place (id and sequence
  preserved); creates POST with the backup's sequence.
- **Forwarding discipline holds in both directions**: export records externally created forward rules
  faithfully but warns; import refuses the whole backup while any entry carries
  forwardTo/forwardAsAttachmentTo/redirectTo (empty arrays don't count). The create/update guards also
  apply on the way in: no conditionless rule may be created or patched (it would match ALL mail), no
  actionless rule created.
- **Diff normalization**: keys sorted, null/empty values dropped, and `senderContains` values uppercased
  before comparison — Graph stores them uppercased (a v4 finding), so without this a backup would diff
  forever against the very rules it created. Asserted in o9.

### The non-live tier split (npm run test:offline) and CI
- A separate entry point (`src/test-offline.ts`, 17 tests) rather than an env flag: `test-tools.ts`
  acquires a live token and reads the mailbox at module load, so flag-gating it would have meant
  restructuring 4,000 lines for no gain. The tier's rule is stated in its header: no Graph, no MSAL
  cache, no KV, no .env, no secrets — it deliberately does not import `src/auth.ts`, so a test that
  accidentally reaches for Graph fails with AuthRequiredError instead of quietly needing a credential.
- Contents: the health check's failure modes (o1–o7: healthy heartbeat, missing/expired/404
  subscription, rotation failure, counters over threshold, KV unreachable, failingSince carry), the
  rules-backup validation/diff (o8–o9), compact classifier fixtures + rails + digest offline (o10–o12,
  including the new error-counter wiring), subscription upkeep and the webhook handshake against stubs
  (o13–o14), the classifier import-boundary walk (o15), annotation-rule assertions (o16 — structural
  rules only; the frozen per-tool table stays duplicated in test-tools.ts, a third copy would add drift
  surface, not safety), and version sync (o17).
- CI (`.github/workflows/ci.yml`): on push — checkout, Node 22, `npm ci`, `npm run typecheck`,
  `npm run test:offline`. No live-mailbox tests, no secrets.
- **CI validation method**: `act` is not installed on this machine, so the workflow was validated by a
  fresh-tree simulation — `git add -A && git write-tree`, `git archive` of that tree into a scratch
  directory (which, like a CI checkout, contains no .env, no .token-cache.json, no .dev.vars), then the
  workflow's exact steps (`npm ci`, `npm run typecheck`, `npm run test:offline`) with AZURE_CLIENT_ID
  and ANTHROPIC_API_KEY explicitly unset. All three steps passed (17/17). Local Node is 24 vs CI's 22;
  nothing in the tier is version-sensitive beyond ES2022.

### The live e2e (r27) and the gate heartbeat
- r27 is headless-runnable by design, which is also how the gate's "real heartbeat" was produced: the
  transport-agnostic `runHealthCheck` runs from Node with a wrangler-backed StateStore over the deployed
  `OUTLOOK_KV`, the local-MSAL Graph transport, and a REAL forced rotation of the KV refresh token (the
  same exchange the Worker performs, rotated token written back first; the consented scope list is
  mirrored in the test with a comment, since `src/worker/ms-token.ts` cannot be imported into the Node
  tsconfig without dragging in workerd's conflicting type universe).
- Flow: capture `sub:mail` byte-exact → overwrite its id with `MCPTEST-bogus-<hex>` (the [MCP TEST]
  marking of the forced state; r20's sweep asserts no MCPTEST remains in the record) → run the checks →
  subscription fails, the other four pass live → the alert draft verified via Graph (isDraft, in the
  INBOX, addressed to the owner, subject/body content) → restore and permanently delete the draft in a
  `finally` → verify the restored record names a live Graph subscription → run the checks AGAIN, all
  green, and assert the healthy heartbeat is the one in deployed KV.
- Restore-race handling: if a concurrent upkeep rewrote `sub:mail` while the bogus record was in place
  (possible only if the 6-hourly cron ticks during the ~1-minute window — no authed MCP requests happen
  during r27), the racer's fresh record is kept instead of clobbered; the liveness assertion covers both
  outcomes.
- The bogus-id subscription check fails on Graph's 400 (malformed id) rather than a 404 — the check
  treats any GraphError on the GET as "not live", which is the honest reading (it cannot be verified).

### Docs/versioning
- `package.json` and `src/core/version.ts` → **0.11.0**; deployed and verified via `/health` and r26.
- Tool count 29 → 30 everywhere it is stated (README header/architecture/annotations table/Claude
  Desktop note, stdio smoke test, offline annotation test); `get_health` added to the frozen annotation
  table in test-tools.ts as read-only.

## Batch 3 — Published (v1.0.0)

Recorded while executing the "Batch 3" publication task on 2026-08-19: the pre-publish history scan,
the packaging decisions, and the public release.

### Pre-publish scan: tooling and findings
- **Tooling.** (1) `gitleaks 8.30.1` over ALL git history (`gitleaks git .` — 60 commits, ~1.76 MB):
  **no leaks found**, both before and after the rewrite below. (2) Manual full-history greps over
  `git log -p --all`: token patterns (`eyJ…` JWTs, `EwB…`/`EwA…` and `M.C5`/`M.R3`/`0.A…` MSA token
  shapes, `Bearer <long>`, `sk-`/`sk-ant-`, `ghp_`/`gho_`, `AKIA…`, `client_secret`) — zero hits
  (the only `refresh_token` matches are the KV **key name** constant `ms:refresh_token`);
  a complete inventory of every email address, GUID and 32-hex string ever committed, each reviewed
  individually (below). (3) A per-path inventory (`git log --all --name-only --diff-filter=A`)
  proving `.env`, `.dev.vars`, `.token-cache.json`, `.mcp-state.json` and any KV dump/seed file
  **never appeared in any commit**. (4) A docs sweep for quoted mailbox content: no real subjects,
  senders or bodies — fixture data only (`[MCP TEST]`, `.invalid`/`example.com` addresses).
- **One finding, scrubbed.** `[redacted-second-alias]` — the owner's own University of Waterloo alias,
  not repeated here for the same reason it was scrubbed —
  appeared on one line of ASSUMPTIONS.md (the app-registration tenant notes) throughout history. It is
  not a secret, but it is a second personal identifier that was never part of the publish decision, so
  it was removed rather than judged: `git filter-repo --replace-text` rewrote all 60 commits, replacing
  the string with `[redacted-second-alias]` (pre-rewrite bundle kept locally outside the repo). This
  happened BEFORE the repo had any remote, so no force-push was involved. **Every commit hash changed**
  (the v0.11.0 Batch 2 commit `9f967e5` became `d641a5a`). Re-scan after the rewrite: gitleaks clean,
  and a history-wide grep for the alias returns zero.
- **Reviewed and accepted (deliberately NOT scrubbed).** The owner's `arthur.yuhao.zhang@outlook.com`
  (the documented mailbox and git author identity) and the `arthur-yuhao-zhang.workers.dev` URL — the
  public-deploy decision already made; the app registration's client id `1d362aa5-…` and its
  Default-Directory tenant id `a289df25-…` (public-client identifiers, documented in this file, not
  credentials); Microsoft's well-known consumers tenant `f8cdef31-…`; the Graph subscription id
  `95f07422-…` and Worker version ids (identifiers, useless without a token); the Cloudflare account
  id `cf54199…` (visible in every dashboard URL, not a secret); the two KV namespace ids in
  wrangler.jsonc (already documented as not secrets); the wrangler-types hash; and two **truncated
  12-hex prefixes of rotated-out refresh tokens** quoted in this file's Batch-C evidence
  (`cfc64c50cb59…` → `382d3b854bc2…`) — both superseded by later rotations and far too short to
  reconstruct anything. No real `clientState` value appears anywhere (fixtures and prose only).

### Packaging decisions
- **LICENSE**: MIT, © 2026 Arthur Zhang.
- **`"private": true` KEPT in package.json.** The repo is public on GitHub but is not an npm package;
  the flag's only effect is to make an accidental `npm publish` fail, which is exactly the guard a
  public-source, non-npm project wants. Removing it would only remove that guard.
- **README first screen** now carries the one-paragraph security model (untrusted-input stance,
  two-step send, soft deletes, no rule forwarding, single-user interactive-only endpoint, fenced
  auto-filing) alongside what-it-is, personal-account support and the SETUP.md link.
- **SECURITY.md**: private reporting via GitHub security advisories (preferred) or email, plus the
  threat-model summary distilled from the README. **CONTRIBUTING.md**: personal tool, PRs/issues
  welcome, the credential-free `typecheck` + `test:offline` pair (what CI runs), and the warning that
  the live suites need your own mailbox and deployment.

### Release
- Repo: **https://github.com/8C9D/outlook-mcp** (public). CI (typecheck + offline tier) green on the
  first push: https://github.com/8C9D/outlook-mcp/actions/runs/32278532958 (main, commit `fcbbfe5`);
  the tag push's run also green (32278533936).
- Version 1.0.0 in package.json AND src/core/version.ts; `npm run deploy` after the bump (same code,
  version only) — `/health` reports `1.0.0` (Worker version id 4e3ad855-4a36-430f-9003-60736d516d60).
  Annotated tag `v1.0.0` on the release commit, pushed.
- Suites after all changes: typecheck green; `test:offline` 17/17; `test:tools` 49/49 (1 designed
  skip); `test:remote` headless 27/27 (14 auth-gated skips). `llm:config` untouched.

## Batch 4 — feature completions (v1.1.0)

Recorded while executing the "Batch 4" task on 2026-08-19: the auto-filer's feedback loop, search
depth, delete_folder, and MCP structured content.

### Correction detection: reconciliation, not events — and where it runs
- **Mechanism chosen: periodic reconciliation of the audit log against reality.** Recent "moved"
  audit entries (the filer's own moves, which since this batch record `folderId`, `newMessageId` and
  `sender`) are compared against each message's CURRENT `parentFolderId`; a mismatch is the user's
  correction. The alternative — change-notification subscriptions on destination folders — was
  rejected: it needs one Graph subscription per folder (against a per-mailbox cap, each with its own
  clientState/renewal lifecycle), and it still cannot see a move between two folders that carry no
  subscription. Reconciliation is one bounded read per watched move, transport-agnostic, and can only
  see a correction late, never miss it.
- **It runs where the filer runs, on the filer's own triggers**: at the top of every accepted
  notification delivery (BEFORE that delivery's messages are classified — so a correction is learned
  before the very next message from that sender is filed, and the remote e2e can drive the whole loop
  through the production webhook with no test-only route), and on the 6-hourly upkeep cron ticks for
  quiet stretches. **No temporary route was needed for the e2e** — r29 sends a second probe and the
  probe's own delivery performs the reconcile, exactly as production does.
- Bounds: only entries younger than 72 h (`RECONCILE_WINDOW_MS`) are watched, at most 20 message
  reads per pass; an entry still in place after 24 h (`CONFIRM_AFTER_MS`) is marked `confirmed` and
  never re-read. Terminal marks on the audit entry (`confirmed` / `corrected` / `gone` / `ignored`)
  are what stop re-checking; the ring is read once and written once per pass (same unlocked
  read-modify-write caveat as every ring buffer here — accepted at single-mailbox scale).
- **What does NOT teach a preference**: a message that vanished (deleted, or its folder was) — a
  delete is not a filing choice; a move into any NEVER_FILE_INTO folder (Deleted Items, Junk, Sent,
  …); and any message whose subject matches the protected (OTP/verification) patterns — the
  compiled-in + configured skip list outranks the feedback loop in BOTH directions: the reconciler
  refuses to learn from protected mail, and classifyAndFile checks the skip list BEFORE the
  preference lookup, so a preference can never act on it either.

### Preference schema and semantics
- KV key `llm:prefs`: an array of `{sender, folderId, folderName, corrections, firstAt, lastAt}`,
  capped at 200 entries (most-recently-touched kept), keyed on the sender's **bare email address,
  lowercased** (`extractAddress` pulls it out of `Name <addr>` forms). Sender-exact rather than
  domain-pattern: it is the signal the correction actually carries; pattern generalization was left
  out rather than guessed at.
- **A preference acts from the FIRST correction.** The user deliberately re-filing a message the
  filer placed is an explicit instruction; making the filer repeat a known-wrong move until a second
  correction arrives would be the feature ignoring its user (and the mandated e2e — one correction,
  then a preference-filed second probe — assumes exactly this). "Repeated corrections … become a
  standing preference" is therefore rendered as: a repeat correction to the same folder increments
  `corrections` and marks the preference **standing** (`corrections >= 2`, shown by
  `list_preferences`); a correction to a DIFFERENT folder replaces the preference and restarts the
  count — the user's latest choice always wins.
- **A correction back to the Inbox is itself a preference**: "leave this sender's mail alone". It is
  stored like any other (folderId = the inbox's), and a hit logs `action: none, source: preference`
  with no model call, rather than re-filing mail the user pulled back.
- On every hit the target folder is re-validated (`getFolder`: still exists, not on NEVER_FILE_INTO);
  a stale or unsafe preference falls through to the model instead of acting.
- The audit `source` field ("llm" vs "preference") is the observable contract: preference decisions
  carry NO `model`/`usage` fields and consume no budget — asserted offline (o18) and live (r29).
- The classifier-side boundary holds: `core/corrections.ts` imports no Graph transport and acts only
  through the same `ClassifierMailbox` port, which grew two READ methods — `getFolder` and
  `findByConversation` — so the port is now "seven methods, still two mutating (move, categorize)".
  o15/v9a walk the new module and pin the new method list.
- **Found live by the first e2e run, then fixed: the user's correction move invalidates the very id
  the reconciler was watching.** A Graph move mints a new message id and the old one 404s
  immediately (probed directly). Without recovery the reconciler marked corrected messages "gone"
  and learned nothing — in production as well as in the test, since real corrections are moves too.
  Fix: `conversationId` (verified stable across moves) is now recorded on every filed move's audit
  entry, and when the direct read misses, the message is re-found via `findByConversation`, keeping
  only candidates with the same audit-truncated subject and skipping copies in never-file folders
  (so the Sent Items copy of self-sent mail can never be mistaken for the filed one).

### Search semantics (verified live on this mailbox, 2026-08-19)
- `$search` combines with **neither** `$filter` (400 `SearchWithFilter`) nor `$orderby`. Query-mode
  filters therefore ride INSIDE the KQL: `received>=YYYY-MM-DD`, `received<=YYYY-MM-DD` and
  `hasattachments:true|false` all work in `$search` (probed individually and in AND-combinations).
- KQL dates are day-granular in an unspecified timezone, so the KQL range is **widened one day each
  way** and the exact America/Toronto boundary is enforced client-side on `receivedDateTime`
  (`torontoMidnightUtc` tries the two possible offsets and asks Intl which one Toronto agrees with —
  exact across DST, since Toronto's transitions happen at 02:00, never midnight). Query mode
  over-fetches (up to 3× max_results, ≤100) before the exact post-filter, then slices.
- `$filter` + `$orderby=receivedDateTime desc` works **only when receivedDateTime leads the filter**:
  `hasAttachments eq true` alone earns `400 InefficientFilter`; prefixing a sentinel
  `receivedDateTime ge 1900-01-01T00:00:00Z` clause fixes it (probed both ways). Latest mode's
  filtering is exact server-side (UTC instants of the Toronto dates), with the same client-side
  check kept as belt and braces.
- `all_folders` maps to `/me/messages` (which includes Sent Items and Deleted Items — stated in the
  schema and output); the default stays the single-folder path, so existing calls behave
  identically. `date_from`/`date_to` are inclusive America/Toronto calendar dates (the mailbox's
  documented timezone), `date_to` implemented as exclusive-next-midnight.

### delete_folder: what Graph actually does, and the tool's honest shape
- **Verified live (twice, with a 20 s consistency wait): `DELETE /me/mailFolders/{id}` on this
  personal account is a PERMANENT delete.** The folder 404s immediately, nothing appears under
  Deleted Items' childFolders, and a message that was inside is gone from `/me/messages` entirely.
  The task brief's expectation ("Graph folder delete moves it to Deleted Items") is FALSE for
  consumer mailboxes, so the tool **never issues that DELETE**. The soft delete is
  `POST /me/mailFolders/{id}/move {destinationId: "deleteditems"}` — verified: the folder keeps its
  id, lands under Deleted Items with its contents intact and reachable. The tool description states
  all of this.
- Guards, in order: well-known folders always refused — matched **by id** against the eleven
  well-known names resolved in ONE `$batch` request (consumer mailFolder has no `wellKnownName`
  property — `$select`ing it 400s, verified; and eleven parallel GETs drew per-mailbox 429s in the
  first live run, hence the batch); subfolders always refused, force or not (no folder tree in one
  call); non-empty refused without `force`; `force` moves the messages (≤500, batched 20 per
  `$batch`) into Deleted Items FIRST — individually visible there, not buried in the deleted folder
  — and the result says so; if any message move fails the folder is NOT deleted.
- Annotated destructive (a soft delete still counts, per the v10 rule), not idempotent, closed-world.

### Structured content: SDK mechanics and the permissive-schema rule
- SDK 1.30.0's `registerTool` accepts `outputSchema` (zod raw shape) and serves it as JSON Schema in
  `tools/list`; handlers return `structuredContent` beside `content`. Two SDK behaviors shaped the
  implementation: the server **requires** structuredContent on every non-error result of a tool that
  declares an outputSchema (a miss becomes a protocol error, not an isError result — so every
  success path, including "no results" and "no report yet", attaches it), and it validates the
  payload against the schema server-side. Schemas are therefore **permissive by rule**: every field
  optional, arrays of `z.looseObject`s, unknown keys tolerated — o21 asserts each schema accepts
  both `{}` and unknown keys, so a schema-validating client can never see a previously-working call
  fail.
- Scope: the five tools whose answers are data a client might render — search_mail, list_folders,
  list_events, list_tasks, get_health. The rest are prose-shaped confirmations and stay text-only.
  isError results deliberately carry no structuredContent (the SDK exempts them). Both transports
  serve identical schemas/payloads by construction (shared registry); asserted over stdio (smoke
  test) and Streamable HTTP (r28).

### Test-side notes
- The remote feedback e2e (r29) is **fully headless**: config via KV write (capture → mutate →
  restore byte-exact, the r24/r25/r20 pattern, now also applied to `llm:prefs`), probes drafted and
  sent via local Graph (the same two-step draft-then-send), the DEPLOYED Worker doing all
  classification off its real webhook. Threshold lowered to 0.5 for the run (the v9 finding: the
  model is rightly warier of [MCP TEST]-marked subjects). The correction-target folder is
  deliberately named nothing like a receipt folder, with a fallback B folder if the model ever picks
  it. Self-sent probes mean the learned preference keys on the owner's own address — which is why
  the e2e restores `llm:prefs` byte-exact in a `finally` and r20's sweep additionally fails if any
  preference still points at a [MCP TEST] folder (removing it) or any [MCP TEST] folder survives.
- v12a (search filters, live) polls `$search` for its probe rather than asserting immediately —
  search indexing lags delivery by seconds-to-minutes; the `$filter` assertions need no poll.
- **The live filer really does race the local harness — fixed with a shield, not a retry.** With
  auto-filing genuinely enabled on the deployed Worker (the owner's real state, threshold 0.8), a
  local-suite send-to-self probe ("[MCP TEST] v2") was classified at 0.85 and filed out of the inbox
  seconds after delivery, timing out the harness's inbox poll (tests b/b2/c, and v5a's delta probe
  the same way). Earlier runs got lucky (the model chose no action on test-marked mail); this run it
  didn't. Fix: every send-to-self probe in the LOCAL suite now carries `FILER_SHIELD` — the phrase
  "verification code probe", which matches the auto-filer's compiled-in protected-subject list — so
  the deployed filer skips them before any model call, deterministically. Test b asserts the phrase
  still matches `isProtectedSubject`, so the shield cannot silently rot. Remote probes (r25, r29)
  deliberately do NOT carry it: being classified is their point.
- Deployed-KV hygiene beyond r20: the local suite's probes transit the real inbox, so the DEPLOYED
  audit/activity rings can pick up [MCP TEST] entries during a local run (and the classifier's
  cleanup races can tick `err:filing:<date>`); after the final local run those were swept from KV
  and `llm:config` re-verified byte-identical to the captured pre-run value.
