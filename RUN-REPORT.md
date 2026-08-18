# RUN-REPORT — outlook-mcp orchestrated run, Batches 2–4

Orchestrated run of 2026-08-18. Each batch: front-loaded auth by the orchestrator,
implementation by a fresh subagent, gate review re-run by the orchestrator.

## Batch 2 — Local feature completion (v0.4.0) — PASSED

**Front-loaded auth (orchestrator).** Added delegated **Tasks.ReadWrite** to the
outlook-mcp app registration in the Azure portal (Microsoft Graph now shows 9
permissions: the previous eight + Tasks.ReadWrite). Added the scope to
`src/auth.ts`, deleted `.token-cache.json`, ran `npm run login`; the user
completed an email-verification challenge mid-flow (expected human step) and
approved the consent screen. Silent re-acquisition verified (`npm run verify`
all green) and a live `/me/todo/lists` call confirmed the new scope works.
Commit `85b4a3d`.

**Shipped (subagent A).**
- 4 new tools (17 → 21): `create_folder` (duplicate-reject naming the existing
  folder id), `manage_categories` (list/create/delete, preset0–24 palette),
  `list_tasks` (grouped overdue/today/upcoming/no-date, Toronto), `manage_task`
  (create/complete/reopen/update/delete, permanent-delete caveat documented;
  `linked_message_id` turns an email into a task with subject/sender/webLink in
  the notes).
- `manage_rules`: new `update` action (PATCH in place, id and sequence
  preserved), `exceptions` on create/update, `enabled` toggle, exceptions in the
  list summary. Forwarding actions still excluded; foreign forward rules still
  flagged.
- `manage_message`: new `categorize` action, replace-not-append, categories
  validated against the master list.
- 2 MCP prompts registered and served over stdio: `triage_inbox` (propose-only,
  approval-gated) and `morning_brief`.
- Notable live-Graph findings fixed during the batch: Graph returns
  400 `ErrorInvalidIdMalformed` (not 404) for malformed ids; To Do rejects
  `$select` on a single task; To Do normalizes due dates to UTC on write
  (reads now send `Prefer: outlook.timezone`).

**Gate review (orchestrator, re-run independently).**
- Live harness: **23/23 PASS** (was 19/19), final sweep zero `[MCP TEST]`
  artifacts across mail, folders, categories, tasks, rules.
- `npx tsc --noEmit` clean; `npm run build` OK.
- `dist/server.js` from `cwd=/tmp`: tools/list returns **21 tools**,
  prompts/list returns **triage_inbox, morning_brief**, stdout clean JSON-RPC.
- Greps: `/me/sendMail` appears only in prohibition comments; forward/redirect
  identifiers only in the read-side foreign-rule flagging.
- README (21-tool table, To Do notes, prompts section, permanent-delete caveat
  in the security model) and ASSUMPTIONS "v4 batch 2" section present;
  version **0.4.0**.

**Commits.** `85b4a3d` (auth), `140a1be` (tool surface + prompts + 0.4.0),
`59d996a` (harness), `a12da2d` (docs).

**Assumptions of consequence.** No `delete_folder` tool (Graph folder delete
takes the message subtree; deserves its own guarded batch). `due_within_days`
keeps overdue and drops no-date tasks; 0 = due today or earlier. `due_date` /
`reminder` can be set but not cleared via update (sentinel deferred). Prompts
are zero-argument. `search_mail` listings carry no read/unread flag —
`morning_brief` instructs a `read_message` follow-up; an unread filter is a
candidate for a future batch.

**Deviations.** Subagent A was twice interrupted by transient API 529 overload
errors and resumed; no impact on the result.

## Batch 3 — Remote hosting on Cloudflare Workers (v0.5.0) — PASSED

**Front-loaded auth (orchestrator).** Cloudflare account confirmed existing and
logged in at dash.cloudflare.com (account id `cf54199aef7e9425e000486ce1fda8bb`,
free tier). `npx wrangler login`: the Cloudflare consent page ignored
automated clicks on Authorize (trusted-gesture requirement) and two OAuth
attempts timed out; the user ran `npx -y wrangler login` themselves and
authorized — the expected human-dependency pause, not a gate violation.
`wrangler whoami` verified (workers / KV / secrets write scopes).

**Shipped (subagent B).**
- Transport-agnostic core refactor: registry, Graph logic, and prompts moved to
  `src/core/*`; tools decoupled from MSAL via a pluggable token provider
  (AsyncLocalStorage per-request on Workers, process default on Node). Both
  entrypoints build the identical `McpServer` from one registry.
- Cloudflare Worker entry (`src/worker/index.ts`) serving MCP over Streamable
  HTTP, deployed at
  **https://outlook-mcp.arthur-yuhao-zhang.workers.dev/mcp**. SDK choice: the
  MCP SDK's `WebStandardStreamableHTTPServerTransport` (stateless, no Durable
  Objects) rather than Cloudflare's deprecated `McpAgent` — keeps one SDK for
  both transports (rationale in ASSUMPTIONS).
- Remote OAuth via `@cloudflare/workers-oauth-provider` (RFC 8414/9728
  metadata, RFC 7591 DCR, S256 PKCE). Identity proved by Microsoft device-code
  flow; a single-user allowlist admits only the owner's Graph `/me` id or
  UPN/mail. No anonymous route reaches Graph.
- Microsoft tokens in Workers KV for remote mode: `npm run seed:kv` pushes the
  local refresh token; the Worker refreshes directly against the consumers
  token endpoint and rotates the stored refresh token on every exchange. Local
  stdio mode keeps MSAL and the disk cache untouched; the two credential
  chains are independent. Secrets live in wrangler secrets only.
- New remote suite `npm run test:remote` (14 tests, self-cleaning: OAuth
  records deleted, revoked bearer proven dead).

**Gate review (orchestrator, re-run independently).**
- Local harness **23/23**; remote suite **14/14** (discovery, anonymous/bogus
  bearer refusal, DCR, allowlist refusal, full PKCE authorize+exchange,
  bad-verifier rejection, initialize, 21-tool parity with stdio, prompts,
  live `list_events` through the KV token, refresh-rotation proven, cleanup,
  revoked-bearer sweep).
- `npx tsc --noEmit` and the worker tsconfig both clean; dist rebuilt; stdio
  smoke from `cwd=/tmp`: server `outlook 0.5.0`, 21 tools, 2 prompts, stdout
  is 3 lines of valid JSON-RPC only.
- Secret hygiene verified: `wrangler.jsonc` holds only KV namespace ids;
  `.gitignore` gained `.dev.vars` and `.wrangler/`; greps found no tokens.
- README "Remote deployment" section (architecture, seeding, exact claude.ai
  connector steps, rotate/revoke) and ASSUMPTIONS batch 3 sections present;
  version **0.5.0**; tree clean.

**Commits.** `0c1fc98` (core refactor), `908f2ba` (remote mode + deploy + docs).

**Assumptions / deviations of consequence.**
- `/authorize` additionally accepts POST with an `ms_access_token` form field
  so the remote harness can complete a real OAuth exchange non-interactively.
  It runs the identical Graph `/me` allowlist check; only a live Microsoft
  access token for the allowlisted account works (which already implies
  mailbox access). Flagged for the user's judgment; documented in ASSUMPTIONS.
- `resourceMetadata.resource` is a hardcoded literal (env not available at
  module scope); renaming the Worker requires editing that line.
- Two tsconfigs (Node vs workerd globals conflict); `npm run typecheck` runs
  both. Generated `worker-configuration.d.ts` committed (binding names only).
- KV eventual consistency handled with bounded polling in two remote tests.
- The claude.ai connector itself is NOT yet added — that is an interactive
  user step (documented in the README).

## Batch 4 — Remote-native features (v0.6.0) — PASSED

**Front-loaded auth (orchestrator).** None needed beyond confirming wrangler
was still authenticated (`wrangler whoami` OK).

**Shipped (subagent C).**
- `check_new_mail` (22nd tool): Graph delta queries per folder. First call (or
  `reset: true`) records a position and lists nothing; later calls return only
  added/changed/removed messages exactly once. Position stored in Workers KV
  remotely and a 0600 `.mcp-state.json` locally (gitignored). Works on both
  transports.
- Graph change notifications: public Worker `/notifications` route handles the
  validation handshake and clientState-checked deliveries into a capped
  (50-entry) KV ring buffer; always answers 202 (no secret-guessing oracle,
  no Graph retry storms). A Workers cron (`17 */6 * * *`) renews the inbox
  subscription before Graph's ~4230-minute cap and re-creates it if lapsed; a
  request-time backstop (`ctx.waitUntil`) also runs the renewal check so the
  subscription bootstraps after deploy and lapses cannot linger.
- `get_mailbox_activity` (23rd tool): reads the ring buffer — pushed activity
  without polling Graph. Remote-only; on stdio it returns a clear isError
  pointing at `check_new_mail`.
- MCP resources on BOTH transports: `outlook://mail/folders` and
  `outlook://mail/inbox/recent` (no SDK limitation encountered).
- Notable findings: `$deltatoken=latest` is OneDrive-only; carrying
  `Prefer: odata.maxpagesize` on every delta hop (Graph drops it from
  nextLink) cut the 1006-message inbox baseline from 92 requests/10.8 s to
  3 requests/0.85 s; consumer accounts do support message subscriptions
  (proven live before wiring).

**Gate review (orchestrator, re-run independently).**
- Local harness **28/28** (delta lifecycle, handshake, ingest, cron renewal
  handler, remote-only guard, resources); remote suite **20/20** including the
  live webhook end-to-end (send-to-self → notification lands in
  `get_mailbox_activity`), KV delta persistence, and the subscription check —
  subscription `95f07422-d465-455d-b03f-14d098692d93` on
  `/me/mailFolders('inbox')/messages`, expiry 2026-08-21T17:31:56Z (~70 h
  out); cron `17 */6 * * *` in the committed wrangler.jsonc and deployed.
- `npm run typecheck` (both tsconfigs) clean; dist rebuilt; stdio smoke from
  `cwd=/tmp`: `outlook 0.6.0`, **23 tools / 2 prompts / 2 resources**, stdout
  entirely valid JSON-RPC.
- Sweeps clean: zero `[MCP TEST]` artifacts locally; KV holds only the
  refresh/access tokens, `sub:mail`, and an empty ring buffer. No secrets in
  the tree (clientState lives only in KV, rotates with the subscription);
  `.mcp-state.json` gitignored; version **0.6.0**; tree clean.

**Commits.** `059f4cd` (surface), `f4cbabc` (local harness), `81577f1`
(remote suite), `71ff882` (docs).

**Assumptions / deviations of consequence.** Request-time subscription
backstop added beyond the brief (rationale above). Ring-buffer writes are
read-modify-write without a lock — fine at single-user scale; a Durable
Object would fix it at scale but was deliberately avoided in Batch 3.

## v6.1 — Production hardening: direct authorize path disabled — PASSED

**Shipped.**
- `POST /authorize` with a caller-supplied `ms_access_token` (the path the
  remote harness used to mint bearers non-interactively) is now gated on
  `ALLOW_DIRECT_AUTHORIZE === "true"` and refused with a `403` "disabled"
  answer **before the request is parsed**. The flag is deliberately absent
  from the deployed Worker (not a var, not a secret — `wrangler secret list`
  re-verified: only the original three), so production authorizes exclusively
  through the interactive device-code flow. Local `wrangler dev` enables the
  path via the gitignored `.dev.vars` (flag + mirrors of the three secrets).
- Remote suite reworked to match: **r5** now asserts the deployed endpoint
  refuses the direct path (403 + /disabled/i, proving the gate fires before
  token validation); **r6** drives the real production `/authorize` page —
  parses the device code and flow id out of the HTML, asks the human running
  the suite to sign in at microsoft.com/devicelogin, polls `/authorize/poll`.
  Headless runs (no TTY, or `MCP_REMOTE_HEADLESS=1`) report the ten
  bearer-dependent tests as SKIP; `MCP_REMOTE_INTERACTIVE=1` forces the
  interactive path.
- Docs: README (security-model bullet, allowlist section, `test:remote`
  description, setup snippet warning) and ASSUMPTIONS.md ("v6.1: production
  hardening" section; the old "deviation worth flagging" bullet marked
  superseded).

**Gate review.**
- `npm run typecheck` (both tsconfigs) clean.
- Gate verified both ways against `wrangler dev --local` before deploying:
  with the flag, DCR + a real MSAL token completed a full direct
  authorization (`status: "ok"` with a code); without `.dev.vars`, the same
  POST got the 403 "disabled" answer.
- Deployed (version `a3417b80`); live probe of the deployed `/authorize`
  confirmed the 403; remote suite **20/20** headless (10 passed live, 10
  SKIP pending an interactive sign-in), including the new r5.
- Webhook state re-verified live after the deploy: subscription
  `95f07422-d465-455d-b03f-14d098692d93` on
  `/me/mailFolders('inbox')/messages`, notification URL correct, expiry
  2026-08-21T17:31:56Z (future), no `renewedAt` yet (fresh from the Batch 4
  run); cron `17 */6 * * *` present in the deploy output's trigger list.

**Assumptions / deviations of consequence.**
- Disabling the direct path removes headless bearer acquisition *by design*;
  full authenticated remote coverage now costs one interactive device-code
  sign-in per run. Judged the right trade: the path let anonymous callers
  relay arbitrary tokens to Graph and could turn a leaked short-lived
  Microsoft access token into a persistent grant.
- Found while re-verifying the webhook: Graph held **four** subscriptions,
  all created within one second (2026-08-18T19:31:56Z, during the Batch 4
  remote run) — the request-time upkeep backstop raced itself across
  concurrent authenticated requests while KV reads were still stale, and
  each created a subscription; the last KV write won. Harmless in effect
  (the three orphans' deliveries fail the clientState check and are
  discarded, and they expire 2026-08-21 with nothing renewing them), but it
  is a real unlocked-check-then-act race, kin to the ring-buffer caveat in
  Batch 4. Orphan deletion was prepared but not executed (tool-permission
  gate); left for the user to run or to let them lapse.

## v6.1b — Upkeep race fixed: Graph as source of truth — PASSED

**Cleanup (user-approved).** The three orphan subscriptions from the Batch 4
race were deleted live via `DELETE /subscriptions/{id}`, keeping
`95f07422-d465-455d-b03f-14d098692d93` — the one whose clientState the
`sub:mail` KV record holds.

**Shipped.**
- `ensureMailSubscription` reworked: whenever the KV record alone cannot
  justify "keep", Graph is the source of truth and KV only a cache. Upkeep
  lists the unexpired subscriptions for this notification URL + resource,
  renews the one whose id the KV record names while sweeping every other,
  and creates only when the record's subscription is not live in Graph —
  after deleting whatever else is there. A stale KV read therefore converges
  on exactly one subscription instead of minting duplicates. The
  per-request fast path is unchanged (healthy record = one KV read, zero
  Graph calls).
- Constraint discovered live and honored: Graph returns `clientState: null`
  when listing, so a racing peer's subscription can never be *reused* — its
  deliveries could never be validated — and is replaced instead. "Reuse the
  existing one" is implemented as "reuse iff KV holds its clientState".
  Residual window: two upkeeps that both list before either creates (true
  bootstrap only) make one each; the next renew-path upkeep sweeps the loser.
- New harness test **v5f**: simulates the incident (second ensure runs on a
  stale, empty KV view after the first created) and asserts exactly one
  subscription survives and it is the one the last KV write describes; also
  drives the simultaneous-bootstrap case via `Promise.all` and asserts the
  follow-up renewal sweeps back down to one. v5d updated for the new
  list-first traffic (create = GET,POST; renew = GET+PATCH; Graph-forgot =
  GET,POST with no PATCH; list/PATCH race covered separately).

**Gate review.**
- `npm run typecheck` clean; local harness **29/29** (was 28) including v5f.
- Deployed (version `4ba94b3c`); Graph re-listed after the deploy: exactly
  **one** subscription — `95f07422…`, correct notification URL, expiry
  2026-08-21T17:31:56Z; cron `17 */6 * * *` in the deploy trigger list.
- Remote suite headless: **20/20** (10 live incl. r5/r13/r16, 10 SKIP
  pending an interactive sign-in).

## Final state

- **Version 0.6.0** — 23 tools, 2 prompts, 2 resources, two transports.
- Local: `node dist/server.js` (Claude Desktop), MSAL + disk cache.
- Remote: **https://outlook-mcp.arthur-yuhao-zhang.workers.dev/mcp** —
  OAuth-gated (single-user allowlist, interactive device-code sign-in only in
  production as of v6.1), tokens in Workers KV with refresh rotation, live
  inbox webhook subscription with cron renewal.
- All four gates passed with the orchestrator re-running every check.
- Run totals: local harness 19/19 → 28/28; remote suite 0 → 20/20.

### The user's remaining manual steps

1. **Cmd+Q and reopen Claude Desktop** so it picks up the new local tools
   (23 tools / 2 prompts / 2 resources over stdio).
2. **Add the remote server as a claude.ai custom connector** following
   README → "Adding it to claude.ai as a custom connector": Settings →
   Connectors → Add custom connector →
   `https://outlook-mcp.arthur-yuhao-zhang.workers.dev/mcp`, then complete
   the Microsoft device-code sign-in when the authorize page asks (only the
   allowlisted account can finish it).
