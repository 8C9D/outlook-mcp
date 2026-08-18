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
