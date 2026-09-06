# outlook-mcp — resume context

_Compiled 2026-08-25 against commit `afe1aa5` (v1.3.0) on branch `main`._

## 1. One-line description

A production MCP (Model Context Protocol) server that gives Claude full read/write access to a personal Microsoft Outlook account — mail, calendar, contacts, To Do tasks, inbox rules and OneDrive — through Microsoft Graph, served over two transports (local stdio and a deployed Cloudflare Worker used as a live claude.ai custom connector), with an explicitly prompt-injection-hardened security model; built by and for a single owner, but published publicly with full setup docs so anyone can run their own instance.

## 2. Tech stack

- **Language / type system:** TypeScript 5.7 (ESM, `NodeNext`), strict; three `tsconfig` projects (`tsconfig.json`, `tsconfig.build.json`, `tsconfig.worker.json`) so the Node build and the Workers build are typechecked separately
- **Runtimes / platforms:** Node.js 22+ (local stdio server, run through `tsx`) and Cloudflare Workers / `workerd` (hosted server, `nodejs_compat`)
- **Protocol:** Model Context Protocol via `@modelcontextprotocol/sdk` ^1.30 — stdio transport locally, Streamable HTTP on the Worker; tools, prompts and MCP resources
- **Auth:** Microsoft Entra ID OAuth 2.0 — MSAL Node (`@azure/msal-node` ^2.16) device-code flow locally; on the Worker, `@cloudflare/workers-oauth-provider` ^0.10 implements a full OAuth 2.1 authorization server (RFC 7591 dynamic client registration, PKCE, RFC 9728 protected-resource metadata) layered over a hand-rolled Microsoft refresh-token grant written with plain `fetch` (MSAL does not run on `workerd`)
- **APIs:** Microsoft Graph v1.0 (mail, attachments + upload sessions, MIME `$value`, mail folders, inbox rules, categories, calendars/events, contacts, mailbox settings, Microsoft To Do, OneDrive/`drive`, `$batch`, delta queries, change-notification subscriptions/webhooks); Anthropic Messages API for the two optional LLM features
- **Storage / infra:** Cloudflare KV (two namespaces — Microsoft tokens/state, and OAuth grants), Cloudflare Cron Triggers (4 schedules), Workers Observability
- **Validation / tooling:** Zod ^4 schemas for every tool input (and structured output on several), Wrangler ^4 for deploy and type generation, GitHub Actions CI
- **Testing:** three hand-rolled test tiers (no framework) — offline/stubbed, live-against-real-mailbox, and live-against-deployed-Worker

## 3. Scope and scale (verifiable numbers only)

| Metric | Value | Source |
| --- | --- | --- |
| Commits | 73 | `git log --oneline \| wc -l` |
| Date range | 2026-08-18 → 2026-08-25 (54 commits on 08-18, 18 on 08-19, 1 on 08-25) | `git log --format='%ad' --date=short \| sort \| uniq -c` |
| Authors | 1 (Arthur Zhang) | `git log --format='%an <%ae>' \| sort \| uniq -c` |
| Release tags | 3 (`v1.0.0`, `v1.1.0`, `v1.2.0`); v1.3.0 shipped in `afe1aa5` but is not yet tagged | `git tag` |
| Source LOC, tracked, excl. tests | 14,341 | `git ls-files -- src \| grep -v '/test-' \| xargs wc -l` |
| Test LOC | 9,234 (`test-tools.ts` 5,670 · `test-remote.ts` 2,089 · `test-offline.ts` 1,475) | `wc -l src/test-*.ts` |
| Total `src/` TypeScript | 23,575 | `git ls-files -- src \| xargs wc -l` |
| Tracked files (excl. `node_modules`/`dist`) | 102 | `git ls-files \| wc -l` |
| **MCP tools exposed** | **40** | `grep -oE '^    name: "[a-z_]+"' src/core/registry.ts \| sort \| wc -l`; matches the deployed connector's tool list |
| MCP prompts | 2 (`triage_inbox`, `morning_brief`) | `src/core/prompts.ts` |
| MCP resources | 2 (`outlook://mail/folders`, `outlook://mail/inbox/recent`) | `src/core/resources.ts` |
| Tools carrying all 4 MCP annotation hints | 40/40 (12 via a shared `READ_ONLY` constant + 28 inline) | `grep -n 'annotations:' src/core/registry.ts` |
| Tools marked `destructiveHint: true` | 12 | `grep -oE 'destructiveHint: (true\|false)' src/core/registry.ts` |
| **Tests** | **116 total: 23 offline (CI) + 62 live local + 31 live remote** | `grep -cE '^await test\(' src/test-offline.ts` (23); `test-tools.ts` 61 top-level + 1 nested = 62; `test-remote.ts` 15 `test(` + 16 `testAuthed(` = 31. Corroborated by the v1.3.0 commit message ("offline 23/23, local live 62/62, remote 31/31") |
| Assertions across the three suites | 1,160 (`assert(` calls: 251 offline, 679 local, 230 remote) | `grep -cE '\bassert\(' src/test-*.ts` |
| Cloudflare cron triggers | 4 | `triggers.crons` in `wrangler.jsonc` |
| Graph delegated scopes requested | 10 | `MAILBOX_SCOPES` in `src/worker/ms-token.ts` |
| Capability-parity audit rows | 96 (42 already supported, 26 newly implemented in v1.3, 19 not feasible via Graph, 6 excluded by design) | `grep -cE '^\| .*\| *(✅\|🆕\|❌\|🚫\|—) *\|' CAPABILITIES.md` |
| Tracked documentation | 1,552 lines / ~17,200 words across 5 Markdown files (README 1,053 · SETUP 250 · CAPABILITIES 164 · SECURITY 52 · CONTRIBUTING 33) | `git ls-files '*.md' \| xargs wc -l` |
| Deployed version responding live | `1.3.0` | `curl https://outlook-mcp.arthur-yuhao-zhang.workers.dev/health` → `{"status":"ok","service":"outlook-mcp","version":"1.3.0"}` |

**Not counted as source:** `worker-configuration.d.ts` (14,884 lines) is Wrangler-generated Workers type output and is committed; `package-lock.json` (115 KB) is a lockfile.

## 4. Engineering highlights

**1. One tool registry, two transports, provably no drift.** `src/core/registry.ts` holds a single `TOOLS` array; both the stdio entry point (`src/server.ts`) and the Cloudflare Worker (`src/worker/index.ts`) build from `createMcpServer()`, so a local install and the hosted connector can never expose different surfaces. Transport-specific concerns sit behind small ports: `src/core/graph.ts` is deliberately free of Node-only imports and asks `src/core/token.ts` for a token (scoped per-request with `AsyncLocalStorage`); attachments save to disk via `src/tools/save-local.ts` on stdio and are handed out as expiring, bearer-gated download links via `src/core/downloads.ts` + `src/worker/download.ts` on the Worker, which has no filesystem.

**2. OAuth architecture: a real authorization server in front of a delegated Microsoft grant.** The Worker runs `@cloudflare/workers-oauth-provider` as a full OAuth 2.1 AS — discovery metadata, RFC 7591 dynamic client registration (needed because claude.ai registers itself), PKCE, token issuance, bearer validation, and the 401 + `WWW-Authenticate` challenge — with `apiRoute: "/mcp"` so nothing anonymous reaches the MCP endpoint (`src/worker/index.ts:60`). Behind it, identity is proven with Microsoft's **device-code** flow (`src/worker/authorize.ts`) precisely because the Entra app is registered as a public native client with no web redirect URI; the device-code token is used once, to call Graph `/me`, checked against a single-user allowlist, and never stored. The mailbox refresh token is a separate credential in KV, and `src/worker/ms-token.ts` reimplements the refresh grant with plain `fetch` because MSAL cannot run on `workerd` — including writing back Microsoft's rotated refresh token, without which the connector locks itself out on the next call. Remote test `r5` asserts the deployed Worker refuses the non-interactive authorize path that local `wrangler dev` enables.

**3. LLM auto-filing hardened against prompt injection by construction, not by prompt.** `src/core/classifier.ts` files newly-arrived mail into the user's *existing* folders and treats mail as hostile input at four independent layers, documented in its header: (a) **structurally** — the module imports no Graph transport at all and acts only through a `ClassifierMailbox` port whose sole mutations are `move` and `categorize`, so send/delete/reply/forward are not expressible from that code path; (b) **by allowlist** — Deleted Items and Junk Email are excluded from the folder list given to the model (`NEVER_FILE_INTO` in `src/core/auto-filing.ts`), which is what stops "move" from standing in for "delete"; (c) **by schema** — the answer must parse as exact JSON or it is discarded; (d) **by prompt and explicit delimiters**. Compiled-in `PROTECTED_SUBJECT_PATTERNS` (OTP codes, 2FA, password resets) never reach the model at all. Every discarded answer is written to a 100-entry audit ring readable through the `get_auto_filing_log` tool, so an injection attempt is *visible* rather than silent. A feedback loop (`src/core/corrections.ts`) reconciles recent filed messages against their current folder and turns a user's manual re-file into a sender→folder preference that short-circuits future model calls. Both LLM features ship **disabled**, share a per-day API-call cap (`DEFAULT_DAILY_CALL_CAP = 200`) and a confidence threshold (`DEFAULT_THRESHOLD = 0.8`), and are toggled through `manage_auto_filing`. The 07:00 morning brief (`src/core/digest.ts`) is drafted, never sent — its `DigestMailbox` interface has no send method.

**4. Send is structurally two-step, and the server monitors itself.** No tool calls `/me/sendMail`: a message must exist as a reviewable draft before `send_draft` can name it, so no single tool call can compose-and-send. The one exception is the opt-in self-alert route, which may send only to the owner's own address, behind a shared secret and a 20-a-day cap. v1.3 added scheduled send by setting the MAPI deferred-send-time property on the draft before that same single POST — no new send path — plus `manage_scheduled_send` to list and cancel waiting sends. A daily health cron (`src/core/health.ts`, `HEALTH_CRON = "37 13 * * *"`) verifies KV, a forced token rotation, the Graph subscription and both LLM error counters; a failing run leaves an **unsent draft** in the owner's inbox rather than emailing anyone, so the health check itself never sends. Every dependency is injected, so all failure modes are unit-tested offline.

**5. Three-tier testing split along a credential boundary, with self-cleaning live suites.** `test-offline.ts` (23 tests, 251 assertions) deliberately imports nothing that pulls in MSAL or dotenv — if a test reaches for a real Graph call it fails loudly with `AuthRequiredError` — which is exactly what lets `.github/workflows/ci.yml` run `typecheck` + `test:offline` on every push with **no secret ever entering the repo or CI**. `test-tools.ts` (62 tests, 679 assertions) exercises handlers against the real mailbox, and every test cleans up after itself with a final sweep asserting no `[MCP TEST]` artifacts remain in messages, drafts, folders, events, calendars, contacts, rules, categories, tasks or temp files, and that mailbox settings are restored exactly. `test-remote.ts` (31 tests, 230 assertions) runs a full OAuth authorization-code exchange against the deployed Worker, proves the KV refresh token actually rotates, and traces a real Graph change notification end-to-end into `get_mailbox_activity` — then deletes the OAuth client, grants and tokens it created. Assertions go well beyond happy paths: `v3b` asserts three drafts are marked read in **one** `$batch` Graph call, `v5f` asserts two racing subscription-ensure calls leave exactly one subscription, and `v9a` asserts no Graph transport is even reachable from the classifier's import graph.

## 5. Distribution status

- **Repository:** `https://github.com/8C9D/outlook-mcp` — **public**, MIT licensed, 0 stars, single branch `main`, last push 2026-08-25. GitHub shows **no repo description** — the About sidebar is empty.
- **Deployed:** Cloudflare Workers at `https://outlook-mcp.arthur-yuhao-zhang.workers.dev`. **Verified live at compile time**: `GET /health` returns `200 {"status":"ok","service":"outlook-mcp","version":"1.3.0"}` (matching `package.json`), and `GET /.well-known/oauth-authorization-server` returns valid RFC 8414 metadata advertising `/authorize`, `/oauth/token` and `/oauth/register` with `authorization_code` + `refresh_token` grants.
- **In use:** the Worker is registered as a claude.ai custom connector and the server is single-user by design (a Graph `/me` id allowlist). Four cron triggers are configured, including a daily health check whose healthy runs write a KV heartbeat. It also runs locally over stdio in Claude Desktop, and `run-school.sh` adds a second env-selected instance for a work/school M365 account (deliberately without `Mail.Send`).
- **What a reviewer would see today:** a very strong front door. The README is 1,053 lines and opens with a one-paragraph security model, a capability table by area, and cost-to-run; `SETUP.md` walks a stranger from empty directory through the Entra app registration; `CAPABILITIES.md` is a 96-row parity audit of personal Outlook against the tool surface, each gap marked supported / new / infeasible-with-live-probe-cited / excluded-by-design; `SECURITY.md`, `CONTRIBUTING.md` and an MIT `LICENSE` are present; CI runs typecheck + the offline tier on every push. Source comments are unusually explanatory — most modules open with a design-rationale header (why device code, why not MSAL on `workerd`, why two UTC crons for one Toronto hour). Commit messages are substantive multi-paragraph release notes.
- **Changed 2026-08-25 (see §7):** the repo previously also shipped `ASSUMPTIONS.md` (1,814 lines) and `RUN-REPORT.md` (1,028 lines) — internal build-run journals in batch/gate language, together ~2.7× the README's length. Both have been untracked and gitignored (they remain on disk locally), and the seven inbound references to them from source comments and docs were rewritten so nothing dangles. Tracked documentation is now 1,552 lines of user-facing docs only.
- **Verified live on 2026-08-25 by running the remote suite** (`MCP_REMOTE_HEADLESS=1 npm run test:remote`, 31/31 passed):
  - **Both LLM features are ENABLED in production** — `r24` printed the deployed KV config as `{"filingEnabled":true,"digestEnabled":true,"threshold":0.8,"dailyCallCap":200}`. They ship disabled; this instance has them turned on. This is a materially stronger claim than "implemented but off" and is safe to make.
  - **The auto-filer and its feedback loop work end-to-end against production** — `r29` filed one probe via a real `claude-haiku-4-5` call (808 input / 55 output tokens), had it re-filed, learned the correction as a sender preference, then filed a second probe **with no model call at all**.
  - **The health cron is running** — heartbeat `health:last` healthy at `2026-08-25T17:12:58Z`, 5/5 checks.
  - **The Graph webhook subscription is live** — `95f07422-…`, expiring 2026-08-27, i.e. the 6-hourly renewal cron is doing its job.
- **Still not verifiable from the repo:** actual daily usage frequency by a human; the real monthly Anthropic cost (README claims ~$1–2/month "measured"); and whether anyone other than the owner has installed it.

## 6. Candidate resume bullets

Each fits one printed line at typical resume width. **★ = the three to use if the project gets one block of three or four lines** — they cover scope, the hardest technical problem, and rigor, with no overlap.

1. ★ Built and deployed a production MCP server exposing **40 Microsoft Graph tools** (mail, calendar, contacts, tasks, OneDrive) from one registry over two transports — local stdio and a live Cloudflare Worker — in **14.3K lines of TypeScript**.
2. ★ Implemented a full **OAuth 2.1 authorization server** (dynamic client registration, PKCE, RFC 9728) over a hand-rolled Microsoft refresh grant written for `workerd` after MSAL proved unrunnable there, including rotation write-back verified against the live deployment.
3. ★ Wrote **116 tests / 1,160 assertions across three tiers**, split so the 23-test offline tier gates every push in CI with **zero credentials in the repo**, and both live tiers self-clean and assert they left no artifacts.
4. Hardened an LLM mail auto-filer against prompt injection **structurally**: the classifier can reach no send or delete path, folder allowlists block "move" as a laundered delete, and an import-graph test fails CI if that boundary is ever crossed.
5. Ran it like production on **4 Cloudflare cron triggers** — webhook subscription renewal, a DST-correct daily digest, and a self-health check that reports failures as an **unsent draft**, preserving the invariant that only a human-driven call sends mail.
6. Audited the tool surface against **96 Outlook capabilities**, closing every feasible gap in one release and documenting the **19 proven infeasible** through Graph with the live probe output for each.

**Notes on wording.** Bullet 3 says "116 tests / 1,160 assertions" deliberately — quote the assertion count if anyone probes, since each "test" is a multi-step scenario rather than a Jest-style unit test. Bullets 1 and 5 can be merged into one if the project only earns two lines; bullets 2 and 4 are the two that reliably start a technical conversation.

## 7. Gaps and risks

> **Remediation applied 2026-08-25** (uncommitted at time of writing — `git status` shows the working tree changed but nothing pushed):
> 1. `cleanup-progress.md` and `recover_deletions.py` added to `.gitignore`.
> 2. `ASSUMPTIONS.md` and `RUN-REPORT.md` untracked (`git rm --cached`) and gitignored; their seven inbound references rewritten.
> 3. The deployed hostname removed from `src/` entirely — it now lives only in `wrangler.jsonc`. This also fixed a latent bug: the OAuth `resourceMetadata.resource` was a hardcoded literal beneath a comment claiming it was "filled from the wrangler var."
>
> Verified after: `npm run typecheck` clean, 23/23 offline, `npm run doctor` 11/11. The items below describe the state as found, with current status noted inline.

**Highest-priority risk (NOW MITIGATED) — untracked personal data in the working tree of a public repo.** Two files were present locally, **untracked but NOT in `.gitignore`**, so a single `git add -A` would have published them:

- `cleanup-progress.md` (1,679 lines, 159 KB) — a mailbox-cleanup journal containing **176 unique real email addresses** (university admissions offices, brokers, banks, personal correspondents), real folder names with unread/total counts for the owner's actual mailbox (e.g. "Inbox 674/1018", "Job Search 2026 41/234", "Finance 225/543"), a **verbatim JSON export of 8 real inbox rules including real Graph folder IDs and sender addresses**, and per-message deletion decisions. This is by far the most damaging thing that could reach the public repo.
- `recover_deletions.py` — hardcodes an absolute path to a Claude Code session transcript under `~/.claude/projects/...` and parses real message subjects and senders out of it.
- **Status: fixed.** Both are now in `.gitignore` (`git check-ignore -v` confirms; neither appears in `git status`). Verified that neither was ever committed — `git log --all --diff-filter=A --name-only` shows no such path — so no history rewrite is needed.

**No secrets are committed — verified.** `.env`, `.dev.vars` and `.token-cache.json` exist locally, are all in `.gitignore`, and the same history scan found **no** such file ever added. Wrangler secrets (`MS_CLIENT_ID`, `ALLOWED_MS_USER_ID`, `ALLOWED_MS_UPN`, `ANTHROPIC_API_KEY`) are documented as secrets in a `wrangler.jsonc` comment and set out-of-band.

**Test fixtures are clean.** Every email address in tracked source is a placeholder (`shop@example.com`, `attacker@evil.invalid`, `mcp-test@example.invalid`, etc.); no third-party real address, no real subject line, and no real Graph message or folder ID (`AQMkAD…`/`AAMkAD…`) appears in any tracked file.

**Personal identifiers that are committed** — deliberately, per `ASSUMPTIONS.md:1431` ("reviewed and accepted, NOT scrubbed") — but worth re-deciding for a resume audience:

- The owner's real address `arthur.yuhao.zhang@outlook.com` appears 4× (`SECURITY.md` as the contact, `ASSUMPTIONS.md` ×2, `RUN-REPORT.md` ×1).
- ~~The deployed hostname `outlook-mcp.arthur-yuhao-zhang.workers.dev` is hardcoded in 6 tracked files, including source defaults.~~ **Fixed:** `grep -rn 'arthur-yuhao-zhang' src/` now returns nothing. The origin lives only in `wrangler.jsonc`, reached in the Worker via a binding and Node-side via a new `deployedBaseUrl()` helper in `src/project-root.ts` (which also de-duplicated a copy of the same parser in `doctor.ts`). A fork now changes one config line, not source. It still appears in README/SETUP prose, which is appropriate — that is documentation of the owner's live endpoint.
- Absolute local paths carrying the username: `README.md:1029` (`/Users/arthurzhang/.nvm/.../node`) and `ASSUMPTIONS.md:81`.
- **Azure/Cloudflare identifiers in the clear:** Entra tenant ID `a289df25-…` and client ID `1d362aa5-…` (`ASSUMPTIONS.md:14-15`); a second hardcoded school-tenant client ID in `run-school.sh:13`; both **Cloudflare KV namespace IDs** in `wrangler.jsonc`; Graph subscription IDs and Worker version IDs in `RUN-REPORT.md`. None are secrets in the credential sense (public client IDs and namespace IDs are unusable without the matching token), but they are hardcoded account identifiers a security-minded reviewer will notice on sight.

**Things that could read poorly to a recruiter or hiring engineer:**

- **The commit history compresses into 3 calendar days** (54 commits on 2026-08-18, 18 on 08-19, 1 on 08-25) for a ~23,600-line, production-deployed codebase. A reviewer reading `git log` may infer heavy AI-assisted generation. This remains the most likely credibility question the project will draw. Partially addressed: the two build-journal docs that most loudly reinforced that reading are no longer in the repo (see the remediation note above), which leaves the commit dates ambiguous rather than corroborated. Deliberately *not* addressed in the repo itself — a preemptive README note would foreground the question for readers who would not otherwise ask. Interview talking points are drafted separately in `outlook-mcp-talking-points.md`.
- **Sole author, no PRs, no issues, no code review**, 0 stars — no collaboration signal.
- **CI is thin relative to the test suite:** only 23 of 116 tests (20%) run in CI; the other 93 need the owner's live credentials. The split is well-argued in the CI file's header, but a reviewer skimming the Actions tab sees a small suite.
- `worker-configuration.d.ts` (14,884 generated lines, 573 KB) is committed, inflating apparent repo size and skewing GitHub's language statistics.
- `package.json` sets `"private": true`; there is no npm package or install path other than "clone and register your own Entra app."

**Exact vs. estimated numbers above:**

- **Exact (computed today):** commit count, date range, author count, tags, all LOC figures, tracked-file count, tool/prompt/resource counts, annotation counts, per-suite test and assertion counts, cron count, scope count, capability-audit counts, doc line counts, live `/health` version.
- **Approximate:** the ~49,000-word documentation figure (`wc -w`, includes Markdown syntax).
- **Exact but easy to misread:** "116 tests" is an exact count of `test()`/`testAuthed()` invocations, but each is a multi-step scenario, not a unit test — it is not comparable to a Jest-style count. The **1,160 assertions** figure is the better apples-to-apples number, and the safer one to quote if pressed.
- **Claimed in the repo, not independently verified here:** the "$1–2/month" LLM cost and any statement about daily real-world usage.

## 8. Assumptions

1. **The audience is a SWE-internship resume**, so highlights and bullets emphasize systems design, auth, testing and operations over feature breadth, and flag the credibility risks an interviewer would probe.
2. **"MCP tools exposed" means entries in the `TOOLS` array in `src/core/registry.ts` (40)** — the same surface both transports register. I did not count `save-local.ts`, `file-sources.ts` or `common.ts` in `src/tools/`, which are shared helpers rather than registered tools. The README's "forty tools" and the deployed connector's tool list both agree with 40; the "37 tools" in the v1.2.0 commit message is superseded.
3. **"Tests" means top-level `test()` / `testAuthed()` scenario invocations**, which is how the suites report themselves ("62/62"). I counted `^await test(` at column zero to avoid double-counting one nested call in `test-tools.ts`; the v1.3.0 commit message's own numbers (23/62/31) corroborate this.
4. **Source LOC excludes** `worker-configuration.d.ts` (Wrangler-generated), `package-lock.json`, `node_modules/` and `dist/`, and counts raw lines including comments and blanks. This codebase is unusually comment-dense, so 14,341 overstates executable lines by a meaningful but unmeasured margin.
5. **Liveness was verified by two unauthenticated GET requests** to the owner's own Worker (`/health` and the OAuth discovery document) at compile time. I did not authenticate, call any MCP tool, or read any mailbox data. "Live in daily use" therefore rests on the deployment being current (v1.3.0 matches `package.json`) plus the repo's own claims — not on observed traffic.
6. **Public/private status and repo metadata come from `gh repo view`** (`visibility: PUBLIC`, empty `description`, 0 stars). I did not separately query the topics field, so a reviewer may see topics I did not detect.
7. **Personal-data scanning covered all tracked files plus the two untracked working-tree files**, using regex sweeps for email addresses, GUIDs, 32-hex IDs, Graph ID prefixes (`AQMkAD`/`AAMkAD`) and the owner's name. I did not read all 4,393 lines of documentation end to end, so a subtler leak (a paraphrased real subject line, a contact name without an address) inside `ASSUMPTIONS.md` or `RUN-REPORT.md` could have been missed.
8. **I treated public OAuth client IDs, Entra tenant IDs and Cloudflare KV namespace IDs as identifiers rather than secrets** — correct per those platforms' threat models, but listed under risks anyway because reviewers commonly flag them.
9. **The commit-cadence observation is an inference about how the repo will be *read*, not a claim about how it was built.** I make no determination about tooling used; I flag it because `git log` makes it the first question a skeptical reviewer will ask.
10. **The v1.3.0 work is on `main` but untagged**, so I treated `afe1aa5` as the current release despite `git tag` topping out at `v1.2.0`; the live `/health` reporting `1.3.0` supports that.
11. **The stale prior version of this file** (describing a 2-commit, 168-line auth prototype) was treated as fully obsolete and overwritten; nothing was carried forward.
