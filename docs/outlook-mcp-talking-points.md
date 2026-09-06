# outlook-mcp — interview talking points

Private notes. Not in the repo, not for distribution.

**What this is for.** `git log` shows 73 commits across 3 calendar days on a 23,600-line deployed system. Some interviewers will ask about that; most who notice won't ask, they'll just quietly discount. These notes are so the answer is ready and unrehearsed-sounding, not so the question gets dodged.

**One rule: don't oversell the authorship.** If you used AI assistance heavily, say so plainly when asked. Interviewers in 2026 assume it; what they're actually probing is whether you can defend the decisions. A candidate who says "yes, heavily — here's the call I made that the model got wrong" is far stronger than one who implies they hand-typed 14K lines. Getting caught shading this is fatal; being straightforward about it costs almost nothing.

---

## Q: "This looks like it was built in three days. How?"

**The shape of an honest answer** (fill in your own specifics — I don't know your actual process):

1. Name the assistance level directly, without apology.
2. Say what the compressed timeline actually was — a focused sprint, prior familiarity with the domain, or work that had been designed in your head before the first commit.
3. Pivot immediately to a decision you made that the tooling would not have: that's the part they're actually evaluating.

**Don't:** call it "just prompting," volunteer a percentage breakdown, or get defensive about the commit density. Sprints are normal.

---

## Q: "Which parts did you actually design?"

These are the decisions in the repo that are genuinely non-obvious — each is a real fork in the road where the alternative was available and worse. Pick two or three you can defend cold. **Verify each still matches the code before an interview.**

### 1. Device code instead of an auth-code redirect
The Entra app is registered as a public native client with no web redirect URI. Adding one means changing the registration; device code needs no redirect at all, so the Worker verifies Microsoft identity against the app exactly as registered. The device-code token is used for one thing — a Graph `/me` call to check the single-user allowlist — and never stored. The mailbox refresh token is a separate credential in KV.
> **The tell that this is real design:** the local `wrangler dev` path enables a non-interactive authorize route that production refuses, and remote test `r5` asserts production refuses it. You built the escape hatch and then wrote a test that it stays shut.

### 2. MSAL doesn't run on `workerd`, so the refresh grant is hand-rolled
MSAL Node reaches for crypto/network APIs `nodejs_compat` doesn't provide. `src/worker/ms-token.ts` issues the refresh-token grant with plain `fetch`. The non-obvious part: Microsoft rotates the refresh token on every exchange and invalidates the old one, so the new value **must** be written back or the connector locks itself out on the very next call. Remote test `r`-tier proves rotation actually happens against the live deployment.
> Good answer to "what broke": this is a failure that only shows up on the *second* request.

### 3. Prompt injection answered structurally, not by prompting
This is the strongest thing in the repo. The argument: mail is attacker-controlled text, and asking a model to "be careful" is not a security boundary. So:
- `src/core/classifier.ts` **imports no Graph transport at all** — it acts through a `ClassifierMailbox` port whose only mutations are `move` and `categorize`. Send, delete, reply and forward are *not expressible* on that code path.
- Deleted Items and Junk are excluded from the folder allowlist handed to the model, which is what stops "move" from being a laundered "delete."
- The answer must parse as exact JSON or it's discarded.
- Offline test `o15` walks the module's **transitive import graph** and fails if Graph ever becomes reachable from the classifier.
> That last one is the point to land: the boundary is enforced by a test, not by a comment. If someone later imports a tool into the classifier, CI fails.

### 4. Two UTC crons for one local hour
Cloudflare crons are UTC-only. 07:00 America/Toronto is 11:00 UTC in EDT and 12:00 in EST, so both are scheduled and the handler drops whichever isn't 07:00 locally. The digest also refuses to draft twice for the same date, so a double-fire can't double up. Nothing drifts across DST and nothing needs redeploying twice a year.
> Small, but it's the kind of thing that shows you thought about the system running unattended for a year.

### 5. Send is structurally two-step
No tool calls `/me/sendMail`. A message must exist as a reviewable draft before `send_draft` can name it, so no single tool call composes-and-sends. The test of whether you meant it: **the health-check cron reports failures as an unsent draft in the inbox** rather than emailing an alert. The one autonomous sender is the opt-in self-alert route, and it is fixed to the owner's own address, gated by a shared secret, and capped at 20 a day, so nothing can reach anyone else.
> Scheduled send (v1.3) was added the same way: set the MAPI deferred-send-time property on the draft before the *same* single POST. New feature, no new send path.

---

## Q: "What's the weakest part?"

Have an answer ready; "nothing" reads as not having looked. Candidates, pick one you actually believe:

- **CI covers 23 of 116 tests (20%).** The other 93 need live Microsoft credentials, and putting those in CI would mean a secret in the repo, which the whole design refuses. Defensible, and the split is argued in the CI file header — but it's a real coverage gap, not just a presentation one. A fix exists (recorded Graph fixtures / a test tenant) and you chose not to build it.
- **Single-user by construction.** The allowlist is one Graph `/me` id. Multi-tenant would need per-user token isolation in KV and a real consent flow. Scoped deliberately, but it means the auth design hasn't been tested against the hard case.
- **The LLM features have limited production mileage.** They are enabled and verified working on the live instance (the deployed config reads `filingEnabled: true, digestEnabled: true`), but the runtime evidence is days old, not months. Don't claim a track record you don't have.

---

## Q: "Walk me through the architecture"

Thirty-second version: one `TOOLS` registry in `src/core/registry.ts`; both the stdio server and the Cloudflare Worker build from it, so the local install and the hosted connector can't drift. Transport differences sit behind narrow ports — `core/graph.ts` has no Node-only imports and asks `core/token.ts` for a token; attachments write to disk on stdio and become expiring bearer-gated links on the Worker, which has no filesystem. OAuth, KV, crons and webhooks are Worker-only and live under `src/worker/`.

**If they push on "why two transports":** stdio for Claude Desktop with no hosting and no OAuth; the Worker so claude.ai can use it as a custom connector, plus the three things only a hosted server can do — receive Graph change notifications, keep the subscription alive on a schedule, and serve attachment bytes it can't save to disk.

---

## Numbers worth having memorized

| | |
|---|---|
| MCP tools | 40 |
| Source LOC (excl. tests, excl. generated) | 14,341 |
| Tests / assertions | 116 across 3 tiers / 1,160 |
| CI tier | 23 offline, zero credentials |
| Capability audit | 96 rows; 19 proven infeasible via Graph |
| Live | `workers.dev`, serving v1.3.0; auto-filing + digest **enabled** in production |

**Don't quote "116 tests" as if it were a Jest count** — each is a multi-step scenario. If pressed on rigor, 1,160 assertions is the honest comparable, and it's the more impressive number anyway.

---

## Repo hygiene done (2026-08-25)

- `cleanup-progress.md` and `recover_deletions.py` gitignored — they quote 176 real email addresses and real mailbox contents. Never were committed.
- `ASSUMPTIONS.md` and `RUN-REPORT.md` untracked (still on disk locally). They were 2,842 lines of batch/gate build-journal in a public repo, and the loudest provenance signal there. The seven inbound references from source comments and docs were rewritten so nothing dangles.
- Deployed hostname moved out of `src/` entirely — it now lives only in `wrangler.jsonc`. This also fixed a real latent bug: the OAuth `resourceMetadata.resource` was a hardcoded literal under a comment claiming it came from the wrangler var.

**Still open, decide before sharing the link:** the repo has no GitHub description or topics; `SECURITY.md` lists your real email as the contact; `README.md:1029` shows an absolute `/Users/arthurzhang/...` path in the Claude Desktop config example.


## Q: "Why build an MCP server at all? Why not use the Microsoft 365 connector?" (added 2026-09-05)
- The first-party Microsoft 365 connector covered work/school tenants, not personal Outlook accounts. Mine is personal, so there was no option; I built the server for my own mailbox, then added the guardrails a mailbox needs (draft-before-send, the auto-filer fenced off from send/delete).
- Claude cannot call Microsoft Graph by itself: it has no credentials, no OAuth flow, no token storage, no schema for the operations. MCP is the protocol that gives a Claude client tools. The alternative is not "direct connection" but curl with a pasted token in a terminal.

## Q: "Why two transports?" (added 2026-09-05)
- A stdio server is launched as a child process by the client on the same machine (Claude Code, the desktop app). A browser tab cannot spawn processes, so claude.ai on the web or phone can only reach a URL. Hence the Cloudflare Worker, and hence the OAuth 2.1 authorization server in front of it (a remote server on the public internet must know who is calling). One codebase, one tool registry, two front doors.

## Plain-language versions of the three page bullets (added 2026-09-05; the page wording is deliberately plainer than the 2026-08-27 form)
- Registry: src/core/registry.ts holds one TOOLS array; both entry points read it, so the local and hosted servers expose the same 40 tools by construction.
- OAuth: two layers. The front door is my OAuth 2.1 authorization server (PKCE, dynamic client registration, RFC 9728 discovery) for MCP clients; the back door is Microsoft's OAuth, kept alive by a refresh-token grant I wrote because MSAL does not run on workerd, with rotated tokens written back to KV.
- Auto-filer isolation: the classifier sits behind a seven-method mailbox port whose only mutations are move and categorize; a CI test walks the import graph from the classifier and fails the build if it can reach the Graph transport (send/delete). Say "import-graph" out loud; the page says "a CI test".
