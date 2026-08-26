# outlook-mcp

MCP server connecting Claude to a **personal** outlook.com mailbox through Microsoft Graph. One shared tool registry is served over two transports: a local stdio server (Node + MSAL, on-disk token cache) and a Cloudflare Worker that claude.ai adds as a custom connector. **This repo is public.**
`README.md` is the tool surface and security model, `SETUP.md` the install from zero, `CAPABILITIES.md` the Outlook-parity audit — read those rather than restating them here.

## Layout

- `src/core/registry.ts` — the single source of the tool/prompt/resource surface; both transports build their server from it, so a tool not registered there (with its MCP annotations) exists on neither.
- `src/tools/` one file per tool; `src/core/` transport-agnostic logic; `src/worker/` Worker-only (KV tokens, OAuth, notifications, LLM); `src/scripts/` doctor and seed-kv.

## Commands worth knowing

- `npm run doctor` — first thing to run when anything is broken; `-- --env-only` needs no credentials.
- `npm run test:offline` — no Graph, no KV, no secrets. The only tier CI runs.
- `npm run test:tools` — hits the **real mailbox**, creating then sweeping `[MCP TEST]` artifacts.
- `npm run test:remote` — hits the deployed Worker; needs an interactive device-code sign-in or reports the authenticated tests as SKIP. `MCP_REMOTE_URL` retargets it at another deployment.
- `npm run seed:kv` — lifts the refresh token from the local token cache (`~/.config/outlook-mcp/token-cache.json`) into KV. Re-run **only** after `npm run login`, never routinely.
- `./run-school.sh [server|login|verify|doctor]` — the second instance (work/school M365 account, its own client id, authority, cache and scopes). Its exports beat `.env`; the personal instance is untouched.

## Standing rules

- No secret, token or live mailbox content ever enters this repo. `ASSUMPTIONS.md`, `RUN-REPORT.md` and `cleanup-progress.md` are gitignored build journals quoting real mail — never commit them, and never lift their content into a tracked file.
- `send_draft` is the only send path and it takes an existing draft id. Never add a compose-and-send tool, never call `/me/sendMail`, and never let an autonomous path send, delete or reply.
- Mailbox deletes stay soft; `manage_task` delete is the one documented exception and says so loudly.
- Mail is untrusted input: the auto-filing rails (`PROTECTED_SUBJECT_PATTERNS`, confidence threshold, daily call cap) may be extended, never weakened, and both LLM features ship disabled.

## Gotchas

- Claude Code runs this as `node dist/server.js` with `OUTLOOK_TOOL_PROFILE=mail`: src edits stay invisible until `npm run build`, and the session sees the mail subset, not all forty tools.
- The Worker rotates the Microsoft refresh token on every exchange and Microsoft invalidates the previous one — a local `npm run login` breaks the KV chain until `npm run seed:kv` reseeds it.
- `wrangler.jsonc`'s `PUBLIC_BASE_URL` is the only place the deployed hostname is written down.
- Never resolve a path from `process.cwd()` — MCP clients launch with an arbitrary one. Use `PROJECT_ROOT`.
- `ALLOW_DIRECT_AUTHORIZE` is local-only by design; remote test r5 asserts the deployment refuses it.
