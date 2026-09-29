# Security policy

## Reporting a vulnerability

Please report vulnerabilities **privately** — do not open a public issue.

Use GitHub's private vulnerability reporting on this repository
(*Security → Report a vulnerability*), which opens a private advisory thread
visible only to you and the maintainer.

This is a personal project with a single maintainer; expect an acknowledgement
within a few days rather than hours. Please include enough detail to reproduce
(the transport — stdio or Worker — and the tool or route involved).

## Threat model in brief

The full reasoning lives in the README's
[Security model in detail](README.md#security-model-in-detail); the shape of it:

- **Mailbox content is untrusted input.** An email can contain text that tries
  to instruct a model into sending, deleting or forwarding things (prompt
  injection). The design answers structurally, not by prompting the model to be
  careful.
- **No tool composes-and-sends.** No tool calls `/me/sendMail`; sending is
  two-step — the complete message must exist as a reviewable draft before
  `send_draft` can name it.
- **One opt-in route sends without a draft, and only to the owner.**
  `POST /self-alert` (and the heartbeat watchdog behind
  `POST /self-alert/heartbeat`, run every 6 hours on the subscription-upkeep
  cron) lets a scheduled job the owner runs elsewhere email the owner. It is
  off — both routes answer `404` — unless the `SELF_ALERT_SECRET` Worker
  secret is set; it requires that secret as a bearer token, compared in
  constant time; the recipient is always `ALLOWED_MS_UPN`,
  with no recipient, cc, bcc or reply-to field in the request and unknown
  fields refused; and it stops at 20 mails per UTC day.
- **Mailbox deletes are soft** (Deleted Items, recoverable). The one
  irreversible operation is `manage_task` delete, which says so in its
  description; per-call approval prompts are the intended backstop for all
  destructive tools.
- **Inbox rules cannot forward.** Rules act on all future mail after one
  approval, so their action list is restricted to move / mark read / soft
  delete.
- **The hosted Worker is single-user and interactive-only.** Nothing anonymous
  reaches `/mcp`; only one Microsoft identity (matched on the Graph `/me` id or
  UPN captured at setup) can complete authorization, and the non-interactive
  `POST /authorize` path is refused in production.
- **`/notifications` is the one public route, and it is write-only and
  content-free.** Deliveries must carry the subscription's random
  `clientState`; forged deliveries are discarded, the route never echoes stored
  state, and it answers `202` either way.
- **The autonomous LLM path (auto-filing) cannot send, delete or reply —
  structurally.** The classifier imports no Graph transport and reaches only a
  five-method read/move/categorize interface with Deleted Items and Junk Email
  removed from its folder allowlist; a test walks the import graph and fails if
  that ever stops being true. Both LLM features ship disabled.
- **No secrets in the repo.** Tokens live in the local MSAL cache
  (`~/.config/outlook-mcp/token-cache.json`, outside the repo, mode 0600) or in Cloudflare KV; CI runs
  only the offline test tier and needs no credentials.

Reports about weaknesses in any of these boundaries are very welcome.
