# Assumptions and decisions

Recorded while executing the "Personal-Outlook MCP foundation" task on 2026-08-18.

## Project location

- Chose `~/dev/outlook-mcp` because `~/dev` already exists and is the established personal-code directory (no `~/projects` or `~/code` present).

## Phase A — BLOCKED: Entra app registration could not be created

The portal work was attempted with Claude in Chrome and is blocked by Microsoft-side restrictions, not by a missed step:

1. **Account chosen for the portal**: `arthur.yuhao.zhang@outlook.com` (the personal Microsoft account matching the user's identity), selected from the Azure portal account picker.
2. **Personal-account context** (consumer pseudo-tenant `f8cdef31-a31e-4b4a-93e4-5f571e91255a`): App registrations list is empty, and the **New registration** dialog refuses with "The ability to create applications outside of a directory has been deprecated", offering only "joining the M365 Developer Program" or "signing up for Azure". Portal settings → Directories confirms the account has **no directories** and no subscriptions.
3. **University of Waterloo tenant** (`[redacted-second-alias]`, also signed in): the App registrations blade returns **401 "You do not have access"** — the tenant blocks members from viewing/creating app registrations.
4. Both remedies Microsoft offers (M365 Developer Program signup, Azure signup) create new accounts/agreements, which the automation deliberately does not perform on the user's behalf. **The user must do one of these once**, then create the `outlook-mcp` registration (or re-run this task).

Consequences:

- No Application (client) ID exists; `.env` contains an empty `AZURE_CLIENT_ID` with a note.
- The portal-state record requested in Phase A step 6 (account type, public client flows, permission list) cannot be provided; the *target* state remains: personal Microsoft accounts allowed, public client flows = Yes, delegated permissions exactly Mail.Read, Mail.ReadWrite, Calendars.ReadWrite, User.Read, offline_access, and **no Mail.Send**.
- The verify runs (Phase B6) could not be executed against Graph; `npm run verify` currently fails fast with the intended "AZURE_CLIENT_ID is not set" message.

## Phase B — implementation decisions

- Node v24.15.0 is installed, satisfying the "Node 20+" target; tsconfig targets ES2022 with NodeNext modules.
- `npm run verify` uses `tsx` (dev dependency) rather than a compiled build, since the project is scaffold-only; `npm run typecheck` runs `tsc --noEmit`.
- The token cache is written with mode 0600 at creation and re-`chmod`ed after every write, since `writeFile`'s `mode` only applies when the file is created.
- `getAccessToken` uses the first cached account if several exist — only one personal account is expected to ever sign in.
- Calendar check window is computed as now → now+7d in UTC ISO strings (URL-encoded), with `Prefer: outlook.timezone="America/Toronto"` controlling the *returned* event times, per the spec.
- Initialized a plain local git repository (no remote), since the spec's gitignore requirements imply version control.
- `.env` is committed to nothing (gitignored) but was still created with an empty value so the fail-fast path and the file's expected format are in place.
