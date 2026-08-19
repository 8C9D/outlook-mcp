# Setting up outlook-mcp

A complete walkthrough from nothing to a working server, for someone who has never seen this
repository. It covers the parts that are genuinely fiddly — the Microsoft app registration — in the
detail they need, including the three errors this project itself hit while getting them wrong.

You need:

- a **personal** Microsoft account (`@outlook.com`, `@hotmail.com`, `@live.com`). This server signs in
  against the `/consumers` authority and is not built for a work or school account.
- **Node 20 or newer**.
- about fifteen minutes for steps 1–3. Steps 4 and 5 (the hosted server) are optional and take longer.

At any point, `npm run doctor` will tell you what is and is not working, and what to do about it.

---

## 1. Register an app with Microsoft

The server signs in *as you*, using an app registration you own. It is free, takes a few minutes, and
two of its settings are the ones people get wrong.

1. Go to <https://entra.microsoft.com> (or the **Microsoft Entra ID** blade of
   <https://portal.azure.com>) and sign in with the personal account whose mailbox this is.
2. **App registrations → New registration**.
3. **Name**: anything — `outlook-mcp` is fine.
4. **Supported account types**: choose **Personal Microsoft accounts only**.
   *This is pitfall #1.* The default is an organizational option, and with it the sign-in fails with
   `AADSTS50020` or `AADSTS700016` no matter what else is right.
5. **Redirect URI**: leave it **empty**. This server uses the device-code flow, which needs none.
6. **Register**. Copy the **Application (client) ID** from the overview page — a GUID. That is the
   only value you need from this portal.
7. Open **Authentication → Advanced settings** and set **Allow public client flows** to **Yes**, then
   **Save**. (In the manifest this is `"allowPublicClient": true`.)
   *This is pitfall #2.* Without it Entra treats the app as confidential, demands a client secret that
   a desktop app cannot keep, and the sign-in fails with `AADSTS70002`.
8. Open **API permissions → Add a permission → Microsoft Graph → Delegated permissions** and add:

   | Permission | What stops working without it |
   | --- | --- |
   | `User.Read` | knowing which account is signed in (and the hosted server's allowlist) |
   | `Mail.Read` | `search_mail`, `read_message`, `read_thread`, `check_new_mail` |
   | `Mail.ReadWrite` | drafts, moves, flags, categories, folders, `.eml` export |
   | `Mail.Send` | `send_draft` — the only send path |
   | `Calendars.ReadWrite` | every calendar tool |
   | `Contacts.ReadWrite` | `search_contacts`, `manage_contact` |
   | `MailboxSettings.ReadWrite` | `auto_reply`, `mailbox_settings` |
   | `Tasks.ReadWrite` | `list_tasks`, `manage_task` |
   | `Files.ReadWrite` | the OneDrive tools (`search_files`, `upload_file`, …) |

   `offline_access` is requested automatically by the sign-in library and does not need adding by
   hand. There is no "grant admin consent" step for a personal account: **you** consent at sign-in, to
   exactly the scopes the app asks for. That is why adding a permission later means running
   `npm run login` again — an already-cached token never gains a scope.

### When sign-in fails, the error number says which setting is wrong

| Error | What it means | Fix |
| --- | --- | --- |
| `AADSTS70002` | Entra wants a client secret, i.e. it thinks this is a confidential client | Authentication → **Allow public client flows: Yes** (step 7) |
| `AADSTS50020` | the personal account does not exist in the directory the app is scoped to | Supported account types must include **personal Microsoft accounts** (step 4) |
| `AADSTS700016` | the client id was not found in the consumers directory | wrong `AZURE_CLIENT_ID`, or the same audience problem as `AADSTS50020` |
| plain `403` from Graph | the token is valid but carries no permission for that call | add the delegated permission (step 8), then **`npm run login` again** — consent is per scope, at sign-in |

`npm run doctor` prints these same translations when it meets one of these errors live.

---

## 2. Install and sign in

```bash
git clone <this repository> outlook-mcp
cd outlook-mcp
npm install

printf 'AZURE_CLIENT_ID=%s\n' "<the Application (client) ID from step 1>" > .env

npm run login     # prints a code; enter it at https://microsoft.com/devicelogin
npm run doctor    # every check should say PASS
```

`.env` and the sign-in cache `.token-cache.json` (mode `0600`) are both gitignored and never leave the
machine. The MCP server itself **never** prompts for sign-in — it only refreshes the cached token
silently — so `npm run login` is the one command that can open a sign-in, and you run it deliberately.

If the doctor is unhappy, its output names the fix. The common ones:

| Doctor says | Do this |
| --- | --- |
| `FAIL AZURE_CLIENT_ID` | create `.env` as above with the client id from step 1 |
| `FAIL token cache` | run `npm run login` |
| `FAIL silent token acquisition` | the cached sign-in expired or was revoked — `npm run login` |
| `FAIL granted scopes` | add the missing permission in step 8, then `npm run login` again |
| `WARN built server` | run `npm run build` (only needed for a `node dist/server.js` client config) |

---

## 3. Point a client at it

The server speaks MCP over stdio. Build it once, then register it with your client.

```bash
npm run build     # produces dist/server.js
which node        # the absolute path your client will need
```

For **Claude Desktop**, add this to
`~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) under `mcpServers`:

```json
"outlook": {
  "command": "/absolute/path/to/node",
  "args": ["/absolute/path/to/outlook-mcp/dist/server.js"]
}
```

- The `command` must be an **absolute path** to the node binary: Claude Desktop does not inherit your
  shell `PATH`. If you use nvm, upgrading node changes this path and the server will stop launching
  until you update it.
- The working directory does not matter — the server finds its own project root, and so its `.env` and
  token cache, from the location of its own module.
- Claude Desktop reads the config only at launch. Quit it fully (Cmd+Q) and reopen.
- Logs land in `~/Library/Logs/Claude/mcp-server-outlook.log`; that is the first place to look if the
  server shows as failed.
- After changing the code, run `npm run build` again — the client runs `dist/`, not `src/`.

That is a complete, working install. Everything below is optional.

---

## 4. Optional: deploy the hosted server

The same tools can run as a Cloudflare Worker, which lets claude.ai reach the mailbox with your laptop
closed, and adds the two things a laptop cannot do: receive Microsoft's push notifications for arriving
mail, and hand out short-lived links to attachment bytes it has nowhere to save. You need a Cloudflare
account (the free plan is enough).

```bash
npx wrangler login

npx wrangler kv namespace create OUTLOOK_KV   # copy the printed id into wrangler.jsonc
npx wrangler kv namespace create OAUTH_KV     # likewise
```

Edit `wrangler.jsonc`:

- put the two namespace ids in `kv_namespaces`;
- set `vars.PUBLIC_BASE_URL` to the URL your Worker will have
  (`https://<worker-name>.<your-subdomain>.workers.dev`) — Microsoft is told to deliver notifications
  to `PUBLIC_BASE_URL + /notifications`, so it must match exactly;
- change `name` if you want a different hostname.

Then set the same URL, path included, as `resourceMetadata.resource` in `src/worker/index.ts`
(`https://…/mcp`). RFC 9728 discovery requires it to match the URL a client is given, character for
character.

```bash
npm run deploy    # prints the workers.dev URL

# Three secrets. Never committed; `wrangler secret put` reads stdin.
printf '%s' "<Application (client) ID>" | npx wrangler secret put MS_CLIENT_ID
printf '%s' "<Graph /me id>"            | npx wrangler secret put ALLOWED_MS_USER_ID
printf '%s' "<your email address>"      | npx wrangler secret put ALLOWED_MS_UPN

npm run seed:kv        # pushes the mailbox refresh token into KV; prints both allowlist values
npm run test:remote    # live checks against the deployed endpoint
```

`npm run doctor` prints the Graph `/me` id and address you need for the two allowlist secrets, and
`npm run seed:kv` prints them too. Together they are what makes the endpoint single-user: no other
Microsoft identity can complete an authorization.

Notes that matter:

- `npm run seed:kv` reads `.token-cache.json`, so run `npm run login` first if the local sign-in is
  stale. Re-run it **only** after a fresh login: at any other time it would overwrite the Worker's
  rotated refresh token with an older one.
- Do **not** set `ALLOW_DIRECT_AUTHORIZE` on the deployed Worker. Leaving it unset is what keeps the
  non-interactive authorize path disabled in production; test `r5` asserts the deployed Worker refuses
  it.
- `ANTHROPIC_API_KEY` is optional and only powers the two LLM features, which **ship disabled**. See
  the README's cost section before setting it.
- Run `npm run test:remote` headlessly (`MCP_REMOTE_HEADLESS=1`) and the checks that need a bearer are
  reported as SKIP; run it in a terminal and it asks you to complete a device-code sign-in.

## 5. Optional: add it to claude.ai as a custom connector

1. In claude.ai open **Settings → Connectors** (on Team/Enterprise: **Organization settings →
   Connectors**, added by an owner).
2. **Add custom connector**.
3. Paste the full URL **including the path**: `https://<your-worker>.workers.dev/mcp`.
4. Leave the OAuth Client ID and Secret fields **empty** — the server supports dynamic client
   registration, so Claude registers itself.
5. **Add**, then **Connect**. A tab opens on the Worker's `/authorize` page, showing a device code.
6. Open <https://microsoft.com/devicelogin> in another tab, enter the code, and sign in **as the
   allowlisted account**. Any other account is refused.
7. The `/authorize` page polls until sign-in completes and returns to claude.ai. The tools, prompts and
   resources are now available there.

## 6. Turning it off again

- **Disconnect one client**: remove the connector in claude.ai, then delete its records from
  `OAUTH_KV` (`npx wrangler kv key list --namespace-id <id> --remote`, then `… key delete`). Edge
  caching means deletions take up to a minute to bite.
- **Revoke the hosted server's mailbox access**: delete `ms:refresh_token` from `OUTLOOK_KV`. Every
  tool call then fails with an auth error; `npm run seed:kv` restores it.
- **Cut Microsoft off entirely**: remove the app at <https://account.live.com/consent/Manage>. That
  invalidates the local cache and the Worker's KV token together.
- **Take the endpoint down**: `npx wrangler delete`. The KV namespaces survive and must be deleted
  separately if you want the stored tokens gone.
- **Locally**: delete `.token-cache.json`.

---

Read [README.md](README.md) next: what the tools do, the security model, and what the optional LLM
features cost.
