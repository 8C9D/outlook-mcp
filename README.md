# outlook-mcp

Foundation for an MCP server that connects Claude to a personal Microsoft (outlook.com) account via Microsoft Graph.
This stage contains only authentication (MSAL device-code flow) and a Graph verification script — no MCP server or tools yet.

## Setup

1. `npm install`
2. Create `.env` in the project root:

   ```
   AZURE_CLIENT_ID=<Application (client) ID of the `outlook-mcp` Entra app registration>
   ```

   The registration must allow personal Microsoft accounts, have public client flows enabled, and have delegated Graph permissions: Mail.Read, Mail.ReadWrite, Calendars.ReadWrite, User.Read, offline_access.
   (As of the last run this registration does not exist yet — see ASSUMPTIONS.md.)

3. `npm run verify`

On first run it prints a device-code sign-in URL and code; open the URL, enter the code, and sign in with the personal Microsoft account.
It then checks `/me`, the inbox, and the next 7 days of calendar, and prints a PASS/FAIL summary.

## Token cache

Tokens are cached in `.token-cache.json` (project root, mode 0600, gitignored), so subsequent runs need no sign-in.
To force re-authentication, delete `.token-cache.json` and run verify again.
