#!/bin/sh
# School-instance wrapper: runs outlook-mcp against a work/school (M365)
# account via the `outlook-mcp-school` multi-tenant app registration. Local
# stdio only — the Worker/remote mode, webhooks, KV and LLM features are out
# of scope for this instance.
#
# Usage: ./run-school.sh [server|login|verify|doctor]   (default: server)
#
# Exports win over .env because dotenv never overrides already-set variables;
# the personal instance (plain `npm run serve` etc.) is untouched.
cd "$(dirname "$0")" || exit 1

export AZURE_CLIENT_ID="20466c8c-aeb5-4724-b671-5ac3e6328540"
export OUTLOOK_MCP_AUTHORITY="organizations"
export OUTLOOK_MCP_TOKEN_CACHE=".token-cache.school.json"
# The school registration deliberately omits Mail.Send and Files.ReadWrite.
export OUTLOOK_MCP_SCOPES="User.Read,Mail.Read,Mail.ReadWrite,Calendars.ReadWrite,Contacts.ReadWrite,MailboxSettings.ReadWrite,Tasks.ReadWrite"

entry="${1:-server}"
case "$entry" in
  server|login|verify) exec npx tsx "src/$entry.ts" ;;
  doctor)              exec npx tsx "src/scripts/doctor.ts" ;;
  *) echo "usage: $0 [server|login|verify|doctor]" >&2; exit 2 ;;
esac
