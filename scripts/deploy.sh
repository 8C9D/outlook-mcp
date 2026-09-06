#!/bin/sh
# Deploys the Worker with PUBLIC_BASE_URL injected from the gitignored .env.
#
# The Worker's public origin names the owner's Cloudflare account, so it is not
# a `vars` entry in wrangler.jsonc (this repository is public). It lives in
# .env as PUBLIC_BASE_URL and reaches the Worker through `wrangler deploy
# --var`; a PUBLIC_BASE_URL already in the environment wins over .env. Extra
# arguments are passed to wrangler unchanged (e.g. `npm run deploy -- --dry-run`).
cd "$(dirname "$0")/.." || exit 1

if [ -z "$PUBLIC_BASE_URL" ] && [ -f .env ]; then
  PUBLIC_BASE_URL=$(sed -n 's/^PUBLIC_BASE_URL=//p' .env | tr -d '"' | tail -n 1)
fi
case "$PUBLIC_BASE_URL" in
  https://*) ;;
  *)
    echo "PUBLIC_BASE_URL is not set: add PUBLIC_BASE_URL=https://<worker-name>.<your-subdomain>.workers.dev to .env (SETUP.md §4)" >&2
    exit 1 ;;
esac

exec npx wrangler deploy --var "PUBLIC_BASE_URL:${PUBLIC_BASE_URL%/}" "$@"
