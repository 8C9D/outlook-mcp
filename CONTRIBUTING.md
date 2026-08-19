# Contributing

This is a personal tool, built for one mailbox and published in case it is
useful to others. Issues and pull requests are welcome all the same — small,
focused changes have the best chance of landing quickly.

## What you can run without any credentials

```sh
npm ci
npm run typecheck      # both tsconfigs (node + worker)
npm run test:offline   # the offline tier — fixtures and stubs, no Graph, no secrets
```

That pair is exactly what CI runs on every push, and it must stay green.

## The live suites

`npm run test:tools` (stdio, against a real mailbox) and `npm run test:remote`
(against a deployed Worker) require **your own** Microsoft account, app
registration and Cloudflare deployment — see [SETUP.md](SETUP.md). They are
self-cleaning, but they do read and write real mailbox state; never point them
at a mailbox you care about without reading what they do first.

## Ground rules

- No secrets, tokens, or real mailbox content in commits, fixtures, or test
  logs. Fixture data uses `.invalid`/`example.com` addresses and `[MCP TEST]`
  subjects.
- Security-relevant behavior (two-step send, soft deletes, the rules action
  allowlist, the single-user Worker auth, the auto-filing capability fence) is
  deliberate; read [SECURITY.md](SECURITY.md) before proposing changes there.
- Vulnerabilities go through [SECURITY.md](SECURITY.md), not the issue tracker.
