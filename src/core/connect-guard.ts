// The checks in front of connecting a new client to the mailbox.
//
// The Microsoft sign-in proves who the owner is, not who started the connection
// or where the resulting grant goes, so on its own it can be phished: a page or
// a code handed to the owner ends with an attacker holding a working grant. Two
// layers answer that:
//
//  - redirectAllowed sends codes only to Claude and to this machine, which stops
//    a grant being delivered straight to a site the attacker runs.
//  - CONNECT_SECRET, opt-in, is asked for before any sign-in starts. It stops the
//    relay, where attackers open the page themselves and pass the owner the code:
//    they can read their own browser's redirect, so only a secret they do not
//    hold keeps them from getting a code to pass on.
//
// spendAnonymousWrite is separate. Registering a client and starting a sign-in
// are both open to anyone and both write to KV, which Workers Free caps at 1,000
// writes a day across the account. Past that, the mailbox token cannot be saved
// after a rotation, so anonymous callers get a small daily allowance of their own.
import type { StateStore } from "./state.js";

/** Where a code may go, beyond loopback. ALLOWED_REDIRECT_HOSTS adds to these. */
export const DEFAULT_REDIRECT_HOSTS = ["claude.ai", "claude.com"];

/** Same floor as SELF_ALERT_SECRET: long enough that guessing online is hopeless. */
export const CONNECT_SECRET_MIN_LENGTH = 32;

/** Anonymous KV writes allowed per UTC day. Connecting a client takes two. */
export const ANONYMOUS_WRITES_PER_DAY = 40;

/** The hostname a redirect URI points at, for display. */
export function redirectHost(redirectUri: string): string {
  try {
    return new URL(redirectUri).hostname;
  } catch {
    return redirectUri;
  }
}

export function redirectAllowed(redirectUri: string, extraHosts?: string): boolean {
  let url: URL;
  try {
    url = new URL(redirectUri);
  } catch {
    return false;
  }
  const host = url.hostname.toLowerCase().replace(/^\[|\]$/g, "");
  if (host === "localhost" || host === "127.0.0.1" || host === "::1") return true;
  if (url.protocol !== "https:") return false;
  const extras = (extraHosts ?? "").split(",").map((h) => h.trim().toLowerCase()).filter(Boolean);
  return [...DEFAULT_REDIRECT_HOSTS, ...extras].some((allowed) => host === allowed || host.endsWith(`.${allowed}`));
}

/**
 * Claims one anonymous write for today. False means refuse; so does a failed
 * write. A soft cap, not an exact one: KV has no compare-and-swap, so concurrent
 * callers can read the same count. KV also refuses a second write to one key
 * within a second, and that refusal lands here, which keeps the overshoot small
 * against a quota of 1,000.
 */
export async function spendAnonymousWrite(store: StateStore, now: Date): Promise<boolean> {
  const key = `anon-writes:${now.toISOString().slice(0, 10)}`;
  const used = Number(await store.get(key)) || 0;
  if (used >= ANONYMOUS_WRITES_PER_DAY) return false;
  try {
    await store.put(key, String(used + 1), { ttlSeconds: 172_800 });
  } catch {
    return false;
  }
  return true;
}
