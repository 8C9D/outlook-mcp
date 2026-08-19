// KV key names shared by the Worker (which reads and rotates them) and the
// local seeding/test scripts (which write and inspect them). Kept free of any
// Worker or Node types so both sides can import it.

/** The current mailbox refresh token. Rotated on every exchange with Microsoft. */
export const KV_REFRESH_TOKEN = "ms:refresh_token";

/** The cached mailbox access token, as {"token","expiresAt"} JSON. */
export const KV_ACCESS_TOKEN = "ms:access_token";

// The keys below are also used by the stdio server, where the same names index
// a local JSON file instead of KV — they are state keys, not KV-only keys.

/** Delta link for one mail folder, keyed by the folder as the caller named it. */
export function deltaKey(folder: string): string {
  return `delta:${folder.toLowerCase()}`;
}

/** The mail change-notification subscription record (id, clientState, expiry). */
export const STATE_SUBSCRIPTION = "sub:mail";

/** The capped ring buffer of received change notifications, newest first. */
export const STATE_ACTIVITY = "activity:mail";

/** Settings for the two LLM features (auto-filing and the morning digest). */
export const STATE_LLM_CONFIG = "llm:config";

/** The capped ring buffer of auto-filing decisions, newest first. */
export const STATE_LLM_AUDIT = "llm:audit";

/** The Toronto date of the last morning brief that was drafted (idempotency). */
export const STATE_DIGEST_LAST = "llm:digest:last";

/** Anthropic API calls made on one America/Toronto calendar day. */
export function llmBudgetKey(torontoDate: string): string {
  return `llm:budget:${torontoDate}`;
}

/** Prefix of the short-lived attachment downloads get_attachment hands out. */
export const DOWNLOAD_PREFIX = "dl:";

/** One pending attachment download, keyed by its unguessable id. */
export function downloadKey(id: string): string {
  return DOWNLOAD_PREFIX + id;
}
