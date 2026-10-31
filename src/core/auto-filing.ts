// Settings, budget rails and the audit trail for the two LLM features.
//
// Deliberately knows nothing about Microsoft Graph or the Anthropic API: it is
// the state layer both features sit on, so the classifier can be unit-tested
// against a memory store with no network at all.
//
// BOTH FEATURES SHIP DISABLED. Nothing here turns itself on; the owner enables
// them with manage_auto_filing after reading what they cost.
import {
  STATE_DIGEST_DRAFT,
  STATE_DIGEST_LAST,
  STATE_LLM_AUDIT,
  STATE_LLM_CONFIG,
  STATE_LLM_PREFS,
  errorCounterKey,
  llmBudgetKey,
} from "./kv-keys.js";
import { readJson, writeJson, type StateStore } from "./state.js";

/** How many decisions the audit ring keeps. Older entries fall off. */
export const AUDIT_CAP = 100;

/** Default minimum confidence for the classifier to act. */
export const DEFAULT_THRESHOLD = 0.8;

/** Default ceiling on Anthropic calls per America/Toronto calendar day. */
export const DEFAULT_DAILY_CALL_CAP = 200;

/** Hard bounds the tool enforces on the tunables. */
export const THRESHOLD_MIN = 0.5;
export const THRESHOLD_MAX = 1;
export const DAILY_CAP_MAX = 2000;

/**
 * Subjects that must never be sent to a model or moved out of the inbox, no
 * matter what the classifier would say. Matched case-insensitively as
 * substrings, because sign-in mail phrases these a dozen ways. This list is
 * compiled in; `skipPatterns` in the config extends it and cannot shrink it.
 */
export const PROTECTED_SUBJECT_PATTERNS = [
  "one-time passcode",
  "one time passcode",
  "verify log in",
  "verify login",
  "single-use code",
  "single use code",
  "security code",
  "verification code",
  "your code is",
  "two-factor",
  "two factor",
  "two-step",
  "2-step",
  "2fa",
  "multi-factor",
  "reset your password",
  "password reset",
  "reset password",
  "one-time code",
  "one time code",
  "one-time password",
  "one time password",
  "login code",
  "log in code",
  "sign-in code",
  "sign in code",
  "authentication code",
  "access code",
  "magic link",
];

/**
 * Wordings a substring cannot state safely: "otp" inside "hotpot", or a bare
 * number beside "code", as in "123456 is your Instagram code".
 */
const PROTECTED_SUBJECT_SHAPES: { name: string; shape: RegExp }[] = [
  { name: "otp", shape: /\b(otp|mfa)\b/ },
  { name: "numeric code", shape: /\b\d{4,8}\b.{0,40}\bcode\b|\bcode\b.{0,20}\b\d{4,8}\b/ },
];

/** What a protected subject becomes in the audit log: the subject may BE the code. */
export const PROTECTED_AUDIT_SUBJECT = "(protected subject)";

/**
 * Folders the classifier may never move mail into, whatever the model says.
 * This is what makes "delete" structurally unreachable from the filing path:
 * a move is the only mutation available, and the two destinations that would
 * amount to a delete (Deleted Items, Junk Email) are not on the allowlist the
 * model is given. Drafts/Sent Items/Outbox are excluded because moving received
 * mail there corrupts mailbox state; Archive is deliberately allowed.
 */
export const NEVER_FILE_INTO = [
  "deleted items",
  "junk email",
  "junk e-mail",
  "drafts",
  "sent items",
  "outbox",
  "conversation history",
  "clutter",
  "sync issues",
];

export type LlmConfig = {
  /** LLM-classified filing on incoming change notifications. */
  filingEnabled: boolean;
  /** The 07:00 America/Toronto morning brief (drafted, never sent). */
  digestEnabled: boolean;
  /** Minimum confidence for the classifier to move or categorize. */
  threshold: number;
  /** Extra never-classify subject substrings, on top of the compiled-in list. */
  skipPatterns: string[];
  /** Ceiling on Anthropic calls per Toronto day, across both features. */
  dailyCallCap: number;
};

export const DEFAULT_LLM_CONFIG: LlmConfig = {
  filingEnabled: false,
  digestEnabled: false,
  threshold: DEFAULT_THRESHOLD,
  skipPatterns: [],
  dailyCallCap: DEFAULT_DAILY_CALL_CAP,
};

/**
 * The stored config, with every field defaulted. A missing, corrupt or
 * partially-written record reads back as "both features off" — the only safe
 * direction for a fail-open/fail-closed choice here.
 */
export async function readLlmConfig(store: StateStore): Promise<LlmConfig> {
  const stored = await readJson<Partial<LlmConfig>>(store, STATE_LLM_CONFIG);
  if (!stored) return { ...DEFAULT_LLM_CONFIG };
  return {
    filingEnabled: stored.filingEnabled === true,
    digestEnabled: stored.digestEnabled === true,
    threshold: clampThreshold(stored.threshold),
    skipPatterns: Array.isArray(stored.skipPatterns)
      ? stored.skipPatterns.filter((p): p is string => typeof p === "string" && p.trim() !== "")
      : [],
    dailyCallCap: clampDailyCap(stored.dailyCallCap),
  };
}

/** Merge a patch into the stored config and return what is now in force. */
export async function writeLlmConfig(
  store: StateStore,
  patch: Partial<LlmConfig>
): Promise<LlmConfig> {
  const merged = { ...(await readLlmConfig(store)), ...patch };
  const next: LlmConfig = {
    filingEnabled: merged.filingEnabled === true,
    digestEnabled: merged.digestEnabled === true,
    threshold: clampThreshold(merged.threshold),
    skipPatterns: [...new Set(merged.skipPatterns.map((p) => p.trim().toLowerCase()))].filter(Boolean),
    dailyCallCap: clampDailyCap(merged.dailyCallCap),
  };
  await writeJson(store, STATE_LLM_CONFIG, next);
  return next;
}

function clampThreshold(value: unknown): number {
  const n = Number(value);
  if (!Number.isFinite(n)) return DEFAULT_THRESHOLD;
  return Math.min(THRESHOLD_MAX, Math.max(THRESHOLD_MIN, n));
}

function clampDailyCap(value: unknown): number {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) return DEFAULT_DAILY_CALL_CAP;
  return Math.min(DAILY_CAP_MAX, Math.floor(n));
}

/**
 * True when this subject must never reach the model. Checks the compiled-in
 * list first so a config that somehow lost its extra patterns still protects
 * sign-in codes.
 */
export function isProtectedSubject(subject: string | undefined, extraPatterns: string[] = []): string | null {
  // NFKC and collapsed whitespace, so a no-break space or full-width letters
  // do not slip past a plain substring.
  const haystack = (subject ?? "").normalize("NFKC").replace(/\s+/g, " ").toLowerCase();
  if (!haystack) return null;
  for (const pattern of [...PROTECTED_SUBJECT_PATTERNS, ...extraPatterns]) {
    const needle = pattern.trim().toLowerCase();
    if (needle && haystack.includes(needle)) return pattern;
  }
  return PROTECTED_SUBJECT_SHAPES.find(({ shape }) => shape.test(haystack))?.name ?? null;
}

/**
 * The same test over a whole message's text: a code in the preview of an
 * innocently titled mail is just as secret, and must not leave the mailbox.
 */
export function isProtectedMail(
  subject: string | undefined,
  preview: string | undefined,
  extraPatterns: string[] = []
): string | null {
  return isProtectedSubject(subject, extraPatterns) ?? isProtectedSubject(preview, extraPatterns);
}

/**
 * Untrusted text with anything shaped like a prompt marker defused, so a mail
 * body cannot close the block it was put in and speak from outside it.
 */
export function neutralizeMarkers(text: string): string {
  return text.replace(/<{3,}|>{3,}/g, "…");
}

/** True when this folder is one the classifier may never file into. */
export function isNeverFileFolder(displayName: string): boolean {
  return NEVER_FILE_INTO.includes(displayName.trim().toLowerCase());
}

/**
 * A date as America/Toronto sees it, YYYY-MM-DD. The budget counter, the digest
 * idempotency marker and the brief's subject line are all keyed on the owner's
 * local day rather than UTC's.
 */
export function torontoDateOf(when: Date): string {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Toronto",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(when);
}

/** The hour (0-23) America/Toronto is currently on. */
export function torontoHourOf(when: Date): number {
  return Number(
    new Intl.DateTimeFormat("en-CA", {
      timeZone: "America/Toronto",
      hour: "2-digit",
      hour12: false,
    }).format(when)
  );
}

// ------------------------------------------------------------- budget rails

export type BudgetVerdict = {
  allowed: boolean;
  /** Calls already made today, after this reservation when it was allowed. */
  used: number;
  cap: number;
};

/**
 * Count one Anthropic call against today's cap, and say whether it may proceed.
 * Read-modify-write, like the activity ring: for one mailbox the lost-update
 * window is smaller than the slack in a 200/day cap, and the cap exists to stop
 * a runaway loop rather than to bill anyone to the cent.
 */
export async function reserveApiCall(
  store: StateStore,
  cap: number,
  today: string
): Promise<BudgetVerdict> {
  const key = llmBudgetKey(today);
  const used = Number((await store.get(key)) ?? "0") || 0;
  if (used >= cap) return { allowed: false, used, cap };
  const next = used + 1;
  // Two days of TTL: enough for the log to stay readable past midnight, short
  // enough that yesterday's counters clean themselves up.
  await store.put(key, String(next), { ttlSeconds: 2 * 24 * 3600 });
  return { allowed: true, used: next, cap };
}

/** Today's call count without reserving anything (for manage_auto_filing status). */
export async function readApiCallsToday(store: StateStore, today: string): Promise<number> {
  return Number((await store.get(llmBudgetKey(today))) ?? "0") || 0;
}

// ----------------------------------------------------------- error counters

/**
 * Errors one feature swallowed today. Both LLM features run in the background
 * with nobody to return an error to, so their failure paths log-and-continue;
 * this counter is what makes a run of such failures visible to the daily
 * health check instead of only to someone reading `wrangler tail`.
 */
export type FeatureErrorRecord = {
  count: number;
  firstAt: string;
  lastAt: string;
  lastReason: string;
};

/**
 * Count one swallowed error against today's counter. Keyed on the
 * America/Toronto date (like the API budget) with a two-day TTL, so counters
 * clean themselves up and the health check always reads "errors today".
 * Never throws: counting an error must not be able to cause one.
 */
export async function recordFeatureError(
  store: StateStore,
  feature: "filing" | "digest",
  reason: string,
  when: Date = new Date()
): Promise<void> {
  try {
    const key = errorCounterKey(feature, torontoDateOf(when));
    const existing = await readJson<FeatureErrorRecord>(store, key);
    const record: FeatureErrorRecord = {
      count: (existing?.count ?? 0) + 1,
      firstAt: existing?.firstAt ?? when.toISOString(),
      lastAt: when.toISOString(),
      lastReason: reason.replace(/\s+/g, " ").trim().slice(0, 300),
    };
    await store.put(key, JSON.stringify(record), { ttlSeconds: 2 * 24 * 3600 });
  } catch {
    // Deliberately swallowed; see above.
  }
}

/** Today's swallowed-error record for one feature, or null when none. */
export async function readFeatureErrors(
  store: StateStore,
  feature: "filing" | "digest",
  torontoDate: string
): Promise<FeatureErrorRecord | null> {
  return readJson<FeatureErrorRecord>(store, errorCounterKey(feature, torontoDate));
}

// ------------------------------------------------------------- audit trail

/** What the classifier (or the digest) did, or why it did nothing. */
export type AuditEntry = {
  /** When the decision was made (ISO 8601, UTC). */
  at: string;
  /** "filing" for a classified message, "digest" for the morning brief. */
  feature: "filing" | "digest";
  action:
    | "moved"
    | "categorized"
    | "moved+categorized"
    | "drafted"
    | "none"
    /** The reconciler noticed the user re-filed something the filer had moved. */
    | "correction";
  messageId?: string;
  /** Truncated; the audit log is read back by a tool, not a mail client. */
  subject?: string;
  folder?: string;
  categories?: string[];
  confidence?: number;
  /** The model's stated reason, or ours when nothing was done. */
  reason: string;
  model?: string;
  usage?: { input: number; output: number };
  /**
   * What decided a filing action: "llm" when the model was consulted,
   * "preference" when a learned sender preference filed it with NO model call.
   * Absent on entries where neither decided (skips, budget stops, corrections).
   */
  source?: "llm" | "preference";
  /** The sender's bare email address, lowercased — what preferences key on. */
  sender?: string;
  /** The id of the folder a move targeted (displayName alone is ambiguous). */
  folderId?: string;
  /** The message's id after a move (a Graph move mints a new one). */
  newMessageId?: string;
  /**
   * The conversation id, which — unlike message ids — survives moves. The
   * correction reconciler falls back to it when the user's own move has
   * invalidated newMessageId (verified live: the old id 404s after a move).
   */
  conversationId?: string;
  /**
   * Set by the correction reconciler once this move needs no further checks:
   * "confirmed" (still where the filer put it after the confirmation window),
   * "corrected" (the user moved it elsewhere; a preference was learned),
   * "gone" (message unreadable — deleted, or its folder was),
   * "ignored" (it landed somewhere no preference should learn from).
   */
  reconciled?: "confirmed" | "corrected" | "gone" | "ignored";
};

/** The audit ring, newest first. */
export async function readAuditLog(store: StateStore): Promise<AuditEntry[]> {
  return (await readJson<AuditEntry[]>(store, STATE_LLM_AUDIT)) ?? [];
}

/** Prepend one decision and trim to the cap. */
export async function appendAudit(store: StateStore, entry: AuditEntry): Promise<AuditEntry[]> {
  const merged = [entry, ...(await readAuditLog(store))].slice(0, AUDIT_CAP);
  await writeJson(store, STATE_LLM_AUDIT, merged);
  return merged;
}

/** Trim a subject to a length that keeps the audit log compact. */
export function auditSubject(subject: string | undefined): string {
  const text = (subject ?? "").replace(/\s+/g, " ").trim();
  return text.length > 120 ? `${text.slice(0, 117)}...` : text || "(no subject)";
}

// ------------------------------------------- learned filing preferences

/** How many sender preferences are kept. Oldest-touched fall off. */
export const PREFS_CAP = 200;

/** A repeat correction to the same folder makes a preference "standing". */
export const STANDING_CORRECTIONS = 2;

/**
 * One learned rule: mail from `sender` goes to `folderId`. Learned from the
 * user's own corrections (moving a message the auto-filer had filed), consulted
 * BEFORE the model on future arrivals — a hit files with no Anthropic call.
 * A preference whose folder is the Inbox means "leave this sender's mail alone".
 */
export type FilingPreference = {
  /** Bare email address, lowercased. */
  sender: string;
  folderId: string;
  /** Display name at learning time, for humans; the id is what moves act on. */
  folderName: string;
  /** How many times the user has corrected this sender to this folder. */
  corrections: number;
  firstAt: string;
  lastAt: string;
};

/** True once repeat corrections have confirmed the preference. */
export function isStandingPreference(pref: FilingPreference): boolean {
  return pref.corrections >= STANDING_CORRECTIONS;
}

/**
 * The bare address inside a "Name <addr@host>" (or bare) sender string,
 * lowercased — the key preferences are stored under. Null when there is none.
 */
export function extractAddress(from: string | undefined): string | null {
  if (!from) return null;
  // The LAST one: the real address is always appended after the display name,
  // which the sender controls and can make look like "<boss@bank.com>".
  const angled = [...from.matchAll(/<([^<>\s]+@[^<>\s]+)>/g)].at(-1);
  const bare = angled?.[1] ?? (/^[^\s<>]+@[^\s<>]+$/.test(from.trim()) ? from.trim() : null);
  return bare ? bare.toLowerCase() : null;
}

/** Every learned preference, most recently touched first. */
export async function readPreferences(store: StateStore): Promise<FilingPreference[]> {
  const raw = await readJson<FilingPreference[]>(store, STATE_LLM_PREFS);
  if (!Array.isArray(raw)) return [];
  return raw.filter(
    (p): p is FilingPreference =>
      !!p && typeof p.sender === "string" && typeof p.folderId === "string"
  );
}

export async function writePreferences(
  store: StateStore,
  prefs: FilingPreference[]
): Promise<void> {
  await writeJson(store, STATE_LLM_PREFS, prefs.slice(0, PREFS_CAP));
}

/** The preference for one sender address (case-insensitive exact match). */
export function findPreference(
  prefs: FilingPreference[],
  sender: string
): FilingPreference | undefined {
  const needle = sender.trim().toLowerCase();
  return prefs.find((p) => p.sender === needle);
}

/**
 * Record one detected correction: the user moved mail from `sender` into
 * `folder`. Same folder again → the correction count grows (toward "standing");
 * a different folder replaces the preference and the count restarts at 1 —
 * the user's latest choice always wins.
 */
export async function upsertPreferenceFromCorrection(
  store: StateStore,
  sender: string,
  folder: { id: string; name: string },
  when: Date = new Date()
): Promise<FilingPreference> {
  const at = when.toISOString();
  const key = sender.trim().toLowerCase();
  const prefs = await readPreferences(store);
  const existing = findPreference(prefs, key);
  const next: FilingPreference =
    existing && existing.folderId === folder.id
      ? { ...existing, folderName: folder.name, corrections: existing.corrections + 1, lastAt: at }
      : { sender: key, folderId: folder.id, folderName: folder.name, corrections: 1, firstAt: at, lastAt: at };
  await writePreferences(store, [next, ...prefs.filter((p) => p.sender !== key)]);
  return next;
}

/** Remove one sender's preference. Returns whether one existed. */
export async function removePreference(store: StateStore, sender: string): Promise<boolean> {
  const key = sender.trim().toLowerCase();
  const prefs = await readPreferences(store);
  const remaining = prefs.filter((p) => p.sender !== key);
  if (remaining.length === prefs.length) return false;
  await writePreferences(store, remaining);
  return true;
}

// ------------------------------------------------- digest idempotency marker

/** The Toronto date of the last brief that was drafted, or null. */
export async function readLastDigestDate(store: StateStore): Promise<string | null> {
  return store.get(STATE_DIGEST_LAST);
}

export async function writeLastDigestDate(store: StateStore, torontoDate: string): Promise<void> {
  await store.put(STATE_DIGEST_LAST, torontoDate);
}

/**
 * A draft the digest wrote, as it left it. Graph's changeKey moves on any later
 * change to the item, which is how the digest tells its own draft from one the
 * owner has since edited or scheduled.
 */
export type DigestDraft = { id: string; changeKey: string };

/** The draft the digest itself last created or rewrote, or null. */
export async function readLastDigestDraft(store: StateStore): Promise<DigestDraft | null> {
  return readJson<DigestDraft>(store, STATE_DIGEST_DRAFT);
}

export async function writeLastDigestDraft(store: StateStore, draft: DigestDraft): Promise<void> {
  await writeJson(store, STATE_DIGEST_DRAFT, draft);
}
