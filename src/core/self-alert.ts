// The self-alert route and its heartbeat watchdog: the one documented
// exception to "send_draft is the only send path".
//
// A scheduled job the owner runs elsewhere (a CI workflow, say) needs two
// things only a mailbox can give it: a way to email the owner when something
// important happens, and someone to notice when the job itself stops running.
// Both are served here, fenced so the exception cannot widen:
//
//  - Off by default. Unless SELF_ALERT_SECRET is set (at least 32 characters)
//    both routes answer 404, exactly like a path that does not exist, and the
//    watchdog does nothing.
//  - Shared-secret gated. `Authorization: Bearer <SELF_ALERT_SECRET>`, compared
//    in constant time; anything else gets a bare 401.
//  - One fixed recipient. Mail goes to ALLOWED_MS_UPN — the owner's own
//    address, already the connector's single-user allowlist — and nowhere else.
//    The body schema has no recipient, cc, bcc, from, reply-to or attachment
//    field and any field outside it is refused with 400, so a caller has
//    nowhere to put a second address. buildSelfAlertMail is the only place the
//    Graph message is shaped.
//  - Capped. At most SELF_ALERT_DAILY_CAP mails per UTC day, watchdog alerts
//    included; over the cap the route answers 429 and nothing is sent.
//
// The watchdog: POST /self-alert/heartbeat records {at, max_age_hours} for a
// source+job, and that record is the job's registration — a job that never
// sent a heartbeat is never watched. The watchdog rides the subscription-upkeep
// cron (UPKEEP_CRON in core/schedule.js, every 6 hours), lists the records and
// alerts once per stale episode on any job silent for longer than it asked;
// the job's next heartbeat clears the flag.
//
// Transport-free: the KV namespace, the clock and the Graph send are injected,
// so every branch is tested offline; src/worker/self-alert.ts supplies the
// real dependencies.
import {
  SELF_ALERT_HEARTBEAT_PREFIX,
  selfAlertCountKey,
  selfAlertHeartbeatKey,
  selfAlertStaleKey,
} from "./kv-keys.js";

export const SELF_ALERT_PATH = "/self-alert";
export const SELF_ALERT_HEARTBEAT_PATH = "/self-alert/heartbeat";

/** Mails per UTC day, the route and the watchdog together. */
export const SELF_ALERT_DAILY_CAP = 20;

/** Request bodies larger than this are refused with 413 before parsing. */
export const SELF_ALERT_MAX_BODY_BYTES = 64 * 1024;

/** A shorter SELF_ALERT_SECRET is treated as unset: the feature stays off. */
export const SELF_ALERT_MIN_SECRET_LENGTH = 32;

/** The source the watchdog's own alerts are sent under. */
export const SELF_ALERT_WATCHDOG_SOURCE = "self-alert-watchdog";

const SUBJECT_MAX = 200;
const TEXT_MAX = 50_000;
const NAME_PATTERN = /^[a-z0-9-]{1,40}$/;
const CONTROL_CHARS = /[\u0000-\u001f\u007f]/;
const MAX_AGE_HOURS_MIN = 1;
const MAX_AGE_HOURS_MAX = 720;
/** Two days: the counter outlives its UTC day briefly, then cleans itself up. */
const COUNTER_TTL_SECONDS = 2 * 24 * 3600;
const HOUR_MS = 3_600_000;
/** One bare address — no list separators, no display-name syntax. */
const SINGLE_ADDRESS = /^[^\s@,;<>"()]+@[^\s@,;<>"()]+$/;

const ALERT_FIELDS = ["subject", "text", "source"] as const;
const HEARTBEAT_FIELDS = ["source", "job", "max_age_hours"] as const;

/** The slice of a KV namespace this feature uses; OUTLOOK_KV satisfies it as is. */
export interface SelfAlertKv {
  get(key: string): Promise<string | null>;
  put(key: string, value: string, options?: { expirationTtl?: number }): Promise<void>;
  delete(key: string): Promise<void>;
  list(options: {
    prefix: string;
    cursor?: string;
  }): Promise<{ keys: { name: string }[]; list_complete: boolean; cursor?: string }>;
}

/** The body of Graph's POST /me/sendMail, exactly as this feature sends it. */
export type SendMailPayload = {
  message: {
    subject: string;
    body: { contentType: "Text"; content: string };
    toRecipients: [{ emailAddress: { address: string } }];
  };
  saveToSentItems: true;
};

export type SelfAlertDeps = {
  /** SELF_ALERT_SECRET. Unset, empty or too short means the feature is off. */
  secret: string | undefined;
  /** ALLOWED_MS_UPN: the only address a self-alert can ever reach. */
  recipient: string | undefined;
  kv: SelfAlertKv;
  /** Deliver one message through Graph; must throw when Graph refuses. */
  sendMail: (payload: SendMailPayload) => Promise<void>;
  now?: () => Date;
};

export type SelfAlert = { subject: string; text: string; source: string };
export type Heartbeat = { source: string; job: string; max_age_hours: number };
export type HeartbeatRecord = { at: string; max_age_hours: number };
/** Bound to the heartbeat it alerted on, so a flag can never outlive its episode. */
type StaleFlag = { alerted_at: string; heartbeat_at: string };

type Validation<T> = { ok: true; value: T } | { ok: false; error: string };

// ------------------------------------------------------------- the gate

/** The configured secret when the feature is on, otherwise undefined. */
export function activeSecret(secret: string | undefined): string | undefined {
  const trimmed = secret?.trim();
  if (!trimmed) return undefined;
  if (trimmed.length < SELF_ALERT_MIN_SECRET_LENGTH) {
    console.error(
      `SELF_ALERT_SECRET is shorter than ${SELF_ALERT_MIN_SECRET_LENGTH} characters, ` +
        "so the self-alert routes and watchdog stay off."
    );
    return undefined;
  }
  return trimmed;
}

export function isSelfAlertPath(pathname: string): boolean {
  return pathname === SELF_ALERT_PATH || pathname === SELF_ALERT_HEARTBEAT_PATH;
}

/**
 * Constant-time string equality. Both sides are hashed first, so the compared
 * buffers always have the same length (no length oracle) and the loop has no
 * early exit.
 */
export async function secretsMatch(presented: string, expected: string): Promise<boolean> {
  const encoder = new TextEncoder();
  const [a, b] = await Promise.all([
    crypto.subtle.digest("SHA-256", encoder.encode(presented)),
    crypto.subtle.digest("SHA-256", encoder.encode(expected)),
  ]);
  const x = new Uint8Array(a);
  const y = new Uint8Array(b);
  let diff = x.length ^ y.length;
  for (let i = 0; i < x.length; i++) diff |= x[i]! ^ y[i]!;
  return diff === 0;
}

async function authorized(request: Request, secret: string): Promise<boolean> {
  const match = /^Bearer +(.+)$/i.exec((request.headers.get("authorization") ?? "").trim());
  const presented = match?.[1]?.trim() ?? "";
  // Compared even when nothing was presented, so both refusals cost the same.
  const same = await secretsMatch(presented, secret);
  return presented.length > 0 && same;
}

// ------------------------------------------------------------ validation

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Exactly `allowed`, all present: anything extra is refused, never ignored. */
function checkFields(body: unknown, allowed: readonly string[]): string | null {
  if (!isPlainObject(body)) return "the body must be a JSON object";
  const extra = Object.keys(body).find((key) => !allowed.includes(key));
  if (extra !== undefined) {
    return `unknown field "${extra.slice(0, 40)}"; only ${allowed.join(", ")} are accepted`;
  }
  const missing = allowed.find((key) => !Object.prototype.hasOwnProperty.call(body, key));
  if (missing !== undefined) return `missing field "${missing}"`;
  return null;
}

function checkName(field: string, value: unknown): string | null {
  return typeof value === "string" && NAME_PATTERN.test(value)
    ? null
    : `${field} must be 1-40 characters of a-z, 0-9 and -`;
}

export function validateAlert(body: unknown): Validation<SelfAlert> {
  const shape = checkFields(body, ALERT_FIELDS);
  if (shape) return { ok: false, error: shape };
  const { subject, text, source } = body as Record<string, unknown>;
  const sourceError = checkName("source", source);
  if (sourceError) return { ok: false, error: sourceError };
  if (
    typeof subject !== "string" ||
    subject.length < 1 ||
    subject.length > SUBJECT_MAX ||
    subject.trim().length === 0 ||
    CONTROL_CHARS.test(subject)
  ) {
    return {
      ok: false,
      error: `subject must be 1-${SUBJECT_MAX} characters, not blank, on one line`,
    };
  }
  if (typeof text !== "string" || text.length < 1 || text.length > TEXT_MAX) {
    return { ok: false, error: `text must be 1-${TEXT_MAX} characters` };
  }
  return { ok: true, value: { subject, text, source: source as string } };
}

export function validateHeartbeat(body: unknown): Validation<Heartbeat> {
  const shape = checkFields(body, HEARTBEAT_FIELDS);
  if (shape) return { ok: false, error: shape };
  const { source, job, max_age_hours } = body as Record<string, unknown>;
  const nameError = checkName("source", source) ?? checkName("job", job);
  if (nameError) return { ok: false, error: nameError };
  if (
    typeof max_age_hours !== "number" ||
    !Number.isFinite(max_age_hours) ||
    max_age_hours < MAX_AGE_HOURS_MIN ||
    max_age_hours > MAX_AGE_HOURS_MAX
  ) {
    return {
      ok: false,
      error: `max_age_hours must be a number from ${MAX_AGE_HOURS_MIN} to ${MAX_AGE_HOURS_MAX}`,
    };
  }
  return {
    ok: true,
    value: { source: source as string, job: job as string, max_age_hours },
  };
}

// ---------------------------------------------------------- sending, capped

/** The Graph sendMail body. The recipient is the configured owner, always. */
export function buildSelfAlertMail(recipient: string, alert: SelfAlert): SendMailPayload {
  return {
    message: {
      subject: `[${alert.source}] ${alert.subject}`,
      body: { contentType: "Text", content: alert.text },
      toRecipients: [{ emailAddress: { address: recipient } }],
    },
    saveToSentItems: true,
  };
}

function utcDateOf(when: Date): string {
  return when.toISOString().slice(0, 10);
}

function secondsUntilNextUtcDay(now: Date): number {
  const next = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate() + 1);
  return Math.max(1, Math.ceil((next - now.getTime()) / 1000));
}

/**
 * Count one send against today's cap, or refuse. Read-modify-write on KV, so
 * two sends racing at the boundary could both pass — the same trade the LLM
 * call budget makes; for a single owner's alerts the ceiling is what matters.
 * A reserved send that Graph then refuses still counts.
 */
async function reserveSend(kv: SelfAlertKv, now: Date): Promise<{ allowed: boolean; used: number }> {
  const key = selfAlertCountKey(utcDateOf(now));
  const used = Number((await kv.get(key)) ?? "0") || 0;
  if (used >= SELF_ALERT_DAILY_CAP) return { allowed: false, used };
  await kv.put(key, String(used + 1), { expirationTtl: COUNTER_TTL_SECONDS });
  return { allowed: true, used: used + 1 };
}

type SendOutcome =
  | { ok: true; sentToday: number }
  | { ok: false; reason: "no_recipient" | "cap" | "send_failed"; detail: string };

function errorText(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

async function sendSelfAlert(deps: SelfAlertDeps, alert: SelfAlert, now: Date): Promise<SendOutcome> {
  const recipient = deps.recipient?.trim();
  if (!recipient || !SINGLE_ADDRESS.test(recipient)) {
    return {
      ok: false,
      reason: "no_recipient",
      detail: "ALLOWED_MS_UPN is not set to a single address",
    };
  }
  const reservation = await reserveSend(deps.kv, now);
  if (!reservation.allowed) {
    return {
      ok: false,
      reason: "cap",
      detail: `the daily cap of ${SELF_ALERT_DAILY_CAP} self-alerts is spent (UTC ${utcDateOf(now)})`,
    };
  }
  try {
    await deps.sendMail(buildSelfAlertMail(recipient, alert));
  } catch (err) {
    return { ok: false, reason: "send_failed", detail: errorText(err) };
  }
  return { ok: true, sentToday: reservation.used };
}

// --------------------------------------------------------------- the route

function json(status: number, body: unknown, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json", ...headers },
  });
}

/** Byte for byte the Worker's own 404, so a disabled route looks absent. */
function notFound(): Response {
  return new Response("Not found", { status: 404 });
}

/** The body as text, or null once it passes `limit` bytes (read no further). */
async function readBodyLimited(request: Request, limit: number): Promise<string | null> {
  const declared = Number(request.headers.get("content-length"));
  if (Number.isFinite(declared) && declared > limit) return null;
  if (!request.body) return "";
  const reader = request.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > limit) {
      await reader.cancel().catch(() => {});
      return null;
    }
    chunks.push(value);
  }
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  // fatal: invalid UTF-8 throws, and the caller answers 400.
  return new TextDecoder("utf-8", { fatal: true, ignoreBOM: false }).decode(bytes);
}

async function readJsonBody(
  request: Request
): Promise<{ ok: true; value: unknown } | { ok: false; status: 400 | 413; error: string }> {
  try {
    const text = await readBodyLimited(request, SELF_ALERT_MAX_BODY_BYTES);
    if (text === null) {
      return { ok: false, status: 413, error: `the body exceeds ${SELF_ALERT_MAX_BODY_BYTES} bytes` };
    }
    return { ok: true, value: JSON.parse(text) };
  } catch {
    return { ok: false, status: 400, error: "the body is not valid JSON" };
  }
}

/**
 * POST /self-alert and POST /self-alert/heartbeat. Order of refusal: feature
 * off (404), wrong method (405), bad credential (401), oversize (413),
 * malformed (400), over the cap (429). Nothing is parsed before the secret
 * checks out, and nothing is sent unless every check passed.
 */
export async function handleSelfAlertRequest(
  request: Request,
  deps: SelfAlertDeps
): Promise<Response> {
  const { pathname } = new URL(request.url);
  const secret = activeSecret(deps.secret);
  if (!secret || !isSelfAlertPath(pathname)) return notFound();
  if (request.method !== "POST") {
    return json(405, { ok: false, error: "method not allowed" }, { allow: "POST" });
  }
  if (!(await authorized(request, secret))) {
    return new Response("Unauthorized", {
      status: 401,
      headers: { "www-authenticate": "Bearer" },
    });
  }

  const body = await readJsonBody(request);
  if (!body.ok) return json(body.status, { ok: false, error: body.error });
  const now = (deps.now ?? (() => new Date()))();

  try {
    return pathname === SELF_ALERT_HEARTBEAT_PATH
      ? await recordHeartbeat(deps.kv, body.value, now)
      : await deliverAlert(deps, body.value, now);
  } catch (err) {
    console.error(`Self-alert route failed: ${errorText(err)}`);
    return json(500, { ok: false, error: "internal error" });
  }
}

async function deliverAlert(deps: SelfAlertDeps, body: unknown, now: Date): Promise<Response> {
  const parsed = validateAlert(body);
  if (!parsed.ok) return json(400, { ok: false, error: parsed.error });
  const alert = parsed.value;

  const outcome = await sendSelfAlert(deps, alert, now);
  if (outcome.ok) {
    console.log(
      `Self-alert sent for source "${alert.source}" (${outcome.sentToday}/${SELF_ALERT_DAILY_CAP} today).`
    );
    return json(202, { ok: true });
  }
  switch (outcome.reason) {
    case "cap":
      console.error(`Self-alert refused for source "${alert.source}": ${outcome.detail}.`);
      return json(
        429,
        { ok: false, error: `daily cap of ${SELF_ALERT_DAILY_CAP} reached; nothing was sent` },
        { "retry-after": String(secondsUntilNextUtcDay(now)) }
      );
    case "no_recipient":
      console.error(`Self-alert refused: ${outcome.detail}.`);
      return json(503, { ok: false, error: "the self-alert recipient is not configured" });
    case "send_failed":
      // Logged in full for `wrangler tail`; Graph's answer is never echoed.
      console.error(`Self-alert send failed for source "${alert.source}": ${outcome.detail}`);
      return json(502, { ok: false, error: "the mail could not be sent" });
  }
}

async function recordHeartbeat(kv: SelfAlertKv, body: unknown, now: Date): Promise<Response> {
  const parsed = validateHeartbeat(body);
  if (!parsed.ok) return json(400, { ok: false, error: parsed.error });
  const { source, job, max_age_hours } = parsed.value;
  const record: HeartbeatRecord = { at: now.toISOString(), max_age_hours };
  await kv.put(selfAlertHeartbeatKey(source, job), JSON.stringify(record));
  await kv.delete(selfAlertStaleKey(source, job));
  return new Response(null, { status: 204 });
}

// ------------------------------------------------------------ the watchdog

function parseJson(raw: string | null): unknown {
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function asHeartbeatRecord(value: unknown): HeartbeatRecord | null {
  if (!isPlainObject(value)) return null;
  const { at, max_age_hours } = value;
  if (typeof at !== "string" || !Number.isFinite(Date.parse(at))) return null;
  if (typeof max_age_hours !== "number" || !Number.isFinite(max_age_hours)) return null;
  return { at, max_age_hours };
}

/** Stale once strictly more than max_age_hours have passed since the heartbeat. */
export function isHeartbeatStale(record: HeartbeatRecord, now: Date): boolean {
  return now.getTime() - Date.parse(record.at) > record.max_age_hours * HOUR_MS;
}

/** The alert the watchdog sends for one stale job. */
export function watchdogAlert(
  source: string,
  job: string,
  record: HeartbeatRecord,
  now: Date
): SelfAlert {
  const hours = Math.floor((now.getTime() - Date.parse(record.at)) / HOUR_MS);
  return {
    source: SELF_ALERT_WATCHDOG_SOURCE,
    subject: `${source}/${job}: no heartbeat for ${hours}h`,
    text: [
      `No heartbeat from job "${job}" (source "${source}") for ${hours} hours; ` +
        `it asked to be flagged after ${record.max_age_hours} hours of silence.`,
      "",
      `Last heartbeat: ${record.at}`,
      `Checked at:     ${now.toISOString()}`,
      "",
      "This alert is sent once per lapse: the job's next heartbeat clears it, and a later " +
        "lapse alerts again. To stop watching a retired job, delete the OUTLOOK_KV key " +
        `${selfAlertHeartbeatKey(source, job)}.`,
      "",
      "—",
      "Sent by the self-alert watchdog on the hosted outlook-mcp server, to its owner's own " +
        "address only.",
    ].join("\n"),
  };
}

export type WatchdogResult =
  | { enabled: false }
  | {
      enabled: true;
      /** Registered jobs with a readable heartbeat record. */
      watched: number;
      /** Of those, how many are past their max_age_hours right now. */
      stale: number;
      /** "source/job" for each alert sent on this run. */
      alerted: string[];
      /** Anything that went wrong; an unsent alert is retried on the next tick. */
      problems: string[];
    };

async function listHeartbeatKeys(kv: SelfAlertKv): Promise<string[]> {
  const names: string[] = [];
  let cursor: string | undefined;
  do {
    const page = await kv.list({
      prefix: SELF_ALERT_HEARTBEAT_PREFIX,
      ...(cursor ? { cursor } : {}),
    });
    names.push(...page.keys.map((key) => key.name));
    cursor = page.list_complete ? undefined : page.cursor;
  } while (cursor);
  return names;
}

/**
 * One watchdog pass, run on every UPKEEP_CRON tick (every 6 hours). Alerts at
 * most once per stale episode: the flag records which heartbeat it alerted on,
 * the next heartbeat deletes it, and a failed or over-cap send sets no flag, so
 * it is retried on the next tick.
 * Per-job failures are collected, not thrown; a KV list failure does throw,
 * and the scheduled handler logs it.
 */
export async function runSelfAlertWatchdog(deps: SelfAlertDeps): Promise<WatchdogResult> {
  if (!activeSecret(deps.secret)) return { enabled: false };
  const now = (deps.now ?? (() => new Date()))();
  const result = {
    enabled: true as const,
    watched: 0,
    stale: 0,
    alerted: [] as string[],
    problems: [] as string[],
  };

  for (const name of await listHeartbeatKeys(deps.kv)) {
    const [source, job, ...rest] = name.slice(SELF_ALERT_HEARTBEAT_PREFIX.length).split(":");
    if (!source || !job || rest.length > 0 || checkName("source", source) || checkName("job", job)) {
      result.problems.push(`unrecognised heartbeat key ${name}`);
      continue;
    }
    const label = `${source}/${job}`;
    try {
      const record = asHeartbeatRecord(parseJson(await deps.kv.get(name)));
      if (!record) {
        result.problems.push(`${label}: unreadable heartbeat record`);
        continue;
      }
      result.watched++;
      if (!isHeartbeatStale(record, now)) continue;
      result.stale++;

      const flagKey = selfAlertStaleKey(source, job);
      const flag = parseJson(await deps.kv.get(flagKey)) as Partial<StaleFlag> | null;
      if (flag?.heartbeat_at === record.at) continue; // already alerted on this lapse

      const outcome = await sendSelfAlert(deps, watchdogAlert(source, job, record, now), now);
      if (!outcome.ok) {
        result.problems.push(`${label}: alert not sent — ${outcome.detail}`);
        continue;
      }
      const raised: StaleFlag = { alerted_at: now.toISOString(), heartbeat_at: record.at };
      await deps.kv.put(flagKey, JSON.stringify(raised));
      result.alerted.push(label);
    } catch (err) {
      result.problems.push(`${label}: ${errorText(err)}`);
    }
  }
  return result;
}
