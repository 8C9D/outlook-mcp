// The Anthropic Messages API over plain fetch.
//
// Deliberately not the @anthropic-ai/sdk: this module is bundled into the
// Cloudflare Worker, one endpoint is all either LLM feature needs, and adding a
// dependency to workerd for a single POST buys nothing. Like core/graph.js it
// is free of Node-only imports so both hosts can load it.
//
// The API key never appears in a thrown message, a log line or an audit entry —
// callers surface `AnthropicError.message`, which carries only the HTTP status
// and Anthropic's own error text.

/** Messages API endpoint. */
const MESSAGES_URL = "https://api.anthropic.com/v1/messages";

/** The API version header every request must carry. */
const API_VERSION = "2023-06-01";

/**
 * The model both features use: the cheapest current Haiku, chosen because
 * classification and a one-paragraph brief are exactly what it is good at and
 * the workload is per-message. Input $1 / output $5 per million tokens.
 *
 * Recorded here as the canonical alias rather than a dated snapshot id — see
 * ASSUMPTIONS.md (Batch C) for the live verification.
 */
export const LLM_MODEL = "claude-haiku-4-5";

/** A non-2xx answer from the Anthropic API. Never carries the API key. */
export class AnthropicError extends Error {
  constructor(
    public readonly status: number,
    public readonly detail: string
  ) {
    super(`Anthropic API error (HTTP ${status}): ${detail}`);
    this.name = "AnthropicError";
  }
}

/** Thrown when no ANTHROPIC_API_KEY is configured on this host. */
export class AnthropicKeyMissingError extends Error {
  constructor() {
    super(
      "No ANTHROPIC_API_KEY is configured on this server, so the LLM features cannot run. " +
        "Set it with `npx wrangler secret put ANTHROPIC_API_KEY`."
    );
    this.name = "AnthropicKeyMissingError";
  }
}

export type AnthropicRequest = {
  apiKey: string;
  /** Defaults to LLM_MODEL. */
  model?: string;
  /** Trusted instructions. Never contains mail content. */
  system: string;
  /** The untrusted payload, already delimited and truncated by the caller. */
  user: string;
  /** Hard ceiling on generated tokens. Both callers keep this small on purpose. */
  maxTokens: number;
};

export type AnthropicReply = {
  text: string;
  model: string;
  usage: { input: number; output: number };
  stopReason: string | null;
};

/** How long to wait when Anthropic answers 429 without a usable Retry-After. */
const DEFAULT_RETRY_SECONDS = 2;

/** Never wait longer than this for a single retry; the caller has a deadline. */
const MAX_RETRY_SECONDS = 30;

/**
 * One Messages API call. Retries once on 429 or 5xx (honouring Retry-After),
 * then gives up: both callers run in the background off a webhook or a cron and
 * a failure costs nothing but the decision itself.
 */
export async function callAnthropic(request: AnthropicRequest): Promise<AnthropicReply> {
  if (!request.apiKey) throw new AnthropicKeyMissingError();
  const model = request.model ?? LLM_MODEL;

  const body = JSON.stringify({
    model,
    max_tokens: request.maxTokens,
    system: request.system,
    messages: [{ role: "user", content: request.user }],
  });

  const doFetch = () =>
    fetch(MESSAGES_URL, {
      method: "POST",
      headers: {
        "x-api-key": request.apiKey,
        "anthropic-version": API_VERSION,
        "content-type": "application/json",
      },
      body,
    });

  let response = await doFetch();
  if (response.status === 429 || response.status >= 500) {
    const header = Number(response.headers.get("retry-after") ?? "");
    const wait = Number.isFinite(header) && header > 0 ? Math.min(header, MAX_RETRY_SECONDS) : DEFAULT_RETRY_SECONDS;
    console.error(`Anthropic answered ${response.status}; retrying once after ${wait}s.`);
    await new Promise((resolve) => setTimeout(resolve, wait * 1000));
    response = await doFetch();
  }

  if (!response.ok) {
    throw new AnthropicError(response.status, await shortErrorText(response));
  }

  const payload = (await response.json()) as any;
  const text = (payload?.content ?? [])
    .filter((block: any) => block?.type === "text")
    .map((block: any) => String(block.text ?? ""))
    .join("");

  return {
    text,
    model: String(payload?.model ?? model),
    usage: {
      input: Number(payload?.usage?.input_tokens ?? 0),
      output: Number(payload?.usage?.output_tokens ?? 0),
    },
    stopReason: payload?.stop_reason ?? null,
  };
}

/** Anthropic's error message, trimmed to something safe to store in an audit log. */
async function shortErrorText(response: Response): Promise<string> {
  const raw = await response.text().catch(() => "");
  try {
    const parsed = JSON.parse(raw);
    const type = parsed?.error?.type;
    const message = parsed?.error?.message;
    if (message) return `${type ?? "error"}: ${String(message).slice(0, 300)}`;
  } catch {
    // fall through to the raw body
  }
  return raw.slice(0, 300) || response.statusText;
}
