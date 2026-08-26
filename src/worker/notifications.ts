// Worker wiring for Microsoft Graph change notifications: the public endpoint
// Graph posts to, and the subscription upkeep the cron trigger drives. The
// logic itself lives in core/notifications.js and core/subscriptions.js, which
// know nothing about Workers and are unit-tested from the Node harness.
//
// The endpoint is deliberately outside the OAuth-protected /mcp route: Graph
// presents no credential, so the shared clientState secret (generated at
// subscription creation and stored in KV) is what authenticates a delivery.
import { handleNotificationRequest, type MessageEnricher } from "../core/notifications.js";
import { classifyNotified } from "./llm.js";
import { runWithTokenProvider } from "../core/token.js";
import { ensureMailSubscription, type EnsureResult } from "../core/subscriptions.js";
import type { Env } from "./env.js";
import { getMailboxAccessToken, mailboxTokenProvider } from "./ms-token.js";
import { kvStateStore } from "./state-kv.js";

/** The path Graph is told to deliver to. Part of the subscription record. */
export const NOTIFICATIONS_PATH = "/notifications";

/** The deployed origin, with any trailing slashes removed. */
export function publicBaseUrl(env: Env): string | undefined {
  const configured = env.PUBLIC_BASE_URL?.trim();
  return configured ? configured.replace(/\/+$/, "") : undefined;
}

/**
 * The absolute URL Graph must post notifications to.
 *
 * There is deliberately no compiled-in fallback: the origin belongs to the
 * deployment, not to the source, so it comes from the PUBLIC_BASE_URL wrangler
 * var and nowhere else. A deploy that forgot it cannot guess its own hostname,
 * and a guess would silently point Graph at someone else's Worker — so this
 * throws, and the scheduled handler logs it.
 */
export function notificationUrl(env: Env): string {
  const base = publicBaseUrl(env);
  if (!base) {
    throw new Error(
      "PUBLIC_BASE_URL is not set, so Graph cannot be told where to deliver change " +
        "notifications. Set it in wrangler.jsonc to this Worker's public origin and redeploy."
    );
  }
  return base + NOTIFICATIONS_PATH;
}

const GRAPH_MESSAGE = "https://graph.microsoft.com/v1.0/me/messages";

/**
 * Turn a notified message id into something a human can read. Uses the KV
 * mailbox token directly rather than the tool layer's provider: this runs
 * outside any MCP request, and a failure here must cost nothing but detail.
 */
function messageEnricher(env: Env): MessageEnricher {
  return async (messageId) => {
    const token = await getMailboxAccessToken(env);
    const response = await fetch(
      `${GRAPH_MESSAGE}/${encodeURIComponent(messageId)}?$select=subject,from,receivedDateTime`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!response.ok) return null;
    const message = (await response.json()) as any;
    const address = message?.from?.emailAddress;
    let from: string | undefined;
    if (address?.address) {
      from =
        address.name && address.name !== address.address
          ? `${address.name} <${address.address}>`
          : address.address;
    }
    return {
      subject: message?.subject ?? undefined,
      from,
      receivedDateTime: message?.receivedDateTime ?? undefined,
    };
  };
}

/**
 * The public /notifications route: validation handshake and deliveries.
 *
 * Anything a delivery triggers beyond recording it — today, LLM classification
 * of the messages it named — is handed to ctx.waitUntil so Graph still gets its
 * 202 immediately. Auto-filing ships disabled, so on a default deployment
 * classifyNotified reads one KV key and returns.
 */
export function handleNotifications(
  request: Request,
  env: Env,
  ctx: ExecutionContext
): Promise<Response> {
  return handleNotificationRequest(request, {
    store: kvStateStore(env),
    enrich: messageEnricher(env),
    onAccepted: (entries) =>
      ctx.waitUntil(
        classifyNotified(env, entries).catch((err) =>
          console.error(`Auto-filing failed for a notification delivery: ${String(err)}`)
        )
      ),
  });
}

/**
 * Create or renew the inbox subscription. Driven by the cron trigger, and also
 * (in the background) by any authenticated MCP request, so the subscription
 * heals itself the moment the connector is used rather than waiting for the
 * next scheduled run.
 */
export async function keepSubscriptionAlive(env: Env): Promise<EnsureResult> {
  const result = await runWithTokenProvider(mailboxTokenProvider(env), () =>
    ensureMailSubscription(kvStateStore(env), notificationUrl(env))
  );
  if (result.action !== "kept") {
    console.log(`Mail subscription ${result.action}: ${result.detail}`);
  }
  return result;
}
