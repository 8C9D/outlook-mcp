// Receiving Microsoft Graph change notifications.
//
// Two things arrive at the same public URL:
//
//  1. The validation handshake. When a subscription is created (or renewed to a
//     new URL) Graph POSTs with a `validationToken` query parameter and expects
//     that exact string echoed back as text/plain within 10 seconds, before any
//     notification is ever sent. It carries no clientState and must be answered
//     without consulting any state.
//  2. Deliveries: a JSON body of notifications. The endpoint is unauthenticated
//     by necessity — Graph will not present a credential — so the clientState
//     secret stored with the subscription is the only proof a delivery is real,
//     and it is checked on every item. Anything that fails the check is dropped.
//
// Graph retries anything that is not answered quickly with 2xx, and duplicate
// deliveries are normal, so this always answers 202 and never blocks on work it
// could avoid.
//
// The received notifications land in a capped ring buffer so a Claude surface
// can ask "what arrived since this morning" from stored state instead of
// polling Graph. Buffer writes are read-modify-write: for one mailbox with a
// human volume of mail the lost-update window is not worth a lock.
import { STATE_ACTIVITY, STATE_SUBSCRIPTION } from "./kv-keys.js";
import { readJson, writeJson, type StateStore } from "./state.js";
import type { SubscriptionRecord } from "./subscriptions.js";

/** How many notifications the ring buffer keeps. Older entries fall off. */
export const ACTIVITY_CAP = 50;

/** At most this many messages per delivery get a subject/sender lookup. */
const ENRICH_PER_DELIVERY = 5;

/** One received notification, as get_mailbox_activity reports it. */
export type ActivityEntry = {
  /** When this server received the notification (ISO 8601, UTC). */
  at: string;
  changeType: string;
  messageId?: string;
  subject?: string;
  from?: string;
  receivedDateTime?: string;
};

type GraphNotification = {
  subscriptionId?: string;
  clientState?: string;
  changeType?: string;
  resource?: string;
  resourceData?: { id?: string };
};

/** Looks a message up for display. Best effort: null means "could not". */
export type MessageEnricher = (
  messageId: string
) => Promise<{ subject?: string; from?: string; receivedDateTime?: string } | null>;

export type NotificationContext = {
  store: StateStore;
  enrich?: MessageEnricher;
  now?: () => Date;
  /**
   * Called once with the notifications that passed the clientState check, after
   * they are recorded and before the 202 goes back. Graph retries anything that
   * is not answered promptly, so an implementation must hand the work off (the
   * Worker uses ctx.waitUntil) rather than await it. Never awaited here, and a
   * throw is swallowed: follow-up work must not cost us the acknowledgement.
   */
  onAccepted?: (entries: ActivityEntry[]) => void;
};

function textResponse(body: string, status: number): Response {
  return new Response(body, {
    status,
    headers: { "content-type": "text/plain; charset=utf-8" },
  });
}

/** The ring buffer, newest first. */
export async function readActivity(store: StateStore): Promise<ActivityEntry[]> {
  return (await readJson<ActivityEntry[]>(store, STATE_ACTIVITY)) ?? [];
}

/** Prepend `entries` (newest first among themselves) and trim to the cap. */
export async function appendActivity(
  store: StateStore,
  entries: ActivityEntry[]
): Promise<ActivityEntry[]> {
  if (entries.length === 0) return readActivity(store);
  const merged = [...entries, ...(await readActivity(store))].slice(0, ACTIVITY_CAP);
  await writeJson(store, STATE_ACTIVITY, merged);
  return merged;
}

/**
 * Handle one request to the notification endpoint: the validation handshake,
 * or a batch of deliveries. Returns the response to send back to Graph.
 */
export async function handleNotificationRequest(
  request: Request,
  context: NotificationContext
): Promise<Response> {
  const url = new URL(request.url);
  const validationToken = url.searchParams.get("validationToken");
  if (validationToken !== null) {
    // Echoed verbatim and nothing else: Graph compares byte for byte.
    return textResponse(validationToken, 200);
  }

  if (request.method !== "POST") {
    return textResponse("Only POST is accepted here.", 405);
  }

  let payload: { value?: GraphNotification[] } | null = null;
  try {
    payload = (await request.json()) as { value?: GraphNotification[] };
  } catch {
    return textResponse("Malformed notification body.", 400);
  }

  const notifications = Array.isArray(payload?.value) ? payload!.value! : [];
  const subscription = await readJson<SubscriptionRecord>(context.store, STATE_SUBSCRIPTION);
  const now = context.now ?? (() => new Date());

  const accepted: ActivityEntry[] = [];
  let rejected = 0;
  for (const notification of notifications) {
    // Constant-time comparison buys nothing here: an attacker cannot see the
    // difference between "wrong secret" and "right secret" in the response.
    const valid =
      !!subscription?.clientState && notification.clientState === subscription.clientState;
    if (!valid) {
      rejected++;
      continue;
    }
    accepted.push({
      at: now().toISOString(),
      changeType: notification.changeType ?? "unknown",
      messageId: notification.resourceData?.id,
    });
  }

  if (context.enrich) {
    for (const entry of accepted.slice(0, ENRICH_PER_DELIVERY)) {
      if (!entry.messageId) continue;
      try {
        const details = await context.enrich(entry.messageId);
        if (details) Object.assign(entry, details);
      } catch {
        // A lookup failure must not cost us the notification itself.
      }
    }
  }

  // Newest first inside this delivery too, so the buffer stays ordered.
  await appendActivity(context.store, accepted.reverse());

  if (context.onAccepted && accepted.length > 0) {
    try {
      context.onAccepted(accepted);
    } catch (err) {
      console.error(`Notification follow-up could not be scheduled: ${String(err)}`);
    }
  }

  if (rejected > 0) {
    console.error(
      `Change notification endpoint discarded ${rejected} delivery item(s) with a bad clientState.`
    );
  }
  // 202 either way: telling an unauthenticated caller which secret was wrong
  // would be a probing oracle, and Graph would retry a non-2xx forever.
  return new Response(JSON.stringify({ accepted: accepted.length, discarded: rejected }), {
    status: 202,
    headers: { "content-type": "application/json" },
  });
}
