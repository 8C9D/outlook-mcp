// The Microsoft Graph change-notification subscription for the inbox.
//
// Graph will only push message-created notifications to a publicly reachable
// URL, so this is a remote-mode feature: the Worker owns the subscription and
// the cron trigger keeps it alive. Nothing here imports Worker types — the
// caller supplies the state store and (in tests) the Graph transport — so the
// renewal logic can be exercised directly from the Node harness.
//
// Mail subscriptions expire in at most 4230 minutes (~2.9 days) and Graph
// deletes them silently once lapsed, so "renew well before expiry" and
// "re-create when it is gone" are both load-bearing.
import { callGraphServer } from "./graph.js";
import { STATE_SUBSCRIPTION } from "./kv-keys.js";
import { readJson, writeJson, type StateStore } from "./state.js";

/** Only the inbox is watched; that is where "what just arrived" means anything. */
export const SUBSCRIPTION_RESOURCE = "/me/mailFolders('inbox')/messages";

/** Graph's own ceiling for message subscriptions. */
export const MAX_EXPIRY_MINUTES = 4230;

/** Ask for slightly less than the ceiling so clock skew can never be rejected. */
export const REQUESTED_EXPIRY_MINUTES = 4200;

/** Renew once less than this much life is left. The cron runs far more often. */
export const RENEW_WITHIN_MINUTES = 24 * 60;

/** What is stored about the live subscription. The clientState is a secret. */
export type SubscriptionRecord = {
  id: string;
  /** Random per subscription; every delivery must echo it or it is discarded. */
  clientState: string;
  expirationDateTime: string;
  notificationUrl: string;
  resource: string;
  createdAt: string;
  renewedAt?: string;
};

/** What ensureMailSubscription decided to do. */
export type SubscriptionAction = "created" | "renewed" | "kept" | "recreated";

/**
 * What Graph reports when subscriptions are listed. The clientState secret is
 * deliberately null in list/get responses (verified live), which is why a
 * subscription this record does not describe can never be reused — its
 * deliveries could never be validated.
 */
type ListedSubscription = {
  id: string;
  resource?: string;
  notificationUrl?: string;
  expirationDateTime?: string;
};

export type EnsureResult = {
  action: SubscriptionAction;
  record: SubscriptionRecord;
  detail: string;
};

type GraphCall = (path: string, init?: RequestInit) => Promise<any>;

export type EnsureOptions = {
  /** Injectable clock and transport so the harness can test every branch. */
  now?: Date;
  graph?: GraphCall;
};

function minutesUntil(expiry: string, now: Date): number {
  const at = Date.parse(expiry);
  if (Number.isNaN(at)) return -Infinity;
  return (at - now.getTime()) / 60000;
}

/**
 * What should happen to `record` right now, without touching Graph. Split out
 * from ensureMailSubscription so the policy is testable on its own.
 */
export function renewalDecision(
  record: SubscriptionRecord | null,
  notificationUrl: string,
  now: Date
): "create" | "renew" | "keep" {
  if (!record?.id) return "create";
  // A moved endpoint invalidates the subscription: Graph would keep posting to
  // the old URL, so the subscription has to be rebuilt rather than renewed.
  if (record.notificationUrl !== notificationUrl) return "create";
  if (record.resource !== SUBSCRIPTION_RESOURCE) return "create";
  const remaining = minutesUntil(record.expirationDateTime, now);
  if (remaining <= 0) return "create";
  return remaining < RENEW_WITHIN_MINUTES ? "renew" : "keep";
}

function expiryFrom(now: Date): string {
  return new Date(now.getTime() + REQUESTED_EXPIRY_MINUTES * 60000).toISOString();
}

/** A fresh, unguessable clientState. Never derived from anything in the repo. */
function newClientState(): string {
  return (crypto.randomUUID() + crypto.randomUUID()).replace(/-/g, "");
}

async function createSubscription(
  store: StateStore,
  notificationUrl: string,
  now: Date,
  graph: GraphCall
): Promise<SubscriptionRecord> {
  const clientState = newClientState();
  const created = await graph("/subscriptions", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      changeType: "created",
      notificationUrl,
      resource: SUBSCRIPTION_RESOURCE,
      expirationDateTime: expiryFrom(now),
      clientState,
      latestSupportedTlsVersion: "v1_2",
    }),
  });
  const record: SubscriptionRecord = {
    id: created.id,
    clientState,
    expirationDateTime: created.expirationDateTime,
    notificationUrl,
    resource: SUBSCRIPTION_RESOURCE,
    createdAt: now.toISOString(),
  };
  await writeJson(store, STATE_SUBSCRIPTION, record);
  return record;
}

/** The unexpired subscriptions Graph holds for this endpoint and resource. */
async function listLiveSubscriptions(
  graph: GraphCall,
  notificationUrl: string,
  now: Date
): Promise<ListedSubscription[]> {
  const listed = await graph("/subscriptions");
  return ((listed?.value ?? []) as ListedSubscription[]).filter(
    (sub) =>
      sub.notificationUrl === notificationUrl &&
      sub.resource === SUBSCRIPTION_RESOURCE &&
      minutesUntil(sub.expirationDateTime ?? "", now) > 0
  );
}

/** Best-effort deletion: a stray that survives is swept on the next upkeep. */
async function deleteSubscriptions(graph: GraphCall, ids: string[]): Promise<void> {
  for (const id of ids) {
    try {
      await graph(`/subscriptions/${encodeURIComponent(id)}`, { method: "DELETE" });
    } catch {
      // Nothing depends on this succeeding now; strays also expire on their own.
    }
  }
}

/**
 * Bring the inbox subscription to a healthy state: create it if there is none
 * (or it lapsed, or the endpoint moved), extend it when it is close to expiry,
 * and otherwise leave it alone. Safe to call as often as you like — the "keep"
 * path costs a single state read and no Graph call.
 *
 * Concurrency-safe by construction: whenever the KV record alone is not
 * enough to say "keep", Graph is treated as the source of truth and KV as a
 * cache. The live subscriptions for this endpoint are listed first, the one
 * this record holds the clientState for is renewed, and duplicates a
 * concurrent upkeep may have created are swept — so a stale KV read can no
 * longer mint an ever-growing pile of subscriptions (which is exactly what
 * a burst of racing upkeep calls once did).
 */
export async function ensureMailSubscription(
  store: StateStore,
  notificationUrl: string,
  options: EnsureOptions = {}
): Promise<EnsureResult> {
  const now = options.now ?? new Date();
  const graph = options.graph ?? callGraphServer;
  const existing = await readJson<SubscriptionRecord>(store, STATE_SUBSCRIPTION);
  const decision = renewalDecision(existing, notificationUrl, now);

  if (decision === "keep") {
    return {
      action: "kept",
      record: existing!,
      detail: `expires ${existing!.expirationDateTime} (${Math.round(
        minutesUntil(existing!.expirationDateTime, now)
      )} min away)`,
    };
  }

  const live = await listLiveSubscriptions(graph, notificationUrl, now);
  const oursIsLive = existing != null && live.some((sub) => sub.id === existing.id);

  if (oursIsLive) {
    // The subscription whose clientState this record holds is alive in Graph
    // (even if KV thought it had lapsed — Graph's view wins). Extend it and
    // sweep everything else.
    await deleteSubscriptions(
      graph,
      live.filter((sub) => sub.id !== existing!.id).map((sub) => sub.id)
    );
    const expirationDateTime = expiryFrom(now);
    try {
      const patched = await graph(`/subscriptions/${encodeURIComponent(existing!.id)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ expirationDateTime }),
      });
      const record: SubscriptionRecord = {
        ...existing!,
        expirationDateTime: patched?.expirationDateTime ?? expirationDateTime,
        renewedAt: now.toISOString(),
      };
      await writeJson(store, STATE_SUBSCRIPTION, record);
      return { action: "renewed", record, detail: `extended to ${record.expirationDateTime}` };
    } catch (err) {
      // Graph drops subscriptions it could not deliver to; a PATCH against one
      // that vanished between the list and now means "start over", not "fail".
      const record = await createSubscription(store, notificationUrl, now, graph);
      return {
        action: "recreated",
        record,
        detail: `renewal failed (${err instanceof Error ? err.message : String(err)}); created ${record.id}`,
      };
    }
  }

  // Anything Graph holds for this endpoint is not ours to use — listed
  // subscriptions come back with clientState null, so an adopted one could
  // never validate a delivery. Replace the lot with one fresh subscription.
  await deleteSubscriptions(graph, live.map((sub) => sub.id));
  const record = await createSubscription(store, notificationUrl, now, graph);
  return {
    action: existing ? "recreated" : "created",
    record,
    detail: live.length
      ? `replaced ${live.length} unusable live subscription(s) with ${record.id}`
      : `subscription ${record.id}`,
  };
}
