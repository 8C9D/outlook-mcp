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

/**
 * Bring the inbox subscription to a healthy state: create it if there is none
 * (or it lapsed, or the endpoint moved), extend it when it is close to expiry,
 * and otherwise leave it alone. Safe to call as often as you like — the "keep"
 * path costs a single state read and no Graph call.
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

  if (decision === "renew") {
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
      // that no longer exists means "start over", not "fail".
      const record = await createSubscription(store, notificationUrl, now, graph);
      return {
        action: "recreated",
        record,
        detail: `renewal failed (${err instanceof Error ? err.message : String(err)}); created ${record.id}`,
      };
    }
  }

  const record = await createSubscription(store, notificationUrl, now, graph);
  return { action: "created", record, detail: `subscription ${record.id}` };
}
