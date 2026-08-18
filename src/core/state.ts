// Transport-agnostic small-state storage, alongside core/token.js.
//
// Two features need to remember something between calls: the delta token that
// makes check_new_mail incremental, and the change-notification ring buffer
// that get_mailbox_activity reads. Neither may know where that state lives —
// on Workers it is KV, on stdio a 0600 JSON file next to the token cache — so
// callers install a store and the tools just read and write keys.
//
// AsyncLocalStorage scopes a store to one request for the same reason the token
// provider is scoped: on Workers the KV binding arrives per-request via `env`.
import { AsyncLocalStorage } from "node:async_hooks";

/**
 * Where the installed store keeps its data. Features that only work when the
 * server is reachable from the internet (change notifications) check this
 * rather than sniffing for a KV binding.
 */
export type StateMode = "local" | "remote";

/** Options for a single write. */
export type PutOptions = {
  /**
   * Drop the entry after this many seconds. Best-effort garbage collection, not
   * a security boundary: KV enforces it server-side (with a 60 s minimum), the
   * other stores approximate it, and callers that care about the deadline must
   * also record it inside the value (see core/downloads.js).
   */
  ttlSeconds?: number;
};

/** A tiny string keyed store. Values are JSON documents written by the caller. */
export interface StateStore {
  readonly mode: StateMode;
  /**
   * Public origin of this server, when it has one (the Worker's PUBLIC_BASE_URL).
   * Alongside `mode` because it answers the same question — what kind of host is
   * this running on — for features that have to hand out a URL to themselves.
   */
  readonly publicBaseUrl?: string;
  get(key: string): Promise<string | null>;
  put(key: string, value: string, options?: PutOptions): Promise<void>;
  delete(key: string): Promise<void>;
}

/** Thrown when a tool needs persistent state and no store has been installed. */
export class StateUnavailableError extends Error {
  constructor(detail?: string) {
    super(
      "This server has no state store installed, so it cannot remember anything between calls." +
        (detail ? ` (${detail})` : "")
    );
    this.name = "StateUnavailableError";
  }
}

const storage = new AsyncLocalStorage<StateStore>();
let defaultStore: StateStore | undefined;

/** Install the process-wide store (stdio server, CLI scripts, test harness). */
export function setDefaultStateStore(store: StateStore): void {
  defaultStore = store;
}

/** Run `fn` with `store` scoped to it, overriding the default (Worker requests). */
export function runWithStateStore<T>(store: StateStore, fn: () => T): T {
  return storage.run(store, fn);
}

/** The store in force here, or undefined when none is installed. */
export function getStateStore(): StateStore | undefined {
  return storage.getStore() ?? defaultStore;
}

/** The store in force here; throws StateUnavailableError when there is none. */
export function requireStateStore(): StateStore {
  const store = getStateStore();
  if (!store) throw new StateUnavailableError();
  return store;
}

/** Read and parse a JSON value, treating anything unparseable as absent. */
export async function readJson<T>(store: StateStore, key: string): Promise<T | null> {
  const raw = await store.get(key);
  if (!raw) return null;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return null;
  }
}

/** Write a JSON value. */
export async function writeJson(store: StateStore, key: string, value: unknown): Promise<void> {
  await store.put(key, JSON.stringify(value));
}

/** An in-process store. Used by the test harness; never by a running server. */
export function createMemoryStateStore(
  mode: StateMode = "remote",
  publicBaseUrl?: string
): StateStore {
  const map = new Map<string, { value: string; expiresAt?: number }>();
  return {
    mode,
    publicBaseUrl,
    async get(key) {
      const entry = map.get(key);
      if (!entry) return null;
      if (entry.expiresAt !== undefined && entry.expiresAt <= Date.now()) {
        map.delete(key);
        return null;
      }
      return entry.value;
    },
    async put(key, value, options) {
      map.set(key, {
        value,
        ...(options?.ttlSeconds ? { expiresAt: Date.now() + options.ttlSeconds * 1000 } : {}),
      });
    },
    async delete(key) {
      map.delete(key);
    },
  };
}
