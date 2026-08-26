// The Worker's state store: OUTLOOK_KV, the same namespace that holds the
// Microsoft tokens. Keys are namespaced by core/kv-keys.js ("delta:", "sub:",
// "activity:", "dl:"), so the token entries and this state never collide.
import type { StateStore } from "../core/state.js";
import type { Env } from "./env.js";
import { publicBaseUrl } from "./notifications.js";

/** KV refuses an expirationTtl below 60 seconds. */
const KV_MIN_TTL_SECONDS = 60;

/** State store for one Worker request or scheduled run. */
export function kvStateStore(env: Env): StateStore {
  return {
    mode: "remote",
    // Left undefined when PUBLIC_BASE_URL is not configured: StateStore declares
    // it optional, and the tools that hand out a link to this server (attachment
    // downloads, .eml exports) already say "Set PUBLIC_BASE_URL and redeploy"
    // rather than emitting a URL nobody can reach.
    publicBaseUrl: publicBaseUrl(env),
    get: (key) => env.OUTLOOK_KV.get(key),
    put: async (key, value, options) => {
      await env.OUTLOOK_KV.put(key, value, {
        ...(options?.ttlSeconds
          ? { expirationTtl: Math.max(KV_MIN_TTL_SECONDS, Math.ceil(options.ttlSeconds)) }
          : {}),
      });
    },
    delete: async (key) => {
      await env.OUTLOOK_KV.delete(key);
    },
  };
}
