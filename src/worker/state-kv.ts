// The Worker's state store: OUTLOOK_KV, the same namespace that holds the
// Microsoft tokens. Keys are namespaced by core/kv-keys.js ("delta:", "sub:",
// "activity:"), so the token entries and this state never collide.
import type { StateStore } from "../core/state.js";
import type { Env } from "./env.js";

/** State store for one Worker request or scheduled run. */
export function kvStateStore(env: Env): StateStore {
  return {
    mode: "remote",
    get: (key) => env.OUTLOOK_KV.get(key),
    put: async (key, value) => {
      await env.OUTLOOK_KV.put(key, value);
    },
    delete: async (key) => {
      await env.OUTLOOK_KV.delete(key);
    },
  };
}
