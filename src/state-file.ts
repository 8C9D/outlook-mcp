// The stdio server's state store: a small JSON file next to the MSAL token
// cache, holding the delta positions check_new_mail advances. It is machine
// state, not configuration — gitignored, 0600, and safe to delete (the next
// call just re-baselines).
//
// Writes are serialised through one promise chain because the whole file is
// rewritten on every put; MCP tool calls can overlap, and last-writer-wins on
// a read-modify-write would silently drop a delta position.
import { promises as fs } from "node:fs";
import path from "node:path";
import { PROJECT_ROOT } from "./project-root.js";
import { setDefaultStateStore, type StateStore } from "./core/state.js";

export const STATE_FILE = path.join(PROJECT_ROOT, ".mcp-state.json");

/**
 * Entries are plain strings; a value written with a TTL is wrapped instead, so
 * the expiry survives a restart. Both shapes are read back transparently.
 */
type StateEntry = string | { value: string; expiresAt: number };
type StateDocument = Record<string, StateEntry>;

/** The live value of an entry, or null when it is absent or expired. */
function liveValue(entry: StateEntry | undefined): string | null {
  if (entry === undefined) return null;
  if (typeof entry === "string") return entry;
  return entry.expiresAt > Date.now() ? entry.value : null;
}

async function readDocument(file: string): Promise<StateDocument> {
  try {
    const parsed = JSON.parse(await fs.readFile(file, "utf8")) as unknown;
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
      return parsed as StateDocument;
    }
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code !== "ENOENT") {
      console.error(`Ignoring unreadable state file ${file}: ${String(err)}`);
    }
  }
  return {};
}

/** A store backed by one JSON file. `file` is overridable so tests stay isolated. */
export function createFileStateStore(file: string = STATE_FILE): StateStore {
  let queue: Promise<unknown> = Promise.resolve();

  function serialise<T>(work: () => Promise<T>): Promise<T> {
    const next = queue.then(work, work);
    // Keep the chain alive after a failed write instead of poisoning it.
    queue = next.catch(() => undefined);
    return next;
  }

  async function mutate(change: (doc: StateDocument) => void): Promise<void> {
    const doc = await readDocument(file);
    // Rewriting the whole file anyway, so drop anything that has expired.
    for (const [key, entry] of Object.entries(doc)) {
      if (liveValue(entry) === null) delete doc[key];
    }
    change(doc);
    await fs.writeFile(file, JSON.stringify(doc, null, 2) + "\n", { mode: 0o600 });
  }

  return {
    mode: "local",
    async get(key) {
      return serialise(async () => liveValue((await readDocument(file))[key]));
    },
    async put(key, value, options) {
      const entry: StateEntry = options?.ttlSeconds
        ? { value, expiresAt: Date.now() + options.ttlSeconds * 1000 }
        : value;
      return serialise(() => mutate((doc) => void (doc[key] = entry)));
    },
    async delete(key) {
      return serialise(() => mutate((doc) => void delete doc[key]));
    },
  };
}

/** Make the on-disk state file the store for this process (Node entry points). */
export function installFileStateStore(file?: string): StateStore {
  const store = createFileStateStore(file);
  setDefaultStateStore(store);
  return store;
}
