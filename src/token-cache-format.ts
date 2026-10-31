// Keeps a token cache written by MSAL Node 2 usable under the current MSAL.
//
// MSAL 2 ended every credential key with one more empty segment than MSAL 7
// does, so a refresh under MSAL 7 stores the rotated token beside the old entry
// rather than over it. Left alone, that fails slowly: MSAL reads the first
// match, which is the stale entry, on every refresh, and it cannot evict that
// entry when it finally expires, because it rebuilds the key to delete in the
// new shape. The sign-in then stays broken until someone runs `npm run login`,
// which also breaks the Worker's refresh-token chain. And `npm run seed:kv`
// would read the second entry as a second account and refuse.
//
// So every read of the cache goes through migrateTokenCache first. A stale
// entry yields to its current-format twin, which can only have been written
// later; an entry with no twin is renamed in place, keeping the only token.
// Deliberately free of MSAL imports, so the offline tier can test it.

const CREDENTIAL_SECTIONS = ["AccessToken", "IdToken", "RefreshToken"] as const;

type CachedCredential = {
  home_account_id?: string;
  environment?: string;
  credential_type?: string;
  client_id?: string;
  family_id?: string;
  realm?: string;
  target?: string;
  token_type?: string;
};

/** The key MSAL 7 gives a credential; mirrors generateCredentialKey in msal-node. */
function currentKey(credential: CachedCredential): string {
  const owner =
    (credential.credential_type === "RefreshToken" && credential.family_id) || credential.client_id;
  const tokenType = credential.token_type?.toLowerCase();
  const scheme = tokenType && tokenType !== "bearer" ? tokenType : "";
  return [
    credential.home_account_id,
    credential.environment,
    credential.credential_type,
    owner,
    credential.realm ?? "",
    credential.target ?? "",
    scheme,
  ]
    .join("-")
    .toLowerCase();
}

/**
 * Returns `serialized` with MSAL 2 credential keys moved to the current shape.
 * Text that does not parse, and a cache with nothing to migrate, come back
 * unchanged, so MSAL still reports a corrupt file in its own words.
 */
export function migrateTokenCache(serialized: string): string {
  let cache: Partial<Record<(typeof CREDENTIAL_SECTIONS)[number], Record<string, CachedCredential>>>;
  try {
    cache = JSON.parse(serialized);
  } catch {
    return serialized;
  }

  let changed = false;
  for (const section of CREDENTIAL_SECTIONS) {
    const entries = cache?.[section];
    if (!entries || typeof entries !== "object") continue;
    for (const [key, credential] of Object.entries(entries)) {
      const current = currentKey(credential);
      if (key !== `${current}-`) continue;
      if (!(current in entries)) entries[current] = credential;
      delete entries[key];
      changed = true;
    }
  }
  return changed ? JSON.stringify(cache) : serialized;
}
