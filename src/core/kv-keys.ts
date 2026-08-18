// KV key names shared by the Worker (which reads and rotates them) and the
// local seeding/test scripts (which write and inspect them). Kept free of any
// Worker or Node types so both sides can import it.

/** The current mailbox refresh token. Rotated on every exchange with Microsoft. */
export const KV_REFRESH_TOKEN = "ms:refresh_token";

/** The cached mailbox access token, as {"token","expiresAt"} JSON. */
export const KV_ACCESS_TOKEN = "ms:access_token";
