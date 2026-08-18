// The authenticated download route: serves the bytes get_attachment parked in
// KV under an unguessable id.
//
// It is registered as an OAuth apiRoute alongside /mcp, so OAuthProvider has
// already validated a bearer before anything here runs and an anonymous request
// is answered with 401 + WWW-Authenticate, never with an attachment. The grant
// check below is the same defence in depth mcp-handler.ts applies: a link is
// only ever useful to the one identity this server belongs to.
import { DOWNLOAD_ROUTE_PREFIX, readDownload } from "../core/downloads.js";
import type { Env } from "./env.js";
import type { GrantProps } from "./mcp-handler.js";
import { kvStateStore } from "./state-kv.js";

type PropsContext = ExecutionContext & { props?: GrantProps };

function plain(status: number, message: string): Response {
  return new Response(message, {
    status,
    headers: { "content-type": "text/plain; charset=utf-8" },
  });
}

/** Quote a filename for Content-Disposition without letting it break the header. */
function contentDisposition(name: string): string {
  const ascii = name.replace(/[^\x20-\x7e]/g, "_").replace(/["\\]/g, "_");
  return `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(name)}`;
}

export const downloadHandler = {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    const props = (ctx as PropsContext).props;
    if (!props?.userId) return plain(401, "Unauthorized: this grant carries no user identity.");
    if (props.userId !== env.ALLOWED_MS_USER_ID) {
      return plain(403, "Forbidden: this grant is not for the allowlisted account.");
    }
    if (request.method !== "GET" && request.method !== "HEAD") {
      return plain(405, "Method not allowed");
    }

    const id = new URL(request.url).pathname.slice(DOWNLOAD_ROUTE_PREFIX.length);
    const record = await readDownload(kvStateStore(env), id);
    if (!record) {
      return plain(404, "This download link has expired or was never issued. Call get_attachment again.");
    }

    const bytes = Uint8Array.from(atob(record.base64), (c) => c.charCodeAt(0));
    return new Response(request.method === "HEAD" ? null : bytes, {
      headers: {
        "content-type": record.contentType || "application/octet-stream",
        "content-length": String(bytes.length),
        "content-disposition": contentDisposition(record.name || "attachment"),
        // Never let a shared cache keep a link that is meant to die.
        "cache-control": "private, no-store",
      },
    });
  },
};
