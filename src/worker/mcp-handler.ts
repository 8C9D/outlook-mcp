// The protected MCP endpoint. OAuthProvider only routes here once a bearer
// token has been validated, so reaching this code means the caller holds a
// grant issued to the allowlisted Microsoft identity.
//
// Stateless Streamable HTTP: a fresh McpServer and transport are built per
// request and discarded afterwards. That is what lets the Worker run without
// Durable Objects — with `sessionIdGenerator: undefined` the transport skips
// session validation entirely, so each POST is self-contained.
import { WebStandardStreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js";
import { createMcpServer } from "../core/registry.js";
import { runWithTokenProvider } from "../core/token.js";
import { VERSION } from "../core/version.js";
import type { Env } from "./env.js";
import { mailboxTokenProvider } from "./ms-token.js";

/** Props stamped onto the grant at authorization time. */
export type GrantProps = {
  userId: string;
  displayName?: string;
  upn?: string;
};

type PropsContext = ExecutionContext & { props?: GrantProps };

function jsonRpcError(status: number, message: string): Response {
  return new Response(
    JSON.stringify({ jsonrpc: "2.0", error: { code: -32000, message }, id: null }),
    { status, headers: { "content-type": "application/json" } }
  );
}

export const mcpHandler = {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    // Defence in depth: OAuthProvider has already rejected anonymous callers,
    // but a grant with no user attached must never reach the mailbox.
    const props = (ctx as PropsContext).props;
    if (!props?.userId) {
      return jsonRpcError(401, "Unauthorized: this grant carries no user identity.");
    }
    if (props.userId !== env.ALLOWED_MS_USER_ID) {
      // Catches a grant minted before the allowlist was narrowed.
      return jsonRpcError(403, "Forbidden: this grant is not for the allowlisted account.");
    }

    const server = createMcpServer(VERSION);
    const transport = new WebStandardStreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true,
    });
    await server.connect(transport);

    try {
      // Scope the KV-backed Graph token to this request; the tool layer reads
      // it back out through core/token.js without knowing where it came from.
      return await runWithTokenProvider(mailboxTokenProvider(env), () =>
        transport.handleRequest(request)
      );
    } finally {
      ctx.waitUntil(server.close());
    }
  },
};
