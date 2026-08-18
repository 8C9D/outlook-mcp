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
import { runWithStateStore } from "../core/state.js";
import { runWithTokenProvider } from "../core/token.js";
import { VERSION } from "../core/version.js";
import type { Env } from "./env.js";
import { mailboxTokenProvider } from "./ms-token.js";
import { keepSubscriptionAlive } from "./notifications.js";
import { kvStateStore } from "./state-kv.js";

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
      // Scope the KV-backed Graph token and state store to this request; the
      // tool layer reads both back out through core/token.js and core/state.js
      // without knowing where they came from.
      return await runWithTokenProvider(mailboxTokenProvider(env), () =>
        runWithStateStore(kvStateStore(env), () => transport.handleRequest(request))
      );
    } finally {
      ctx.waitUntil(server.close());
      // Belt and braces for the cron: using the connector at all is enough to
      // notice a lapsed subscription. Costs one KV read when nothing is due.
      ctx.waitUntil(
        keepSubscriptionAlive(env).catch((err) =>
          console.error(`Background subscription upkeep failed: ${String(err)}`)
        )
      );
    }
  },
};
