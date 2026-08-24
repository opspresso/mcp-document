/**
 * Everything a request passes before the protocol sees it.
 *
 * Its own module because `server.ts` boots on import — it reads the environment
 * and binds a port at the top level, which is what a process entry point should
 * do and what makes it unimportable. The routing that lived there was therefore
 * the one part of this server no test could reach: the health probe, the
 * cluster-origin refusal, the shared-secret gate and the size ceiling were four
 * decisions asserted by nothing. Building the listener here, and letting
 * `server.ts` do nothing but bind it, is what makes them testable without a
 * process.
 *
 * Order is deliberate and is the cheapest-first rule: the path decides whether
 * any of the rest applies, the two header gates cost a comparison, and the size
 * ceiling comes last among them but still ahead of the transport — an oversized
 * body refused here is memory this process never allocates.
 */

import type { IncomingMessage, ServerResponse } from "node:http";
import { createMcpHandler } from "@modelcontextprotocol/server";
import { toNodeHandler } from "@modelcontextprotocol/node";
import { authorizes, authorizesOrigin } from "./auth.js";
import type { Config } from "./config.js";
import { MAX_BODY_BYTES } from "./limits.js";
import { logError } from "./log.js";
import { buildServer } from "./mcp.js";

function send(response: ServerResponse, status: number, body: unknown): void {
  response.writeHead(status, { "content-type": "application/json" });
  response.end(JSON.stringify(body));
}

/** A refusal in the envelope a JSON-RPC caller can parse, whatever the status. */
function refuse(response: ServerResponse, status: number, code: number, message: string): void {
  send(response, status, { jsonrpc: "2.0", id: null, error: { code, message } });
}

/** Methods the transport uses without a body: nothing to weigh. */
const BODYLESS = new Set(["GET", "HEAD", "DELETE", "OPTIONS"]);

/**
 * Why this body may not be read, or nothing.
 *
 * `MAX_SOURCE_BYTES` bounds a *document*, and it can only be applied once the
 * whole request has been buffered and its base64 decoded — by which point a
 * single-threaded process has already held whatever was sent. `MAX_BODY_BYTES`
 * is the ceiling that keeps that from happening, and it was a constant nothing
 * consulted: neither the routing nor the SDK's Node adapter weighed a request,
 * so the 413 `limits.ts` describes did not exist.
 *
 * **Declared, not counted.** Attaching a byte counter to the request stream
 * would switch it into flowing mode, and the adapter builds its reader lazily —
 * the first chunks would be gone before it looked. So `Content-Length` is the
 * check, and a body that declines to state its length declines the only check
 * there is. Every client that POSTs a JSON-RPC document sends one.
 */
export function refusesBody(request: {
  method?: string | undefined;
  headers: { "content-length"?: string | undefined };
}): { status: number; message: string } | undefined {
  if (BODYLESS.has(request.method ?? "")) {
    return undefined;
  }
  const declared = Number(request.headers["content-length"]);
  if (!Number.isInteger(declared) || declared < 0) {
    return {
      status: 411,
      message: "the request body must declare its length with a Content-Length header",
    };
  }
  return declared > MAX_BODY_BYTES
    ? {
        status: 413,
        message:
          `the request body is ${declared.toLocaleString("en-US")} bytes, over the ` +
          `${MAX_BODY_BYTES.toLocaleString("en-US")} byte limit — a document sent inline as ` +
          "base64 costs a third more than its bytes",
      }
    : undefined;
}

/** The request listener `server.ts` binds, and the tests drive without binding. */
export function createRequestListener(
  config: Config,
): (request: IncomingMessage, response: ServerResponse) => void {
  const mcp = toNodeHandler(
    createMcpHandler(buildServer, { onerror: (error) => logError("mcp_handler_failed", error) }),
  );
  return (request, response) => {
    void (async () => {
      // On the path alone: a probe or a proxy is free to append a query string,
      // and matching the whole target turned `/health?x=1` into a 404.
      const path = (request.url ?? "").split("?", 1)[0] ?? "";
      if (path === "/health") {
        send(response, 200, { status: "ok" });
        return;
      }
      if (path !== "/mcp") {
        send(response, 404, { error: "not found" });
        return;
      }
      if (!authorizesOrigin(request.headers.origin)) {
        refuse(response, 403, -32002, "browser origins are not allowed");
        return;
      }
      if (!authorizes(config.apiKey, request.headers.authorization)) {
        response.setHeader("www-authenticate", 'Bearer realm="mcp"');
        refuse(response, 401, -32001, "missing or invalid bearer token");
        return;
      }
      const refusal = refusesBody(request);
      if (refusal) {
        refuse(response, refusal.status, -32003, refusal.message);
        return;
      }
      await mcp(request, response);
    })();
  };
}
