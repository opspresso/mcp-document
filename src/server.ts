/**
 * An MCP server that reads documents and writes them.
 *
 * It exists because a document is not its text, and a report is not a file. An
 * agent handed a `.hwp` or a `.docx` cannot open it, and an agent that has
 * written a report has no way to hand it to a person in a form they can open.
 * This closes both gaps, and only those.
 *
 * The protocol comes from `@modelcontextprotocol/server`, which serves **both
 * eras from one endpoint**: a client that opens with `server/discover` gets
 * revision `2026-07-28`, and one that opens with the `initialize` handshake is
 * served statelessly as before. Why that replaced a hand-written protocol is in
 * `mcp.ts`, beside the registration it replaced it with.
 *
 * What stays here is what the SDK has no opinion about: reading configuration,
 * the health probe, the shared-secret gate, and the routing between them.
 */

import { createServer, type ServerResponse } from "node:http";
import { createMcpHandler } from "@modelcontextprotocol/server";
import { toNodeHandler } from "@modelcontextprotocol/node";
import { authorizes, authorizesOrigin, describeAuth } from "./auth.js";
import { ConfigError, loadConfig, type Config } from "./config.js";
import { logError } from "./log.js";
import { buildServer } from "./mcp.js";
import { SERVER_NAME, SERVER_VERSION } from "./version.js";

function send(response: ServerResponse, status: number, body: unknown): void {
  const text = JSON.stringify(body);
  response.writeHead(status, { "content-type": "application/json" });
  response.end(text);
}

function start(config: Config): void {
  const mcp = toNodeHandler(
    createMcpHandler(buildServer, { onerror: (error) => logError("mcp_handler_failed", error) }),
  );
  const server = createServer((request, response) => {
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
        send(response, 403, {
          jsonrpc: "2.0",
          id: null,
          error: { code: -32002, message: "browser origins are not allowed" },
        });
        return;
      }
      if (!authorizes(config.apiKey, request.headers.authorization)) {
        response.setHeader("www-authenticate", 'Bearer realm="mcp"');
        send(response, 401, {
          jsonrpc: "2.0",
          id: null,
          error: { code: -32001, message: "missing or invalid bearer token" },
        });
        return;
      }
      await mcp(request, response);
    })();
  });

  server.listen(config.port, () => {
    console.log(`${SERVER_NAME} v${SERVER_VERSION} listening on :${config.port} (POST /mcp)`);
    console.log("documents are returned to the caller; this server stores nothing");
    // Always, not only when open: an operator reading logs to find out which
    // mode an instance is in should not have to infer it from a missing line.
    const notice = describeAuth(config.apiKey);
    if (config.apiKey) {
      console.log(notice);
    } else {
      console.warn(notice);
    }
  });

  for (const signal of ["SIGTERM", "SIGINT"] as const) {
    process.on(signal, () => server.close(() => process.exit(0)));
  }
}

try {
  start(loadConfig());
} catch (error) {
  // Fail-fast, and loudly: a missing bucket name should stop a rollout at the
  // readiness probe rather than surface inside somebody's agent run later.
  if (error instanceof ConfigError) {
    console.error(`${SERVER_NAME}: ${error.message}`);
    process.exit(1);
  }
  throw error;
}
