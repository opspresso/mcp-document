/**
 * An MCP server that reads documents and writes them.
 *
 * It exists because a document is not its text, and a report is not a file. An
 * agent handed a `.hwp` or a `.docx` cannot open it, and an agent that has
 * written a report has no way to hand it to a person in a form they can open.
 * This closes both gaps, and only those.
 *
 * The protocol is implemented directly rather than through an SDK, as in
 * `mcp-url-fetch`. The surface is four methods, and the one dependency that
 * mattered there — an SDK whose schema generation changed under a server
 * written against an older release — is exactly what broke the off-the-shelf
 * alternative.
 */

import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { authorizes, describeAuth } from "./auth.js";
import { ConfigError, loadConfig, type Config } from "./config.js";
import { MAX_BODY_BYTES } from "./limits.js";
import { callTool, TOOLS } from "./tools.js";
import { SERVER_NAME, SERVER_VERSION } from "./version.js";

const PROTOCOL_VERSION = "2025-06-18";

interface JsonRpcRequest {
  jsonrpc: string;
  id?: number | string;
  method: string;
  params?: Record<string, unknown>;
}

async function handle(message: JsonRpcRequest): Promise<unknown> {
  switch (message.method) {
    case "initialize":
      return {
        protocolVersion: PROTOCOL_VERSION,
        capabilities: { tools: { listChanged: false } },
        serverInfo: { name: SERVER_NAME, version: SERVER_VERSION },
      };
    case "tools/list":
      return { tools: TOOLS };
    // Not gated behind a capability: ping is part of the base protocol and the
    // receiver must answer it. A client using it as a keepalive reads an error
    // here as a dead connection.
    case "ping":
      return {};
    case "tools/call": {
      const name = message.params?.name;
      if (!TOOLS.some((tool) => tool.name === name)) {
        throw new Error(`unknown tool: ${String(name)}`);
      }
      return callTool(name, (message.params?.arguments ?? {}) as Record<string, unknown>);
    }
    default:
      throw new Error(`unsupported method: ${message.method}`);
  }
}

/** Distinguished from a parse failure so the caller is not sent to debug its JSON. */
class BodyTooLarge extends Error {}

async function readBody(request: IncomingMessage): Promise<string> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of request) {
    size += (chunk as Buffer).byteLength;
    if (size > MAX_BODY_BYTES) {
      throw new BodyTooLarge(
        `request body is over the ${MAX_BODY_BYTES.toLocaleString("en-US")} byte limit — a ` +
          "document that large has to be passed as `url` rather than inline",
      );
    }
    chunks.push(chunk as Buffer);
  }
  return Buffer.concat(chunks).toString("utf8");
}

function send(response: ServerResponse, status: number, body: unknown): void {
  const text = JSON.stringify(body);
  response.writeHead(status, { "content-type": "application/json" });
  response.end(text);
}

function start(config: Config): void {
  const server = createServer((request, response) => {
    void (async () => {
      // On the path alone: a probe or a proxy is free to append a query string,
      // and matching the whole target turned `/health?x=1` into a 404.
      const path = (request.url ?? "").split("?", 1)[0] ?? "";
      if (path === "/health") {
        send(response, 200, { status: "ok" });
        return;
      }
      if (!path.startsWith("/mcp")) {
        send(response, 404, { error: "not found" });
        return;
      }
      if (!authorizes(config.apiKey, request.headers.authorization)) {
        send(response, 401, {
          jsonrpc: "2.0",
          id: null,
          error: { code: -32001, message: "missing or invalid bearer token" },
        });
        return;
      }
      if (request.method === "DELETE") {
        // Session teardown: this server is stateless, so there is nothing to release.
        response.writeHead(204).end();
        return;
      }
      if (request.method !== "POST") {
        send(response, 405, { error: "method not allowed" });
        return;
      }
      let body: string;
      try {
        body = await readBody(request);
      } catch (error) {
        const tooLarge = error instanceof BodyTooLarge;
        send(response, tooLarge ? 413 : 400, {
          jsonrpc: "2.0",
          id: null,
          error: {
            code: -32600,
            message: tooLarge ? error.message : "could not read the request body",
          },
        });
        return;
      }
      let message: JsonRpcRequest;
      try {
        message = JSON.parse(body) as JsonRpcRequest;
      } catch {
        send(response, 400, {
          jsonrpc: "2.0",
          id: null,
          error: { code: -32700, message: "parse error" },
        });
        return;
      }
      // A notification carries no id and expects no reply.
      if (message.id === undefined) {
        response.writeHead(202).end();
        return;
      }
      try {
        send(response, 200, {
          jsonrpc: "2.0",
          id: message.id,
          result: await handle(message),
        });
      } catch (error) {
        send(response, 200, {
          jsonrpc: "2.0",
          id: message.id,
          error: { code: -32601, message: (error as Error).message },
        });
      }
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
