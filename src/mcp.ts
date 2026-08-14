/**
 * What this server offers a model, registered on the SDK's `McpServer`.
 *
 * Separate from `server.ts` so it can be built without binding a port — the
 * tests connect to it over HTTP, which is the only way to assert what a client
 * on either protocol era actually sees.
 *
 * **The protocol comes from the SDK.** It used to be implemented by hand, on
 * the grounds that the surface was four JSON-RPC methods and an SDK whose
 * schema generation changed under a server written against an older release is
 * what broke the off-the-shelf alternative. Revision `2026-07-28` ended that
 * trade: it removed the `initialize` handshake and added a per-request `_meta`
 * envelope, `server/discover`, `resultType` on every result, the
 * `ttlMs`/`cacheScope` hints the list verbs now require, `Mcp-Param-*`
 * mirroring from a tool's own schema, and multi round-trip results. Four
 * methods became a moving surface, and following it by hand is the larger risk.
 *
 * The tool schemas stay the JSON Schema objects `tools.ts` already declares,
 * converted rather than rewritten: they carry long, load-bearing descriptions —
 * the `render_document` one teaches a model the whole Markdown dialect — and
 * restating them in another schema language would be a transcription exercise
 * with no upside and every chance of a quiet omission.
 */

import { fromJsonSchema, McpServer, type CallToolResult } from "@modelcontextprotocol/server";
import { AjvJsonSchemaValidator } from "@modelcontextprotocol/server/validators/ajv";
import { callTool, TOOLS } from "./tools.js";
import { SERVER_NAME, SERVER_VERSION } from "./version.js";

/** One validator for every tool: compiling per registration would repeat the work. */
const validator = new AjvJsonSchemaValidator();

/**
 * A fresh server per request.
 *
 * `createMcpHandler` takes a factory rather than an instance because a
 * 2025-era request is served statelessly from its own instance. Nothing here
 * holds state between calls, so building one is cheap.
 */
export function buildServer(): McpServer {
  const server = new McpServer({ name: SERVER_NAME, version: SERVER_VERSION });
  for (const tool of TOOLS) {
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: fromJsonSchema(tool.inputSchema, validator),
      },
      // The arguments arrive validated against the same schema a client read.
      // The cast is the one seam between the two type systems: `ToolResult` is
      // this repository's own `{ content, isError? }`, which is a CallToolResult
      // — but not the whole union the SDK's callback may return, and TypeScript
      // will not narrow to a member on its own.
      async (args) => callTool(tool.name, args as Record<string, unknown>) as Promise<CallToolResult>,
    );
  }
  return server;
}
