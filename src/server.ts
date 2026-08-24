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
 * What is left here is the process: read the configuration, bind the listener
 * `http.ts` builds, say which authentication mode it came up in, and close on a
 * signal. Nothing in this file is importable without booting — which is why the
 * routing is not in it.
 */

import { createServer } from "node:http";
import { ConfigError, loadConfig, type Config } from "./config.js";
import { describeAuth } from "./auth.js";
import { createRequestListener } from "./http.js";
import { SERVER_NAME, SERVER_VERSION } from "./version.js";

function start(config: Config): void {
  const server = createServer(createRequestListener(config));

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
  // Fail-fast, and loudly: a malformed PORT should stop a rollout at the
  // readiness probe rather than surface inside somebody's agent run later.
  if (error instanceof ConfigError) {
    console.error(`${SERVER_NAME}: ${error.message}`);
    process.exit(1);
  }
  throw error;
}
