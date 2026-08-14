/**
 * What a client actually sees, over HTTP and a real client.
 *
 * The protocol moved to the SDK, so the thing worth asserting is no longer the
 * shape of a JSON-RPC envelope this repository writes — it is that a client
 * connects and finds the tools, on **both** eras. This server is reached by
 * clients that speak `2026-07-28` and by clients that still open with the
 * `initialize` handshake, and the SDK serves both from one endpoint; a change
 * that quietly dropped either would look exactly like everything working.
 *
 * Over HTTP rather than an in-memory pair, because the era is decided by the
 * transport: an in-memory link cannot tell the two apart, so it would assert
 * nothing about the endpoint that actually ships.
 */

import { strict as assert } from "node:assert";
import { createServer, type Server } from "node:http";
import type { AddressInfo } from "node:net";
import { after, before, describe, it } from "node:test";
import { Client, StreamableHTTPClientTransport } from "@modelcontextprotocol/client";
import { createMcpHandler } from "@modelcontextprotocol/server";
import { toNodeHandler } from "@modelcontextprotocol/node";
import { buildServer } from "./mcp.js";
import { TOOLS } from "./tools.js";
import { SERVER_NAME, SERVER_VERSION } from "./version.js";

let http: Server;
let url: URL;

before(async () => {
  const handler = toNodeHandler(createMcpHandler(buildServer));
  http = createServer((request, response) => void handler(request, response));
  await new Promise<void>((ready) => http.listen(0, "127.0.0.1", ready));
  url = new URL(`http://127.0.0.1:${(http.address() as AddressInfo).port}/mcp`);
});

after(async () => {
  await new Promise<void>((done) => http.close(() => done()));
});

async function connect(mode: "auto" | "legacy"): Promise<Client> {
  const client = new Client(
    { name: "test-client", version: "1.0.0" },
    { versionNegotiation: { mode } },
  );
  await client.connect(new StreamableHTTPClientTransport(url), { timeout: 5_000 });
  return client;
}

describe("what a client is offered", () => {
  it("lists every tool this server declares, with its schema intact", async () => {
    const client = await connect("auto");

    const { tools } = await client.listTools();

    assert.deepEqual(
      tools.map((tool) => tool.name).sort(),
      TOOLS.map((tool) => tool.name).sort(),
    );
    // The descriptions are the tools' documentation — `render_document`'s
    // teaches a model the whole Markdown dialect — and they travel through a
    // schema conversion now. A truncation here would be invisible in use until
    // a model stopped knowing how to ask for something.
    for (const declared of TOOLS) {
      const offered = tools.find((tool) => tool.name === declared.name);
      assert.equal(offered?.description, declared.description);
      assert.deepEqual(offered?.inputSchema.required, declared.inputSchema.required);
    }
    await client.close();
  });

  it("identifies itself with the version that was published", async () => {
    const client = await connect("auto");

    assert.deepEqual(client.getServerVersion(), {
      name: SERVER_NAME,
      version: SERVER_VERSION,
    });
    await client.close();
  });

  it("serves a client that opens with the 2026-07-28 probe", async () => {
    // The era AgentDure speaks, and the reason this server moved to the SDK.
    const client = await connect("auto");

    assert.equal(client.getProtocolEra(), "modern");
    assert.equal((await client.listTools()).tools.length, TOOLS.length);
    await client.close();
  });

  it("still serves a client that opens with the handshake", async () => {
    // The era every other client is still on. Dropping it would be invisible
    // until somebody's existing configuration stopped working.
    const client = await connect("legacy");

    assert.equal(client.getProtocolEra(), "legacy");
    assert.equal((await client.listTools()).tools.length, TOOLS.length);
    await client.close();
  });

  it("reads a document through the tool a client calls", async () => {
    // End to end, so the argument validation and the result shape are both the
    // SDK's: a plain-text file is refused by name, which is the one answer
    // `read_document` gives without touching a parser.
    const client = await connect("auto");

    const result = await client.callTool({
      name: "read_document",
      arguments: { content: Buffer.from("hello").toString("base64"), filename: "notes.txt" },
    });

    assert.equal(result.isError, true);
    await client.close();
  });
});
