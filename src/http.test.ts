/**
 * The gates a request passes before the protocol sees it.
 *
 * Asserted through a bound socket rather than by calling the predicates,
 * because three of the four are only meaningful as HTTP: a status code, a
 * `www-authenticate` header, and — the one that made this file necessary —
 * whether a real MCP client still connects through them. The size ceiling is
 * enforced on `Content-Length`, which is a bet that every client states it;
 * `a real client still connects` is that bet, written down.
 */

import { strict as assert } from "node:assert";
import { createServer, type Server } from "node:http";
import { connect, type AddressInfo } from "node:net";
import { after, before, test } from "node:test";
import { Client, StreamableHTTPClientTransport } from "@modelcontextprotocol/client";
import { createRequestListener, refusesBody } from "./http.js";
import { MAX_BODY_BYTES } from "./limits.js";

const KEY = "s3cr3t";

let http: Server;
let origin: string;
let port: number;

before(async () => {
  http = createServer(createRequestListener({ port: 0, apiKey: KEY }));
  await new Promise<void>((ready) => http.listen(0, "127.0.0.1", ready));
  port = (http.address() as AddressInfo).port;
  origin = `http://127.0.0.1:${port}`;
});

after(async () => {
  await new Promise<void>((done) => http.close(() => done()));
});

/** A JSON-RPC POST with the headers a client would send, bar the ones named. */
async function post(
  headers: Record<string, string> = {},
  body = '{"jsonrpc":"2.0","id":1,"method":"tools/list"}',
): Promise<{ status: number; headers: Headers; error?: { code: number; message: string } }> {
  const response = await fetch(`${origin}/mcp`, {
    method: "POST",
    headers: { "content-type": "application/json", ...headers },
    body,
  });
  const parsed = (await response.json()) as { error?: { code: number; message: string } };
  return { status: response.status, headers: response.headers, ...parsed };
}

/**
 * A POST's headers alone, down a raw socket, and what came back.
 *
 * `fetch` will not send a `Content-Length` that disagrees with the body it
 * holds, and will not omit one either — and both are exactly the request this
 * gate exists to judge, since the whole point is that nothing reads the body
 * before the header is read.
 */
async function announce(...extra: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    const socket = connect(port, "127.0.0.1", () => {
      socket.write(
        "POST /mcp HTTP/1.1\r\n" +
          "host: 127.0.0.1\r\n" +
          `authorization: Bearer ${KEY}\r\n` +
          "content-type: application/json\r\n" +
          `${extra.join("\r\n")}\r\n\r\n`,
      );
    });
    let received = "";
    socket.on("data", (chunk) => {
      received += chunk.toString("utf8");
      if (received.includes("\r\n\r\n")) {
        socket.destroy();
        resolve(received);
      }
    });
    socket.on("error", reject);
  });
}

test("the health probe answers, with or without a query string", async () => {
  for (const path of ["/health", "/health?ready=1"]) {
    const response = await fetch(`${origin}${path}`);
    assert.equal(response.status, 200);
    assert.deepEqual(await response.json(), { status: "ok" });
  }
});

test("a path this server does not serve is a 404, not a protocol error", async () => {
  const response = await fetch(`${origin}/`);
  assert.equal(response.status, 404);
});

test("a browser origin is refused before the key is even considered", async () => {
  const response = await post({ authorization: `Bearer ${KEY}`, origin: "https://example.com" });
  assert.equal(response.status, 403);
  assert.equal(response.error?.code, -32002);
});

test("a missing or wrong key is a 401 that says how to present one", async () => {
  const cases: Record<string, string>[] = [{}, { authorization: "Bearer wrong" }, { authorization: KEY }];
  for (const headers of cases) {
    const response = await post(headers);
    assert.equal(response.status, 401);
    assert.equal(response.headers.get("www-authenticate"), 'Bearer realm="mcp"');
    assert.equal(response.error?.code, -32001);
  }
});

test("a body over the ceiling is refused on what it declared, before it is read", async () => {
  const answer = await announce(`content-length: ${MAX_BODY_BYTES + 1}`);
  assert.match(answer, /^HTTP\/1\.1 413 /);
  assert.match(answer, /over the .* byte limit/);
});

test("a chunked body states no length, so there is nothing to weigh and it is refused", async () => {
  // The one shape that reaches the gate without a length: a malformed
  // `Content-Length` never gets this far, because Node's own parser answers it
  // with a 400 before the listener is called.
  const answer = await announce("transfer-encoding: chunked");
  assert.match(answer, /^HTTP\/1\.1 411 /);
});

test("a body at the ceiling is not refused for its size", async () => {
  assert.equal(refusesBody({ method: "POST", headers: { "content-length": String(MAX_BODY_BYTES) } }), undefined);
  assert.equal(
    refusesBody({ method: "POST", headers: { "content-length": String(MAX_BODY_BYTES + 1) } })?.status,
    413,
  );
});

test("a body that will not say how long it is cannot be weighed, so it is refused", async () => {
  assert.equal(refusesBody({ method: "POST", headers: {} })?.status, 411);
  assert.equal(refusesBody({ method: "POST", headers: { "content-length": "nonsense" } })?.status, 411);
  // The transport's own bodyless verbs carry no length and must still pass.
  for (const method of ["GET", "HEAD", "DELETE", "OPTIONS"]) {
    assert.equal(refusesBody({ method, headers: {} }), undefined);
  }
});

test("a real client connects and lists the tools through every gate", async () => {
  // The load-bearing one. `refusesBody` refuses a body with no `Content-Length`,
  // which is only safe while the clients that matter send one — so this asserts
  // that the SDK's own transport, carrying the key, gets all the way through.
  const client = new Client({ name: "gate-test", version: "1.0.0" });
  await client.connect(
    new StreamableHTTPClientTransport(new URL(`${origin}/mcp`), {
      requestInit: { headers: { authorization: `Bearer ${KEY}` } },
    }),
    { timeout: 5_000 },
  );
  const { tools } = await client.listTools();
  assert.ok(tools.length > 0);
  await client.close();
});
