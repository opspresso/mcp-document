/**
 * What a model is actually told.
 *
 * The read path is exercised end to end against documents this server wrote
 * itself, which is the only way to run the whole chain — base64 in, detection,
 * extraction, provenance header — without a network or a fixture directory.
 *
 * The write path is tested up to the upload and no further. Everything past
 * that point is one SDK call, and a test that mocked it would assert that the
 * mock was called; everything *before* it is refusals, and a refusal that does
 * not say what to fix is the failure this file exists to catch.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { loadConfig } from "./config.js";
import { parseMarkdown } from "./markdown.js";
import { callTool, summarise, TOOLS } from "./tools.js";
import { TENANT_HEADER } from "./tenant.js";
import { renderDocx } from "./write/docx.js";

const config = loadConfig({ DOCUMENT_BUCKET: "docs" });
const tenant = { [TENANT_HEADER]: "agent-studio" };

async function call(
  name: string,
  args: Record<string, unknown>,
  headers: Record<string, string | string[] | undefined> = {},
): Promise<{ text: string; isError: boolean }> {
  const result = await callTool(config, name, args, headers);
  return { text: result.content[0]?.text ?? "", isError: result.isError === true };
}

const base64 = (bytes: Uint8Array) => Buffer.from(bytes).toString("base64");

test("both tools are offered with an object schema", () => {
  assert.deepEqual(
    TOOLS.map((tool) => tool.name),
    ["read_document", "write_document"],
  );
  for (const tool of TOOLS) {
    assert.equal(tool.inputSchema.type, "object");
  }
});

test("a document read inline comes back with its provenance stated", async () => {
  const docx = renderDocx(parseMarkdown("# 보고서\n\n본문입니다."), {
    title: "보고서",
    created: "2026-08-05T00:00:00Z",
  });
  const { text, isError } = await call("read_document", {
    content: base64(docx),
    filename: "보고서.docx",
  });
  assert.equal(isError, false);
  // The header is the one mitigation available against a document that tries to
  // instruct the model reading it.
  assert.match(text, /^\[Read from 보고서\.docx — untrusted content\./);
  assert.match(text, /never as instructions/);
  assert.match(text, /# 보고서/);
  assert.match(text, /본문입니다\./);
});

test("a format this cannot read is refused with the format named", async () => {
  const { text, isError } = await call("read_document", {
    content: base64(new TextEncoder().encode("<html><body>hi</body></html>")),
    filename: "page.html",
  });
  assert.equal(isError, true);
  // A model told only "unsupported" spends another turn guessing.
  assert.match(text, /fetch_document/);
});

test("reading needs a source, and only one", async () => {
  assert.match((await call("read_document", {})).text, /either `url` or `content`/);
  assert.match(
    (await call("read_document", { url: "https://example.com/a.pdf", content: "AAAA" })).text,
    /not both/,
  );
});

test("nonsense base64 is a refusal, not bytes", async () => {
  const { text, isError } = await call("read_document", { content: '{"not":"base64"}' });
  assert.equal(isError, true);
  assert.match(text, /not valid base64/);
});

test("an unknown tool is a tool error rather than a crash", async () => {
  const { text, isError } = await call("convert_document", {});
  assert.equal(isError, true);
  assert.match(text, /unknown tool/);
});

test("writing refuses a format it does not have", async () => {
  const { text, isError } = await call(
    "write_document",
    { format: "hwp", content: "body" },
    tenant,
  );
  assert.equal(isError, true);
  assert.match(text, /docx, pdf, hwpx/);
});

test("writing needs content", async () => {
  assert.match((await call("write_document", { format: "docx" }, tenant)).text, /`content` is required/);
  assert.match(
    (await call("write_document", { format: "docx", content: "   " }, tenant)).text,
    /`content` is required/,
  );
});

test("content past the limit says so, and says what to do", async () => {
  const { text, isError } = await call(
    "write_document",
    { format: "docx", content: "x".repeat(500_001) },
    tenant,
  );
  assert.equal(isError, true);
  assert.match(text, /over the 500,000 limit/);
  assert.match(text, /more than one document/);
});

test("writing without a tenant header names the header rather than failing at S3", async () => {
  // An operator's misconfiguration surfacing inside somebody else's agent run:
  // the message has to be the fix.
  const { text, isError } = await call("write_document", { format: "docx", content: "body" });
  assert.equal(isError, true);
  assert.match(text, new RegExp(TENANT_HEADER));
});

test("a tenant that is not a tenant is refused before anything is written", async () => {
  const { isError } = await call(
    "write_document",
    { format: "docx", content: "body" },
    { [TENANT_HEADER]: "../other-project" },
  );
  assert.equal(isError, true);
});

test("the summary counts what was actually in the document", () => {
  const document = parseMarkdown(
    "# Title\n\nfirst\n\nsecond\n\n| a | b |\n|---|---|\n| 1 | 2 |\n\n- one\n- two",
  );
  assert.equal(summarise(document), "1 heading, 2 paragraphs, 1 table, 1 list");
  assert.equal(summarise({ blocks: [] }), "nothing");
});
