/**
 * What a model is actually told.
 *
 * The read path is exercised end to end against documents this server wrote
 * itself, which is the only way to run the whole chain — base64 in, detection,
 * extraction, provenance header — without a network or a fixture directory.
 *
 * The write path now goes end to end too: there is no upload past it any more,
 * so the bytes it produces are the answer and can simply be checked. What still
 * matters most are the refusals — one that does not say what to fix is the
 * failure this file exists to catch.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { parseMarkdown } from "./markdown.js";
import { callTool, CONTENT_TYPES, summarise, TOOLS } from "./tools.js";
import { renderDocx } from "./write/docx.js";

async function call(
  name: string,
  args: Record<string, unknown>,
): Promise<{ text: string; isError: boolean; file?: { mimeType: string; blob: string; uri: string } }> {
  const result = await callTool(name, args);
  const first = result.content[0];
  const resource = result.content.find((block) => block.type === "resource");
  return {
    text: first?.type === "text" ? first.text : "",
    isError: result.isError === true,
    ...(resource?.type === "resource" ? { file: resource.resource } : {}),
  };
}

const base64 = (bytes: Uint8Array) => Buffer.from(bytes).toString("base64");

test("both tools are offered with an object schema", () => {
  assert.deepEqual(
    TOOLS.map((tool) => tool.name),
    ["read_document", "render_document"],
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
  assert.match(text, /the caller reads for itself/);
});

test("reading needs the bytes, and says so", async () => {
  assert.match((await call("read_document", {})).text, /`content` is required/);
});

test("a PDF is refused by name, pointing at who does read it", async () => {
  // Silence here would read as "this document is unreadable", which is a
  // different and much more damaging claim than "not my job".
  const { text, isError } = await call("read_document", {
    content: base64(new TextEncoder().encode("%PDF-1.4 whatever")),
    filename: "report.pdf",
  });
  assert.equal(isError, true);
  assert.match(text, /this is a PDF, which the caller reads for itself/);
  assert.match(text, /DOCX, XLSX, PPTX, HWP, HWPX/);
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
  const { text, isError } = await call("render_document", { format: "hwp", content: "body" });
  assert.equal(isError, true);
  assert.match(text, /docx, pdf, hwpx/);
});

test("writing needs content", async () => {
  assert.match((await call("render_document", { format: "docx" })).text, /`content` is required/);
  assert.match(
    (await call("render_document", { format: "docx", content: "   " })).text,
    /`content` is required/,
  );
});

test("content past the limit says so, and says what to do", async () => {
  const { text, isError } = await call("render_document", { format: "docx", content: "x".repeat(500_001) });
  assert.equal(isError, true);
  assert.match(text, /over the 500,000 limit/);
  assert.match(text, /more than one document/);
});



test("the summary counts what was actually in the document", () => {
  const document = parseMarkdown(
    "# Title\n\nfirst\n\nsecond\n\n| a | b |\n|---|---|\n| 1 | 2 |\n\n- one\n- two",
  );
  assert.equal(summarise(document), "1 heading, 2 paragraphs, 1 table, 1 list");
  assert.equal(summarise({ blocks: [] }), "nothing");
});

test("a rendered document comes back as bytes, not a link", async () => {
  // This is the change: the server no longer knows a bucket exists. The caller
  // stores what it receives, beside every other byte one of its runs produced.
  const { text, isError, file } = await call("render_document", {
    format: "docx",
    content: "# 보고서\n\n본문입니다.",
    filename: "보고서",
  });
  assert.equal(isError, false);
  assert.equal(file?.mimeType, CONTENT_TYPES.docx);
  assert.match(file?.uri ?? "", /docx$/);
  // A real DOCX: the zip magic survives the base64 round trip.
  const bytes = Buffer.from(file?.blob ?? "", "base64");
  assert.deepEqual([...bytes.subarray(0, 2)], [0x50, 0x4b]);
  // And the text tells the model what happened without promising a link.
  assert.match(text, /보고서\.docx/);
  assert.doesNotMatch(text, /https?:\/\//);
});
