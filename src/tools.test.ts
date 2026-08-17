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
import { mock, test } from "node:test";
import { parseMarkdown } from "./markdown.js";
import { callTool, CONTENT_TYPES, PROFILES, summarise, TOOLS } from "./tools.js";
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

test("render_document offers the five professional profiles", () => {
  const render = TOOLS.find((tool) => tool.name === "render_document");
  assert.deepEqual(render?.inputSchema.properties.profile.enum, [...PROFILES]);
  assert.deepEqual(render?.inputSchema.required, ["format", "content"]);
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

test("writing refuses an unknown profile and names every choice", async () => {
  const { text, isError } = await call("render_document", {
    format: "docx",
    profile: "luxury",
    content: "## 결론\n\n진행한다.",
  });
  assert.equal(isError, true);
  assert.match(text, /`profile` must be one of executive, consulting, formal, technical, standard/);
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

test("a deck comes back as a deck, and says how many slides it is", async () => {
  const { text, isError, file } = await call("render_document", {
    format: "pptx",
    content: "# 표지\n\n## 첫째\n\n내용\n\n## 둘째\n\n내용",
  });
  assert.equal(isError, false);
  assert.equal(file?.mimeType, CONTENT_TYPES.pptx);
  assert.match(file?.uri ?? "", /pptx$/);
  // The slide count is what a page count is for a PDF: the unit the format has.
  assert.match(text, /3 slides/);
});

test("assets ride with pptx and docx, and an SVG is turned away with the fix named", async () => {
  const png = {
    mimeType: "image/png",
    content: Buffer.from([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52,
      0, 0, 0x02, 0x80, 0, 0, 0x01, 0x90, 8, 6, 0, 0, 0,
    ]).toString("base64"),
  };
  const wrongFormat = await call("render_document", {
    format: "hwpx",
    content: "# 보고서",
    assets: { "a.png": png },
  });
  assert.ok(wrongFormat.isError && wrongFormat.text.includes("pptx and docx only"), wrongFormat.text);

  const svg = await call("render_document", {
    format: "pptx",
    content: "## 그림\n\n![x](asset://a.svg)",
    assets: { "a.svg": { mimeType: "image/svg+xml", content: "PHN2Zz48L3N2Zz4=" } },
  });
  assert.ok(svg.isError && svg.text.includes("Rasterise"), svg.text);

  for (const format of ["pptx", "docx"] as const) {
    const embedded = await call("render_document", {
      format,
      content: "## 구조도\n\n![전체 구조](asset://a.png)",
      assets: { "a.png": png },
    });
    assert.equal(embedded.isError, false, embedded.text);
    assert.ok(embedded.file, `the ${format} came back with the picture inside`);
  }
});

test("a deck that references an asset nobody sent is refused by name", async () => {
  const result = await call("render_document", {
    format: "pptx",
    content: "## 그림\n\n![x](asset://ghost.png)",
  });
  assert.ok(result.isError && result.text.includes("asset://ghost.png"), result.text);
});

/** The `tool_call` lines written while `run` executes, parsed. */
async function linesDuring(run: () => Promise<unknown>): Promise<Record<string, unknown>[]> {
  const write = mock.method(console, "log", () => {});
  const complain = mock.method(console, "error", () => {});
  try {
    await run();
    return write.mock.calls
      .map((call) => JSON.parse(String(call.arguments[0])) as Record<string, unknown>)
      .filter((line) => line.event === "tool_call");
  } finally {
    write.mock.restore();
    complain.mock.restore();
  }
}

test("every call leaves one line naming the tool and the format, never the content", async () => {
  const lines = await linesDuring(async () => {
    await call("render_document", { format: "docx", content: "# Report\n\nthe secret paragraph" });
    await call("read_document", { filename: "Quarterly.HWP", content: base64(new Uint8Array([1, 2, 3])) });
  });
  assert.equal(lines.length, 2);

  const [rendered, read] = lines;
  assert.equal(rendered?.tool, "render_document");
  assert.equal(rendered?.format, "docx");
  assert.equal(rendered?.ok, true);
  assert.equal(typeof rendered?.ms, "number");

  // The read is refused, and a refusal is still a call — with the extension it
  // was named by, lower-cased, so one format's failures group together.
  assert.equal(read?.tool, "read_document");
  assert.equal(read?.format, "hwp");
  assert.equal(read?.ok, false);

  assert.doesNotMatch(JSON.stringify(lines), /secret|Quarterly/);
});
