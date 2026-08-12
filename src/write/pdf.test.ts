/**
 * PDF generation, checked by extracting the text back out with the same reader
 * this server offers.
 *
 * The assertions are on content and structure rather than on exact layout: a
 * PDF's text comes back in the order it was drawn, broken at the lines this
 * renderer chose, so pinning the whole string would pin every measurement to a
 * font's metrics. What matters is that everything written is in there, that
 * Korean survives at all — the entire reason a font is embedded — and that the
 * layout terminates on inputs designed to make it not.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { parseMarkdown } from "../markdown.js";
import { MAX_TEXT_CHARS } from "../limits.js";
import { PDFDocument } from "pdf-lib";
import { columnWidths, renderPdf, usesBold } from "./pdf.js";

const CREATED = new Date("2026-08-05T00:00:00Z");

/**
 * `unpdf` directly, rather than through a reader of ours.
 *
 * The PDF *reader* left with the URL side — the caller extracts PDFs in-process
 * and routing one here would have been a network round trip to reach this same
 * library. The round trip is still the only honest check that what this writes
 * is a PDF something can read, so the test keeps the dependency the source no
 * longer needs.
 */
/**
 * `Math.sumPrecise` is a TC39 proposal no Node this runs on has. The PDF.js
 * build inside `unpdf` calls it while rebuilding an embedded font's glyph
 * tables, so every font throws a TypeError it catches and reports as a warning —
 * one line per font, which buries everything else in the test output. Neumaier
 * summation is more than enough for glyph byte counts, and it is installed only
 * if absent so a future runtime's own wins.
 *
 * It lived beside the PDF *reader* until that left with the URL side; only this
 * round trip still needs it.
 */
const math = Math as unknown as { sumPrecise?: (values: Iterable<number>) => number };
if (typeof math.sumPrecise !== "function") {
  math.sumPrecise = (values) => {
    let sum = 0;
    let compensation = 0;
    for (const value of values) {
      const next = sum + value;
      compensation +=
        Math.abs(sum) >= Math.abs(value) ? sum - next + value : value - next + sum;
      sum = next;
    }
    return sum + compensation;
  };
}

/**
 * Page by page, joined the way the reader that used to live here did — a merged
 * extract loses the line structure one of these tests is entirely about.
 */
async function extractLines(bytes: Uint8Array): Promise<string> {
  const { extractText, getDocumentProxy } = await import("unpdf");
  const pdf = await getDocumentProxy(new Uint8Array(bytes));
  const { text } = await extractText(pdf, { mergePages: false });
  return text.join("\n\n");
}

async function roundTrip(markdown: string): Promise<string> {
  const { bytes } = await renderPdf(parseMarkdown(markdown), { title: "test", created: CREATED });
  const text = await extractLines(bytes);
  // Line breaks are the renderer's, not the document's, so they are not what
  // these assertions are about.
  return text.replace(/\s+/g, " ").trim();
}

test("what goes in comes back out", async () => {
  const text = await roundTrip("# Title\n\nA paragraph of body text.\n\n- one\n- two");
  assert.match(text, /Title/);
  assert.match(text, /A paragraph of body text\./);
  assert.match(text, /one/);
  assert.match(text, /two/);
});

test("Korean survives, which is the whole reason a font is embedded", async () => {
  // With one of PDF's built-in fonts this comes back empty or as a row of
  // boxes: they encode Latin-1 and nothing else.
  const text = await roundTrip("# 분기 보고서\n\n한글 본문이 그대로 남아 있어야 한다.");
  assert.match(text, /분기 보고서/);
  assert.match(text, /한글 본문이 그대로 남아 있어야 한다/);
});

test("the Hangul face is embedded whole, because subsetting it loses glyphs", async () => {
  // The regression, and the reason this is asserted on size rather than on
  // content: `@pdf-lib/fontkit`'s subsetter drops most Hangul glyphs *silently*.
  // The text layer stays perfect, so the round trip above passes either way and
  // the page shows blanks where two thirds of the characters should be. A
  // whole face is the only thing that distinguishes the two, and its weight is
  // the only thing a test can see.
  const { bytes } = await renderPdf(parseMarkdown("한글 문서"), { title: "t", created: CREATED });
  assert.ok(
    bytes.byteLength > 400_000,
    `expected a whole embedded face, got ${bytes.byteLength} bytes — subsetting is back`,
  );
});

test("the bold face is embedded only when something is bold", async () => {
  const plain = await renderPdf(parseMarkdown("본문뿐인 문서"), { title: "t", created: CREATED });
  const bold = await renderPdf(parseMarkdown("# 제목\n\n본문뿐인 문서"), {
    title: "t",
    created: CREATED,
  });
  assert.ok(usesBold(parseMarkdown("# 제목")), "a heading is bold");
  assert.equal(usesBold(parseMarkdown("본문")), false);
  assert.ok(
    bold.bytes.byteLength > plain.bytes.byteLength * 1.5,
    "the document with a heading should carry a second face",
  );
});

test("a long Korean paragraph wraps instead of running off the page", async () => {
  // Korean prose has no spaces to break at, so a breaker that waits for one
  // produces a single line the width of the document.
  const sentence = "이 문장은 공백 없이".replace(/ /g, "") .repeat(60);
  const { bytes } = await renderPdf(parseMarkdown(sentence), { title: "t", created: CREATED });
  const text = await extractLines(bytes);
  assert.ok(text.includes("\n"), "the paragraph should have been broken across lines");
  assert.equal(text.replace(/\s+/g, ""), sentence);
});

test("a document longer than a page gets more pages", async () => {
  const long = Array.from({ length: 120 }, (_, index) => `Paragraph number ${index}.`).join("\n\n");
  const { pages } = await renderPdf(parseMarkdown(long), { title: "t", created: CREATED });
  assert.ok(pages > 1, `expected more than one page, got ${pages}`);
});

test("tables, code, quotes and rules all render without losing their text", async () => {
  const text = await roundTrip(
    "| 이름 | 값 |\n|---|---|\n| 가 | 1 |\n\n```\nconst a = 1;\n```\n\n> 인용문\n\n---\n\n끝",
  );
  for (const expected of ["이름", "값", "가", "1", "const a = 1;", "인용문", "끝"]) {
    assert.ok(text.includes(expected), `expected ${JSON.stringify(expected)} in the output`);
  }
});

test("a link is one clickable annotation, not one per word", async () => {
  const { bytes } = await renderPdf(parseMarkdown("see [the whole spec](https://example.com/s)"), {
    title: "t",
    created: CREATED,
  });
  // Read back rather than searched for as text: pdf-lib saves with object
  // streams, so every dictionary in the file is inside a Flate stream.
  const loaded = await PDFDocument.load(bytes);
  const annotations = loaded.getPage(0).node.Annots();
  // Colouring text blue only makes it look like a link; the annotation is what
  // makes clicking it do something — and three words must not become three
  // links with dead gaps between them.
  assert.equal(annotations?.size(), 1);
  assert.match(annotations!.toString(), /\d+ 0 R/);
});

test("a word wider than the page gets its own line rather than an endless loop", async () => {
  const text = await roundTrip(`start ${"x".repeat(400)} end`);
  assert.match(text, /start/);
  assert.match(text, /end/);
});

test("an empty document is still a valid PDF", async () => {
  const { bytes, pages } = await renderPdf({ blocks: [] }, { title: "empty", created: CREATED });
  assert.equal(pages, 1);
  assert.equal(Buffer.from(bytes.subarray(0, 4)).toString("latin1"), "%PDF");
});

test("column widths follow the content and stay inside their clamps", () => {
  const cell = (text: string) => [[{ text }]];
  const rows = [
    [[{ text: "id" }], [{ text: "description" }]],
    [[{ text: "1" }], [{ text: "a much longer piece of prose in this column" }]],
  ];
  const widths = columnWidths(rows, 2);
  assert.ok(widths[1]! > widths[0]!, "the wider column should get more room");
  const total = widths.reduce((sum, width) => sum + width, 0);
  assert.ok(Math.abs(total - 481.88) < 1, `columns should fill the content width, got ${total}`);
  // One enormous cell must not squeeze the others to nothing.
  const lopsided = columnWidths([[cell("a")[0]!, [{ text: "x".repeat(5000) }]]], 2);
  assert.ok(lopsided[0]! / total > 0.05, "the small column keeps a usable share");
});
