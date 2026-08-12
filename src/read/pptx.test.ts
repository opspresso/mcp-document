import { strict as assert } from "node:assert";
import { test } from "node:test";
import { buildZip } from "../zip.js";
import { pptxToText, PptxError, slidesOf } from "./pptx.js";

const utf8 = (value: string) => new TextEncoder().encode(value);

function deck(...slides: string[]): Uint8Array {
  const parts: Record<string, Uint8Array> = {
    "ppt/presentation.xml": utf8("<p:presentation/>"),
  };
  slides.forEach((body, index) => {
    parts[`ppt/slides/slide${index + 1}.xml`] = utf8(
      `<?xml version="1.0"?><p:sld><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`,
    );
  });
  return buildZip(parts);
}

const para = (...runs: string[]) => `<a:p>${runs.map((r) => `<a:r><a:t>${r}</a:t></a:r>`).join("")}</a:p>`;

test("slides come back in deck order, which is numeric and not lexical", () => {
  const entries = [
    "ppt/slides/slide10.xml",
    "ppt/slides/slide2.xml",
    "ppt/slides/slide1.xml",
    "ppt/slides/_rels/slide1.xml.rels",
  ].map((name) => ({ name, originalSize: 0 }));
  assert.deepEqual(slidesOf(entries), [
    "ppt/slides/slide1.xml",
    "ppt/slides/slide2.xml",
    "ppt/slides/slide10.xml",
  ]);
});

test("each slide is numbered, because that is how a person addresses one", () => {
  const { text, slides } = pptxToText(deck(para("Title"), para("Second")));
  assert.equal(slides, 2);
  assert.equal(text, "## Slide 1\nTitle\n\n## Slide 2\nSecond");
});

test("runs inside a paragraph join into one line", () => {
  // A deck splits a sentence across runs whenever formatting changes mid-line.
  const { text } = pptxToText(deck(para("Revenue ", "rose ", "12%")));
  assert.equal(text, "## Slide 1\nRevenue rose 12%");
});

test("a soft break is the line the author put there", () => {
  const { text } = pptxToText(deck(`<a:p><a:r><a:t>one</a:t></a:r><a:br/><a:r><a:t>two</a:t></a:r></a:p>`));
  assert.equal(text, "## Slide 1\none\ntwo");
});

test("table cells are separated the way every other reader separates them", () => {
  const row = `<a:tr><a:tc>${para("A")}</a:tc><a:tc>${para("B")}</a:tc></a:tr>`;
  const { text } = pptxToText(deck(`<a:tbl>${row}</a:tbl>`));
  assert.match(text, /A \| B/);
});

test("an empty slide keeps its number rather than vanishing", () => {
  // Otherwise "slide 3" in the text means slide 4 in the file.
  const { text } = pptxToText(deck(para("first"), "", para("third")));
  assert.match(text, /## Slide 2\n\n## Slide 3/);
});

test("a zip with no slides is not a deck, and says so", () => {
  assert.throws(() => pptxToText(buildZip({ "notes.txt": utf8("hi") })), PptxError);
});
