/**
 * DOCX generation, checked by reading it back.
 *
 * The round trip is the strongest end-to-end check available without opening
 * Word: it runs the real zip, the real OOXML and the real extractor, and it
 * fails if either side stops agreeing with the other. What it cannot check is
 * whether Word calls the file valid — that is a manual step, and the structural
 * assertions below stand in for the parts of it that have actually gone wrong
 * (a cell with no paragraph in it, a hyperlink with no relationship).
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { docxToText } from "../read/docx.js";
import { parseMarkdown } from "../markdown.js";
import { listEntries, readEntries } from "../zip.js";
import { renderDocx } from "./docx.js";

const CREATED = "2026-08-05T00:00:00Z";

function build(markdown: string): Uint8Array {
  return renderDocx(parseMarkdown(markdown), { title: "test", created: CREATED });
}

function roundTrip(markdown: string): string {
  return docxToText(build(markdown)).text;
}

function partOf(bytes: Uint8Array, name: string): string {
  const part = readEntries(bytes, [name]).get(name);
  assert.ok(part, `expected the archive to hold ${name}`);
  return new TextDecoder().decode(part);
}

test("a document has exactly the parts its content types declare", () => {
  const bytes = build("# Title\n\nbody");
  const names = listEntries(bytes).map((entry) => entry.name).sort();
  assert.deepEqual(names, [
    "[Content_Types].xml",
    "_rels/.rels",
    "docProps/app.xml",
    "docProps/core.xml",
    "word/_rels/document.xml.rels",
    "word/document.xml",
    "word/footer1.xml",
    "word/styles.xml",
  ]);
  const types = partOf(bytes, "[Content_Types].xml");
  for (const declared of [
    "/word/document.xml",
    "/word/styles.xml",
    "/word/footer1.xml",
    "/docProps/core.xml",
    "/docProps/app.xml",
  ]) {
    assert.ok(types.includes(declared), `content types should name ${declared}`);
  }
});

test("the page number is a field, and the footer it sits in is wired to the section", () => {
  const bytes = build("# Title\n\nbody");
  // A number written here would be wrong the moment the text reflowed; `PAGE`
  // is computed by Word when the document is opened.
  assert.match(partOf(bytes, "word/footer1.xml"), /<w:fldSimple w:instr=" PAGE ">/);
  const id = /<w:footerReference w:type="default" r:id="(rId\d+)"\/>/.exec(
    partOf(bytes, "word/document.xml"),
  )?.[1];
  assert.ok(id, "the section must reference the footer");
  const rels = partOf(bytes, "word/_rels/document.xml.rels");
  assert.ok(rels.includes(`Id="${id}"`) && rels.includes('Target="footer1.xml"'));
});

test("headings survive with their level", () => {
  assert.equal(roundTrip("# 제목\n\n## 부제\n\n본문"), "# 제목\n## 부제\n본문");
});

test("styled text keeps its characters, and the styling is in the markup", () => {
  const bytes = build("plain **bold** and *italic* and `code`");
  assert.equal(docxToText(bytes).text, "plain bold and italic and code");
  const body = partOf(bytes, "word/document.xml");
  assert.ok(body.includes("<w:b/>"), "bold should be a run property");
  assert.ok(body.includes("<w:i/>"), "italic should be a run property");
  assert.ok(body.includes('<w:rStyle w:val="CodeChar"/>'), "code should use the code character style");
});

test("a run's leading and trailing spaces are preserved", () => {
  // Without `xml:space="preserve"` Word drops them and closes the gap, so
  // `**bold** text` comes out as `boldtext`.
  assert.equal(roundTrip("**bold** text"), "bold text");
  assert.ok(partOf(build("a b"), "word/document.xml").includes('xml:space="preserve"'));
});

test("lists come back as the markers they were written with", () => {
  assert.equal(roundTrip("- one\n- two"), "- one\n- two");
  assert.equal(roundTrip("1. first\n2. second"), "1. first\n2. second");
});

test("a nested list numbers from one each time it is entered", () => {
  const text = roundTrip("1. a\n  1. a.1\n  2. a.2\n2. b\n  1. b.1");
  assert.equal(text, "1. a\n1. a.1\n2. a.2\n2. b\n1. b.1");
});

test("a table becomes a table, and every cell has a paragraph in it", () => {
  const bytes = build("| 이름 | 값 |\n|---|---|\n| a | 1 |\n| b |  |");
  assert.equal(docxToText(bytes).text, "이름 | 값\na | 1\nb");
  const body = partOf(bytes, "word/document.xml");
  // A `w:tc` with no `w:p` inside is what makes Word call a file corrupt.
  assert.equal(/<w:tc>(?:(?!<w:p[ />]).)*<\/w:tc>/s.test(body), false);
});

test("a column asked to be set right is set right, and a plain one is untouched", () => {
  const aligned = partOf(build("| a | n |\n|---|---:|\n| x | 42 |"), "word/document.xml");
  assert.equal(aligned.match(/<w:jc w:val="right"\/>/g)?.length, 2, "header and cell both move");
  const plain = partOf(build("| a | n |\n|---|---|\n| x | 42 |"), "word/document.xml");
  assert.equal(plain.includes("<w:jc"), false, "left is Word's default and needs no element");
});

test("a link is a hyperlink with a relationship behind it", () => {
  const bytes = build("see [the spec](https://example.com/s)");
  assert.equal(docxToText(bytes).text, "see the spec");
  const body = partOf(bytes, "word/document.xml");
  const id = /<w:hyperlink r:id="(rId\d+)">/.exec(body)?.[1];
  assert.ok(id, "the link should be a w:hyperlink");
  const rels = partOf(bytes, "word/_rels/document.xml.rels");
  assert.ok(
    rels.includes(`Id="${id}"`) && rels.includes('Target="https://example.com/s"'),
    "the hyperlink id must resolve to an external relationship",
  );
});

test("one link across several runs is one hyperlink, not one per run", () => {
  const body = partOf(build("[**bold** and plain](https://example.com)"), "word/document.xml");
  assert.equal(body.match(/<w:hyperlink /g)?.length, 1);
});

test("the same target twice reuses its relationship", () => {
  const bytes = build("[a](https://example.com) and [b](https://example.com)");
  const rels = partOf(bytes, "word/_rels/document.xml.rels");
  assert.equal(rels.match(/Type="[^"]*\/hyperlink"/g)?.length, 1);
});

test("a code block keeps its lines", () => {
  assert.equal(roundTrip("```ts\nconst a = 1;\nconst b = 2;\n```"), "const a = 1;\nconst b = 2;");
});

test("XML metacharacters in the text do not become markup", () => {
  const text = roundTrip('a < b & c > d "quoted"');
  assert.equal(text, 'a < b & c > d "quoted"');
});

test("the title lands in the document properties, escaped", () => {
  const bytes = renderDocx(parseMarkdown("body"), { title: 'A & B <report>', created: CREATED });
  const core = partOf(bytes, "docProps/core.xml");
  assert.ok(core.includes("<dc:title>A &amp; B &lt;report&gt;</dc:title>"));
  assert.ok(core.includes(CREATED));
});

test("the file records what wrote it", () => {
  // Which release made a document is the first question asked about one that
  // renders oddly, and OOXML has a field for exactly that.
  assert.match(partOf(build("body"), "docProps/app.xml"), /<Application>mcp-document \d+\.\d+\.\d+<\/Application>/);
});

test("the same input twice produces the same bytes", () => {
  // The created timestamp is passed in rather than read from the clock, so a
  // document is a function of its input — which is what makes a diff of two
  // outputs mean something.
  assert.deepEqual(build("# same\n\ntext"), build("# same\n\ntext"));
});

test("an empty document is still a document", () => {
  const bytes = renderDocx({ blocks: [] }, { title: "empty", created: CREATED });
  assert.ok(listEntries(bytes).length > 0);
  assert.ok(partOf(bytes, "word/document.xml").includes("<w:body>"));
});

test("a pptx directive renders as its contents, with no fence in the page", () => {
  // `:::cards` is a deck-planning hint; on a page the contents stand where it
  // stood, and the fences never reach the text.
  const text = roundTrip(":::cards\n### 하나\n\n설명\n:::\n\n뒤 문단");
  assert.ok(text.includes("하나"));
  assert.ok(text.includes("설명"));
  assert.ok(text.includes("뒤 문단"));
  assert.equal(text.includes(":::"), false);
});
