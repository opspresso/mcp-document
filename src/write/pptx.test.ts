/**
 * PPTX generation, checked by reading it back.
 *
 * The same round trip the DOCX tests use, with one thing more to check: a deck
 * is a sequence of boxes, so *how many slides* and *what landed on which* are
 * part of the output rather than a detail of it. `read/pptx.ts` numbers the
 * slides it finds, which makes the assertion a plain string comparison.
 *
 * What it cannot check is whether PowerPoint calls the file valid — a manual
 * step. The structural assertions below stand in for the parts of it that
 * actually go wrong: a cell with no paragraph, a hyperlink with no relationship,
 * a slide with no layout behind it.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { detect } from "../detect.js";
import { parseMarkdown } from "../markdown.js";
import { pptxToText } from "../read/pptx.js";
import { listEntries, readEntries } from "../zip.js";
import { renderPptx } from "./pptx.js";

const CREATED = "2026-08-13T00:00:00Z";

function build(markdown: string): Uint8Array {
  return renderPptx(parseMarkdown(markdown), { title: "test", created: CREATED }).bytes;
}

function roundTrip(markdown: string): string {
  return pptxToText(build(markdown)).text;
}

function partOf(bytes: Uint8Array, name: string): string {
  const part = readEntries(bytes, [name]).get(name);
  assert.ok(part, `expected the archive to hold ${name}`);
  return new TextDecoder().decode(part);
}

test("a deck has the scaffolding its content types declare", () => {
  const bytes = build("# Title\n\nbody");
  const names = listEntries(bytes).map((entry) => entry.name).sort();
  assert.deepEqual(names, [
    "[Content_Types].xml",
    "_rels/.rels",
    "docProps/core.xml",
    "ppt/_rels/presentation.xml.rels",
    "ppt/presentation.xml",
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
    "ppt/slideLayouts/slideLayout1.xml",
    "ppt/slideLayouts/slideLayout2.xml",
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    "ppt/slideMasters/slideMaster1.xml",
    "ppt/slides/_rels/slide1.xml.rels",
    "ppt/slides/slide1.xml",
    "ppt/theme/theme1.xml",
  ]);
  const types = partOf(bytes, "[Content_Types].xml");
  for (const declared of [
    "/ppt/presentation.xml",
    "/ppt/slideMasters/slideMaster1.xml",
    "/ppt/slides/slide1.xml",
    "/ppt/theme/theme1.xml",
  ]) {
    assert.ok(types.includes(declared), `content types should name ${declared}`);
  }
});

test("what this writes, this reads", () => {
  const bytes = build("# 제목\n\n## 첫째\n\n본문");
  assert.equal(detect(bytes, "", "deck.pptx").format, "pptx");
});

test("a level 1 or 2 heading opens a slide and becomes its title", () => {
  const rendered = renderPptx(parseMarkdown("# 표지\n\n## 하나\n\na\n\n## 둘\n\nb"), {
    title: "t",
    created: CREATED,
  });
  assert.equal(rendered.slides, 3);
  assert.equal(
    pptxToText(rendered.bytes).text,
    "## Slide 1\n표지\n\n## Slide 2\n하나\na\n\n## Slide 3\n둘\nb",
  );
});

test("a level 3 heading stays in the body rather than opening a slide", () => {
  assert.equal(roundTrip("## 하나\n\n### 안쪽\n\n본문"), "## Slide 1\n하나\n안쪽\n본문");
});

test("the deck opens on a cover only when the document does", () => {
  // A `#` at the top is a cover and takes the title layout; the same heading
  // halfway down is an ordinary section.
  const cover = build("# 표지\n\n## 다음");
  assert.ok(partOf(cover, "ppt/slides/_rels/slide1.xml.rels").includes("slideLayout1.xml"));
  assert.ok(partOf(cover, "ppt/slides/_rels/slide2.xml.rels").includes("slideLayout2.xml"));
  const plain = build("본문\n\n# 나중 제목");
  assert.ok(partOf(plain, "ppt/slides/_rels/slide1.xml.rels").includes("slideLayout2.xml"));
  assert.ok(partOf(plain, "ppt/slides/_rels/slide2.xml.rels").includes("slideLayout2.xml"));
});

test("every slide is registered in the presentation, and every one has a layout", () => {
  const bytes = build("## a\n\n## b\n\n## c");
  const presentation = partOf(bytes, "ppt/presentation.xml");
  assert.equal(presentation.match(/<p:sldId /g)?.length, 3);
  const rels = partOf(bytes, "ppt/_rels/presentation.xml.rels");
  for (let slide = 1; slide <= 3; slide += 1) {
    assert.ok(rels.includes(`Target="slides/slide${slide}.xml"`));
    assert.ok(partOf(bytes, `ppt/slides/_rels/slide${slide}.xml.rels`).includes("slideLayout"));
  }
});

test("styled text keeps its characters, and the styling is in the markup", () => {
  const bytes = build("## s\n\nplain **bold** and *italic* and `code`");
  assert.ok(pptxToText(bytes).text.includes("plain bold and italic and code"));
  const slide = partOf(bytes, "ppt/slides/slide1.xml");
  assert.ok(slide.includes(' b="1"'), "bold should be a run property");
  assert.ok(slide.includes(' i="1"'), "italic should be a run property");
  assert.ok(slide.includes('<a:latin typeface="Consolas"/>'), "code should change the face");
});

test("a link is a run property with a relationship behind it", () => {
  const bytes = build("## s\n\nsee [the spec](https://example.com/s)");
  const slide = partOf(bytes, "ppt/slides/slide1.xml");
  const id = /<a:hlinkClick[^>]* r:id="(rId\d+)"/.exec(slide)?.[1];
  assert.ok(id, "the link should be an a:hlinkClick");
  const rels = partOf(bytes, "ppt/slides/_rels/slide1.xml.rels");
  assert.ok(
    rels.includes(`Id="${id}"`) && rels.includes('Target="https://example.com/s"'),
    "the hyperlink id must resolve to an external relationship",
  );
});

test("the same target twice reuses its relationship", () => {
  const bytes = build("## s\n\n[a](https://example.com) and [b](https://example.com)");
  const rels = partOf(bytes, "ppt/slides/_rels/slide1.xml.rels");
  assert.equal(rels.match(/Type="[^"]*\/hyperlink"/g)?.length, 1);
});

test("a table becomes a table, and every cell has a paragraph in it", () => {
  const bytes = build("## s\n\n| 이름 | 값 |\n|---|---|\n| a | 1 |\n| b |  |");
  assert.ok(pptxToText(bytes).text.includes("이름 | 값\na | 1\nb"));
  const slide = partOf(bytes, "ppt/slides/slide1.xml");
  assert.ok(slide.includes("<a:tbl>"), "a table should be a graphic frame holding a:tbl");
  // A cell with no paragraph in it is what makes PowerPoint call a file corrupt.
  assert.equal(/<a:tc>(?:(?!<a:p[ />]).)*<\/a:tc>/s.test(slide), false);
});

test("a column asked to be set right is set right", () => {
  // Scoped to the table: the slide-number field is right-aligned too, and
  // counting it would make this pass for the wrong reason.
  const tableOf = (markdown: string): string =>
    /<a:tbl>[\s\S]*<\/a:tbl>/.exec(partOf(build(markdown), "ppt/slides/slide1.xml"))?.[0] ?? "";
  const aligned = tableOf("## s\n\n| a | n |\n|---|---:|\n| x | 42 |");
  assert.equal(aligned.match(/algn="r"/g)?.length, 2, "header and cell both move");
  const plain = tableOf("## s\n\n| a | n |\n|---|---|\n| x | 42 |");
  assert.equal(plain.includes('algn="r"'), false);
});

test("a slide is numbered by a field that carries no text of its own", () => {
  const bytes = build("# 표지\n\n## 본문\n\n내용");
  const numbered = partOf(bytes, "ppt/slides/slide2.xml");
  assert.match(numbered, /<a:fld id="\{[0-9A-F-]+\}" type="slidenum">/);
  // No `a:t` inside the field: the literal would be picked up by every text
  // extractor, including this server's own reader.
  assert.equal(/<a:fld[^>]*>[\s\S]*?<a:t>/.test(numbered), false);
  assert.equal(
    partOf(bytes, "ppt/slides/slide1.xml").includes("slidenum"),
    false,
    "the cover is not numbered",
  );
  // And the round trip is unchanged by any of it.
  assert.equal(pptxToText(bytes).text, "## Slide 1\n표지\n\n## Slide 2\n본문\n내용");
});

test("lists carry their markers, and a nested list numbers from one", () => {
  assert.equal(roundTrip("## s\n\n- one\n- two"), "## Slide 1\ns\n• one\n• two");
  assert.equal(
    roundTrip("## s\n\n1. a\n  1. a.1\n  2. a.2\n2. b"),
    "## Slide 1\ns\n1. a\n1. a.1\n2. a.2\n2. b",
  );
});

test("content past the bottom continues on the next slide, titled as such", () => {
  const items = Array.from({ length: 20 }, (_, index) => `- 항목 ${index + 1}`).join("\n");
  const rendered = renderPptx(parseMarkdown(`## 목록\n\n${items}`), {
    title: "t",
    created: CREATED,
  });
  assert.equal(rendered.slides, 2);
  const text = pptxToText(rendered.bytes).text;
  assert.ok(text.includes("## Slide 2\n목록 (계속)"), text);
  // Nothing is dropped in the move: every item is on one slide or the other.
  for (let item = 1; item <= 20; item += 1) {
    assert.ok(text.includes(`• 항목 ${item}\n`) || text.endsWith(`• 항목 ${item}`));
  }
});

test("a numbered list split across slides keeps counting", () => {
  // The marker is resolved before the deck is packed, which is the whole reason
  // it can: a second list block would start again at 1.
  const items = Array.from({ length: 20 }, (_, index) => `${index + 1}. 항목`).join("\n");
  const text = pptxToText(build(`## 목록\n\n${items}`)).text;
  assert.ok(text.includes("15. 항목"), text);
  assert.equal(text.match(/\n1\. 항목/g)?.length, 1, "only the first item is numbered 1");
});

test("a table too tall for one slide splits by row, with its header repeated", () => {
  const rows = Array.from({ length: 20 }, (_, index) => `| 행 ${index + 1} | ${index} |`).join("\n");
  const rendered = renderPptx(parseMarkdown(`## 표\n\n| 이름 | 값 |\n|---|---|\n${rows}`), {
    title: "t",
    created: CREATED,
  });
  assert.ok(rendered.slides > 1, "twenty rows should not fit on one slide");
  const text = pptxToText(rendered.bytes).text;
  assert.equal(text.match(/이름 \| 값/g)?.length, rendered.slides, "each slide repeats the header");
  assert.ok(text.includes("행 20"), "no row is lost in the split");
});

test("a code block keeps its lines", () => {
  assert.equal(
    roundTrip("## s\n\n```ts\nconst a = 1;\nconst b = 2;\n```"),
    "## Slide 1\ns\nconst a = 1;\nconst b = 2;",
  );
});

test("XML metacharacters in the text do not become markup", () => {
  assert.equal(roundTrip('## s\n\na < b & c > d "quoted"'), '## Slide 1\ns\na < b & c > d "quoted"');
});

test("the title lands in the document properties, escaped", () => {
  const bytes = renderPptx(parseMarkdown("body"), {
    title: "A & B <deck>",
    created: CREATED,
  }).bytes;
  const core = partOf(bytes, "docProps/core.xml");
  assert.ok(core.includes("<dc:title>A &amp; B &lt;deck&gt;</dc:title>"));
  assert.ok(core.includes(CREATED));
});

test("the same input twice produces the same bytes", () => {
  assert.deepEqual(build("## same\n\ntext"), build("## same\n\ntext"));
});

test("an empty document is still a deck of one slide", () => {
  const rendered = renderPptx({ blocks: [] }, { title: "empty", created: CREATED });
  assert.equal(rendered.slides, 1);
  const slide = partOf(rendered.bytes, "ppt/slides/slide1.xml");
  assert.ok(slide.includes("<p:spTree>"));
  assert.ok(slide.includes("<a:p>"), "an empty slide still has a paragraph to select");
});
