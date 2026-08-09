/**
 * The parser every renderer is downstream of. A mistake here appears in all
 * three formats at once and looks like three separate bugs.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { parseInline, parseMarkdown, plainTextOf, type Block } from "./markdown.js";

function blocks(source: string): Block[] {
  return parseMarkdown(source).blocks;
}

test("headings carry their level, and the first h1 becomes the title", () => {
  const document = parseMarkdown("# 보고서\n\n## 배경\n\n본문");
  assert.equal(document.title, "보고서");
  assert.deepEqual(
    document.blocks.map((block) => block.kind),
    ["heading", "heading", "paragraph"],
  );
  assert.equal(document.blocks[1]?.kind === "heading" ? document.blocks[1].level : 0, 2);
});

test("lines with no blank between them are one paragraph, as Markdown says", () => {
  const parsed = blocks("first line\nsecond line\n\nnew paragraph");
  assert.equal(parsed.length, 2);
  assert.equal(parsed[0]?.kind === "paragraph" ? plainTextOf(parsed[0].runs) : "", "first line second line");
});

test("emphasis nests, and adjacent runs of one style are one run", () => {
  const runs = parseInline("plain **bold and *both* more** tail");
  assert.deepEqual(runs, [
    { text: "plain " },
    { text: "bold and ", bold: true },
    { text: "both", bold: true, italic: true },
    { text: " more", bold: true },
    { text: " tail" },
  ]);
});

test("a code span is literal, so markers inside it are characters", () => {
  assert.deepEqual(parseInline("run `a * b ** c` now"), [
    { text: "run " },
    { text: "a * b ** c", code: true },
    { text: " now" },
  ]);
});

test("an underscore inside a word is not emphasis", () => {
  // The regression this exists for: `snake_case_name` parsed as emphasis loses
  // its underscores and silently becomes a different identifier.
  assert.deepEqual(parseInline("call snake_case_name here"), [{ text: "call snake_case_name here" }]);
  assert.deepEqual(parseInline("_really_ emphasised"), [
    { text: "really", italic: true },
    { text: " emphasised" },
  ]);
});

test("links keep their target on every run they cover", () => {
  assert.deepEqual(parseInline("see [the **spec**](https://example.com/s)"), [
    { text: "see " },
    { text: "the ", href: "https://example.com/s" },
    { text: "spec", bold: true, href: "https://example.com/s" },
  ]);
  // An empty label is the URL: a link with no visible text is not a link.
  assert.deepEqual(parseInline("[](https://example.com)"), [
    { text: "https://example.com", href: "https://example.com" },
  ]);
});

test("a backslash escapes the punctuation Markdown gives meaning to", () => {
  assert.deepEqual(parseInline("2 \\* 3 \\*\\* 4"), [{ text: "2 * 3 ** 4" }]);
});

test("a fenced block keeps its lines and its language", () => {
  const parsed = blocks("```ts\nconst a = 1;\n\n  indented\n```\nafter");
  assert.deepEqual(parsed[0], { kind: "code", text: "const a = 1;\n\n  indented", language: "ts" });
  assert.equal(parsed[1]?.kind, "paragraph");
});

test("an unclosed fence takes the rest of the document, which is what it says", () => {
  const parsed = blocks("text\n\n```\nnever closed\nstill code");
  assert.deepEqual(parsed[1], { kind: "code", text: "never closed\nstill code" });
});

test("bullets and numbers are separate lists, and indentation is depth", () => {
  const parsed = blocks("- one\n  - nested\n- two\n\n1. first\n2. second");
  assert.deepEqual(parsed[0], {
    kind: "list",
    ordered: false,
    items: [
      { runs: [{ text: "one" }], depth: 0 },
      { runs: [{ text: "nested" }], depth: 1 },
      { runs: [{ text: "two" }], depth: 0 },
    ],
  });
  assert.equal(parsed[1]?.kind === "list" ? parsed[1].ordered : false, true);
});

test("a bullet list touching a numbered one does not merge into it", () => {
  // Merging would renumber one of them out of existence.
  const parsed = blocks("- bullet\n1. number");
  assert.equal(parsed.length, 2);
  assert.equal(parsed[0]?.kind === "list" ? parsed[0].ordered : true, false);
  assert.equal(parsed[1]?.kind === "list" ? parsed[1].ordered : false, true);
});

test("a table needs its divider, and its cells are parsed as inline", () => {
  const parsed = blocks("| 이름 | 값 |\n|---|---:|\n| **a** | 1 |\n| b | 2 |");
  assert.equal(parsed[0]?.kind, "table");
  if (parsed[0]?.kind !== "table") {
    return;
  }
  assert.deepEqual(parsed[0].header.map(plainTextOf), ["이름", "값"]);
  assert.equal(parsed[0].rows.length, 2);
  assert.deepEqual(parsed[0].rows[0]?.[0], [{ text: "a", bold: true }]);
});

test("a pipe row with no divider under it is a paragraph, not a table", () => {
  const parsed = blocks("a | b | c");
  assert.equal(parsed[0]?.kind, "paragraph");
});

test("an escaped pipe stays inside its cell", () => {
  const parsed = blocks("| expr | note |\n|---|---|\n| a \\| b | or |");
  assert.equal(parsed[0]?.kind === "table" ? plainTextOf(parsed[0].rows[0]?.[0] ?? []) : "", "a | b");
});

test("consecutive quote lines are one quote", () => {
  const parsed = blocks("> first\n> second\n\nafter");
  assert.deepEqual(parsed[0], { kind: "quote", runs: [{ text: "first second" }] });
});

test("a rule is a rule, and a dash with text after it is a list item", () => {
  assert.deepEqual(blocks("---")[0], { kind: "rule" });
  assert.equal(blocks("- item")[0]?.kind, "list");
});

test("syntax this does not support survives as the characters that were written", () => {
  // Refusing to produce a document over one line of it would be a much worse
  // outcome than an unrendered `<div>` a reader can see.
  const parsed = blocks("<div>raw</div>\n\n[^1]: a footnote");
  assert.equal(plainTextOf(parsed[0]?.kind === "paragraph" ? parsed[0].runs : []), "<div>raw</div>");
  assert.equal(plainTextOf(parsed[1]?.kind === "paragraph" ? parsed[1].runs : []), "[^1]: a footnote");
});

test("an image becomes a link to the picture, labelled with its alt text", () => {
  // Nothing here embeds pictures. Dropping it would lose both the description
  // and the address; a literal `![alt](url)` in a rendered document is noise.
  assert.deepEqual(parseInline("see ![도표 1](https://example.com/chart.png)"), [
    { text: "see " },
    { text: "도표 1", href: "https://example.com/chart.png" },
  ]);
  assert.deepEqual(parseInline("![](x.png)"), [{ text: "image", href: "x.png" }]);
});

test("an empty document has no blocks and no title", () => {
  assert.deepEqual(parseMarkdown("   \n\n  "), { blocks: [] });
});
