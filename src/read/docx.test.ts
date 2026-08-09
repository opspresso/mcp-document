/**
 * The walk over `word/document.xml`, which decides what a DOCX says.
 *
 * Tested on the XML rather than on a file: the zip layer has its own tests, and
 * every decision worth making here is about which element means a line, a cell
 * or nothing at all.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { documentXmlToText } from "./docx.js";

const paragraph = (...runs: string[]) =>
  `<w:p>${runs.map((run) => `<w:r><w:t>${run}</w:t></w:r>`).join("")}</w:p>`;

test("each paragraph is a line", () => {
  const { text, paragraphs } = documentXmlToText(
    `<w:document><w:body>${paragraph("first")}${paragraph("second")}</w:body></w:document>`,
  );
  assert.equal(text, "first\nsecond");
  assert.equal(paragraphs, 2);
});

test("runs inside one paragraph join without a gap", () => {
  // Word splits a sentence across runs at every formatting change, so a space
  // between them would appear in the middle of words a user typed together.
  assert.equal(documentXmlToText(paragraph("한", "글", " 문서")).text, "한글 문서");
});

test("tabs and breaks survive as themselves", () => {
  const xml = "<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r></w:p>";
  assert.equal(documentXmlToText(xml).text, "a\tb\nc");
});

test("a heading style becomes its Markdown prefix", () => {
  const xml =
    '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Background</w:t></w:r></w:p>';
  assert.equal(documentXmlToText(xml).text, "## Background");
});

test("a list paragraph is marked, whatever its numbering says", () => {
  const xml =
    "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>";
  assert.equal(documentXmlToText(xml).text, "- item");
});

test("a table row is one line with its cells separated", () => {
  const xml =
    `<w:tbl><w:tr><w:tc>${paragraph("a")}</w:tc><w:tc>${paragraph("b")}</w:tc></w:tr>` +
    `<w:tr><w:tc>${paragraph("c")}</w:tc><w:tc>${paragraph("d")}</w:tc></w:tr></w:tbl>`;
  assert.equal(documentXmlToText(xml).text, "a | b\nc | d");
});

test("a cell holding two paragraphs stays one cell", () => {
  // The regression a naive `</w:p>` → newline produces: a multi-paragraph cell
  // breaks its own row in half, and every column after it shifts.
  const xml = `<w:tbl><w:tr><w:tc>${paragraph("one")}${paragraph("two")}</w:tc><w:tc>${paragraph("x")}</w:tc></w:tr></w:tbl>`;
  assert.equal(documentXmlToText(xml).text, "one two | x");
});

test("field codes and deleted text are not the document's text", () => {
  const xml =
    "<w:p><w:r><w:instrText> HYPERLINK \\l bookmark </w:instrText></w:r>" +
    "<w:del><w:r><w:delText>removed</w:delText></w:r></w:del>" +
    "<w:r><w:t>kept</w:t></w:r></w:p>";
  assert.equal(documentXmlToText(xml).text, "kept");
});

test("runs of blank paragraphs collapse, and the ends are trimmed", () => {
  const xml = `${paragraph("")}${paragraph("")}${paragraph("a")}${paragraph("")}${paragraph("")}${paragraph("b")}${paragraph("")}`;
  assert.equal(documentXmlToText(xml).text, "a\n\nb");
});

test("a document with no text comes back empty rather than with markup in it", () => {
  const { text } = documentXmlToText(
    '<w:document><w:body><w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p></w:body></w:document>',
  );
  assert.equal(text, "");
});

test("`xml:space` and escaped characters are handled", () => {
  const xml = '<w:p><w:r><w:t xml:space="preserve">a &amp; b </w:t><w:t>c</w:t></w:r></w:p>';
  assert.equal(documentXmlToText(xml).text, "a & b c");
});
