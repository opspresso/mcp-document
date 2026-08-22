import { strict as assert } from "node:assert";
import { test } from "node:test";
import { parseMarkdown } from "./markdown.js";
import { validateRenderedDocument, ValidationError } from "./validate.js";
import { buildZip, readEntries } from "./zip.js";
import { renderDocx } from "./write/docx.js";
import { renderPdf } from "./write/pdf.js";
import { renderPptx } from "./write/pptx/index.js";
import { renderXlsx } from "./write/xlsx.js";

const CREATED = "2026-08-23T00:00:00.000Z";

test("a generated DOCX reopens and all internal relationships resolve", async () => {
  const bytes = renderDocx(parseMarkdown("# 보고서\n\n[근거](https://example.com)"), {
    title: "보고서",
    created: CREATED,
  });
  const report = await validateRenderedDocument("docx", bytes);
  assert.equal(report.structure, "passed");
  assert.equal(report.content, "reopened");
  assert.equal(report.externalRelationships, 1);
});

test("a generated PPTX reopens and keeps its declared slide relationships", async () => {
  const rendered = renderPptx(parseMarkdown("# 표지\n\n## 결론\n\n본문"), {
    title: "발표",
    created: CREATED,
  });
  const report = await validateRenderedDocument("pptx", rendered.bytes);
  assert.equal(report.structure, "passed");
  assert.equal(report.content, "reopened");
});

test("a missing relationship target is a hard validation failure", async () => {
  const original = renderDocx(parseMarkdown("[근거](https://example.com)"), {
    title: "보고서",
    created: CREATED,
  });
  const parts = Object.fromEntries(readEntries(original, [
    "[Content_Types].xml",
    "_rels/.rels",
    "docProps/core.xml",
    "docProps/app.xml",
    "word/document.xml",
    "word/styles.xml",
    "word/_rels/document.xml.rels",
    "word/header1.xml",
    "word/footer1.xml",
  ]));
  parts["word/_rels/document.xml.rels"] = new TextEncoder().encode(
    '<Relationships><Relationship Id="bad" Target="missing.xml"/></Relationships>',
  );
  await assert.rejects(() => validateRenderedDocument("docx", buildZip(parts)), ValidationError);
});

test("a generated PDF reopens with the page count the renderer reported", async () => {
  const rendered = await renderPdf(parseMarkdown("# 보고서\n\n본문"), {
    title: "보고서",
    created: new Date(CREATED),
  });
  const report = await validateRenderedDocument("pdf", rendered.bytes, rendered.pages);
  assert.equal(report.structure, "passed");
  assert.equal(report.content, "not_checked");
});

test("a generated XLSX reopens and all worksheet relationships resolve", async () => {
  const rendered = renderXlsx([{ name: "Data", rows: [["value"], [1]] }], {
    title: "Data",
    created: CREATED,
  });
  const report = await validateRenderedDocument("xlsx", rendered.bytes);
  assert.equal(report.structure, "passed");
  assert.equal(report.content, "reopened");
});
