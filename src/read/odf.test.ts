import { strict as assert } from "node:assert";
import { test } from "node:test";
import { buildZip, stored } from "../zip.js";
import { odfKindOf, odfToText, OdfError } from "./odf.js";

const utf8 = (value: string) => new TextEncoder().encode(value);

const MIME = {
  text: "application/vnd.oasis.opendocument.text",
  spreadsheet: "application/vnd.oasis.opendocument.spreadsheet",
  presentation: "application/vnd.oasis.opendocument.presentation",
};

function odf(mimetype: string, body: string): Uint8Array {
  return buildZip({
    // First and stored, as the packaging rule requires.
    mimetype: stored(utf8(mimetype)),
    "content.xml": utf8(
      `<?xml version="1.0"?><office:document-content><office:body>${body}</office:body></office:document-content>`,
    ),
  });
}

test("the package says which kind it is", () => {
  assert.equal(odfKindOf(MIME.text), "text");
  assert.equal(odfKindOf(MIME.spreadsheet), "spreadsheet");
  assert.equal(odfKindOf(MIME.presentation), "presentation");
  assert.equal(odfKindOf("application/zip"), undefined);
});

test("a text document comes back as paragraphs", () => {
  const bytes = odf(MIME.text, `<text:h>보고서</text:h><text:p>본문입니다.</text:p>`);
  const { text, kind } = odfToText(bytes);
  assert.equal(kind, "text");
  assert.equal(text, "보고서\n본문입니다.");
});

test("a spreadsheet names each sheet and separates cells", () => {
  const row = `<table:table-row><table:table-cell><text:p>A</text:p></table:table-cell><table:table-cell><text:p>B</text:p></table:table-cell></table:table-row>`;
  const bytes = odf(MIME.spreadsheet, `<table:table table:name="Data">${row}</table:table>`);
  const { text, kind, parts } = odfToText(bytes);
  assert.equal(kind, "spreadsheet");
  assert.equal(parts, 1);
  assert.equal(text, "## Data\nA | B");
});

test("a presentation numbers its slides", () => {
  const bytes = odf(
    MIME.presentation,
    `<draw:page><text:p>first</text:p></draw:page><draw:page><text:p>second</text:p></draw:page>`,
  );
  const { text, parts } = odfToText(bytes);
  assert.equal(parts, 2);
  // A blank line between slides: the heading flushes the previous one, which is
  // what keeps two decks' worth of text from reading as one continuous page.
  assert.equal(text, "## Slide 1\nfirst\n\n## Slide 2\nsecond");
});

test("an encoded run of spaces survives, since XML would have collapsed it", () => {
  const bytes = odf(MIME.text, `<text:p>a<text:s text:c="3"/>b</text:p>`);
  // Normalised to one space on the way out, but not lost entirely.
  assert.equal(odfToText(bytes).text, "a b");
});

test("a tab is kept as a tab, because it is how columns are laid out", () => {
  const bytes = odf(MIME.text, `<text:p>name<text:tab/>value</text:p>`);
  assert.equal(odfToText(bytes).text, "name\tvalue");
});

test("a package that does not say what it is, is refused", () => {
  const bytes = buildZip({ mimetype: stored(utf8("application/zip")), "content.xml": utf8("<x/>") });
  assert.throws(() => odfToText(bytes), OdfError);
});

test("an empty document is refused rather than returned empty", () => {
  assert.throws(() => odfToText(odf(MIME.text, "")), OdfError);
});

test("a run of repeated cells occupies its columns, so the values after it do not shift", () => {
  // LibreOffice stores a run of identical or empty cells once, with a count.
  // Emitting one separator for the element puts every later value in a column
  // it does not belong to — a table that still looks like a table and says
  // something else.
  const row =
    `<table:table-row><table:table-cell><text:p>A</text:p></table:table-cell>` +
    `<table:table-cell table:number-columns-repeated="3"/>` +
    `<table:table-cell><text:p>B</text:p></table:table-cell></table:table-row>`;
  const bytes = odf(MIME.spreadsheet, `<table:table table:name="Data">${row}</table:table>`);
  // Five columns: B is the fifth, not the second.
  assert.equal(odfToText(bytes).text, "## Data\nA | | | | B");
});

test("the repeat that pads a row to the sheet's width costs nothing", () => {
  // Every ODS row ends with one of these. Expanding it eagerly would draw
  // sixteen thousand empty columns for a row holding one value.
  const row =
    `<table:table-row><table:table-cell><text:p>A</text:p></table:table-cell>` +
    `<table:table-cell table:number-columns-repeated="16384"/></table:table-row>`;
  const bytes = odf(MIME.spreadsheet, `<table:table table:name="Data">${row}</table:table>`);
  assert.equal(odfToText(bytes).text, "## Data\nA");
});

test("a cell covered by a merge still holds its column", () => {
  const row =
    `<table:table-row>` +
    `<table:table-cell table:number-columns-spanned="2"><text:p>A</text:p></table:table-cell>` +
    `<table:covered-table-cell/>` +
    `<table:table-cell><text:p>B</text:p></table:table-cell></table:table-row>`;
  const bytes = odf(MIME.spreadsheet, `<table:table table:name="Data">${row}</table:table>`);
  assert.equal(odfToText(bytes).text, "## Data\nA | | B");
});

test("text a change tracker deleted is not the document's text", () => {
  // `text:tracked-changes` sits at the top of the body and holds whole deleted
  // paragraphs. Returning them puts text the author removed at the top of the
  // document, as if it were current.
  const bytes = odf(
    MIME.text,
    `<office:text><text:tracked-changes><text:changed-region><text:deletion>` +
      `<office:change-info><dc:creator>김철수</dc:creator><dc:date>2026-01-02</dc:date></office:change-info>` +
      `<text:p>removed</text:p></text:deletion></text:changed-region></text:tracked-changes>` +
      `<text:p>kept</text:p></office:text>`,
  );
  assert.equal(odfToText(bytes).text, "kept");
});

test("a comment is not the paragraph it is anchored in", () => {
  const bytes = odf(
    MIME.text,
    `<office:text><text:p>본문<office:annotation><dc:creator>김철수</dc:creator>` +
      `<text:p>확인 필요</text:p></office:annotation>입니다.</text:p></office:text>`,
  );
  assert.equal(odfToText(bytes).text, "본문입니다.");
});

test("a footnote does not split the paragraph it hangs from", () => {
  // The footnote's own `</text:p>` used to flush, cutting the host paragraph in
  // half with the note's text between the pieces.
  const bytes = odf(
    MIME.text,
    `<office:text><text:p>앞<text:note text:note-class="footnote">` +
      `<text:note-citation>1</text:note-citation>` +
      `<text:note-body><text:p>각주 본문</text:p></text:note-body>` +
      `</text:note>뒤</text:p></office:text>`,
  );
  assert.equal(odfToText(bytes).text, "앞뒤");
});

test("speaker notes are not slide content", () => {
  // `document.ts` promises a deck comes back without speaker notes. The ODP
  // path said so and did the opposite.
  const bytes = odf(
    MIME.presentation,
    `<draw:page><text:p>slide</text:p>` +
      `<presentation:notes><draw:frame><draw:text-box><text:p>말할 것</text:p></draw:text-box></draw:frame></presentation:notes>` +
      `</draw:page>`,
  );
  assert.equal(odfToText(bytes).text, "## Slide 1\nslide");
});
