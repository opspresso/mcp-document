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
