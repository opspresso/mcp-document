/**
 * What a pile of bytes is, and — when it is nothing this reads — what the
 * refusal says.
 *
 * The refusals are tested as carefully as the successes. A model that is told
 * "unsupported" spends another turn guessing; one that is told "this is an
 * Excel workbook" or "use fetch_document for web pages" does not. Every case
 * below is a wrong answer somebody would otherwise have to debug from a tool
 * result that said nothing.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { detect, extensionOf, looksLikeText } from "./detect.js";
import { buildZip, stored } from "./zip.js";

const utf8 = (value: string) => new TextEncoder().encode(value);
const OLE = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]);

function reasonOf(bytes: Uint8Array, mimeType = "", filename?: string): string {
  const detection = detect(bytes, mimeType, filename);
  assert.equal(detection.format, "unsupported", "expected a refusal");
  return detection.format === "unsupported" ? detection.reason : "";
}

test("magic bytes decide, whatever the header said", () => {
  // A PDF behind a download endpoint is served as octet-stream constantly, and
  // a `.hwp` almost always is. A header that disagrees with the file's own
  // first bytes is wrong about the file.
  assert.equal(detect(utf8("%PDF-1.7\n..."), "application/octet-stream", "x.bin").format, "pdf");
  assert.equal(detect(utf8("%PDF-1.4"), "text/plain", "notes.txt").format, "pdf");
});

test("a zip is identified by the part it carries", () => {
  const docx = buildZip({ "word/document.xml": utf8("<w:document/>") });
  const hwpx = buildZip({
    mimetype: stored(utf8("application/hwp+zip")),
    "Contents/section0.xml": utf8("<hp:sec/>"),
  });
  assert.equal(detect(docx, "", undefined).format, "docx");
  assert.equal(detect(hwpx, "", undefined).format, "hwpx");
});

test("a zip that is a different office format is named rather than dismissed", () => {
  const xlsx = buildZip({ "xl/workbook.xml": utf8("<workbook/>") });
  assert.match(reasonOf(xlsx), /Excel workbook \(\.xlsx\)/);
  const pptx = buildZip({ "ppt/presentation.xml": utf8("<p/>") });
  assert.match(reasonOf(pptx), /PowerPoint deck \(\.pptx\)/);
  const epub = buildZip({ "META-INF/container.xml": utf8("<container/>") });
  assert.match(reasonOf(epub), /EPUB/);
  const plain = buildZip({ "notes.txt": utf8("hello") });
  assert.match(reasonOf(plain), /zip archive/);
});

test("a compound file is left for the HWP reader to identify", () => {
  // Every Office binary from 97-2003 is one of these too, and telling them
  // apart means reading the directory — which the HWP reader parses anyway.
  assert.equal(detect(OLE, "", "report.doc").format, "hwp");
  assert.equal(detect(OLE, "application/x-hwp", undefined).format, "hwp");
});

test("HWP 3.0 names its version and the way out of it", () => {
  const hwp3 = utf8("HWP Document File V3.00 ");
  const reason = reasonOf(hwp3, "", "old.hwp");
  assert.match(reason, /HWP 3\.0/);
  assert.match(reason, /\.hwpx/);
});

test("a web page is sent to the server that reads web pages", () => {
  assert.match(reasonOf(utf8("<html><body>hi</body></html>"), "text/html"), /fetch_document/);
  assert.match(reasonOf(utf8("<html/>"), "", "page.htm"), /fetch_document/);
});

test("an image is sent to the tool that returns pictures", () => {
  assert.match(reasonOf(Uint8Array.from([1, 2, 3, 4]), "image/png"), /fetch_image/);
});

test("text is recognised from its bytes when nothing declared it", () => {
  assert.equal(detect(utf8("# Notes\n\nplain text"), "", undefined).format, "text");
  assert.equal(detect(utf8('{"a":1}'), "application/json", undefined).format, "text");
  assert.equal(detect(utf8("a,b,c\n1,2,3"), "", "data.csv").format, "text");
});

test("EUC-KR bytes are text, though they are not valid UTF-8", () => {
  // Decoding is the decoder's problem. Judging text by whether UTF-8 succeeds
  // would refuse every Korean document written before UTF-8 won.
  const eucKr = Uint8Array.from([0xc7, 0xd1, 0xb1, 0xdb, 0x0a, 0x74, 0x65, 0x78, 0x74]);
  assert.equal(detect(eucKr, "", "memo.txt").format, "text");
  assert.equal(looksLikeText(eucKr), true);
});

test("binary with no signature is refused, and says what was declared", () => {
  const binary = Uint8Array.from(Array.from({ length: 64 }, (_, index) => index % 7));
  assert.match(reasonOf(binary, "application/octet-stream"), /application\/octet-stream/);
  assert.equal(looksLikeText(binary), false);
});

test("an empty document is empty, not unrecognised", () => {
  assert.match(reasonOf(new Uint8Array(0)), /empty/);
});

test("a name that promises one thing and bytes that are another says so", () => {
  assert.match(reasonOf(utf8("hello"), "", "report.pdf"), /named \.pdf, but its contents are not/);
  assert.match(
    reasonOf(utf8("hello"), "application/pdf"),
    /served as application\/pdf, but its contents are not/,
  );
});

test("a legacy binary named by its extension is named in the refusal", () => {
  assert.match(reasonOf(utf8("hello"), "", "old.rtf"), /Rich Text Format/);
  assert.match(reasonOf(utf8("hello"), "", "sheet.xls"), /Excel 97-2003/);
});

test("extensions are read off the end of a name, case-insensitively", () => {
  assert.equal(extensionOf("report.HWPX"), "hwpx");
  assert.equal(extensionOf("a.b.docx"), "docx");
  assert.equal(extensionOf("noextension"), undefined);
  assert.equal(extensionOf(undefined), undefined);
});
