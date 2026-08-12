/**
 * The RTF reader, which exists for a failure mode the others do not have.
 *
 * RTF is a text file, so without this it is not refused — it is *read as plain
 * text* and reaches the model as thousands of control words with the prose
 * scattered through them. Every case below is one of the ways that garbage gets
 * in: a font table's contents, a generator's version string, an escape left
 * unresolved.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { rtfToText, RtfError } from "./rtf.js";

const rtf = (body: string) => new TextEncoder().encode(`{\\rtf1\\ansi${body}}`);
/** Latin-1, which is what a writer emitting `\\'hh` produced. */
const latin1 = (source: string) => new Uint8Array(Buffer.from(source, "latin1"));

test("prose comes back without its control words", () => {
  assert.equal(rtfToText(rtf("\\pard Hello world.\\par")).text, "Hello world.");
});

test("a font table's contents stay out of the text", () => {
  // The defect a naive control-word strip produces: "Times New Roman" in the
  // middle of somebody's letter.
  const bytes = rtf("{\\fonttbl{\\f0\\froman Times New Roman;}}\\pard Body text.\\par");
  const { text } = rtfToText(bytes);
  assert.equal(text, "Body text.");
  assert.doesNotMatch(text, /Times/);
});

test("an ignorable destination is skipped whole, whatever it is named", () => {
  const bytes = rtf("{\\*\\generator Riched20 10.0;}\\pard Real content.\\par");
  const { text } = rtfToText(bytes);
  assert.equal(text, "Real content.");
  assert.doesNotMatch(text, /Riched20/);
});

test("a colour table and a stylesheet are destinations too", () => {
  const bytes = rtf(
    "{\\colortbl;\\red0\\green0\\blue0;}{\\stylesheet{\\s0 Normal;}}\\pard Text.\\par",
  );
  assert.equal(rtfToText(bytes).text, "Text.");
});

test("a hex escape becomes its character", () => {
  // Windows-1252 é, which is what `\\'e9` means in an `\\ansi` document.
  assert.equal(rtfToText(latin1("{\\rtf1\\ansi caf\\'e9\\par}")).text, "café");
});

test("a unicode escape wins over the fallback that follows it", () => {
  // `\\u54620?` is 한 with `?` as the substitute for readers that cannot show it.
  const { text } = rtfToText(rtf("\\pard \\u54620?\\u44544?\\par"));
  assert.equal(text, "한글");
  assert.doesNotMatch(text, /\?/);
});

test("a negative unicode code point is the signed 16-bit form", () => {
  // Writers emit negative numbers for anything past U+7FFF; -11384 is 54152.
  assert.equal(rtfToText(rtf("\\pard \\u-11384?\\par")).text, String.fromCodePoint(54152));
});

test("escaped braces and backslashes are literal", () => {
  assert.equal(rtfToText(rtf("\\pard a \\{b\\} \\\\c\\par")).text, "a {b} \\c");
});

test("paragraphs and tabs become the lines and columns they are", () => {
  const { text } = rtfToText(rtf("\\pard one\\par two\\line three\\par name\\tab value\\par"));
  assert.equal(text, "one\ntwo\nthree\nname\tvalue");
});

test("table cells are separated the way every other reader separates them", () => {
  assert.match(rtfToText(rtf("\\pard A\\cell B\\cell\\row")).text, /A \| B/);
});

test("a file that does not begin with the header is refused", () => {
  assert.throws(() => rtfToText(new TextEncoder().encode("just text")), RtfError);
});

test("a document with only markup is refused rather than returned empty", () => {
  assert.throws(() => rtfToText(rtf("{\\fonttbl{\\f0 Arial;}}")), RtfError);
});
