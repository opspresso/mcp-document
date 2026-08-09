/**
 * HWPX generation, checked by reading it back.
 *
 * The round trip is what it is elsewhere. What is different here is how much
 * these tests are *not* able to say: HWPX has one reader that matters, and
 * nothing in this file knows whether 한글 opens the result. So the assertions
 * cover the parts a wrong file gets wrong in a way that is checkable — the
 * package layout, the id references resolving against the header, a cell
 * holding a paragraph — and the rest is a manual step the README names.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { hwpxToText } from "../read/hwpx.js";
import { parseMarkdown } from "../markdown.js";
import { detect } from "../detect.js";
import { listEntries, readEntries } from "../zip.js";
import { renderHwpx } from "./hwpx.js";

const OPTIONS = { title: "test", created: "2026-08-05T00:00:00Z", application: "0.1.0" };

function build(markdown: string): Uint8Array {
  return renderHwpx(parseMarkdown(markdown), OPTIONS);
}

function roundTrip(markdown: string): string {
  return hwpxToText(build(markdown)).text;
}

function partOf(bytes: Uint8Array, name: string): string {
  const part = readEntries(bytes, [name]).get(name);
  assert.ok(part, `expected the archive to hold ${name}`);
  return new TextDecoder().decode(part);
}

test("the package has the parts OWPML requires, with mimetype first and stored", () => {
  const bytes = build("# 제목\n\n본문");
  const names = listEntries(bytes).map((entry) => entry.name);
  assert.equal(names[0], "mimetype");
  for (const required of [
    "version.xml",
    "settings.xml",
    "META-INF/container.xml",
    "META-INF/manifest.xml",
    "Contents/content.hpf",
    "Contents/header.xml",
    "Contents/section0.xml",
  ]) {
    assert.ok(names.includes(required), `the package should hold ${required}`);
  }
  // Stored, so the type is readable at a fixed offset without inflating.
  assert.equal(
    Buffer.from(bytes.subarray(30, 38)).toString("latin1"),
    "mimetype",
    "mimetype must be the first entry",
  );
  assert.equal(partOf(bytes, "mimetype"), "application/hwp+zip");
});

test("what this writes, this server recognises as HWPX", () => {
  assert.equal(detect(build("본문"), "", undefined).format, "hwpx");
});

test("headings and paragraphs come back as their text", () => {
  assert.equal(roundTrip("# 분기 보고서\n\n첫 문단\n\n둘째 문단"), "분기 보고서\n첫 문단\n둘째 문단");
});

test("lists come back as the markers they were written with", () => {
  assert.equal(roundTrip("- 하나\n- 둘"), "- 하나\n- 둘");
  assert.equal(roundTrip("1. 첫째\n2. 둘째"), "1. 첫째\n2. 둘째");
});

test("a table round-trips, and every cell holds a paragraph", () => {
  const bytes = build("| 이름 | 값 |\n|---|---|\n| 가 | 1 |");
  assert.equal(hwpxToText(bytes).text, "이름 | 값\n가 | 1");
  const section = partOf(bytes, "Contents/section0.xml");
  // A `hp:tc` with no paragraph inside is what makes 한글 refuse a file.
  assert.equal(/<hp:tc[^>]*>(?:(?!<hp:p[ >]).)*<\/hp:tc>/s.test(section), false);
  assert.match(section, /<hp:cellAddr colAddr="1" rowAddr="1"\/>/);
});

test("every id a paragraph or run refers to exists in the header", () => {
  // The reference that does not resolve is the failure mode of this format:
  // nothing in the body says what a style is, only which numbered one it wants.
  const bytes = build("# 제목\n\n**굵게** *기울임* `코드` [링크](https://example.com)\n\n> 인용\n\n- 목록\n\n```\ncode\n```");
  const header = partOf(bytes, "Contents/header.xml");
  const section = partOf(bytes, "Contents/section0.xml");
  const declared = (xml: string, pattern: RegExp): Set<string> =>
    new Set([...xml.matchAll(pattern)].map((match) => match[1]!));
  const charIds = declared(header, /<hh:charPr id="(\d+)"/g);
  const paraIds = declared(header, /<hh:paraPr id="(\d+)"/g);
  for (const id of declared(section, /charPrIDRef="(\d+)"/g)) {
    assert.ok(charIds.has(id), `charPr ${id} is referenced but not declared`);
  }
  for (const id of declared(section, /paraPrIDRef="(\d+)"/g)) {
    assert.ok(paraIds.has(id), `paraPr ${id} is referenced but not declared`);
  }
});

test("the section properties are written exactly once", () => {
  // They ride inside the first run of the document. Twice is a second section
  // definition; never is a document with no page size.
  const section = partOf(build("# 제목\n\n본문\n\n또 본문"), "Contents/section0.xml");
  assert.equal(section.match(/<hp:secPr /g)?.length, 1);
});

test("an empty document still carries its section", () => {
  const bytes = renderHwpx({ blocks: [] }, OPTIONS);
  const section = partOf(bytes, "Contents/section0.xml");
  assert.equal(section.match(/<hp:secPr /g)?.length, 1);
  assert.match(section, /<hp:p /);
});

test("a tab inside text becomes the element HWPX uses for one", () => {
  const section = partOf(build("before\tafter"), "Contents/section0.xml");
  assert.match(section, /<hp:tab /);
  assert.equal(section.includes("before\tafter"), false);
});

test("XML metacharacters in the text do not become markup", () => {
  assert.equal(roundTrip('a < b & c > d "quoted"'), 'a < b & c > d "quoted"');
});

test("the title lands in the package metadata, escaped", () => {
  const bytes = renderHwpx(parseMarkdown("본문"), { ...OPTIONS, title: "A & B <보고서>" });
  assert.match(partOf(bytes, "Contents/content.hpf"), /<opf:title>A &amp; B &lt;보고서&gt;<\/opf:title>/);
});

test("the same input twice produces the same bytes", () => {
  assert.deepEqual(build("# same\n\ntext"), build("# same\n\ntext"));
});
