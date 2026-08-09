/**
 * The walk over `Contents/sectionN.xml`, and the order those sections are read
 * in — which is numeric, because a lexical sort puts `section10` between
 * `section1` and `section2` and silently reorders any document long enough to
 * have ten of them.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { sectionsOf, sectionXmlToText } from "./hwpx.js";

const paragraph = (...runs: string[]) =>
  `<hp:p>${runs.map((run) => `<hp:run><hp:t>${run}</hp:t></hp:run>`).join("")}</hp:p>`;

test("each paragraph is a line", () => {
  assert.equal(sectionXmlToText(paragraph("첫째") + paragraph("둘째")), "첫째\n둘째");
});

test("runs inside one paragraph join without a gap", () => {
  assert.equal(sectionXmlToText(paragraph("한", "글", " 문서")), "한글 문서");
});

test("tabs and line breaks survive", () => {
  const xml = "<hp:p><hp:run><hp:t>a</hp:t><hp:tab/><hp:t>b</hp:t><hp:lineBreak/><hp:t>c</hp:t></hp:run></hp:p>";
  assert.equal(sectionXmlToText(xml), "a\tb\nc");
});

test("a table row is one line with its cells separated", () => {
  const xml =
    `<hp:tbl><hp:tr><hp:tc><hp:subList>${paragraph("가")}</hp:subList></hp:tc>` +
    `<hp:tc><hp:subList>${paragraph("나")}</hp:subList></hp:tc></hp:tr></hp:tbl>`;
  assert.equal(sectionXmlToText(xml), "가 | 나");
});

test("elements are matched on their local name, not on the `hp:` prefix", () => {
  // The prefix is conventional, not required. Keying on it would return "no
  // text" for a valid document rather than an error anybody could act on.
  assert.equal(sectionXmlToText("<p><run><t>bound elsewhere</t></run></p>"), "bound elsewhere");
  assert.equal(sectionXmlToText("<x:p><x:run><x:t>또는 이렇게</x:t></x:run></x:p>"), "또는 이렇게");
});

test("everything outside a text element is dropped", () => {
  const xml =
    `<hp:sec><hp:secPr><hp:pagePr width="59528"/></hp:secPr>${paragraph("본문")}</hp:sec>`;
  assert.equal(sectionXmlToText(xml), "본문");
});

test("sections are ordered by their number, not by their name", () => {
  const entries = [
    "Contents/section10.xml",
    "Contents/section2.xml",
    "Contents/section0.xml",
    "Contents/header.xml",
    "mimetype",
  ].map((name) => ({ name, originalSize: 1 }));
  assert.deepEqual(sectionsOf(entries), [
    "Contents/section0.xml",
    "Contents/section2.xml",
    "Contents/section10.xml",
  ]);
});
