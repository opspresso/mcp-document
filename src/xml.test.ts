/**
 * The walker both office readers are built on. Every case here is a way for
 * markup to end up in the output as prose, which is the failure that looks like
 * success: the model reads it, cannot tell it apart from the document, and
 * neither can anyone reviewing the result.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { decodeXmlEntities, escapeXml, walkXml, XmlError, type XmlHandler } from "./xml.js";

interface Event {
  kind: "text" | "open" | "close";
  name?: string;
  value?: string;
  selfClosing?: boolean;
}

function record(xml: string): Event[] {
  const events: Event[] = [];
  const handler: XmlHandler = {
    text: (value) => events.push({ kind: "text", value }),
    open: (name, _attributes, selfClosing) => events.push({ kind: "open", name, selfClosing }),
    close: (name) => events.push({ kind: "close", name }),
  };
  walkXml(xml, handler);
  return events;
}

function textOf(xml: string): string {
  return record(xml)
    .filter((event) => event.kind === "text")
    .map((event) => event.value)
    .join("");
}

test("opens, closes and the characters between them are reported in order", () => {
  assert.deepEqual(record("<a><b>hi</b></a>"), [
    { kind: "open", name: "a", selfClosing: false },
    { kind: "open", name: "b", selfClosing: false },
    { kind: "text", value: "hi" },
    { kind: "close", name: "b" },
    { kind: "close", name: "a" },
  ]);
});

test("a self-closing tag opens and never closes", () => {
  assert.deepEqual(record("<w:tab/>"), [{ kind: "open", name: "w:tab", selfClosing: true }]);
  assert.deepEqual(record("<w:br />"), [{ kind: "open", name: "w:br", selfClosing: true }]);
});

test("attributes are separated from the name", () => {
  const seen: string[] = [];
  walkXml('<w:pStyle w:val="Heading1"/>', {
    text: () => {},
    open: (name, attributes) => seen.push(`${name}|${attributes}`),
    close: () => {},
  });
  assert.deepEqual(seen, ['w:pStyle|w:val="Heading1"']);
});

test("a `>` inside an attribute value does not cut the tag short", () => {
  // The regression this exists for: `<[^>]*>` ends the tag at the quoted `>`,
  // and everything after it — the rest of the attributes, and the tag's own
  // closing bracket — is then reported as the document's prose.
  assert.equal(textOf('<a title="1 > 0">body</a>'), "body");
});

test("declarations and comments contribute nothing", () => {
  assert.equal(textOf('<?xml version="1.0" encoding="UTF-8"?><a>x</a>'), "x");
  assert.equal(textOf("<a>x<!-- a comment -->y</a>"), "xy");
});

test("DTD and entity declarations are refused", () => {
  assert.throws(() => textOf("<!DOCTYPE html><a>x</a>"), XmlError);
  assert.throws(() => textOf('<!ENTITY x "secret"><a>&x;</a>'), XmlError);
});

test("CDATA is text, and an unterminated one takes the rest of the document literally", () => {
  assert.equal(textOf("<a><![CDATA[<b>not a tag</b>]]></a>"), "<b>not a tag</b>");
  assert.equal(textOf("<a><![CDATA[tail"), "tail");
});

test("a document cut mid-tag loses the tag rather than turning it into prose", () => {
  assert.equal(textOf("<a>kept</a><b attr=\"un"), "kept");
});

test("entities are decoded, and an escaped ampersand does not introduce a second one", () => {
  assert.equal(decodeXmlEntities("a &lt; b &amp;&amp; c"), "a < b && c");
  assert.equal(decodeXmlEntities("&amp;lt;"), "&lt;");
  assert.equal(decodeXmlEntities("&#54620;&#xAE00;"), "한글");
  // Not an entity this knows: left alone rather than eaten, so the text still
  // says what the document said.
  assert.equal(decodeXmlEntities("&unknown;"), "&unknown;");
});

test("a numeric reference that cannot be a character is dropped, not made U+FFFD", () => {
  assert.equal(decodeXmlEntities("a&#xD800;b"), "ab");
  assert.equal(decodeXmlEntities("a&#1114112;b"), "ab");
});

test("escaping round-trips through decoding", () => {
  const raw = `<a href="x">&'한글'</a>`;
  assert.equal(decodeXmlEntities(escapeXml(raw)), raw);
});
