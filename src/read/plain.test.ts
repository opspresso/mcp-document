/**
 * Deciding what encoding a pile of text bytes is in, which is the difference
 * between a Korean document and a screen of replacement characters.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { charsetFromDeclaration, decodeText, plainToText } from "./plain.js";

const utf8 = (value: string) => new TextEncoder().encode(value);

test("UTF-8 is the assumption, and it is stated as the one used", () => {
  const result = decodeText(utf8("한글 문서"));
  assert.equal(result.text, "한글 문서");
  assert.equal(result.charset, "utf-8");
});

test("the transport's charset is honoured", () => {
  // EUC-KR for 한글: guessing UTF-8 here returns replacement characters, which
  // is not a result a reader can recover from.
  const eucKr = Uint8Array.from([0xc7, 0xd1, 0xb1, 0xdb]);
  assert.equal(decodeText(eucKr, "euc-kr").text, "한글");
  assert.equal(decodeText(eucKr, "euc-kr").charset, "euc-kr");
});

test("a BOM outranks the header, because it cannot be stale", () => {
  const withBom = Buffer.concat([Uint8Array.from([0xef, 0xbb, 0xbf]), utf8("한글")]);
  const result = decodeText(withBom, "euc-kr");
  assert.equal(result.text, "한글");
  assert.equal(result.charset, "utf-8");
});

test("UTF-16 is read from its BOM, and the BOM does not survive into the text", () => {
  const le = Buffer.from("﻿한글", "utf16le");
  assert.equal(decodeText(le).text, "한글");
  assert.equal(decodeText(le).charset, "utf-16le");
  const be = Buffer.from([0xfe, 0xff, 0xd5, 0x5c, 0xae, 0x00]);
  assert.equal(decodeText(be).text, "한글");
});

test("a declaration inside the document is used when the transport said nothing", () => {
  const xml = Buffer.concat([
    Buffer.from('<?xml version="1.0" encoding="EUC-KR"?><a>', "latin1"),
    Buffer.from([0xc7, 0xd1, 0xb1, 0xdb]),
    Buffer.from("</a>", "latin1"),
  ]);
  assert.equal(charsetFromDeclaration(xml), "euc-kr");
  assert.match(decodeText(xml).text, /한글/);
});

test("an encoding this build does not have falls back rather than failing", () => {
  // Mojibake is something a reader works around; a hard error gives them
  // nothing at all.
  const result = decodeText(utf8("plain"), "x-not-a-charset");
  assert.equal(result.text, "plain");
  assert.equal(result.charset, "utf-8");
});

test("line endings are normalised and trailing spaces go, but blank lines stay", () => {
  // In a plain text file the blank lines are the only structure there is.
  const { text } = plainToText(utf8("\n\ntitle  \r\n\r\nbody\t\n\n"));
  assert.equal(text, "title\n\nbody");
});
