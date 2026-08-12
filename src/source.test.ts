/**
 * The decisions on the way in. None of them touch the network — and now there is
 * no network here to touch: the URL side left with the outbound boundary, which
 * was a byte-for-byte copy of the caller's.
 *
 * The base64 cases are the ones that matter. `Buffer.from(x, "base64")` drops
 * every character outside the alphabet rather than failing, so without the
 * check here a caller who sent JSON — or a `data:` URL, or the document's own
 * text — would have it silently turned into bytes and handed to a detector as
 * though it were a file.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { decodeBase64, loadSource, SourceError } from "./source.js";


test("base64 round-trips", () => {
  const bytes = decodeBase64(Buffer.from("hello 한글").toString("base64"));
  assert.equal(Buffer.from(bytes).toString("utf8"), "hello 한글");
});

test("whitespace inside base64 is tolerated, since a wrapped payload is still base64", () => {
  const wrapped = Buffer.from("hello").toString("base64").split("").join("\n");
  assert.equal(Buffer.from(decodeBase64(wrapped)).toString("utf8"), "hello");
});

test("anything that is not base64 is refused rather than silently decoded", () => {
  for (const value of ['{"a":1}', "hello world!", "****", "abcde", ""]) {
    assert.throws(
      () => decodeBase64(value),
      SourceError,
      `expected ${JSON.stringify(value)} to be refused`,
    );
  }
});

test("a data: URL is refused with the fix in the message", () => {
  assert.throws(
    () => decodeBase64("data:application/pdf;base64,JVBERi0="),
    (error: unknown) => error instanceof SourceError && /after the comma/.test(error.message),
  );
});

test("content is required, and says what it wants", () => {
  assert.throws(
    () => loadSource({}),
    (error: unknown) => error instanceof SourceError && /base64/.test(error.message),
  );
});

test("an inline document is labelled by its filename, or said to be uploaded", () => {
  const named = loadSource({ content: "aGVsbG8=", filename: "report.docx" });
  assert.equal(named.label, "report.docx");
  assert.equal(named.filename, "report.docx");
  // Nothing is invented when the caller named nothing: a made-up filename would
  // appear in the provenance header as if it were a fact.
  const anonymous = loadSource({ content: "aGVsbG8=" });
  assert.equal(anonymous.label, "an uploaded document");
  assert.equal(anonymous.filename, undefined);
  assert.equal(anonymous.mimeType, "");
});
