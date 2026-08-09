/**
 * The decisions on the way in, none of which touch the network.
 *
 * The base64 cases are the ones that matter. `Buffer.from(x, "base64")` drops
 * every character outside the alphabet rather than failing, so without the
 * check here a caller who sent JSON — or a `data:` URL, or the document's own
 * text — would have it silently turned into bytes and handed to a detector as
 * though it were a file.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { decodeBase64, filenameFromUrl, loadSource, parseContentType, SourceError } from "./source.js";

test("a content type is split from its charset", () => {
  assert.deepEqual(parseContentType("application/pdf"), { mimeType: "application/pdf" });
  assert.deepEqual(parseContentType("text/plain; charset=EUC-KR"), {
    mimeType: "text/plain",
    charset: "euc-kr",
  });
  assert.deepEqual(parseContentType('text/csv;charset="utf-8"'), {
    mimeType: "text/csv",
    charset: "utf-8",
  });
  assert.deepEqual(parseContentType(null), { mimeType: "" });
});

test("a filename hint is taken from the URL only when it carries an extension", () => {
  assert.equal(filenameFromUrl("https://example.com/a/report.hwp"), "report.hwp");
  assert.equal(filenameFromUrl("https://example.com/a/report.hwp?v=2"), "report.hwp");
  assert.equal(filenameFromUrl("https://example.com/%ED%95%9C%EA%B8%80.docx"), "한글.docx");
  // Nothing to learn from these, and offering them would make the detector
  // confident about something it was told nothing about.
  assert.equal(filenameFromUrl("https://example.com/download"), undefined);
  assert.equal(filenameFromUrl("https://example.com/"), undefined);
  assert.equal(filenameFromUrl("not a url"), undefined);
});

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

test("exactly one source is required", async () => {
  await assert.rejects(() => loadSource({}), SourceError);
  await assert.rejects(
    () => loadSource({ url: "https://example.com/a.pdf", content: "AAAA" }),
    (error: unknown) => error instanceof SourceError && /not both/.test(error.message),
  );
});

test("an inline document is labelled by its filename, or said to be uploaded", async () => {
  const named = await loadSource({ content: "aGVsbG8=", filename: "report.docx" });
  assert.equal(named.label, "report.docx");
  assert.equal(named.filename, "report.docx");
  // Nothing is invented when the caller named nothing: a made-up filename would
  // appear in the provenance header as if it were a fact.
  const anonymous = await loadSource({ content: "aGVsbG8=" });
  assert.equal(anonymous.label, "an uploaded document");
  assert.equal(anonymous.filename, undefined);
  assert.equal(anonymous.mimeType, "");
});
