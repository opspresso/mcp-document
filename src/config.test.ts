/**
 * Boot-time validation, which is the only thing standing between a mistyped
 * environment and a server that runs while doing the wrong thing.
 *
 * The prefix cases are the ones with teeth: it is joined with a tenant to make
 * an S3 key, so a segment a path reads as movement is a way for one deployment
 * to write under another's documents.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { ConfigError, loadConfig, MAX_DOWNLOAD_TTL_SECONDS } from "./config.js";

const base = { DOCUMENT_BUCKET: "docs" };

test("the defaults are the deployment this is built for", () => {
  const config = loadConfig({ ...base });
  assert.equal(config.port, 3000);
  assert.equal(config.region, "ap-northeast-2");
  assert.equal(config.prefix, "documents");
  assert.equal(config.downloadTtlSeconds, MAX_DOWNLOAD_TTL_SECONDS);
  assert.equal(config.apiKey, undefined);
});

test("a missing bucket stops the rollout rather than the first write", () => {
  assert.throws(() => loadConfig({}), ConfigError);
  assert.throws(() => loadConfig({ DOCUMENT_BUCKET: "   " }), ConfigError);
});

test("an empty prefix means no prefix", () => {
  assert.equal(loadConfig({ ...base, DOCUMENT_PREFIX: "" }).prefix, "");
  assert.equal(loadConfig({ ...base, DOCUMENT_PREFIX: "/" }).prefix, "");
});

test("surrounding slashes are normalised away, so the key joins cleanly", () => {
  assert.equal(loadConfig({ ...base, DOCUMENT_PREFIX: "/a/b/" }).prefix, "a/b");
});

test("a prefix a path would read as movement is refused", () => {
  for (const prefix of ["../escape", "a/../b", "a/./b", "a//b", "-lead", "a b"]) {
    assert.throws(
      () => loadConfig({ ...base, DOCUMENT_PREFIX: prefix }),
      ConfigError,
      `expected ${JSON.stringify(prefix)} to be refused`,
    );
  }
});

test("a TTL past what SigV4 will sign is a configuration error, not a dead link", () => {
  assert.throws(
    () => loadConfig({ ...base, DOWNLOAD_TTL_SECONDS: String(MAX_DOWNLOAD_TTL_SECONDS + 1) }),
    ConfigError,
  );
  assert.equal(loadConfig({ ...base, DOWNLOAD_TTL_SECONDS: "3600" }).downloadTtlSeconds, 3600);
});

test("a non-integer or non-positive number is refused rather than coerced", () => {
  for (const value of ["0", "-1", "1.5", "many"]) {
    assert.throws(() => loadConfig({ ...base, PORT: value }), ConfigError);
  }
});
