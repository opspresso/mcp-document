/**
 * Boot-time validation, which is the only thing standing between a mistyped
 * environment and a server that runs while doing the wrong thing.
 *
 * Small, because the configuration is: the bucket, its prefix, the region and
 * the download TTL all left when the renderer stopped storing anything. The
 * prefix cases had teeth — it was joined with a tenant to make an S3 key, so a
 * segment a path reads as movement was a way for one deployment to write under
 * another's documents — and they are gone with the key they built.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { ConfigError, loadConfig } from "./config.js";

test("the defaults are the deployment this is built for", () => {
  const config = loadConfig({});
  assert.equal(config.port, 3000);
  // Unset means open, which is the deployment this is built for: a Deployment
  // behind a ClusterIP with no ingress. See auth.ts.
  assert.equal(config.apiKey, undefined);
});

test("nothing is required, because nothing outside this process is", () => {
  // It used to refuse to boot without DOCUMENT_BUCKET — half a server was not a
  // state worth being able to deploy. There is no half now.
  assert.doesNotThrow(() => loadConfig({}));
});

test("a blank api key reads as unset rather than as an empty secret", () => {
  assert.equal(loadConfig({ MCP_API_KEY: "   " }).apiKey, undefined);
  assert.equal(loadConfig({ MCP_API_KEY: " s3cret " }).apiKey, "s3cret");
});

test("a non-integer or non-positive port is refused rather than coerced", () => {
  for (const value of ["0", "-1", "1.5", "many"]) {
    assert.throws(() => loadConfig({ PORT: value }), ConfigError);
  }
});
