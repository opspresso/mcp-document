/**
 * The one value that decides which prefix a written document lands under. Every
 * case here is a way for a document to be written somewhere it should not be.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { parseTenant, TenantError, TENANT_HEADER } from "./tenant.js";

test("an ordinary project name passes through", () => {
  assert.equal(parseTenant("agent-studio"), "agent-studio");
  assert.equal(parseTenant("  spaced  "), "spaced");
  assert.equal(parseTenant("a.b_c-1"), "a.b_c-1");
});

test("an absent tenant is refused rather than defaulted", () => {
  for (const value of [undefined, "", "   "]) {
    assert.throws(() => parseTenant(value), TenantError);
  }
  // The message has to name what to fix: this is an operator's misconfiguration
  // surfacing inside somebody else's agent run.
  assert.throws(() => parseTenant(undefined), (error: unknown) =>
    error instanceof TenantError && error.message.includes(TENANT_HEADER));
});

test("a repeated header is ambiguous, so it is refused rather than guessed", () => {
  assert.throws(() => parseTenant(["a", "b"]), TenantError);
  // Even when both values agree: a caller sending it twice is a caller whose
  // configuration says something this cannot see.
  assert.throws(() => parseTenant(["a", "a"]), TenantError);
});

test("anything a path would read as movement is refused", () => {
  for (const value of ["..", ".", "a/b", "../a", "a\\b", "a b", "-lead", "/abs"]) {
    assert.throws(
      () => parseTenant(value),
      TenantError,
      `expected ${JSON.stringify(value)} to be refused`,
    );
  }
});

test("an overlong tenant is refused", () => {
  assert.equal(parseTenant("a".repeat(128)).length, 128);
  assert.throws(() => parseTenant("a".repeat(129)), TenantError);
});
