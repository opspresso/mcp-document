/**
 * Time-sortable identifiers for the keys written documents land under.
 *
 * ULID rather than a UUID because the key is the only ordering S3 has: a
 * prefix listing walks lexicographically, so a key that carries the time is
 * what makes "this project's documents, newest first" answerable at all —
 * without it, finding a document written this morning means listing everything
 * the project has ever produced and reading each object's metadata.
 *
 * Hand-written rather than a dependency, and the same twenty lines
 * `mcp-memory/src/id.ts` carries: the surface is one function and the encoding
 * is fixed by a spec that cannot drift under it.
 */

import { randomBytes } from "node:crypto";

const ALPHABET = "0123456789ABCDEFGHJKMNPQRSTVWXYZ";
const TIME_CHARS = 10;
const RANDOM_CHARS = 16;
/** 48 bits, the ULID timestamp width. */
const MAX_TIME = 2 ** 48 - 1;

function encodeTime(ms: number, length: number): string {
  let remaining = ms;
  let out = "";
  for (let index = 0; index < length; index += 1) {
    out = ALPHABET[remaining % 32] + out;
    remaining = Math.floor(remaining / 32);
  }
  return out;
}

function encodeRandom(length: number): string {
  const bytes = randomBytes(length);
  let out = "";
  for (let index = 0; index < length; index += 1) {
    // One byte per character, folded into the alphabet. Wastes three bits of
    // each byte, which costs nothing here and keeps the mapping obvious.
    out += ALPHABET[bytes[index]! % 32];
  }
  return out;
}

/** A ULID for `at` (default now). */
export function ulid(at: number = Date.now()): string {
  if (!Number.isFinite(at) || at < 0 || at > MAX_TIME) {
    throw new RangeError(`timestamp out of ULID range: ${at}`);
  }
  return encodeTime(Math.floor(at), TIME_CHARS) + encodeRandom(RANDOM_CHARS);
}
