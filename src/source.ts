/**
 * Where the bytes of a document come from.
 *
 * One way in now: base64 the caller already holds. `url` is gone, and with it
 * this server's entire outbound boundary — an SSRF guard, a pinned-DNS fetch, a
 * redirect policy, all of it a byte-for-byte copy of the sibling's. One copy of
 * that code is the right number, and it belongs where the addresses a *model*
 * chooses are already governed rather than in a parser.
 *
 * What comes out is bytes, whatever the caller declared about them, and a label
 * honest enough to put in front of the extracted text.
 */

import { DocumentError } from "./errors.js";

export class SourceError extends DocumentError {}

export interface DocumentSource {
  bytes: Uint8Array;
  /** Declared media type, lower-cased. Empty when nothing declared one. */
  mimeType: string;
  /** Charset the caller declared, when it declared one. */
  charset?: string;
  /** What the provenance header names as the origin of this text. */
  label: string;
  /** Filename, for the extension hint the detector may need. */
  filename?: string;
}

/**
 * Decode base64, and refuse anything that is not.
 *
 * `Buffer.from(x, "base64")` silently drops anything that is not a base64
 * character, so a caller that sent JSON, or a data: URL, or the document's text
 * would get bytes back rather than an error — and those bytes would then be
 * offered to a detector as if they were a file. Checking the alphabet first is
 * what turns that into the message the caller needs.
 */
export function decodeBase64(content: string): Uint8Array {
  const compact = content.replace(/\s+/g, "");
  if (compact === "") {
    throw new SourceError("`content` is empty");
  }
  if (compact.startsWith("data:")) {
    throw new SourceError(
      "`content` must be base64 alone, without a `data:` URL prefix — pass the part after the comma",
    );
  }
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(compact) || compact.length % 4 !== 0) {
    throw new SourceError("`content` is not valid base64");
  }
  return Buffer.from(compact, "base64");
}

/** The document to read, from the caller's own bytes. */
export function loadSource(args: {
  content?: string | undefined;
  filename?: string | undefined;
}): DocumentSource {
  const { content, filename } = args;
  if (!content) {
    throw new SourceError("`content` is required: the document's bytes, base64-encoded");
  }
  return {
    bytes: decodeBase64(content),
    mimeType: "",
    label: filename ?? "an uploaded document",
    ...(filename ? { filename } : {}),
  };
}
