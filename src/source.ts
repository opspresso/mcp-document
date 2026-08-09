/**
 * Where the bytes of a document come from.
 *
 * Two ways in, and they are not interchangeable. A `url` is fetched through the
 * outbound boundary (`publicFetch.ts`), because the address is one the *model*
 * chose and nobody approved. `content` is base64 the caller already holds,
 * which crosses no network and is bounded by the request body instead.
 *
 * Both come out as the same thing: bytes, whatever the source declared about
 * them, and a label honest enough to put in front of the extracted text.
 */

import { fetchPublicUrl } from "./publicFetch.js";
import { FETCH_TIMEOUT_MS, MAX_DOCUMENT_BYTES } from "./limits.js";
import { SERVER_NAME, SERVER_VERSION } from "./version.js";
import { DocumentError } from "./errors.js";

/**
 * Node's fetch sends `user-agent: node`, which a good share of the edge blocks
 * or challenges outright — and that shows up as "the server answered 403" with
 * the cause on this side. Naming the build and linking the source is also what
 * lets a site owner who sees this in their logs find out what it is.
 */
const USER_AGENT = `${SERVER_NAME}/${SERVER_VERSION} (+https://github.com/opspresso/mcp-document)`;

/**
 * Sent on every fetch: accept anything, rather than a list of the types this
 * reads. A server that honours Accept and holds a `.hwp` has nothing in any
 * such list to offer and would answer 406 — which carries no content-type, so
 * the failure would arrive with less information than the document it refused.
 */
const ACCEPT = "*/*";

export class SourceError extends DocumentError {}

export interface DocumentSource {
  bytes: Uint8Array;
  /** Declared media type, lower-cased. Empty when nothing declared one. */
  mimeType: string;
  /** Charset the transport declared, when it declared one. */
  charset?: string;
  /** What the provenance header names as the origin of this text. */
  label: string;
  /** Filename, from the caller or from the URL path, for extension hints. */
  filename?: string;
}

/** `text/plain; charset=EUC-KR` → `{ mimeType: "text/plain", charset: "euc-kr" }`. */
export function parseContentType(header: string | null): { mimeType: string; charset?: string } {
  const [type = "", ...parameters] = (header ?? "").split(";");
  const charset = parameters
    .map((parameter) => /^\s*charset\s*=\s*"?([^";]+)"?\s*$/i.exec(parameter)?.[1])
    .find((value): value is string => value !== undefined);
  return {
    mimeType: type.trim().toLowerCase(),
    ...(charset ? { charset: charset.trim().toLowerCase() } : {}),
  };
}

/**
 * The last path segment of a URL, when it looks like a filename.
 *
 * Only a hint for `detect.ts`, and never a path: what it is used for is the
 * extension. A segment with no dot in it is not offered at all rather than
 * offered as an extensionless name, which would make the detector confident
 * about something it was told nothing about.
 */
export function filenameFromUrl(url: string): string | undefined {
  try {
    const last = new URL(url).pathname.split("/").filter(Boolean).at(-1);
    if (!last) {
      return undefined;
    }
    const name = decodeURIComponent(last);
    return name.includes(".") ? name : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Read a body with a hard ceiling.
 *
 * The declared length is checked first and separately: a `content-length` that
 * lies must not decide how much is pulled into memory, and a body with no
 * declared length at all must still be bounded. So the stream is cut the moment
 * it goes over, rather than buffered and measured afterwards.
 */
async function readCapped(response: Response, maxBytes: number): Promise<Uint8Array> {
  const declared = Number(response.headers.get("content-length") ?? "");
  if (Number.isFinite(declared) && declared > maxBytes) {
    throw new SourceError(
      `the document is ${declared.toLocaleString("en-US")} bytes, over the ` +
        `${maxBytes.toLocaleString("en-US")} limit`,
    );
  }
  const body = response.body;
  if (!body) {
    return new Uint8Array(0);
  }
  const chunks: Uint8Array[] = [];
  let size = 0;
  const reader = body.getReader();
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) {
        break;
      }
      size += value.byteLength;
      if (size > maxBytes) {
        throw new SourceError(
          `the document is larger than the ${maxBytes.toLocaleString("en-US")} byte limit`,
        );
      }
      chunks.push(value);
    }
  } finally {
    // Releasing the lock lets the connection be reused; cancelling an
    // already-finished body is a no-op.
    reader.releaseLock();
    await body.cancel().catch(() => {});
  }
  return Buffer.concat(chunks);
}

async function fromUrl(url: string, filename: string | undefined): Promise<DocumentSource> {
  const response = await fetchPublicUrl(url, {
    signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    headers: { Accept: ACCEPT, "User-Agent": USER_AGENT },
  });
  if (!response.ok) {
    await response.body?.cancel().catch(() => {});
    throw new SourceError(`the server answered ${response.status}`);
  }
  const { mimeType, charset } = parseContentType(response.headers.get("content-type"));
  const bytes = await readCapped(response, MAX_DOCUMENT_BYTES);
  const named = filename ?? filenameFromUrl(url);
  return {
    bytes,
    mimeType,
    ...(charset ? { charset } : {}),
    label: url,
    ...(named ? { filename: named } : {}),
  };
}

/**
 * Base64 with nothing guessed.
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

/**
 * Resolve exactly one of `url` and `content` into bytes.
 *
 * Both together is an error rather than a precedence rule: a caller that sent
 * both has two ideas about which document this is, and picking one silently
 * reads a file it was not asked to read.
 */
export async function loadSource(args: {
  url?: string | undefined;
  content?: string | undefined;
  filename?: string | undefined;
}): Promise<DocumentSource> {
  const { url, content, filename } = args;
  if (url && content) {
    throw new SourceError("pass either `url` or `content`, not both");
  }
  if (url) {
    return fromUrl(url, filename);
  }
  if (content) {
    const bytes = decodeBase64(content);
    return {
      bytes,
      mimeType: "",
      label: filename ?? "an uploaded document",
      ...(filename ? { filename } : {}),
    };
  }
  throw new SourceError("either `url` or `content` is required");
}
