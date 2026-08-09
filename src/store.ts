/**
 * Where a written document goes, and how a person gets to it.
 *
 * A link, because a link is the only thing that survives the trip. Agent
 * Studio's MCP layer drops a non-text blob outright — it decodes what it is
 * given as UTF-8, finds that a `.docx` is not text, and returns
 * `[binary resource omitted]` (`infrastructure/mcp/toolManager.ts`). So the
 * bytes have to land somewhere addressable and the tool result has to be the
 * address. That is not a workaround; it is the only shape that reaches a user.
 *
 * The key is composed here and never by the caller: `<prefix>/<tenant>/<ulid>/
 * <name>`. The tenant comes from a header an operator set (`tenant.ts`), the
 * ULID makes the key unique and time-ordered, and the name is the model's —
 * sanitised, because it is the one part of the key chosen by something that
 * read a document a moment ago.
 */

import { GetObjectCommand, PutObjectCommand, S3Client } from "@aws-sdk/client-s3";
import { getSignedUrl } from "@aws-sdk/s3-request-presigner";
import type { Config } from "./config.js";
import { DocumentError } from "./errors.js";
import { ulid } from "./id.js";

export class StoreError extends DocumentError {}

/** Long enough to stay recognisable, short enough to keep a key manageable. */
const MAX_NAME_CHARS = 80;

export interface StoredDocument {
  key: string;
  /** A presigned GET, valid for the configured lifetime. */
  url: string;
  bytes: number;
  expiresAt: Date;
}

/**
 * A filename that is safe to be part of a key and readable when downloaded.
 *
 * Korean is kept: this server exists partly to write 한글 documents, and
 * transliterating a title into ASCII would produce a file nobody can find
 * again. What goes is anything a path or a header reads as structure —
 * separators, control characters, quotes — and any leading dot, so the result
 * cannot be a name that hides.
 */
export function safeFilename(name: string, extension: string): string {
  const stem = name
    .normalize("NFC")
    .replace(new RegExp(`\\.${extension}$`, "i"), "")
    // Control characters, which a filename must not carry into a header.
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/[\\/:*?"'<>|]/g, " ")
    // A word made only of dots was a path segment before the separators became
    // spaces, and it names nothing. Dropping the whole word is what keeps
    // `../../etc/passwd` from arriving as `.. .. etc passwd`.
    .split(/\s+/)
    .filter((word) => word !== "" && !/^\.+$/.test(word))
    .join(" ")
    .replace(/^\.+/, "")
    .replace(/\.+$/, "")
    .slice(0, MAX_NAME_CHARS)
    .trim();
  return `${stem === "" ? "document" : stem}.${extension}`;
}

/**
 * One client for the process. Created on first use rather than at import so
 * that a test — or a read-only call — never has to have credentials.
 */
let client: S3Client | undefined;

function clientFor(config: Config): S3Client {
  client ??= new S3Client({ region: config.region });
  return client;
}

export function documentKey(config: Config, tenant: string, filename: string): string {
  return [config.prefix, tenant, ulid(), filename].filter((part) => part !== "").join("/");
}

export async function storeDocument(
  config: Config,
  tenant: string,
  filename: string,
  bytes: Uint8Array,
  contentType: string,
): Promise<StoredDocument> {
  const key = documentKey(config, tenant, filename);
  try {
    const command = new PutObjectCommand({
      Bucket: config.bucket,
      Key: key,
      Body: bytes,
      ContentType: contentType,
      // RFC 5987, so a Korean filename arrives intact rather than as the key's
      // last segment percent-decoded by whatever the browser felt like.
      ContentDisposition: `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
    });
    await clientFor(config).send(command);
    // A GET, obviously — but worth stating, because presigning the PutObject
    // that was just sent also produces a working-looking URL, and what it
    // authorises is *overwriting* the document rather than reading it.
    const url = await getSignedUrl(
      clientFor(config),
      new GetObjectCommand({ Bucket: config.bucket, Key: key }),
      { expiresIn: config.downloadTtlSeconds },
    );
    return {
      key,
      url,
      bytes: bytes.byteLength,
      expiresAt: new Date(Date.now() + config.downloadTtlSeconds * 1000),
    };
  } catch (error) {
    // The reason is kept: "Access Denied" and "NoSuchBucket" are the two that
    // actually happen, and both are an operator's to fix — a result that only
    // said "storage failed" would send them to read logs for a fact the tool
    // already had.
    throw new StoreError(
      `the document could not be stored — ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}
