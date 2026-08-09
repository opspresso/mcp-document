/**
 * Bytes that are already the text, once somebody decides what encoding they are
 * in.
 *
 * The order of evidence is BOM, then the transport's charset, then a
 * declaration inside the document, then UTF-8. A BOM outranks the header
 * because it is a statement the file makes about itself and cannot be stale;
 * everything after it is somebody's configuration.
 *
 * An unknown charset label falls back to UTF-8 rather than failing. Mojibake is
 * something a reader can work around and often still read through; a hard error
 * gives them nothing.
 */

export interface DecodedText {
  text: string;
  /** The encoding actually used, so a caller can say what it assumed. */
  charset: string;
}

/** A byte-order mark, and what it settles. */
function fromBom(bytes: Uint8Array): { charset: string; offset: number } | undefined {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return { charset: "utf-8", offset: 3 };
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return { charset: "utf-16le", offset: 2 };
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return { charset: "utf-16be", offset: 2 };
  }
  return undefined;
}

/**
 * The charset an XML declaration states.
 *
 * Only the head is scanned, and as latin1, because the declaration is ASCII in
 * every encoding it could be announcing. Still common on Korean documents,
 * where guessing UTF-8 turns the whole file into replacement characters.
 */
export function charsetFromDeclaration(bytes: Uint8Array): string | undefined {
  const head = Buffer.from(bytes.subarray(0, 1024)).toString("latin1");
  return /<\?xml[^>]+encoding\s*=\s*["']([a-z0-9_\-]+)/i.exec(head)?.[1]?.toLowerCase();
}

export function decodeText(bytes: Uint8Array, declared?: string): DecodedText {
  const bom = fromBom(bytes);
  const body = bom ? bytes.subarray(bom.offset) : bytes;
  const charset = bom?.charset ?? declared ?? charsetFromDeclaration(bytes) ?? "utf-8";
  if (charset !== "utf-8" && charset !== "utf8") {
    try {
      return { text: new TextDecoder(charset).decode(body), charset };
    } catch {
      // RangeError: this build has no such encoding. Fall through to UTF-8.
    }
  }
  return { text: new TextDecoder("utf-8").decode(body), charset: "utf-8" };
}

/**
 * Trailing whitespace on every line, and the blank lines at either end, are
 * removed; blank lines *inside* the document are kept, because in a plain text
 * file they are the only structure there is.
 */
export function plainToText(bytes: Uint8Array, declared?: string): DecodedText {
  const { text, charset } = decodeText(bytes, declared);
  return {
    text: text
      .replace(/\r\n?/g, "\n")
      .split("\n")
      .map((line) => line.replace(/[^\S\n]+$/, ""))
      .join("\n")
      .replace(/^\n+|\n+$/g, ""),
    charset,
  };
}
