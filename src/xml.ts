import { DocumentError } from "./errors.js";
import { MAX_XML_DEPTH, MAX_XML_EVENTS } from "./limits.js";

/**
 * A tag walker, which is all that reading DOCX and HWPX needs.
 *
 * Not a DOM parser, for the same reason `mcp-url-fetch` does not parse HTML: a
 * real one is a dependency with its own attack surface, and the job here is
 * narrow. Both formats put their prose inside one element (`w:t`, `hp:t`) and
 * mark their structure with others, so a scan that reports opens, closes and
 * the characters between them is the whole interface — a handler decides which
 * of those mean something.
 *
 * It handles what these two formats actually contain: declarations, comments,
 * CDATA, self-closing tags, attributes with `>` inside quotes. It does not
 * validate: an unbalanced document is read as far as it goes rather than
 * refused, because half a document's text is worth more than an error about
 * markup nobody will look at.
 */

export interface XmlHandler {
  /** Character data between tags, already entity-decoded. */
  text(value: string): void;
  /** `name` excludes the brackets and the slash; `selfClosing` gets no `close`. */
  open(name: string, attributes: string, selfClosing: boolean): void;
  close(name: string): void;
}

export class XmlError extends DocumentError {}

const NAMED_ENTITIES: Record<string, string> = {
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  nbsp: " ",
};

export function decodeXmlEntities(value: string): string {
  if (!value.includes("&")) {
    return value;
  }
  return (
    value
      .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => codePoint(parseInt(hex, 16)))
      .replace(/&#(\d+);/g, (_, dec: string) => codePoint(Number(dec)))
      .replace(/&([a-z]+);/gi, (match, name: string) => NAMED_ENTITIES[name.toLowerCase()] ?? match)
      // Last: an escaped ampersand may itself introduce an entity that was never
      // meant to be decoded.
      .replace(/&amp;/gi, "&")
  );
}

/** Out of range or a surrogate is dropped rather than becoming U+FFFD, which reads as corruption. */
function codePoint(value: number): string {
  if (!Number.isFinite(value) || value < 0 || value > 0x10ffff) {
    return "";
  }
  if (value >= 0xd800 && value <= 0xdfff) {
    return "";
  }
  return String.fromCodePoint(value);
}

/** Escape for use in an XML text node or a double-quoted attribute value. */
export function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * The end of a tag that starts at `from`.
 *
 * Scanned rather than matched with `[^>]*>` because an attribute value may
 * contain `>` — rare in these formats, but a mis-cut tag turns the rest of a
 * document's markup into prose, which is the failure that looks like success.
 */
function endOfTag(xml: string, from: number): number {
  let quote: string | undefined;
  for (let index = from + 1; index < xml.length; index += 1) {
    const character = xml[index];
    if (quote) {
      if (character === quote) {
        quote = undefined;
      }
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }
    if (character === ">") {
      return index;
    }
  }
  return -1;
}

export function walkXml(xml: string, handler: XmlHandler): void {
  let index = 0;
  let depth = 0;
  let events = 0;
  while (index < xml.length) {
    const start = xml.indexOf("<", index);
    if (start === -1) {
      emit(handler, xml.slice(index));
      return;
    }
    emit(handler, xml.slice(index, start));

    if (xml.startsWith("<![CDATA[", start)) {
      const end = xml.indexOf("]]>", start);
      // Unterminated CDATA takes the rest of the document with it, which is what
      // it says: everything after it is literal.
      handler.text(end === -1 ? xml.slice(start + 9) : xml.slice(start + 9, end));
      index = end === -1 ? xml.length : end + 3;
      continue;
    }
    if (xml.startsWith("<!--", start)) {
      const end = xml.indexOf("-->", start);
      index = end === -1 ? xml.length : end + 3;
      continue;
    }
    if (xml.startsWith("<!DOCTYPE", start) || xml.startsWith("<!ENTITY", start)) {
      throw new XmlError("document XML may not declare a DTD or entity");
    }
    // `<?xml ?>`: markup that names nothing this reads.
    if (xml.startsWith("<?", start) || xml.startsWith("<!", start)) {
      const end = endOfTag(xml, start);
      index = end === -1 ? xml.length : end + 1;
      continue;
    }

    const end = endOfTag(xml, start);
    if (end === -1) {
      // A tag cut off by a truncated document is not prose; dropping it is what
      // a parser would do with an unterminated element.
      return;
    }
    const inner = xml.slice(start + 1, end);
    index = end + 1;
    if (inner.startsWith("/")) {
      depth = Math.max(0, depth - 1);
      events += 1;
      if (events > MAX_XML_EVENTS) {
        throw new XmlError(`document XML has more than ${MAX_XML_EVENTS.toLocaleString("en-US")} elements`);
      }
      handler.close(inner.slice(1).trim());
      continue;
    }
    const selfClosing = inner.endsWith("/");
    const body = selfClosing ? inner.slice(0, -1) : inner;
    const space = /\s/.exec(body)?.index ?? body.length;
    events += 1;
    if (events > MAX_XML_EVENTS) {
      throw new XmlError(`document XML has more than ${MAX_XML_EVENTS.toLocaleString("en-US")} elements`);
    }
    if (!selfClosing) {
      depth += 1;
      if (depth > MAX_XML_DEPTH) {
        throw new XmlError(`document XML is nested more than ${MAX_XML_DEPTH} elements deep`);
      }
    }
    handler.open(body.slice(0, space), body.slice(space).trim(), selfClosing);
  }
}

function emit(handler: XmlHandler, raw: string): void {
  if (raw !== "") {
    handler.text(decodeXmlEntities(raw));
  }
}
