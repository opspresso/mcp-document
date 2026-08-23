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
 * A qualified name without its prefix. `table:table-cell` → `table-cell`.
 *
 * Four readers had a copy of this and `docx.ts` deliberately does not use it:
 * `word/document.xml` can carry DrawingML inside `mc:AlternateContent`, where
 * `a:t` is a shape's text, so matching the local name `t` there would leak
 * WordArt and fallback graphics into the body. Everywhere else the prefix is
 * conventional rather than required, and keying on it is what breaks when a
 * writer binds the namespace to a different one.
 */
export function localName(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

/**
 * Compiled once per attribute name.
 *
 * `walkXml` hands attributes over as one raw string, so every lookup is a
 * match. That was two per document while nothing but a heading style depended
 * on one; a reader that resolves spans, levels, list ids and link targets asks
 * several of every element it passes. The names are the readers' own literals,
 * never anything a document chose, so the map is bounded by the source code.
 */
const ATTRIBUTE_PATTERNS = new Map<string, RegExp>();

function attributePattern(name: string): RegExp {
  const known = ATTRIBUTE_PATTERNS.get(name);
  if (known) {
    return known;
  }
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`(?:^|\\s)${escaped}\\s*=\\s*(["'])([^]*?)\\1`);
  ATTRIBUTE_PATTERNS.set(name, pattern);
  return pattern;
}

/**
 * One attribute's value, entity-decoded.
 *
 * Single *or* double quoted, because XML allows both — `endOfTag` above already
 * tracks both, and only the extractors did not. A value the pattern missed
 * reads as "no heading style", "no colspan", "no link": the failure that looks
 * like success, which is why three readers carrying a double-quote-only copy
 * of this was worth ending.
 *
 * Decoded, because the values structure depends on are author prose
 * (`wp:docPr/@descr`), sheet names and hyperlink targets, and all three carry
 * `&amp;` as a matter of course. `<sheet name="A&amp;B">` used to come back as
 * the heading `## A&amp;B`.
 */
export function attributeOf(attributes: string, name: string): string | undefined {
  const match = attributePattern(name).exec(attributes);
  return match?.[2] === undefined ? undefined : decodeXmlEntities(match[2]);
}

/**
 * Every attribute at once, for an element that is asked about three or more.
 *
 * A duplicated name keeps its first value, which is what a parser that refused
 * the document would have called an error and what a reader that has to answer
 * calls the one the writer meant.
 */
export function attributesOf(attributes: string): Map<string, string> {
  const found = new Map<string, string>();
  const pattern = /([^\s=/>]+)\s*=\s*(["'])([^]*?)\2/g;
  for (const match of attributes.matchAll(pattern)) {
    const [, name, , value] = match;
    if (name !== undefined && value !== undefined && !found.has(name)) {
      found.set(name, decodeXmlEntities(value));
    }
  }
  return found;
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
