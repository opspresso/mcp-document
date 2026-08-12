/**
 * OpenDocument — text, spreadsheet and presentation — to the text a model
 * should read.
 *
 * One reader for all three, which is not a shortcut: ODF puts the body of every
 * kind in a single `content.xml` and marks structure with the same handful of
 * elements. A paragraph is `text:p` whether it sits in a document, a cell or a
 * slide; a cell is `table:table-cell` in a spreadsheet and in a text document's
 * table alike. Three readers would be the same reader three times, and the
 * places they would drift apart are exactly the shared elements.
 *
 * What differs by kind is only what a heading means — a sheet name, a slide
 * number, nothing — so that is the only thing this branches on.
 */

import { walkXml, type XmlHandler } from "../xml.js";
import { openZip } from "../zip.js";
import { normalize } from "./lines.js";
import { DocumentError } from "../errors.js";

export class OdfError extends DocumentError {}

const CONTENT = "content.xml";
const MIMETYPE = "mimetype";

export type OdfKind = "text" | "spreadsheet" | "presentation";

export interface OdfText {
  text: string;
  kind: OdfKind;
  /** Sheets or slides that contributed; absent for a text document. */
  parts?: number;
}

function localName(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

function attribute(attributes: string, name: string): string | undefined {
  const match = new RegExp(`(?:^|\\s)${name}\\s*=\\s*"([^"]*)"`).exec(attributes);
  return match?.[1];
}

/** What the package says it is. The `mimetype` entry is required to be first. */
export function odfKindOf(mimetype: string): OdfKind | undefined {
  if (mimetype.includes("opendocument.text")) {
    return "text";
  }
  if (mimetype.includes("opendocument.spreadsheet")) {
    return "spreadsheet";
  }
  if (mimetype.includes("opendocument.presentation")) {
    return "presentation";
  }
  return undefined;
}

class Extractor implements XmlHandler {
  private readonly lines: string[] = [];
  private buffer = "";
  private cellDepth = 0;
  /** Sheets or slides seen, for the note. */
  parts = 0;

  constructor(private readonly kind: OdfKind) {}

  text(value: string): void {
    this.buffer += value;
  }

  open(name: string, attributes: string, selfClosing: boolean): void {
    switch (localName(name)) {
      case "tab":
        this.buffer += "\t";
        return;
      case "line-break":
        this.buffer += "\n";
        return;
      // `<text:s text:c="4"/>` is a run of spaces the format encodes rather
      // than storing, because XML would collapse them.
      case "s": {
        const count = Number(attribute(attributes, "text:c") ?? "1");
        this.buffer += " ".repeat(Number.isInteger(count) && count > 0 ? Math.min(count, 80) : 1);
        return;
      }
      case "table": {
        if (this.kind !== "spreadsheet") {
          return;
        }
        this.parts += 1;
        const sheet = attribute(attributes, "table:name");
        this.flush();
        this.lines.push(sheet ? `## ${sheet}` : `## Sheet ${this.parts}`);
        return;
      }
      case "page":
        if (this.kind !== "presentation") {
          return;
        }
        this.parts += 1;
        this.flush();
        this.lines.push(`## Slide ${this.parts}`);
        return;
      case "table-cell":
        if (!selfClosing) {
          this.cellDepth += 1;
        }
        return;
      default:
        return;
    }
  }

  close(name: string): void {
    switch (localName(name)) {
      case "p":
      case "h":
        // Inside a cell a paragraph is a line *within* the cell, not the end of
        // the row — flushing here would put every cell on its own line.
        if (this.cellDepth === 0) {
          this.flush();
        }
        return;
      case "table-cell":
        this.buffer += " | ";
        if (this.cellDepth > 0) {
          this.cellDepth -= 1;
        }
        return;
      case "table-row":
        this.flush();
        return;
      default:
        return;
    }
  }

  private flush(): void {
    this.lines.push(this.buffer);
    this.buffer = "";
  }

  done(): string[] {
    this.flush();
    return this.lines;
  }
}

export function odfToText(bytes: Uint8Array): OdfText {
  const { entries, read } = openZip(bytes);
  const names = entries.map((entry) => entry.name);
  if (!names.includes(CONTENT)) {
    throw new OdfError("it has no content part — the archive is not an OpenDocument file");
  }
  const parts = read([CONTENT, MIMETYPE]);
  const decoder = new TextDecoder();
  const declared = parts.get(MIMETYPE);
  const kind = declared ? odfKindOf(decoder.decode(declared).trim()) : undefined;
  if (!kind) {
    throw new OdfError("the package does not say which kind of OpenDocument file it is");
  }
  const content = parts.get(CONTENT);
  if (!content) {
    throw new OdfError("its content part could not be read");
  }
  const extractor = new Extractor(kind);
  walkXml(decoder.decode(content), extractor);
  const text = normalize(extractor.done());
  if (text === "") {
    throw new OdfError("it has no readable text");
  }
  return { text, kind, ...(extractor.parts > 0 ? { parts: extractor.parts } : {}) };
}
