/**
 * HWPX to text.
 *
 * HWPX is OWPML in an ODF-style zip: the body is `Contents/section0.xml`,
 * `section1.xml` and so on, and they are read in numeric order because that is
 * the order the document is in — a section list sorted as strings puts
 * `section10` between `section1` and `section2`, which silently reorders any
 * document long enough to have ten of them.
 *
 * Elements are matched on their **local name**, without the `hp:` prefix.
 * The prefix is conventional rather than required — a document is free to bind
 * the namespace to another one — and a reader that keyed on `hp:t` would return
 * "no text" for a valid file rather than an error anyone could act on.
 */

import { walkXml, type XmlHandler } from "../xml.js";
import { openZip, type ZipEntry } from "../zip.js";
import { normalize } from "./lines.js";
import { DocumentError } from "../errors.js";

const SECTION = /^Contents\/section(\d+)\.xml$/;

export class HwpxError extends DocumentError {}

export interface HwpxText {
  text: string;
  /** How many section parts contributed, which is what "the whole body" means here. */
  sections: number;
}

function localName(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

/** Section parts in document order, which is numeric and not lexical. */
export function sectionsOf(entries: readonly ZipEntry[]): string[] {
  return entries
    .map((entry) => ({ name: entry.name, index: Number(SECTION.exec(entry.name)?.[1] ?? NaN) }))
    .filter((entry) => Number.isInteger(entry.index))
    .sort((a, b) => a.index - b.index)
    .map((entry) => entry.name);
}

class Extractor implements XmlHandler {
  private readonly lines: string[] = [];
  private buffer = "";
  private textDepth = 0;
  private cellDepth = 0;

  text(value: string): void {
    if (this.textDepth > 0) {
      this.buffer += value;
    }
  }

  open(name: string, _attributes: string, selfClosing: boolean): void {
    switch (localName(name)) {
      case "t":
        if (!selfClosing) {
          this.textDepth += 1;
        }
        return;
      case "tab":
        this.buffer += "\t";
        return;
      case "lineBreak":
        this.buffer += "\n";
        return;
      case "tc":
        this.cellDepth += 1;
        return;
      default:
        return;
    }
  }

  close(name: string): void {
    switch (localName(name)) {
      case "t":
        this.textDepth = Math.max(0, this.textDepth - 1);
        return;
      case "p":
        // Inside a cell a paragraph is a line wrap, not a row break.
        if (this.cellDepth > 0) {
          this.buffer += " ";
          return;
        }
        this.endLine();
        return;
      case "tc":
        this.cellDepth = Math.max(0, this.cellDepth - 1);
        this.buffer += " | ";
        return;
      case "tr":
        this.endLine();
        return;
      default:
        return;
    }
  }

  private endLine(): void {
    const line = this.buffer;
    this.buffer = "";
    for (const part of line.split("\n")) {
      this.lines.push(part);
    }
  }

  finish(): string {
    this.endLine();
    return normalize(this.lines);
  }
}

/** One section's XML, separated so the walk above can be tested without a zip. */
export function sectionXmlToText(xml: string): string {
  const extractor = new Extractor();
  walkXml(xml, extractor);
  return extractor.finish();
}

export function hwpxToText(bytes: Uint8Array): HwpxText {
  const archive = openZip(bytes);
  const names = sectionsOf(archive.entries);
  if (names.length === 0) {
    throw new HwpxError(
      "this .hwpx has no Contents/section*.xml part, so it has no body — it is not a document 한글 wrote",
    );
  }
  const parts = archive.read(names);
  const decoder = new TextDecoder("utf-8");
  const text = normalize(
    names.flatMap((name) => {
      const part = parts.get(name);
      return part ? sectionXmlToText(decoder.decode(part)).split("\n") : [];
    }),
  );
  if (text === "") {
    throw new HwpxError(
      `this .hwpx has ${names.length} section(s) but no text in any of them — its content is ` +
        "most likely images, which need OCR rather than text extraction",
    );
  }
  return { text, sections: names.length };
}
