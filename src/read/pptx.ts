/**
 * PPTX to the text a model should read.
 *
 * The cheapest reader here, because a deck is the same shape as an HWPX body:
 * one XML part per slide, numbered, and the prose sits in one element. What
 * differs is what a slide *is* — a page of boxes with no reading order the file
 * commits to. Shapes come out in the order the file stores them, which is the
 * order they were added, not the order a person's eye takes them.
 *
 * So the text is honest and the layout is not recoverable, and the note says
 * how many slides answered rather than pretending the deck has been understood.
 * Speaker notes are a separate part and are left out: they are the presenter's
 * script, not the slide, and interleaving them would make the deck say things
 * the audience never saw.
 */

import { walkXml, type XmlHandler } from "../xml.js";
import { openZip, type ZipEntry } from "../zip.js";
import { normalize } from "./lines.js";
import { DocumentError } from "../errors.js";

export class PptxError extends DocumentError {}

const SLIDE = /^ppt\/slides\/slide(\d+)\.xml$/;

export interface PptxText {
  text: string;
  /** How many slides contributed, which is what "the whole deck" means here. */
  slides: number;
}

function localName(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

/** Slide parts in deck order, which is numeric and not lexical. */
export function slidesOf(entries: readonly ZipEntry[]): string[] {
  return entries
    .map((entry) => ({ name: entry.name, index: Number(SLIDE.exec(entry.name)?.[1] ?? NaN) }))
    .filter((entry) => Number.isInteger(entry.index))
    .sort((a, b) => a.index - b.index)
    .map((entry) => entry.name);
}

class Extractor implements XmlHandler {
  private readonly lines: string[] = [];
  private buffer = "";
  private textDepth = 0;
  private cellDepth = 0;
  private fieldDepth = 0;

  text(value: string): void {
    // Text inside `a:fld` is a render cache, not content: a slide-number field
    // carries the digit PowerPoint last computed, and reading it out would put
    // a stray "7" on the end of every slide.
    if (this.textDepth > 0 && this.fieldDepth === 0) {
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
      // A soft break inside a run, which is a line the author put there.
      case "br":
        this.buffer += "\n";
        return;
      case "fld":
        if (!selfClosing) {
          this.fieldDepth += 1;
        }
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
        if (this.textDepth > 0) {
          this.textDepth -= 1;
        }
        return;
      case "fld":
        if (this.fieldDepth > 0) {
          this.fieldDepth -= 1;
        }
        return;
      case "tc":
        // Cells are separated the way every other reader here separates them,
        // so a table reads the same whichever format it arrived in.
        this.buffer += " | ";
        if (this.cellDepth > 0) {
          this.cellDepth -= 1;
        }
        return;
      // A paragraph ends a line — except inside a cell, where it is a line
      // *within* the cell and flushing would put every cell on its own row.
      case "p":
        if (this.cellDepth === 0) {
          this.flush();
        }
        return;
      // A row ends a line, so cells do not run on into the next one.
      case "tr":
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

export function pptxToText(bytes: Uint8Array): PptxText {
  const { entries, read } = openZip(bytes);
  const names = slidesOf(entries);
  if (names.length === 0) {
    // A zip with no slides is not a deck. Saying so beats an empty success,
    // which would read as "this presentation is blank".
    throw new PptxError("it has no slides — the archive is not a PPTX presentation");
  }
  const parts = read(names);
  const lines: string[] = [];
  let slides = 0;
  for (const [index, name] of names.entries()) {
    const part = parts.get(name);
    if (!part) {
      continue;
    }
    const extractor = new Extractor();
    walkXml(new TextDecoder().decode(part), extractor);
    const slideText = normalize(extractor.done());
    slides += 1;
    // Numbered, because a slide is how a person refers to a place in a deck —
    // "the chart on slide 7" is the only address this format has.
    lines.push(`## Slide ${index + 1}`, ...(slideText === "" ? [] : [slideText]), "");
  }
  return { text: normalize(lines), slides };
}
