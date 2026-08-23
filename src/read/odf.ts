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
import { MAX_REPEATED_COLUMNS } from "../limits.js";
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

/**
 * Subtrees whose text is not the document's text.
 *
 * Every other reader here gates on a text element — `w:t`, `hp:t`, `a:t` — and
 * this one accumulated every character in `content.xml`. What that returned was
 * not stray whitespace: `text:tracked-changes` holds whole *deleted* paragraphs
 * and sits at the top of the body, so an edited document came back with text
 * the author removed presented as current, above the text they kept. An
 * annotation contributed its author's name and the comment body mid-sentence,
 * a footnote's own `</text:p>` flushed and cut its host paragraph in half, and
 * an ODP's speaker notes arrived as slide content while `document.ts` promised
 * they would not.
 *
 * Matched on the local name, like everything else here: `text:tracked-changes`,
 * `office:annotation`, `office:annotation-end`, `text:note`,
 * `presentation:notes`.
 */
const SKIPPED = new Set(["tracked-changes", "annotation", "annotation-end", "note", "notes"]);

/** A cell's `table:number-columns-repeated`, bounded and never zero. */
function repeatOf(attributes: string): number {
  const count = Number(attribute(attributes, "table:number-columns-repeated") ?? "1");
  if (!Number.isInteger(count) || count < 1) {
    return 1;
  }
  return Math.min(count, MAX_REPEATED_COLUMNS);
}

class Extractor implements XmlHandler {
  private readonly lines: string[] = [];
  private buffer = "";
  private cellDepth = 0;
  /** Open `text:p` / `text:h` elements. Text outside one is not prose. */
  private paraDepth = 0;
  /** Open `SKIPPED` subtrees. Counters keep counting inside one; effects stop. */
  private skipDepth = 0;
  /** Where the outermost open cell's content starts in `buffer`. */
  private cellStart = 0;
  /** How many columns that cell stands for. */
  private cellRepeat = 1;
  /**
   * Columns a repeat run owes, held until a later cell in the row needs them.
   *
   * Materialising eagerly would draw the sheet's full width for every row, so
   * the run is paid for only when something has to sit to the right of it —
   * which is exactly when its width starts to matter.
   */
  private owed = 0;
  /** Sheets or slides seen, for the note. */
  parts = 0;

  constructor(private readonly kind: OdfKind) {}

  /** Prose is what a paragraph holds, outside a subtree we are skipping. */
  private get capturing(): boolean {
    return this.paraDepth > 0 && this.skipDepth === 0;
  }

  text(value: string): void {
    if (this.capturing) {
      this.buffer += value;
    }
  }

  open(name: string, attributes: string, selfClosing: boolean): void {
    const local = localName(name);
    if (SKIPPED.has(local)) {
      if (!selfClosing) {
        this.skipDepth += 1;
      }
      return;
    }
    switch (local) {
      case "p":
      case "h":
        if (!selfClosing) {
          this.paraDepth += 1;
        }
        return;
      case "tab":
        if (this.capturing) {
          this.buffer += "\t";
        }
        return;
      case "line-break":
        if (this.capturing) {
          this.buffer += "\n";
        }
        return;
      // `<text:s text:c="4"/>` is a run of spaces the format encodes rather
      // than storing, because XML would collapse them.
      case "s": {
        if (!this.capturing) {
          return;
        }
        const count = Number(attribute(attributes, "text:c") ?? "1");
        this.buffer += " ".repeat(Number.isInteger(count) && count > 0 ? Math.min(count, 80) : 1);
        return;
      }
      case "table": {
        if (this.kind !== "spreadsheet" || this.skipDepth > 0) {
          return;
        }
        this.parts += 1;
        const sheet = attribute(attributes, "table:name");
        this.flush();
        this.lines.push(sheet ? `## ${sheet}` : `## Sheet ${this.parts}`);
        return;
      }
      case "page":
        if (this.kind !== "presentation" || this.skipDepth > 0) {
          return;
        }
        this.parts += 1;
        this.flush();
        this.lines.push(`## Slide ${this.parts}`);
        return;
      // A covered cell is one a merge swallowed. It holds no text and it holds
      // its column, so it separates like any other — without it, every value
      // to the right of a merge moves left by the width of the merge.
      case "table-cell":
      case "covered-table-cell": {
        if (this.skipDepth > 0) {
          return;
        }
        const repeat = repeatOf(attributes);
        if (selfClosing) {
          // An empty cell: nothing to place, only columns to owe.
          this.owed += repeat;
          return;
        }
        if (this.cellDepth === 0) {
          this.settle();
          this.cellStart = this.buffer.length;
          this.cellRepeat = repeat;
        }
        this.cellDepth += 1;
        return;
      }
      default:
        return;
    }
  }

  close(name: string): void {
    const local = localName(name);
    if (SKIPPED.has(local)) {
      if (this.skipDepth > 0) {
        this.skipDepth -= 1;
      }
      return;
    }
    switch (local) {
      case "p":
      case "h":
        if (this.paraDepth > 0) {
          this.paraDepth -= 1;
        }
        // Inside a cell a paragraph is a line *within* the cell, not the end of
        // the row — flushing here would put every cell on its own line. Inside
        // a skipped subtree it is not this document's paragraph at all.
        if (this.cellDepth === 0 && this.skipDepth === 0) {
          this.flush();
        }
        return;
      case "table-cell":
      case "covered-table-cell": {
        if (this.skipDepth > 0) {
          return;
        }
        if (this.cellDepth > 0) {
          this.cellDepth -= 1;
        }
        const content = this.cellDepth === 0 ? this.buffer.slice(this.cellStart) : "";
        this.buffer += " | ";
        // A repeated cell with content stands for that value in each of its
        // columns; the separator is what carries the empty ones.
        for (let repeat = 1; repeat < this.cellRepeat; repeat += 1) {
          this.buffer += `${content} | `;
        }
        this.cellRepeat = 1;
        return;
      }
      case "table-row":
        if (this.skipDepth > 0) {
          return;
        }
        // Columns owed at the end of a row are the padding every ODS row
        // carries. Nothing sits to their right, so nobody is waiting on them.
        this.owed = 0;
        this.flush();
        return;
      default:
        return;
    }
  }

  /** Pay for the repeat run standing between the last cell and this one. */
  private settle(): void {
    if (this.owed > 0) {
      this.buffer += " | ".repeat(this.owed);
      this.owed = 0;
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
