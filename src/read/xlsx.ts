/**
 * XLSX to the text a model should read.
 *
 * The reader with the most in it, because a spreadsheet is the format where
 * "the text" is least obviously defined. Three decisions carry it:
 *
 * **Values, never formulas.** A cell holds both — `<f>SUM(B2:B9)</f>` and the
 * `<v>` Excel last computed. The formula is how the number was made; the number
 * is the answer, and a model asked what the total is has no use for the recipe.
 *
 * **Position is reconstructed, not assumed.** A cell states its own address
 * (`r="C7"`) and empty cells are simply absent, so the parts arrive as a sparse
 * list. Emitting them in file order would silently shift every value left into
 * a column it does not belong to — a table that still looks like a table and
 * says something different. Columns are rebuilt from the addresses.
 *
 * **The budget is spent in rows.** A sheet can be far larger than any text
 * budget, so this cuts on a row boundary and reports what came back in the
 * document's own units — the same contract the other readers keep.
 */

import { walkXml, type XmlHandler } from "../xml.js";
import { openZip } from "../zip.js";
import { DocumentError } from "../errors.js";

export class XlsxError extends DocumentError {}

const WORKBOOK = "xl/workbook.xml";
const WORKBOOK_RELS = "xl/_rels/workbook.xml.rels";
const SHARED_STRINGS = "xl/sharedStrings.xml";

export interface XlsxText {
  text: string;
  /** Sheets that contributed, and how many the workbook holds. */
  sheets: number;
  totalSheets: number;
  /** Rows kept and rows the workbook held, across every sheet read. */
  rows: number;
  totalRows: number;
}

function localName(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

function attribute(attributes: string, name: string): string | undefined {
  const match = new RegExp(`(?:^|\\s)${name}\\s*=\\s*"([^"]*)"`).exec(attributes);
  return match?.[1];
}

/**
 * The column an address names, zero-based. `A` → 0, `Z` → 25, `AA` → 26.
 *
 * Base-26 with no zero digit, so it is not quite what a naive parse gives: the
 * letters are 1-based and the result is shifted back at the end.
 */
export function columnOf(reference: string): number {
  let column = 0;
  for (const character of reference) {
    const value = character.toUpperCase().charCodeAt(0) - 64;
    if (value < 1 || value > 26) {
      break;
    }
    column = column * 26 + value;
  }
  return column - 1;
}

/** `<si>` entries in order — cells with `t="s"` index into this. */
class SharedStrings implements XmlHandler {
  private readonly values: string[] = [];
  private buffer = "";
  private textDepth = 0;
  private inItem = false;

  text(value: string): void {
    if (this.textDepth > 0) {
      this.buffer += value;
    }
  }

  open(name: string, _attributes: string, selfClosing: boolean): void {
    const local = localName(name);
    if (local === "si") {
      this.inItem = true;
      this.buffer = "";
      return;
    }
    // A single `si` can hold several runs, each with its own `t`; they
    // concatenate into one string rather than becoming separate entries.
    if (local === "t" && this.inItem && !selfClosing) {
      this.textDepth += 1;
    }
  }

  close(name: string): void {
    const local = localName(name);
    if (local === "t" && this.textDepth > 0) {
      this.textDepth -= 1;
      return;
    }
    if (local === "si") {
      this.values.push(this.buffer);
      this.buffer = "";
      this.inItem = false;
    }
  }

  done(): string[] {
    return this.values;
  }
}

/** One worksheet's cells, as rows of already-positioned strings. */
class Sheet implements XmlHandler {
  private readonly rows: string[][] = [];
  private cells: string[] = [];
  private column = 0;
  private type = "";
  private buffer = "";
  private capturing = false;
  /** `<v>` inside `<f>` does not exist, but `<is><t>` does — both are values. */
  private inValue = false;

  constructor(private readonly shared: readonly string[]) {}

  text(value: string): void {
    if (this.capturing) {
      this.buffer += value;
    }
  }

  open(name: string, attributes: string, selfClosing: boolean): void {
    switch (localName(name)) {
      case "row":
        this.cells = [];
        return;
      case "c": {
        this.type = attribute(attributes, "t") ?? "";
        const reference = attribute(attributes, "r");
        // Absent addresses mean "the next column", which is what a writer that
        // omits them intends.
        this.column = reference ? columnOf(reference) : this.cells.length;
        this.buffer = "";
        this.inValue = false;
        if (selfClosing) {
          this.place("");
        }
        return;
      }
      case "v":
      case "t":
        if (!selfClosing) {
          this.capturing = true;
          this.inValue = true;
        }
        return;
      // Deliberately not captured: this is how the value was computed, not
      // what it is.
      case "f":
        this.capturing = false;
        return;
      default:
        return;
    }
  }

  close(name: string): void {
    switch (localName(name)) {
      case "v":
      case "t":
        this.capturing = false;
        return;
      case "c":
        this.place(this.resolve());
        return;
      case "row":
        this.rows.push(this.cells);
        this.cells = [];
        return;
      default:
        return;
    }
  }

  /** A shared-string index, or the literal the cell carried. */
  private resolve(): string {
    if (!this.inValue) {
      return "";
    }
    if (this.type === "s") {
      const index = Number(this.buffer);
      return Number.isInteger(index) ? (this.shared[index] ?? "") : "";
    }
    return this.buffer;
  }

  /** Pad to the cell's own column, so a sparse row keeps its shape. */
  private place(value: string): void {
    while (this.cells.length < this.column) {
      this.cells.push("");
    }
    this.cells[this.column] = value;
  }

  done(): string[][] {
    return this.rows;
  }
}

/** Sheet name → part path, in workbook order. */
function sheetParts(
  workbook: string | undefined,
  rels: string | undefined,
): Array<{ name: string; path: string }> {
  if (!workbook) {
    return [];
  }
  const targets = new Map<string, string>();
  if (rels) {
    for (const match of rels.matchAll(/<Relationship\b([^>]*)>/g)) {
      const id = attribute(match[1] ?? "", "Id");
      const target = attribute(match[1] ?? "", "Target");
      if (id && target) {
        // Targets are relative to `xl/`, and some writers make that explicit.
        targets.set(id, `xl/${target.replace(/^\/?(xl\/)?/, "")}`);
      }
    }
  }
  const sheets: Array<{ name: string; path: string }> = [];
  for (const match of workbook.matchAll(/<(?:\w+:)?sheet\b([^>]*)\/?>/g)) {
    const attributes = match[1] ?? "";
    const name = attribute(attributes, "name");
    if (!name) {
      continue;
    }
    const id = attribute(attributes, "r:id") ?? attribute(attributes, "id");
    // The relationship is authoritative; the conventional path is the fallback
    // for a workbook whose rels part is missing or unreadable.
    const path = (id && targets.get(id)) || `xl/worksheets/sheet${sheets.length + 1}.xml`;
    sheets.push({ name, path });
  }
  return sheets;
}

/**
 * Trailing empty cells are the grid's, not the row's.
 *
 * Trimmed *before* the separators are stripped: the join leaves a space after
 * the last one, and an anchored pattern will not reach past it.
 */
function rowText(cells: readonly string[]): string {
  return cells.join(" | ").trim().replace(/(?:\s*\|)+$/, "").trim();
}

export function xlsxToText(bytes: Uint8Array, maxChars: number): XlsxText {
  const { entries, read } = openZip(bytes);
  const names = entries.map((entry) => entry.name);
  if (!names.includes(WORKBOOK)) {
    throw new XlsxError("it has no workbook part — the archive is not an XLSX workbook");
  }

  const decoder = new TextDecoder();
  const head = read([WORKBOOK, WORKBOOK_RELS, SHARED_STRINGS]);
  const workbookXml = head.get(WORKBOOK);
  const sheets = sheetParts(
    workbookXml ? decoder.decode(workbookXml) : undefined,
    head.get(WORKBOOK_RELS) ? decoder.decode(head.get(WORKBOOK_RELS)!) : undefined,
  );
  if (sheets.length === 0) {
    throw new XlsxError("the workbook declares no sheets");
  }

  const sharedXml = head.get(SHARED_STRINGS);
  const shared = new SharedStrings();
  if (sharedXml) {
    walkXml(decoder.decode(sharedXml), shared);
  }
  const strings = shared.done();

  const wanted = sheets.map((sheet) => sheet.path).filter((path) => names.includes(path));
  const parts = read(wanted);

  const lines: string[] = [];
  let length = 0;
  let kept = 0;
  let total = 0;
  let sheetsRead = 0;
  let full = true;

  for (const sheet of sheets) {
    const part = parts.get(sheet.path);
    if (!part) {
      continue;
    }
    const reader = new Sheet(strings);
    walkXml(decoder.decode(part), reader);
    const rows = reader.done();
    total += rows.length;
    if (!full) {
      // Past the budget: still counted, so the note can say how much was left.
      continue;
    }
    sheetsRead += 1;
    // Named, because a workbook's sheets are how a person addresses part of it.
    const heading = `## ${sheet.name}`;
    lines.push(heading);
    length += heading.length + 1;
    for (const cells of rows) {
      const line = rowText(cells);
      if (length + line.length + 1 > maxChars) {
        full = false;
        break;
      }
      lines.push(line);
      length += line.length + 1;
      kept += 1;
    }
    lines.push("");
  }

  if (kept === 0) {
    // Every sheet empty, or the first row alone past the budget. Either way an
    // empty success would read as "this workbook has no data".
    throw new XlsxError("it has no readable cells");
  }
  return {
    text: lines.join("\n").trim(),
    sheets: sheetsRead,
    totalSheets: sheets.length,
    rows: kept,
    totalRows: total,
  };
}
