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
import {
  MAX_INSPECTED_CELLS,
  MAX_SPREADSHEET_CELLS,
  MAX_SPREADSHEET_ROWS,
} from "../limits.js";

export class XlsxError extends DocumentError {}

const WORKBOOK = "xl/workbook.xml";
const WORKBOOK_RELS = "xl/_rels/workbook.xml.rels";
const SHARED_STRINGS = "xl/sharedStrings.xml";

export interface XlsxText {
  text: string;
  /** Sheets that contributed, and how many the workbook holds. */
  sheets: number;
  totalSheets: number;
  hiddenSheets: number;
  /** Rows kept and rows the workbook held, across every sheet read. */
  rows: number;
  totalRows: number;
}

export interface InspectedCell {
  address: string;
  value: string;
  formula?: string;
  error?: string;
}

export interface InspectedSheet {
  name: string;
  state: "visible" | "hidden" | "veryHidden";
  cells: InspectedCell[];
  totalCells: number;
}

export interface XlsxInspection {
  sheets: InspectedSheet[];
  totalSheets: number;
  hiddenSheets: number;
  complete: boolean;
  externalLinks: number;
  macroEnabled: boolean;
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
  private readonly inspected: InspectedCell[] = [];
  private cells: string[] = [];
  private column = 0;
  private row = 0;
  private address = "";
  private type = "";
  private buffer = "";
  private formula = "";
  private capturing = false;
  private capturingFormula = false;
  /** `<v>` inside `<f>` does not exist, but `<is><t>` does — both are values. */
  private inValue = false;
  private cellCount = 0;
  private rowCount = 0;
  private inspectionComplete = true;

  constructor(
    private readonly shared: readonly string[],
    private readonly keepRows = true,
    private readonly inspectionLimit = 0,
  ) {}

  text(value: string): void {
    if (this.capturing) {
      this.buffer += value;
    }
    if (this.capturingFormula) {
      this.formula += value;
    }
  }

  open(name: string, attributes: string, selfClosing: boolean): void {
    switch (localName(name)) {
      case "row":
        this.cells = [];
        this.row = Number(attribute(attributes, "r")) || this.rowCount + 1;
        return;
      case "c": {
        this.cellCount += 1;
        if (this.cellCount > MAX_SPREADSHEET_CELLS) {
          throw new XlsxError(
            `a worksheet has more than ${MAX_SPREADSHEET_CELLS.toLocaleString("en-US")} cells`,
          );
        }
        this.type = attribute(attributes, "t") ?? "";
        const reference = attribute(attributes, "r");
        // Absent addresses mean "the next column", which is what a writer that
        // omits them intends.
        this.column = reference ? columnOf(reference) : this.cells.length;
        this.address = reference ?? `${columnName(this.column)}${this.row}`;
        this.buffer = "";
        this.formula = "";
        this.inValue = false;
        if (selfClosing) {
          this.finishCell();
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
        if (!selfClosing) {
          this.capturingFormula = true;
        }
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
      case "f":
        this.capturingFormula = false;
        return;
      case "c":
        this.finishCell();
        return;
      case "row":
        this.rowCount += 1;
        if (this.rowCount > MAX_SPREADSHEET_ROWS) {
          throw new XlsxError(
            `a worksheet has more than ${MAX_SPREADSHEET_ROWS.toLocaleString("en-US")} rows`,
          );
        }
        if (this.keepRows) {
          this.rows.push(this.cells);
        }
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

  private finishCell(): void {
    const value = this.resolve();
    if (this.keepRows) {
      this.place(value);
    }
    if (this.inspectionLimit > 0) {
      if (this.inspected.length < this.inspectionLimit) {
        this.inspected.push({
          address: this.address,
          value,
          ...(this.formula ? { formula: this.formula } : {}),
          ...(this.type === "e" ? { error: value } : {}),
        });
      } else {
        this.inspectionComplete = false;
      }
    }
  }

  done(): string[][] {
    return this.rows;
  }

  inspection(): { cells: InspectedCell[]; totalCells: number; complete: boolean } {
    return { cells: this.inspected, totalCells: this.cellCount, complete: this.inspectionComplete };
  }
}

function columnName(column: number): string {
  let value = column + 1;
  let name = "";
  while (value > 0) {
    value -= 1;
    name = String.fromCharCode(65 + (value % 26)) + name;
    value = Math.floor(value / 26);
  }
  return name;
}

/** Sheet name → part path, in workbook order. */
function sheetParts(
  workbook: string | undefined,
  rels: string | undefined,
): Array<{ name: string; path: string; state: "visible" | "hidden" | "veryHidden" }> {
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
  const sheets: Array<{
    name: string;
    path: string;
    state: "visible" | "hidden" | "veryHidden";
  }> = [];
  for (const match of workbook.matchAll(/<(?:\w+:)?sheet\b([^>]*)\/?>/g)) {
    const attributes = match[1] ?? "";
    const name = attribute(attributes, "name");
    if (!name) {
      continue;
    }
    const id = attribute(attributes, "r:id") ?? attribute(attributes, "id");
    const declaredState = attribute(attributes, "state");
    const state =
      declaredState === "hidden" || declaredState === "veryHidden" ? declaredState : "visible";
    // The relationship is authoritative; the conventional path is the fallback
    // for a workbook whose rels part is missing or unreadable.
    const path = (id && targets.get(id)) || `xl/worksheets/sheet${sheets.length + 1}.xml`;
    sheets.push({ name, path, state });
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

  const visibleSheets = sheets.filter((sheet) => sheet.state === "visible");
  const hiddenSheets = sheets.length - visibleSheets.length;
  const wanted = visibleSheets.map((sheet) => sheet.path).filter((path) => names.includes(path));
  const parts = read(wanted);

  const lines: string[] = [];
  let length = 0;
  let kept = 0;
  let total = 0;
  let sheetsRead = 0;
  let full = true;

  for (const sheet of visibleSheets) {
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
    hiddenSheets,
    rows: kept,
    totalRows: total,
  };
}

export function inspectXlsx(bytes: Uint8Array, includeHidden = false): XlsxInspection {
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

  const shared = new SharedStrings();
  const sharedXml = head.get(SHARED_STRINGS);
  if (sharedXml) {
    walkXml(decoder.decode(sharedXml), shared);
  }
  const strings = shared.done();

  const visible = includeHidden ? sheets : sheets.filter((sheet) => sheet.state === "visible");
  const wanted = visible.map((sheet) => sheet.path).filter((path) => names.includes(path));
  const parts = read(wanted);
  const inspected: InspectedSheet[] = [];
  let remaining = MAX_INSPECTED_CELLS;
  let complete = true;
  for (const sheet of visible) {
    if (remaining === 0) {
      complete = false;
      break;
    }
    const part = parts.get(sheet.path);
    if (!part) {
      continue;
    }
    const reader = new Sheet(strings, false, remaining);
    walkXml(decoder.decode(part), reader);
    const result = reader.inspection();
    inspected.push({ name: sheet.name, state: sheet.state, cells: result.cells, totalCells: result.totalCells });
    remaining = Math.max(0, remaining - result.cells.length);
    complete &&= result.complete;
  }
  return {
    sheets: inspected,
    totalSheets: sheets.length,
    hiddenSheets: sheets.filter((sheet) => sheet.state !== "visible").length,
    complete,
    externalLinks: names.filter((name) => name.startsWith("xl/externalLinks/") && name.endsWith(".xml")).length,
    macroEnabled: names.includes("xl/vbaProject.bin"),
  };
}
