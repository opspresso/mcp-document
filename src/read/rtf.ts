/**
 * RTF to what a model should read.
 *
 * This one earns its place differently from the others. RTF *is* text, so
 * without a reader it does not get refused — it gets through whatever path
 * handles plain files and reaches the model as
 * `{\rtf1\ansi\deff0{\fonttbl{\f0\froman Times;}}...`, thousands of control
 * words with the prose scattered through them. A format that fails by producing
 * garbage rather than an error is worth more than one that simply cannot be
 * opened.
 *
 * The parser is a state machine over four things: control words (`\par`),
 * groups (`{...}`), escapes (`\'e9`, `\\`) and literal text. What makes it more
 * than a `\\\w+` strip is **destinations** — groups whose contents are not prose
 * at all. `{\fonttbl ...}` holds font names, `{\*\generator ...}` holds the
 * writer's version string, and emitting either puts "Times New Roman" in the
 * middle of somebody's letter.
 *
 * **Structure comes from three additions, and one of them is mandatory with the
 * others.** `\outlinelevelN` says a paragraph is a heading — and `\pard`, which
 * resets paragraph properties, has to be honoured in the same change or one
 * heading turns every paragraph after it into a heading too. `{\listtext ...}`
 * is the marker the *writer* rendered for a list item, which is what makes
 * ordered-versus-bullet answerable without touching `\listtable` at all: it
 * reports what was drawn rather than reconstructing a counter. And character
 * formatting is group-scoped — `{` saves it and `}` restores it — so emphasis
 * needs a stack rather than a flag.
 */

import { MAX_TEXT_CHARS } from "../limits.js";
import type { Align, Run } from "../markdown.js";
import { DocumentError } from "../errors.js";
import { drawnMarker, type ReadBlock, type ReadCell, type ReadRow } from "./blocks.js";
import { collapseRuns } from "./lines.js";
import { blocksToMarkdown } from "./serialize.js";

export class RtfError extends DocumentError {}

export interface RtfBlocks {
  blocks: ReadBlock[];
  observed: string[];
}

/**
 * Groups whose contents describe the document rather than say anything.
 *
 * `\*` marks an ignorable destination generally, and these are the named ones
 * every writer emits. Skipping the group wholesale — not just the control word —
 * is the point: the font table's *contents* are what would otherwise appear.
 *
 * `nonshppict` joins them because it is the *second* copy of a picture the
 * `\*\shppict` beside it already described; without it every figure is reported
 * twice.
 */
const DESTINATIONS = new Set([
  "fonttbl",
  "colortbl",
  "stylesheet",
  "listtable",
  "listoverridetable",
  "info",
  "object",
  "themedata",
  "colorschememapping",
  "latentstyles",
  "datastore",
  "generator",
  "xmlnstbl",
  "revtbl",
  "header",
  "footer",
  "headerl",
  "headerr",
  "footerl",
  "footerr",
  "footnote",
  "annotation",
  "comment",
  "nonshppict",
]);

/** Control words that produce a character rather than describing one. */
const LITERALS: Record<string, string> = {
  tab: "\t",
  emdash: "—",
  endash: "–",
  emspace: " ",
  enspace: " ",
  qmspace: " ",
  bullet: "•",
  lquote: "‘",
  rquote: "’",
  ldblquote: "“",
  rdblquote: "”",
  "~": " ",
  "-": "",
  _: "-",
};

interface Emphasis {
  bold?: boolean;
  italic?: boolean;
}

/** `\b` is on and `\b0` is off; a bare word carries no parameter. */
function toggle(parameter: string | undefined): boolean {
  return parameter === undefined || Number(parameter) !== 0;
}

class Reader {
  private readonly blocks: ReadBlock[] = [];
  private runs: Run[] = [];
  private pending = "";
  private emphasis: Emphasis = {};
  /** Character formatting is group-scoped: `{` saves it and `}` restores it. */
  private readonly saved: Emphasis[] = [];
  /** Paragraph properties, which `\pard` resets and `\par` does not. */
  private outline: number | undefined;
  private inTable = false;
  /** The marker the writer drew for this item, captured from `{\listtext …}`. */
  private marker: string | undefined;
  private capturing = false;
  private rows: ReadRow[] = [];
  private cells: ReadCell[] = [];
  private columns = 0;
  readonly observed = new Set<string>();

  emit(value: string): void {
    if (this.capturing) {
      this.marker = (this.marker ?? "") + value;
      return;
    }
    this.pending += value;
  }

  private cut(): void {
    if (this.pending !== "") {
      this.runs.push({ text: this.pending, ...this.emphasis });
      this.pending = "";
    }
  }

  save(): void {
    this.saved.push({ ...this.emphasis });
  }

  restore(): void {
    this.cut();
    this.emphasis = this.saved.pop() ?? {};
    if (this.capturing) {
      this.capturing = false;
    }
  }

  captureMarker(): void {
    this.cut();
    this.marker = "";
    this.capturing = true;
  }

  setEmphasis(word: string, parameter: string | undefined): void {
    this.cut();
    if (word === "plain") {
      this.emphasis = {};
      return;
    }
    const on = toggle(parameter);
    const next = { ...this.emphasis };
    if (word === "b") {
      next.bold = on || undefined;
    } else {
      next.italic = on || undefined;
    }
    this.emphasis = next;
  }

  /** `\pard` — the reset that keeps one heading from spreading to the rest. */
  resetParagraph(): void {
    this.outline = undefined;
    this.inTable = false;
  }

  setOutline(parameter: string | undefined): void {
    const level = Number(parameter ?? "0");
    this.outline = Number.isInteger(level) && level >= 0 && level <= 8 ? level : undefined;
  }

  enterTable(): void {
    this.inTable = true;
  }

  picture(): void {
    this.endParagraph();
    this.blocks.push({ kind: "image", alt: "image" });
  }

  endCell(): void {
    this.cut();
    const runs = collapseRuns(this.runs);
    this.runs = [];
    this.cells.push({ runs });
  }

  endRow(): void {
    if (this.pending !== "" || this.runs.length > 0) {
      this.endCell();
    }
    this.columns = Math.max(this.columns, this.cells.length);
    this.rows.push({ cells: this.cells });
    this.cells = [];
  }

  private finishTable(): void {
    if (this.rows.length === 0 || this.columns === 0) {
      this.rows = [];
      this.columns = 0;
      return;
    }
    this.blocks.push({
      kind: "table",
      rows: this.rows,
      columns: this.columns,
      align: Array.from({ length: this.columns }, () => "left" as Align),
      totalRows: this.rows.length,
      merged: false,
    });
    this.rows = [];
    this.columns = 0;
  }

  endParagraph(): void {
    this.cut();
    const runs = collapseRuns(this.runs);
    this.runs = [];
    const outline = this.outline;
    const marker = this.marker;
    this.marker = undefined;
    // A row that is still open belongs to the table; a paragraph outside one
    // closes whatever table came before it.
    if (this.inTable) {
      return;
    }
    this.finishTable();
    if (runs.length === 0) {
      return;
    }
    if (outline !== undefined) {
      const level = Math.min(outline + 1, 6) as 1 | 2 | 3 | 4 | 5 | 6;
      this.blocks.push({ kind: "heading", level, runs });
      return;
    }
    // `{\listtext 1.\tab}` is the marker the writer rendered. It is reported
    // rather than recounted: the numbering definition is in a destination this
    // reader skips, and inventing a count that disagrees with the file would
    // be the plausible-but-wrong failure the whole reader is built to avoid.
    const drawn = marker === undefined ? undefined : drawnMarker([{ text: `${marker.trim()} ` }]);
    if (drawn) {
      const last = this.blocks[this.blocks.length - 1];
      const item = { runs, depth: 0 };
      if (last?.kind === "list" && last.ordered === drawn.ordered) {
        last.items.push(item);
        return;
      }
      this.blocks.push({
        kind: "list",
        ordered: drawn.ordered,
        items: [item],
        ...(drawn.start !== undefined && drawn.start !== 1 ? { marks: { start: drawn.start } } : {}),
      });
      return;
    }
    this.blocks.push({ kind: "paragraph", runs });
  }

  done(): ReadBlock[] {
    this.endParagraph();
    this.inTable = false;
    if (this.cells.length > 0) {
      this.endRow();
    }
    this.finishTable();
    return this.blocks;
  }
}

export function rtfToBlocks(bytes: Uint8Array): RtfBlocks {
  // Latin-1, not UTF-8: RTF is 7-bit ASCII with everything else escaped, and a
  // stray high byte in a `\'hh` sequence must not become U+FFFD before it is
  // read. Non-ASCII is resolved through the escapes below.
  const source = Buffer.from(bytes).toString("latin1");
  if (!source.trimStart().startsWith("{\\rtf")) {
    throw new RtfError("it does not begin with an RTF header");
  }

  const reader = new Reader();
  /** Depth at which the current destination began; -1 when emitting normally. */
  let skipDepth = -1;
  let depth = 0;
  /** `\uN` is followed by replacement characters this many units wide. */
  let skipUnits = 0;

  const emit = (value: string): void => {
    if (skipDepth !== -1) {
      return;
    }
    if (skipUnits > 0) {
      // The fallback for a Unicode character this reader already took.
      skipUnits -= value.length;
      return;
    }
    reader.emit(value);
  };

  for (let index = 0; index < source.length; index += 1) {
    const character = source[index]!;

    if (character === "{") {
      depth += 1;
      reader.save();
      continue;
    }
    if (character === "}") {
      if (skipDepth !== -1 && depth <= skipDepth) {
        skipDepth = -1;
      }
      depth -= 1;
      reader.restore();
      continue;
    }
    if (character !== "\\") {
      if (character === "\r" || character === "\n") {
        // Source line breaks are formatting of the file, not of the document.
        continue;
      }
      emit(character);
      continue;
    }

    // From here: a control word, a control symbol, or an escape.
    const next = source[index + 1];
    if (next === undefined) {
      break;
    }
    if (next === "\\" || next === "{" || next === "}") {
      emit(next);
      index += 1;
      continue;
    }
    if (next === "'") {
      const hex = source.slice(index + 2, index + 4);
      const code = Number.parseInt(hex, 16);
      // Windows-1252 is what `\ansi` means in practice, and it is what every
      // writer that emits these actually used.
      emit(Number.isNaN(code) ? "" : Buffer.from([code]).toString("latin1"));
      index += 3;
      continue;
    }
    if (next === "*") {
      // `{\*\name ...}` — ignorable whatever `name` turns out to be, with one
      // exception: `\*\shppict` wraps the picture a reader should say was
      // there. Skipping the group is still right; announcing it first is what
      // keeps a figure from vanishing without a trace.
      if (/^\\\*\\shppict\b/.test(source.slice(index, index + 11))) {
        reader.picture();
      }
      if (skipDepth === -1) {
        skipDepth = depth;
      }
      index += 1;
      continue;
    }

    const match = /^([a-zA-Z]+)(-?\d+)? ?/.exec(source.slice(index + 1));
    if (!match) {
      // A control symbol this reader has no meaning for.
      index += 1;
      continue;
    }
    const word = match[1]!;
    const parameter = match[2];
    index += match[0].length;

    if (skipDepth !== -1) {
      continue;
    }
    if (DESTINATIONS.has(word)) {
      skipDepth = depth;
      continue;
    }
    if (word === "pict") {
      // The bytes are megabytes of hex and are skipped; that a picture stood
      // here is the part a reader needs.
      reader.picture();
      skipDepth = depth;
      continue;
    }
    if (word === "listtext" || word === "pntext") {
      reader.captureMarker();
      continue;
    }
    if (word === "u" && parameter !== undefined) {
      // Signed 16-bit: writers emit negative numbers for anything past U+7FFF.
      const code = Number(parameter);
      const point = code < 0 ? code + 65536 : code;
      emit(String.fromCodePoint(point));
      // `\ucN` sets how many fallback characters follow; 1 is the default and
      // is what writers overwhelmingly emit.
      skipUnits = 1;
      continue;
    }
    switch (word) {
      case "pard":
        reader.resetParagraph();
        continue;
      case "outlinelevel":
        reader.setOutline(parameter);
        continue;
      case "intbl":
      case "trowd":
        reader.enterTable();
        continue;
      case "cell":
        reader.endCell();
        continue;
      case "row":
      case "nestrow":
        reader.endRow();
        continue;
      case "par":
      case "line":
      case "page":
      case "sect":
        reader.endParagraph();
        continue;
      case "b":
      case "i":
      case "plain":
        reader.setEmphasis(word, parameter);
        continue;
      default:
        break;
    }
    const literal = LITERALS[word];
    if (literal !== undefined) {
      emit(literal);
    }
  }

  const blocks = reader.done();
  if (blocks.length === 0) {
    throw new RtfError("it has no readable text");
  }
  return { blocks, observed: [...reader.observed] };
}

export interface RtfText {
  text: string;
}

/** The same read, written out — what `document.ts` asks for. */
export function rtfToText(bytes: Uint8Array): RtfText {
  const { text } = blocksToMarkdown(rtfToBlocks(bytes).blocks, MAX_TEXT_CHARS);
  if (text === "") {
    throw new RtfError("it has no readable text");
  }
  return { text };
}
