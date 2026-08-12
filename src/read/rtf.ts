/**
 * RTF to the text a model should read.
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
 */

import { DocumentError } from "../errors.js";
import { normalize } from "./lines.js";

export class RtfError extends DocumentError {}

export interface RtfText {
  text: string;
}

/**
 * Groups whose contents describe the document rather than say anything.
 *
 * `\*` marks an ignorable destination generally, and these are the named ones
 * every writer emits. Skipping the group wholesale — not just the control word —
 * is the point: the font table's *contents* are what would otherwise appear.
 */
const DESTINATIONS = new Set([
  "fonttbl",
  "colortbl",
  "stylesheet",
  "listtable",
  "listoverridetable",
  "info",
  "pict",
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
]);

/** Control words that produce a character rather than describing one. */
const LITERALS: Record<string, string> = {
  par: "\n",
  line: "\n",
  tab: "\t",
  page: "\n",
  sect: "\n",
  cell: " | ",
  row: "\n",
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
  "~": " ",
  "-": "",
  _: "-",
};

export function rtfToText(bytes: Uint8Array): RtfText {
  // Latin-1, not UTF-8: RTF is 7-bit ASCII with everything else escaped, and a
  // stray high byte in a `\'hh` sequence must not become U+FFFD before it is
  // read. Non-ASCII is resolved through the escapes below.
  const source = Buffer.from(bytes).toString("latin1");
  if (!source.trimStart().startsWith("{\\rtf")) {
    throw new RtfError("it does not begin with an RTF header");
  }

  const out: string[] = [];
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
    out.push(value);
  };

  for (let index = 0; index < source.length; index += 1) {
    const character = source[index]!;

    if (character === "{") {
      depth += 1;
      continue;
    }
    if (character === "}") {
      if (skipDepth !== -1 && depth <= skipDepth) {
        skipDepth = -1;
      }
      depth -= 1;
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
      // `{\*\name ...}` — ignorable whatever `name` turns out to be.
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
    const literal = LITERALS[word];
    if (literal !== undefined) {
      emit(literal);
    }
  }

  const text = normalize(out.join("").split("\n"));
  if (text === "") {
    throw new RtfError("it has no readable text");
  }
  return { text };
}
