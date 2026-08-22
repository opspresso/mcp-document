/**
 * One document in, one piece of text out.
 *
 * The dispatch is here rather than in `tools.ts` so that the note each format
 * produces is written next to the reader that knows what it means. "All 12
 * pages" and "the whole body, without headers or footers" answer the same
 * question — did this reach the end of the document — in the only units their
 * format has.
 *
 * Every path returns text or raises. None of them returns an empty success: an
 * empty string reads as "the document is empty", which is a different and much
 * more damaging claim than "I could not read it".
 *
 * Office formats that need a parser. PDF and plain text left with the
 * URL side: they need no parser Agent Studio lacks, so routing one here was a
 * network round trip to reach the same `unpdf` — and a third copy of the
 * extraction to keep in step.
 */

import { detect, type Format } from "../detect.js";
import { MAX_TEXT_CHARS, truncateText } from "../limits.js";
import type { DocumentSource } from "../source.js";
import { DocumentError } from "../errors.js";
import { docxToText } from "./docx.js";
import { hwpToText } from "./hwp5.js";
import { hwpxToText } from "./hwpx.js";
import { odfToText } from "./odf.js";
import { pptxToText } from "./pptx.js";
import { rtfToText } from "./rtf.js";
import { xlsxToText } from "./xlsx.js";

export class UnsupportedDocument extends DocumentError {}

export interface ReadResult {
  /** Already inside the character budget. The provenance header is not applied here. */
  text: string;
  format: Format;
  /** What came back, in the document's own units. */
  note?: string;
  /** Whether every readable unit fitted inside the text budget. */
  complete: boolean;
  /** Parts deliberately left out of the text representation. */
  omissions: string[];
  /** Counts in the format's own units. */
  counts?: Record<string, number>;
}

/**
 * Cut to the budget, and say which of the two things happened.
 *
 * `whole` is the note for a document that fitted — it is not "nothing to say":
 * the formats that leave parts out (DOCX's headers, a section list) have to say
 * so on the successful path, because that is the path where nobody is looking
 * for a caveat.
 */
function fit(text: string, whole?: string): { text: string; note?: string; complete: boolean } {
  const cut = truncateText(text, MAX_TEXT_CHARS);
  if (cut.note) {
    return { ...cut, complete: false };
  }
  return whole ? { text: cut.text, note: whole, complete: true } : { text: cut.text, complete: true };
}

export async function readDocument(source: DocumentSource): Promise<ReadResult> {
  const detection = detect(source.bytes, source.mimeType, source.filename);
  if (detection.format === "unsupported") {
    throw new UnsupportedDocument(detection.reason);
  }
  const format = detection.format;

  if (format === "docx") {
    const { text } = docxToText(source.bytes);
    return {
      format,
      ...fit(text, "the document body, without headers, footers or footnotes"),
      omissions: ["headers", "footers", "footnotes", "comments", "tracked deletions", "formatting"],
    };
  }

  if (format === "hwpx") {
    const { text, sections } = hwpxToText(source.bytes);
    return {
      format,
      ...fit(text, `all ${sections} section(s)`),
      omissions: ["formatting"],
      counts: { sections },
    };
  }

  if (format === "xlsx") {
    // The only reader besides the old PDF one that budgets for itself: a sheet
    // can dwarf any text budget, and cutting mid-row would leave a line whose
    // columns no longer line up with its neighbours'.
    const { text, sheets, totalSheets, hiddenSheets, rows, totalRows } = xlsxToText(
      source.bytes,
      MAX_TEXT_CHARS,
    );
    const whole = sheets === totalSheets - hiddenSheets && rows === totalRows;
    return {
      text,
      format,
      complete: whole,
      omissions: [
        "formulas",
        "cell formatting",
        "comments",
        "macros",
        ...(hiddenSheets > 0 ? ["hidden sheets"] : []),
      ],
      counts: { sheets, totalSheets, hiddenSheets, rows, totalRows },
      note: whole
        ? `all ${sheets} visible sheet(s), ${totalRows} row(s)`
        : `${rows} of ${totalRows} row(s) across ${sheets} visible sheet(s)`,
    };
  }

  if (format === "pptx") {
    const { text, slides } = pptxToText(source.bytes);
    return {
      format,
      ...fit(text, `all ${slides} slide(s), without speaker notes`),
      omissions: ["speaker notes", "comments", "animations", "formatting"],
      counts: { slides },
    };
  }

  if (format === "odf") {
    const { text, kind, parts } = odfToText(source.bytes);
    const unit = kind === "spreadsheet" ? "sheet" : "slide";
    return {
      format,
      ...fit(
        text,
        parts === undefined
          ? "the document body, without headers or footers"
          : `all ${parts} ${unit}(s)`,
      ),
      omissions: ["formatting", "headers", "footers"],
      ...(parts === undefined ? {} : { counts: { [unit === "sheet" ? "sheets" : "slides"]: parts } }),
    };
  }

  if (format === "rtf") {
    const { text } = rtfToText(source.bytes);
    return {
      format,
      ...fit(text, "the document body, without headers or footers"),
      omissions: ["headers", "footers", "formatting"],
    };
  }

  const { text, sections, version } = hwpToText(source.bytes);
  return {
    format,
    ...fit(text, `all ${sections} section(s) of an HWP ${version} document`),
    omissions: ["formatting"],
    counts: { sections },
  };
}
