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
 */

import { detect, type Format } from "../detect.js";
import { MAX_TEXT_CHARS, truncateText } from "../limits.js";
import type { DocumentSource } from "../source.js";
import { DocumentError } from "../errors.js";
import { docxToText } from "./docx.js";
import { hwpToText } from "./hwp5.js";
import { hwpxToText } from "./hwpx.js";
import { pdfToText } from "./pdf.js";
import { plainToText } from "./plain.js";

export class UnsupportedDocument extends DocumentError {}

export interface ReadResult {
  /** Already inside the character budget. The provenance header is not applied here. */
  text: string;
  format: Format;
  /** What came back, in the document's own units. */
  note?: string;
}

/**
 * Cut to the budget, and say which of the two things happened.
 *
 * `whole` is the note for a document that fitted — it is not "nothing to say":
 * the formats that leave parts out (DOCX's headers, a section list) have to say
 * so on the successful path, because that is the path where nobody is looking
 * for a caveat.
 */
function fit(text: string, whole?: string): { text: string; note?: string } {
  const cut = truncateText(text, MAX_TEXT_CHARS);
  if (cut.note) {
    return cut;
  }
  return whole ? { text: cut.text, note: whole } : { text: cut.text };
}

export async function readDocument(source: DocumentSource): Promise<ReadResult> {
  const detection = detect(source.bytes, source.mimeType, source.filename);
  if (detection.format === "unsupported") {
    throw new UnsupportedDocument(detection.reason);
  }
  const format = detection.format;

  if (format === "pdf") {
    // The only reader that budgets for itself: it drops whole pages rather than
    // cutting mid-page, so the cut and the note have to be made together.
    const { text, note } = await pdfToText(source.bytes, MAX_TEXT_CHARS);
    return { text, format, note };
  }

  if (format === "docx") {
    const { text } = docxToText(source.bytes);
    return { format, ...fit(text, "the document body, without headers, footers or footnotes") };
  }

  if (format === "hwpx") {
    const { text, sections } = hwpxToText(source.bytes);
    return { format, ...fit(text, `all ${sections} section(s)`) };
  }

  if (format === "hwp") {
    const { text, sections, version } = hwpToText(source.bytes);
    return { format, ...fit(text, `all ${sections} section(s) of an HWP ${version} document`) };
  }

  const { text, charset } = plainToText(source.bytes, source.charset);
  if (text === "") {
    throw new UnsupportedDocument("the document has no readable text");
  }
  // The charset is stated only when it was not the one everybody assumes.
  // Saying "decoded as utf-8" on every plain file is noise; saying it on the
  // EUC-KR one is the difference between trusting the text and not.
  return { format, ...fit(text, charset === "utf-8" ? undefined : `text decoded as ${charset}`) };
}
