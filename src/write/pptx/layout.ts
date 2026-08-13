/**
 * Slide geometry, and the arithmetic that estimates what fits.
 *
 * Everything here is in EMU — 914,400 to the inch, which is what every
 * DrawingML measurement is in — derived from the point scale `write/theme.ts`
 * decides. The line count is an estimate: nothing here measures a font, because
 * unlike the PDF renderer this package embeds none and cannot know what the
 * reader's PowerPoint will substitute. Being a line out puts a line nearer the
 * edge than intended, which is a better failure than letting a slide overflow
 * and lose it.
 */

import type { Run } from "../../markdown.js";
import { DECK, centiPoints, emu } from "../theme.js";
import type { Piece } from "./types.js";

/** EMU, which is what every DrawingML measurement is in: 914,400 to the inch. */
export const EMU_PER_POINT = 12700;

/** 16:9, which is what every deck has been since PowerPoint 2013. */
export const SLIDE_WIDTH = 12192000;
export const SLIDE_HEIGHT = 6858000;

export const SIDE_MARGIN = 685800;
export const CONTENT_WIDTH = SLIDE_WIDTH - SIDE_MARGIN * 2;

/** The title bar of an ordinary slide, and the body beneath it. */
export const TITLE_BOX = { y: 457200, height: 1143000 };
export const BODY_BOX = { y: 1714500, height: 4457700 };

/** The opening slide, whose title sits where a reader expects a cover's to. */
export const COVER_TITLE_BOX = { y: 2133600, height: 1371600 };
export const COVER_BODY_BOX = { y: 3657600, height: 1371600 };

/** Hundredths of a point, which is what `sz` is in. */
export const TITLE_SIZE = centiPoints(DECK.title);
export const COVER_TITLE_SIZE = centiPoints(DECK.coverTitle);
export const BODY_SIZE = centiPoints(DECK.body);
export const CODE_SIZE = centiPoints(DECK.code);
/** Headings 3 to 6, which stay in the body rather than opening a slide. */
export const SUBHEADING_SIZES = DECK.subheadings.map(centiPoints);

/**
 * The bar under a slide's title.
 *
 * This is where the console's lavender page ended up. A deck is the one output
 * that can afford a full-bleed tint — a slide is shown, not printed by the
 * hundred — so the cover and section slides take the tint, and every ordinary
 * slide gets this instead: a short brand rule under the title, which says the
 * same thing without putting a colour field behind every bullet.
 */
export const TITLE_RULE = { width: emu(48), height: emu(3), gap: emu(6) };

/** The slide-number box, sitting in the bottom margin. */
export const NUMBER_BOX = { width: emu(48), height: emu(16) };

/** A line of body text, with the leading that goes with it. */
export const LINE_HEIGHT = Math.round(((BODY_SIZE * 1.35) / 100) * EMU_PER_POINT);
/** A table row, which is worth one and a half lines of prose. */
export const ROW_HEIGHT = Math.round(LINE_HEIGHT * 1.5);

/**
 * How many body lines a slide holds, and how wide one is in character units.
 *
 * The width is in units of one CJK character at the body size: 852pt of box
 * divided by 18pt a glyph, less a little for the fact that a line ending exactly
 * at the edge wraps. Latin counts as half, which is close enough for prose and
 * is the same approximation `write/table.ts` makes when it shares out columns.
 */
export const BODY_LINES = Math.floor(BODY_BOX.height / LINE_HEIGHT);
export const COVER_LINES = Math.floor(COVER_BODY_BOX.height / LINE_HEIGHT);
export const COLUMNS = 45;

/** One indent step, as a share of the line — a list level, or a quote's bar. */
export const INDENT_COLUMNS = 2;
export const INDENT_EMU = 274320;

/**
 * Scripts that take a full character width.
 *
 * The same set `write/pdf.ts` breaks lines on, and deliberately a second copy:
 * importing it from there would pull `pdf-lib` and a 2MB font loader into a
 * renderer that embeds nothing.
 */
const WIDE =
  /[ᄀ-ᇿ⺀-〿぀-ヿ㄰-㆏㐀-䶿一-鿿ꥠ-꥿가-퟿豈-﫿︰-﹏＀-｠￠-￦]/;

/** Text width in character units: a wide character is one, everything else is half. */
export function widthOf(text: string): number {
  let width = 0;
  for (const character of text) {
    width += WIDE.test(character) ? 1 : 0.5;
  }
  return width;
}

export function textOf(runs: readonly Run[]): string {
  return runs.map((run) => run.text).join("");
}

/** How many body lines a piece takes, which is the currency slides are filled in. */
export function linesOf(piece: Piece): number {
  if (piece.kind === "table") {
    return Math.ceil(((piece.rows.length + 1) * ROW_HEIGHT) / LINE_HEIGHT);
  }
  const width = Math.max(1, COLUMNS - piece.style.indent * INDENT_COLUMNS);
  const wrapped = Math.max(1, Math.ceil(widthOf(textOf(piece.runs)) / width));
  return wrapped + (piece.style.before ? 1 : 0);
}
