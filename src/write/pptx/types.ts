/**
 * The presentation AST — what sits between the Markdown document and the XML.
 *
 * The planner produces this and the renderer consumes it, and neither sees the
 * other. The split exists because the two jobs change for different reasons: a
 * new slide archetype is a planning decision (what does this content *mean*?)
 * and a rendering decision (what shapes say that?), and holding both in one
 * module is how the original single-file renderer reached a thousand lines.
 */

import type { Align, Run } from "../../markdown.js";

/** How a line of body text is set. Runs carry their own bold and italic on top. */
export interface Style {
  size: number;
  bold?: boolean;
  italic?: boolean;
  mono?: boolean;
  colour?: string;
  /** Opacity in DrawingML's thousandths of a percent — 45000 is 45%. */
  alpha?: number;
  /** A colour for the first run only, which is where a list keeps its marker. */
  marker?: string;
  /** Indent levels, each `INDENT_COLUMNS` wide. */
  indent: number;
  /** Space above, in hundredths of a point. */
  before?: number;
  /** Table cells and the closing slide set this; body text is flush left. */
  align?: Align;
}

/**
 * One thing that occupies vertical space.
 *
 * Blocks are flattened to these before a slide is filled, because packing is
 * about lines and a block is not a line. A table is the exception that stays
 * whole: its rows have to line up, so it is split by row rather than by line.
 */
export type Piece =
  | { kind: "text"; runs: Run[]; style: Style }
  | { kind: "table"; header: Run[][]; rows: Run[][][]; align: Align[] };

/**
 * One slide, typed by what it is for rather than by what is on it.
 *
 * Four archetypes so far. A `cover` holds a title and at most a subtitle —
 * everything else the author wrote under the opening `#` moves to the slides
 * after it. A `section` is a mid-document `#`: a divider, numbered in the order
 * the dividers appear. A `closing` is the last section when its title says it
 * is one — 감사합니다, Thank you, Q&A — set centred on the cover's ground.
 * Everything else is `content`.
 */
export type Slide =
  | { type: "cover"; title: Run[]; subtitle?: Run[] }
  | { type: "section"; title: Run[]; ordinal: number }
  | { type: "content"; title?: Run[]; pieces: Piece[] }
  | { type: "closing"; title: Run[]; pieces: Piece[] };

export interface Presentation {
  slides: Slide[];
}
