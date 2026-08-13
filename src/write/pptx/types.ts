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
  /** Indent levels, each `INDENT_COLUMNS` wide. */
  indent: number;
  /** Space above, in hundredths of a point. */
  before?: number;
  /** Only a table cell sets this; body text is always flush left. */
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
 * `cover` and `content` are the two the original renderer knew. The union is
 * what the professional archetypes — section, cards, metrics, comparison —
 * extend, each carrying the payload its layout needs and nothing else.
 */
export type Slide =
  | { type: "cover"; title?: Run[]; pieces: Piece[] }
  | { type: "content"; title?: Run[]; pieces: Piece[] };

export interface Presentation {
  slides: Slide[];
}
