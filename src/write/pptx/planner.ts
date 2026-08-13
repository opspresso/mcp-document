/**
 * Markdown in, a presentation out — the deck's decisions, with no XML in them.
 *
 * A deck is a sequence of fixed-size boxes, so unlike the DOCX and HWPX writers
 * this planner has to decide **where a slide ends** — there is no reflow to fall
 * back on, and text past the bottom of a box is simply not on the screen.
 *
 * Three decisions follow from that.
 *
 * **A slide opens at a level 1 or 2 heading.** Nothing else in Markdown says
 * "new slide", and the alternative — a horizontal rule, as Marp uses — turns a
 * decorative divider into a page break in documents that were never written as
 * decks. Level 3 and below stay in the body as bold sub-headings. It also makes
 * the round trip symmetric: `read/pptx.ts` emits a deck as `## Slide N`, so
 * reading a deck and writing it back produces the same slides.
 *
 * **Content that does not fit continues on the next slide**, titled with
 * `(계속)`. The line count is `layout.ts`'s estimate, in the same
 * character-width currency `write/table.ts` uses for column widths.
 *
 * **Blocks are flattened to lines before anything is packed.** A list becomes one
 * piece per item with its marker already resolved, which is what lets a numbered
 * list be split across two slides and still count 4, 5, 6 rather than starting
 * again at 1. Only a table stays whole, and it splits by row with its header
 * repeated.
 */

import type { Block, MarkdownDocument, Run } from "../../markdown.js";
import { PALETTE } from "../theme.js";
import {
  BODY_LINES,
  BODY_SIZE,
  CODE_SIZE,
  COVER_LINES,
  LINE_HEIGHT,
  ROW_HEIGHT,
  SUBHEADING_SIZES,
  linesOf,
} from "./layout.js";
import type { Piece, Presentation, Slide, Style } from "./types.js";

/** What a continued slide's title says, so the reader knows it is not a new topic. */
export const CONTINUED = " (계속)";

const BODY_STYLE: Style = { size: BODY_SIZE, indent: 0 };

function piecesOf(block: Block): Piece[] {
  switch (block.kind) {
    case "heading":
      // Levels 1 and 2 opened a slide and never reach here; 3 to 6 are set as
      // bold sub-headings, which is what they read as on a slide.
      return [
        {
          kind: "text",
          runs: block.runs,
          style: {
            size: SUBHEADING_SIZES[Math.min(block.level, 6) - 3] ?? BODY_SIZE,
            bold: true,
            colour: PALETTE.brand,
            indent: 0,
            before: 600,
          },
        },
      ];
    case "paragraph":
      return [{ kind: "text", runs: block.runs, style: BODY_STYLE }];
    case "list": {
      // The marker is resolved here rather than at render time, which is what
      // lets a numbered list survive being split across two slides.
      const counters: number[] = [];
      return block.items.map((item) => {
        counters.length = item.depth + 1;
        counters[item.depth] = (counters[item.depth] ?? 0) + 1;
        const marker = block.ordered ? `${counters[item.depth]}. ` : "• ";
        return {
          kind: "text",
          runs: [{ text: marker }, ...item.runs],
          style: { size: BODY_SIZE, indent: item.depth + 1 },
        };
      });
    }
    case "code":
      return (block.text === "" ? [""] : block.text.split("\n")).map((line) => ({
        kind: "text",
        runs: [{ text: line }],
        style: { size: CODE_SIZE, mono: true, indent: 1 },
      }));
    case "quote":
      return [
        {
          kind: "text",
          runs: block.runs,
          style: { size: BODY_SIZE, italic: true, colour: PALETTE.inkMuted, indent: 1 },
        },
      ];
    case "table":
      return [{ kind: "table", header: block.header, rows: block.rows, align: block.align }];
    case "rule":
      // A drawn line would be a shape of its own in the middle of a text box,
      // which is a second layout problem for a mark this small. The same row of
      // dashes `write/hwpx.ts` settles for.
      return [{ kind: "text", runs: [{ text: "─".repeat(30) }], style: { size: BODY_SIZE, colour: PALETTE.rule, indent: 0 } }];
  }
}

/**
 * A table cut to `rows` data rows, and what is left of it.
 *
 * The header repeats on the second half, because a column of numbers with no
 * heading above it is a column nobody can read.
 */
function splitTable(
  piece: Extract<Piece, { kind: "table" }>,
  rows: number,
): { head: Piece; rest: Piece } | undefined {
  if (rows < 1 || rows >= piece.rows.length) {
    return undefined;
  }
  return {
    head: { kind: "table", header: piece.header, rows: piece.rows.slice(0, rows), align: piece.align },
    rest: { kind: "table", header: piece.header, rows: piece.rows.slice(rows), align: piece.align },
  };
}

/**
 * Fill slides with pieces, greedily.
 *
 * `budget` is a function of the slide number because a cover holds less than an
 * ordinary slide: its title sits in the middle of the page, so what follows has
 * a quarter of the room.
 */
function pack(pieces: readonly Piece[], budget: (slide: number) => number): Piece[][] {
  const slides: Piece[][] = [];
  let current: Piece[] = [];
  let used = 0;

  const flush = (): void => {
    if (current.length > 0) {
      slides.push(current);
      current = [];
      used = 0;
    }
  };

  for (const piece of pieces) {
    let pending: Piece | undefined = piece;
    while (pending) {
      const room = budget(slides.length) - used;
      const cost = linesOf(pending);
      if (cost <= room) {
        current.push(pending);
        used += cost;
        pending = undefined;
        continue;
      }
      if (room > 0 && pending.kind === "table") {
        // A table is the one piece worth cutting mid-way: its rows are
        // independent, and moving the whole thing leaves a slide half empty.
        const rows = Math.max(1, Math.floor((room * LINE_HEIGHT) / ROW_HEIGHT) - 1);
        const split = splitTable(pending, rows);
        if (split) {
          current.push(split.head);
          flush();
          pending = split.rest;
          continue;
        }
      }
      if (current.length === 0) {
        // Nothing to move it past: a single piece taller than a slide goes on
        // one anyway, and PowerPoint's autofit shrinks what is left over.
        current.push(pending);
        flush();
        pending = undefined;
        continue;
      }
      flush();
    }
  }
  flush();
  return slides.length > 0 ? slides : [[]];
}

interface Section {
  title?: Run[];
  cover: boolean;
  blocks: Block[];
}

/** Cut the document at every level 1 and 2 heading. */
function sectionsOf(document: MarkdownDocument): Section[] {
  const sections: Section[] = [];
  let current: Section | undefined;
  for (const block of document.blocks) {
    if (block.kind === "heading" && block.level <= 2) {
      // Only a document that *opens* with a level 1 heading gets a cover: a `#`
      // halfway down is a new section, and centring its title would read as the
      // deck starting over.
      current = { title: block.runs, cover: sections.length === 0 && block.level === 1, blocks: [] };
      sections.push(current);
      continue;
    }
    if (!current) {
      current = { cover: false, blocks: [] };
      sections.push(current);
    }
    current.blocks.push(block);
  }
  return sections.length > 0 ? sections : [{ cover: false, blocks: [] }];
}

export function plan(document: MarkdownDocument): Presentation {
  const slides = sectionsOf(document).flatMap((section): Slide[] => {
    const pieces = section.blocks.flatMap(piecesOf);
    const pages = pack(pieces, (slide) =>
      section.cover && slide === 0 ? COVER_LINES : BODY_LINES,
    );
    return pages.map((page, index) => ({
      type: section.cover && index === 0 ? "cover" : "content",
      pieces: page,
      ...(section.title
        ? { title: index === 0 ? section.title : [...section.title, { text: CONTINUED }] }
        : {}),
    }));
  });
  return { slides };
}
