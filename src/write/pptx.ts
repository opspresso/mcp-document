/**
 * A document AST to PPTX.
 *
 * The fourth renderer, and the first whose output is not a page of prose. A deck
 * is a sequence of fixed-size boxes, so unlike the DOCX and HWPX writers this one
 * has to decide **where a slide ends** — there is no reflow to fall back on, and
 * text past the bottom of a box is simply not on the screen.
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
 * `(계속)`. The line count is an estimate — nothing here measures a font, because
 * unlike the PDF renderer this one embeds none and cannot know which face the
 * reader's PowerPoint will substitute — so a character-width approximation
 * decides, the same currency `write/table.ts` already uses for column widths.
 * Being one line out puts a line closer to the edge than intended; the
 * alternative, letting a slide overflow, loses the line entirely.
 *
 * **Blocks are flattened to lines before anything is packed.** A list becomes one
 * piece per item with its marker already resolved, which is what lets a numbered
 * list be split across two slides and still count 4, 5, 6 rather than starting
 * again at 1. Only a table stays whole, and it splits by row with its header
 * repeated.
 *
 * The package itself is written by hand, as the other OOXML writer here is. A
 * deck needs more scaffolding than a document — PowerPoint refuses a file with no
 * theme or no slide master — so the master, the two layouts and the theme are
 * fixed constants, and only the slides are a function of the input.
 */

import { escapeXml } from "../xml.js";
import { buildZip } from "../zip.js";
import type { Align, Block, MarkdownDocument, Run } from "../markdown.js";
import { columnShares } from "./table.js";
import { CHART, DECK, PALETTE, centiPoints, emu } from "./theme.js";
import { PRODUCER } from "../version.js";

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** EMU, which is what every DrawingML measurement is in: 914,400 to the inch. */
const EMU_PER_POINT = 12700;

/** 16:9, which is what every deck has been since PowerPoint 2013. */
const SLIDE_WIDTH = 12192000;
const SLIDE_HEIGHT = 6858000;

const SIDE_MARGIN = 685800;
const CONTENT_WIDTH = SLIDE_WIDTH - SIDE_MARGIN * 2;

/** The title bar of an ordinary slide, and the body beneath it. */
const TITLE_BOX = { y: 457200, height: 1143000 };
const BODY_BOX = { y: 1714500, height: 4457700 };

/** The opening slide, whose title sits where a reader expects a cover's to. */
const COVER_TITLE_BOX = { y: 2133600, height: 1371600 };
const COVER_BODY_BOX = { y: 3657600, height: 1371600 };

/** Hundredths of a point, which is what `sz` is in. */
const TITLE_SIZE = centiPoints(DECK.title);
const COVER_TITLE_SIZE = centiPoints(DECK.coverTitle);
const BODY_SIZE = centiPoints(DECK.body);
const CODE_SIZE = centiPoints(DECK.code);
/** Headings 3 to 6, which stay in the body rather than opening a slide. */
const SUBHEADING_SIZES = DECK.subheadings.map(centiPoints);

const INK = PALETTE.ink;
const LINK_COLOUR = PALETTE.brandDeep;
const MUTED = PALETTE.inkMuted;
const BORDER = PALETTE.rule;
const HEADER_FILL = PALETTE.brand;

/**
 * The bar under a slide's title.
 *
 * This is where the console's lavender page ended up. A deck is the one output
 * that can afford a full-bleed tint — a slide is shown, not printed by the
 * hundred — so the cover and section slides take the tint, and every ordinary
 * slide gets this instead: a short brand rule under the title, which says the
 * same thing without putting a colour field behind every bullet.
 */
const TITLE_RULE = { width: emu(48), height: emu(3), gap: emu(6) };

/**
 * The slide-number box, sitting in the bottom margin.
 *
 * The field id is a fixed GUID rather than a generated one, for the reason the
 * created date is passed in: the bytes have to follow from the input alone, and
 * nothing distinguishes one deck's number field from another's.
 */
const NUMBER_BOX = { width: emu(48), height: emu(16) };

/**
 * PowerPoint's own default table style, referenced by GUID.
 *
 * The GUID names a style built into PowerPoint itself; the `tableStyles.xml`
 * part this package carries declares it as the default and defines nothing.
 * That is exactly what PowerPoint writes for a fresh file — a table whose
 * `tblPr` names no style is a table no native file contains, and Windows
 * PowerPoint treats several such never-written shapes as damage.
 */
const TABLE_STYLE_ID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
const NUMBER_FIELD_ID = "{7B4A2F5C-0E31-4C6D-9A2B-51D0C8E3F614}";

/** A line of body text, with the leading that goes with it. */
const LINE_HEIGHT = Math.round(((BODY_SIZE * 1.35) / 100) * EMU_PER_POINT);
/** A table row, which is worth one and a half lines of prose. */
const ROW_HEIGHT = Math.round(LINE_HEIGHT * 1.5);

/**
 * How many body lines a slide holds, and how wide one is in character units.
 *
 * The width is in units of one CJK character at the body size: 852pt of box
 * divided by 18pt a glyph, less a little for the fact that a line ending exactly
 * at the edge wraps. Latin counts as half, which is close enough for prose and
 * is the same approximation `write/table.ts` makes when it shares out columns.
 */
const BODY_LINES = Math.floor(BODY_BOX.height / LINE_HEIGHT);
const COVER_LINES = Math.floor(COVER_BODY_BOX.height / LINE_HEIGHT);
const COLUMNS = 45;

/** One indent step, as a share of the line — a list level, or a quote's bar. */
const INDENT_COLUMNS = 2;
const INDENT_EMU = 274320;

/** What a continued slide's title says, so the reader knows it is not a new topic. */
const CONTINUED = " (계속)";

/**
 * Scripts that take a full character width.
 *
 * The same set `write/pdf.ts` breaks lines on, and deliberately a second copy:
 * importing it from there would pull `pdf-lib` and a 2MB font loader into a
 * renderer that embeds nothing.
 */
const WIDE =
  /[ᄀ-ᇿ⺀-〿぀-ヿ㄰-㆏㐀-䶿一-鿿ꥠ-꥿가-퟿豈-﫿︰-﹏＀-｠￠-￦]/;

const encoder = new TextEncoder();

function part(xml: string): Uint8Array {
  return encoder.encode(DECLARATION + xml);
}

/** Text width in character units: a wide character is one, everything else is half. */
function widthOf(text: string): number {
  let width = 0;
  for (const character of text) {
    width += WIDE.test(character) ? 1 : 0.5;
  }
  return width;
}

function textOf(runs: readonly Run[]): string {
  return runs.map((run) => run.text).join("");
}

/* ------------------------------------------------------------------ pieces */

/** How a line of body text is set. Runs carry their own bold and italic on top. */
interface Style {
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
type Piece =
  | { kind: "text"; runs: Run[]; style: Style }
  | { kind: "table"; header: Run[][]; rows: Run[][][]; align: Align[] };

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
          style: { size: BODY_SIZE, italic: true, colour: MUTED, indent: 1 },
        },
      ];
    case "table":
      return [{ kind: "table", header: block.header, rows: block.rows, align: block.align }];
    case "rule":
      // A drawn line would be a shape of its own in the middle of a text box,
      // which is a second layout problem for a mark this small. The same row of
      // dashes `write/hwpx.ts` settles for.
      return [{ kind: "text", runs: [{ text: "─".repeat(30) }], style: { size: BODY_SIZE, colour: BORDER, indent: 0 } }];
  }
}

/** How many body lines a piece takes, which is the currency slides are filled in. */
function linesOf(piece: Piece): number {
  if (piece.kind === "table") {
    return Math.ceil(((piece.rows.length + 1) * ROW_HEIGHT) / LINE_HEIGHT);
  }
  const width = Math.max(1, COLUMNS - piece.style.indent * INDENT_COLUMNS);
  const wrapped = Math.max(1, Math.ceil(widthOf(textOf(piece.runs)) / width));
  return wrapped + (piece.style.before ? 1 : 0);
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

/* ------------------------------------------------------------------ slides */

interface Slide {
  title?: Run[];
  cover: boolean;
  pieces: Piece[];
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

function slidesOf(document: MarkdownDocument): Slide[] {
  return sectionsOf(document).flatMap((section) => {
    const pieces = section.blocks.flatMap(piecesOf);
    const pages = pack(pieces, (slide) =>
      section.cover && slide === 0 ? COVER_LINES : BODY_LINES,
    );
    return pages.map((page, index) => ({
      cover: section.cover && index === 0,
      pieces: page,
      ...(section.title
        ? { title: index === 0 ? section.title : [...section.title, { text: CONTINUED }] }
        : {}),
    }));
  });
}

/* --------------------------------------------------------------- rendering */

/** `sz` and the rest of `a:rPr`, for one run inside a paragraph of a given style. */
function runProperties(run: Run, style: Style, linkId?: string): string {
  const bold = run.bold || style.bold;
  const italic = run.italic || style.italic;
  const mono = run.code || style.mono;
  const colour = run.href ? LINK_COLOUR : (style.colour ?? INK);
  return (
    `<a:rPr lang="en-US" sz="${style.size}"${bold ? ' b="1"' : ""}${italic ? ' i="1"' : ""}` +
    `${run.href ? ' u="sng"' : ""} dirty="0">` +
    `<a:solidFill><a:srgbClr val="${colour}"/></a:solidFill>` +
    // Only the Latin face is named, and only for code. Naming a face for prose
    // is what `write/docx.ts` refuses to do for the same reason: a font chosen
    // here is a font the reader's machine may not have, and the substitute is
    // then chosen by nobody.
    (mono ? '<a:latin typeface="Consolas"/><a:cs typeface="Consolas"/>' : "") +
    (linkId ? `<a:hlinkClick xmlns:r="${R}" r:id="${linkId}"/>` : "") +
    "</a:rPr>"
  );
}

/** `a:t`, with a newline inside a run becoming the break it was. */
function runElement(run: Run, style: Style, linkId?: string): string {
  return run.text
    .split("\n")
    .map((line) => `<a:r>${runProperties(run, style, linkId)}<a:t>${escapeXml(line)}</a:t></a:r>`)
    .join(`<a:br>${runProperties(run, style, linkId)}</a:br>`);
}

class Renderer {
  /** Hyperlink targets for the slide being written, in the order first seen. */
  private links: string[] = [];
  /** Shape ids are unique within a slide's tree; 1 is the tree itself. */
  private nextId = 2;

  /** Start a slide, discarding the previous one's links and ids. */
  private reset(): void {
    this.links = [];
    this.nextId = 2;
  }

  /** rId1 is the layout, so hyperlinks start above it. */
  private relationshipFor(href: string): string {
    let index = this.links.indexOf(href);
    if (index === -1) {
      index = this.links.push(href) - 1;
    }
    return `rId${index + 2}`;
  }

  private paragraph(runs: readonly Run[], style: Style): string {
    const marginLeft = style.indent * INDENT_EMU;
    const algn = style.align === "right" ? "r" : style.align === "center" ? "ctr" : "l";
    const properties =
      `<a:pPr marL="${marginLeft}" indent="0" algn="${algn}">` +
      (style.before ? `<a:spcBef><a:spcPts val="${style.before}"/></a:spcBef>` : "") +
      // Every list marker here is literal text, as in the other three renderers,
      // so PowerPoint's own bullet has to be turned off or every line gets two.
      "<a:buNone/></a:pPr>";
    const body =
      runs.length === 0
        ? `<a:endParaRPr lang="en-US" sz="${style.size}"/>`
        : runs
            .map((run) =>
              runElement(run, style, run.href ? this.relationshipFor(run.href) : undefined),
            )
            .join("");
    return `<a:p>${properties}${body}</a:p>`;
  }

  /** A text box, positioned absolutely. `placeholder` fills the layout's slot. */
  private textShape(
    name: string,
    box: { x: number; y: number; width: number; height: number },
    paragraphs: string,
    placeholder?: string,
  ): string {
    const id = this.nextId;
    this.nextId += 1;
    return (
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/>` +
      `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
      `<p:nvPr>${placeholder ?? ""}</p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${box.x}" y="${box.y}"/>` +
      `<a:ext cx="${box.width}" cy="${box.height}"/></a:xfrm>` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      // No insets, so the box the text is laid out in is the box measured above.
      // `normAutofit` is the safety net for an estimate that came out short.
      '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0">' +
      "<a:normAutofit/></a:bodyPr><a:lstStyle/>" +
      paragraphs +
      "</p:txBody></p:sp>"
    );
  }

  private table(piece: Extract<Piece, { kind: "table" }>, y: number): string {
    const rows = [piece.header, ...piece.rows];
    const columns = Math.max(1, ...rows.map((row) => row.length));
    const widths = columnShares(rows, columns).map((share) => Math.round(share * CONTENT_WIDTH));
    const total = widths.reduce((sum, width) => sum + width, 0);
    const id = this.nextId;
    this.nextId += 1;

    // Horizontal rules only. A full grid boxes every number in, and the eye
    // reads a table by its rows — the column gaps do what the vertical lines
    // would have.
    const border = (edge: string): string =>
      `<a:${edge} w="12700" cap="flat" cmpd="sng" algn="ctr">` +
      `<a:solidFill><a:srgbClr val="${BORDER}"/></a:solidFill>` +
      `<a:prstDash val="solid"/></a:${edge}>`;
    const noBorder = (edge: string): string => `<a:${edge}><a:noFill/></a:${edge}>`;
    const cell = (cells: readonly Run[][], column: number, row: number): string => {
      const header = row === 0;
      const runs = cells[column] ?? [];
      // Zebra counted from the header, so the first data row is the plain one.
      const fill = header ? HEADER_FILL : row % 2 === 0 ? PALETTE.brandTint : undefined;
      // A cell with no paragraph in it is what makes PowerPoint call the file
      // corrupt — the same trap `w:tc` has in DOCX.
      return (
        "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>" +
        this.paragraph(runs, {
          ...BODY_STYLE,
          align: piece.align[column],
          ...(header ? { bold: true, colour: PALETTE.onBrand } : {}),
        }) +
        "</a:txBody>" +
        '<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720" anchor="ctr">' +
        noBorder("lnL") +
        noBorder("lnR") +
        border("lnT") +
        border("lnB") +
        (fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : "") +
        "</a:tcPr></a:tc>"
      );
    };

    const body = rows
      .map(
        (cells, index) =>
          `<a:tr h="${ROW_HEIGHT}">` +
          Array.from({ length: columns }, (_, column) => cell(cells, column, index)).join("") +
          "</a:tr>",
      )
      .join("");

    return (
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/>` +
      '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
      "<p:nvPr/></p:nvGraphicFramePr>" +
      `<p:xfrm><a:off x="${SIDE_MARGIN}" y="${y}"/>` +
      `<a:ext cx="${total}" cy="${ROW_HEIGHT * rows.length}"/></p:xfrm>` +
      // The literal URI, not `${A}/table`: the table's graphicData URI is
      // `.../drawingml/2006/table` — the namespace with `main` swapped out, not
      // appended to. PowerPoint identifies the embedded object by this string
      // alone, and an unknown URI is not "a table it renders differently", it is
      // damage: the repair flow throws away the whole slide's content. No schema
      // catches it, because to a schema a URI is just a token.
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>' +
      // The style id resolves against `tableStyles.xml`; every visible border
      // is still stated per cell, so the style only supplies what native files
      // always carry.
      `<a:tblPr firstRow="1"><a:tableStyleId>${TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>` +
      `<a:tblGrid>${widths.map((width) => `<a:gridCol w="${width}"/>`).join("")}</a:tblGrid>` +
      body +
      "</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"
    );
  }

  /**
   * The slide number, bottom right.
   *
   * A `slidenum` field rather than a digit, so a deck that gets a slide inserted
   * renumbers itself. It carries **no `a:t`** — the element is optional, and the
   * literal it would hold is what every text extractor picks up: this server's
   * own reader would then return "7" as a line of the slide's content. The empty
   * paragraph it leaves behind lands at the end of the slide, where `normalize`
   * drops it.
   */
  private slideNumber(ordinal: number): string {
    const id = this.nextId;
    this.nextId += 1;
    const box = {
      x: SLIDE_WIDTH - SIDE_MARGIN - NUMBER_BOX.width,
      y: SLIDE_HEIGHT - SIDE_MARGIN,
      width: NUMBER_BOX.width,
      height: NUMBER_BOX.height,
    };
    return (
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Slide Number"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${box.x}" y="${box.y}"/>` +
      `<a:ext cx="${box.width}" cy="${box.height}"/></a:xfrm>` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      '<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/>' +
      '<a:p><a:pPr algn="r"/>' +
      `<a:fld id="${NUMBER_FIELD_ID}" type="slidenum">` +
      `<a:rPr lang="en-US" sz="${centiPoints(DECK.caption)}">` +
      `<a:solidFill><a:srgbClr val="${MUTED}"/></a:solidFill></a:rPr>` +
      // The cached text every native fld carries. PowerPoint recomputes it on
      // open; the reader skips fld contents, so it never reaches extraction.
      `<a:t>${ordinal}</a:t>` +
      "</a:fld></a:p></p:txBody></p:sp>"
    );
  }

  /** The short brand rule under a slide's title. A shape, because a slide has no borders. */
  private accentBar(y: number): string {
    const id = this.nextId;
    this.nextId += 1;
    return (
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Accent ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${SIDE_MARGIN}" y="${y}"/>` +
      `<a:ext cx="${TITLE_RULE.width}" cy="${TITLE_RULE.height}"/></a:xfrm>` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      `<a:solidFill><a:srgbClr val="${PALETTE.brandLight}"/></a:solidFill>` +
      "</p:spPr>" +
      // The empty body is not optional in practice. `p:txBody` is `minOccurs="0"`
      // in the schema, and PowerPoint still calls a `p:sp` without one damaged
      // and offers to repair the file. The empty `a:p` it carries is what every
      // text extractor reads as a blank line, which is why this shape and the
      // slide number are written last — `read/pptx.ts` drops trailing blanks.
      "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody>" +
      "</p:sp>"
    );
  }

  /** One slide, and the hyperlink targets it turned out to need. */
  slide(slide: Slide, index: number): { xml: string; links: readonly string[] } {
    this.reset();
    const titleBox = slide.cover ? COVER_TITLE_BOX : TITLE_BOX;
    const bodyBox = slide.cover ? COVER_BODY_BOX : BODY_BOX;
    const shapes: string[] = [];
    /**
     * Shapes that carry no content, appended after everything else.
     *
     * Their position on the slide is set by `a:xfrm` rather than by their place
     * in the tree, so writing them last costs nothing visually — and it is what
     * keeps their empty paragraphs out of the extracted text, where they would
     * otherwise land as blank lines in the middle of a slide.
     */
    const decorations: string[] = [];

    if (slide.title) {
      shapes.push(
        this.textShape(
          `Title ${index + 1}`,
          { x: SIDE_MARGIN, y: titleBox.y, width: CONTENT_WIDTH, height: titleBox.height },
          this.paragraph(slide.title, {
            size: slide.cover ? COVER_TITLE_SIZE : TITLE_SIZE,
            bold: true,
            colour: PALETTE.brand,
            indent: 0,
          }),
          '<p:ph type="title"/>',
        ),
      );
      if (!slide.cover) {
        // The cover has a tinted ground of its own and does not need the rule;
        // on every other slide this is the whole of the identity.
        decorations.push(this.accentBar(titleBox.y + titleBox.height + TITLE_RULE.gap));
      }
    }

    // Runs of text sit in one box; a table is a frame of its own, so the body is
    // stacked rather than being a single shape.
    let y = bodyBox.y;
    let paragraphs: string[] = [];
    let lines = 0;
    /** The first text box claims the layout's body placeholder; later ones cannot. */
    let placeholderUsed = false;

    const flushText = (): void => {
      if (paragraphs.length === 0) {
        return;
      }
      const height = Math.max(LINE_HEIGHT, lines * LINE_HEIGHT);
      shapes.push(
        this.textShape(
          `Body ${shapes.length + 1}`,
          { x: SIDE_MARGIN, y, width: CONTENT_WIDTH, height },
          paragraphs.join(""),
          placeholderUsed ? undefined : '<p:ph type="body" idx="1"/>',
        ),
      );
      placeholderUsed = true;
      y += height;
      paragraphs = [];
      lines = 0;
    };

    for (const piece of slide.pieces) {
      if (piece.kind === "table") {
        flushText();
        shapes.push(this.table(piece, y));
        y += linesOf(piece) * LINE_HEIGHT;
        continue;
      }
      paragraphs.push(this.paragraph(piece.runs, piece.style));
      lines += linesOf(piece);
    }
    flushText();

    // Not on the cover: a title page with a "1" on it is a title page nobody
    // designed.
    if (!slide.cover) {
      decorations.push(this.slideNumber(index + 1));
    }

    if (shapes.length === 0) {
      // A slide with no shape at all opens, but there is nothing to select and
      // nothing in the outline. An empty body box is the honest empty slide.
      shapes.push(
        this.textShape(
          "Body 1",
          { x: SIDE_MARGIN, y: bodyBox.y, width: CONTENT_WIDTH, height: LINE_HEIGHT },
          this.paragraph([], BODY_STYLE),
          '<p:ph type="body" idx="1"/>',
        ),
      );
    }
    shapes.push(...decorations);

    return {
      xml:
        `<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld>` +
        // The console's lavender, kept for the one slide that can carry a
        // full-bleed field without costing anything: the cover.
        (slide.cover
          ? `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${PALETTE.surfaceTint}"/></a:solidFill>` +
            "<a:effectLst/></p:bgPr></p:bg>"
          : "") +
        "<p:spTree>" +
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
        shapes.join("") +
        "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>",
      links: this.links,
    };
  }
}

/* --------------------------------------------------------------- packaging */

const SLIDE_LAYOUT_TYPE = `${R}/slideLayout`;
const SLIDE_MASTER_TYPE = `${R}/slideMaster`;
const SLIDE_TYPE = `${R}/slide`;
const THEME_TYPE = `${R}/theme`;
const HYPERLINK_TYPE = `${R}/hyperlink`;

function relationships(entries: readonly string[]): string {
  return `<Relationships xmlns="${RELATIONSHIPS}">${entries.join("")}</Relationships>`;
}

function relationship(id: string, type: string, target: string, external = false): string {
  return (
    `<Relationship Id="${id}" Type="${type}" Target="${escapeXml(target)}"` +
    `${external ? ' TargetMode="External"' : ""}/>`
  );
}

function contentTypesXml(slides: number): string {
  const override = (path: string, type: string): string =>
    `<Override PartName="${path}" ContentType="${type}"/>`;
  const presentationml = "application/vnd.openxmlformats-officedocument.presentationml";
  return (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    override("/ppt/presentation.xml", `${presentationml}.presentation.main+xml`) +
    override("/ppt/slideMasters/slideMaster1.xml", `${presentationml}.slideMaster+xml`) +
    override("/ppt/slideLayouts/slideLayout1.xml", `${presentationml}.slideLayout+xml`) +
    override("/ppt/slideLayouts/slideLayout2.xml", `${presentationml}.slideLayout+xml`) +
    Array.from({ length: slides }, (_, index) =>
      override(`/ppt/slides/slide${index + 1}.xml`, `${presentationml}.slide+xml`),
    ).join("") +
    override(
      "/ppt/theme/theme1.xml",
      "application/vnd.openxmlformats-officedocument.theme+xml",
    ) +
    override("/ppt/presProps.xml", `${presentationml}.presProps+xml`) +
    override("/ppt/viewProps.xml", `${presentationml}.viewProps+xml`) +
    override("/ppt/tableStyles.xml", `${presentationml}.tableStyles+xml`) +
    override(
      "/docProps/core.xml",
      "application/vnd.openxmlformats-package.core-properties+xml",
    ) +
    override(
      "/docProps/app.xml",
      "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    ) +
    "</Types>"
  );
}

/**
 * `docProps/app.xml`, whose only job here is to say what wrote the file.
 *
 * OOXML keeps the producer separate from the core properties: `dc:title` and the
 * dates are the document's, `<Application>` is the tool's. Word fills in a dozen
 * more fields — word counts, template names — and none of them are things this
 * renderer knows or a reader needs. When somebody turns up with a file that
 * renders oddly, this is the line that says which release made it.
 */
function appPropertiesXml(): string {
  return (
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    `<Application>${escapeXml(PRODUCER)}</Application>` +
    "</Properties>"
  );
}

function packageRelsXml(): string {
  return relationships([
    relationship("rId1", `${R}/officeDocument`, "ppt/presentation.xml"),
    relationship("rId2", `${RELATIONSHIPS}/metadata/core-properties`, "docProps/core.xml"),
    relationship("rId3", `${R}/extended-properties`, "docProps/app.xml"),
  ]);
}

function corePropertiesXml(title: string, created: string): string {
  return (
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${escapeXml(title)}</dc:title>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${created}</dcterms:modified>` +
    "</cp:coreProperties>"
  );
}

/**
 * The parts PowerPoint writes into every file and reads back from every file.
 *
 * All three are effectively empty — no styles defined, no view state worth
 * keeping — but their absence is a package shape no native file has ever had,
 * and Windows PowerPoint reads absence as damage where the Mac build shrugs.
 */
function presPropsXml(): string {
  return `<p:presentationPr xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"/>`;
}

function viewPropsXml(): string {
  return `<p:viewPr xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"/>`;
}

function tableStylesXml(): string {
  return `<a:tblStyleLst xmlns:a="${A}" def="${TABLE_STYLE_ID}"/>`;
}

function presentationXml(slides: number): string {
  return (
    `<p:presentation xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" saveSubsetFonts="1">` +
    '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
    "<p:sldIdLst>" +
    // Slide ids start at 256, which the format requires, and the relationship
    // ids continue past the master's rId1.
    Array.from(
      { length: slides },
      (_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`,
    ).join("") +
    "</p:sldIdLst>" +
    `<p:sldSz cx="${SLIDE_WIDTH}" cy="${SLIDE_HEIGHT}"/>` +
    '<p:notesSz cx="6858000" cy="9144000"/>' +
    "</p:presentation>"
  );
}

function presentationRelsXml(slides: number): string {
  // The theme is reachable through the master, but native files relate it from
  // the presentation part as well, and the other three live only here.
  const next = slides + 2;
  return relationships([
    relationship("rId1", SLIDE_MASTER_TYPE, "slideMasters/slideMaster1.xml"),
    ...Array.from({ length: slides }, (_, index) =>
      relationship(`rId${index + 2}`, SLIDE_TYPE, `slides/slide${index + 1}.xml`),
    ),
    relationship(`rId${next}`, THEME_TYPE, "theme/theme1.xml"),
    relationship(`rId${next + 1}`, `${R}/presProps`, "presProps.xml"),
    relationship(`rId${next + 2}`, `${R}/viewProps`, "viewProps.xml"),
    relationship(`rId${next + 3}`, `${R}/tableStyles`, "tableStyles.xml"),
  ]);
}

/**
 * The theme, which is not optional.
 *
 * PowerPoint refuses a package whose master has no theme behind it, and the
 * schema wants three entries in each of the four format lists whether or not
 * anything refers to them. No east-Asian face is named, for the reason
 * `write/docx.ts` names none: the substitute PowerPoint picks on the reader's
 * machine beats one picked here from a font that may not be installed.
 */
function themeXml(): string {
  const accents = CHART.slice(0, 6);
  const line = (width: number): string =>
    `<a:ln w="${width}" cap="flat" cmpd="sng" algn="ctr">` +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>';
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  return (
    `<a:theme xmlns:a="${A}" name="Office">` +
    "<a:themeElements>" +
    '<a:clrScheme name="Office">' +
    `<a:dk1><a:sysClr val="windowText" lastClr="${PALETTE.ink}"/></a:dk1>` +
    `<a:lt1><a:sysClr val="window" lastClr="${PALETTE.onBrand}"/></a:lt1>` +
    `<a:dk2><a:srgbClr val="${PALETTE.brand}"/></a:dk2>` +
    `<a:lt2><a:srgbClr val="${PALETTE.brandTint}"/></a:lt2>` +
    accents
      .map((colour, index) => `<a:accent${index + 1}><a:srgbClr val="${colour}"/></a:accent${index + 1}>`)
      .join("") +
    `<a:hlink><a:srgbClr val="${LINK_COLOUR}"/></a:hlink>` +
    `<a:folHlink><a:srgbClr val="${PALETTE.brandDeep}"/></a:folHlink>` +
    "</a:clrScheme>" +
    '<a:fontScheme name="Office">' +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    "</a:fontScheme>" +
    '<a:fmtScheme name="Office">' +
    `<a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst>` +
    `<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
    "<a:effectStyleLst>" +
    "<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3) +
    "</a:effectStyleLst>" +
    `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst>` +
    "</a:fmtScheme>" +
    "</a:themeElements>" +
    "</a:theme>"
  );
}

/** An empty placeholder in a layout, which the slide's own shape inherits its slot from. */
function layoutPlaceholder(id: number, name: string, placeholder: string, box: { y: number; height: number }): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/>` +
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    `<p:nvPr>${placeholder}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${SIDE_MARGIN}" y="${box.y}"/>` +
    `<a:ext cx="${CONTENT_WIDTH}" cy="${box.height}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/>' +
    '<a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
  );
}

function slideLayoutXml(cover: boolean): string {
  const title = cover ? COVER_TITLE_BOX : TITLE_BOX;
  const body = cover ? COVER_BODY_BOX : BODY_BOX;
  return (
    `<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" ` +
    `type="${cover ? "title" : "obj"}" preserve="1">` +
    `<p:cSld name="${cover ? "Title Slide" : "Title and Content"}"><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    layoutPlaceholder(2, "Title 1", '<p:ph type="title"/>', title) +
    layoutPlaceholder(3, "Body 2", '<p:ph type="body" idx="1"/>', body) +
    "</p:spTree></p:cSld>" +
    "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"
  );
}

function slideLayoutRelsXml(): string {
  return relationships([
    relationship("rId1", SLIDE_MASTER_TYPE, "../slideMasters/slideMaster1.xml"),
  ]);
}

function slideMasterXml(): string {
  const style = (size: number): string =>
    `<a:lvl1pPr><a:defRPr sz="${size}"><a:solidFill><a:srgbClr val="${INK}"/></a:solidFill>` +
    "</a:defRPr></a:lvl1pPr>";
  return (
    `<p:sldMaster xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">` +
    '<p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill>' +
    "<a:effectLst/></p:bgPr></p:bg><p:spTree>" +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    // The top of the inheritance chain. A slide's placeholder points at its
    // layout's, and the layout's at these — native masters always carry them,
    // and a chain that ends nowhere is another shape no real file has.
    layoutPlaceholder(2, "Title Placeholder 1", '<p:ph type="title"/>', TITLE_BOX) +
    layoutPlaceholder(3, "Body Placeholder 2", '<p:ph type="body" idx="1"/>', BODY_BOX) +
    "</p:spTree></p:cSld>" +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ' +
    'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" ' +
    'folHlink="folHlink"/>' +
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/>' +
    '<p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>' +
    `<p:txStyles><p:titleStyle>${style(TITLE_SIZE)}</p:titleStyle>` +
    `<p:bodyStyle>${style(BODY_SIZE)}</p:bodyStyle>` +
    `<p:otherStyle>${style(BODY_SIZE)}</p:otherStyle></p:txStyles>` +
    "</p:sldMaster>"
  );
}

function slideMasterRelsXml(): string {
  return relationships([
    relationship("rId1", SLIDE_LAYOUT_TYPE, "../slideLayouts/slideLayout1.xml"),
    relationship("rId2", SLIDE_LAYOUT_TYPE, "../slideLayouts/slideLayout2.xml"),
    relationship("rId3", THEME_TYPE, "../theme/theme1.xml"),
  ]);
}

function slideRelsXml(cover: boolean, links: readonly string[]): string {
  return relationships([
    relationship(
      "rId1",
      SLIDE_LAYOUT_TYPE,
      `../slideLayouts/slideLayout${cover ? 1 : 2}.xml`,
    ),
    ...links.map((href, index) =>
      relationship(`rId${index + 2}`, HYPERLINK_TYPE, href, true),
    ),
  ]);
}

export interface PptxOptions {
  title: string;
  /** ISO 8601, passed in so the bytes are a function of the input alone. */
  created: string;
}

export interface RenderedPptx {
  bytes: Uint8Array;
  /** How many slides the Markdown turned into, which is what the caller reports. */
  slides: number;
}

export function renderPptx(document: MarkdownDocument, options: PptxOptions): RenderedPptx {
  const slides = slidesOf(document);
  const renderer = new Renderer();
  const parts: Record<string, Uint8Array> = {
    "[Content_Types].xml": part(contentTypesXml(slides.length)),
    "_rels/.rels": part(packageRelsXml()),
    "docProps/core.xml": part(corePropertiesXml(options.title, options.created)),
    "docProps/app.xml": part(appPropertiesXml()),
    "ppt/presentation.xml": part(presentationXml(slides.length)),
    "ppt/_rels/presentation.xml.rels": part(presentationRelsXml(slides.length)),
    "ppt/theme/theme1.xml": part(themeXml()),
    "ppt/presProps.xml": part(presPropsXml()),
    "ppt/viewProps.xml": part(viewPropsXml()),
    "ppt/tableStyles.xml": part(tableStylesXml()),
    "ppt/slideMasters/slideMaster1.xml": part(slideMasterXml()),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": part(slideMasterRelsXml()),
    "ppt/slideLayouts/slideLayout1.xml": part(slideLayoutXml(true)),
    "ppt/slideLayouts/slideLayout2.xml": part(slideLayoutXml(false)),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": part(slideLayoutRelsXml()),
    "ppt/slideLayouts/_rels/slideLayout2.xml.rels": part(slideLayoutRelsXml()),
  };

  slides.forEach((slide, index) => {
    const rendered = renderer.slide(slide, index);
    parts[`ppt/slides/slide${index + 1}.xml`] = part(rendered.xml);
    parts[`ppt/slides/_rels/slide${index + 1}.xml.rels`] = part(
      slideRelsXml(slide.cover, rendered.links),
    );
  });

  return { bytes: buildZip(parts), slides: slides.length };
}
