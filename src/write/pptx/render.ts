/**
 * A planned slide to its XML part.
 *
 * Everything here is shape arithmetic: the planner has already decided what is
 * on the slide, and this module decides what DrawingML says it. Runs of text
 * sit in one box; a table is a frame of its own, so a body is stacked rather
 * than being a single shape.
 */

import { escapeXml } from "../../xml.js";
import type { Run } from "../../markdown.js";
import { columnShares } from "../table.js";
import { DECK, PALETTE, centiPoints } from "../theme.js";
import {
  BODY_BOX,
  BODY_SIZE,
  CONTENT_WIDTH,
  COVER_BODY_BOX,
  COVER_TITLE_BOX,
  COVER_TITLE_SIZE,
  INDENT_EMU,
  LINE_HEIGHT,
  NUMBER_BOX,
  ROW_HEIGHT,
  SIDE_MARGIN,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  TITLE_BOX,
  TITLE_RULE,
  TITLE_SIZE,
  linesOf,
} from "./layout.js";
import { A, NUMBER_FIELD_ID, P, R, TABLE_STYLE_ID } from "./ooxml.js";
import type { Piece, Slide, Style } from "./types.js";

const INK = PALETTE.ink;
const LINK_COLOUR = PALETTE.brandDeep;
const MUTED = PALETTE.inkMuted;
const BORDER = PALETTE.rule;
const HEADER_FILL = PALETTE.brand;

const BODY_STYLE: Style = { size: BODY_SIZE, indent: 0 };

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

export class Renderer {
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
    const cover = slide.type === "cover";
    const titleBox = cover ? COVER_TITLE_BOX : TITLE_BOX;
    const bodyBox = cover ? COVER_BODY_BOX : BODY_BOX;
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
            size: cover ? COVER_TITLE_SIZE : TITLE_SIZE,
            bold: true,
            colour: PALETTE.brand,
            indent: 0,
          }),
          '<p:ph type="title"/>',
        ),
      );
      if (!cover) {
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
    if (!cover) {
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
        (cover
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
