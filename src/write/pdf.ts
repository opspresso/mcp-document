/**
 * A document AST to PDF.
 *
 * PDF has no idea what a paragraph is, so unlike the other two renderers this
 * one is a layout engine: it measures, breaks lines, and decides where a page
 * ends. Everything here follows from that.
 *
 * **The Korean font is embedded whole, and that is not optional.** PDF's
 * built-in fonts cover Latin-1 and nothing else, so a document with a single
 * Hangul syllable in it needs a real font in the file. Nanum Gothic ships in
 * this repository (`assets/fonts`, SIL OFL). Noto Sans KR would be the more
 * obvious choice and is not here for a mechanical reason: Google Fonts now
 * publishes it as a variable font, and putting one of those through fontkit is
 * a path with more ways to go wrong than a static TTF has.
 *
 * Whole, rather than subset, because `@pdf-lib/fontkit`'s subsetter **silently
 * drops most Hangul glyphs**. It does not fail: the text layer is intact, so
 * extraction returns the document perfectly, and the page shows blanks where
 * two thirds of the characters should be — "2026년 1분기 보고서" renders as
 * "6년 서". A document that reads correctly to a machine and is unreadable to a
 * person is the worst shape this could take, so the ~750KB a Flate-compressed
 * face costs is paid on every PDF. The bold face is embedded only when
 * something is bold, which is what keeps a plain document to one of them.
 *
 * **Line breaking is per script.** Latin breaks at spaces; CJK breaks between
 * any two characters, because Korean and Chinese prose has no spaces to break
 * at and a line breaker that waits for one produces a single line running off
 * the page. So the text is split into atoms — a word, or one CJK character —
 * and lines are filled greedily.
 *
 * **Italic is synthetic.** Nanum Gothic has no italic face, and shipping a
 * third file to slant some text is not a trade worth making; the text is
 * sheared instead. Bold is a real face, because a faked one is visibly wrong.
 */

import { readFileSync } from "node:fs";
import fontkit from "@pdf-lib/fontkit";
import {
  degrees,
  PDFDocument,
  PDFFont,
  PDFName,
  PDFPage,
  PDFString,
  rgb,
  StandardFonts,
  type RGB,
} from "pdf-lib";
import type { Block, MarkdownDocument, Run } from "../markdown.js";
import { columnShares } from "./table.js";
import { DOC, rgbOf, type ColourName } from "./theme.js";
import { PRODUCER } from "../version.js";

/** A4 in points, and a 2cm margin. */
const PAGE_WIDTH = 595.28;
const PAGE_HEIGHT = 841.89;
const MARGIN = 56.7;
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2;

const BODY_SIZE = DOC.body;
const CODE_SIZE = DOC.code;
const CAPTION_SIZE = DOC.caption;
const HEADING_SIZES = DOC.headings;
const LINE_RATIO = 1.5;
/** Space above a heading, as a multiple of its own size. */
const HEADING_SPACE_ABOVE = 0.8;
const PARAGRAPH_SPACE = 6;
const INDENT_STEP = 18;

/** The shared palette, in the three floats `pdf-lib` wants. */
function ink(name: ColourName): RGB {
  const { r, g, b } = rgbOf(name);
  return rgb(r, g, b);
}

const INK = ink("ink");
const LINK_COLOUR = ink("brandDeep");
const MUTED = ink("inkMuted");
const RULE_COLOUR = ink("rule");
const BRAND = ink("brand");
const BRAND_LIGHT = ink("brandLight");
const ON_BRAND = ink("onBrand");
const TINT = ink("brandTint");

/**
 * Scripts that break between any two characters.
 *
 * Hangul, Han, Kana, and the CJK punctuation and full-width forms that travel
 * with them. A Latin word inside Korean prose is still one atom, which is what
 * keeps `mcp-document` from being broken across two lines.
 */
const CJK =
  /[ᄀ-ᇿ⺀-〿぀-ヿ㄰-㆏㐀-䶿一-鿿ꥠ-꥿가-퟿豈-﫿︰-﹏＀-｠￠-￦]/;

/** Characters the built-in Courier can encode, which is what makes a listing monospaced. */
const LATIN1_ONLY = /^[\x20-\x7e]*$/;

const FONT_DIRECTORY = new URL("../../assets/fonts/", import.meta.url);

/** Read once per process: two files of about 2MB each. */
let fontBytes: { regular: Uint8Array; bold: Uint8Array } | undefined;

function loadFontBytes(): { regular: Uint8Array; bold: Uint8Array } {
  fontBytes ??= {
    regular: readFileSync(new URL("NanumGothic-Regular.ttf", FONT_DIRECTORY)),
    bold: readFileSync(new URL("NanumGothic-Bold.ttf", FONT_DIRECTORY)),
  };
  return fontBytes;
}

interface Fonts {
  regular: PDFFont;
  bold: PDFFont;
  mono: PDFFont;
}

/** One drawable piece: a word, a run of spaces, or a single CJK character. */
interface Atom {
  text: string;
  run: Run;
  width: number;
  /** A run of spaces, which is dropped when it lands at the end of a line. */
  space: boolean;
}

interface Line {
  atoms: Atom[];
  width: number;
}

function fontFor(fonts: Fonts, run: Run): PDFFont {
  return run.bold ? fonts.bold : fonts.regular;
}

/**
 * `override` exists for text on a brand-filled surface, which has to be white
 * whatever the run says. A link keeps its own colour even there — losing the
 * only signal that it is clickable would be the worse trade.
 */
function colourFor(run: Run, override?: RGB): RGB {
  if (run.href) {
    return override ?? LINK_COLOUR;
  }
  return override ?? INK;
}

/** Split a run into the smallest pieces a line may break between. */
function atomsOf(run: Run, fonts: Fonts, size: number): Atom[] {
  const font = fontFor(fonts, run);
  const atoms: Atom[] = [];
  const push = (text: string, space: boolean): void => {
    if (text !== "") {
      atoms.push({ text, run, width: font.widthOfTextAtSize(text, size), space });
    }
  };
  let word = "";
  for (const character of run.text.replace(/\n/g, " ")) {
    if (/\s/.test(character)) {
      push(word, false);
      word = "";
      push(character, true);
      continue;
    }
    if (CJK.test(character)) {
      push(word, false);
      word = "";
      push(character, false);
      continue;
    }
    word += character;
  }
  push(word, false);
  return atoms;
}

/** Greedy fill. A single atom wider than the line gets a line of its own rather than a loop. */
function wrap(atoms: readonly Atom[], maxWidth: number): Line[] {
  const lines: Line[] = [];
  let current: Atom[] = [];
  let width = 0;
  for (const atom of atoms) {
    if (current.length === 0 && atom.space) {
      // A break already happened here; the space that caused it is not indent.
      continue;
    }
    if (width + atom.width > maxWidth && current.length > 0) {
      while (current.length > 0 && current[current.length - 1]!.space) {
        width -= current.pop()!.width;
      }
      lines.push({ atoms: current, width });
      current = atom.space ? [] : [atom];
      width = atom.space ? 0 : atom.width;
      continue;
    }
    current.push(atom);
    width += atom.width;
  }
  if (current.length > 0) {
    lines.push({ atoms: current, width });
  }
  return lines;
}

class Writer {
  private page: PDFPage;
  private y: number;
  private pages = 1;

  constructor(
    private readonly document: PDFDocument,
    private readonly fonts: Fonts,
  ) {
    this.page = document.addPage([PAGE_WIDTH, PAGE_HEIGHT]);
    this.y = PAGE_HEIGHT - MARGIN;
  }

  pageCount(): number {
    return this.pages;
  }

  /**
   * Page numbers, drawn once the count is known.
   *
   * After the body rather than during it, because a footer written while the
   * page is being filled would be one the layout then has to avoid — and the
   * number of a page is not knowable until the page exists. A single-page
   * document gets none: "1" under a one-page memo is furniture.
   */
  numberPages(): void {
    if (this.pages < 2) {
      return;
    }
    for (const [index, page] of this.document.getPages().entries()) {
      const label = `${index + 1}`;
      page.drawText(label, {
        x: PAGE_WIDTH - MARGIN - this.fonts.regular.widthOfTextAtSize(label, CAPTION_SIZE),
        // Inside the bottom margin, which nothing else is allowed to enter.
        y: MARGIN * 0.55,
        size: CAPTION_SIZE,
        font: this.fonts.regular,
        color: MUTED,
      });
    }
  }

  /** Make room for `height`, starting a page when there is none. */
  private reserve(height: number): void {
    if (this.y - height < MARGIN) {
      this.page = this.document.addPage([PAGE_WIDTH, PAGE_HEIGHT]);
      this.pages += 1;
      this.y = PAGE_HEIGHT - MARGIN;
    }
    this.y -= height;
  }

  private space(height: number): void {
    // Never at the top of a page: leading whitespace there is a margin nobody
    // asked for, and it accumulates every time a block starts a page.
    if (this.y < PAGE_HEIGHT - MARGIN) {
      this.y -= height;
    }
  }

  /**
   * A clickable region over text already drawn.
   *
   * The annotation is what makes a link a link — colouring text blue only makes
   * it look like one, and a reader who clicks gets nothing.
   */
  private annotate(href: string, x: number, y: number, width: number, height: number): void {
    const annotation = this.document.context.register(
      this.document.context.obj({
        Type: "Annot",
        Subtype: "Link",
        Rect: [x, y, x + width, y + height],
        Border: [0, 0, 0],
        A: { Type: "Action", S: "URI", URI: PDFString.of(href) },
      }),
    );
    const existing = this.page.node.lookup(PDFName.of("Annots"));
    if (existing && "push" in existing && typeof existing.push === "function") {
      (existing as { push: (value: unknown) => void }).push(annotation);
      return;
    }
    this.page.node.set(PDFName.of("Annots"), this.document.context.obj([annotation]));
  }

  private drawLine(
    line: Line,
    left: number,
    size: number,
    baseline: number,
    colour?: RGB,
  ): void {
    let x = left;
    /** Where the current link started, so it is underlined and annotated once. */
    let linkFrom: number | undefined;
    let linkHref: string | undefined;

    const closeLink = (to: number): void => {
      if (linkFrom === undefined || linkHref === undefined) {
        return;
      }
      // One rectangle for the whole link, not one per word: two annotations
      // with the space between them left out is a link with a dead gap in it.
      this.page.drawLine({
        start: { x: linkFrom, y: baseline - size * 0.12 },
        end: { x: to, y: baseline - size * 0.12 },
        thickness: 0.5,
        color: LINK_COLOUR,
      });
      this.annotate(linkHref, linkFrom, baseline - size * 0.2, to - linkFrom, size);
      linkFrom = undefined;
      linkHref = undefined;
    };

    for (const atom of line.atoms) {
      if (atom.run.href !== linkHref) {
        closeLink(x);
        if (atom.run.href) {
          linkFrom = x;
          linkHref = atom.run.href;
        }
      }
      if (!atom.space) {
        this.page.drawText(atom.text, {
          x,
          y: baseline,
          size,
          font: fontFor(this.fonts, atom.run),
          color: colourFor(atom.run, colour),
          // Nanum Gothic has no italic face; shearing the glyphs is what stands
          // in for one.
          ...(atom.run.italic ? { xSkew: degrees(12) } : {}),
        });
      }
      x += atom.width;
    }
    closeLink(x);
  }

  /** Lay out runs into the given width and draw them, returning the height used. */
  private paragraph(
    runs: readonly Run[],
    options: {
      size: number;
      left: number;
      width: number;
      firstLinePrefix?: Atom[];
      colour?: RGB;
    },
  ): void {
    const { size, left, width } = options;
    const atoms = [
      ...(options.firstLinePrefix ?? []),
      ...runs.flatMap((run) => atomsOf(run, this.fonts, size)),
    ];
    const lines = wrap(atoms, width);
    const height = size * LINE_RATIO;
    for (const line of lines.length > 0 ? lines : [{ atoms: [], width: 0 }]) {
      this.reserve(height);
      // The baseline sits above the descender, not on the line's bottom edge.
      this.drawLine(line, left, size, this.y + height * 0.25, options.colour);
    }
  }

  private codeBlock(text: string): void {
    const lines = text === "" ? [""] : text.split("\n");
    // One decision for the whole block: a listing whose lines changed font
    // halfway down reads as a rendering fault rather than as a choice.
    const monospaced = lines.every((line) => LATIN1_ONLY.test(line));
    const font = monospaced ? this.fonts.mono : this.fonts.regular;
    const height = CODE_SIZE * 1.35;
    this.space(PARAGRAPH_SPACE);
    for (const line of lines) {
      this.reserve(height);
      this.page.drawRectangle({
        x: MARGIN,
        y: this.y,
        width: CONTENT_WIDTH,
        height,
        color: TINT,
      });
      // Cut rather than wrapped: a wrapped line of code is a line of code that
      // says something different, and there is no continuation marker in a PDF
      // that would say otherwise.
      let visible = line;
      while (visible !== "" && font.widthOfTextAtSize(visible, CODE_SIZE) > CONTENT_WIDTH - 12) {
        visible = visible.slice(0, -1);
      }
      this.page.drawText(visible, {
        x: MARGIN + 6,
        y: this.y + height * 0.28,
        size: CODE_SIZE,
        font,
        color: INK,
      });
    }
    this.space(PARAGRAPH_SPACE);
  }

  private table(block: Extract<Block, { kind: "table" }>): void {
    const rows = [block.header, ...block.rows];
    const columns = Math.max(1, ...rows.map((row) => row.length));
    const widths = columnWidths(rows, columns);
    this.space(PARAGRAPH_SPACE);
    rows.forEach((cells, index) => {
      const header = index === 0;
      const laid = Array.from({ length: columns }, (_, column) => {
        const runs = (cells[column] ?? []).map((run) => (header ? { ...run, bold: true } : run));
        return wrap(
          runs.flatMap((run) => atomsOf(run, this.fonts, BODY_SIZE)),
          widths[column]! - 10,
        );
      });
      const height = Math.max(1, ...laid.map((lines) => lines.length)) * BODY_SIZE * 1.35 + 6;
      this.reserve(height);
      // One fill per row, not one box per cell. A full grid boxes every number
      // in; the eye reads a table by its rows, and the column gaps are already
      // doing what the vertical lines would.
      const fill = header ? BRAND : index % 2 === 0 ? TINT : undefined;
      if (fill) {
        this.page.drawRectangle({
          x: MARGIN,
          y: this.y,
          width: CONTENT_WIDTH,
          height,
          color: fill,
        });
      }
      this.page.drawLine({
        start: { x: MARGIN, y: this.y },
        end: { x: MARGIN + CONTENT_WIDTH, y: this.y },
        thickness: 0.5,
        color: RULE_COLOUR,
      });
      let x = MARGIN;
      laid.forEach((lines, column) => {
        const width = widths[column]!;
        const align = block.align[column];
        lines.forEach((line, lineIndex) => {
          // Every line of a wrapped cell is placed on its own, so a two-line
          // right-aligned cell has both lines flush to the same edge.
          const left =
            align === "right"
              ? x + width - 5 - line.width
              : align === "center"
                ? x + (width - line.width) / 2
                : x + 5;
          this.drawLine(
            line,
            left,
            BODY_SIZE,
            this.y + height - 4 - (lineIndex + 1) * BODY_SIZE * 1.35 + BODY_SIZE * 0.35,
            header ? ON_BRAND : undefined,
          );
        });
        x += width;
      });
    });
    this.space(PARAGRAPH_SPACE);
  }

  block(block: Block): void {
    switch (block.kind) {
      case "heading": {
        const size = HEADING_SIZES[block.level - 1]!;
        this.space(size * HEADING_SPACE_ABOVE);
        this.paragraph(
          block.runs.map((run) => ({ ...run, bold: true })),
          { size, left: MARGIN, width: CONTENT_WIDTH, colour: BRAND },
        );
        // A hairline under the top two levels, in the brand colour. It is what
        // the console's lavender page became on paper: the same signal at a
        // hundredth of the ink.
        if (block.level <= 2) {
          this.space(3);
          this.reserve(1);
          this.page.drawLine({
            start: { x: MARGIN, y: this.y },
            end: { x: PAGE_WIDTH - MARGIN, y: this.y },
            thickness: 0.75,
            color: BRAND_LIGHT,
          });
        }
        this.space(2);
        return;
      }
      case "paragraph":
        this.paragraph(block.runs, { size: BODY_SIZE, left: MARGIN, width: CONTENT_WIDTH });
        this.space(PARAGRAPH_SPACE);
        return;
      case "list": {
        const counters: number[] = [];
        for (const item of block.items) {
          counters.length = item.depth + 1;
          counters[item.depth] = (counters[item.depth] ?? 0) + 1;
          const marker = block.ordered ? `${counters[item.depth]}.  ` : "•  ";
          const left = MARGIN + INDENT_STEP * item.depth;
          const prefix = atomsOf({ text: marker }, this.fonts, BODY_SIZE);
          this.paragraph(item.runs, {
            size: BODY_SIZE,
            left,
            width: CONTENT_WIDTH - INDENT_STEP * item.depth,
            firstLinePrefix: prefix,
          });
        }
        this.space(PARAGRAPH_SPACE);
        return;
      }
      case "code":
        this.codeBlock(block.text);
        return;
      case "quote": {
        this.space(PARAGRAPH_SPACE / 2);
        const top = this.y;
        this.paragraph(
          block.runs.map((run) => ({ ...run, italic: true })),
          {
            size: BODY_SIZE,
            left: MARGIN + INDENT_STEP,
            width: CONTENT_WIDTH - INDENT_STEP,
            colour: MUTED,
          },
        );
        // Only when the quote stayed on one page: a bar drawn from a `top` that
        // belongs to the previous page runs the length of this one.
        if (this.y < top) {
          this.page.drawRectangle({
            x: MARGIN,
            y: this.y,
            width: 2,
            height: top - this.y,
            color: BRAND_LIGHT,
          });
        }
        this.space(PARAGRAPH_SPACE);
        return;
      }
      case "table":
        this.table(block);
        return;
      case "rule":
        this.space(PARAGRAPH_SPACE);
        this.reserve(1);
        this.page.drawLine({
          start: { x: MARGIN, y: this.y },
          end: { x: PAGE_WIDTH - MARGIN, y: this.y },
          thickness: 0.5,
          color: RULE_COLOUR,
        });
        this.space(PARAGRAPH_SPACE);
        return;
    }
  }
}

/** The shared column shares, in points. */
export function columnWidths(rows: readonly Run[][][], columns: number): number[] {
  return columnShares(rows, columns).map((share) => share * CONTENT_WIDTH);
}

export interface PdfOptions {
  title: string;
  /** Passed in rather than read from the clock, so the bytes follow from the input. */
  created: Date;
}

export interface RenderedPdf {
  bytes: Uint8Array;
  pages: number;
}

/**
 * Whether anything in the document is set in bold.
 *
 * Asked so the bold face can be left out when it is not used, which halves the
 * font weight of a plain document. Headings and table headers are bolded by
 * this renderer rather than by the parser, so they count here even though no
 * run in them says so.
 */
export function usesBold(document: MarkdownDocument): boolean {
  return document.blocks.some(
    (block) =>
      block.kind === "heading" ||
      block.kind === "table" ||
      (block.kind === "paragraph" && block.runs.some((run) => run.bold)) ||
      (block.kind === "quote" && block.runs.some((run) => run.bold)) ||
      (block.kind === "list" && block.items.some((item) => item.runs.some((run) => run.bold))),
  );
}

export async function renderPdf(
  document: MarkdownDocument,
  options: PdfOptions,
): Promise<RenderedPdf> {
  const pdf = await PDFDocument.create();
  pdf.registerFontkit(fontkit);
  const bytes = loadFontBytes();
  // `subset: true` is what one would want and is not usable: it drops most
  // Hangul glyphs without erroring. See the note at the top of this file.
  const regular = await pdf.embedFont(bytes.regular, { subset: false });
  const fonts: Fonts = {
    regular,
    bold: usesBold(document) ? await pdf.embedFont(bytes.bold, { subset: false }) : regular,
    mono: await pdf.embedFont(StandardFonts.Courier),
  };
  pdf.setTitle(options.title);
  // Both fields, because readers disagree about which one they show: Preview
  // reads `Creator`, most others `Producer`.
  pdf.setProducer(PRODUCER);
  pdf.setCreator(PRODUCER);
  pdf.setCreationDate(options.created);
  pdf.setModificationDate(options.created);

  const writer = new Writer(pdf, fonts);
  for (const block of document.blocks) {
    writer.block(block);
  }
  writer.numberPages();
  return { bytes: await pdf.save(), pages: writer.pageCount() };
}
