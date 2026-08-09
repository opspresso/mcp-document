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

/** A4 in points, and a 2cm margin. */
const PAGE_WIDTH = 595.28;
const PAGE_HEIGHT = 841.89;
const MARGIN = 56.7;
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2;

const BODY_SIZE = 11;
const CODE_SIZE = 9.5;
const HEADING_SIZES = [20, 17, 15, 13, 12, 11];
const LINE_RATIO = 1.5;
/** Space above a heading, as a multiple of its own size. */
const HEADING_SPACE_ABOVE = 0.8;
const PARAGRAPH_SPACE = 6;
const INDENT_STEP = 18;

const INK = rgb(0.13, 0.13, 0.13);
const LINK_COLOUR = rgb(0.02, 0.39, 0.76);
const MUTED = rgb(0.4, 0.4, 0.4);
const RULE_COLOUR = rgb(0.75, 0.75, 0.75);
const CODE_BACKGROUND = rgb(0.96, 0.96, 0.96);

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

function colourFor(run: Run): RGB {
  return run.href ? LINK_COLOUR : INK;
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

  private drawLine(line: Line, left: number, size: number, baseline: number): void {
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
          color: colourFor(atom.run),
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
    options: { size: number; left: number; width: number; firstLinePrefix?: Atom[] },
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
      this.drawLine(line, left, size, this.y + height * 0.25);
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
        color: CODE_BACKGROUND,
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
      let x = MARGIN;
      laid.forEach((lines, column) => {
        this.page.drawRectangle({
          x,
          y: this.y,
          width: widths[column]!,
          height,
          borderColor: RULE_COLOUR,
          borderWidth: 0.5,
          ...(header ? { color: rgb(0.95, 0.95, 0.95) } : {}),
        });
        lines.forEach((line, lineIndex) => {
          this.drawLine(
            line,
            x + 5,
            BODY_SIZE,
            this.y + height - 4 - (lineIndex + 1) * BODY_SIZE * 1.35 + BODY_SIZE * 0.35,
          );
        });
        x += widths[column]!;
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
          { size, left: MARGIN, width: CONTENT_WIDTH },
        );
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
          { size: BODY_SIZE, left: MARGIN + INDENT_STEP, width: CONTENT_WIDTH - INDENT_STEP },
        );
        // Only when the quote stayed on one page: a bar drawn from a `top` that
        // belongs to the previous page runs the length of this one.
        if (this.y < top) {
          this.page.drawRectangle({
            x: MARGIN,
            y: this.y,
            width: 2,
            height: top - this.y,
            color: MUTED,
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
  pdf.setProducer("mcp-document");
  pdf.setCreationDate(options.created);
  pdf.setModificationDate(options.created);

  const writer = new Writer(pdf, fonts);
  for (const block of document.blocks) {
    writer.block(block);
  }
  return { bytes: await pdf.save(), pages: writer.pageCount() };
}
