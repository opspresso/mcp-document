/**
 * A document AST to DOCX.
 *
 * The parts are written by hand rather than through a library, for the reason
 * the rest of this repository is: the surface actually used is small — six
 * parts, one of which is the body — and a library's idea of a paragraph is one
 * more thing between the AST and the bytes.
 *
 * Two decisions are worth naming.
 *
 * **Lists are literal markers with a hanging indent**, not `w:numPr` against a
 * `numbering.xml`. Real numbering is a second part, a definition per level per
 * list, and an abstract-numbering indirection — and what it buys is Word
 * renumbering a list the reader edits. Nothing here is edited before it is
 * read, and the literal form is what survives being extracted back to text,
 * which is how this server's own round trip checks itself.
 *
 * **No font is named.** The theme's default already carries an east-Asian face
 * on every platform Word runs on, and naming one — `Malgun Gothic`, say — is a
 * Windows font that a Mac substitutes for something else. Substitution chosen
 * by Word is better than substitution chosen here.
 */

import { escapeXml } from "../xml.js";
import { buildZip } from "../zip.js";
import type { Align, Block, MarkdownDocument, Run } from "../markdown.js";
import { columnShares } from "./table.js";
import { DOC, PALETTE, halfPoints } from "./theme.js";
import { PRODUCER } from "../version.js";

/** A4, in twentieths of a point, with a 2.5cm margin. */
const PAGE = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/>';

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** One indent level, in twentieths of a point: 0.63cm, Word's own list step. */
const INDENT_STEP = 360;

/** rId1 is the styles part and rId2 the footer; hyperlinks take what is left. */
const FOOTER_RELATIONSHIP = "rId2";
const FIRST_LINK_RELATIONSHIP = 3;

/** What is left of the page once the margins are taken out — a full-width table. */
const TABLE_WIDTH = 11906 - 1418 * 2;

const encoder = new TextEncoder();

function part(xml: string): Uint8Array {
  return encoder.encode(DECLARATION + xml);
}

/**
 * Text inside a `w:t`.
 *
 * `xml:space="preserve"` on every one of them: a run that begins or ends with a
 * space is ordinary — `**bold** text` produces exactly that — and without the
 * attribute Word drops it, closing up the gap between two words.
 */
function textElement(value: string): string {
  return value
    .split("\n")
    .map((line) => `<w:t xml:space="preserve">${escapeXml(line)}</w:t>`)
    .join("<w:br/>");
}

/** A column's alignment as `w:jc`. Left is Word's default and needs no element. */
function justification(align: Align | undefined): string {
  if (align === "right") {
    return '<w:jc w:val="right"/>';
  }
  return align === "center" ? '<w:jc w:val="center"/>' : "";
}

/**
 * `colour` overrides whatever the run's styles would have given it, and exists
 * for the one case a `Run` cannot express: text on a brand-filled table header,
 * which has to be white regardless of what it is. Putting it in the AST would
 * make colour a property of the document rather than of the rendering.
 */
function runProperties(run: Run, colour?: string): string {
  const parts = [
    run.bold ? "<w:b/>" : "",
    run.italic ? "<w:i/>" : "",
    run.code ? '<w:rStyle w:val="CodeChar"/>' : "",
    run.href ? '<w:rStyle w:val="Hyperlink"/>' : "",
    colour ? `<w:color w:val="${colour}"/>` : "",
  ].join("");
  return parts ? `<w:rPr>${parts}</w:rPr>` : "";
}

class Renderer {
  /** Hyperlink targets, in the order they were first seen; the index is the id. */
  private readonly links: string[] = [];

  private relationshipFor(href: string): string {
    let index = this.links.indexOf(href);
    if (index === -1) {
      index = this.links.push(href) - 1;
    }
    // rId1 is styles.xml and rId2 the footer, so links start above both.
    return `rId${index + FIRST_LINK_RELATIONSHIP}`;
  }

  linkTargets(): readonly string[] {
    return this.links;
  }

  /**
   * Runs, with consecutive links to the same target wrapped in one
   * `w:hyperlink`. Word treats two adjacent hyperlink elements as two links, so
   * `[**bold** text](url)` would otherwise be two clickable pieces with a seam.
   */
  private runs(runs: readonly Run[], colour?: string): string {
    let out = "";
    for (let index = 0; index < runs.length; ) {
      const href = runs[index]!.href;
      if (!href) {
        out += `<w:r>${runProperties(runs[index]!, colour)}${textElement(runs[index]!.text)}</w:r>`;
        index += 1;
        continue;
      }
      let end = index;
      while (end < runs.length && runs[end]!.href === href) {
        end += 1;
      }
      const inner = runs
        .slice(index, end)
        .map((run) => `<w:r>${runProperties(run, colour)}${textElement(run.text)}</w:r>`)
        .join("");
      out += `<w:hyperlink r:id="${this.relationshipFor(href)}">${inner}</w:hyperlink>`;
      index = end;
    }
    return out;
  }

  private paragraph(runs: readonly Run[], properties = "", colour?: string): string {
    return `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ""}${this.runs(runs, colour)}</w:p>`;
  }

  private list(block: Extract<Block, { kind: "list" }>): string {
    // One counter per depth, reset whenever the list steps back out — so a
    // nested list numbers from 1 each time it is entered rather than carrying
    // on from where the previous nested list left off.
    const counters: number[] = [];
    return block.items
      .map((item) => {
        counters.length = item.depth + 1;
        counters[item.depth] = (counters[item.depth] ?? 0) + 1;
        const marker = block.ordered ? `${counters[item.depth]}. ` : "- ";
        const left = INDENT_STEP * (item.depth + 1);
        return this.paragraph([{ text: marker }, ...item.runs], `<w:ind w:left="${left}" w:hanging="${INDENT_STEP}"/>`);
      })
      .join("");
  }

  private table(block: Extract<Block, { kind: "table" }>): string {
    const width = Math.max(block.header.length, ...block.rows.map((row) => row.length), 1);
    const widths = columnShares([block.header, ...block.rows], width).map((share) =>
      Math.round(share * TABLE_WIDTH),
    );
    const cell = (runs: readonly Run[], row: number, column: number): string => {
      const header = row === 0;
      // Zebra on alternate data rows. Counted from the header so the first data
      // row is the plain one — a tint immediately under a filled header reads as
      // the header being two rows tall.
      const fill = header ? PALETTE.brand : row % 2 === 0 ? PALETTE.brandTint : undefined;
      return (
        `<w:tc><w:tcPr><w:tcW w:w="${widths[column]}" w:type="dxa"/>` +
        (fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : "") +
        "</w:tcPr>" +
        // A `w:tc` with no `w:p` in it is what makes Word call a file corrupt, so
        // an empty cell still gets an empty paragraph.
        this.paragraph(
          header ? runs.map((run) => ({ ...run, bold: true })) : runs,
          justification(block.align[column]),
          header ? PALETTE.onBrand : undefined,
        ) +
        "</w:tc>"
      );
    };
    const row = (cells: readonly Run[][], index: number): string =>
      "<w:tr>" +
      // The header repeats when the table breaks across pages; a column of
      // numbers with no heading over it is a column nobody can read.
      (index === 0 ? "<w:trPr><w:tblHeader/></w:trPr>" : "") +
      Array.from({ length: width }, (_, column) => cell(cells[column] ?? [], index, column)).join("") +
      "</w:tr>";
    const grid = `<w:tblGrid>${widths.map((value) => `<w:gridCol w:w="${value}"/>`).join("")}</w:tblGrid>`;
    // Horizontal rules only. A full grid boxes every number in; the eye reads a
    // table by its rows, and the vertical lines are doing no work the column
    // gaps are not already doing.
    const edge = (name: string): string =>
      `<w:${name} w:val="single" w:sz="4" w:color="${PALETTE.rule}"/>`;
    const borders =
      "<w:tblBorders>" +
      edge("top") +
      edge("bottom") +
      edge("insideH") +
      '<w:left w:val="none" w:sz="0" w:space="0"/><w:right w:val="none" w:sz="0" w:space="0"/>' +
      '<w:insideV w:val="none" w:sz="0" w:space="0"/></w:tblBorders>';
    return (
      // A fixed layout with a grid, rather than `auto`: autofit sizes columns to
      // whatever is shortest, which for a two-column table of a label and a
      // sentence produces a table an inch wide in the corner of the page.
      `<w:tbl><w:tblPr><w:tblW w:w="${TABLE_WIDTH}" w:type="dxa"/>` +
      `<w:tblLayout w:type="fixed"/>${borders}</w:tblPr>${grid}` +
      row(block.header, 0) +
      block.rows.map((cells, index) => row(cells, index + 1)).join("") +
      // Word wants a paragraph between a table and whatever follows it,
      // including the end of the body.
      "</w:tbl><w:p/>"
    );
  }

  block(block: Block): string {
    switch (block.kind) {
      case "heading":
        return this.paragraph(block.runs, `<w:pStyle w:val="Heading${block.level}"/>`);
      case "paragraph":
        return this.paragraph(block.runs);
      case "list":
        return this.list(block);
      case "code":
        // One paragraph per line rather than one paragraph with breaks: a long
        // listing then breaks across pages instead of overflowing one.
        return (block.text === "" ? [""] : block.text.split("\n"))
          .map((line) => this.paragraph([{ text: line }], '<w:pStyle w:val="Code"/>'))
          .join("");
      case "quote":
        return this.paragraph(block.runs, '<w:pStyle w:val="Quote"/>');
      case "table":
        return this.table(block);
      case "rule":
        return this.paragraph(
          [],
          `<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="${PALETTE.rule}"/></w:pBdr>`,
        );
    }
  }
}

function documentXml(document: MarkdownDocument, renderer: Renderer): string {
  const body = document.blocks.map((block) => renderer.block(block)).join("");
  return (
    `<w:document xmlns:w="${W}" xmlns:r="${R}">` +
    // The footer reference comes before the page size: `w:sectPr` puts its
    // header and footer references first, and Word rejects the other order.
    `<w:body>${body}<w:sectPr>` +
    `<w:footerReference w:type="default" r:id="${FOOTER_RELATIONSHIP}"/>${PAGE}` +
    "</w:sectPr></w:body></w:document>"
  );
}

/**
 * The page number, as a field rather than as a number.
 *
 * `PAGE` is computed by Word when the document is opened or printed, which is
 * the only way a footer can be right in a document whose length this renderer
 * does not decide — text reflows to the reader's fonts, and a number written
 * here would be wrong the first time it did.
 */
function footerXml(): string {
  const properties =
    `<w:rPr><w:color w:val="${PALETTE.inkMuted}"/><w:sz w:val="${halfPoints(DOC.caption)}"/></w:rPr>`;
  return (
    `<w:ftr xmlns:w="${W}" xmlns:r="${R}"><w:p>` +
    '<w:pPr><w:jc w:val="right"/><w:spacing w:after="0"/></w:pPr>' +
    '<w:fldSimple w:instr=" PAGE ">' +
    `<w:r>${properties}<w:t>1</w:t></w:r>` +
    "</w:fldSimple></w:p></w:ftr>"
  );
}

function stylesXml(): string {
  // Levels 1 and 2 carry a hairline under them. It is the whole of what the
  // console's lavender page became here: a full-bleed tint costs ink on every
  // page and survives no photocopier, and a rule in the brand colour says the
  // same thing in a hundredth of the area.
  const underline =
    `<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="${PALETTE.brandLight}"/></w:pBdr>`;
  const heading = (level: number): string =>
    `<w:style w:type="paragraph" w:styleId="Heading${level}">` +
    `<w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/>` +
    `<w:pPr><w:keepNext/><w:spacing w:before="${level === 1 ? 240 : 200}" w:after="80"/>` +
    (level <= 2 ? underline : "") +
    `<w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
    `<w:rPr><w:b/><w:color w:val="${PALETTE.brand}"/>` +
    `<w:sz w:val="${halfPoints(DOC.headings[level - 1] ?? DOC.body)}"/></w:rPr></w:style>`;
  return (
    `<w:styles xmlns:w="${W}">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:color w:val="${PALETTE.ink}"/>` +
    `<w:sz w:val="${halfPoints(DOC.body)}"/></w:rPr></w:rPrDefault>` +
    '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    [1, 2, 3, 4, 5, 6].map(heading).join("") +
    // A quote is set off by a brand bar rather than by indentation alone, which
    // is the console's own way of marking an aside.
    '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:ind w:left="720"/>' +
    `<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="${PALETTE.brandLight}"/></w:pBdr></w:pPr>` +
    `<w:rPr><w:i/><w:color w:val="${PALETTE.inkMuted}"/></w:rPr></w:style>` +
    '<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="360"/>' +
    `<w:shd w:val="clear" w:color="auto" w:fill="${PALETTE.brandTint}"/></w:pPr>` +
    '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>' +
    `<w:sz w:val="${halfPoints(DOC.code)}"/></w:rPr></w:style>` +
    '<w:style w:type="character" w:styleId="CodeChar"><w:name w:val="Code Char"/>' +
    '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>' +
    `<w:sz w:val="${halfPoints(DOC.code)}"/><w:color w:val="${PALETTE.brandDeep}"/></w:rPr></w:style>` +
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>' +
    `<w:rPr><w:color w:val="${PALETTE.brandDeep}"/><w:u w:val="single"/></w:rPr></w:style>` +
    "</w:styles>"
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

function contentTypesXml(): string {
  return (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    "</Types>"
  );
}

function packageRelsXml(): string {
  return (
    `<Relationships xmlns="${RELATIONSHIPS}">` +
    `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>` +
    `<Relationship Id="rId2" Type="${RELATIONSHIPS}/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="${R}/extended-properties" Target="docProps/app.xml"/>` +
    "</Relationships>"
  );
}

function documentRelsXml(links: readonly string[]): string {
  return (
    `<Relationships xmlns="${RELATIONSHIPS}">` +
    `<Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/>` +
    `<Relationship Id="${FOOTER_RELATIONSHIP}" Type="${R}/footer" Target="footer1.xml"/>` +
    links
      .map(
        (href, index) =>
          `<Relationship Id="rId${index + FIRST_LINK_RELATIONSHIP}" Type="${R}/hyperlink" Target="${escapeXml(href)}" TargetMode="External"/>`,
      )
      .join("") +
    "</Relationships>"
  );
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

export interface DocxOptions {
  title: string;
  /** ISO 8601, passed in so the bytes are a function of the input alone. */
  created: string;
}

export function renderDocx(document: MarkdownDocument, options: DocxOptions): Uint8Array {
  const renderer = new Renderer();
  // Before the relationships part: rendering is what discovers the hyperlinks.
  const body = documentXml(document, renderer);
  return buildZip({
    "[Content_Types].xml": part(contentTypesXml()),
    "_rels/.rels": part(packageRelsXml()),
    "docProps/core.xml": part(corePropertiesXml(options.title, options.created)),
    "docProps/app.xml": part(appPropertiesXml()),
    "word/document.xml": part(body),
    "word/_rels/document.xml.rels": part(documentRelsXml(renderer.linkTargets())),
    "word/styles.xml": part(stylesXml()),
    "word/footer1.xml": part(footerXml()),
  });
}
