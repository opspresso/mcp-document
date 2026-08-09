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
import type { Block, MarkdownDocument, Run } from "../markdown.js";
import { columnShares } from "./table.js";

/** A4, in twentieths of a point, with a 2.5cm margin. */
const PAGE = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/>';

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** One indent level, in twentieths of a point: 0.63cm, Word's own list step. */
const INDENT_STEP = 360;

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

function runProperties(run: Run): string {
  const parts = [
    run.bold ? "<w:b/>" : "",
    run.italic ? "<w:i/>" : "",
    run.code ? '<w:rStyle w:val="CodeChar"/>' : "",
    run.href ? '<w:rStyle w:val="Hyperlink"/>' : "",
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
    // rId1 is styles.xml, so links start above it.
    return `rId${index + 2}`;
  }

  linkTargets(): readonly string[] {
    return this.links;
  }

  /**
   * Runs, with consecutive links to the same target wrapped in one
   * `w:hyperlink`. Word treats two adjacent hyperlink elements as two links, so
   * `[**bold** text](url)` would otherwise be two clickable pieces with a seam.
   */
  private runs(runs: readonly Run[]): string {
    let out = "";
    for (let index = 0; index < runs.length; ) {
      const href = runs[index]!.href;
      if (!href) {
        out += `<w:r>${runProperties(runs[index]!)}${textElement(runs[index]!.text)}</w:r>`;
        index += 1;
        continue;
      }
      let end = index;
      while (end < runs.length && runs[end]!.href === href) {
        end += 1;
      }
      const inner = runs
        .slice(index, end)
        .map((run) => `<w:r>${runProperties(run)}${textElement(run.text)}</w:r>`)
        .join("");
      out += `<w:hyperlink r:id="${this.relationshipFor(href)}">${inner}</w:hyperlink>`;
      index = end;
    }
    return out;
  }

  private paragraph(runs: readonly Run[], properties = ""): string {
    return `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ""}${this.runs(runs)}</w:p>`;
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
    const cell = (runs: readonly Run[], header: boolean, column: number): string =>
      `<w:tc><w:tcPr><w:tcW w:w="${widths[column]}" w:type="dxa"/></w:tcPr>` +
      // A `w:tc` with no `w:p` in it is what makes Word call a file corrupt, so
      // an empty cell still gets an empty paragraph.
      `${this.paragraph(header ? runs.map((run) => ({ ...run, bold: true })) : runs)}</w:tc>`;
    const row = (cells: readonly Run[][], header: boolean): string =>
      `<w:tr>${Array.from({ length: width }, (_, index) => cell(cells[index] ?? [], header, index)).join("")}</w:tr>`;
    const grid = `<w:tblGrid>${widths.map((value) => `<w:gridCol w:w="${value}"/>`).join("")}</w:tblGrid>`;
    const borders =
      '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:color="BFBFBF"/>' +
      '<w:bottom w:val="single" w:sz="4" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:color="BFBFBF"/>' +
      '<w:insideH w:val="single" w:sz="4" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:color="BFBFBF"/></w:tblBorders>';
    return (
      // A fixed layout with a grid, rather than `auto`: autofit sizes columns to
      // whatever is shortest, which for a two-column table of a label and a
      // sentence produces a table an inch wide in the corner of the page.
      `<w:tbl><w:tblPr><w:tblW w:w="${TABLE_WIDTH}" w:type="dxa"/>` +
      `<w:tblLayout w:type="fixed"/>${borders}</w:tblPr>${grid}` +
      row(block.header, true) +
      block.rows.map((cells) => row(cells, false)).join("") +
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
          '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="BFBFBF"/></w:pBdr>',
        );
    }
  }
}

function documentXml(document: MarkdownDocument, renderer: Renderer): string {
  const body = document.blocks.map((block) => renderer.block(block)).join("");
  return (
    `<w:document xmlns:w="${W}" xmlns:r="${R}">` +
    `<w:body>${body}<w:sectPr>${PAGE}</w:sectPr></w:body></w:document>`
  );
}

const HEADING_SIZES = [32, 28, 26, 24, 22, 22];

function stylesXml(): string {
  const heading = (level: number): string =>
    `<w:style w:type="paragraph" w:styleId="Heading${level}">` +
    `<w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/>` +
    `<w:pPr><w:keepNext/><w:spacing w:before="${level === 1 ? 240 : 200}" w:after="80"/>` +
    `<w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
    `<w:rPr><w:b/><w:sz w:val="${HEADING_SIZES[level - 1]}"/></w:rPr></w:style>`;
  return (
    `<w:styles xmlns:w="${W}">` +
    '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    [1, 2, 3, 4, 5, 6].map(heading).join("") +
    '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/><w:color w:val="595959"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="360"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="19"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="CodeChar"><w:name w:val="Code Char"/>' +
    '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="19"/><w:color w:val="C7254E"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>' +
    '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
    "</w:styles>"
  );
}

function contentTypesXml(): string {
  return (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    "</Types>"
  );
}

function packageRelsXml(): string {
  return (
    `<Relationships xmlns="${RELATIONSHIPS}">` +
    `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>` +
    `<Relationship Id="rId2" Type="${RELATIONSHIPS}/metadata/core-properties" Target="docProps/core.xml"/>` +
    "</Relationships>"
  );
}

function documentRelsXml(links: readonly string[]): string {
  return (
    `<Relationships xmlns="${RELATIONSHIPS}">` +
    `<Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/>` +
    links
      .map(
        (href, index) =>
          `<Relationship Id="rId${index + 2}" Type="${R}/hyperlink" Target="${escapeXml(href)}" TargetMode="External"/>`,
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
    "word/document.xml": part(body),
    "word/_rels/document.xml.rels": part(documentRelsXml(renderer.linkTargets())),
    "word/styles.xml": part(stylesXml()),
  });
}
