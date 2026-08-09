/**
 * DOCX to the text a model should read.
 *
 * Only `word/document.xml`. Headers, footers, footnotes and comments live in
 * their own parts, and pulling them in would interleave running heads with
 * body prose at every page boundary — text that reads as the document saying
 * something it does not say. What is left out is left out visibly: the note
 * this returns says the body is what came back.
 *
 * Two pieces of structure are kept because they carry meaning a reader would
 * otherwise have to guess at: a heading style becomes its Markdown `#` prefix,
 * and a numbered or bulleted paragraph becomes `- `. Everything else — fonts,
 * colours, spacing — is presentation, and a model has no use for it.
 */

import { walkXml, type XmlHandler } from "../xml.js";
import { openZip } from "../zip.js";
import { normalize } from "./lines.js";
import { DocumentError } from "../errors.js";

const DOCUMENT_PART = "word/document.xml";

export class DocxError extends DocumentError {}

export interface DocxText {
  text: string;
  /** How many paragraphs the body held, for a caller that wants to say so. */
  paragraphs: number;
}

function attribute(attributes: string, name: string): string | undefined {
  const match = new RegExp(`(?:^|\\s)${name}\\s*=\\s*"([^"]*)"`).exec(attributes);
  return match?.[1];
}

class Extractor implements XmlHandler {
  private readonly lines: string[] = [];
  private buffer = "";
  private prefix = "";
  private textDepth = 0;
  private cellDepth = 0;
  paragraphs = 0;

  text(value: string): void {
    if (this.textDepth > 0) {
      this.buffer += value;
    }
  }

  open(name: string, attributes: string, selfClosing: boolean): void {
    switch (name) {
      case "w:t":
        // Not self-closing `<w:t/>`, which holds nothing and would leave the
        // depth raised for the rest of the document.
        if (!selfClosing) {
          this.textDepth += 1;
        }
        return;
      case "w:tab":
        this.buffer += "\t";
        return;
      case "w:br":
      case "w:cr":
        this.buffer += "\n";
        return;
      case "w:p":
        this.prefix = "";
        return;
      case "w:tc":
        this.cellDepth += 1;
        return;
      case "w:numPr":
        this.prefix = "- ";
        return;
      case "w:pStyle": {
        const level = /^(?:Heading|heading)\s*([1-6])$/.exec(attribute(attributes, "w:val") ?? "");
        if (level?.[1]) {
          this.prefix = `${"#".repeat(Number(level[1]))} `;
        }
        return;
      }
      default:
        return;
    }
  }

  close(name: string): void {
    switch (name) {
      case "w:t":
        this.textDepth = Math.max(0, this.textDepth - 1);
        return;
      case "w:p":
        this.paragraphs += 1;
        // Inside a cell a paragraph is a line wrap, not a row break: ending the
        // line here would turn every multi-paragraph cell into its own row.
        if (this.cellDepth > 0) {
          this.buffer += " ";
          return;
        }
        this.endLine();
        return;
      case "w:tc":
        this.cellDepth = Math.max(0, this.cellDepth - 1);
        // Cell boundaries carry meaning in a table — a row of values run
        // together is not readable as a row.
        this.buffer += " | ";
        return;
      case "w:tr":
        this.endLine();
        return;
      default:
        return;
    }
  }

  private endLine(): void {
    const line = this.prefix + this.buffer;
    this.buffer = "";
    this.prefix = "";
    for (const part of line.split("\n")) {
      this.lines.push(part);
    }
  }

  finish(): string {
    this.endLine();
    return normalize(this.lines);
  }
}

/** The body's XML, separated so the walk above can be tested without a zip. */
export function documentXmlToText(xml: string): DocxText {
  const extractor = new Extractor();
  walkXml(xml, extractor);
  return { text: extractor.finish(), paragraphs: extractor.paragraphs };
}

export function docxToText(bytes: Uint8Array): DocxText {
  const archive = openZip(bytes);
  const part = archive.read([DOCUMENT_PART]).get(DOCUMENT_PART);
  if (!part) {
    throw new DocxError(
      `this .docx has no ${DOCUMENT_PART}, so it has no body — it is not a document Word wrote`,
    );
  }
  // OOXML parts are XML, and XML without a declared encoding is UTF-8. Word
  // writes the declaration and writes UTF-8; nothing here has ever needed the
  // charset dance the plain-text path does.
  const result = documentXmlToText(new TextDecoder("utf-8").decode(part));
  if (result.text === "") {
    throw new DocxError(
      result.paragraphs > 0
        ? `this .docx has ${result.paragraphs} paragraph(s) but no text in any of them — its ` +
          `content is most likely images, which need OCR rather than text extraction`
        : "this .docx has no text in its body",
    );
  }
  return result;
}
