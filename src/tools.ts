/**
 * The tools, and everything between a JSON-RPC `tools/call` and the answer.
 *
 * Separated from `server.ts` because the descriptions are the part of this
 * server a model actually reads. They say what each tool takes, what it
 * returns, and — as much as anything else — when *not* to reach for it: a
 * model that sends a web page here has spent a turn learning something the
 * description could have told it.
 *
 * Every failure comes back as a tool error with a reason rather than as a
 * protocol error. A failed read is the model's problem to react to, not the
 * run's, and a JSON-RPC error would take the whole turn down with it.
 */

import { extensionOf } from "./detect.js";
import { DocumentError } from "./errors.js";
import {
  asUntrustedContent,
  MAX_ASSET_COUNT,
  MAX_ASSET_TOTAL_BYTES,
  MAX_MARKDOWN_CHARS,
  MAX_RENDERED_BYTES,
  MAX_TEXT_CHARS,
} from "./limits.js";
import { elapsedMs, logError, logInfo, logWarn } from "./log.js";
import { parseMarkdown, withoutDirectives, type Block, type MarkdownDocument } from "./markdown.js";
import { readDocument, UnsupportedDocument } from "./read/document.js";
import { inspectXlsx, XlsxError, type InspectedCell, type XlsxInspection } from "./read/xlsx.js";
import { decodeBase64, loadSource, SourceError } from "./source.js";
import { safeFilename } from "./filename.js";
import { renderDocx } from "./write/docx.js";
import { renderHwpx } from "./write/hwpx.js";
import { renderPdf } from "./write/pdf.js";
import { renderPptx } from "./write/pptx/index.js";
import { renderXlsx } from "./write/xlsx.js";
import type { ImageAsset } from "./write/image.js";
import { validateRenderedDocument, ValidationError } from "./validate.js";
import {
  DEFAULT_PROFILE,
  DOCUMENT_PROFILES,
  type DocumentProfile,
} from "./write/theme.js";

/**
 * A text block, or an embedded resource carrying bytes.
 *
 * The second is new, and it is what replaced uploading to S3 and answering with
 * a link: this server no longer knows a bucket exists. The caller stores what it
 * receives, alongside every other byte one of its runs produced, and owns the
 * retention and the delete button that go with that.
 */
export type ToolContent =
  | { type: "text"; text: string }
  | {
      type: "resource";
      resource: { uri: string; mimeType: string; blob: string };
    };

export interface ToolResult {
  content: ToolContent[];
  structuredContent?: Record<string, unknown>;
  isError?: true;
}

export const FORMATS = ["docx", "pdf", "hwpx", "pptx"] as const;
export type Format = (typeof FORMATS)[number];
export const PROFILES = DOCUMENT_PROFILES;

export const CONTENT_TYPES: Record<Format, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pdf: "application/pdf",
  hwpx: "application/hwp+zip",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

export const XLSX_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export const TOOL_OUTPUT_SCHEMA = {
  type: "object",
  properties: {
    operation: { type: "string" },
  },
  required: ["operation"],
  additionalProperties: true,
} as const;

export const TOOLS = [
  {
    name: "read_document",
    description:
      "Read a document and return its text. Handles DOCX (Word), XLSX (Excel), PPTX " +
      "(PowerPoint), HWP and HWPX (한글), OpenDocument (ODT/ODS/ODP) and RTF. Pass the file's " +
      "bytes as base64 in `content`. Use this when you have a document you cannot otherwise " +
      "open: a report, a spec, a contract, a spreadsheet, a deck. Returns the contents, not a " +
      "summary. PDF, plain text and web pages are not read here — the caller extracts those " +
      "itself, and sending one gets a refusal saying so. A spreadsheet comes back as values, " +
      "never formulas, one visible sheet per heading.",
    inputSchema: {
      type: "object",
      properties: {
        content: { type: "string", description: "The document's bytes, base64-encoded." },
        filename: {
          type: "string",
          description:
            "The document's filename, e.g. 'report.hwp'. Optional, and only a hint: it helps " +
            "identify a format whose bytes are ambiguous.",
        },
      },
      required: ["content"],
    },
  },
  {
    name: "inspect_spreadsheet",
    description:
      "Inspect an XLSX workbook without executing it. Returns cell addresses and either cached " +
      "values, formulas, or both; formulas are never recalculated. Hidden and very-hidden sheets " +
      "are excluded unless `includeHidden` is true. Reports macros and external workbook links " +
      "without opening or following them. Use this for formula review, error-cell diagnosis, and " +
      "workbook structure. Use read_document instead when values as simple sheet text are enough.",
    inputSchema: {
      type: "object",
      properties: {
        content: { type: "string", description: "The XLSX workbook bytes, base64-encoded." },
        filename: {
          type: "string",
          description: "The workbook filename, used only in the untrusted-content label.",
        },
        mode: {
          type: "string",
          enum: ["values", "formulas", "both"],
          description: "What each addressed cell returns. Defaults to both.",
        },
        includeHidden: {
          type: "boolean",
          description: "Include hidden and very-hidden sheets. Defaults to false.",
        },
      },
      required: ["content"],
    },
  },
  {
    name: "render_spreadsheet",
    description:
      "Create a new XLSX workbook and return the file. Pass one or more named sheets whose rows " +
      "are arrays of strings, finite numbers, booleans or null. A string beginning with `=` stays " +
      "literal; a formula must be explicit as `{ formula: \"SUM(B2:B5)\", cachedValue: 42 }`. " +
      "`cachedValue` is optional and is never calculated or verified here. The workbook requests a " +
      "full recalculation when opened. Use this for a new portable workbook, not for editing or " +
      "preserving an existing workbook's styles, charts, macros, comments or external links.",
    inputSchema: {
      type: "object",
      properties: {
        sheets: {
          type: "array",
          minItems: 1,
          maxItems: 50,
          items: {
            type: "object",
            properties: {
              name: { type: "string", minLength: 1, maxLength: 31 },
              rows: {
                type: "array",
                items: {
                  type: "array",
                  items: {
                    oneOf: [
                      { type: "string" },
                      { type: "number" },
                      { type: "boolean" },
                      { type: "null" },
                      {
                        type: "object",
                        properties: {
                          formula: { type: "string", minLength: 1 },
                          cachedValue: {
                            oneOf: [
                              { type: "string" },
                              { type: "number" },
                              { type: "boolean" },
                              { type: "null" },
                            ],
                          },
                        },
                        required: ["formula"],
                        additionalProperties: false,
                      },
                    ],
                  },
                },
              },
            },
            required: ["name", "rows"],
            additionalProperties: false,
          },
        },
        title: {
          type: "string",
          description: "Workbook title used in package metadata. Defaults to workbook.",
        },
        filename: {
          type: "string",
          description: "Filename without the extension. Defaults to the title.",
        },
      },
      required: ["sheets"],
    },
  },
  {
    name: "render_document",
    description:
      "Write a document and return the file. Takes Markdown and produces DOCX " +
      "(Word), PDF, HWPX (한글) or PPTX (PowerPoint) — pick the one the reader will open. Use " +
      "this to deliver something a person will read or share: a report, a summary, meeting " +
      "notes, a proposal, a slide deck. Pick a `profile` for its purpose: executive for " +
      "leadership decisions, consulting for strategy, formal for submissions, technical for " +
      "engineering, or standard for the classic corporate style; executive is the default. " +
      "Supports headings, paragraphs, bold, italic, inline " +
      "code, links, bullet and numbered lists, tables, block quotes, fenced code blocks and " +
      "horizontal rules. A table column is set flush right with `---:` in the divider row and " +
      "centred with `:---:` — set columns of numbers right, or their digits do not line up. " +
      "Two lines with no blank line between them are one paragraph, as in " +
      "Markdown. In docx, pdf and hwpx the Markdown becomes a designed report: an opening " +
      "`#` is a cover page with the first paragraph as its subtitle, every later `#` a " +
      "numbered chapter on a fresh page, and a cover plus three or more `#`/`##` headings " +
      "adds a contents page listing them. Block quotes are callout boxes; in docx " +
      "`:::metrics` renders a key-figure strip and `:::comparison` a two-column table. " +
      "In pptx the Markdown " +
      "becomes a designed deck: an opening `#` is the cover and its first paragraph the " +
      "subtitle, every later `#` a numbered section divider, every `##` a slide. A slide " +
      "whose shape says what it is gets a matching layout — two to four `###`s with a short " +
      "line each become cards, a short bullet list of figures (`- 99.99% Availability`) " +
      "becomes big-number metrics, a lone block quote (with `— author`) a quote slide, two " +
      "`###`s under an 'A vs B' title a two-column comparison, three to five short numbered " +
      "steps a process flow, date-led steps (`1. Q1 파일럿`) a timeline, and a final 감사합니다 " +
      "or Thank-you heading the closing slide. Wrap one such group in `:::cards` … `:::` " +
      "(also metrics, comparison, process, timeline, quote) to force the layout when the " +
      "shape alone would not; formats with no treatment for a directive render its contents " +
      "as if the fences were never written. In pptx, docx and pdf an image whose bytes were sent " +
      "in `assets` and which stands alone becomes an embedded, captioned figure. " +
      "Returns the file itself; the caller delivers it to the user.",
    inputSchema: {
      type: "object",
      properties: {
        format: {
          type: "string",
          enum: [...FORMATS],
          description:
            "docx for Word, pdf to be read as-is, hwpx for 한글, pptx for a slide deck.",
        },
        profile: {
          type: "string",
          enum: [...PROFILES],
          description:
            "Document design and editorial profile. executive for leadership decisions, " +
            "consulting for strategy and proposals, formal for public or external submissions, " +
            "technical for architecture and engineering, standard for the classic corporate style. " +
            "Defaults to executive.",
        },
        content: { type: "string", description: "The document body, as Markdown." },
        title: {
          type: "string",
          description:
            "The document's title, used in its metadata and as the filename. Defaults to the " +
            "first level-1 heading in the content.",
        },
        filename: {
          type: "string",
          description: "Filename without the extension. Defaults to the title.",
        },
        assets: {
          type: "object",
          description:
            "Images to embed in pptx, docx and pdf: each key is a name the Markdown references " +
            "as `![caption](asset://name)`, each value the image. A slide or paragraph that " +
            "is exactly one such image becomes a full-width figure with the caption under it. " +
            "PNG and JPEG only — rasterise an SVG before sending it.",
          additionalProperties: {
            type: "object",
            properties: {
              mimeType: { type: "string", enum: ["image/png", "image/jpeg"] },
              content: { type: "string", description: "The image's bytes, base64-encoded." },
            },
            required: ["mimeType", "content"],
          },
        },
      },
      required: ["format", "content"],
    },
  },
] as const;

function ok(text: string, structuredContent?: Record<string, unknown>): ToolResult {
  return {
    content: [{ type: "text", text }],
    ...(structuredContent ? { structuredContent } : {}),
  };
}

function failed(
  text: string,
  details: {
    operation: string;
    code: string;
    retryable?: boolean;
    field?: string;
    suggestedFix?: string;
  } = { operation: "unknown", code: "INVALID_ARGUMENT" },
): ToolResult {
  return {
    content: [{ type: "text", text }],
    structuredContent: {
      operation: details.operation,
      error: {
        code: details.code,
        retryable: details.retryable ?? false,
        ...(details.field ? { field: details.field } : {}),
        ...(details.suggestedFix ? { suggestedFix: details.suggestedFix } : {}),
      },
    },
    isError: true,
  };
}

function failureCode(error: unknown): string {
  if (error instanceof UnsupportedDocument) {
    return "UNSUPPORTED_FORMAT";
  }
  if (error instanceof ValidationError) {
    return "STRUCTURE_VALIDATION_FAILED";
  }
  if (error instanceof SourceError) {
    return /over the .*limit/.test(error.message) ? "INPUT_TOO_LARGE" : "INVALID_INPUT";
  }
  if (error instanceof XlsxError) {
    return "INVALID_WORKBOOK";
  }
  return error instanceof DocumentError ? "INVALID_DOCUMENT" : "INTERNAL_ERROR";
}

/**
 * The origin, and nothing else.
 *
 * A URL the model read out of another tool's output can carry its capability in
 * the URL itself — a signed query string, or a secret path segment as a Slack
 * webhook does. A log line is the last place either should come to rest, and
 * the question this line answers — "everything to that host started failing on
 * Tuesday" — is asked about the host.
 */
function originOf(url: string | undefined): string {
  if (!url) {
    return "(inline content)";
  }
  try {
    return new URL(url).origin;
  } catch {
    return "(unparseable url)";
  }
}

function describe(error: unknown): string {
  if (error instanceof Error) {
    return error.name === "TimeoutError" || error.name === "AbortError"
      ? "the request timed out"
      : error.message;
  }
  return String(error);
}

async function read(args: Record<string, unknown>): Promise<ToolResult> {
  const content = typeof args.content === "string" && args.content ? args.content : undefined;
  const filename = typeof args.filename === "string" && args.filename ? args.filename : undefined;
  try {
    const source = loadSource({ content, filename });
    const result = await readDocument(source);
    return ok(asUntrustedContent(source.label, result.text, result.note), {
      operation: "read_document",
      sourceFormat: result.format,
      complete: result.complete,
      omissions: result.omissions,
      ...(result.counts ? { counts: result.counts } : {}),
    });
  } catch (error) {
    // The caller is told; without this line the operator is not, and a format
    // that started failing has no evidence behind it anywhere. A filename is
    // safe to log — unlike a URL, it carries no capability.
    (error instanceof DocumentError ? logWarn : logError)("tool_failed", error, {
      tool: "read_document",
      filename,
    });
    return failed(
      `Error: could not read the document — ${
        error instanceof DocumentError ? error.message : describe(error)
      }`,
      {
        operation: "read_document",
        code: failureCode(error),
        suggestedFix: "Pass supported document bytes as base64 and follow the reason in the message.",
      },
    );
  }
}

type InspectionMode = "values" | "formulas" | "both";

function inspectedLine(cell: InspectedCell, mode: InspectionMode): string | undefined {
  if (mode === "formulas") {
    return cell.formula ? `${cell.address}: =${cell.formula}` : undefined;
  }
  if (mode === "values") {
    return `${cell.address}: ${cell.value}`;
  }
  return cell.formula
    ? `${cell.address}: =${cell.formula} → ${cell.value}${cell.error ? ` (${cell.error})` : ""}`
    : `${cell.address}: ${cell.value}`;
}

function formatInspection(
  inspection: XlsxInspection,
  mode: InspectionMode,
): { text: string; sheets: Array<{ name: string; state: string; cells: InspectedCell[] }>; complete: boolean } {
  const lines: string[] = [];
  const sheets: Array<{ name: string; state: string; cells: InspectedCell[] }> = [];
  let length = 0;
  let complete = inspection.complete;
  for (const sheet of inspection.sheets) {
    const heading = `## ${sheet.name}${sheet.state === "visible" ? "" : ` [${sheet.state}]`}`;
    if (length + heading.length + 1 > MAX_TEXT_CHARS) {
      complete = false;
      break;
    }
    lines.push(heading);
    length += heading.length + 1;
    const cells: InspectedCell[] = [];
    for (const cell of sheet.cells) {
      const line = inspectedLine(cell, mode);
      if (!line) {
        continue;
      }
      if (length + line.length + 1 > MAX_TEXT_CHARS) {
        complete = false;
        break;
      }
      lines.push(line);
      cells.push(cell);
      length += line.length + 1;
    }
    sheets.push({ name: sheet.name, state: sheet.state, cells });
    lines.push("");
    length += 1;
    if (!complete) {
      break;
    }
  }
  return { text: lines.join("\n").trim(), sheets, complete };
}

async function inspectSpreadsheet(args: Record<string, unknown>): Promise<ToolResult> {
  const content = typeof args.content === "string" && args.content ? args.content : undefined;
  const filename = typeof args.filename === "string" && args.filename ? args.filename : undefined;
  const requestedMode = args.mode;
  if (
    requestedMode !== undefined &&
    requestedMode !== "values" &&
    requestedMode !== "formulas" &&
    requestedMode !== "both"
  ) {
    return failed("Error: `mode` must be one of values, formulas, both.", {
      operation: "inspect_spreadsheet",
      code: "INVALID_ARGUMENT",
      field: "mode",
      suggestedFix: "Use values, formulas or both.",
    });
  }
  if (args.includeHidden !== undefined && typeof args.includeHidden !== "boolean") {
    return failed("Error: `includeHidden` must be a boolean.", {
      operation: "inspect_spreadsheet",
      code: "INVALID_ARGUMENT",
      field: "includeHidden",
      suggestedFix: "Pass true or false.",
    });
  }
  const mode = (requestedMode ?? "both") as InspectionMode;
  const includeHidden = args.includeHidden === true;
  try {
    const source = loadSource({ content, filename });
    const inspection = inspectXlsx(source.bytes, includeHidden);
    const formatted = formatInspection(inspection, mode);
    const warnings = [
      ...(inspection.hiddenSheets > 0 && !includeHidden
        ? [`${inspection.hiddenSheets} hidden sheet(s) were omitted.`]
        : []),
      ...(inspection.externalLinks > 0
        ? [`${inspection.externalLinks} external workbook link part(s) were detected but not followed.`]
        : []),
      ...(inspection.macroEnabled ? ["A VBA project was detected but not executed."] : []),
      "Formulas were read as text and were not recalculated.",
    ];
    const note = `${formatted.complete ? "all inspected cells" : "a bounded prefix of cells"}; ${warnings.join(" ")}`;
    return ok(asUntrustedContent(source.label, formatted.text, note), {
      operation: "inspect_spreadsheet",
      sourceFormat: "xlsx",
      mode,
      complete: formatted.complete,
      hiddenIncluded: includeHidden,
      totalSheets: inspection.totalSheets,
      hiddenSheets: inspection.hiddenSheets,
      externalLinks: inspection.externalLinks,
      macroEnabled: inspection.macroEnabled,
      warnings,
      sheets: formatted.sheets,
    });
  } catch (error) {
    (error instanceof DocumentError ? logWarn : logError)("tool_failed", error, {
      tool: "inspect_spreadsheet",
      filename,
    });
    return failed(
      `Error: could not inspect the workbook — ${
        error instanceof DocumentError ? error.message : describe(error)
      }`,
      {
        operation: "inspect_spreadsheet",
        code: failureCode(error),
        suggestedFix: "Pass an XLSX workbook; formulas and active content are inspected but never executed.",
      },
    );
  }
}

async function renderSpreadsheet(args: Record<string, unknown>): Promise<ToolResult> {
  const title =
    (typeof args.title === "string" && args.title.trim()) ||
    (typeof args.filename === "string" && args.filename.trim()) ||
    "workbook";
  const filename = safeFilename(
    (typeof args.filename === "string" && args.filename.trim()) || title,
    "xlsx",
  );
  try {
    const rendered = renderXlsx(args.sheets, {
      title,
      created: new Date().toISOString(),
    });
    if (rendered.bytes.byteLength > MAX_RENDERED_BYTES) {
      return failed(
        `Error: the rendered xlsx is ${rendered.bytes.byteLength.toLocaleString("en-US")} bytes, ` +
          `over the ${MAX_RENDERED_BYTES.toLocaleString("en-US")} limit for one response.`,
        {
          operation: "render_spreadsheet",
          code: "OUTPUT_TOO_LARGE",
          suggestedFix: "Split the workbook or reduce its cells.",
        },
      );
    }
    const validation = await validateRenderedDocument("xlsx", rendered.bytes);
    const warnings = [
      ...validation.warnings,
      ...(rendered.formulas > 0
        ? ["Formula results were not calculated or verified; the workbook requests recalculation on open."]
        : []),
    ];
    return {
      content: [
        {
          type: "text",
          text:
            `Wrote ${filename} (${rendered.bytes.byteLength.toLocaleString("en-US")} bytes, ` +
            `${rendered.sheets} sheet(s), ${rendered.cells.toLocaleString("en-US")} cells, ` +
            `${rendered.formulas.toLocaleString("en-US")} formulas). The file is attached. ` +
            "Package validation passed; formulas were not calculated and visual layout was not rendered.",
        },
        {
          type: "resource",
          resource: {
            uri: `file:///${encodeURIComponent(filename)}`,
            mimeType: XLSX_CONTENT_TYPE,
            blob: Buffer.from(rendered.bytes).toString("base64"),
          },
        },
      ],
      structuredContent: {
        operation: "render_spreadsheet",
        format: "xlsx",
        filename,
        bytes: rendered.bytes.byteLength,
        counts: {
          sheets: rendered.sheets,
          rows: rendered.rows,
          cells: rendered.cells,
          formulas: rendered.formulas,
        },
        validation: { ...validation, warnings },
      },
    };
  } catch (error) {
    (error instanceof DocumentError ? logWarn : logError)("tool_failed", error, {
      tool: "render_spreadsheet",
    });
    return failed(
      `Error: could not write the workbook — ${
        error instanceof DocumentError ? error.message : describe(error)
      }`,
      {
        operation: "render_spreadsheet",
        code: failureCode(error),
        suggestedFix: "Fix the named sheet, row, cell or formula problem and render again.",
      },
    );
  }
}

/** What the document turned out to be, for the line the caller reads. */
export function summarise(document: MarkdownDocument): string {
  // Directives spliced open, so their contents are counted as what they are.
  const { blocks } = withoutDirectives(document);
  const count = (kind: Block["kind"]): number =>
    blocks.filter((block) => block.kind === kind).length;
  const parts = [
    [count("heading"), "heading"],
    [count("paragraph"), "paragraph"],
    [count("table"), "table"],
    [count("list"), "list"],
    [count("code"), "code block"],
  ] as const;
  const said = parts
    .filter(([number]) => number > 0)
    .map(([number, name]) => `${number} ${name}${number === 1 ? "" : "s"}`);
  return said.length > 0 ? said.join(", ") : "nothing";
}

async function render(args: Record<string, unknown>): Promise<ToolResult> {
  const requested = args.format;
  if (typeof requested !== "string" || !FORMATS.includes(requested as Format)) {
    return failed(`Error: \`format\` must be one of ${FORMATS.join(", ")}.`, {
      operation: "render_document",
      code: "INVALID_ARGUMENT",
      field: "format",
      suggestedFix: `Use one of ${FORMATS.join(", ")}.`,
    });
  }
  const format = requested as Format;
  const requestedProfile = args.profile;
  if (
    requestedProfile !== undefined &&
    (typeof requestedProfile !== "string" || !PROFILES.includes(requestedProfile as DocumentProfile))
  ) {
    return failed(`Error: \`profile\` must be one of ${PROFILES.join(", ")}.`, {
      operation: "render_document",
      code: "INVALID_ARGUMENT",
      field: "profile",
      suggestedFix: `Use one of ${PROFILES.join(", ")}.`,
    });
  }
  const profile = (requestedProfile ?? DEFAULT_PROFILE) as DocumentProfile;
  const markdown = args.content;
  if (typeof markdown !== "string" || markdown.trim() === "") {
    return failed("Error: `content` is required, as Markdown.", {
      operation: "render_document",
      code: "INVALID_ARGUMENT",
      field: "content",
      suggestedFix: "Pass non-empty Markdown content.",
    });
  }
  if (markdown.length > MAX_MARKDOWN_CHARS) {
    return failed(
      `Error: \`content\` is ${markdown.length.toLocaleString("en-US")} characters, over the ` +
        `${MAX_MARKDOWN_CHARS.toLocaleString("en-US")} limit. Write it as more than one document.`,
      {
        operation: "render_document",
        code: "INPUT_TOO_LARGE",
        field: "content",
        suggestedFix: "Split the content into more than one document.",
      },
    );
  }

  try {
    const document = parseMarkdown(markdown);
    const title =
      (typeof args.title === "string" && args.title.trim()) ||
      document.title ||
      (typeof args.filename === "string" && args.filename.trim()) ||
      "document";
    const filename = safeFilename(
      (typeof args.filename === "string" && args.filename.trim()) || title,
      format,
    );
    const created = new Date();
    const assets = parseAssets(args.assets, format);

    return await renderTo(format, document, { title, created, filename, assets, profile });
  } catch (error) {
    (error instanceof DocumentError ? logWarn : logError)("tool_failed", error, {
      tool: "render_document",
      format,
    });
    return failed(
      `Error: could not write the document — ${
        error instanceof DocumentError ? error.message : describe(error)
      }`,
      {
        operation: "render_document",
        code: failureCode(error),
        suggestedFix: "Follow the named format, asset or validation problem and render again.",
      },
    );
  }
}

/** A name the Markdown can reference and a zip entry can carry. */
const ASSET_NAME = /^[A-Za-z0-9가-힣][A-Za-z0-9가-힣._-]{0,63}$/;

const ASSET_MIMES = ["image/png", "image/jpeg"] as const;

/**
 * The `assets` argument, decoded and bounded.
 *
 * Everything wrong with it is refused by name — the asset, and what about it —
 * because "invalid assets" spends the model's next turn guessing. Assets are a
 * pptx, docx and pdf feature; sent with any other format they would be silently dropped
 * pictures, so that is refused too rather than ignored.
 */
function parseAssets(
  raw: unknown,
  format: Format,
): Record<string, ImageAsset> | undefined {
  if (raw === undefined || raw === null) {
    return undefined;
  }
  if (typeof raw !== "object" || Array.isArray(raw)) {
    throw new DocumentError("`assets` must be an object of name → { mimeType, content }.");
  }
  const entries = Object.entries(raw as Record<string, unknown>);
  if (entries.length === 0) {
    return undefined;
  }
  if (format !== "pptx" && format !== "docx" && format !== "pdf") {
    throw new DocumentError(
      "`assets` are embedded in pptx, docx and pdf only — hwpx renders an image as a link.",
    );
  }
  if (entries.length > MAX_ASSET_COUNT) {
    throw new DocumentError(`${entries.length} assets is over the limit of ${MAX_ASSET_COUNT}.`);
  }
  const assets: Record<string, ImageAsset> = {};
  let total = 0;
  for (const [name, value] of entries) {
    if (!ASSET_NAME.test(name)) {
      throw new DocumentError(
        `asset name "${name}" is not usable — letters, digits, 한글, dot, dash and underscore only.`,
      );
    }
    const asset = value as { mimeType?: unknown; content?: unknown };
    const mime = ASSET_MIMES.find((allowed) => allowed === asset?.mimeType);
    if (!mime) {
      throw new DocumentError(
        `asset "${name}" declares ${JSON.stringify(asset?.mimeType)} — only image/png and ` +
          "image/jpeg can be embedded. Rasterise an SVG before sending it.",
      );
    }
    if (typeof asset.content !== "string" || asset.content === "") {
      throw new DocumentError(`asset "${name}" carries no base64 content.`);
    }
    const bytes = decodeBase64(asset.content);
    total += bytes.byteLength;
    if (total > MAX_ASSET_TOTAL_BYTES) {
      throw new DocumentError(
        `the assets total more than ${MAX_ASSET_TOTAL_BYTES.toLocaleString("en-US")} bytes decoded.`,
      );
    }
    assets[name] = { mimeType: mime, bytes };
  }
  return assets;
}

/**
 * The bytes, as an embedded resource.
 *
 * This is what replaced uploading to a bucket and answering with a link. The
 * server no longer knows a bucket exists — the caller stores what it receives,
 * beside every other byte one of its runs produced, and owns the retention and
 * the delete button that come with that. It also means one fewer place holding
 * an AWS credential.
 */
async function renderTo(
  format: Format,
  document: MarkdownDocument,
  meta: {
    title: string;
    created: Date;
    filename: string;
    assets?: Record<string, ImageAsset>;
    profile: DocumentProfile;
  },
): Promise<ToolResult> {
  let bytes: Uint8Array;
  let extra = "";
  let counts: Record<string, number> = {};
  if (format === "docx") {
    bytes = renderDocx(document, {
      title: meta.title,
      created: meta.created.toISOString(),
      profile: meta.profile,
      ...(meta.assets ? { assets: meta.assets } : {}),
    });
  } else if (format === "hwpx") {
    bytes = renderHwpx(document, {
      title: meta.title,
      created: meta.created.toISOString(),
      profile: meta.profile,
    });
  } else if (format === "pptx") {
    const rendered = renderPptx(document, {
      title: meta.title,
      created: meta.created.toISOString(),
      profile: meta.profile,
      ...(meta.assets ? { assets: meta.assets } : {}),
    });
    bytes = rendered.bytes;
    counts = { slides: rendered.slides, continuations: rendered.continuations };
    extra = `, ${rendered.slides} slide${rendered.slides === 1 ? "" : "s"}`;
  } else {
    const rendered = await renderPdf(document, {
      title: meta.title,
      created: meta.created,
      profile: meta.profile,
      ...(meta.assets ? { assets: meta.assets } : {}),
    });
    bytes = rendered.bytes;
    counts = { pages: rendered.pages };
    extra = `, ${rendered.pages} page${rendered.pages === 1 ? "" : "s"}`;
  }

  if (bytes.byteLength > MAX_RENDERED_BYTES) {
    // Refused here, with a reason. The transport bounds the whole JSON-RPC
    // envelope and base64 inflates by 4/3, so a document past this arrives at
    // the caller as a parse failure — which says nothing at all about the
    // document having been too large.
    return failed(
      `Error: the rendered ${format} is ${bytes.byteLength.toLocaleString("en-US")} bytes, over ` +
        `the ${MAX_RENDERED_BYTES.toLocaleString("en-US")} limit for one response. Write it as ` +
        "more than one document.",
      {
        operation: "render_document",
        code: "OUTPUT_TOO_LARGE",
        suggestedFix: "Split the content into more than one document or reduce embedded assets.",
      },
    );
  }

  const checked = await validateRenderedDocument(format, bytes, counts.pages);
  const validation =
    counts.continuations && counts.continuations > 0
      ? {
          ...checked,
          warnings: [
            ...checked.warnings,
            `${counts.continuations} continuation slide(s) were created; review deck density.`,
          ],
        }
      : checked;

  return {
    content: [
      {
        type: "text",
        text:
          `Wrote ${meta.filename} with the ${meta.profile} profile ` +
          `(${bytes.byteLength.toLocaleString("en-US")} bytes${extra}) — ` +
          `${summarise(document)}. The file is attached and the caller delivers it to the user, ` +
          "so describe it rather than offering a link. Package validation passed; visual " +
          "layout was not rendered.",
      },
      {
        type: "resource",
        resource: {
          // A name, not an address: nothing here serves it. The caller reads the
          // last segment to name the file it stores.
          uri: `file:///${encodeURIComponent(meta.filename)}`,
          mimeType: CONTENT_TYPES[format],
          blob: Buffer.from(bytes).toString("base64"),
        },
      },
    ],
    structuredContent: {
      operation: "render_document",
      format,
      profile: meta.profile,
      filename: meta.filename,
      bytes: bytes.byteLength,
      counts,
      validation,
    },
  };
}

/**
 * No config and no headers.
 *
 * Both left with the bucket. `x-document-tenant` existed to partition storage
 * this server no longer owns, and a tenant argument was never an option — a
 * model that can name its own tenant can write into another project's prefix,
 * including one talked into it by a document it read a moment earlier. Removing
 * the storage removed the question.
 */
export async function callTool(name: unknown, args: Record<string, unknown>): Promise<ToolResult> {
  const started = performance.now();
  const result = await dispatch(name, args);
  // One line per call, whatever the outcome — the tool, the format, how long it
  // took and whether it answered. Not the content: what was read or written is
  // the caller's, and `log.ts` says why it never reaches a log line.
  logInfo("tool_call", {
    tool: String(name),
    format: formatOf(name, args),
    ms: elapsedMs(started),
    ok: result.isError !== true,
  });
  return result;
}

/** The format a call was about, for its log line: requested of a writer, or named by a read file. */
function formatOf(name: unknown, args: Record<string, unknown>): string | undefined {
  if (name === "render_document") {
    return typeof args.format === "string" ? args.format : undefined;
  }
  if (name === "read_document") {
    return extensionOf(typeof args.filename === "string" ? args.filename : undefined);
  }
  if (name === "inspect_spreadsheet") {
    return extensionOf(typeof args.filename === "string" ? args.filename : undefined) ?? "xlsx";
  }
  if (name === "render_spreadsheet") {
    return "xlsx";
  }
  return undefined;
}

async function dispatch(name: unknown, args: Record<string, unknown>): Promise<ToolResult> {
  if (name === "read_document") {
    return read(args);
  }
  if (name === "inspect_spreadsheet") {
    return inspectSpreadsheet(args);
  }
  if (name === "render_spreadsheet") {
    return renderSpreadsheet(args);
  }
  if (name === "render_document") {
    return render(args);
  }
  return failed(`Error: unknown tool ${String(name)}.`, {
    operation: String(name),
    code: "UNKNOWN_TOOL",
    suggestedFix: `Use one of ${TOOLS.map((tool) => tool.name).join(", ")}.`,
  });
}
