/**
 * The two tools, and everything between a JSON-RPC `tools/call` and the answer.
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

import type { Config } from "./config.js";
import { DocumentError } from "./errors.js";
import { asUntrustedContent, MAX_MARKDOWN_CHARS } from "./limits.js";
import { parseMarkdown, type Block, type MarkdownDocument } from "./markdown.js";
import { readDocument } from "./read/document.js";
import { loadSource } from "./source.js";
import { safeFilename, storeDocument } from "./store.js";
import { parseTenant, TENANT_HEADER } from "./tenant.js";
import { SERVER_VERSION } from "./version.js";
import { renderDocx } from "./write/docx.js";
import { renderHwpx } from "./write/hwpx.js";
import { renderPdf } from "./write/pdf.js";

export interface ToolResult {
  content: { type: "text"; text: string }[];
  isError?: true;
}

export const FORMATS = ["docx", "pdf", "hwpx"] as const;
export type Format = (typeof FORMATS)[number];

const CONTENT_TYPES: Record<Format, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pdf: "application/pdf",
  hwpx: "application/hwp+zip",
};

export const TOOLS = [
  {
    name: "read_document",
    description:
      "Read a document and return its text. Handles PDF, DOCX (Word), HWP and HWPX (한글), and " +
      "plain text formats — Markdown, CSV, TSV, JSON, XML. Give it either `url` (an http(s) " +
      "address) or `content` (the file's bytes as base64), not both. Use this when you have a " +
      "document you cannot otherwise open: a report, a spec, a contract, a data file. Returns " +
      "the contents, not a summary. For a web page use mcp-url-fetch's fetch_document instead, " +
      "and for an image use its fetch_image. A scanned PDF with no text layer comes back as an " +
      "error rather than as an empty document.",
    inputSchema: {
      type: "object",
      properties: {
        url: { type: "string", description: "Absolute http(s) URL of the document." },
        content: { type: "string", description: "The document's bytes, base64-encoded." },
        filename: {
          type: "string",
          description:
            "The document's filename, e.g. 'report.hwp'. Optional, and only a hint: it helps " +
            "identify a format whose bytes are ambiguous.",
        },
      },
    },
  },
  {
    name: "write_document",
    description:
      "Write a document and return a link to download it. Takes Markdown and produces DOCX " +
      "(Word), PDF, or HWPX (한글) — pick the one the reader will open. Use this to deliver " +
      "something a person will read or share: a report, a summary, meeting notes, a proposal. " +
      "Supports headings, paragraphs, bold, italic, inline code, links, bullet and numbered " +
      "lists, tables, block quotes, fenced code blocks and horizontal rules. Two lines with no " +
      "blank line between them are one paragraph, as in Markdown. Images are not embedded; an " +
      "image becomes a link. Returns a download URL that expires, so pass it on rather than " +
      "storing it.",
    inputSchema: {
      type: "object",
      properties: {
        format: {
          type: "string",
          enum: [...FORMATS],
          description: "docx for Word, pdf to be read as-is, hwpx for 한글.",
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
      },
      required: ["format", "content"],
    },
  },
] as const;

function ok(text: string): ToolResult {
  return { content: [{ type: "text", text }] };
}

function failed(text: string): ToolResult {
  return { content: [{ type: "text", text }], isError: true };
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
  const url = typeof args.url === "string" && args.url ? args.url : undefined;
  const content = typeof args.content === "string" && args.content ? args.content : undefined;
  const filename = typeof args.filename === "string" && args.filename ? args.filename : undefined;
  try {
    const source = await loadSource({ url, content, filename });
    const result = await readDocument(source);
    return ok(asUntrustedContent(source.label, result.text, result.note));
  } catch (error) {
    // The model is told; without this the operator is not, and a host that
    // started failing on Tuesday has no evidence behind it anywhere.
    console.warn(`read_document failed: ${originOf(url)} — ${describe(error)}`);
    return failed(
      `Error: could not read the document — ${
        error instanceof DocumentError ? error.message : describe(error)
      }`,
    );
  }
}

/** What the document turned out to be, for the line the caller reads. */
export function summarise(document: MarkdownDocument): string {
  const count = (kind: Block["kind"]): number =>
    document.blocks.filter((block) => block.kind === kind).length;
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

async function write(
  config: Config,
  args: Record<string, unknown>,
  headers: Record<string, string | string[] | undefined>,
): Promise<ToolResult> {
  const requested = args.format;
  if (typeof requested !== "string" || !FORMATS.includes(requested as Format)) {
    return failed(`Error: \`format\` must be one of ${FORMATS.join(", ")}.`);
  }
  const format = requested as Format;
  const markdown = args.content;
  if (typeof markdown !== "string" || markdown.trim() === "") {
    return failed("Error: `content` is required, as Markdown.");
  }
  if (markdown.length > MAX_MARKDOWN_CHARS) {
    return failed(
      `Error: \`content\` is ${markdown.length.toLocaleString("en-US")} characters, over the ` +
        `${MAX_MARKDOWN_CHARS.toLocaleString("en-US")} limit. Write it as more than one document.`,
    );
  }

  let tenant: string;
  try {
    tenant = parseTenant(headers[TENANT_HEADER]);
  } catch (error) {
    return failed(`Error: ${describe(error)}`);
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

    let bytes: Uint8Array;
    let extra = "";
    if (format === "docx") {
      bytes = renderDocx(document, { title, created: created.toISOString() });
    } else if (format === "hwpx") {
      bytes = renderHwpx(document, {
        title,
        created: created.toISOString(),
        application: SERVER_VERSION,
      });
    } else {
      const rendered = await renderPdf(document, { title, created });
      bytes = rendered.bytes;
      extra = `, ${rendered.pages} page${rendered.pages === 1 ? "" : "s"}`;
    }

    const stored = await storeDocument(config, tenant, filename, bytes, CONTENT_TYPES[format]);
    return ok(
      `Wrote ${filename} (${stored.bytes.toLocaleString("en-US")} bytes${extra}) — ` +
        `${summarise(document)}.\n\n${stored.url}\n\n` +
        `That link is valid until ${stored.expiresAt.toISOString()}. Give it to the user; it is ` +
        `the only way they can get the file.`,
    );
  } catch (error) {
    console.warn(`write_document failed: ${format} for ${tenant} — ${describe(error)}`);
    return failed(
      `Error: could not write the document — ${
        error instanceof DocumentError ? error.message : describe(error)
      }`,
    );
  }
}

export async function callTool(
  config: Config,
  name: unknown,
  args: Record<string, unknown>,
  headers: Record<string, string | string[] | undefined>,
): Promise<ToolResult> {
  if (name === "read_document") {
    return read(args);
  }
  if (name === "write_document") {
    return write(config, args, headers);
  }
  return failed(`Error: unknown tool ${String(name)}.`);
}
