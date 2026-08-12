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

import { DocumentError } from "./errors.js";
import { asUntrustedContent, MAX_MARKDOWN_CHARS, MAX_RENDERED_BYTES } from "./limits.js";
import { parseMarkdown, type Block, type MarkdownDocument } from "./markdown.js";
import { readDocument } from "./read/document.js";
import { loadSource } from "./source.js";
import { safeFilename } from "./filename.js";
import { SERVER_VERSION } from "./version.js";
import { renderDocx } from "./write/docx.js";
import { renderHwpx } from "./write/hwpx.js";
import { renderPdf } from "./write/pdf.js";
import { renderPptx } from "./write/pptx.js";

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
  isError?: true;
}

export const FORMATS = ["docx", "pdf", "hwpx", "pptx"] as const;
export type Format = (typeof FORMATS)[number];

export const CONTENT_TYPES: Record<Format, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pdf: "application/pdf",
  hwpx: "application/hwp+zip",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

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
      "never formulas, one sheet per heading.",
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
    name: "render_document",
    description:
      "Write a document and return the file. Takes Markdown and produces DOCX " +
      "(Word), PDF, HWPX (한글) or PPTX (PowerPoint) — pick the one the reader will open. Use " +
      "this to deliver something a person will read or share: a report, a summary, meeting " +
      "notes, a proposal, a slide deck. Supports headings, paragraphs, bold, italic, inline " +
      "code, links, bullet and numbered lists, tables, block quotes, fenced code blocks and " +
      "horizontal rules. A table column is set flush right with `---:` in the divider row and " +
      "centred with `:---:` — set columns of numbers right, or their digits do not line up. " +
      "Two lines with no blank line between them are one paragraph, as in " +
      "Markdown. Images are not embedded; an image becomes a link. In pptx every level 1 or 2 " +
      "heading starts a new slide and becomes its title, so write one heading per slide and " +
      "keep what follows it short — a few bullets or a small table; deeper headings stay in " +
      "the body, and anything that does not fit continues on the next slide. Returns the file " +
      "itself; the caller delivers it to the user.",
    inputSchema: {
      type: "object",
      properties: {
        format: {
          type: "string",
          enum: [...FORMATS],
          description:
            "docx for Word, pdf to be read as-is, hwpx for 한글, pptx for a slide deck.",
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
  const content = typeof args.content === "string" && args.content ? args.content : undefined;
  const filename = typeof args.filename === "string" && args.filename ? args.filename : undefined;
  try {
    const source = loadSource({ content, filename });
    const result = await readDocument(source);
    return ok(asUntrustedContent(source.label, result.text, result.note));
  } catch (error) {
    // The caller is told; without this line the operator is not, and a format
    // that started failing has no evidence behind it anywhere. A filename is
    // safe to log — unlike a URL, it carries no capability.
    console.warn(`read_document failed: ${filename ?? "(unnamed)"} — ${describe(error)}`);
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

async function render(args: Record<string, unknown>): Promise<ToolResult> {
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

    return await renderTo(format, document, { title, created, filename });
  } catch (error) {
    console.warn(`render_document failed: ${format} — ${describe(error)}`);
    return failed(
      `Error: could not write the document — ${
        error instanceof DocumentError ? error.message : describe(error)
      }`,
    );
  }
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
  meta: { title: string; created: Date; filename: string },
): Promise<ToolResult> {
  let bytes: Uint8Array;
  let extra = "";
  if (format === "docx") {
    bytes = renderDocx(document, { title: meta.title, created: meta.created.toISOString() });
  } else if (format === "hwpx") {
    bytes = renderHwpx(document, {
      title: meta.title,
      created: meta.created.toISOString(),
      application: SERVER_VERSION,
    });
  } else if (format === "pptx") {
    const rendered = renderPptx(document, {
      title: meta.title,
      created: meta.created.toISOString(),
    });
    bytes = rendered.bytes;
    extra = `, ${rendered.slides} slide${rendered.slides === 1 ? "" : "s"}`;
  } else {
    const rendered = await renderPdf(document, { title: meta.title, created: meta.created });
    bytes = rendered.bytes;
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
    );
  }

  return {
    content: [
      {
        type: "text",
        text:
          `Wrote ${meta.filename} (${bytes.byteLength.toLocaleString("en-US")} bytes${extra}) — ` +
          `${summarise(document)}. The file is attached and the caller delivers it to the user, ` +
          "so describe it rather than offering a link.",
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
  if (name === "read_document") {
    return read(args);
  }
  if (name === "render_document") {
    return render(args);
  }
  return failed(`Error: unknown tool ${String(name)}.`);
}
