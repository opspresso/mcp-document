/**
 * What a pile of bytes actually is.
 *
 * The order is deliberate: **magic bytes first, then the declared type, then the
 * filename**. A document served as `application/octet-stream` is ordinary — a
 * `.hwp` behind a download endpoint almost always is — so a declared type that
 * disagrees with the file's own header is wrong about the file, not about the
 * header. The filename comes last because it is the only one of the three the
 * caller can be careless with at no cost.
 *
 * Refusals carry a reason and, where one exists, the thing to do instead. A
 * format this cannot read is a fact the model can act on; "unsupported" alone
 * buys another turn spent guessing.
 */

import { listEntries, looksLikeZip } from "./zip.js";

export type Format = "pdf" | "docx" | "hwpx" | "hwp" | "text";

export type Detection = { format: Format } | { format: "unsupported"; reason: string };

/** How much of the head is examined when deciding whether bytes are text. */
const TEXT_SAMPLE_BYTES = 4096;

/** `%PDF` — the four bytes every PDF starts with. */
function looksLikePdf(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 4 &&
    bytes[0] === 0x25 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x44 &&
    bytes[3] === 0x46
  );
}

/**
 * The OLE compound file header. HWP 5.0 is one, and so is every Office binary
 * from the 97-2003 generation — which is why this only says "compound file"
 * and leaves `read/hwp5.ts` to say which one, from the directory it is about to
 * parse anyway.
 */
const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

function looksLikeOle(bytes: Uint8Array): boolean {
  return bytes.length >= 8 && OLE_SIGNATURE.every((byte, index) => bytes[index] === byte);
}

/**
 * HWP 3.0, which is a flat binary rather than a compound file and so is not
 * caught by anything above. It announces itself in ASCII, which is the only
 * convenient thing about it.
 */
const HWP3_SIGNATURE = "HWP Document File V3.00";

function looksLikeHwp3(bytes: Uint8Array): boolean {
  return (
    bytes.length >= HWP3_SIGNATURE.length &&
    Buffer.from(bytes.subarray(0, HWP3_SIGNATURE.length)).toString("latin1") === HWP3_SIGNATURE
  );
}

/**
 * Bytes that read as text in *some* encoding.
 *
 * Judged on control characters rather than on whether UTF-8 decoding succeeds:
 * a EUC-KR page is text, and decoding it as UTF-8 produces a screen of
 * replacement characters. Its bytes are all in 0x80-0xFF, which no control
 * check objects to, so the encoding question is left to the decoder that
 * actually has the charset.
 */
export function looksLikeText(bytes: Uint8Array): boolean {
  const sample = bytes.subarray(0, TEXT_SAMPLE_BYTES);
  if (sample.length === 0) {
    return false;
  }
  let control = 0;
  for (const byte of sample) {
    // A NUL is the one byte no text encoding this could be ever emits — UTF-16
    // would, but a UTF-16 document without a BOM is not something to guess at.
    if (byte === 0) {
      return false;
    }
    if (byte < 0x09 || (byte > 0x0d && byte < 0x20) || byte === 0x7f) {
      control += 1;
    }
  }
  return control / sample.length < 0.05;
}

const BY_MIME_TYPE: Record<string, Format> = {
  "application/pdf": "pdf",
  "application/x-pdf": "pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
  "application/hwp+zip": "hwpx",
  "application/vnd.hancom.hwpx": "hwpx",
  "application/haansofthwpx": "hwpx",
  "application/x-hwp": "hwp",
  "application/vnd.hancom.hwp": "hwp",
  "application/haansofthwp": "hwp",
  "application/json": "text",
  "application/ld+json": "text",
  "application/xml": "text",
  "application/yaml": "text",
  "application/x-yaml": "text",
  "application/x-ndjson": "text",
};

const BY_EXTENSION: Record<string, Format> = {
  pdf: "pdf",
  docx: "docx",
  hwpx: "hwpx",
  hwp: "hwp",
  txt: "text",
  text: "text",
  md: "text",
  markdown: "text",
  csv: "text",
  tsv: "text",
  json: "text",
  jsonl: "text",
  ndjson: "text",
  xml: "text",
  yaml: "text",
  yml: "text",
  log: "text",
  ini: "text",
  toml: "text",
};

/** Formats named so a refusal can say what this is, rather than only that it is not readable. */
const NAMED_REFUSALS: Record<string, string> = {
  doc: "a Word 97-2003 document (.doc)",
  xls: "an Excel 97-2003 workbook (.xls)",
  ppt: "a PowerPoint 97-2003 deck (.ppt)",
  xlsx: "an Excel workbook (.xlsx)",
  pptx: "a PowerPoint deck (.pptx)",
  odt: "an OpenDocument text file (.odt)",
  ods: "an OpenDocument spreadsheet (.ods)",
  odp: "an OpenDocument presentation (.odp)",
  rtf: "a Rich Text Format file (.rtf)",
  epub: "an EPUB book (.epub)",
};

const READS = "This tool reads PDF, DOCX, HWP, HWPX and plain text.";

function refuse(reason: string): Detection {
  return { format: "unsupported", reason };
}

/** Said in one place so every refusal names the same set of formats. */
function cannotRead(what: string): Detection {
  return refuse(`this is ${what}, which this tool cannot read. ${READS}`);
}

/** A web page belongs to the other server, and saying so saves a turn. */
function isWebPage(): Detection {
  return refuse(
    "this is an HTML page, not a document — mcp-url-fetch's fetch_document reads web pages and " +
      "converts them to text",
  );
}

export function extensionOf(filename: string | undefined): string | undefined {
  const match = /\.([A-Za-z0-9]+)$/.exec(filename?.trim() ?? "");
  return match?.[1]?.toLowerCase();
}

/**
 * A zip is a container, so what it *is* depends on what is in it. The names
 * alone decide: every one of these formats declares itself by the presence of a
 * part at a fixed path.
 */
function insideZip(bytes: Uint8Array): Detection {
  const names = new Set(listEntries(bytes).map((entry) => entry.name));
  if (names.has("word/document.xml")) {
    return { format: "docx" };
  }
  if ([...names].some((name) => /^Contents\/section\d+\.xml$/.test(name))) {
    return { format: "hwpx" };
  }
  if (names.has("xl/workbook.xml")) {
    return cannotRead("an Excel workbook (.xlsx)");
  }
  if (names.has("ppt/presentation.xml")) {
    return cannotRead("a PowerPoint deck (.pptx)");
  }
  if (names.has("content.xml") && names.has("mimetype")) {
    return cannotRead("an OpenDocument file");
  }
  if (names.has("META-INF/container.xml")) {
    return cannotRead("an EPUB book (.epub)");
  }
  return refuse(`this is a zip archive, but not a document format this tool reads. ${READS}`);
}

export function detect(
  bytes: Uint8Array,
  mimeType: string,
  filename: string | undefined,
): Detection {
  if (bytes.length === 0) {
    return refuse("the document is empty");
  }
  if (looksLikePdf(bytes)) {
    return { format: "pdf" };
  }
  if (looksLikeZip(bytes)) {
    return insideZip(bytes);
  }
  if (looksLikeHwp3(bytes)) {
    // Its own case because the extension branch below would otherwise report a
    // `.hwp` whose contents are not HWP — true, and useless. This one names the
    // version and the way out of it.
    return refuse(
      "this is an HWP 3.0 document, a format from before 한글 moved to compound files. This tool " +
        "reads HWP 5.x — open it in 한글 and save it as .hwpx",
    );
  }
  if (looksLikeOle(bytes)) {
    // Which compound file it is comes from its directory, which the HWP reader
    // parses to do its own job. Deciding here would mean parsing it twice and
    // disagreeing about it once.
    return { format: "hwp" };
  }

  if (mimeType.startsWith("image/")) {
    return refuse(
      `this is ${mimeType}, an image rather than a document — mcp-url-fetch's fetch_image ` +
        "returns the picture itself",
    );
  }
  if (mimeType === "text/html" || mimeType === "application/xhtml+xml") {
    return isWebPage();
  }
  const declared = BY_MIME_TYPE[mimeType];
  if (declared) {
    // Only the text branch can be honoured on the header alone: every other
    // format here has magic bytes, and not finding them means the body is not
    // what the header claimed.
    if (declared === "text") {
      return { format: "text" };
    }
    return refuse(
      `this was served as ${mimeType}, but its contents are not ${declared.toUpperCase()}`,
    );
  }

  const extension = extensionOf(filename);
  if (extension === "html" || extension === "htm") {
    return isWebPage();
  }
  const named = extension ? NAMED_REFUSALS[extension] : undefined;
  if (named) {
    return cannotRead(named);
  }
  const byExtension = extension ? BY_EXTENSION[extension] : undefined;
  if (byExtension === "text") {
    return { format: "text" };
  }
  if (byExtension) {
    return refuse(
      `this is named .${extension}, but its contents are not ${byExtension.toUpperCase()}`,
    );
  }

  if (mimeType.startsWith("text/") || looksLikeText(bytes)) {
    return { format: "text" };
  }
  return refuse(
    `this is not a document format this tool recognises${mimeType ? ` (served as ${mimeType})` : ""}. ${READS}`,
  );
}
