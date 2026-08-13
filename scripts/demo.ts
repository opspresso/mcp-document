/**
 * The four files the tests cannot see.
 *
 * What tests cannot cover is what a document *looks like* — README says so —
 * and every format has a reader that either opens a file or does not. This
 * renders the two demo sources into all four outputs under `build/`, so a
 * change to any renderer can be looked at before it is trusted:
 *
 *     build/demo.pptx   demo-deck.md — every slide archetype
 *     build/demo.docx   demo-doc.md  — every document device
 *     build/demo.pdf    demo-doc.md  — the same report, laid out here
 *     build/demo.hwpx   demo-doc.md  — the same report, for 한글
 *
 * Run it with `npm run demo`. The PDF and HWPX take the same source as the
 * DOCX on purpose: one report in three formats is what shows a drifted colour
 * or a diverged heading scale, which is exactly what `write/theme.ts` exists
 * to prevent. Assets go only to the formats that embed them — in the PDF and
 * HWPX the figure renders as the link it is documented to be.
 */

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { parseMarkdown } from "../src/markdown.js";
import { renderDocx } from "../src/write/docx.js";
import { renderHwpx } from "../src/write/hwpx.js";
import { renderPdf } from "../src/write/pdf.js";
import { renderPptx } from "../src/write/pptx/index.js";
import { flatPng } from "./flat-png.js";

const read = (name: string): string => readFileSync(new URL(name, import.meta.url), "utf8");

const deck = parseMarkdown(read("./demo-deck.md"));
const report = parseMarkdown(read("./demo-doc.md"));

const created = new Date();
const assets = {
  // The console's lavender, so even the placeholder is on palette.
  "architecture.png": { mimeType: "image/png" as const, bytes: flatPng(960, 540, [0xf4, 0xf3, 0xfe]) },
};

const DECK_TITLE = "AI Agent Platform";
const REPORT_TITLE = "AI Agent Platform 도입 보고서";

mkdirSync("build", { recursive: true });

const pptx = renderPptx(deck, { title: DECK_TITLE, created: created.toISOString(), assets });
writeFileSync("build/demo.pptx", pptx.bytes);
console.log(`build/demo.pptx — ${pptx.slides} slides, ${pptx.bytes.byteLength.toLocaleString("en-US")} bytes`);

const docx = renderDocx(report, { title: REPORT_TITLE, created: created.toISOString(), assets });
writeFileSync("build/demo.docx", docx);
console.log(`build/demo.docx — ${docx.byteLength.toLocaleString("en-US")} bytes`);

const pdf = await renderPdf(report, { title: REPORT_TITLE, created });
writeFileSync("build/demo.pdf", pdf.bytes);
console.log(`build/demo.pdf — ${pdf.pages} pages, ${pdf.bytes.byteLength.toLocaleString("en-US")} bytes`);

const hwpx = renderHwpx(report, { title: REPORT_TITLE, created: created.toISOString() });
writeFileSync("build/demo.hwpx", hwpx);
console.log(`build/demo.hwpx — ${hwpx.byteLength.toLocaleString("en-US")} bytes`);

writeFileSync("build/demo-deck.md", read("./demo-deck.md"));
writeFileSync("build/demo-doc.md", read("./demo-doc.md"));
