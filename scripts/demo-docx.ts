/**
 * The report the tests cannot see.
 *
 * The DOCX twin of `demo-pptx.ts`: renders `demo-doc.md` — a report that uses
 * every document device — to `build/demo.docx`, to be opened in Word or Pages
 * before a change to the renderer is trusted. Run it with `npm run demo:docx`.
 */

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { parseMarkdown } from "../src/markdown.js";
import { renderDocx } from "../src/write/docx.js";
import { flatPng } from "./flat-png.js";

const source = new URL("./demo-doc.md", import.meta.url);
const markdown = readFileSync(source, "utf8");
const bytes = renderDocx(parseMarkdown(markdown), {
  title: "AI Agent Platform 도입 보고서",
  created: new Date().toISOString(),
  assets: {
    // The console's lavender, so even the placeholder is on palette.
    "architecture.png": { mimeType: "image/png", bytes: flatPng(960, 540, [0xf4, 0xf3, 0xfe]) },
  },
});

mkdirSync("build", { recursive: true });
writeFileSync("build/demo.docx", bytes);
writeFileSync("build/demo-doc.md", markdown);
console.log(`build/demo.docx — ${bytes.byteLength.toLocaleString("en-US")} bytes`);
