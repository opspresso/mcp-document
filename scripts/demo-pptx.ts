/**
 * The deck the tests cannot see.
 *
 * What tests cannot cover is what a document *looks like* — README says so, and
 * it matters most for pptx, whose line counting is an estimate. This renders
 * `demo-deck.md` to `build/demo.pptx` so a change to the renderer can be opened
 * in PowerPoint or Keynote before it is trusted. Run it with `npm run demo:pptx`.
 */

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { parseMarkdown } from "../src/markdown.js";
import { renderPptx } from "../src/write/pptx/index.js";
import { flatPng } from "./flat-png.js";

const source = new URL("./demo-deck.md", import.meta.url);
const markdown = readFileSync(source, "utf8");
const { bytes, slides } = renderPptx(parseMarkdown(markdown), {
  title: "AI Agent Platform",
  created: new Date().toISOString(),
  assets: {
    // The console's lavender, so even the placeholder is on palette.
    "architecture.png": { mimeType: "image/png", bytes: flatPng(960, 540, [0xf4, 0xf3, 0xfe]) },
  },
});

mkdirSync("build", { recursive: true });
writeFileSync("build/demo.pptx", bytes);
writeFileSync("build/demo-deck.md", markdown);
console.log(`build/demo.pptx — ${slides} slides, ${bytes.byteLength.toLocaleString("en-US")} bytes`);
