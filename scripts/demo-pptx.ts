/**
 * The deck the tests cannot see.
 *
 * What tests cannot cover is what a document *looks like* — README says so, and
 * it matters most for pptx, whose line counting is an estimate. This renders
 * `demo-deck.md` to `build/demo.pptx` so a change to the renderer can be opened
 * in PowerPoint or Keynote before it is trusted. Run it with `npm run demo:pptx`.
 *
 * The one asset the deck references is generated below rather than committed:
 * a flat PNG is thirty lines of code, and a binary in the repository is a file
 * nobody can review.
 */

import { deflateSync } from "node:zlib";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { parseMarkdown } from "../src/markdown.js";
import { renderPptx } from "../src/write/pptx/index.js";

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  out.set([...type].map((c) => c.charCodeAt(0)), 4);
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
  return out;
}

/** A flat-colour PNG, enough of a picture to see the image slide's layout. */
function flatPng(width: number, height: number, rgb: [number, number, number]): Uint8Array {
  const ihdr = new Uint8Array(13);
  const view = new DataView(ihdr.buffer);
  view.setUint32(0, width);
  view.setUint32(4, height);
  ihdr.set([8, 2, 0, 0, 0], 8);
  const raw = new Uint8Array(height * (1 + width * 3));
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      raw.set(rgb, y * (1 + width * 3) + 1 + x * 3);
    }
  }
  const signature = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const parts = [
    signature,
    chunk("IHDR", ihdr),
    chunk("IDAT", new Uint8Array(deflateSync(raw))),
    chunk("IEND", new Uint8Array(0)),
  ];
  const png = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let at = 0;
  for (const part of parts) {
    png.set(part, at);
    at += part.length;
  }
  return png;
}

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
