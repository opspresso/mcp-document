/**
 * The one place a colour or a size is decided.
 *
 * Before this file the four renderers each carried their own literals, and they
 * had already drifted: the table border was `BFBFBF` in three of them and a
 * fourth grey in the PDF, the quote grey differed by a shade, and nothing in any
 * of them said which product had produced the document. A reader could not tell
 * an AgentDure report from any other tool's output.
 *
 * **The palette is AgentDure's own, taken from the console.** The brand ramp is
 * indigo-violet at OKLCH hue 290 — `theme.ts` in the app re-hued an inherited
 * blue rather than re-picking it, so every shade keeps its lightness curve and
 * shade 6/7 stay the ones chosen for white-on-brand contrast. Shade 7 is what
 * this file uses for filled surfaces, because a document is read on paper and on
 * a screen that is not the console's lavender.
 *
 * **What did not survive the translation from screen to page**, and why:
 *
 * - *The lavender page.* The console's body is `#f4f5ff`, not white, and that is
 *   its most recognisable trait. Laid under a whole document it is ink somebody
 *   pays for and a scan artefact in every photocopy, so it is kept for covers and
 *   section slides — the two places a full-bleed field earns its cost — and the
 *   body pages stay white. A hairline of brand under a heading carries the
 *   identity instead.
 * - *The grain.* A tiled `feTurbulence` at a few percent is what stops a flat
 *   lavender field reading as a blank div. Reproducing it here means embedding a
 *   raster in every file for a texture nobody would name if asked. Dropped.
 * - *Pure black.* Never used. The console's shadows are violet-tinted
 *   (`rgba(38, 26, 74, …)`) rather than neutral, and the ink here follows: a
 *   near-black with the same hue in it.
 *
 * **No font is named, in any format.** `write/docx.ts` reached this conclusion
 * on its own — a face named here is a face the reader's machine may not have,
 * and the substitute is then chosen by nobody. AgentDure's own faces make it
 * worse rather than better: Figtree, Chakra Petch and JetBrains Mono carry no
 * Hangul at all, so a Korean document set in them is a document set in whatever
 * the system falls back to. The identity is carried by colour and layout, which
 * survive every substitution. The PDF is the exception it always was: it embeds
 * Nanum Gothic because PDF has no system stack to fall back to.
 */

/* --------------------------------------------------------------- palette */

/**
 * Six-digit hex, no `#`.
 *
 * DOCX and PPTX want it exactly like this, HWPX wants a `#` in front, and the
 * PDF wants three floats — `hashed` and `rgbOf` below do those two conversions
 * so the values themselves are stated once.
 */
export const PALETTE = {
  /** brand-7. Filled surfaces: a table header, a section slide. */
  brand: "6B3DD8",
  /** brand-6, the console's light-mode primary. Rules and accents. */
  brandLight: "805FE9",
  /** brand-8. Links, which need to stay legible as small underlined text. */
  brandDeep: "5B33B8",
  /** brand-0. Zebra rows, code blocks, quote grounds. */
  brandTint: "F4F3FE",
  /** The console's page. Covers and section slides only — never a body page. */
  surfaceTint: "F4F5FF",
  /** Near-black with the brand's hue in it, rather than a neutral grey. */
  ink: "1F1D2B",
  /** Sub-heads, captions, sources, quotes. */
  inkMuted: "6B6880",
  /** Hairlines: table rules, the line under a heading, a horizontal rule. */
  rule: "E3E1EE",
  /** Anything on a brand-filled surface. */
  onBrand: "FFFFFF",
  /** chart-3 and chart-8 from the console, already validated for CVD. */
  positive: "1BAF7A",
  negative: "E34948",
} as const;

export type ColourName = keyof typeof PALETTE;

/**
 * The console's categorical chart palette, validated there for colour-vision
 * deficiency and for contrast on a light surface.
 *
 * Nothing in this repository draws a chart. It is here because a PPTX carries a
 * theme, and that theme is the swatch list a reader sees when they add a shape
 * or a chart to the deck we handed them — leaving Office's defaults there means
 * their first edit is off-brand. Six of the eight fit the theme's six accent
 * slots; `#008300` is dropped for being a second green, and the order is kept
 * otherwise.
 */
export const CHART = [
  "2A78D6",
  "EB6834",
  "1BAF7A",
  "EDA100",
  "E87BA4",
  "4A3AA7",
  "E34948",
  "898781",
] as const;

/** `#RRGGBB`, which is the form HWPX and the OOXML `srgbClr` attribute want. */
export function hashed(name: ColourName): string {
  return `#${PALETTE[name]}`;
}

/**
 * Channels as floats in 0-1, for `pdf-lib`'s `rgb()`.
 *
 * Returned as a plain object rather than a `RGB`: this module is imported by
 * three renderers that have no business loading `pdf-lib`, and the PDF one
 * wraps it in a single call at the point of use.
 */
export function rgbOf(name: ColourName): { r: number; g: number; b: number } {
  const hex = PALETTE[name];
  return {
    r: parseInt(hex.slice(0, 2), 16) / 255,
    g: parseInt(hex.slice(2, 4), 16) / 255,
    b: parseInt(hex.slice(4, 6), 16) / 255,
  };
}

/* ------------------------------------------------------------------ scale */

/**
 * Two scales, because a page and a screen are read from different distances.
 *
 * Both are in **points**, which is the only unit all four formats can be
 * derived from, and the converters below do the deriving.
 *
 * The three document renderers did not agree before this: DOCX set a level 1
 * heading at 16pt, HWPX at 18pt and the PDF at 20pt, and the body was 11pt in
 * two of them and 10pt in the third. The same Markdown produced three documents
 * of visibly different weight, which is exactly what a shared scale is for. The
 * PDF's is the one kept — it is the format whose output was laid out against a
 * real page rather than against a default.
 */
export const DOC = {
  body: 11,
  code: 9.5,
  /** Levels 1 to 6. */
  headings: [20, 17, 15, 13, 12, 11],
  /** Page numbers, sources, anything set below the body. */
  caption: 8.5,
} as const;

export const DECK = {
  coverTitle: 44,
  title: 32,
  body: 18,
  code: 14,
  /** Headings 3 to 6, which stay in a slide's body rather than opening one. */
  subheadings: [22, 20, 18, 18],
  caption: 11,
} as const;

/* -------------------------------------------------------------- converters */

/** DOCX `w:sz`, which counts half-points. */
export function halfPoints(points: number): number {
  return Math.round(points * 2);
}

/** DOCX lengths, and HWPX line spacing: twentieths of a point. */
export function twips(points: number): number {
  return Math.round(points * 20);
}

/**
 * HWPX `height` and DrawingML `sz`, both hundredths of a point.
 *
 * The same number serves as HWPUNIT for a *character* size — HWPUNIT is
 * 1/7200 inch and a point is 1/72, so one point is one hundred of them. It is
 * not the converter for a *length*; `hwpunit` below is, and they agree by
 * arithmetic rather than by coincidence.
 */
export function centiPoints(points: number): number {
  return Math.round(points * 100);
}

/** HWPX lengths: 1/7200 inch. */
export function hwpunit(points: number): number {
  return Math.round(points * 100);
}

/** DrawingML lengths: 914,400 to the inch. */
export function emu(points: number): number {
  return Math.round(points * 12700);
}
