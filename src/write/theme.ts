/**
 * The one place a colour or a size is decided.
 *
 * Before this file the four renderers each carried their own literals, and they
 * had already drifted: the table border was `BFBFBF` in three of them and a
 * fourth grey in the PDF, the quote grey differed by a shade, and a level 1
 * heading was three different sizes in three formats.
 *
 * **The palette is deliberately nobody's brand.** It used to be AgentDure's
 * indigo-violet, taken from the console — and a document handed to a customer,
 * a partner or a public office should not arrive dressed in its tooling's
 * colours. What replaced it is the *average* of what professional documents
 * already look like: a restrained corporate blue for structure and emphasis,
 * neutral near-black ink, quiet grey rules. Blue because it is the one hue
 * every reader has seen a thousand reports in — it signals "document", not
 * "product". The result should look designed and belong to no one.
 *
 * What that means concretely:
 *
 * - *White pages.* A tinted ground is kept for the two surfaces that can carry
 *   a full-bleed field for free — a cover slide, a section divider — and body
 *   pages stay white: a tint under a whole document is ink somebody pays for
 *   and an artefact in every photocopy.
 * - *No decoration for its own sake.* The identity is hierarchy, spacing and
 *   a few hairlines; every colour below earns a WCAG ratio in `theme.test.ts`
 *   or it does not ship.
 * - *Pure black is still never used* — near-black reads as ink, full black on
 *   a screen glares.
 *
 * **No font is named, in any format.** `write/docx.ts` reached this conclusion
 * on its own — a face named here is a face the reader's machine may not have,
 * and the substitute is then chosen by nobody. The identity is carried by
 * colour and layout, which survive every substitution. The PDF is the
 * exception it always was: it embeds Nanum Gothic because PDF has no system
 * stack to fall back to.
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
  /** Deep corporate blue. Filled surfaces: a table header, a section slide. */
  brand: "1F4E79",
  /** The lighter working blue. Rules, accents, quote bars, chapter ordinals. */
  brandLight: "4472C4",
  /** Links, which need to stay legible as small underlined text. */
  brandDeep: "0563C1",
  /** Barely-blue ground. Zebra rows, code blocks, cards. */
  brandTint: "EEF3F9",
  /** Covers and section slides only — never a body page. */
  surfaceTint: "F4F6F9",
  /** Neutral near-black. */
  ink: "212529",
  /** Sub-heads, captions, sources, quotes. */
  inkMuted: "595959",
  /** Hairlines: table rules, the line under a heading, a horizontal rule. */
  rule: "D9DEE5",
  /** Anything on a brand-filled surface. */
  onBrand: "FFFFFF",
  /** Universal signal colours, validated for colour-vision deficiency. */
  positive: "1BAF7A",
  negative: "E34948",
} as const;

export type ColourName = keyof typeof PALETTE;

/**
 * A categorical chart palette, validated for colour-vision
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
  /** The cover's title and its second line, a page read at arm's length. */
  coverTitle: 30,
  subtitle: 13,
  /** A chapter's "01" — a mark at the top of the page, not text. */
  ordinal: 36,
  /** A key figure in a metrics strip: the biggest thing on a page of prose. */
  metric: 26,
} as const;

export const DECK = {
  coverTitle: 44,
  /** The cover's second line, under the title and quieter than it. */
  subtitle: 20,
  title: 32,
  /** A section divider's title, sitting alone on a brand field. */
  sectionTitle: 40,
  /** The divider's "01" — a numeral this large is a mark, not text. */
  ordinal: 66,
  /** The closing slide's one line, between a title and a cover in weight. */
  closingTitle: 36,
  body: 18,
  code: 14,
  /** Headings 3 to 6, which stay in a slide's body rather than opening one. */
  subheadings: [22, 20, 18, 18],
  caption: 11,
  /** A card's heading, and the line or two under it. */
  cardTitle: 18,
  cardBody: 14,
  /** A metric's number — the biggest thing on its slide bar the title. */
  metric: 44,
  metricLabel: 13,
  /** A pulled quote, set alone: bigger than body, smaller than a title. */
  quote: 24,
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
