/**
 * A document AST to PPTX.
 *
 * The fourth renderer, and the first whose output is not a page of prose — so it
 * is a package rather than a file. The pipeline is the spec of the split:
 *
 *     MarkdownDocument → planner.ts → Presentation → render.ts → slide XML
 *                                                    package.ts → everything else
 *
 * `planner.ts` decides what each slide *is* (its archetype, and where one ends),
 * `render.ts` decides what shapes say that, `package.ts` carries the scaffolding
 * PowerPoint refuses to open a file without, and `layout.ts` is the geometry
 * both sides measure against. This module only assembles the zip.
 */

import { buildZip } from "../../zip.js";
import type { MarkdownDocument } from "../../markdown.js";
import { part } from "./ooxml.js";
import {
  LAYOUT_COUNT,
  appPropertiesXml,
  contentTypesXml,
  corePropertiesXml,
  packageRelsXml,
  presPropsXml,
  presentationRelsXml,
  presentationXml,
  slideLayoutRelsXml,
  slideLayoutXml,
  slideMasterRelsXml,
  slideMasterXml,
  slideRelsXml,
  tableStylesXml,
  themeXml,
  viewPropsXml,
  type LayoutIndex,
} from "./package.js";
import { plan } from "./planner.js";
import { Renderer } from "./render.js";
import type { Slide } from "./types.js";

/**
 * Which layout part carries each archetype's design.
 *
 * The recognised archetypes — cards, metrics, quote, comparison — share the
 * content layout: same ground, same footer, same title chrome. Their geometry
 * is per-slide shape work, not layout furniture.
 */
const LAYOUT_OF: Record<Slide["type"], LayoutIndex> = {
  cover: 1,
  content: 2,
  cards: 2,
  metrics: 2,
  quote: 2,
  comparison: 2,
  section: 3,
  closing: 4,
};

export interface PptxOptions {
  title: string;
  /** ISO 8601, passed in so the bytes are a function of the input alone. */
  created: string;
}

export interface RenderedPptx {
  bytes: Uint8Array;
  /** How many slides the Markdown turned into, which is what the caller reports. */
  slides: number;
}

export function renderPptx(document: MarkdownDocument, options: PptxOptions): RenderedPptx {
  const { slides } = plan(document);
  const renderer = new Renderer();
  const parts: Record<string, Uint8Array> = {
    "[Content_Types].xml": part(contentTypesXml(slides.length)),
    "_rels/.rels": part(packageRelsXml()),
    "docProps/core.xml": part(corePropertiesXml(options.title, options.created)),
    "docProps/app.xml": part(appPropertiesXml()),
    "ppt/presentation.xml": part(presentationXml(slides.length)),
    "ppt/_rels/presentation.xml.rels": part(presentationRelsXml(slides.length)),
    "ppt/theme/theme1.xml": part(themeXml()),
    "ppt/presProps.xml": part(presPropsXml()),
    "ppt/viewProps.xml": part(viewPropsXml()),
    "ppt/tableStyles.xml": part(tableStylesXml()),
    "ppt/slideMasters/slideMaster1.xml": part(slideMasterXml()),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": part(slideMasterRelsXml()),
  };

  for (let layout = 1; layout <= LAYOUT_COUNT; layout += 1) {
    parts[`ppt/slideLayouts/slideLayout${layout}.xml`] = part(
      slideLayoutXml(layout as LayoutIndex, options.title),
    );
    parts[`ppt/slideLayouts/_rels/slideLayout${layout}.xml.rels`] = part(slideLayoutRelsXml());
  }

  slides.forEach((slide, index) => {
    const rendered = renderer.slide(slide, index);
    parts[`ppt/slides/slide${index + 1}.xml`] = part(rendered.xml);
    parts[`ppt/slides/_rels/slide${index + 1}.xml.rels`] = part(
      slideRelsXml(LAYOUT_OF[slide.type], rendered.links),
    );
  });

  return { bytes: buildZip(parts), slides: slides.length };
}
