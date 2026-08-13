/**
 * The parts of a PPTX that are not slides.
 *
 * A deck needs more scaffolding than a document — PowerPoint refuses a file
 * with no theme or no slide master — so the master, the layouts and the theme
 * are fixed constants here, and only the slides are a function of the input.
 * All of the "empty" parts (presProps, viewProps, tableStyles) are carried
 * because their absence is a package shape no native file has ever had, and
 * Windows PowerPoint reads absence as damage where the Mac build shrugs.
 */

import { escapeXml } from "../../xml.js";
import { CHART, PALETTE } from "../theme.js";
import { PRODUCER } from "../../version.js";
import {
  BODY_BOX,
  BODY_SIZE,
  CONTENT_WIDTH,
  COVER_BODY_BOX,
  COVER_TITLE_BOX,
  SIDE_MARGIN,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  TITLE_BOX,
  TITLE_SIZE,
} from "./layout.js";
import {
  A,
  HYPERLINK_TYPE,
  P,
  R,
  RELATIONSHIPS,
  SLIDE_LAYOUT_TYPE,
  SLIDE_MASTER_TYPE,
  SLIDE_TYPE,
  TABLE_STYLE_ID,
  THEME_TYPE,
  relationship,
  relationships,
} from "./ooxml.js";

export function contentTypesXml(slides: number): string {
  const override = (path: string, type: string): string =>
    `<Override PartName="${path}" ContentType="${type}"/>`;
  const presentationml = "application/vnd.openxmlformats-officedocument.presentationml";
  return (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    override("/ppt/presentation.xml", `${presentationml}.presentation.main+xml`) +
    override("/ppt/slideMasters/slideMaster1.xml", `${presentationml}.slideMaster+xml`) +
    override("/ppt/slideLayouts/slideLayout1.xml", `${presentationml}.slideLayout+xml`) +
    override("/ppt/slideLayouts/slideLayout2.xml", `${presentationml}.slideLayout+xml`) +
    Array.from({ length: slides }, (_, index) =>
      override(`/ppt/slides/slide${index + 1}.xml`, `${presentationml}.slide+xml`),
    ).join("") +
    override(
      "/ppt/theme/theme1.xml",
      "application/vnd.openxmlformats-officedocument.theme+xml",
    ) +
    override("/ppt/presProps.xml", `${presentationml}.presProps+xml`) +
    override("/ppt/viewProps.xml", `${presentationml}.viewProps+xml`) +
    override("/ppt/tableStyles.xml", `${presentationml}.tableStyles+xml`) +
    override(
      "/docProps/core.xml",
      "application/vnd.openxmlformats-package.core-properties+xml",
    ) +
    override(
      "/docProps/app.xml",
      "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    ) +
    "</Types>"
  );
}

/**
 * `docProps/app.xml`, whose only job here is to say what wrote the file.
 *
 * OOXML keeps the producer separate from the core properties: `dc:title` and the
 * dates are the document's, `<Application>` is the tool's. Word fills in a dozen
 * more fields — word counts, template names — and none of them are things this
 * renderer knows or a reader needs. When somebody turns up with a file that
 * renders oddly, this is the line that says which release made it.
 */
export function appPropertiesXml(): string {
  return (
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    `<Application>${escapeXml(PRODUCER)}</Application>` +
    "</Properties>"
  );
}

export function packageRelsXml(): string {
  return relationships([
    relationship("rId1", `${R}/officeDocument`, "ppt/presentation.xml"),
    relationship("rId2", `${RELATIONSHIPS}/metadata/core-properties`, "docProps/core.xml"),
    relationship("rId3", `${R}/extended-properties`, "docProps/app.xml"),
  ]);
}

export function corePropertiesXml(title: string, created: string): string {
  return (
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${escapeXml(title)}</dc:title>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${created}</dcterms:modified>` +
    "</cp:coreProperties>"
  );
}

export function presPropsXml(): string {
  return `<p:presentationPr xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"/>`;
}

export function viewPropsXml(): string {
  return `<p:viewPr xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"/>`;
}

export function tableStylesXml(): string {
  return `<a:tblStyleLst xmlns:a="${A}" def="${TABLE_STYLE_ID}"/>`;
}

export function presentationXml(slides: number): string {
  return (
    `<p:presentation xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" saveSubsetFonts="1">` +
    '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
    "<p:sldIdLst>" +
    // Slide ids start at 256, which the format requires, and the relationship
    // ids continue past the master's rId1.
    Array.from(
      { length: slides },
      (_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`,
    ).join("") +
    "</p:sldIdLst>" +
    `<p:sldSz cx="${SLIDE_WIDTH}" cy="${SLIDE_HEIGHT}"/>` +
    '<p:notesSz cx="6858000" cy="9144000"/>' +
    "</p:presentation>"
  );
}

export function presentationRelsXml(slides: number): string {
  // The theme is reachable through the master, but native files relate it from
  // the presentation part as well, and the other three live only here.
  const next = slides + 2;
  return relationships([
    relationship("rId1", SLIDE_MASTER_TYPE, "slideMasters/slideMaster1.xml"),
    ...Array.from({ length: slides }, (_, index) =>
      relationship(`rId${index + 2}`, SLIDE_TYPE, `slides/slide${index + 1}.xml`),
    ),
    relationship(`rId${next}`, THEME_TYPE, "theme/theme1.xml"),
    relationship(`rId${next + 1}`, `${R}/presProps`, "presProps.xml"),
    relationship(`rId${next + 2}`, `${R}/viewProps`, "viewProps.xml"),
    relationship(`rId${next + 3}`, `${R}/tableStyles`, "tableStyles.xml"),
  ]);
}

/**
 * The theme, which is not optional.
 *
 * PowerPoint refuses a package whose master has no theme behind it, and the
 * schema wants three entries in each of the four format lists whether or not
 * anything refers to them. No east-Asian face is named, for the reason
 * `write/docx.ts` names none: the substitute PowerPoint picks on the reader's
 * machine beats one picked here from a font that may not be installed.
 */
export function themeXml(): string {
  const accents = CHART.slice(0, 6);
  const line = (width: number): string =>
    `<a:ln w="${width}" cap="flat" cmpd="sng" algn="ctr">` +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>';
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  return (
    `<a:theme xmlns:a="${A}" name="Office">` +
    "<a:themeElements>" +
    '<a:clrScheme name="Office">' +
    `<a:dk1><a:sysClr val="windowText" lastClr="${PALETTE.ink}"/></a:dk1>` +
    `<a:lt1><a:sysClr val="window" lastClr="${PALETTE.onBrand}"/></a:lt1>` +
    `<a:dk2><a:srgbClr val="${PALETTE.brand}"/></a:dk2>` +
    `<a:lt2><a:srgbClr val="${PALETTE.brandTint}"/></a:lt2>` +
    accents
      .map((colour, index) => `<a:accent${index + 1}><a:srgbClr val="${colour}"/></a:accent${index + 1}>`)
      .join("") +
    `<a:hlink><a:srgbClr val="${PALETTE.brandDeep}"/></a:hlink>` +
    `<a:folHlink><a:srgbClr val="${PALETTE.brandDeep}"/></a:folHlink>` +
    "</a:clrScheme>" +
    '<a:fontScheme name="Office">' +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    "</a:fontScheme>" +
    '<a:fmtScheme name="Office">' +
    `<a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst>` +
    `<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
    "<a:effectStyleLst>" +
    "<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3) +
    "</a:effectStyleLst>" +
    `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst>` +
    "</a:fmtScheme>" +
    "</a:themeElements>" +
    "</a:theme>"
  );
}

/** An empty placeholder in a layout, which the slide's own shape inherits its slot from. */
function layoutPlaceholder(id: number, name: string, placeholder: string, box: { y: number; height: number }): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/>` +
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    `<p:nvPr>${placeholder}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${SIDE_MARGIN}" y="${box.y}"/>` +
    `<a:ext cx="${CONTENT_WIDTH}" cy="${box.height}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/>' +
    '<a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
  );
}

export function slideLayoutXml(cover: boolean): string {
  const title = cover ? COVER_TITLE_BOX : TITLE_BOX;
  const body = cover ? COVER_BODY_BOX : BODY_BOX;
  return (
    `<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" ` +
    `type="${cover ? "title" : "obj"}" preserve="1">` +
    `<p:cSld name="${cover ? "Title Slide" : "Title and Content"}"><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    layoutPlaceholder(2, "Title 1", '<p:ph type="title"/>', title) +
    layoutPlaceholder(3, "Body 2", '<p:ph type="body" idx="1"/>', body) +
    "</p:spTree></p:cSld>" +
    "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"
  );
}

export function slideLayoutRelsXml(): string {
  return relationships([
    relationship("rId1", SLIDE_MASTER_TYPE, "../slideMasters/slideMaster1.xml"),
  ]);
}

export function slideMasterXml(): string {
  const style = (size: number): string =>
    `<a:lvl1pPr><a:defRPr sz="${size}"><a:solidFill><a:srgbClr val="${PALETTE.ink}"/></a:solidFill>` +
    "</a:defRPr></a:lvl1pPr>";
  return (
    `<p:sldMaster xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">` +
    '<p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill>' +
    "<a:effectLst/></p:bgPr></p:bg><p:spTree>" +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    // The top of the inheritance chain. A slide's placeholder points at its
    // layout's, and the layout's at these — native masters always carry them,
    // and a chain that ends nowhere is another shape no real file has.
    layoutPlaceholder(2, "Title Placeholder 1", '<p:ph type="title"/>', TITLE_BOX) +
    layoutPlaceholder(3, "Body Placeholder 2", '<p:ph type="body" idx="1"/>', BODY_BOX) +
    "</p:spTree></p:cSld>" +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ' +
    'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" ' +
    'folHlink="folHlink"/>' +
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/>' +
    '<p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>' +
    `<p:txStyles><p:titleStyle>${style(TITLE_SIZE)}</p:titleStyle>` +
    `<p:bodyStyle>${style(BODY_SIZE)}</p:bodyStyle>` +
    `<p:otherStyle>${style(BODY_SIZE)}</p:otherStyle></p:txStyles>` +
    "</p:sldMaster>"
  );
}

export function slideMasterRelsXml(): string {
  return relationships([
    relationship("rId1", SLIDE_LAYOUT_TYPE, "../slideLayouts/slideLayout1.xml"),
    relationship("rId2", SLIDE_LAYOUT_TYPE, "../slideLayouts/slideLayout2.xml"),
    relationship("rId3", THEME_TYPE, "../theme/theme1.xml"),
  ]);
}

export function slideRelsXml(cover: boolean, links: readonly string[]): string {
  return relationships([
    relationship(
      "rId1",
      SLIDE_LAYOUT_TYPE,
      `../slideLayouts/slideLayout${cover ? 1 : 2}.xml`,
    ),
    ...links.map((href, index) =>
      relationship(`rId${index + 2}`, HYPERLINK_TYPE, href, true),
    ),
  ]);
}
