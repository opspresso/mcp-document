/**
 * What a section's content *is*, recognised from its shape.
 *
 * Markdown says what to show and none of how; these functions read the
 * structure the author already wrote — three sub-headings with a line each, a
 * list of numbers, a lone block quote — and name the slide archetype that
 * structure was reaching for. No directive syntax is needed for any of it,
 * which is the point: the deck improves without the input contract changing.
 *
 * **Every rule here is conservative.** A section that does not match a pattern
 * *exactly* stays a content slide, because a layout forced onto content it
 * does not fit is worse than a plain slide — a "cards" slide with a code block
 * wedged into a card is how generated decks get their reputation. When a rule
 * says 2 to 4, a fifth card means the section was prose after all.
 */

import type { Block, Run } from "../../markdown.js";
import { plainTextOf } from "../../markdown.js";
import type { Card, CompareColumn, Metric, Milestone, Slide } from "./types.js";

/** How much text fits a card before the card is prose. */
const CARD_BODY_LIMIT = 160;
/** A metric line: short enough to be a figure with a name. */
const METRIC_LIMIT = 44;
const METRIC_VALUE_LIMIT = 12;
/** A metric's name is a few words; more words make it a sentence. */
const METRIC_LABEL_TOKENS = 4;
/** An attribution is a line, not a paragraph, and it announces itself. */
const ATTRIBUTION_LIMIT = 80;
const ATTRIBUTION_LEAD = /^[—–-]/;
/** Lines one comparison column holds before it is a table's job. */
const COMPARE_LINE_LIMIT = 6;

/** A process step has to fit in a node; a sentence does not. */
const STEP_LIMIT = 48;
/** What a timeline says happened has to fit under its station. */
const MILESTONE_LIMIT = 40;

/**
 * A token that names a time: 2026, 2026년, Q3, 8월, 3분기, 2주차. The whole
 * token must be the date — "2026년의" is prose about a year, not a station.
 */
const WHEN = /^(20\d{2}년?|q[1-4]|\d{1,2}월|\d{1,2}분기|\d{1,2}주차?)$/i;

/** A comparison needs the title to say it is one: "A vs B", "도입 비교". */
const COMPARISON_TITLE = /(^|\s)vs\.?(\s|$)|비교/i;

function isSubheading(block: Block): block is Extract<Block, { kind: "heading" }> {
  return block.kind === "heading" && block.level >= 3;
}

/**
 * Cards: 2 to 4 sub-headings, each followed by at most one short paragraph.
 *
 * The all-heading shape is what an author writes when they mean a set of
 * parallel things — 핵심 가치, three pillars, four features. Anything else in
 * the section (a list, a code block, a second paragraph) means the headings
 * were structure, not a set, and the section stays content.
 */
export function asCards(title: readonly Run[], blocks: readonly Block[]): Slide | undefined {
  const cards: Card[] = [];
  for (const block of blocks) {
    if (isSubheading(block)) {
      cards.push({ title: block.runs });
      continue;
    }
    const last = cards[cards.length - 1];
    if (
      block.kind === "paragraph" &&
      last &&
      !last.body &&
      plainTextOf(block.runs).length <= CARD_BODY_LIMIT
    ) {
      last.body = block.runs;
      continue;
    }
    return undefined;
  }
  if (cards.length < 2 || cards.length > 4) {
    return undefined;
  }
  return { type: "cards", title: [...title], cards };
}

/**
 * Metrics: one bullet list of 2 to 4 short lines, each a figure with a name —
 * "99.99% Availability", "가용성 99.99%". The figure may lead or trail; it must
 * be one token with a digit in it, and the name must be the rest.
 */
export function asMetrics(title: readonly Run[], blocks: readonly Block[]): Slide | undefined {
  const [list] = blocks;
  if (blocks.length !== 1 || list?.kind !== "list" || list.ordered) {
    return undefined;
  }
  if (list.items.length < 2 || list.items.length > 4) {
    return undefined;
  }
  const metrics: Metric[] = [];
  for (const item of list.items) {
    if (item.depth !== 0) {
      return undefined;
    }
    const text = plainTextOf(item.runs).trim();
    if (text.length > METRIC_LIMIT) {
      return undefined;
    }
    const tokens = text.split(/\s+/);
    if (tokens.length < 2 || tokens.length > METRIC_LABEL_TOKENS + 1) {
      return undefined;
    }
    const leads = /\d/.test(tokens[0]!) && tokens[0]!.length <= METRIC_VALUE_LIMIT;
    const trails = /\d/.test(tokens[tokens.length - 1]!) && tokens[tokens.length - 1]!.length <= METRIC_VALUE_LIMIT;
    if (leads) {
      metrics.push({ value: tokens[0]!, label: [{ text: tokens.slice(1).join(" ") }] });
    } else if (trails) {
      metrics.push({ value: tokens[tokens.length - 1]!, label: [{ text: tokens.slice(0, -1).join(" ") }] });
    } else {
      return undefined;
    }
  }
  return { type: "metrics", title: [...title], metrics };
}

/**
 * A quote slide: one block quote alone, or with a single short dash-led line
 * after it — "— 운영팀 리드". A paragraph that does not announce itself as an
 * attribution means the quote was part of an argument, and arguments are
 * content slides.
 */
export function asQuote(title: readonly Run[], blocks: readonly Block[]): Slide | undefined {
  const [quote, attribution, ...rest] = blocks;
  if (quote?.kind !== "quote" || rest.length > 0) {
    return undefined;
  }
  if (attribution === undefined) {
    return { type: "quote", title: [...title], quote: quote.runs };
  }
  if (attribution.kind !== "paragraph") {
    return undefined;
  }
  const line = plainTextOf(attribution.runs).trim();
  if (!ATTRIBUTION_LEAD.test(line) || line.length > ATTRIBUTION_LIMIT) {
    return undefined;
  }
  return { type: "quote", title: [...title], quote: quote.runs, attribution: attribution.runs };
}

/**
 * A comparison: the title says "vs" (or 비교), and the section is exactly two
 * sub-headings, each followed by short lists or paragraphs. Two named columns
 * is the shape; three is a cards section, and prose under one heading is not
 * a comparison at all.
 */
export function asComparison(title: readonly Run[], blocks: readonly Block[]): Slide | undefined {
  if (!COMPARISON_TITLE.test(plainTextOf(title))) {
    return undefined;
  }
  const columns: CompareColumn[] = [];
  for (const block of blocks) {
    if (isSubheading(block)) {
      columns.push({ title: block.runs, lines: [] });
      continue;
    }
    const column = columns[columns.length - 1];
    if (!column) {
      return undefined;
    }
    if (block.kind === "paragraph") {
      column.lines.push({ runs: block.runs, bullet: false });
      continue;
    }
    if (block.kind === "list") {
      for (const item of block.items) {
        column.lines.push({ runs: item.runs, bullet: true });
      }
      continue;
    }
    return undefined;
  }
  if (columns.length !== 2) {
    return undefined;
  }
  for (const column of columns) {
    if (column.lines.length === 0 || column.lines.length > COMPARE_LINE_LIMIT) {
      return undefined;
    }
  }
  return { type: "comparison", title: [...title], columns: [columns[0]!, columns[1]!] };
}

/** The lone flat ordered list a process or a timeline is made of, or nothing. */
function orderedItems(blocks: readonly Block[], most: number): Run[][] | undefined {
  const [list] = blocks;
  if (blocks.length !== 1 || list?.kind !== "list" || !list.ordered) {
    return undefined;
  }
  if (list.items.length < 3 || list.items.length > most) {
    return undefined;
  }
  if (list.items.some((item) => item.depth !== 0)) {
    return undefined;
  }
  return list.items.map((item) => item.runs);
}

/**
 * A timeline: an ordered list whose every step opens with a date — 2026년,
 * Q3, 8월. All of them, not most: one undated step means the author was
 * writing a sequence of actions, which is a process.
 */
export function asTimeline(title: readonly Run[], blocks: readonly Block[]): Slide | undefined {
  const items = orderedItems(blocks, 6);
  if (!items) {
    return undefined;
  }
  const milestones: Milestone[] = [];
  for (const runs of items) {
    const text = plainTextOf(runs).trim();
    const tokens = text.split(/\s+/);
    if (tokens.length < 2 || !WHEN.test(tokens[0]!)) {
      return undefined;
    }
    const rest = tokens.slice(1).join(" ");
    if (rest.length > MILESTONE_LIMIT) {
      return undefined;
    }
    milestones.push({ when: tokens[0]!, what: [{ text: rest }] });
  }
  return { type: "timeline", title: [...title], milestones };
}

/**
 * A process: three to five short ordered steps, each of which fits in a node.
 * The numbers the author wrote become the numbers in the nodes, so nothing is
 * renumbered and nothing is lost if the section falls back to a plain list.
 */
export function asProcess(title: readonly Run[], blocks: readonly Block[]): Slide | undefined {
  const items = orderedItems(blocks, 5);
  if (!items) {
    return undefined;
  }
  if (items.some((runs) => plainTextOf(runs).length > STEP_LIMIT)) {
    return undefined;
  }
  return { type: "process", title: [...title], steps: items.map((runs) => [...runs]) };
}

/**
 * The first archetype the section's shape matches, or nothing.
 *
 * Order matters only where patterns could overlap, and they barely can: a
 * comparison requires its title to say so, so it is asked first; cards and
 * metrics are disjoint (headings against a list); a quote shares nothing with
 * either.
 */
export function specialise(title: readonly Run[] | undefined, blocks: readonly Block[]): Slide | undefined {
  if (!title || blocks.length === 0) {
    return undefined;
  }
  return (
    asComparison(title, blocks) ??
    asCards(title, blocks) ??
    asMetrics(title, blocks) ??
    asTimeline(title, blocks) ??
    asProcess(title, blocks) ??
    asQuote(title, blocks)
  );
}
