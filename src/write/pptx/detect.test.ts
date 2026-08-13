/**
 * The archetype rules, tested at the boundary that matters: what matches.
 *
 * Every rule is conservative by design, so most of these tests are about what
 * must *not* match — the cost of a wrong layout is a slide that misrepresents
 * its content, which is worse than any plain slide.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { parseMarkdown } from "../../markdown.js";
import { specialise } from "./detect.js";

function detect(markdown: string) {
  const { blocks } = parseMarkdown(markdown);
  const [first, ...rest] = blocks;
  assert.ok(first?.kind === "heading");
  return specialise(first.runs, rest);
}

test("two to four sub-headings with short lines are cards", () => {
  const slide = detect(
    "## 핵심 가치\n\n### Automation\n\n반복 작업 자동화\n\n### Integration\n\nMCP 기반 연결\n\n### Intelligence\n\nLLM 기반 판단",
  );
  assert.equal(slide?.type, "cards");
  assert.ok(slide?.type === "cards" && slide.cards.length === 3);
  assert.ok(slide?.type === "cards" && slide.cards[1]?.body !== undefined);
});

test("a heading without a line still makes a card, but one heading is not a set", () => {
  assert.equal(detect("## 구성\n\n### 하나\n\n### 둘")?.type, "cards");
  assert.equal(detect("## 구성\n\n### 하나\n\n설명"), undefined);
});

test("a code block or a list under a sub-heading keeps the section as content", () => {
  assert.equal(detect("## s\n\n### A\n\n```\ncode\n```\n\n### B\n\nb"), undefined);
  assert.equal(detect("## s\n\n### A\n\n- 항목\n\n### B\n\nb"), undefined);
});

test("five headings are a list of topics, not a row of cards", () => {
  assert.equal(detect("## s\n\n### 1\n\n### 2\n\n### 3\n\n### 4\n\n### 5"), undefined);
});

test("short numeric bullets are metrics, figure leading or trailing", () => {
  const leading = detect("## 주요 성과\n\n- 99.99% Availability\n- 43% Cost Reduction\n- 2.4x Deployment Speed");
  assert.equal(leading?.type, "metrics");
  assert.ok(leading?.type === "metrics" && leading.metrics[0]?.value === "99.99%");
  const trailing = detect("## 성과\n\n- 가용성 99.99%\n- 절감률 43%");
  assert.equal(trailing?.type, "metrics");
  assert.ok(trailing?.type === "metrics" && trailing.metrics[0]?.value === "99.99%");
});

test("a list is not metrics when its lines are prose, numbered, or figure-free", () => {
  assert.equal(detect("## s\n\n- 지표를 개선했다\n- 비용을 줄였다"), undefined);
  assert.equal(detect("## s\n\n1. 99% a\n2. 43% b"), undefined, "an ordered list is a sequence");
  assert.equal(
    detect("## s\n\n- 99.99% 가용성을 확보해 어떤 상황에도 서비스가 계속된다는 뜻이다\n- 43% 절감"),
    undefined,
    "a long line is a sentence, not a metric",
  );
});

test("a lone block quote is a quote slide, with at most a line of attribution", () => {
  const bare = detect("## 고객의 말\n\n> 도입 후 반복 업무가 사라졌다.");
  assert.equal(bare?.type, "quote");
  const attributed = detect("## 고객의 말\n\n> 도입 후 반복 업무가 사라졌다.\n\n— 운영팀 리드");
  assert.equal(attributed?.type, "quote");
  assert.ok(attributed?.type === "quote" && attributed.attribution !== undefined);
  assert.equal(
    detect("## s\n\n> 인용\n\n이어지는 본문 문단이 인용을 논증의 일부로 만든다. 이 문단은 충분히 길어서 출처 표기로 볼 수 없다."),
    undefined,
  );
});

test("a vs-titled pair of sub-headings is a comparison; without vs it is cards", () => {
  const markdown = (title: string) =>
    `## ${title}\n\n### IRSA\n\n- 표준 방식\n- 넓은 지원\n\n### Pod Identity\n\n- 간단한 설정\n- 신규 권장`;
  const compared = detect(markdown("IRSA vs Pod Identity"));
  assert.equal(compared?.type, "comparison");
  assert.ok(compared?.type === "comparison" && compared.columns[0].lines.length === 2);
  const plain = detect(markdown("인증 방식"));
  assert.notEqual(plain?.type, "comparison");
});

test("a comparison needs exactly two columns with something in each", () => {
  assert.equal(detect("## A vs B\n\n### A\n\n- a"), undefined, "one column");
  assert.equal(detect("## A vs B\n\n### A\n\n- a\n\n### B\n\n- b\n\n### C\n\n- c"), undefined, "three columns");
  assert.equal(detect("## A vs B\n\n### A\n\n### B\n\n- b"), undefined, "an empty column");
});
