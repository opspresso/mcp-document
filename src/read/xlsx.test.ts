/**
 * The spreadsheet reader, whose three decisions are all invisible when they go
 * wrong: a formula returned instead of its value, a sparse row shifted left
 * into the wrong columns, and a budget spent mid-row.
 */

import { strict as assert } from "node:assert";
import { test } from "node:test";
import { buildZip } from "../zip.js";
import { columnOf, xlsxToText, XlsxError } from "./xlsx.js";

const utf8 = (value: string) => new TextEncoder().encode(value);

const RELS = `<?xml version="1.0"?><Relationships>
  <Relationship Id="rId1" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Target="worksheets/sheet2.xml"/>
</Relationships>`;

function workbook(...names: string[]): string {
  const sheets = names
    .map((name, index) => `<sheet name="${name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
    .join("");
  return `<?xml version="1.0"?><workbook><sheets>${sheets}</sheets></workbook>`;
}

function sheet(rows: string): string {
  return `<?xml version="1.0"?><worksheet><sheetData>${rows}</sheetData></worksheet>`;
}

/** A workbook with one sheet, built from raw `<row>` markup. */
function oneSheet(rows: string, shared?: string[]): Uint8Array {
  const parts: Record<string, Uint8Array> = {
    "xl/workbook.xml": utf8(workbook("Sheet1")),
    "xl/_rels/workbook.xml.rels": utf8(RELS),
    "xl/worksheets/sheet1.xml": utf8(sheet(rows)),
  };
  if (shared) {
    const items = shared.map((value) => `<si><t>${value}</t></si>`).join("");
    parts["xl/sharedStrings.xml"] = utf8(`<?xml version="1.0"?><sst>${items}</sst>`);
  }
  return buildZip(parts);
}

const read = (bytes: Uint8Array, max = 90_000) => xlsxToText(bytes, max);

test("columnOf reads the format's base-26 with no zero digit", () => {
  assert.equal(columnOf("A1"), 0);
  assert.equal(columnOf("B2"), 1);
  assert.equal(columnOf("Z9"), 25);
  // The case a naive base-26 parse gets wrong: the letters are 1-based.
  assert.equal(columnOf("AA1"), 26);
  assert.equal(columnOf("AB1"), 27);
});

test("a shared string is resolved to its text, not left as an index", () => {
  const bytes = oneSheet(
    `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>`,
    ["이름", "부서"],
  );
  assert.equal(read(bytes).text, "## Sheet1\n이름 | 부서");
});

test("the stored value comes back, never the formula that made it", () => {
  // `=SUM(B2:B9)` is how the number was made; the number is the answer.
  const bytes = oneSheet(`<row r="1"><c r="A1"><f>SUM(B2:B9)</f><v>42</v></c></row>`);
  const { text } = read(bytes);
  assert.equal(text, "## Sheet1\n42");
  assert.doesNotMatch(text, /SUM/);
});

test("a sparse row keeps its columns instead of shifting left", () => {
  // The failure this prevents is silent: emitting in file order would put 9
  // under the first column, and the table would still look like a table.
  const bytes = oneSheet(`<row r="1"><c r="A1"><v>1</v></c><c r="D1"><v>9</v></c></row>`);
  assert.equal(read(bytes).text, "## Sheet1\n1 |  |  | 9");
});

test("an inline string is read like any other value", () => {
  const bytes = oneSheet(`<row r="1"><c r="A1" t="inlineStr"><is><t>inline</t></is></c></row>`);
  assert.equal(read(bytes).text, "## Sheet1\ninline");
});

test("every sheet is named, in workbook order", () => {
  const bytes = buildZip({
    "xl/workbook.xml": utf8(workbook("First", "Second")),
    "xl/_rels/workbook.xml.rels": utf8(RELS),
    "xl/worksheets/sheet1.xml": utf8(sheet(`<row r="1"><c r="A1"><v>1</v></c></row>`)),
    "xl/worksheets/sheet2.xml": utf8(sheet(`<row r="1"><c r="A1"><v>2</v></c></row>`)),
  });
  const { text, sheets, totalSheets } = read(bytes);
  assert.equal(text, "## First\n1\n\n## Second\n2");
  assert.equal(sheets, 2);
  assert.equal(totalSheets, 2);
});

test("the budget is spent in whole rows, and what was left is counted", () => {
  const rows = Array.from(
    { length: 50 },
    (_, index) => `<row r="${index + 1}"><c r="A${index + 1}"><v>${"x".repeat(20)}</v></c></row>`,
  ).join("");
  const { text, rows: kept, totalRows } = read(oneSheet(rows), 200);
  assert.ok(kept > 0 && kept < totalRows, `expected a partial read, got ${kept}/${totalRows}`);
  assert.equal(totalRows, 50);
  // Cut on a row boundary: no line is half a row's columns.
  for (const line of text.split("\n").slice(1)) {
    assert.ok(line === "" || line === "x".repeat(20), `unexpected partial line: ${line}`);
  }
});

test("trailing empty cells do not become trailing separators", () => {
  const bytes = oneSheet(`<row r="1"><c r="A1"><v>1</v></c><c r="C1"><v></v></c></row>`);
  assert.equal(read(bytes).text, "## Sheet1\n1");
});

test("a zip with no workbook part is refused as not being one", () => {
  assert.throws(() => read(buildZip({ "notes.txt": utf8("hi") })), XlsxError);
});

test("a workbook with nothing in it is refused rather than returned empty", () => {
  // An empty success reads as "this workbook has no data", which is a different
  // claim from "I could not read it".
  assert.throws(() => read(oneSheet("")), XlsxError);
});
