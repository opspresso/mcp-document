# Spreadsheets

The spreadsheet tools separate three operations that should not be confused:
simple text extraction, formula inspection, and creation of a new workbook.

## Read values as text

`read_document` returns visible XLSX sheets as row text. It uses stored values,
not formulas, and bounds output on whole rows. This is the smallest result for a
model that only needs to understand a table. The structured result names
omissions so it cannot be mistaken for an original-preserving workbook read.

## Inspect without executing

`inspect_spreadsheet` returns addressed cells in `values`, `formulas` or `both`
mode. A formula is text and its value is whatever the workbook cached when it
was last calculated; this server does not calculate or verify the pair. Hidden
and very-hidden sheets require `includeHidden: true`. External workbook links
and VBA projects are reported as presence only and are never followed or run.

The inspection is bounded to 10,000 cells and reports whether it is complete.
Use the result for formula review and diagnosis, not as evidence that a model
currently calculates correctly.

## Create a new workbook

`render_spreadsheet` takes named sheets and rows of scalar cells. A string that
begins with `=` remains a literal string. A formula must be explicit:

```json
{
  "sheets": [
    {
      "name": "Summary",
      "rows": [
        ["Item", "Value"],
        ["Revenue", 40],
        ["Cost", 10],
        ["Profit", { "formula": "B2-B3", "cachedValue": 30 }]
      ]
    }
  ]
}
```

`cachedValue` is optional. Supply it only when it comes from a known calculation;
it is not recomputed here. The generated workbook requests a full calculation
when opened, freezes a non-empty header row and applies bounded column widths.
It is structurally reopened before it is returned.

This creates a new workbook. It does not edit or preserve an input workbook's
styles, charts, comments, macros, external links, named ranges or pivot tables.
