# Reading

Give it `content` — the file's bytes as base64. `filename` is optional and only
a hint. There is no `url`: fetching left with the outbound boundary, and an
address a model chose is governed where the caller already governs them.

**Magic bytes decide what a file is, then the declared type, then the name.** A
document served as `application/octet-stream` is ordinary — a `.hwp` behind a
download endpoint almost always is — so a header that disagrees with the file's
own first bytes is wrong about the file.

| Format | How |
|---|---|
| DOCX | `word/document.xml`, plus `styles.xml`, `numbering.xml` and the body's rels in the same `read()` — headers, footers and footnotes stay out, because they would interleave running heads with prose at every page boundary |
| XLSX | cached values from `xl/worksheets/*`, resolved through `xl/sharedStrings.xml`, one heading per visible sheet |
| PPTX | `ppt/slides/slide*.xml` in the order `presentation.xml` states, numbered; speaker notes left out |
| HWPX | `Contents/section*.xml` in numeric order, resolved against `Contents/header.xml` |
| HWP 5.x | OLE compound file → deflate per section → `HWPTAG_PARA_TEXT` records |
| ODT / ODS / ODP | one `content.xml`, one reader — ODF marks structure the same way whichever kind it is; a paragraph's text only, so tracked deletions, comments, footnotes and speaker notes stay out |
| RTF | control words, groups and escapes, with destinations (`\fonttbl`, `{\*\generator}`) skipped whole |

PDF, plain text and HTML are the caller's: none needs a parser Agent Studio lacks,
so routing one here would be a network round trip to reach the same library —
and a second copy of the extraction to keep in step with the first.

**A spreadsheet is where "the text" is least obviously defined**, and three
decisions carry it. Values, never formulas — `=SUM(B2:B9)` is how the number was
made, and the number is the answer. Position is rebuilt from each cell's own
address, because empty cells are simply absent and emitting them in file order
would shift every value into a column it does not belong to: a table that still
looks like a table and says something else. And the budget is spent in whole
rows, so a cut never leaves a line whose columns no longer line up.

That simple read is intentionally lossy. Use `inspect_spreadsheet` when formulas,
cell addresses, error cells, hidden sheets, external-link presence or macro
presence matter. It never recalculates formulas, follows an external link or
executes VBA. Hidden and very-hidden sheets are omitted unless explicitly
requested. See [Spreadsheets](spreadsheets.md).

Every read result also states `complete`, format-specific `omissions`, and
available counts in `structuredContent`. Text extraction is content recovery,
not original-preserving editing: styles, comments, tracked changes, speaker
notes and similar parts may be omitted by design. Re-rendering extracted text
creates a new document and does not preserve the source package.

**RTF is here for a different reason from the rest.** It is a text file, so
without a reader it is not refused — it is read as plain text and reaches the
model as thousands of control words with the prose scattered through them. A
format that fails by producing garbage is worth more than one that cannot be
opened at all.

## What survives

A document's shape is the part a reader was going to use, so it comes back as
Markdown rather than as lines: a heading keeps its level, a list keeps its order
and its depth, a table is a table with its column alignment, a link keeps its
target, emphasis stays emphasis, and a picture leaves a mark saying it was there
and what it was called. Everything else — fonts, colours, spacing — is
presentation, and a model has no use for it.

**Where a level comes from is decided per format, and never guessed.** DOCX
reads direct `w:outlineLvl`, then the style id, then `styles.xml` following
`w:basedOn`, then the style's name — which is what makes `제목 1` and
`Überschrift 1` work, since matching `Heading1` never did. HWPX reads
`hh:heading` out of `Contents/header.xml`, 한글's one outline mechanism. ODF
reads `text:outline-level`, which sits on the element being opened and was
being thrown away. PPTX reads `p:ph/@type`, the only trustworthy title signal a
deck has. RTF reads `\outlinelevel`. **HWP 5.x reads none**: the record layouts
that would carry one are not verified against the spec, and a wrong field
offset resolves to a real shape and answers confidently with the wrong level —
the `.doc` failure this reader refuses to ship. It says so in `omissions`.

Every extra part is *enrichment*. A missing, malformed or self-contradicting
one leaves the reader exactly where it was: a flat paragraph, never a guessed
level.

**A list whose markers are characters is still a list.** Three writers draw the
marker themselves rather than use their format's numbering — this repository's
own DOCX and PPTX renderers among them — and each format has its own way of
saying so: a hanging indent in DOCX and HWPX, an explicit `a:buNone` in PPTX, a
`{\listtext …}` group in RTF. The marker the writer *drew* is what is reported;
a counter reconstructed from a definition the file may not carry would be a
number that disagrees with the document. A paragraph that merely opens with a
dash is a paragraph, and it is escaped so it reads back as one.

**A heading that is also numbered is a heading.** Word writes `w:pStyle` before
`w:numPr`, so a marker chosen while walking past each element overwrote the
level — and numbering the headings is the ordinary shape of a Korean or a legal
template. The paragraph's properties are collected and read at its end instead.
`w:numId="0"`, which is how Word says a paragraph's numbering was removed, is
not a list.

**A column a cell does not fill still holds its place.** ODF stores a run of
identical or empty cells once, with `table:number-columns-repeated`, and a merge
leaves `table:covered-table-cell` behind; a reader that emits one separator per
element puts every value after either of them in a column it does not belong to.
The run is paid for when a later cell in the row needs it, which is also why the
`16384`-wide padding every ODS row ends with costs nothing.

**A tracked insertion says it is one.** Its text is `w:t` like any other, so
nothing in the words says the paragraph is a proposal rather than the document
— `inspect_document` is where a reader learns it.

**The cut falls on a block boundary, and inside a table on a row.** A table cut
between its header and its divider is not a table when it is read back, and a
row cut in half is a row whose columns no longer line up — the same failure the
spreadsheet reader spends its budget in whole rows to avoid.

## What Markdown cannot say

`inspect_document` answers the questions the text cannot: which cells a merge
covers and how far, whether a row is a header because the document said so or
because it came first, how deep a list nests, which style a paragraph wears,
where a shape sat on its slide, which part a picture points at. One line per
block, addressed by ordinal, with each block's true character count beside a
preview of its text — `from` and `to` page through a long document.

It is a line grammar rather than JSON for one reason: the caller cuts a tool
result at a fixed length, and a truncated JSON document is a total loss where a
truncated line grammar loses its last line. XLSX is refused by name — a
workbook's structure is `inspect_spreadsheet`'s question.

**What it refuses, it names.** A PDF, a web page and a text file are each
identified and sent back to the caller that reads them; the 97-2003 binaries
(`.doc` `.xls` `.ppt`), an `.epub` and an HWP 3.0 file are each identified by
name with what to do instead. A password-protected or distribution (배포용)
`.hwp` says which it is. "Unsupported" on its own buys the model another turn
spent guessing.

The 97-2003 formats are deliberately absent. They are OLE record streams, not
containers — `.doc` scatters its text through a piece table, `.xls` is a BIFF
stream — and a half-right parse of either produces something that *looks* like
text. That failure is worse than the refusal.

An HWP body with a truncated record header, extended length or payload is
refused. Returning its readable prefix as a complete document would conceal
the missing content.
