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
| DOCX | `word/document.xml` only — headers, footers and footnotes would interleave running heads with prose at every page boundary |
| XLSX | values from `xl/worksheets/*`, resolved through `xl/sharedStrings.xml`, one heading per sheet |
| PPTX | `ppt/slides/slide*.xml` in deck order, numbered; speaker notes left out |
| HWPX | `Contents/section*.xml`, in numeric order |
| HWP 5.x | OLE compound file → deflate per section → `HWPTAG_PARA_TEXT` records |
| ODT / ODS / ODP | one `content.xml`, one reader — ODF marks structure the same way whichever kind it is |
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

**RTF is here for a different reason from the rest.** It is a text file, so
without a reader it is not refused — it is read as plain text and reaches the
model as thousands of control words with the prose scattered through them. A
format that fails by producing garbage is worth more than one that cannot be
opened at all.

A heading style becomes its Markdown `#`, a list paragraph becomes `- `, and
table cells are separated by ` | `. Everything else — fonts, colours, spacing —
is presentation, and a model has no use for it.

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
