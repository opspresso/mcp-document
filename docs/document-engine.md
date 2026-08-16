# The document engine

The three page formats — `docx`, `pdf`, `hwpx` — read the same structure the
presentation engine does and answer with a *report* rather than a deck:
because a page reflows where a slide is a fixed box, the machinery differs
entirely — no packing, no layout selection, just the devices a document earns.
The cover, the numbered chapters and the contents threshold are one decision
(`write/semantics.ts`); what follows describes DOCX, and then where the other
two differ.

The selected design profile changes the page's visual register without changing
that structure. It supplies the colours, cover-rule proportion and table-header
treatment to all three renderers. `formal` and `technical` use light headers
with dark text; the other profiles use a filled header. HWPX expresses the same
decision through its `borderFill` and `charPr` tables rather than approximating
it in the body.

**An opening `#` is the cover**: the title oversized on its own page, the
first paragraph under it as the subtitle, a brand rule above, and no running
head or page number on it (`titlePg`). The cover title is `Heading1` with its
size overridden inline, deliberately — the style id is what this server's own
reader maps back to `#`, and a cover styled any other way reads back as a
document that lost its title. **Every later `#` opens a chapter** on a fresh
page, a large light-brand ordinal above the heading. **The running head**
names the document on every page but the cover, in the header part where no
text extractor looks. A quote is a **callout** — the brand bar and the tint
ground marking an editorial aside without interrupting the main text.

**A report with a cover and three or more level 1-2 headings gets a contents
page** — without page numbers, which is the decision that makes it honest.
Page numbers are the one thing about a TOC this renderer cannot know (text
reflows to the reader's fonts), and asking Word to fill them in on open —
`settings.xml`'s `updateFields`, the usual generator's answer — greets every
reader with a dialog about fields referring to other files. A list of the
headings with the numbers omitted (`TOC \n`) is one the renderer can finish
itself: correct as written, nothing to update, no dialog. The field wrapper
stays, so a reader who edits the headings can press F9 and have Word
regenerate the list. A document with less structure gets no contents page.

**Directives get a page treatment only when asked.** `:::metrics` becomes a
key-figure strip — the numbers large in the brand colour, names beneath, a
borderless table doing the alignment; `:::comparison` becomes a two-column
table with the column names as its header row. Nothing is recognised
automatically here, deliberately: a document is prose, and prose transformed
unasked is prose misquoted. The deck recognises; the page waits to be told.
The other directive names unwrap — a page already reads a list as a list.

**A paragraph that is exactly one `![caption](asset://…)` image is a figure**:
the picture centred at its aspect ratio, never upscaled past its pixels, the
caption under it in the muted caption size. An image inside prose stays a
link, as it always was.

**The PDF carries the same cover and chapters — and the one contents page
with real page numbers.** It is the format whose pages this repository lays
out itself, so the contents page is reserved after the cover, the body is laid
knowing which page every heading lands on, and the list is drawn last with the
numbers as facts. The cover is counted but not numbered, as every title page
is. **HWPX carries them too, conservatively**: the page break is the `hp:p`
attribute 한글 itself writes, the cover and ordinal are two more entries in
the `charPr`/`paraPr` tables the body already resolves against, and the
contents list goes without page numbers, as the DOCX one does. The cover's
short brand rule is a bottom border on an empty paragraph whose right margin
sets the profile length; it avoids introducing a drawing control solely for
decoration.

**Three of the devices above are DOCX's alone.** Neither of the other two
writes a header part, so there is no running head in either. Neither gives
`:::metrics` or `:::comparison` a treatment — a directive is unwrapped and its
contents stand where they stood. And `assets` are accepted for `docx` and
`pptx` only, so in the PDF and HWPX an image is the link it is everywhere else.

`npm run demo` renders `scripts/demo-doc.md` — a report that uses every
device — into `build/demo.docx`, and the same report into `build/demo.pdf`
and `build/demo.hwpx`: one source in three page formats is what shows a
drifted colour or a diverged heading scale, which is what `write/theme.ts`
exists to prevent.

## The Korean font, and a bug worth knowing about

PDF's built-in fonts cover Latin-1 and nothing else, so a Korean document needs a
real font inside the file. Nanum Gothic ships in `assets/fonts` (SIL OFL 1.1).

It is embedded **whole**, not subset, because `@pdf-lib/fontkit`'s subsetter
silently drops most Hangul glyphs. It does not fail: the text layer stays
perfect, so extraction returns the document exactly, and the page shows blanks
where two thirds of the characters should be — `2026년 1분기 보고서` renders as
`6년      서`. A document that reads correctly to a machine and is unreadable to
a person is the worst shape this could take, so a whole face is embedded and a
Korean PDF costs about 750KB. The bold face is embedded only when something is
bold, which is what keeps a plain document to one of them.

Nanum Gothic rather than Noto Sans KR for a mechanical reason: Google Fonts now
publishes Noto Sans KR as a variable font, and putting one of those through
fontkit has more ways to go wrong than a static TTF does.

Line breaking is per script — Latin at spaces, CJK between any two characters.
Korean prose has no spaces to break at, and a breaker that waits for one
produces a single line running off the page.

## HWPX has not been opened in 한글

DOCX and PDF have many independent implementations and all of them are
forgiving. HWPX has essentially one reader that matters, and it either opens a
file or it does not.

What is verified: the package layout, the `mimetype` entry first and stored, and
every `charPrIDRef` and `paraPrIDRef` in the body resolving against
`Contents/header.xml` — plus a round trip through this server's own HWPX reader.
What is **not** verified is whether 한글 accepts the result, and nothing in CI
can tell you. Open one before relying on it.
