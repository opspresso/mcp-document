# The design system

Every colour and size is decided in `src/write/theme.ts`, and the four renderers
read it rather than carrying their own. Before that they had already drifted —
the table border was one grey in three of them and a different grey in the PDF,
and a level 1 heading was 16pt in DOCX, 18pt in HWPX and 20pt in the PDF. The
same Markdown produced three documents of visibly different weight.

**The palette is deliberately nobody's brand.** It used to be AgentDure's
indigo-violet, taken from the console — and a document handed to a customer, a
partner or a public office should not arrive dressed in its tooling's colours.
What replaced it is the *average* of what professional documents already look
like: a restrained corporate blue for structure, neutral near-black ink, quiet
grey rules — blue because it is the hue every reader has seen a thousand
reports in, so it signals "document" rather than "product". Every pairing
still earns its WCAG ratio in `theme.test.ts`, and a CVD-validated categorical
palette rides in the PPTX theme so a reader who adds a chart to the deck gets
sane colours rather than Office's defaults.

| Token | | Used for |
|---|---|---|
| `brand` | `#1F4E79` | Headings, table headers, filled surfaces |
| `brandLight` | `#4472C4` | The hairline under a heading, quote bars, ordinals |
| `brandDeep` | `#0563C1` | Links and inline code |
| `brandTint` | `#EEF3F9` | Zebra rows, code grounds, cards |
| `surfaceTint` | `#F4F6F9` | Cover and section slides only |
| `ink` / `inkMuted` | `#212529` / `#595959` | Body, then captions and quotes |
| `rule` | `#D9DEE5` | Table rules, horizontal rules |

**Body pages are white.** A tinted ground under a whole document is ink
somebody pays for and an artefact in every photocopy, so the tint is kept for
the surfaces that carry it free — a cover slide, a section divider — and a
brand hairline under each heading does the work on every other page. Pure
black is never used: near-black reads as ink, full black on a screen glares.

**No font is named, in any format.** A face named here is a face the reader's
machine may not have, and the substitute is then chosen by nobody. AgentDure's
own faces make that worse rather than better — Figtree, Chakra Petch and
JetBrains Mono carry no Hangul at all, so a Korean document set in them is a
document set in whatever the system falls back to. Colour and layout carry the
identity instead, and they survive every substitution. The PDF is the exception
it always was: it embeds Nanum Gothic, because PDF has no system stack behind it.

**The language is stated, so the system default is the *right* one.** Not
naming a face leaves the choice to the reader's Office — which picks its
east-Asian default by the run's language, and an unlabelled run falls back to
the *reader's locale*: 한글 through a Chinese or Japanese face on a non-Korean
machine. A document with Korean in it carries `themeFontLang eastAsia="ko-KR"`
and a run-language default in DOCX, and per-run `lang="ko-KR"` in PPTX — on a
Korean Windows that resolves to 맑은 고딕, on a Mac to Apple SD Gothic Neo,
and on every machine to *its* Korean face rather than a guess. Still no face
named anywhere.

Tables are set with **horizontal rules only** — a full grid boxes every number
in, and the eye reads a table by its rows. The header is filled with the brand
colour and repeats when a table breaks across pages.

Page numbers are a `PAGE` field in DOCX and a `slidenum` field in PPTX, so a
document that reflows or a deck that gains a slide renumbers itself; the PDF,
whose layout this repository decides, draws the number directly and only when
there is more than one page. **HWPX gets none** — OWPML puts a footer in a
control with its own sub-list anchored to the section, and that is a shape to get
exactly right against the one reader that either opens a file or does not.
