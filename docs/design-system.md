# The design system

Every colour, size, spacing rule and profile is decided in `src/write/theme.ts`,
and the four renderers read it rather than carrying their own. A profile is a
complete editorial treatment rather than a colour skin: it owns the palette,
table header contrast, cover accent proportion and slide-card geometry. Type
scales and content semantics stay shared, so changing profile never changes
the words or makes one format feel like a different document.

**Every profile is deliberately nobody's brand.** These are purpose-driven
house styles a document can safely carry to a customer, partner or public
office. Every text pairing earns its WCAG ratio in `theme.test.ts`, and a
CVD-validated categorical palette rides in the PPTX theme so a reader who adds
a chart gets sane colours rather than Office's defaults.

| Profile | Character | Distinguishing decisions |
|---|---|---|
| `executive` (default) | editorial, authoritative | deep navy, very narrow cover band, restrained corners, solid table header |
| `consulting` | directional, presentation-led | stronger blue-teal contrast, wider cover accent, more expressive cards |
| `formal` | conservative, print-first | no cover band, square geometry, light table header |
| `technical` | precise, systems-oriented | teal structure, narrow cover mark, near-square cards, light table header |
| `standard` | familiar corporate | the classic corporate-blue palette and rounded-card treatment |

The default `executive` palette is:

| Token | | Used for |
|---|---|---|
| `brand` | `#17324D` | Headings, table headers, filled surfaces |
| `brandLight` | `#2D6A78` | The hairline under a heading, quote bars, ordinals |
| `brandDeep` | `#0B5D7A` | Links and inline code |
| `brandTint` | `#EAF1F3` | Zebra rows, code grounds, cards |
| `surfaceTint` | `#F5F7F8` | Cover and section slides only |
| `ink` / `inkMuted` | `#18222B` / `#4F5D68` | Body, then captions and quotes |
| `rule` | `#CBD5DB` | Table rules, horizontal rules |

**Body pages are white.** A tinted ground under a whole document is ink
somebody pays for and an artefact in every photocopy, so the tint is kept for
the surfaces that carry it free — a cover slide, a section divider — and a
brand hairline under each heading does the work on every other page. Pure
black is never used: near-black reads as ink, full black on a screen glares.

**No font is named, in any format.** A face named here is a face the reader's
machine may not have, and the substitute is then chosen by nobody. Agent Studio's
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

**Spacing follows the reading role, not an application's default.** Page body
copy uses 1.5 leading; headings use 1.2, cover titles 1.15, subtitles 1.4, and
compact table or code text 1.35. Deck body copy uses 1.35, titles 1.15, cover
titles 1.1 and subtitles 1.3. DOCX, HWPX and PPTX state those values in their
native paragraph properties, while PDF uses the same ratios in its own layout
arithmetic. Tracking stays at zero so Hangul keeps its natural square advance;
native kerning is enabled from 12pt for Latin display text. No renderer
compresses characters to make a line fit.

**Covers share one information grammar.** A short profile-colour rule anchors
the title, the title carries the decision, and one quieter subtitle supplies
scope. Page formats keep that composition on white; PPTX moves it to the
profile tint and may add a restrained edge band because a projected slide can
carry a full-bleed field. The rule marks hierarchy and the band identifies the
deck profile; neither is repeated as arbitrary page decoration.

Tables are set with **horizontal rules only** — a full grid boxes every number
in, and the eye reads a table by its rows. The header is either a solid brand
field or a light brand tint, as the profile requires, and repeats when a table
breaks across pages.

Page numbers are a `PAGE` field in DOCX and a `slidenum` field in PPTX, so a
document that reflows or a deck that gains a slide renumbers itself; the PDF,
whose layout this repository decides, draws the number directly and only when
there is more than one page. **HWPX gets none** — OWPML puts a footer in a
control with its own sub-list anchored to the section, and that is a shape to get
exactly right against the one reader that either opens a file or does not.
