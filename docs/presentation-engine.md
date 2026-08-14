# The presentation engine

`pptx` is the one output that is not a page of prose, so it does more than set
type: it reads the Markdown's structure and plans a deck from it. The pipeline
is planner → presentation AST → renderer, and the planner's first decision is
what each section *is*.

**Heading levels are the deck's skeleton.** A `#` that opens the document is
the cover, and its first paragraph becomes the subtitle — everything else moves
past the cover, because a list on a title page is a list that belongs on the
next slide. Every later `#` becomes a full-bleed **section divider**, numbered
01, 02 in order. Every `##` opens a slide; levels 3 and below stay in the body.
A final 감사합니다 / Thank you / Q&A heading becomes a centred **closing**
slide. Nothing else says "new slide" — not a horizontal rule, which is what
Marp uses, because that turns a decorative divider into a page break in every
document that was not written as a deck.

**A slide whose shape says what it is gets a matching layout.** Two to four
`###`s with a short line each become a row of **cards**; a short bullet list of
figures (`- 99.99% Availability`) becomes big-number **metrics**; a lone block
quote with a `— author` line becomes a **quote** slide; two `###`s under an
"A vs B" title become a two-column **comparison**; three to five short numbered
steps become a **process** flow with arrows; date-led steps (`1. Q1 파일럿`)
become a **timeline**; a section that is exactly one `![caption](asset://…)`
image becomes an **image** slide, the picture placed at its aspect ratio with
the caption under it. Every rule is conservative — a section that does not
match exactly stays a plain content slide, because a layout forced onto content
it does not fit is worse than no layout. Wrapping one such group in
`:::cards` … `:::` (also `metrics`, `comparison`, `process`, `timeline`,
`quote`) forces the archetype where recognition would not fire; the other three
formats render the contents as if the fences were never written. Everything is
native PowerPoint objects — shapes, tables, text — so a reader can edit any of
it.

**The design lives on the slide layouts**, which is where PowerPoint itself
keeps a template's: the cover's tinted ground and brand band, the divider's
brand field, the footer that names the deck on every content slide. Slides
carry only their content, so editing one never means stepping around
furniture — and this server's own reader, which reads only the slide parts,
never sees the decoration.

**What does not fit continues on the next slide.** The break prefers the last
sub-heading on the slide — the topic moves whole, and its heading becomes the
continuation's title (`아키텍처 — Control Plane`); only a break with no
sub-heading to move falls back to `… (계속)`. The line count is an estimate:
nothing here measures a font, because unlike the PDF renderer this one embeds
none and cannot know what the reader's PowerPoint will substitute. Being a line
out puts a line nearer the edge than intended, which is a better failure than
letting a slide overflow and lose it. Blocks are flattened to lines before a
slide is filled, which is what lets a numbered list split across two slides and
still count 4, 5, 6; a table splits by row and repeats its header.

`npm run demo` renders `scripts/demo-deck.md` — one deck that uses every
archetype — into `build/demo.pptx`, which is the review surface for any change
to a renderer: the tests prove the package round-trips, and the demo deck is
what a person opens to see the design.
