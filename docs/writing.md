# Writing

`render_document` takes Markdown and produces one of four formats. Supported:
ATX headings, paragraphs, `**bold**`, `*italic*`, `` `code` ``, links, bullet
and numbered lists nested up to four levels (two spaces of indent to a level),
GFM tables **with column alignment** (`---:` sets a column flush right, `:---:`
centres it), block quotes, fenced code blocks, horizontal rules, and
`:::name` … `:::` directives — a container that names what its contents are.
The presentation engine reads the name as a slide archetype; the DOCX renderer
gives two of them a page treatment of its own (`:::metrics`, `:::comparison`);
the PDF and HWPX renderers unwrap the fences and render the contents as if they
were never written.

Two lines with no blank line between them are **one paragraph**, as Markdown
says. A renderer cannot recover a distinction the parser threw away, so the tool
description states it rather than guessing.

Syntax that is not supported — footnotes, raw HTML — passes through as the
characters that were written. Refusing to produce a document over one line of it
is a much worse outcome than a stray `<div>` a reader can see. An image is the
exception: nothing here fetches pictures, so `![alt](url)` becomes a link
labelled with its alt text, which keeps both the description and the address.
The one way an image gets *into* a document is `render_document`'s optional
`assets` argument — PNG or JPEG bytes sent by name alongside the Markdown,
referenced as `![caption](asset://name)`, and embedded by the pptx and docx
renderers. A referenced asset nobody sent is refused by name, and SVG is refused
with the fix stated: PowerPoint's `svgBlip` demands a raster fallback part
beside the vector one, and producing that means rasterising, which this
repository does not do. Rasterise first, send the PNG.

Lists carry **literal markers** in all four formats rather than a numbering
definition. What real numbering buys is the reader's editor renumbering a list
they edit; nothing here is edited before it is read, and the literal form is
what survives extraction back to text — which is how this server's own round
trip checks itself.

What each format then *makes* of that structure:

- [The design system](design-system.md) — the colour, type and language every renderer shares
- [The document engine](document-engine.md) — `docx`, `pdf` and `hwpx` as a report
- [The presentation engine](presentation-engine.md) — `pptx` as a planned deck
