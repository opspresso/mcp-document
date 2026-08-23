# Why it is shaped this way

## Why the bytes, and never a link

`render_document` and `render_spreadsheet` return the file inside the tool result, as an MCP `resource`
block. It used to upload to S3 and answer with a presigned URL, because a client
that received a non-image blob would decode it as UTF-8 and hand the model a
page of replacement characters — or, once that was fixed, an omission notice:

```
[binary resource omitted: application/vnd…, 24576 bytes — not text, so it
cannot be read here. Ask the server for a text representation.]
```

The caller carries the bytes now. Agent Studio stores what a tool returns beside
every other byte one of its runs produced, with one retention window, one
gallery and one delete button — which is what makes a rendered document
something a person can find again rather than a link that quietly expires.

One rule survives the change and runs both ways: what cannot be produced or
extracted is reported as a tool error with the reason, never as an empty
success. "The document has no text layer, it is a scan" is actionable; an empty
string reads as "the document is empty".

The bytes are bounded — `MAX_RENDERED_BYTES` in `limits.ts` — and a document
past it is refused *here*, with a sentence. The caller's transport bounds the
whole JSON-RPC envelope, and base64 inflates by 4/3, so letting it be cut there
turns "the document is large" into a parse failure that says nothing at all.

Generated artifacts are not returned before a second parser reopens them.
OOXML and HWPX validation checks required parts and every internal relationship;
PDF validation reloads the file and checks the reported page count. The result
reports this separately from visual validation, which the server does not
perform.

## Why the read model is not the write model

`markdown.ts` holds a document model already, and the reading side does not use
its union. The difference is direction. That one is, in its own header's words,
"the greatest common denominator of DOCX, PDF and HWPX, and anything richer
would be a feature one renderer could honour and the others would silently
drop" — a claim about what four renderers can *draw*. A reader's subject is what
a file *contained*, which no renderer bounds.

Three kinds prove it. A cell that spans columns has no GFM syntax, so
`parseMarkdown` could never produce one and the four renderers would gain
nothing from a shape that carried it. An image is deliberately a link on the
write side — `tools.ts` refuses assets for HWPX because that renderer draws a
picture as a link — so an `image` block is exactly the "one renderer honours it,
the others drop it" the comment warns about. And a slide boundary is a fact
about a deck rather than a thing to lay out.

So `read/blocks.ts` imports the five kinds where a document and a renderer agree
and the whole inline vocabulary — `Run`, `ListItem`, `Align` — and adds only what
reading needs. The two models meet where they already both speak: Markdown text.
`blocksToMarkdown` writes `![alt](x)` and `parseMarkdown` reads it back. Neither
imports the other's union, and the four renderers keep exhaustive `kind`
switches with no `default`, which is what makes "handled everywhere" a compile
error rather than a promise.

The escaping that makes that meeting honest lives beside the parser, in
`markdown.ts`, and answers with the parser's own regexes. Two files with two
lists of the same punctuation drift the first time one of them gains a rule.

## The protocol is the SDK's; the formats are not

The protocol surface was four methods, and the one dependency that mattered in
the sibling repository was an SDK whose schema generation changed under a server
written against an older release — so it was written by hand here. Revision
`2026-07-28` ended that trade. It removed the `initialize` handshake and added a
per-request `_meta` envelope, `server/discover`, `resultType` on every result,
the `ttlMs`/`cacheScope` hints the list verbs now require, `Mcp-Param-*`
mirroring from a tool's own schema, and multi round-trip results. Four methods
became a moving surface, and following it by hand is the larger risk now.

`@modelcontextprotocol/server` serves **both eras from one endpoint**: a client
opening with `server/discover` gets `2026-07-28`, one opening with the handshake
is served statelessly as before. The tool schemas stay the JSON Schema objects
`tools.ts` declares — converted, not rewritten, because their descriptions are
the tools' documentation and restating them elsewhere is a transcription
exercise with every chance of a quiet omission.

The old reasoning still decides the format layer. DOCX, HWPX, PPTX and XLSX are written by
composing their parts directly, and read by a tag walker rather than a DOM
parse: what these formats are actually used for here is a dozen elements, and a
library's idea of a paragraph is one more thing between the document and the
bytes. PDF is the exception — it is a layout problem, not a markup one, so
`pdf-lib` does the object model and this repository does the line breaking.
