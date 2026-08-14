# Why it is shaped this way

## Why the bytes, and never a link

`render_document` returns the file inside the tool result, as an MCP `resource`
block. It used to upload to S3 and answer with a presigned URL, because a client
that received a non-image blob would decode it as UTF-8 and hand the model a
page of replacement characters — or, once that was fixed, an omission notice:

```
[binary resource omitted: application/vnd…, 24576 bytes — not text, so it
cannot be read here. Ask the server for a text representation.]
```

The caller carries the bytes now. AgentDure stores what a tool returns beside
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

The old reasoning still decides the format layer. DOCX, HWPX and PPTX are written by
composing their parts directly, and read by a tag walker rather than a DOM
parse: what these formats are actually used for here is a dozen elements, and a
library's idea of a paragraph is one more thing between the document and the
bytes. PDF is the exception — it is a layout problem, not a markup one, so
`pdf-lib` does the object model and this repository does the line breaking.
