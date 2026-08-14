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

## Why not an SDK

The protocol surface is four methods, and the one dependency that mattered in
the sibling repository was an SDK whose schema generation changed under a server
written against an older release. A hand-written handler has no such drift.

The same reasoning decides the format layer. DOCX, HWPX and PPTX are written by
composing their parts directly, and read by a tag walker rather than a DOM
parse: what these formats are actually used for here is a dozen elements, and a
library's idea of a paragraph is one more thing between the document and the
bytes. PDF is the exception — it is a layout problem, not a markup one, so
`pdf-lib` does the object model and this repository does the line breaking.
