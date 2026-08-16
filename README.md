# mcp-document

An MCP server that parses office documents and writes them.

| Tool | Takes | Returns |
|---|---|---|
| `read_document(content, filename?)` | DOCX, PPTX, XLSX, HWP, HWPX, ODT/ODS/ODP, RTF — bytes as base64 | the **text**, as an MCP `text` block |
| `render_document(format, content, profile?, title?, filename?, assets?)` | Markdown → `docx` `pptx` `pdf` `hwpx` | the **file**, as an MCP `resource` block |

It exists because a document is not its text, and a report is not a file. An
agent handed a `.hwp` or a `.docx` cannot open it — the format is a container it
has no way through. And an agent that has *written* a report has nothing to hand
to a person: text in a chat window is not a document somebody can file, print or
send on. This closes both gaps, and only those.

**What it deliberately no longer does: fetch, store, or read a PDF.** All three
left in the same change, and for one reason — each was a second copy of
something the caller already had. The outbound boundary here (an SSRF guard, a
pinned-DNS fetch, a redirect policy) was byte-for-byte the sibling's. The PDF
reader was byte-for-byte the caller's, around the same `unpdf`. And the S3
upload meant a second bucket, a second retention policy and a second AWS
credential for bytes the caller was already storing everything else in.

So this is the parser, and only the parser: the formats that genuinely need
one. A caller sending a PDF, a web page or a text file gets a refusal that names
the format and says who reads it — silence would read as "this document is
unreadable", which is a different and much more damaging claim.

## The documentation

| | |
|---|---|
| [Architecture](docs/architecture.md) | Why there is no SDK, and why a rendered file comes back as bytes |
| [Reading](docs/reading.md) | What `read_document` takes, how a format is decided, what it refuses and why |
| [Writing](docs/writing.md) | The Markdown `render_document` understands, and the images it embeds |
| [The design system](docs/design-system.md) | The palette, the type scale and the language every renderer reads from `theme.ts` |
| [The document engine](docs/document-engine.md) | How `docx`, `pdf` and `hwpx` become a report — and the Korean font inside the PDF |
| [The presentation engine](docs/presentation-engine.md) | How `pptx` becomes a planned deck |
| [Safety](docs/safety.md) | The threat model, the limits, and the two authentication modes |
| [Operations](docs/operations.md) | Configuration, the endpoints, the image, and registering it with a client |
| [Development](docs/development.md) | The commands, the release, and what the tests do and do not cover |

## Run

```bash
npm install          # Node >= 24
npm run dev          # tsx, no build step
npm test             # node --test, no test framework

MCP_API_KEY=<secret> node dist/server.js   # after npm run build
```

`PORT` (default 3000) and `MCP_API_KEY` are the only two settings, and with the
key unset the server answers anyone that can reach it —
[Safety](docs/safety.md#authentication-has-two-modes) says what that mode
assumes. Everything else about deploying it is in
[Operations](docs/operations.md).
