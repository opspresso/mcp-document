# Operations

## Configure

```
PORT=3000                      default 3000
MCP_API_KEY=<secret>           unset means no authentication — see below
```

Two settings, and it used to be six. The bucket, its prefix, the region and the
download TTL left with the storage; the pod needs no AWS role at all now, and no
network access beyond the port it listens on.

Every other number — the size ceilings, the extraction limits — is a constant in
`src/limits.ts` rather than a knob. Each one would otherwise be a way for two
deployments to behave differently for a reason nobody wrote down.

With `MCP_API_KEY` unset the server answers anyone that can reach it, which is
only safe while nothing routes to it from outside the cluster —
[Safety](safety.md#authentication-has-two-modes) says what that mode assumes.

## Run

    MCP_API_KEY=<secret> node dist/server.js   # authenticated
    node dist/server.js                        # open — cluster-internal only

    POST   /mcp      JSON-RPC; Authorization: Bearer <MCP_API_KEY> when a key is set
    DELETE /mcp      session teardown; 204, since this server holds no session
    GET    /health   liveness

A tag publishes a `linux/amd64` image to GHCR and to a private ECR mirror,
creates a GitHub Release whose notes are the commit subjects, and dispatches the
released version to the GitOps repository that deploys it. The image runs as the
unprivileged `node` user and needs no writable volume:

    docker run -e MCP_API_KEY=<secret> -p 3000:3000 \
      ghcr.io/opspresso/mcp-document:latest

## Connect from an MCP client

```json
{
  "mcpServers": {
    "document": {
      "url": "https://<host>/mcp",
      "headers": { "Authorization": "Bearer <MCP_API_KEY>" }
    }
  }
}
```

No `uv` or local command is required. Clients that only support local `stdio`
servers need an HTTP-to-stdio bridge.

## Register in AgentDure

Tools → register with the URL ending in `/mcp` and a header
`Authorization: Bearer <MCP_API_KEY>`. No tenant header: there is nothing left
to partition.

`read_document` returns a `text` block, which flows straight into the
turn. `render_document` returns a `resource` block carrying the bytes, which
AgentDure stores as an artifact and delivers to the user — so the model should
describe what it wrote rather than offer a link.
