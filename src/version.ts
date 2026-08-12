/**
 * What this server calls itself.
 *
 * Its own file because two places need it and neither should reach the other:
 * `server.ts` tells a client what it connected to, and `write/hwpx.ts` stamps
 * it into a document's metadata as the application that produced it. (It was a
 * User-Agent as well, until fetching left.)
 *
 * `SERVER_VERSION` restates package.json's `version`. A test pins the two
 * together: nothing else would notice them drifting, and what a client is told
 * it is talking to would quietly stop being true.
 */

export const SERVER_NAME = "mcp-document";
export const SERVER_VERSION = "0.2.2";
