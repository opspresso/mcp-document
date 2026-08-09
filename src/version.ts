/**
 * What this server calls itself.
 *
 * Its own file because both ends need it and they cannot reach each other:
 * `server.ts` tells a client what it connected to, and `source.ts` tells every
 * site it fetches the same thing in a User-Agent.
 *
 * `SERVER_VERSION` restates package.json's `version`. A test pins the two
 * together: nothing else would notice them drifting, and what a client is told
 * it is talking to would quietly stop being true.
 */

export const SERVER_NAME = "mcp-document";
export const SERVER_VERSION = "0.1.0";
