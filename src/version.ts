/**
 * What this server calls itself.
 *
 * Its own file because two kinds of caller need it and neither should reach the
 * other: `server.ts` tells a client what it connected to, and all four renderers
 * stamp it into a document's metadata as the thing that produced it. (It was a
 * User-Agent as well, until fetching left.)
 *
 * `SERVER_VERSION` restates package.json's `version`, and `npm version` keeps
 * the two in step through `scripts/sync-version.mjs`. `version.test.ts` is the
 * backstop for a release made some other way: nothing at runtime compares them,
 * so without it what a client is told it is talking to could quietly stop being
 * true.
 */

export const SERVER_NAME = "mcp-document";
export const SERVER_VERSION = "0.5.0";

/**
 * How a produced file names what made it.
 *
 * One string, because every format has exactly one field for this and they all
 * want a human-readable name rather than a structured pair — PDF's `Producer`,
 * OOXML's `<Application>`, OWPML's `application`. When somebody turns up with a
 * document that renders oddly, this is what tells them which release wrote it.
 */
export const PRODUCER = `${SERVER_NAME} ${SERVER_VERSION}`;
