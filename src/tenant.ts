/**
 * Whose documents a written file belongs to.
 *
 * **The tenant comes from a header and never from tool arguments.** Agent
 * Studio stores per-server headers encrypted and merges a version's overrides
 * into them at dispatch (`application/execution/mcpTools.ts`), so the header is
 * something an operator configured. A tool argument is something the *model*
 * chose, and a model that can name its own tenant can write into another
 * project's prefix — including a model that was talked into it by a document it
 * read a moment earlier. No amount of validation fixes that; the channel is
 * wrong.
 *
 * A write with no tenant is refused rather than defaulted. A default here would
 * be a shared prefix that every misconfigured binding silently falls into.
 *
 * Reading is not scoped by it at all, and does not ask for it. The tenant
 * exists to partition storage, and reading touches no storage — requiring it
 * there would make a server that only reads unusable for the sake of a
 * boundary it never crosses.
 */

export const TENANT_HEADER = "x-document-tenant";

/** Long enough for a repository-style name, short enough to keep S3 keys sane. */
const MAX_LENGTH = 128;
/** Restricted to what is unambiguous in an S3 key path. */
const ALLOWED = /^[A-Za-z0-9][A-Za-z0-9._-]*$/;

export class TenantError extends Error {}

/**
 * Validate a tenant identifier.
 *
 * @throws TenantError when absent or malformed — the caller turns this into a
 * tool error naming the header, so a misconfigured binding says what to fix.
 */
export function parseTenant(raw: string | string[] | undefined): string {
  // A repeated header is ambiguous about which value was meant, and guessing is
  // how the wrong tenant gets picked. Node hands duplicates back as an array.
  if (Array.isArray(raw)) {
    throw new TenantError(`${TENANT_HEADER} was sent more than once`);
  }
  const value = raw?.trim();
  if (!value) {
    throw new TenantError(
      `${TENANT_HEADER} header is required to write a document — set it on the MCP server entry`,
    );
  }
  if (value.length > MAX_LENGTH) {
    throw new TenantError(`${TENANT_HEADER} is too long (max ${MAX_LENGTH} characters)`);
  }
  if (!ALLOWED.test(value)) {
    throw new TenantError(
      `${TENANT_HEADER} must start with a letter or digit and contain only letters, digits, '.', '_' or '-'`,
    );
  }
  // `.` and `..` pass the pattern and mean something to a path. Nothing else
  // does: the pattern already bars '/', so no other traversal spelling exists.
  if (value === "." || value === "..") {
    throw new TenantError(`${TENANT_HEADER} is not a valid tenant`);
  }
  return value;
}
