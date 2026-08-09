/**
 * Everything this server reads from its environment, validated once at boot.
 *
 * Fail-fast rather than on first use: a missing bucket name should stop a
 * rollout at the readiness probe, not surface as a tool error inside somebody's
 * agent run half an hour later. The bucket is required even though reading
 * never touches it — half a server is not a state worth being able to deploy,
 * and `write_document` failing on the one call that needed it is the version of
 * this that gets discovered by a user rather than by a rollout.
 *
 * The set is deliberately small. Every knob here is a way for two deployments
 * to behave differently for a reason nobody wrote down, so the size ceilings
 * and the extraction limits are constants in `limits.ts` rather than
 * environment variables. What remains is what a deployment genuinely has to
 * say: where its bucket is, how long a link should live, and who may call it.
 */

export interface Config {
  port: number;
  /** Shared secret callers must present. Unset means no authentication — see `auth.ts`. */
  apiKey: string | undefined;
  region: string;
  /** Bucket the written documents land in. */
  bucket: string;
  /** Key prefix inside that bucket, ahead of the tenant. */
  prefix: string;
  /** How long a download link stays valid. */
  downloadTtlSeconds: number;
}

/**
 * SigV4's own ceiling on a presigned URL. Asking for more does not produce a
 * longer-lived link; it produces one S3 rejects, which would surface as a dead
 * download rather than as a configuration error.
 */
export const MAX_DOWNLOAD_TTL_SECONDS = 7 * 24 * 60 * 60;

export class ConfigError extends Error {}

function required(env: NodeJS.ProcessEnv, name: string): string {
  const value = env[name]?.trim();
  if (!value) {
    throw new ConfigError(`${name} is required`);
  }
  return value;
}

function integer(
  env: NodeJS.ProcessEnv,
  name: string,
  fallback: number,
  max?: number,
): number {
  const raw = env[name]?.trim();
  if (!raw) {
    return fallback;
  }
  const value = Number(raw);
  if (!Number.isInteger(value) || value <= 0) {
    throw new ConfigError(`${name} must be a positive integer, got "${raw}"`);
  }
  if (max !== undefined && value > max) {
    throw new ConfigError(`${name} must be at most ${max}, got "${raw}"`);
  }
  return value;
}

/**
 * A key prefix, normalised to have no leading or trailing slash.
 *
 * The empty string is allowed and means "no prefix" — `store.ts` joins the
 * parts, so an empty one simply drops out. What is not allowed is anything a
 * path would read as movement: the prefix is joined with a tenant and a
 * filename, and `..` there would let one deployment's documents be written
 * under another's.
 */
function keyPrefix(env: NodeJS.ProcessEnv, name: string, fallback: string): string {
  const raw = env[name];
  const value = (raw === undefined ? fallback : raw).trim().replace(/^\/+|\/+$/g, "");
  if (value === "") {
    return "";
  }
  if (!/^[A-Za-z0-9][A-Za-z0-9._\-/]*$/.test(value)) {
    throw new ConfigError(
      `${name} must start with a letter or digit and contain only letters, digits, '.', '_', '-' or '/'`,
    );
  }
  if (value.split("/").some((segment) => segment === "" || segment === "." || segment === "..")) {
    throw new ConfigError(`${name} must not contain empty, '.' or '..' path segments`);
  }
  return value;
}

export function loadConfig(env: NodeJS.ProcessEnv = process.env): Config {
  return {
    port: integer(env, "PORT", 3000),
    apiKey: env.MCP_API_KEY?.trim() || undefined,
    region: env.AWS_REGION?.trim() || "ap-northeast-2",
    bucket: required(env, "DOCUMENT_BUCKET"),
    prefix: keyPrefix(env, "DOCUMENT_PREFIX", "documents"),
    downloadTtlSeconds: integer(
      env,
      "DOWNLOAD_TTL_SECONDS",
      MAX_DOWNLOAD_TTL_SECONDS,
      MAX_DOWNLOAD_TTL_SECONDS,
    ),
  };
}
