/**
 * Everything this server reads from its environment, validated once at boot.
 *
 * Two settings, and it used to be six. The bucket, its prefix, the region and
 * the download TTL all left with the storage: the renderer hands its bytes back
 * over MCP now, so the caller keeps them beside every other byte one of its runs
 * produced — and owns the retention and the delete button that go with that.
 * One fewer process holding an AWS credential, and one fewer place where two
 * deployments can quietly disagree about how long a document lives.
 *
 * What remains is what a deployment genuinely has to say: which port, and who
 * may call it. The size ceilings and the extraction limits stay constants in
 * `limits.ts` rather than knobs, for the same reason they always were — every
 * one of them is a way for two deployments to behave differently for a reason
 * nobody wrote down.
 */

export interface Config {
  port: number;
  /** Shared secret callers must present. Unset means no authentication — see `auth.ts`. */
  apiKey: string | undefined;
}

export class ConfigError extends Error {}

function integer(env: NodeJS.ProcessEnv, name: string, fallback: number): number {
  const raw = env[name]?.trim();
  if (!raw) {
    return fallback;
  }
  const value = Number(raw);
  if (!Number.isInteger(value) || value <= 0) {
    throw new ConfigError(`${name} must be a positive integer, got "${raw}"`);
  }
  return value;
}

export function loadConfig(env: NodeJS.ProcessEnv = process.env): Config {
  return {
    port: integer(env, "PORT", 3000),
    apiKey: env.MCP_API_KEY?.trim() || undefined,
  };
}
