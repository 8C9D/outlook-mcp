// Claude Desktop (and other MCP clients) launch this server with an arbitrary
// cwd, so nothing may resolve paths relative to process.cwd(). The project root
// is discovered from this module's own location by walking up to the directory
// that contains package.json — which works from src/ (tsx) and dist/ (node)
// alike, whatever depth the compiled file lands at.
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

function findProjectRoot(startDir: string): string {
  let dir = startDir;
  for (;;) {
    if (existsSync(path.join(dir, "package.json"))) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) {
      throw new Error(`Could not find package.json in any directory above ${startDir}`);
    }
    dir = parent;
  }
}

export const PROJECT_ROOT = findProjectRoot(path.dirname(fileURLToPath(import.meta.url)));

/**
 * The public origin of this checkout's deployed Worker. It is written down in
 * one place: PUBLIC_BASE_URL in the gitignored .env (the hostname names the
 * owner's account, so it is kept out of the tracked wrangler.jsonc, and
 * scripts/deploy.sh injects it at deploy time). Node-side callers only (the
 * test harnesses and the doctor, whose src/auth.ts import loads .env); inside
 * the Worker the same value arrives as a binding.
 *
 * Resolution order: MCP_REMOTE_URL (to aim a suite at another deployment),
 * then PUBLIC_BASE_URL from the environment, then a `vars` entry in
 * wrangler.jsonc for a checkout that chose to keep it there. Returns undefined
 * when none is set, or the config cannot be parsed by this crude
 * comment-stripper, so a caller can degrade rather than fail.
 */
export function deployedBaseUrl(): string | undefined {
  const override = process.env.MCP_REMOTE_URL?.trim();
  if (override) return override.replace(/\/+$/, "");
  const fromEnv = process.env.PUBLIC_BASE_URL?.trim();
  if (fromEnv) return fromEnv.replace(/\/+$/, "");
  try {
    const raw = readFileSync(path.join(PROJECT_ROOT, "wrangler.jsonc"), "utf8");
    const config = JSON.parse(raw.replace(/^\s*\/\/.*$/gm, "")) as {
      vars?: { PUBLIC_BASE_URL?: string };
    };
    const configured = config.vars?.PUBLIC_BASE_URL?.trim();
    return configured ? configured.replace(/\/+$/, "") : undefined;
  } catch {
    return undefined;
  }
}
