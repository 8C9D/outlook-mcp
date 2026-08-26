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
 * The public origin of this checkout's deployed Worker, read from the
 * PUBLIC_BASE_URL var in wrangler.jsonc — the one place the deployed hostname
 * is written down. Node-side callers only (the test harnesses and the doctor);
 * inside the Worker the same value arrives as a binding.
 *
 * Returns undefined for a checkout with no wrangler.jsonc, no var, or a config
 * this crude comment-stripper cannot parse, so a caller can degrade rather than
 * fail. MCP_REMOTE_URL overrides it for pointing a suite at another deployment.
 */
export function deployedBaseUrl(): string | undefined {
  const override = process.env.MCP_REMOTE_URL?.trim();
  if (override) return override.replace(/\/+$/, "");
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
