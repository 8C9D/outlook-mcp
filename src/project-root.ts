// Claude Desktop (and other MCP clients) launch this server with an arbitrary
// cwd, so nothing may resolve paths relative to process.cwd(). The project root
// is discovered from this module's own location by walking up to the directory
// that contains package.json — which works from src/ (tsx) and dist/ (node)
// alike, whatever depth the compiled file lands at.
import { existsSync } from "node:fs";
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
