import fs from "node:fs";
import path from "node:path";

export const bundledNodeModules = String.raw`C:\Users\86182\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules`;

export function assertRuntimeAvailable() {
  for (const name of ["@oai/artifact-tool", "jszip"]) {
    const target = path.join(bundledNodeModules, ...name.split("/"));
    if (!fs.existsSync(target)) throw new Error(`Missing bundled module: ${name}`);
  }
  return bundledNodeModules;
}
