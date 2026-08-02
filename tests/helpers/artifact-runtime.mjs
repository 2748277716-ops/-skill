import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
export const inputReaderModule = "jszip";

export function assertRuntimeAvailable() {
  const resolvedModules = {};
  for (const name of ["@oai/artifact-tool", inputReaderModule]) {
    try {
      resolvedModules[name] = require.resolve(name);
    } catch {
      throw new Error("Missing bundled module: " + name);
    }
  }
  return resolvedModules;
}
