import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

const TEXT_EXTENSIONS = new Set([".js", ".json", ".mjs", ".yaml", ".yml"]);

export function normalizeGeneratedTree(root, options = {}) {
  for (const file of filesUnder(path.resolve(root))) {
    if (!TEXT_EXTENSIONS.has(path.extname(file))) continue;
    const source = fs.readFileSync(file, "utf8");
    const withoutSourceMaps = options.stripSourceMapComments
      ? source.replace(/^\s*\/\/[#@]\s*sourceMappingURL=.*$/gm, "")
      : source;
    const normalized = `${withoutSourceMaps.replace(/[ \t]+$/gm, "").trimEnd()}\n`;
    if (normalized !== source) fs.writeFileSync(file, normalized);
  }
}

function filesUnder(target) {
  const stat = fs.statSync(target);
  if (!stat.isDirectory()) return [target];
  return fs.readdirSync(target, { withFileTypes: true }).flatMap((entry) => {
    const child = path.join(target, entry.name);
    return entry.isDirectory() ? filesUnder(child) : [child];
  });
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  for (const target of process.argv.slice(2)) normalizeGeneratedTree(target);
}
