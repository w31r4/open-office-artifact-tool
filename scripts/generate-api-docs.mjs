import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { HELP_CATALOG } from "../src/index.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputPath = path.join(repoRoot, "docs", "api.md");

const byKind = new Map();
for (const item of HELP_CATALOG) {
  const list = byKind.get(item.artifactKind) || [];
  list.push(item);
  byKind.set(item.artifactKind, list);
}

const lines = [
  "# API catalog",
  "",
  "Generated from `HELP_CATALOG` in `src/index.mjs`.",
  "",
];

for (const artifactKind of [...byKind.keys()].sort()) {
  lines.push(`## ${artifactKind}`);
  lines.push("");
  lines.push("| Name | Kind | Summary |");
  lines.push("| --- | --- | --- |");
  for (const item of byKind.get(artifactKind).sort((a, b) => a.name.localeCompare(b.name))) {
    const summary = String(item.summary || "").replaceAll("|", "\\|");
    lines.push(`| \`${item.name}\` | ${item.kind} | ${summary} |`);
  }
  lines.push("");
}

await fs.writeFile(outputPath, `${lines.join("\n")}\n`, "utf8");
console.log(outputPath);
