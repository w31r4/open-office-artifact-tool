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

function escapeTable(value) {
  return String(value || "").replaceAll("|", "\\|");
}

function hasDetails(item) {
  return Boolean(item.examples?.length || item.options?.length || item.params?.length || item.notes?.length || item.returns || item.schema);
}

function detailList(label, values) {
  if (!values?.length) return [];
  return [`**${label}:**`, "", ...values.map((value) => `- ${value}`), ""];
}

for (const artifactKind of [...byKind.keys()].sort()) {
  const items = byKind.get(artifactKind).sort((a, b) => a.name.localeCompare(b.name));
  lines.push(`## ${artifactKind}`);
  lines.push("");
  lines.push("| Name | Kind | Summary |");
  lines.push("| --- | --- | --- |");
  for (const item of items) {
    lines.push(`| \`${item.name}\` | ${item.kind} | ${escapeTable(item.summary)} |`);
  }
  lines.push("");
  const detailed = items.filter(hasDetails);
  if (detailed.length) {
    lines.push(`### ${artifactKind} details`);
    lines.push("");
    for (const item of detailed) {
      lines.push(`#### \`${item.name}\``);
      lines.push("");
      lines.push(item.summary || "");
      lines.push("");
      lines.push(...detailList("Examples", item.examples));
      lines.push(...detailList("Options", item.options || item.params));
      if (item.returns) lines.push("**Returns:**", "", item.returns, "");
      if (item.schema) lines.push("**Schema:**", "", "```json", JSON.stringify(item.schema, null, 2), "```", "");
      lines.push(...detailList("Notes", item.notes));
    }
  }
}

await fs.writeFile(outputPath, `${lines.join("\n")}\n`, "utf8");
console.log(outputPath);
