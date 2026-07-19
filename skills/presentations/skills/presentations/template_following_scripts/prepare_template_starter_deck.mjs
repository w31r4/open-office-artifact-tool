#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";

import { parseArgs, requireArg } from "../container_tools/artifact_tool_utils.mjs";

function usage() {
  return [
    "Usage:",
    "  node template_following_scripts/prepare_template_starter_deck.mjs --workspace <dir> --pptx <source.pptx> --map <template-frame-map.json> --out <starter.pptx> [options]",
    "",
    "Options:",
    "  --preview-dir <dir>     Render starter slide PNGs. Defaults to <workspace>/template-starter-preview.",
    "  --layout-dir <dir>      Write starter layout JSON. Defaults to <workspace>/template-starter-layout.",
    "  --inspect <path>        template-inspect.ndjson. Defaults to <workspace>/template-inspect/template-inspect.ndjson.",
    "  --contact-sheet <path>  Optional PNG contact sheet path.",
    "  --scale <n>            Render scale. Defaults to 1.",
    "",
    "This command currently performs a read-only path/input preflight, then fails",
    "closed because broad imported-slide graph clone/delete is not available.",
  ].join("\n");
}

function isWithin(child, parent) {
  const relative = path.relative(parent, child);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    console.log(usage());
    return;
  }

  const workspaceDir = path.resolve(requireArg(args, "workspace"));
  const pptxPath = path.resolve(requireArg(args, "pptx"));
  const mapPath = path.resolve(requireArg(args, "map"));
  const out = path.resolve(requireArg(args, "out"));
  const previewDir = args["preview-dir"]
    ? path.resolve(args["preview-dir"])
    : path.join(workspaceDir, "template-starter-preview");
  const layoutDir = args["layout-dir"]
    ? path.resolve(args["layout-dir"])
    : path.join(workspaceDir, "template-starter-layout");
  const contactSheetPath = args["contact-sheet"] ? path.resolve(args["contact-sheet"]) : undefined;
  const scale = args.scale ? Number.parseFloat(args.scale) : 1;

  if (!Number.isFinite(scale) || scale <= 0) {
    throw new Error("--scale must be a positive number");
  }
  for (const writePath of [out, previewDir, layoutDir, contactSheetPath].filter(Boolean)) {
    if (!isWithin(writePath, workspaceDir)) {
      throw new Error(`Refusing to write starter artifacts outside workspace: ${writePath}`);
    }
  }

  const sourceStat = await fs.stat(pptxPath).catch(() => undefined);
  if (!sourceStat?.isFile()) {
    throw new Error(`Missing source PPTX: ${pptxPath}`);
  }
  const mapStat = await fs.stat(mapPath).catch(() => undefined);
  if (!mapStat?.isFile()) {
    throw new Error(`Missing template frame map: ${mapPath}`);
  }

  throw new Error(
    "Template starter generation requires source-preserving imported-slide duplication plus broad graph deletion. " +
    "The current OpenChestnut codec has only an isolated layout-only delete profile and an unchanged shape/inline-table/image/recursive-group/closed-notes clone that requires export/reimport; it will not reconstruct or share a broad clone graph. No output was written.",
  );
}

main().catch((error) => {
  console.error(error.stack || error.message || String(error));
  console.error(usage());
  process.exit(1);
});
