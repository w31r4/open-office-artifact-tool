import { pathToFileURL } from "node:url";

import { verifyWorkbookFile } from "./workflow.mjs";

function parseArgs(argv) {
  const options = {};
  for (let index = 0; index < argv.length; index += 1) {
    const key = argv[index];
    if (!key.startsWith("--")) continue;
    options[key.slice(2)] = argv[index + 1];
    index += 1;
  }
  return options;
}

export async function main(argv = process.argv.slice(2)) {
  const args = parseArgs(argv);
  if (!args.input) throw new Error("Usage: verify-workbook.mjs --input workbook.xlsx [--output-dir dir] [--sheet name] [--range A1:D20] [--render-format svg|png|webp|jpeg|pdf] [--all-sheets true] [--native-render auto|required|off] [--baseline-dir dir] [--write-baseline true]");
  const result = await verifyWorkbookFile(args.input, {
    outputDir: args["output-dir"],
    sheetName: args.sheet,
    range: args.range,
    renderFormat: args["render-format"],
    baselineDir: args["baseline-dir"],
    writeBaseline: args["write-baseline"] === "true",
    pixelThreshold: args["pixel-threshold"] ? Number(args["pixel-threshold"]) : undefined,
    diffAlignment: args["diff-alignment"],
    diffPalette: args["diff-color"] || args["diff-unchanged-color"] ? { changed: args["diff-color"], unchanged: args["diff-unchanged-color"] } : undefined,
    pixelRegistration: args["registration-offset"] ? { maxOffset: Number(args["registration-offset"]), minImprovementRatio: args["registration-improvement"] ? Number(args["registration-improvement"]) : undefined } : undefined,
    allSheets: args["all-sheets"] === "true",
    nativeRender: args["native-render"],
    maxChars: args["max-chars"] ? Number(args["max-chars"]) : undefined,
  });
  console.log(JSON.stringify(result.summary));
  return result;
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(error.stack || error.message);
    process.exitCode = 1;
  });
}
