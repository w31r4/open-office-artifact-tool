import { pathToFileURL } from "node:url";

import { verifyPdfFile } from "./workflow.mjs";

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
  if (!args.input) throw new Error("Usage: verify-pdf.mjs --input artifact.pdf [--output-dir dir] [--native-render auto|required|off] [--pdfjs auto|required|off] [--baseline-dir dir] [--write-baseline true]");
  const result = await verifyPdfFile(args.input, { outputDir: args["output-dir"], nativeRender: args["native-render"], pdfjs: args.pdfjs, baselineDir: args["baseline-dir"], writeBaseline: args["write-baseline"] === "true", pixelThreshold: args["pixel-threshold"] ? Number(args["pixel-threshold"]) : undefined, diffAlignment: args["diff-alignment"], diffPalette: args["diff-color"] || args["diff-unchanged-color"] ? { changed: args["diff-color"], unchanged: args["diff-unchanged-color"] } : undefined, pixelRegistration: args["registration-offset"] ? { maxOffset: Number(args["registration-offset"]), minImprovementRatio: args["registration-improvement"] ? Number(args["registration-improvement"]) : undefined } : undefined, maxChars: args["max-chars"] ? Number(args["max-chars"]) : undefined });
  console.log(JSON.stringify(result.summary));
  return result;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) main().catch((error) => { console.error(error.stack || error.message); process.exitCode = 1; });
