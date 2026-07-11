import { pathToFileURL } from "node:url";

import { runSpreadsheetFixture } from "./workflow.mjs";

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
  if (!args.fixture) throw new Error("Usage: run-fixture.mjs --fixture fixture.json [--output-dir dir] [--render-format svg|png|webp|jpeg|pdf] [--all-sheets true] [--native-render auto|required|off] [--baseline-dir dir] [--write-baseline true]");
  const result = await runSpreadsheetFixture(args.fixture, {
    outputDir: args["output-dir"],
    sheetName: args.sheet,
    range: args.range,
    renderFormat: args["render-format"],
    baselineDir: args["baseline-dir"],
    writeBaseline: args["write-baseline"] === "true",
    pixelThreshold: args["pixel-threshold"] ? Number(args["pixel-threshold"]) : undefined,
    diffAlignment: args["diff-alignment"],
    diffPalette: args["diff-color"] || args["diff-unchanged-color"] ? { changed: args["diff-color"], unchanged: args["diff-unchanged-color"] } : undefined,
    allSheets: args["all-sheets"] === "true",
    nativeRender: args["native-render"],
  });
  console.log(JSON.stringify({ fixture: result.fixture.name, workbook: result.workbookPath, qa: result.qa.summary }));
  return result;
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(error.stack || error.message);
    process.exitCode = 1;
  });
}
