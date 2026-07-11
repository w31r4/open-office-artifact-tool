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
  if (!args.input) throw new Error("Usage: verify-workbook.mjs --input workbook.xlsx [--output-dir dir] [--sheet name] [--range A1:D20] [--render-format svg|png|webp|jpeg|pdf]");
  const result = await verifyWorkbookFile(args.input, {
    outputDir: args["output-dir"],
    sheetName: args.sheet,
    range: args.range,
    renderFormat: args["render-format"],
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

