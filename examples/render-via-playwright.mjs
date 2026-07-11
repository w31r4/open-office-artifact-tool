import os from "node:os";
import path from "node:path";

import { DocumentModel, renderArtifact } from "open-office-artifact-tool";
import { createPlaywrightRenderer } from "open-office-artifact-tool/renderers/playwright";

const outputDir = process.env.OUTPUT_DIR || path.join(os.tmpdir(), "open-office-artifact-examples");
const document = DocumentModel.create({ paragraphs: ["Rendered through Playwright", "Install playwright + chromium to produce PNG/WebP/PDF previews."] });
const renderer = createPlaywrightRenderer({ viewport: { width: 900, height: 1200 }, deviceScaleFactor: 1 });

try {
  const png = await renderArtifact(document, { format: "png", renderer });
  await png.save(path.join(outputDir, "playwright-render.png"));
  console.log(`saved ${path.join(outputDir, "playwright-render.png")}`);
} catch (error) {
  if (/playwright|Chromium|Executable doesn't exist|install/i.test(String(error?.message || error))) {
    console.log(`skipped Playwright render example: ${error.message}`);
  } else {
    throw error;
  }
}
