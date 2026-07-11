import assert from "node:assert/strict";

import { DocumentModel, FileBlob, renderArtifact } from "open-office-artifact-tool";
import { createPlaywrightRenderer, playwrightRenderer } from "open-office-artifact-tool/renderers/playwright";

function shouldSkipPlaywright(error) {
  return /optional peer dependency|Cannot find package 'playwright'|Executable doesn't exist|playwright install|browserType\.launch/i.test(String(error?.message || error));
}

function assertPng(bytes) {
  assert.deepEqual([...bytes.slice(0, 4)], [0x89, 0x50, 0x4e, 0x47]);
}

function assertWebp(bytes) {
  assert.equal(new TextDecoder().decode(bytes.slice(0, 4)), "RIFF");
  assert.equal(new TextDecoder().decode(bytes.slice(8, 12)), "WEBP");
}

function assertJpeg(bytes) {
  assert.deepEqual([...bytes.slice(0, 3)], [0xff, 0xd8, 0xff]);
}

function assertPdf(bytes) {
  assert.equal(new TextDecoder().decode(bytes.slice(0, 4)), "%PDF");
}

const document = DocumentModel.create({ paragraphs: ["Playwright render smoke"] });
const renderer = createPlaywrightRenderer({ viewport: { width: 360, height: 220 }, deviceScaleFactor: 1, timeout: 20_000 });
assert.equal(typeof playwrightRenderer, "function");
await assert.rejects(
  () => renderer({ input: new FileBlob(new Uint8Array([0]), { type: "image/png" }), inputType: "image/png", outputType: "image/png", format: "png" }),
  /supports SVG or HTML input/,
);

let png;
try {
  png = await renderArtifact(document, { format: "png", renderer });
} catch (error) {
  if (!process.env.PLAYWRIGHT_RENDERER_TESTS && shouldSkipPlaywright(error)) {
    console.log(`playwright renderer smoke skipped: ${error.message}`);
    process.exit(0);
  }
  throw error;
}

assert.equal(png.type, "image/png");
assert.equal(png.metadata.renderer, "playwright");
assert.equal(png.metadata.artifactKind, "document");
assert.equal(png.metadata.format, "png");
assert.equal(png.metadata.inputType, "image/svg+xml");
assert.deepEqual(png.metadata.viewport, { width: 360, height: 220 });
assertPng(png.bytes);

const webp = await renderArtifact(document, { format: "webp", renderer });
assert.equal(webp.type, "image/webp");
assert.equal(webp.metadata.renderer, "playwright");
assertWebp(webp.bytes);

const jpeg = await renderArtifact(document, { format: "jpeg", renderer });
assert.equal(jpeg.type, "image/jpeg");
assert.equal(jpeg.metadata.renderer, "playwright");
assertJpeg(jpeg.bytes);

const pdf = await renderArtifact(document, { format: "pdf", renderer });
assert.equal(pdf.type, "application/pdf");
assert.equal(pdf.metadata.renderer, "playwright");
assertPdf(pdf.bytes);

console.log("playwright renderer smoke ok");
