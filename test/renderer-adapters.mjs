import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { DocumentModel, FileBlob, renderArtifact } from "open-office-artifact-tool";
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";
import { createSharpRenderer } from "open-office-artifact-tool/renderers/sharp";

function fakeSharp(input) {
  assert.ok(Buffer.isBuffer(input));
  return {
    resize() { return this; },
    flatten() { return this; },
    png() { this.format = "png"; return this; },
    webp() { this.format = "webp"; return this; },
    jpeg() { this.format = "jpeg"; return this; },
    async toBuffer() {
      if (this.format === "webp") return Buffer.from("RIFFxxxxWEBP", "utf8");
      if (this.format === "jpeg") return Buffer.from([0xff, 0xd8, 0xff, 0xd9]);
      return Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a]);
    },
  };
}

const document = DocumentModel.create({ paragraphs: ["Sharp renderer smoke"] });
const sharpRenderer = createSharpRenderer({ sharp: fakeSharp, flatten: true });
const png = await renderArtifact(document, { format: "png", renderer: sharpRenderer });
assert.equal(png.type, "image/png");
assert.equal(png.metadata.renderer, "sharp");
assert.equal(png.metadata.inputType, "image/svg+xml");
assert.deepEqual([...png.bytes.slice(0, 4)], [0x89, 0x50, 0x4e, 0x47]);
const webp = await renderArtifact(document, { format: "webp", renderer: sharpRenderer });
assert.equal(webp.type, "image/webp");
assert.equal(new TextDecoder().decode(webp.bytes.slice(0, 4)), "RIFF");
await assert.rejects(
  () => sharpRenderer({ input: new FileBlob(new Uint8Array([1]), { type: "application/pdf" }), inputType: "application/pdf", outputType: "image/png", format: "png" }),
  /supports SVG, PNG, JPEG, and WebP input/,
);

const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-poppler-test-"));
const mockPoppler = path.join(tempDir, "mock-poppler.mjs");
await fs.writeFile(mockPoppler, `
import fs from 'node:fs/promises';
const outputPrefix = process.argv.at(-1);
await fs.writeFile(outputPrefix + '.png', new Uint8Array([0x89,0x50,0x4e,0x47,0x0d,0x0a]));
`, "utf8");
const popplerRenderer = createPopplerRenderer({ command: process.execPath, args: [mockPoppler], timeoutMs: 10_000 });
const pdfInput = new FileBlob(new TextEncoder().encode("%PDF-1.4\n%%EOF"), { type: "application/pdf" });
const raster = await popplerRenderer({ input: pdfInput, inputType: "application/pdf", outputType: "image/png", format: "png", artifactKind: "pdf", options: { pageIndex: 0 } });
assert.equal(raster.type, "image/png");
assert.equal(raster.metadata.renderer, "poppler");
assert.equal(raster.metadata.page, 1);
assert.deepEqual([...raster.bytes.slice(0, 4)], [0x89, 0x50, 0x4e, 0x47]);
await assert.rejects(
  () => popplerRenderer({ input: new FileBlob("<svg/>", { type: "image/svg+xml" }), inputType: "image/svg+xml", outputType: "image/png", format: "png" }),
  /supports application\/pdf input/,
);
await fs.rm(tempDir, { recursive: true, force: true });

console.log("renderer adapters smoke ok");
