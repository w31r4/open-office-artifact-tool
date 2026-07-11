import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { DocumentModel, FileBlob, renderArtifact } from "open-office-artifact-tool";
import { createLibreOfficeRenderer } from "open-office-artifact-tool/renderers/libreoffice";
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";
import { createSharpRenderer } from "open-office-artifact-tool/renderers/sharp";
import { createCanvasRenderer } from "open-office-artifact-tool/renderers/canvas";

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
const mockLibreOffice = path.join(tempDir, "mock-libreoffice.mjs");
await fs.writeFile(mockLibreOffice, `
import fs from 'node:fs/promises';
import path from 'node:path';
const outDir = process.argv[process.argv.indexOf('--outdir') + 1];
const inputPath = process.argv.at(-1);
const base = path.basename(inputPath, path.extname(inputPath));
await fs.mkdir(outDir, { recursive: true });
await fs.writeFile(path.join(outDir, base + '.pdf'), new TextEncoder().encode('%PDF-libreoffice-mock'));
`, "utf8");
const libreOfficeRenderer = createLibreOfficeRenderer({ command: process.execPath, args: [mockLibreOffice], timeoutMs: 10_000 });
const docxInput = new FileBlob(new Uint8Array([1, 2, 3]), { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
const officePdf = await libreOfficeRenderer({ input: docxInput, inputType: docxInput.type, outputType: "application/pdf", format: "pdf", artifactKind: "document", options: { pageIndex: 0 } });
assert.equal(officePdf.type, "application/pdf");
assert.equal(officePdf.metadata.renderer, "libreoffice");
assert.equal(officePdf.metadata.inputType, docxInput.type);
assert.match(await officePdf.text(), /%PDF-libreoffice-mock/);

function fakeCanvasLib({ width = 120, height = 80 } = {}) {
  return {
    async loadImage(buf) {
      assert.ok(Buffer.isBuffer(buf));
      return { width, height, naturalWidth: width, naturalHeight: height };
    },
    createCanvas(w, h) {
      assert.ok(w > 0 && h > 0, "canvas renderer must create a positive-size canvas");
      return {
        width: w,
        height: h,
        getContext() {
          return {
            fillStyle: null,
            fillRect() {},
            drawImage(image, dx, dy, dw, dh) {
              assert.equal(dx, 0);
              assert.equal(dy, 0);
              assert.equal(dw, w);
              assert.equal(dh, h);
            },
          };
        },
        toBuffer(mime, opts) {
          if (mime === "image/jpeg") return Buffer.from([0xff, 0xd8, 0xff, 0xd9]);
          return Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a]);
        },
      };
    },
  };
}

const canvasDoc = DocumentModel.create({ paragraphs: ["Canvas renderer smoke"] });
const canvasRenderer = createCanvasRenderer({ canvas: fakeCanvasLib(), background: "white" });
const canvasPng = await renderArtifact(canvasDoc, { format: "png", renderer: canvasRenderer });
assert.equal(canvasPng.type, "image/png");
assert.equal(canvasPng.metadata.renderer, "canvas");
assert.equal(canvasPng.metadata.inputType, "image/svg+xml");
assert.equal(canvasPng.metadata.format, "png");
assert.ok(canvasPng.metadata.width > 0 && canvasPng.metadata.height > 0);
assert.deepEqual([...canvasPng.bytes.slice(0, 4)], [0x89, 0x50, 0x4e, 0x47]);
const canvasJpeg = await renderArtifact(canvasDoc, { format: "jpeg", renderer: canvasRenderer });
assert.equal(canvasJpeg.type, "image/jpeg");
assert.equal(canvasJpeg.metadata.format, "jpeg");
assert.deepEqual([...canvasJpeg.bytes.slice(0, 2)], [0xff, 0xd8]);
const svgBlob = new FileBlob('<svg xmlns="http://www.w3.org/2000/svg" width="40" height="30"><rect width="40" height="30"/></svg>', { type: "image/svg+xml" });
const canvasFromBlob = await canvasRenderer({ input: svgBlob, inputType: "image/svg+xml", outputType: "image/png", format: "png", artifactKind: "document" });
assert.equal(canvasFromBlob.metadata.renderer, "canvas");
assert.equal(canvasFromBlob.metadata.width, 120);
assert.equal(canvasFromBlob.metadata.height, 80);
await assert.rejects(
  () => canvasRenderer({ input: svgBlob, inputType: "image/svg+xml", outputType: "image/webp", format: "webp" }),
  /supported formats are png and jpeg/,
);
await assert.rejects(
  () => canvasRenderer({ input: new FileBlob(new Uint8Array([1]), { type: "application/pdf" }), inputType: "application/pdf", outputType: "image/png", format: "png" }),
  /supports SVG, PNG, JPEG, and WebP input/,
);
const smallCanvasLib = fakeCanvasLib({ width: 0, height: 0 });
delete smallCanvasLib.loadImage;
smallCanvasLib.loadImage = async () => ({ width: 0, height: 0, naturalWidth: 0, naturalHeight: 0 });
const sizedRenderer = createCanvasRenderer({ canvas: smallCanvasLib, background: "white" });
const sizedPng = await sizedRenderer({ input: svgBlob, inputType: "image/svg+xml", outputType: "image/png", format: "png", artifactKind: "document" });
assert.equal(sizedPng.metadata.width, 40);
assert.equal(sizedPng.metadata.height, 30);

await fs.rm(tempDir, { recursive: true, force: true });

console.log("renderer adapters smoke ok");
