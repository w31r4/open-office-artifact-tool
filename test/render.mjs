import assert from "node:assert/strict";
import { deflateSync } from "node:zlib";

import {
  DocumentModel,
  FileBlob,
  PdfArtifact,
  Presentation,
  renderArtifact,
  visualQaArtifact,
  Workbook,
} from "open-office-artifact-tool";

function pngChunk(type, data = Buffer.alloc(0)) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  return Buffer.concat([length, Buffer.from(type, "ascii"), data, Buffer.alloc(4)]);
}

function makePng(width, height, rgba) {
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const rowBytes = width * 4;
  const raw = Buffer.alloc((rowBytes + 1) * height);
  for (let y = 0; y < height; y += 1) {
    raw[y * (rowBytes + 1)] = 0;
    Buffer.from(rgba).copy(raw, y * (rowBytes + 1) + 1, y * rowBytes, (y + 1) * rowBytes);
  }
  return new Uint8Array(Buffer.concat([signature, pngChunk("IHDR", ihdr), pngChunk("IDAT", deflateSync(raw)), pngChunk("IEND")]));
}

const blackPixelPng = makePng(1, 1, [0, 0, 0, 255]);
const whitePixelPng = makePng(1, 1, [255, 255, 255, 255]);

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet1");
sheet.getRange("A1:B2").values = [["A", "B"], [1, 2]];
const workbookPreview = await renderArtifact(workbook, { sheetName: "Sheet1", range: "A1:B2" });
assert.equal(workbookPreview.type, "image/svg+xml");
assert.equal(workbookPreview.metadata.artifactKind, "workbook");
assert.equal(workbookPreview.metadata.sheetName, "Sheet1");
assert.match(await workbookPreview.text(), /A/);

const presentation = Presentation.create({ slideSize: { width: 320, height: 180 } });
const slide = presentation.slides.add();
slide.shapes.add({ name: "render-shape", position: { left: 20, top: 20, width: 120, height: 60 }, text: "Render me" });
const presentationPreview = await renderArtifact(presentation, { slide, format: "svg" });
assert.equal(presentationPreview.type, "image/svg+xml");
assert.equal(presentationPreview.metadata.artifactKind, "presentation");
assert.match(await presentationPreview.text(), /Render me/);

const document = DocumentModel.create({ paragraphs: ["Render document"] });
const documentPreview = await renderArtifact(document);
assert.equal(documentPreview.type, "image/svg+xml");
assert.equal(documentPreview.metadata.artifactKind, "document");
assert.match(await documentPreview.text(), /Render document/);
let adapterCalls = 0;
const rasterPreview = await renderArtifact(document, {
  format: "png",
  renderer: async ({ input, inputType, outputType, artifactKind }) => {
    adapterCalls += 1;
    assert.equal(inputType, "image/svg+xml");
    assert.equal(outputType, "image/png");
    assert.equal(artifactKind, "document");
    assert.match(await input.text(), /Render document/);
    return new FileBlob(new Uint8Array([0x89, 0x50, 0x4e, 0x47]), { type: "image/png", metadata: { adapter: "fake-png" } });
  },
});
assert.equal(adapterCalls, 1);
assert.equal(rasterPreview.type, "image/png");
assert.equal(rasterPreview.metadata.artifactKind, "document");
assert.equal(rasterPreview.metadata.format, "png");
assert.equal(rasterPreview.metadata.renderedFrom, "image/svg+xml");
assert.equal(rasterPreview.metadata.adapter, "fake-png");
const visualQa = await visualQaArtifact(document, { minBytes: 20 });
assert.equal(visualQa.ok, true);
assert.equal(visualQa.summary.type, "image/svg+xml");
assert.equal(visualQa.summary.width, 612);
assert.match(visualQa.ndjson, /"kind":"visualQa"/);
const changedVisualQa = await visualQaArtifact(document, { baseline: new FileBlob("different", { type: "image/svg+xml" }) });
assert.equal(changedVisualQa.ok, false);
assert.match(changedVisualQa.ndjson, /visualDiff/);
const pngArtifact = { render: () => new FileBlob(whitePixelPng, { type: "image/png" }) };
const pixelQa = await visualQaArtifact(pngArtifact, { baseline: new FileBlob(blackPixelPng, { type: "image/png" }), pixelDiff: true, maxChars: 4000 });
assert.equal(pixelQa.ok, false);
assert.equal(pixelQa.summary.pixelDiff.differentPixels, 1);
assert.equal(pixelQa.summary.pixelDiff.mismatchRatio, 1);
assert.match(pixelQa.ndjson, /visualPixelDiff/);
const unchangedPixelQa = await visualQaArtifact(pngArtifact, { baseline: new FileBlob(whitePixelPng, { type: "image/png" }), pixelDiff: true });
assert.equal(unchangedPixelQa.ok, true);
assert.equal(unchangedPixelQa.summary.pixelDiff.changed, false);
await assert.rejects(() => renderArtifact(document, { format: "webp" }), /no renderer adapter/);

const pdf = PdfArtifact.create({ pages: [{ text: "Render PDF", tables: [{ values: [["Metric", "Value"]] }] }] });
const pdfPreview = await renderArtifact(pdf, { pageIndex: 0 });
assert.equal(pdfPreview.type, "image/svg+xml");
assert.equal(pdfPreview.metadata.artifactKind, "pdf");
assert.equal(pdfPreview.metadata.pageIndex, 0);
assert.match(await pdfPreview.text(), /Render PDF/);

await assert.rejects(() => renderArtifact({}), /render\(\) or export\(\)/);

console.log("render smoke ok");
