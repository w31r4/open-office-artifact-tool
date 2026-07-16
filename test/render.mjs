import assert from "node:assert/strict";
import { deflateSync } from "node:zlib";
import sharp from "sharp";

import {
  DocumentModel,
  FileBlob,
  PdfArtifact,
  Presentation,
  renderArtifact,
  visualQaArtifact,
  Workbook,
} from "open-office-artifact-tool";
import { createArtifactVisualQaApi } from "../src/qa/artifact-visual.mjs";

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

function makePpm(width, height, rgb) {
  return new Uint8Array(Buffer.concat([Buffer.from(`P6\n${width} ${height}\n255\n`, "ascii"), Buffer.from(rgb)]));
}

const blackPixelPng = makePng(1, 1, [0, 0, 0, 255]);
const whitePixelPng = makePng(1, 1, [255, 255, 255, 255]);
const blackPixelPpm = makePpm(1, 1, [0, 0, 0]);
const whitePixelPpm = makePpm(1, 1, [255, 255, 255]);

assert.throws(() => createArtifactVisualQaApi({}), /requires inferArtifactKind/);
let inferredArtifactKinds = 0;
const injectedVisualApi = createArtifactVisualQaApi({
  inferArtifactKind: () => {
    inferredArtifactKinds += 1;
    return "fixture";
  },
});
const injectedPreview = await injectedVisualApi.renderArtifact({ render: () => new FileBlob(whitePixelPng, { type: "image/png" }) });
assert.equal(inferredArtifactKinds, 1);
assert.equal(injectedPreview.metadata.artifactKind, "fixture");

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
assert.equal(pixelQa.diffBlob.type, "image/png");
assert.ok(pixelQa.summary.diff.bytes > 0);
assert.deepEqual(await sharp(pixelQa.diffBlob.bytes).metadata().then(({ width, height, format }) => ({ width, height, format })), { width: 1, height: 1, format: "png" });
assert.match(pixelQa.ndjson, /visualPixelDiff/);
const paletteQa = await visualQaArtifact(pngArtifact, { baseline: new FileBlob(blackPixelPng, { type: "image/png" }), pixelDiff: { diffPalette: { changed: "#00aaff", changedAlpha: 128 } } });
const palettePixels = await sharp(paletteQa.diffBlob.bytes).raw().toBuffer({ resolveWithObject: true });
assert.deepEqual([...palettePixels.data], [0, 170, 255, 128]);
assert.deepEqual(paletteQa.summary.pixelDiff.diffPalette.changed, [0, 170, 255]);
assert.deepEqual(paletteQa.diffBlob.metadata.palette.changed, [0, 170, 255]);
const wideBlackPng = new Uint8Array(await sharp({ create: { width: 2, height: 1, channels: 4, background: { r: 0, g: 0, b: 0, alpha: 1 } } }).png().toBuffer());
const alignedQa = await visualQaArtifact(pngArtifact, { baseline: new FileBlob(wideBlackPng, { type: "image/png" }), pixelDiff: { diffAlignment: "center", diffPalette: { changed: [255, 128, 0], unchanged: "#112233" }, pixelRegistration: 2 } });
assert.equal(alignedQa.summary.pixelDiff.dimensionMismatch, true);
assert.equal(alignedQa.summary.pixelDiff.alignment, "center");
assert.equal(alignedQa.summary.pixelDiff.diffWidth, 2);
assert.equal(alignedQa.summary.pixelDiff.diffHeight, 1);
assert.equal(alignedQa.diffBlob.metadata.alignment, "center");
assert.equal(alignedQa.summary.pixelDiff.registration.skipped, "dimensionMismatch");
assert.deepEqual(await sharp(alignedQa.diffBlob.bytes).metadata().then(({ width, height }) => ({ width, height })), { width: 2, height: 1 });
const strictDimensionQa = await visualQaArtifact(pngArtifact, { baseline: new FileBlob(wideBlackPng, { type: "image/png" }), pixelDiff: true });
assert.equal(strictDimensionQa.summary.pixelDiff.alignment, "strict");
assert.equal(strictDimensionQa.diffBlob, undefined);
const markerPng = async (markerX, extraX) => {
  const pixels = Buffer.alloc(5 * 5 * 4, 255);
  const marker = (2 * 5 + markerX) * 4;
  pixels[marker] = 0; pixels[marker + 1] = 0; pixels[marker + 2] = 0;
  if (extraX != null) { const extra = (2 * 5 + extraX) * 4; pixels[extra] = 255; pixels[extra + 1] = 0; pixels[extra + 2] = 0; }
  return new Uint8Array(await sharp(pixels, { raw: { width: 5, height: 5, channels: 4 } }).png().toBuffer());
};
const baselineMarkerPng = await markerPng(1);
const shiftedMarkerPng = await markerPng(2);
const shiftedMarkerArtifact = { render: () => new FileBlob(shiftedMarkerPng, { type: "image/png" }) };
const unregisteredMarkerQa = await visualQaArtifact(shiftedMarkerArtifact, { baseline: new FileBlob(baselineMarkerPng, { type: "image/png" }), pixelDiff: true });
assert.equal(unregisteredMarkerQa.summary.pixelDiff.differentPixels, 2);
const registeredMarkerQa = await visualQaArtifact(shiftedMarkerArtifact, { baseline: new FileBlob(baselineMarkerPng, { type: "image/png" }), pixelDiff: { pixelRegistration: { maxOffset: 2, maxSamples: 1_000 } } });
assert.equal(registeredMarkerQa.ok, true);
assert.equal(registeredMarkerQa.summary.byteChanged, true);
assert.equal(registeredMarkerQa.summary.changed, false);
assert.equal(registeredMarkerQa.summary.pixelDiff.registration.applied, true);
assert.deepEqual(registeredMarkerQa.summary.pixelDiff.registration.offset, { x: 1, y: 0 });
assert.equal(registeredMarkerQa.summary.pixelDiff.registration.differentPixelsBefore, 2);
assert.equal(registeredMarkerQa.summary.pixelDiff.registration.differentPixelsAfter, 0);
assert.equal(registeredMarkerQa.summary.pixelDiff.registration.ignoredEdgePixels, 5);
const shiftedChangedMarkerPng = await markerPng(2, 3);
const registeredChangedQa = await visualQaArtifact({ render: () => new FileBlob(shiftedChangedMarkerPng, { type: "image/png" }) }, { baseline: new FileBlob(baselineMarkerPng, { type: "image/png" }), pixelDiff: { pixelRegistration: 2 } });
assert.equal(registeredChangedQa.summary.pixelDiff.registration.applied, true);
assert.equal(registeredChangedQa.summary.pixelDiff.differentPixels, 1);
assert.deepEqual(registeredChangedQa.diffBlob.metadata.registration.offset, { x: 1, y: 0 });
const rejectedRegistrationQa = await visualQaArtifact({ render: () => new FileBlob(shiftedChangedMarkerPng, { type: "image/png" }) }, { baseline: new FileBlob(baselineMarkerPng, { type: "image/png" }), pixelDiff: { pixelRegistration: { maxOffset: 2, minImprovementRatio: 0.9 } } });
assert.equal(rejectedRegistrationQa.summary.pixelDiff.registration.applied, false);
assert.deepEqual(rejectedRegistrationQa.summary.pixelDiff.registration.offset, { x: 0, y: 0 });
assert.equal(rejectedRegistrationQa.summary.pixelDiff.differentPixels, 3);
const noDiffImageQa = await visualQaArtifact(pngArtifact, { baseline: new FileBlob(blackPixelPng, { type: "image/png" }), pixelDiff: true, diffImage: false });
assert.equal(noDiffImageQa.diffBlob, undefined);
const unchangedPixelQa = await visualQaArtifact(pngArtifact, { baseline: new FileBlob(whitePixelPng, { type: "image/png" }), pixelDiff: true });
assert.equal(unchangedPixelQa.ok, true);
assert.equal(unchangedPixelQa.summary.pixelDiff.changed, false);
assert.equal(unchangedPixelQa.diffBlob, undefined);
const ppmArtifact = { render: () => new FileBlob(whitePixelPpm, { type: "image/x-portable-pixmap" }) };
const ppmQa = await visualQaArtifact(ppmArtifact, { baseline: new FileBlob(blackPixelPpm, { type: "image/x-portable-pixmap" }), pixelDiff: true, maxChars: 4000 });
assert.equal(ppmQa.ok, false);
assert.equal(ppmQa.summary.pixelDiff.format, "ppm");
assert.equal(ppmQa.summary.pixelDiff.differentPixels, 1);
assert.match(ppmQa.ndjson, /visualPixelDiff/);
const whiteRaw = { create: { width: 1, height: 1, channels: 3, background: { r: 255, g: 255, b: 255 } } };
const blackRaw = { create: { width: 1, height: 1, channels: 3, background: { r: 0, g: 0, b: 0 } } };
const whiteJpeg = new Uint8Array(await sharp(whiteRaw).jpeg({ quality: 100, chromaSubsampling: "4:4:4" }).toBuffer());
const whiteJpegWithMetadata = new Uint8Array(await sharp(whiteRaw).withMetadata({ orientation: 1 }).jpeg({ quality: 100, chromaSubsampling: "4:4:4" }).toBuffer());
const whiteWebp = new Uint8Array(await sharp(whiteRaw).webp({ lossless: true }).toBuffer());
const blackWebp = new Uint8Array(await sharp(blackRaw).webp({ lossless: true }).toBuffer());
assert.notDeepEqual(whiteJpeg, whiteJpegWithMetadata);
const jpegArtifact = { render: () => new FileBlob(whiteJpeg, { type: "image/jpeg" }) };
const jpegMetadataQa = await visualQaArtifact(jpegArtifact, { baseline: new FileBlob(whiteJpegWithMetadata, { type: "image/jpeg" }), pixelDiff: true });
assert.equal(jpegMetadataQa.ok, true);
assert.equal(jpegMetadataQa.summary.byteChanged, true);
assert.equal(jpegMetadataQa.summary.changed, false);
assert.equal(jpegMetadataQa.summary.pixelDiff.format, "jpeg");
const mixedRasterQa = await visualQaArtifact(jpegArtifact, { baseline: new FileBlob(whiteWebp, { type: "image/webp" }), pixelDiff: true });
assert.equal(mixedRasterQa.ok, true);
assert.equal(mixedRasterQa.summary.pixelDiff.format, "jpeg/webp");
assert.equal(mixedRasterQa.summary.pixelDiff.changed, false);
const webpArtifact = { render: () => new FileBlob(whiteWebp, { type: "image/webp" }) };
const changedWebpQa = await visualQaArtifact(webpArtifact, { baseline: new FileBlob(blackWebp, { type: "image/webp" }), pixelDiff: true });
assert.equal(changedWebpQa.ok, false);
assert.equal(changedWebpQa.summary.pixelDiff.format, "webp");
assert.equal(changedWebpQa.summary.pixelDiff.differentPixels, 1);
assert.equal(changedWebpQa.diffBlob.type, "image/png");
assert.match(changedWebpQa.ndjson, /visualPixelDiff/);
await assert.rejects(() => renderArtifact(document, { format: "webp" }), /no renderer adapter/);

const pdf = PdfArtifact.create({ pages: [{ text: "Render PDF", tables: [{ values: [["Metric", "Value"]] }] }] });
const pdfPreview = await renderArtifact(pdf, { pageIndex: 0 });
assert.equal(pdfPreview.type, "image/svg+xml");
assert.equal(pdfPreview.metadata.artifactKind, "pdf");
assert.equal(pdfPreview.metadata.pageIndex, 0);
assert.match(await pdfPreview.text(), /Render PDF/);

await assert.rejects(() => renderArtifact({}), /render\(\) or export\(\)/);

console.log("render smoke ok");
