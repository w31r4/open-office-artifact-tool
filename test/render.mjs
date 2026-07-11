import assert from "node:assert/strict";

import {
  DocumentModel,
  FileBlob,
  PdfArtifact,
  Presentation,
  renderArtifact,
  visualQaArtifact,
  Workbook,
} from "open-office-artifact-tool";

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
await assert.rejects(() => renderArtifact(document, { format: "webp" }), /no renderer adapter/);

const pdf = PdfArtifact.create({ pages: [{ text: "Render PDF", tables: [{ values: [["Metric", "Value"]] }] }] });
const pdfPreview = await renderArtifact(pdf, { pageIndex: 0 });
assert.equal(pdfPreview.type, "image/svg+xml");
assert.equal(pdfPreview.metadata.artifactKind, "pdf");
assert.equal(pdfPreview.metadata.pageIndex, 0);
assert.match(await pdfPreview.text(), /Render PDF/);

await assert.rejects(() => renderArtifact({}), /render\(\) or export\(\)/);

console.log("render smoke ok");
