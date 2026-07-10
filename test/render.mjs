import assert from "node:assert/strict";

import {
  DocumentModel,
  PdfArtifact,
  Presentation,
  renderArtifact,
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

const pdf = PdfArtifact.create({ pages: [{ text: "Render PDF", tables: [{ values: [["Metric", "Value"]] }] }] });
const pdfPreview = await renderArtifact(pdf, { pageIndex: 0 });
assert.equal(pdfPreview.type, "image/svg+xml");
assert.equal(pdfPreview.metadata.artifactKind, "pdf");
assert.equal(pdfPreview.metadata.pageIndex, 0);
assert.match(await pdfPreview.text(), /Render PDF/);

await assert.rejects(() => renderArtifact({}), /render\(\) or export\(\)/);

console.log("render smoke ok");
