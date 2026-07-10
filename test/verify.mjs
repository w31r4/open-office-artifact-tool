import assert from "node:assert/strict";

import {
  DocumentModel,
  PdfArtifact,
  Presentation,
  verifyArtifact,
  Workbook,
} from "open-office-artifact-tool";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet1");
sheet.getRange("A1:B2").values = [["A", "B"], [1, 2]];
sheet.getRange("C2").formulas = [["=A2+B2"]];
assert.equal(workbook.verify().ok, true);
sheet.getRange("D2").formulas = [["=UNKNOWN(A2)"]];
const workbookIssues = verifyArtifact(workbook).ndjson;
assert.match(workbookIssues, /formulaError/);
assert.match(workbookIssues, /#NAME\?/);

const presentation = Presentation.create({ slideSize: { width: 320, height: 180 } });
const slide = presentation.slides.add();
slide.shapes.add({ name: "ok", position: { left: 20, top: 20, width: 80, height: 60 }, text: "OK" });
assert.equal(presentation.verify().ok, true);
slide.shapes.add({ name: "bad-overlap", position: { left: 40, top: 30, width: 80, height: 40 }, text: "Bad" });
assert.match(verifyArtifact(presentation).ndjson, /overlap/);

const document = DocumentModel.create({ paragraphs: ["Title"] });
document.addParagraph("- fake bullet");
document.addHyperlink("bad", "ftp://example.com");
const documentIssues = document.verify({ maxChars: 8000 }).ndjson;
assert.match(documentIssues, /fakeList/);
assert.match(documentIssues, /invalidHyperlink/);

const pdf = PdfArtifact.create({ pages: [{ text: "This uses an en dash – bad", tables: [{ values: [[]], bbox: [0, 0, 0, 10] }] }] });
const pdfIssues = verifyArtifact(pdf, { maxChars: 8000 }).ndjson;
assert.match(pdfIssues, /unicodeDash/);
assert.match(pdfIssues, /emptyTable|tableOutOfBounds/);

const unsupported = verifyArtifact({});
assert.equal(unsupported.ok, false);
assert.match(unsupported.ndjson, /unsupportedArtifact/);

console.log("verify smoke ok");
