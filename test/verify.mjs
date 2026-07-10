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
document.addParagraph("Styled with missing style", { styleId: "MissingStyle" });
document.addHyperlink("bad", "ftp://example.com");
document.addCitation("Bad cite", { url: "mailto:source" });
const badImage = document.addImage({ name: "bad-image", dataUrl: "data:not-base64" });
badImage.widthPx = 0;
badImage.heightPx = -1;
const badSection = document.addSection({ name: "bad-section" });
badSection.orientation = "sideways";
badSection.breakType = "somewhere";
badSection.pageSize.widthTwips = 0;
badSection.margins.left = -1;
document.addTable({ name: "empty-table", values: [] });
document.addTable({ name: "ragged-table", values: [["A", "B"], ["only one cell"]] });
const documentIssues = document.verify({ maxChars: 16000 }).ndjson;
assert.match(documentIssues, /fakeList/);
assert.match(documentIssues, /unknownStyle/);
assert.match(documentIssues, /invalidHyperlink/);
assert.match(documentIssues, /invalidCitationUrl/);
assert.match(documentIssues, /invalidImageDataUrl/);
assert.match(documentIssues, /invalidImageDimensions/);
assert.match(documentIssues, /invalidSectionOrientation/);
assert.match(documentIssues, /invalidSectionBreak/);
assert.match(documentIssues, /invalidSectionPageSize/);
assert.match(documentIssues, /invalidSectionMargin/);
assert.match(documentIssues, /emptyTable/);
assert.match(documentIssues, /raggedTableRows/);

const pdf = PdfArtifact.create({ pages: [{ text: "This uses an en dash – bad", tables: [{ values: [[]], bbox: [0, 0, 0, 10] }] }] });
const pdfIssues = verifyArtifact(pdf, { maxChars: 8000 }).ndjson;
assert.match(pdfIssues, /unicodeDash/);
assert.match(pdfIssues, /emptyTable|tableOutOfBounds/);

const unsupported = verifyArtifact({});
assert.equal(unsupported.ok, false);
assert.match(unsupported.ndjson, /unsupportedArtifact/);

console.log("verify smoke ok");
