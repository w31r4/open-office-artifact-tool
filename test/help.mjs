import assert from "node:assert/strict";

import {
  DocumentModel,
  helpArtifact,
  HELP_CATALOG,
  PdfArtifact,
  Presentation,
  Workbook,
} from "open-office-artifact-tool";

assert.ok(HELP_CATALOG.length >= 40);
assert.ok(HELP_CATALOG.some((item) => item.name === "Workbook.create"));
assert.ok(HELP_CATALOG.some((item) => item.name === "slide.compose"));
assert.ok(HELP_CATALOG.some((item) => item.name === "document.addHyperlink"));
assert.ok(HELP_CATALOG.some((item) => item.name === "pdf.extractTables"));
assert.ok(HELP_CATALOG.some((item) => item.name === "verifyArtifact"));
assert.ok(HELP_CATALOG.some((item) => item.name === "renderArtifact"));
assert.ok(HELP_CATALOG.some((item) => item.name === "workbook.formulaGraph"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.AVERAGE"));
assert.ok(HELP_CATALOG.some((item) => item.name === "createPlaywrightRenderer"));

const workbook = Workbook.create();
const presentation = Presentation.create();
const document = DocumentModel.create({ paragraphs: ["Doc"] });
const pdf = PdfArtifact.create({ text: "PDF" });

assert.match(workbook.help("sheet.charts.add").ndjson, /worksheet chart/);
assert.match(workbook.help("workbook.formulaGraph").ndjson, /dependency graph/);
assert.match(workbook.help("fx.AVERAGE").ndjson, /Average numeric values/);
assert.match(presentation.help("slide.compose").ndjson, /compose tree/);
assert.match(document.help("document.addField").ndjson, /fldSimple/);
assert.match(pdf.help("extractTables").ndjson, /table values/);
assert.match(helpArtifact("*", "renderArtifact").ndjson, /FileBlob metadata/);
assert.match(helpArtifact("shared", "createPlaywrightRenderer").ndjson, /Playwright renderer adapter/);
assert.match(helpArtifact(workbook, "fx.PMT").ndjson, /financial/);
assert.equal(helpArtifact("presentation", "sheet.charts.add").ndjson, "");

console.log("help smoke ok");
