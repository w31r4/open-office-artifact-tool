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
assert.ok(HELP_CATALOG.some((item) => item.name === "slide.addNotes"));
assert.ok(HELP_CATALOG.some((item) => item.name === "slide.comments.addThread"));
assert.ok(HELP_CATALOG.some((item) => item.name === "slide.connectors.add"));
assert.ok(HELP_CATALOG.some((item) => item.name === "presentation.theme"));
assert.ok(HELP_CATALOG.some((item) => item.name === "presentation.layouts.add"));
assert.ok(HELP_CATALOG.some((item) => item.name === "slide.applyLayout"));
assert.ok(HELP_CATALOG.some((item) => item.name === "document.addHyperlink"));
assert.ok(HELP_CATALOG.some((item) => item.name === "document.applyDesignPreset"));
assert.ok(HELP_CATALOG.some((item) => item.name === "document.layoutJson"));
assert.ok(HELP_CATALOG.some((item) => item.name === "pdf.extractTables"));
assert.ok(HELP_CATALOG.some((item) => item.name === "createPdfjsParser"));
assert.ok(HELP_CATALOG.some((item) => item.name === "verifyArtifact"));
assert.ok(HELP_CATALOG.some((item) => item.name === "renderArtifact"));
assert.ok(HELP_CATALOG.some((item) => item.name === "range.format"));
assert.ok(HELP_CATALOG.some((item) => item.name === "workbook.formulaGraph"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.AVERAGE"));
assert.ok(HELP_CATALOG.some((item) => item.name === "createPlaywrightRenderer"));
assert.ok(HELP_CATALOG.some((item) => item.name === "createSharpRenderer"));
assert.ok(HELP_CATALOG.some((item) => item.name === "createPopplerRenderer"));
assert.ok(HELP_CATALOG.some((item) => item.name === "createNativeOfficeRenderer"));
assert.ok(HELP_CATALOG.some((item) => item.name === "renderFileWithNativeOffice"));

const workbook = Workbook.create();
const presentation = Presentation.create();
const document = DocumentModel.create({ paragraphs: ["Doc"] });
const pdf = PdfArtifact.create({ text: "PDF" });

assert.match(workbook.help("sheet.charts.add").ndjson, /worksheet chart/);
assert.match(workbook.help("workbook.formulaGraph").ndjson, /dependency graph/);
assert.match(workbook.help("fx.AVERAGE").ndjson, /Average numeric values/);
assert.match(presentation.help("slide.compose").ndjson, /compose tree/);
assert.match(presentation.help("slide.addNotes").ndjson, /speaker notes/);
assert.match(presentation.help("slide.connectors.add").ndjson, /connector line/);
assert.match(presentation.help("presentation.theme").ndjson, /theme colors/);
assert.match(presentation.help("presentation.layouts.add").ndjson, /slide layout/);
assert.match(document.help("document.addField").ndjson, /fldSimple/);
assert.match(document.help("document.applyDesignPreset").ndjson, /design preset/);
assert.match(document.help("document.layoutJson").ndjson, /layout JSON/);
assert.match(pdf.help("extractTables").ndjson, /table values/);
assert.match(pdf.help("createPdfjsParser").ndjson, /positioned text/);
assert.match(helpArtifact("*", "renderArtifact").ndjson, /FileBlob metadata/);
assert.match(helpArtifact("shared", "createPlaywrightRenderer").ndjson, /Playwright renderer adapter/);
assert.match(helpArtifact("shared", "createSharpRenderer").ndjson, /sharp renderer adapter/);
assert.match(helpArtifact("shared", "createPopplerRenderer").ndjson, /Poppler CLI renderer adapter/);
assert.match(helpArtifact("shared", "createNativeOfficeRenderer").ndjson, /native Office renderer adapter/);
assert.match(helpArtifact("shared", "renderFileWithNativeOffice").ndjson, /native Office bridge command/);
assert.match(helpArtifact(workbook, "fx.PMT").ndjson, /financial/);
assert.equal(helpArtifact("presentation", "sheet.charts.add").ndjson, "");

console.log("help smoke ok");
