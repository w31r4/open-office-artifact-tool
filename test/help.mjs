import assert from "node:assert/strict";
import fs from "node:fs/promises";

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
assert.ok(HELP_CATALOG.some((item) => item.name === "visualQaArtifact"));
assert.ok(HELP_CATALOG.some((item) => item.name === "renderArtifact"));
assert.ok(HELP_CATALOG.some((item) => item.name === "range.format"));
assert.ok(HELP_CATALOG.some((item) => item.name === "workbook.formulaGraph"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.AVERAGE"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.XLOOKUP"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.TEXTJOIN"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.SUMPRODUCT"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.HLOOKUP"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.IFERROR"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.ISNUMBER"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.ISTEXT"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.ISBLANK"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.ISERROR"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.AVERAGEIF"));
assert.ok(HELP_CATALOG.some((item) => item.name === "fx.AVERAGEIFS"));
assert.ok(HELP_CATALOG.some((item) => item.name === "createPlaywrightRenderer"));
assert.ok(HELP_CATALOG.some((item) => item.name === "createSharpRenderer"));
assert.ok(HELP_CATALOG.some((item) => item.name === "createPopplerRenderer"));
assert.ok(HELP_CATALOG.some((item) => item.name === "createLibreOfficeRenderer"));
assert.ok(HELP_CATALOG.some((item) => item.name === "createNativeOfficeRenderer"));
assert.ok(HELP_CATALOG.some((item) => item.name === "renderFileWithNativeOffice"));
assert.ok(HELP_CATALOG.find((item) => item.name === "workbook.inspect")?.options?.includes("include/fields"));
assert.ok(HELP_CATALOG.find((item) => item.name === "workbook.inspect")?.schema?.parameters?.kind);
assert.ok(HELP_CATALOG.find((item) => item.name === "renderArtifact")?.schema?.returns?.blob);
assert.ok(HELP_CATALOG.find((item) => item.name === "range.conditionalFormats.add")?.schema?.parameters?.colors);
assert.ok(HELP_CATALOG.find((item) => item.name === "DocumentFile.patchDocx")?.schema?.parameters?.patches);
assert.ok(HELP_CATALOG.find((item) => item.name === "PdfFile.importPdf")?.schema?.returns?.pdf);
assert.ok(HELP_CATALOG.find((item) => item.name === "renderArtifact")?.returns?.includes("FileBlob"));
assert.ok(HELP_CATALOG.find((item) => item.name === "visualQaArtifact")?.examples?.some((example) => example.includes("pixelDiff")));

const workbook = Workbook.create();
const presentation = Presentation.create();
const document = DocumentModel.create({ paragraphs: ["Doc"] });
const pdf = PdfArtifact.create({ text: "PDF" });

assert.match(workbook.help("sheet.charts.add").ndjson, /worksheet chart/);
assert.match(workbook.help("workbook.formulaGraph").ndjson, /dependency graph/);
assert.match(workbook.help("fx.AVERAGE").ndjson, /Average numeric values/);
assert.match(workbook.help("fx.XLOOKUP").ndjson, /lookup/);
assert.match(workbook.help("fx.TEXTJOIN").ndjson, /delimiter/);
assert.match(workbook.help("fx.SUMPRODUCT").ndjson, /products/);
assert.match(workbook.help("fx.HLOOKUP").ndjson, /first row/);
assert.match(workbook.help("fx.IFERROR").ndjson, /formula error/);
assert.match(workbook.help("fx.AVERAGEIFS").ndjson, /criteria ranges/);
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
assert.match(helpArtifact("shared", "visualQaArtifact").ndjson, /baseline render/);
assert.match(helpArtifact("shared", "createPlaywrightRenderer").ndjson, /Playwright renderer adapter/);
assert.match(helpArtifact("shared", "createSharpRenderer").ndjson, /sharp renderer adapter/);
assert.match(helpArtifact("shared", "createPopplerRenderer").ndjson, /Poppler CLI renderer adapter/);
assert.match(helpArtifact("shared", "createLibreOfficeRenderer").ndjson, /LibreOffice CLI renderer adapter/);
assert.match(helpArtifact("shared", "createNativeOfficeRenderer").ndjson, /native Office renderer adapter/);
assert.match(helpArtifact("shared", "renderFileWithNativeOffice").ndjson, /native Office bridge command/);
assert.match(helpArtifact("shared", "pixelDiff").ndjson, /visualQaArtifact/);
assert.match(helpArtifact("workbook", "include\/fields").ndjson, /workbook.inspect/);
assert.match(helpArtifact("workbook", "colorScale").ndjson, /range\.conditionalFormats\.add/);
assert.match(helpArtifact("document", "maxPatchBytes").ndjson, /DocumentFile\.patchDocx/);
assert.match(helpArtifact(workbook, "fx.PMT").ndjson, /financial/);
assert.equal(helpArtifact("presentation", "sheet.charts.add").ndjson, "");
const apiDocs = await fs.readFile(new URL("../docs/api.md", import.meta.url), "utf8");
assert.match(apiDocs, /### shared details/);
assert.match(apiDocs, /#### `renderArtifact`/);
assert.match(apiDocs, /await renderArtifact\(document/);
assert.match(apiDocs, /#### `workbook.inspect`/);
assert.match(apiDocs, /include\/fields/);
assert.match(apiDocs, /\*\*Schema parameters:\*\*/);
assert.match(apiDocs, /`kind` \(string\)/);
assert.match(apiDocs, /`blob` \(FileBlob\)/);
assert.match(apiDocs, /DocumentFile\.patchDocx/);

console.log("help smoke ok");
