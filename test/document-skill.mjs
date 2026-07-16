import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { DocumentFile, FileBlob } from "open-office-artifact-tool";
import {
  createDocumentFromFixture,
  nativeDocumentRenderStatus,
  runDocumentFixture,
  verifyDocumentFile,
} from "../skills/documents/scripts/workflow.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const fixturesDir = path.join(repoRoot, "skills", "documents", "fixtures");
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-document-skill-"));
const baselineDir = path.join(outputDir, "baselines");
const nativeStatus = nativeDocumentRenderStatus();

async function runFixture(name, options = {}) {
  const result = await runDocumentFixture(path.join(fixturesDir, `${name}.json`), {
    outputDir: path.join(outputDir, name),
    nativeRender: "off",
    ...options,
  });
  assert.deepEqual(Object.keys(result).sort(), ["docxPath", "fixture", "qa"]);
  assert.equal(result.fixture.name, name);
  assert.equal(result.qa.summary.packageOk, true);
  assert.equal(result.qa.summary.verifyOk, true);
  assert.equal(result.qa.summary.visualQaOk, true);
  return result;
}

try {
  assert.throws(
    () => createDocumentFromFixture({ settings: { trackRevisions: true }, blocks: [] }),
    /limited to evenAndOddHeaders.*trackRevisions.*read-only/i,
  );
  const business = await runFixture("business-brief", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(business.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) {
    assert.equal(business.qa.summary.nativeRender.ok, true);
    assert.ok(business.qa.summary.nativeRender.pageCount >= 1);
    assert.equal(business.qa.summary.nativeRender.pages.length, business.qa.summary.nativeRender.pageCount);
  }
  for (const filePath of Object.values(business.qa.summary.files)) {
    const stat = await fs.stat(filePath);
    assert.ok(stat.isFile() && stat.size > 0, `Expected non-empty document skill output ${filePath}`);
  }

  const document = await DocumentFile.importDocx(await FileBlob.load(business.docxPath));
  assert.equal(document.defaultRunStyle.fontFamily, "Aptos");
  assert.equal(document.styles.values().some((style) => style.id === "BriefLead" && style.basedOn === "Normal"), true);
  const editedLead = document.blocks.find((block) => block.text === "Create, inspect, render, and verify the canonical DOCX path.");
  assert.equal(editedLead?.kind, "paragraph");
  assert.equal(editedLead?.paragraphFormat.alignment, "left");
  assert.equal(editedLead?.runs[0].style.color, "#9c2b2e");
  assert.equal(document.blocks.some((block) => block.kind === "listItem" && block.text === "Inspect stable document blocks and fields."), true);
  assert.equal(document.blocks.find((block) => block.kind === "table")?.values[1][2], "Pass");
  const hyperlink = document.blocks.find((block) => block.kind === "hyperlink");
  assert.equal(hyperlink?.url, "https://learn.microsoft.com/office/open-xml/word-processing");
  assert.equal(hyperlink?.tooltip, "Edited through the canonical Office path");
  assert.equal(hyperlink?.history, false);
  const image = document.blocks.find((block) => block.kind === "image");
  assert.equal(image?.alt, "Edited green status mark");
  assert.equal(image?.widthPx, 48);
  assert.equal(document.blocks.find((block) => block.kind === "field")?.instruction, "NUMPAGES");
  assert.equal(document.blocks.find((block) => block.kind === "section")?.margins.left, 1200);
  assert.equal(document.comments[0]?.author, "Lead reviewer");
  assert.equal(document.comments[0]?.text, "Delivery evidence approved.");
  assert.equal(document.settings.evenAndOddHeaders, true);
  assert.equal(document.headers.some((item) => item.referenceType === "first" && item.variantActive), true);
  assert.equal(document.footers.some((item) => item.referenceType === "even" && item.fieldInstruction === "PAGE"), true);

  const businessZip = await JSZip.loadAsync(await fs.readFile(business.docxPath));
  for (const part of [
    "word/document.xml",
    "word/styles.xml",
    "word/numbering.xml",
    "word/comments.xml",
    "word/settings.xml",
  ]) assert.ok(businessZip.file(part), `Expected ${part}`);
  assert.ok(Object.keys(businessZip.files).some((name) => /(^|\/)media\//.test(name)));
  assert.ok(Object.keys(businessZip.files).filter((name) => /^word\/header\d+\.xml$/.test(name)).length >= 2);
  assert.ok(Object.keys(businessZip.files).some((name) => /^word\/footer\d+\.xml$/.test(name)));
  const businessXml = await businessZip.file("word/document.xml").async("text");
  assert.match(businessXml, /w:instr="NUMPAGES"/);
  assert.match(businessXml, /<w:drawing>/);
  assert.match(businessXml, /<w:sectPr>/);
  assert.match(await fs.readFile(business.qa.summary.files.packageInspect, "utf8"), /word\/document\.xml/);

  const merged = await runFixture("open-chestnut-merged-table");
  const mergedDocument = await DocumentFile.importDocx(await FileBlob.load(merged.docxPath));
  const mergedTable = mergedDocument.blocks.find((block) => block.kind === "table");
  assert.equal(mergedTable?.values[0][0], "Edited merged owner");
  assert.equal(mergedTable?.getCell(0, 0).columnSpan, 2);
  assert.equal(mergedTable?.getCell(0, 0).rowSpan, 2);
  assert.equal(mergedTable?.getCell(1, 0).verticalMerge, "continue");
  assert.equal(mergedTable?.getCell(1, 0).editable, false);
  assert.equal(mergedTable?.widthDxa, 9300);
  assert.deepEqual(mergedTable?.columnWidthsDxa, [2500, 3100, 3700]);
  assert.equal(mergedTable?.borderColor, "884400");

  const numbering = await runFixture("open-chestnut-numbering-edit");
  const numberingDocument = await DocumentFile.importDocx(await FileBlob.load(numbering.docxPath));
  const numberedItems = numberingDocument.blocks.filter((block) => block.kind === "listItem");
  assert.equal(numberedItems.length, 2);
  assert.equal(numberedItems[0].text, "Edited first grouped item");
  assert.equal(numberedItems.every((block) => block.numberFormat === "lowerRoman" && block.start === 5 && block.levelText === "%1."), true);

  const comments = await runFixture("open-chestnut-comments");
  const commentsDocument = await DocumentFile.importDocx(await FileBlob.load(comments.docxPath));
  assert.equal(commentsDocument.comments.length, 1);
  assert.equal(commentsDocument.comments[0].author, "Lead reviewer");
  assert.equal(commentsDocument.comments[0].initials, "LR");
  assert.equal(commentsDocument.comments[0].text, "Approved after source-bound review.");

  const classicFixture = await runFixture("package-comments");
  const classicDocument = await DocumentFile.importDocx(await FileBlob.load(classicFixture.docxPath));
  assert.equal(classicDocument.comments.length, 1);
  assert.equal(classicDocument.comments[0].author, "QA Lead");
  assert.equal(classicDocument.comments[0].text, "Decision paragraph confirmed.");

  const directNumbering = await runFixture("package-numbering");
  const directNumberingDocument = await DocumentFile.importDocx(await FileBlob.load(directNumbering.docxPath));
  assert.equal(directNumberingDocument.blocks.filter((block) => block.kind === "listItem").length, 2);
  assert.equal(directNumberingDocument.blocks.some((block) => block.text === "Confirm the edited second item."), true);

  const sectionSettings = await runFixture("package-settings");
  const settingsDocument = await DocumentFile.importDocx(await FileBlob.load(sectionSettings.docxPath));
  assert.equal(settingsDocument.settings.evenAndOddHeaders, true);
  assert.equal(settingsDocument.sectionSettings[0]?.differentFirstPage, true);
  assert.equal(settingsDocument.headers.some((item) => item.referenceType === "first"), true);
  assert.equal(settingsDocument.headers.some((item) => item.referenceType === "even"), true);
  assert.equal(settingsDocument.footers[0]?.fieldInstruction, "PAGE");

  const baselineWrite = await verifyDocumentFile(business.docxPath, {
    outputDir: path.join(outputDir, "baseline-write"),
    previewFormat: "png",
    nativeRender: "off",
    baselineDir,
    writeBaseline: true,
  });
  assert.equal(baselineWrite.summary.writeBaseline, true);
  assert.ok((await fs.stat(baselineWrite.summary.modelBaselinePath)).size > 100);
  const baselineCompare = await verifyDocumentFile(business.docxPath, {
    outputDir: path.join(outputDir, "baseline-compare"),
    previewFormat: "png",
    nativeRender: "off",
    baselineDir,
  });
  assert.equal(baselineCompare.summary.modelBaselineCompared, true);
  assert.equal(baselineCompare.summary.modelPixelDiff.changed, false);
  assert.equal(baselineCompare.summary.visualQaOk, true);

  const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
  assert.ok(packageJson.files.includes("skills/**"));
  const skillText = await fs.readFile(path.join(repoRoot, "skills", "documents", "SKILL.md"), "utf8");
  assert.match(skillText, /canonical OpenChestnut Office path/);
  assert.match(skillText, /LibreOffice PDF plus Poppler page PNGs/);
  assert.match(skillText, /baseline-dir/);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("document skill smoke ok");
