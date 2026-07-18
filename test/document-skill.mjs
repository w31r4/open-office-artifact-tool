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
} from "./skill-harness/documents/scripts/workflow.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const fixturesDir = path.join(repoRoot, "test", "skill-harness", "documents", "fixtures");
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
    /limited to evenAndOddHeaders and updateFields.*trackRevisions.*read-only/i,
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

  const controls = await runFixture("open-chestnut-content-controls");
  const controlsDocument = await DocumentFile.importDocx(await FileBlob.load(controls.docxPath));
  assert.deepEqual(controlsDocument.contentControls.map((control) => [control.tag, control.text]), [
    ["CUSTOMER_NAME", "Grace Hopper"],
    ["ACCOUNT_ID", "AC-2048"],
  ]);
  assert.equal(controlsDocument.inspect({ kind: "contentControl" }).ndjson.includes("Customer name"), true);
  const controlsZip = await JSZip.loadAsync(await fs.readFile(controls.docxPath));
  const controlsXml = await controlsZip.file("word/document.xml").async("text");
  assert.equal((controlsXml.match(/<w:sdt>/g) || []).length, 2);
  assert.match(controlsXml, /<w:tag w:val="CUSTOMER_NAME"\s*\/>/);
  assert.match(controlsXml, /<w:tag w:val="ACCOUNT_ID"\s*\/>/);

  const bibliography = await runFixture("open-chestnut-bibliography");
  const bibliographyDocument = await DocumentFile.importDocx(await FileBlob.load(bibliography.docxPath));
  assert.equal(bibliographyDocument.bibliography.styleName, "APA");
  assert.deepEqual(bibliographyDocument.bibliographySources.map((source) => [source.tag, source.title, source.authors[0].first]), [
    ["AgentSource", "Notes on the Analytical Engine", "Augusta Ada"],
  ]);
  assert.equal(bibliographyDocument.blocks.find((block) => block.kind === "citation")?.text, "(Lovelace, 1843, revised)");
  const bibliographyZip = await JSZip.loadAsync(await fs.readFile(bibliography.docxPath));
  const bibliographyParts = Object.keys(bibliographyZip.files).filter((name) => /^customXml\/item\d*\.xml$/.test(name));
  assert.equal(bibliographyParts.length, 1);
  assert.match(await bibliographyZip.file(bibliographyParts[0]).async("text"), /<Sources\b[^>]*xmlns="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/bibliography"/);
  assert.match(await bibliographyZip.file("word/document.xml").async("text"), /w:instr=" CITATION AgentSource "/);

  const toc = await runFixture("open-chestnut-toc", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  const tocDocument = await DocumentFile.importDocx(await FileBlob.load(toc.docxPath));
  const tocField = tocDocument.blocks.find((block) => block.kind === "field");
  assert.equal(tocDocument.settings.updateFields, true);
  assert.equal(tocField?.complex, true);
  assert.equal(tocField?.instruction, 'TOC \\o "1-4" \\h \\z \\u');
  assert.equal(tocField?.display, "Update this table of contents in Word");
  assert.equal(toc.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  const tocZip = await JSZip.loadAsync(await fs.readFile(toc.docxPath));
  const tocXml = await tocZip.file("word/document.xml").async("text");
  const tocSettings = await tocZip.file("word/settings.xml").async("text");
  assert.match(tocXml, /w:fldCharType="begin"/);
  assert.match(tocXml, /<w:instrText[^>]*> TOC \\o (?:"|&quot;)1-4(?:"|&quot;) \\h \\z \\u <\/w:instrText>/);
  assert.match(tocXml, /w:fldCharType="separate"/);
  assert.match(tocXml, /w:fldCharType="end"/);
  assert.match(tocSettings, /<w:updateFields\b[^>]*w:val="true"/);

  const inlineFields = await runFixture("open-chestnut-inline-fields", {
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  const inlineFieldDocument = await DocumentFile.importDocx(await FileBlob.load(inlineFields.docxPath));
  const inlineFieldParagraph = inlineFieldDocument.blocks.find((block) => block.name === "field-caption" || block.text.startsWith("Figure 1:"));
  assert.equal(inlineFieldParagraph?.text, "Figure 1: Updated revenue. See figure 1 on page 1.");
  assert.deepEqual(inlineFieldParagraph?.runs.filter((run) => run.inlineField).map((run) => run.inlineField.instruction), [
    "SEQ Figure \\* ARABIC",
    "REF fig1 \\h",
    "PAGEREF fig1 \\h",
  ]);
  assert.equal(inlineFieldParagraph?.runs.find((run) => run.inlineField?.instruction.startsWith("SEQ "))?.inlineField.bookmarkName, "fig1");
  assert.equal(inlineFieldParagraph?.runs.find((run) => run.inlineField?.instruction.startsWith("SEQ "))?.inlineField.bookmarkNativeId, 0);
  assert.equal(inlineFields.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  const inlineFieldZip = await JSZip.loadAsync(await fs.readFile(inlineFields.docxPath));
  const inlineFieldXml = await inlineFieldZip.file("word/document.xml").async("text");
  assert.equal((inlineFieldXml.match(/w:fldCharType="begin"/g) || []).length, 3);
  assert.equal((inlineFieldXml.match(/w:fldCharType="separate"/g) || []).length, 3);
  assert.equal((inlineFieldXml.match(/w:fldCharType="end"/g) || []).length, 3);
  assert.match(inlineFieldXml, /SEQ Figure \\[*] ARABIC/);
  assert.match(inlineFieldXml, /REF fig1 \\h/);
  assert.match(inlineFieldXml, /PAGEREF fig1 \\h/);
  const inlineBookmarkStart = inlineFieldXml.indexOf('w:name="fig1"');
  const inlineBookmarkResult = inlineFieldXml.indexOf("<w:t>1</w:t>", inlineBookmarkStart);
  const inlineBookmarkEnd = inlineFieldXml.indexOf("<w:bookmarkEnd", inlineBookmarkStart);
  assert.ok(inlineBookmarkStart >= 0 && inlineBookmarkResult > inlineBookmarkStart && inlineBookmarkEnd > inlineBookmarkResult, "Caption-number bookmark must wrap only the SEQ cached-result run.");

  const classicFixture = await runFixture("package-comments");
  const classicDocument = await DocumentFile.importDocx(await FileBlob.load(classicFixture.docxPath));
  assert.equal(classicDocument.comments.length, 1);
  assert.equal(classicDocument.comments[0].author, "QA Lead");
  assert.equal(classicDocument.comments[0].text, "Decision paragraph confirmed.");
  const classicSourceBytes = await fs.readFile(classicFixture.docxPath);
  const classicTarget = classicDocument.blocks.find((block) => block.id === classicDocument.comments[0].targetId);
  assert.equal(classicTarget?.kind, "paragraph");
  const { editClassicComment } = await import(
    "../skills/documents/skills/documents/examples/openchestnut-classic-comment-edit-workflow.mjs"
  );
  const classicWorkflowOutput = path.join(outputDir, "classic-comment-updated.docx");
  const classicWorkflowAudit = path.join(outputDir, "classic-comment-audit.json");
  const classicWorkflow = await editClassicComment({
    inputPath: classicFixture.docxPath,
    outputPath: classicWorkflowOutput,
    auditPath: classicWorkflowAudit,
    anchorText: classicTarget.text,
    expectedCommentText: classicDocument.comments[0].text,
    replacementText: "Decision paragraph approved after QA.",
  });
  assert.equal(classicWorkflow.audit.provider.actual, "open-chestnut");
  assert.equal(classicWorkflow.audit.validation.reimport.ok, true);
  assert.equal(classicWorkflow.audit.validation.modelRender.renderer, "model-svg");
  assert.deepEqual(await fs.readFile(classicFixture.docxPath), classicSourceBytes);
  const classicWorkflowDocument = await DocumentFile.importDocx(await FileBlob.load(classicWorkflowOutput));
  assert.equal(classicWorkflowDocument.comments.length, 1);
  assert.equal(classicWorkflowDocument.comments[0].id, classicDocument.comments[0].id);
  assert.equal(classicWorkflowDocument.comments[0].targetId, classicDocument.comments[0].targetId);
  assert.equal(classicWorkflowDocument.comments[0].author, "QA Lead");
  assert.equal(classicWorkflowDocument.comments[0].text, "Decision paragraph approved after QA.");

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
  const skillText = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "SKILL.md"), "utf8");
  const pluginReadme = await fs.readFile(path.join(repoRoot, "skills", "documents", "README.md"), "utf8");
  assert.match(pluginReadme, /open-office-artifact-tool/);
  assert.match(pluginReadme, /OpenChestnut/);
  assert.match(skillText, /render_docx\.py/);
  assert.match(skillText, /DocumentModel/);
  assert.match(skillText, /DocumentFile/);
  assert.match(skillText, /OpenChestnut/);
  assert.match(skillText, /artifact_tool\/API_QUICK_START\.md/);
  assert.match(skillText, /document\.addInsertion/);
  assert.match(skillText, /document\.addDeletion/);
  assert.match(skillText, /paragraph\.addTextContentControl/);
  assert.match(skillText, /document\.fillContentControls/);
  assert.match(skillText, /document\.addBibliographySource/);
  assert.match(skillText, /document\.addCitation/);
  assert.match(skillText, /document\.addTableOfContents/);
  assert.match(skillText, /paragraph\.addField/);
  assert.match(skillText, /openchestnut-classic-comment-edit-workflow\.mjs/);
  assert.doesNotMatch(skillText, /Author\/edit with `python-docx`|Default tool: python-docx/);
  const commentsGuide = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "tasks", "comments_manage.md"), "utf8");
  assert.match(commentsGuide, /document\.addComment/);
  assert.doesNotMatch(commentsGuide, /If the task is to \*insert\* new comments.+use the OOXML-level guide/);
  const controlsGuide = await fs.readFile(path.join(repoRoot, "skills", "documents", "skills", "documents", "tasks", "forms_content_controls.md"), "utf8");
  assert.match(controlsGuide, /paragraph\.addTextContentControl/);
  assert.match(controlsGuide, /document\.fillContentControls/);
  assert.match(controlsGuide, /Rich.*block.*cell.*dropdown.*date.*checkbox/is);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("document skill smoke ok");
