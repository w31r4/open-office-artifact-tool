import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { DocumentFile, FileBlob } from "open-office-artifact-tool";
import {
  nativeDocumentRenderStatus,
  runDocumentFixture,
  verifyDocumentFile,
} from "../skills/documents/scripts/workflow.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const fixturePath = path.join(repoRoot, "skills", "documents", "fixtures", "business-brief.json");
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-document-skill-"));
const baselineDir = path.join(outputDir, "baselines");

try {
  const result = await runDocumentFixture(fixturePath, { outputDir, nativeRender: "off" });
  assert.equal(result.fixture.name, "business-brief");
  assert.equal(result.qa.summary.verifyOk, true);
  assert.equal(result.qa.summary.packageOk, true);
  assert.equal(result.qa.summary.visualQaOk, true);
  assert.equal(result.qa.summary.nativeRender.status, "skipped");
  for (const filePath of Object.values(result.qa.summary.files)) {
    const stat = await fs.stat(filePath);
    assert.ok(stat.isFile() && stat.size > 0, `Expected non-empty document skill output ${filePath}`);
  }
  const imported = await DocumentFile.importDocx(await FileBlob.load(result.docxPath));
  const inspect = imported.inspect({ kind: "theme,paragraph,listItem,table,comment,header,hyperlink,citation,image,field,section", maxChars: 24_000 }).ndjson;
  assert.match(inspect, /Office artifact readiness brief/);
  assert.match(inspect, /readiness-table/);
  assert.match(inspect, /native render review/);
  assert.match(inspect, /Opening section evidence/);
  assert.match(inspect, /Business Brief Theme/);
  assert.match(inspect, /Theme fidelity/);
  assert.equal(imported.headers.find((item) => item.name === "opening-header")?.sectionIndex, 0);
  assert.match(await fs.readFile(result.qa.summary.files.packageInspect, "utf8"), /word\/document\.xml/);
  assert.match(await fs.readFile(result.qa.summary.files.packageInspect, "utf8"), /word\/commentsExtended\.xml/);
  assert.match(await fs.readFile(result.qa.summary.files.packageInspect, "utf8"), /word\/commentsIds\.xml/);
  assert.match(await fs.readFile(result.qa.summary.files.packageInspect, "utf8"), /word\/commentsExtensible\.xml/);
  assert.match(await fs.readFile(result.qa.summary.files.packageInspect, "utf8"), /word\/people\.xml/);
  assert.match(await fs.readFile(result.qa.summary.files.preview, "utf8"), /<svg/);
  const nativePreferred = await verifyDocumentFile(result.docxPath, { outputDir: path.join(outputDir, "native-preferred"), preferNative: true, nativeRender: "off" });
  assert.equal(nativePreferred.summary.verifyOk, true);
  assert.match(nativePreferred.inspect.ndjson, /Office artifact readiness brief/);
  const nativePreferredDocument = await DocumentFile.importDocx(await FileBlob.load(result.docxPath), { preferNative: true });
  assert.equal(nativePreferredDocument.theme.name, "Business Brief Theme");
  assert.equal(nativePreferredDocument.theme.fonts.majorEastAsia, "Arial Unicode MS");
  const themeRuns = nativePreferredDocument.blocks.find((item) => item.text === "Theme fidelity East Asia")?.runs;
  assert.equal(themeRuns?.[0].style.resolvedColor, "#99b3cc");
  assert.equal(themeRuns?.[1].style.resolvedFontFamilyEastAsia, "Arial Unicode MS");
  const complexThemeRun = nativePreferredDocument.blocks.find((item) => item.text === "العربية")?.runs[0];
  assert.equal(complexThemeRun?.style.resolvedFontFamilyComplexScript, "Geeza Pro");
  assert.equal(complexThemeRun?.style.boldComplexScript, true);
  const inheritedRun = nativePreferredDocument.blocks.find((item) => item.text === "Inherited character style")?.runs[0];
  assert.equal(nativePreferredDocument.defaultRunStyle.fontFamily, "Default Serif");
  assert.equal(inheritedRun?.style.runStyleId, "BriefCharacter");
  assert.equal(inheritedRun?.style.resolvedColor, "#cc3300");
  assert.equal(inheritedRun?.style.resolvedFontFamily, "Source Serif 4");
  assert.equal(inheritedRun?.style.fontSize, 26);
  assert.equal(inheritedRun?.style.bold, false);
  assert.equal(inheritedRun?.style.italic, false);
  assert.equal(nativePreferredDocument.headers.find((item) => item.text === "Opening section evidence")?.sectionIndex, 0);
  assert.equal(nativePreferredDocument.headers.find((item) => item.text === "Clean-room document workflow")?.sectionIndex, 1);
  assert.equal(nativePreferredDocument.comments.find((item) => item.text.includes("native render review"))?.author, "QA Agent");
  assert.equal(nativePreferredDocument.comments.find((item) => item.text.includes("native render review"))?.date, "2026-07-11T00:20:00.000Z");
  const nativeReviewRoot = nativePreferredDocument.comments.find((item) => item.text.includes("native render review"));
  const nativeReviewReply = nativePreferredDocument.comments.find((item) => item.text.includes("wording is approved"));
  assert.equal(nativeReviewRoot?.resolved, true);
  assert.equal(nativeReviewRoot?.durableId, "0010A001");
  assert.equal(nativeReviewRoot?.dateUtc, "2026-07-11T00:20:00.000Z");
  assert.deepEqual(nativeReviewRoot?.person, { providerId: "None", userId: "qa-agent@example.test" });
  assert.equal(nativeReviewReply?.parentId, nativeReviewRoot?.id);
  assert.equal(nativeReviewReply?.durableId, "0010A002");
  assert.deepEqual(nativeReviewReply?.person, { providerId: "None", userId: "maintainer@example.test" });
  assert.equal(nativeReviewReply?.targetId, nativeReviewRoot?.targetId);
  const businessBriefZip = await JSZip.loadAsync(await fs.readFile(result.docxPath));
  assert.ok(businessBriefZip.file("word/commentsExtended.xml"));
  assert.ok(businessBriefZip.file("word/commentsIds.xml"));
  assert.ok(businessBriefZip.file("word/commentsExtensible.xml"));
  assert.ok(businessBriefZip.file("word/people.xml"));
  assert.match(await businessBriefZip.file("word/commentsExtended.xml").async("text"), /w15:paraIdParent=/);
  assert.match(await businessBriefZip.file("word/commentsIds.xml").async("text"), /w16cid:durableId="0010A001"/);
  assert.match(await businessBriefZip.file("word/commentsExtensible.xml").async("text"), /w16cex:dateUtc="2026-07-11T00:20:00\.000Z"/);
  assert.match(await businessBriefZip.file("word/people.xml").async("text"), /w15:userId="qa-agent@example\.test"/);
  const businessBriefRels = await businessBriefZip.file("word/_rels/document.xml.rels").async("text");
  assert.match(businessBriefRels, /relationships\/commentsExtended/);
  assert.match(businessBriefRels, /relationships\/commentsIds/);
  assert.match(businessBriefRels, /relationships\/commentsExtensible/);
  assert.match(businessBriefRels, /relationships\/people/);
  const nativeTableComment = nativePreferredDocument.comments.find((item) => item.text.includes("table comment anchor"));
  assert.equal(nativeTableComment?.author, "Maintainer");
  assert.equal(nativePreferredDocument.resolve(nativeTableComment?.targetId)?.kind, "table");
  assert.equal(nativePreferredDocument.blocks.find((item) => item.kind === "hyperlink")?.url, "https://learn.microsoft.com/office/open-xml/open-xml-sdk");
  assert.equal(nativePreferredDocument.blocks.find((item) => item.kind === "field")?.instruction, "PAGE");
  assert.match(nativePreferredDocument.blocks.find((item) => item.kind === "citation")?.metadata?.bookmark || "", /^OpenOfficeCitation_/);
  assert.equal(nativePreferredDocument.blocks.find((item) => item.text === "Preserve native numbering definitions.")?.numberFormat, "upperLetter");
  assert.equal(nativePreferredDocument.blocks.find((item) => item.text === "Resolve nested numbering levels.")?.numberFormat, "lowerRoman");
  assert.equal(nativePreferredDocument.blocks.find((item) => item.text === "Resolve nested numbering levels.")?.level, 1);

  const nativeStatus = nativeDocumentRenderStatus();
  const packageComments = await runDocumentFixture(path.join(repoRoot, "skills", "documents", "fixtures", "package-comments.json"), {
    outputDir: path.join(outputDir, "package-comments"),
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(packageComments.qa.summary.packageOk, true);
  assert.equal(packageComments.qa.summary.verifyOk, true);
  assert.ok(packageComments.qa.packageInspect.parts.some((part) => part.path === "word/review/agent-comments.xml"));
  assert.ok(packageComments.qa.packageInspect.parts.some((part) => part.path === "word/review/agent-comments-extended.xml"));
  const packageCommentDocument = packageComments.qa.document;
  const paragraphComment = packageCommentDocument.comments.find((comment) => comment.text === "Confirm the decision paragraph.");
  const tableComment = packageCommentDocument.comments.find((comment) => comment.text === "Confirm the table-cell anchor.");
  assert.equal(paragraphComment?.author, "QA Agent");
  assert.equal(paragraphComment?.resolved, true);
  assert.equal(paragraphComment?.paraId, "A0000015");
  assert.equal(packageCommentDocument.resolve(paragraphComment?.targetId)?.text, "Approve after native package review.");
  assert.equal(tableComment?.author, "Maintainer");
  assert.equal(packageCommentDocument.resolve(tableComment?.targetId)?.kind, "table");
  const paragraphReply = packageCommentDocument.comments.find((comment) => comment.text === "The decision paragraph is approved.");
  assert.equal(paragraphReply?.parentId, paragraphComment?.id);
  assert.equal(paragraphReply?.targetId, paragraphComment?.targetId);
  assert.equal(packageComments.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const packageNumbering = await runDocumentFixture(path.join(repoRoot, "skills", "documents", "fixtures", "package-numbering.json"), {
    outputDir: path.join(outputDir, "package-numbering"),
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(packageNumbering.qa.summary.packageOk, true);
  assert.equal(packageNumbering.qa.summary.verifyOk, true);
  assert.ok(packageNumbering.qa.packageInspect.parts.some((part) => part.path === "word/review/agent-numbering.xml"));
  const packageNumberingDocument = packageNumbering.qa.document;
  const primaryStep = packageNumberingDocument.blocks.find((block) => block.text === "Validate the numbering definition.");
  const nestedStep = packageNumberingDocument.blocks.find((block) => block.text === "Confirm the nested level.");
  assert.equal(primaryStep?.kind, "listItem");
  assert.equal(primaryStep?.numberFormat, "upperLetter");
  assert.equal(primaryStep?.start, 2);
  assert.equal(primaryStep?.numberingId, 77);
  assert.equal(nestedStep?.kind, "listItem");
  assert.equal(nestedStep?.numberFormat, "lowerRoman");
  assert.equal(nestedStep?.level, 1);
  assert.equal(nestedStep?.start, 3);
  assert.equal(packageNumbering.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const packageSettings = await runDocumentFixture(path.join(repoRoot, "skills", "documents", "fixtures", "package-settings.json"), {
    outputDir: path.join(outputDir, "package-settings"),
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(packageSettings.qa.summary.packageOk, true);
  assert.equal(packageSettings.qa.summary.verifyOk, true);
  assert.ok(packageSettings.qa.packageInspect.parts.some((part) => part.path === "word/review/agent-settings.xml"));
  assert.deepEqual(packageSettings.qa.document.settings, {
    trackRevisions: true,
    updateFields: true,
    evenAndOddHeaders: true,
    mirrorMargins: true,
    documentProtection: { edit: "comments", enforcement: true, formatting: false },
  });
  assert.match(packageSettings.qa.inspect.ndjson, /"kind":"settings"/);
  assert.equal(packageSettings.qa.summary.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const baselineWrite = await verifyDocumentFile(result.docxPath, {
    outputDir: path.join(outputDir, "baseline-write"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "off",
    baselineDir,
    writeBaseline: true,
  });
  assert.equal(baselineWrite.summary.writeBaseline, true);
  assert.ok((await fs.stat(baselineWrite.summary.modelBaselinePath)).size > 100);
  const baselineCompare = await verifyDocumentFile(result.docxPath, {
    outputDir: path.join(outputDir, "baseline-compare"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "off",
    baselineDir,
  });
  assert.equal(baselineCompare.summary.modelBaselineCompared, true);
  assert.equal(baselineCompare.summary.modelPixelDiff.changed, false);
  assert.equal(baselineCompare.summary.visualQaOk, true);
  if (nativeStatus.available) {
    assert.equal(baselineCompare.summary.nativeRender.status, "passed");
    assert.equal(baselineCompare.summary.nativeRender.ok, true);
    assert.equal(baselineCompare.summary.nativeRender.pageCountMatches, true);
    assert.ok(baselineCompare.summary.nativeRender.pageCount >= 1);
    for (const page of baselineCompare.summary.nativeRender.pages) {
      assert.equal(page.baselineCompared, true);
      assert.equal(page.pixelDiff.changed, false);
      assert.equal(page.ok, true);
      assert.ok((await fs.stat(page.path)).size > 100);
    }
  }

  const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
  assert.ok(packageJson.files.includes("skills/**"));
  const skillText = await fs.readFile(path.join(repoRoot, "skills", "documents", "SKILL.md"), "utf8");
  assert.match(skillText, /LibreOffice PDF plus Poppler page PNGs/);
  assert.match(skillText, /baseline-dir/);
  assert.match(skillText, /preferNative/);
  assert.match(skillText, /package-numbering/);
  assert.match(skillText, /package-settings/);
  assert.match(skillText, /themeColor/);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("document skill smoke ok");
