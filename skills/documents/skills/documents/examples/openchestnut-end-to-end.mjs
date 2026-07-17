import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import {
  DocumentFile,
  DocumentModel,
} from "open-office-artifact-tool";

export const DEFAULT_BRIEF = Object.freeze({
  title: "Launch readiness decision brief",
  subtitle: "A concise, evidence-led recommendation for the next release gate",
  owner: "Artifact Platform",
  date: "16 July 2026",
  recommendation: "Approve the controlled rollout after the final compatibility review.",
  summary: "The core authoring, OpenChestnut codec, package, and render gates are stable. Remaining work is concentrated in environment-specific compatibility and release operations.",
  evidence: [
    ["Core workflow", "Pass", "Create, import, edit, export, and second import"],
    ["Package integrity", "Pass", "Bundled runtime works without a local .NET SDK"],
    ["Visual QA", "Pass", "DOCX is rendered and reviewed before delivery"],
  ],
  nextSteps: [
    "Complete the final application-compatibility review.",
    "Record the release owner and publish window.",
    "Archive the rendered QA evidence with the release record.",
  ],
  sourceLabel: "Project release evidence",
  sourceUrl: "https://github.com/w31r4/open-office-artifact-tool",
});

function addStyles(document) {
  const styles = {
    BriefTitle: {
      name: "Brief Title",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos Display",
      fontSize: 23,
      bold: true,
      color: "#123B5D",
      spaceAfterTwips: 80,
      keepNext: true,
    },
    BriefSubtitle: {
      name: "Brief Subtitle",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos",
      fontSize: 10.5,
      color: "#486779",
      spaceAfterTwips: 160,
      keepNext: true,
    },
    BriefMeta: {
      name: "Brief Metadata",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos",
      fontSize: 8.5,
      bold: true,
      color: "#647784",
      spaceAfterTwips: 120,
      keepNext: true,
    },
    BriefHeading1: {
      name: "Brief Heading 1",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos Display",
      fontSize: 13,
      bold: true,
      color: "#123B5D",
      spaceBeforeTwips: 160,
      spaceAfterTwips: 70,
      keepNext: true,
    },
    BriefCallout: {
      name: "Brief Callout",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos",
      fontSize: 10,
      bold: true,
      color: "#0F5E6B",
      spaceBeforeTwips: 70,
      spaceAfterTwips: 140,
      keepNext: true,
    },
    TableGrid: {
      name: "Table Grid",
      type: "table",
      fontFamily: "Aptos",
      fontSize: 8.5,
    },
  };
  for (const [id, style] of Object.entries(styles)) document.styles.add(id, style);
}

export function buildDocument(spec = DEFAULT_BRIEF) {
  const document = DocumentModel.create({
    name: spec.title,
    designPreset: "standard_business_brief",
    defaultRunStyle: { fontFamily: "Aptos", fontSize: 9.5, color: "#172033" },
    blocks: [],
  });
  addStyles(document);

  document.addParagraph(spec.title, {
    name: "brief-title",
    styleId: "BriefTitle",
  });
  document.addParagraph(spec.subtitle, {
    name: "brief-subtitle",
    styleId: "BriefSubtitle",
  });
  const metadata = document.addParagraph("", {
    name: "brief-metadata",
    styleId: "BriefMeta",
    runs: [
      { text: "OWNER  " },
      {
        text: "{{OWNER}}",
        contentControl: {
          id: "brief-owner",
          tag: "OWNER",
          alias: "Brief owner",
        },
      },
      { text: `    |    DATE  ${spec.date}` },
    ],
  });
  assert.equal(metadata.text.includes("{{OWNER}}"), true);

  const decisionHeading = document.addParagraph("Decision", { styleId: "BriefHeading1" });
  const decisionBookmark = document.addBookmark(decisionHeading, "DecisionSection");
  const recommendation = document.addParagraph(`Recommendation: ${spec.recommendation}`, {
    name: "recommendation",
    styleId: "BriefCallout",
    paragraphFormat: {
      leftIndentTwips: 240,
      rightIndentTwips: 240,
      spaceBeforeTwips: 70,
      spaceAfterTwips: 140,
      keepNext: true,
    },
  });
  document.addParagraph(spec.summary, {
    name: "executive-summary",
    styleId: "Normal",
    paragraphFormat: { spaceAfterTwips: 100 },
  });
  const footnoteTarget = document.addParagraph("The decision depends on a final application-compatibility gate.", {
    name: "footnote-target",
    styleId: "Normal",
  });
  document.addFootnote(footnoteTarget, "The final gate includes native rendering and package validation.", {
    name: "release-gate-footnote",
  });
  const endnoteTarget = document.addParagraph("The evidence snapshot is retained with the release record.", {
    name: "endnote-target",
    styleId: "Normal",
  });
  document.addEndnote(endnoteTarget, "Evidence snapshot dated 2026-07-17.", {
    name: "evidence-endnote",
  });
  document.addInsertion("Final compatibility review is required before rollout.", {
    name: "tracked-release-condition",
    styleId: "Normal",
    author: "Release reviewer",
    date: "2026-07-16T08:10:00Z",
  });
  document.addDeletion("Immediate unrestricted rollout.", {
    name: "tracked-superseded-condition",
    styleId: "Normal",
    author: "Release reviewer",
    date: "2026-07-16T08:11:00Z",
  });

  document.addParagraph("Evidence", { styleId: "BriefHeading1" });
  document.addTable({
    name: "readiness-evidence",
    styleId: "TableGrid",
    widthDxa: 9000,
    indentDxa: 120,
    columnWidthsDxa: [2100, 1800, 5100],
    cellMarginsDxa: { top: 60, right: 100, bottom: 60, left: 100 },
    borderColor: "AFC1CC",
    borderSize: 6,
    headerFill: "DCEAF3",
    values: [["Gate", "Status", "Evidence"], ...spec.evidence],
  });

  document.addParagraph("Next steps", { styleId: "BriefHeading1" });
  spec.nextSteps.forEach((step, index) => document.addListItem(step, {
    name: `next-step-${index + 1}`,
    styleId: "Normal",
    listType: "number",
    numberFormat: "decimal",
    start: 1,
    level: 0,
    levelText: "%1.",
    numberingId: 41,
    abstractNumberingId: 4,
  }));

  document.addHyperlink(`Source — ${spec.sourceLabel}`, spec.sourceUrl, {
    name: "project-source",
    styleId: "Normal",
    tooltip: "Open project evidence",
    history: true,
  });
  document.addHyperlink("Back to decision", decisionBookmark, {
    name: "decision-jump",
    styleId: "Normal",
    tooltip: "Jump to the decision section",
    history: true,
  });
  document.addBibliographySource({
    id: "bibliography/ProjectEvidence",
    tag: "ProjectEvidence",
    sourceType: "InternetSite",
    title: spec.sourceLabel,
    year: "2026",
    url: spec.sourceUrl,
    corporateAuthor: spec.owner,
  });
  document.addCitation(`(${spec.owner}, 2026)`, { tag: "ProjectEvidence" }, {
    id: "citation/project-evidence",
    styleId: "Normal",
  });

  document.addComment(recommendation, "Confirm the recommendation wording before publication.", {
    author: "Release reviewer",
    initials: "RR",
    date: "2026-07-16T08:00:00Z",
  });
  document.addHeader("LAUNCH READINESS | DECISION BRIEF", {
    referenceType: "default",
    sectionIndex: 0,
  });
  document.addFooter("1", {
    referenceType: "default",
    sectionIndex: 0,
    fieldInstruction: "PAGE",
  });
  return document;
}

export async function createDocument(outputPath, spec = DEFAULT_BRIEF) {
  const authored = buildDocument(spec);
  const authoredReport = authored.verify({ visualQa: true });
  assert.equal(authoredReport.ok, true, authoredReport.ndjson || JSON.stringify(authoredReport.issues));

  const firstDocx = await DocumentFile.exportDocx(authored);
  const imported = await DocumentFile.importDocx(firstDocx);
  assert.deepEqual(imported.fillContentControls({ OWNER: spec.owner }), {
    updated: 1,
    matchedTags: ["OWNER"],
    missingTags: [],
  });
  assert.equal(imported.bookmarks.length, 2);
  const importedDecisionBookmark = imported.bookmarks.find((bookmark) => bookmark.name === "DecisionSection");
  assert.ok(importedDecisionBookmark);
  assert.equal(
    imported.resolve(importedDecisionBookmark.id).targetId,
    imported.blocks.find((block) => block.text === "Decision")?.id,
  );
  assert.equal(
    imported.blocks.some((block) => block.kind === "hyperlink" && block.anchor === "DecisionSection"),
    true,
  );
  assert.deepEqual(imported.notes.map((note) => [note.kind, note.text]), [
    ["footnote", "The final gate includes native rendering and package validation."],
    ["endnote", "Evidence snapshot dated 2026-07-17."],
  ]);
  const recommendation = imported.blocks.find(
    (block) => block.kind === "paragraph" && block.text.startsWith("Recommendation:"),
  );
  assert.ok(recommendation, "OpenChestnut must re-import the recommendation paragraph");
  imported.resolve(`${recommendation.id}/text`).text = `Recommendation: ${spec.recommendation}`;

  const evidence = imported.blocks.find((block) => block.kind === "table");
  assert.ok(evidence, "OpenChestnut must re-import the evidence table");
  evidence.getCell(1, 1).value = "Verified";
  imported.comments[0].text = "Recommendation wording verified for the release record.";
  const trackedInsertion = imported.blocks.find(
    (block) => block.kind === "change" && block.changeType === "insert",
  );
  assert.ok(trackedInsertion, "OpenChestnut must re-import the tracked insertion");
  trackedInsertion.text = "Final application-compatibility review is required before rollout.";
  trackedInsertion.author = "Lead reviewer";
  trackedInsertion.date = "2026-07-16T09:00:00Z";
  imported.notes[0].text = "The final gate includes native rendering, package validation, and semantic re-import.";
  imported.notes[1].text = "Evidence snapshot dated 2026-07-17; retained with the release record.";
  imported.bibliographySources[0].title = `${spec.sourceLabel} — verified`;
  imported.blocks.find((block) => block.kind === "citation").text = `(${spec.owner}, 2026, verified)`;

  const finalDocx = await DocumentFile.exportDocx(imported);
  const finalDocument = await DocumentFile.importDocx(finalDocx);
  const finalReport = finalDocument.verify({ visualQa: true });
  assert.equal(finalReport.ok, true, finalReport.ndjson || JSON.stringify(finalReport.issues));
  assert.equal(finalDocument.blocks.find((block) => block.kind === "table")?.getCell(1, 1).value, "Verified");
  assert.equal(finalDocument.comments[0]?.text, "Recommendation wording verified for the release record.");
  assert.equal(finalDocument.bookmarks.some((bookmark) => bookmark.name === "DecisionSection"), true);
  assert.equal(finalDocument.bibliographySources[0]?.title, `${spec.sourceLabel} — verified`);
  assert.equal(finalDocument.blocks.find((block) => block.kind === "citation")?.text, `(${spec.owner}, 2026, verified)`);
  assert.deepEqual(finalDocument.contentControls.map((control) => [control.tag, control.alias, control.text]), [
    ["OWNER", "Brief owner", spec.owner],
  ]);
  assert.equal(
    finalDocument.blocks.some((block) => block.kind === "hyperlink" && block.anchor === "DecisionSection"),
    true,
  );
  assert.deepEqual(finalDocument.notes.map((note) => [note.kind, note.text]), [
    ["footnote", "The final gate includes native rendering, package validation, and semantic re-import."],
    ["endnote", "Evidence snapshot dated 2026-07-17; retained with the release record."],
  ]);
  assert.deepEqual(
    finalDocument.blocks.filter((block) => block.kind === "change").map(
      (block) => [block.changeType, block.text, block.author],
    ),
    [
      ["insert", "Final application-compatibility review is required before rollout.", "Lead reviewer"],
      ["delete", "Immediate unrestricted rollout.", "Release reviewer"],
    ],
  );

  const inspection = finalDocument.inspect({
    kind: "document,paragraph,listItem,table,comment,bookmark,note,contentControl,header,footer,hyperlink,citation,bibliographySource,change,layout",
    maxChars: 32_000,
  });
  for (const expected of [spec.title, "Verified", "Recommendation wording verified", "application-compatibility", "semantic re-import", "Evidence snapshot", "DecisionSection", "ProjectEvidence", "2026, verified", "OWNER", "LAUNCH READINESS"]) {
    assert.match(inspection.ndjson, new RegExp(expected));
  }

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await finalDocx.save(outputPath);
  return { document: finalDocument, file: finalDocx, inspection, verification: finalReport };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-decision-brief.docx");
  const result = await createDocument(outputPath);
  console.log(JSON.stringify({
    outputPath,
    bytes: result.file.bytes.length,
    blocks: result.document.blocks.length,
    comments: result.document.comments.length,
    notes: result.document.notes.length,
    verified: result.verification.ok,
  }));
}
