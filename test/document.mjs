import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { DocumentFile, DocumentModel, FileBlob } from "../src/index.mjs";

const document = DocumentModel.create({
  name: "Research memo",
  paragraphs: ["Research memo", "This document exercises the clean-room DOCX facade."],
});
document.styles.add("Callout", { name: "Callout", fontSize: 24, bold: true, fontFamily: "Aptos" });
const header = document.addHeader("Confidential research memo", { name: "default-header" });
const footer = document.addFooter("Page footer", { name: "default-footer" });
const heading = document.addParagraph("Findings", { styleId: "Heading1", name: "findings-heading" });
const bullet = document.addListItem("Use real numbering definitions", { listType: "bullet", name: "numbering-rule" });
const numbered = document.addListItem("Render and verify", { listType: "number", name: "render-step" });
const table = document.addTable({
  name: "evidence-table",
  styleId: "TableGrid",
  values: [["Area", "Status"], ["DOCX styles", "partial"], ["Comments", "roundtrip"]],
});
table.getCell(2, 1).value = "anchored";
const comment = document.addComment(heading, "Check this heading before final export.", { author: "Reviewer" });

const inspect = document.inspect({ kind: "paragraph,table,comment,style,listItem,header,footer", maxChars: 10000 }).ndjson;
assert.match(inspect, /Research memo/);
assert.match(inspect, /findings-heading/);
assert.match(inspect, /numbering-rule/);
assert.match(inspect, /render-step/);
assert.match(inspect, /Confidential research memo/);
assert.match(inspect, /Page footer/);
assert.match(inspect, /evidence-table/);
assert.match(inspect, /Check this heading/);
assert.match(inspect, /Callout/);
assert.equal(document.resolve(heading.id).styleId, "Heading1");
assert.equal(document.resolve(bullet.id).listType, "bullet");
assert.equal(document.resolve(numbered.id).listType, "number");
assert.equal(document.resolve(header.id).text, "Confidential research memo");
assert.equal(document.resolve(footer.id).text, "Page footer");
assert.equal(document.resolve(table.id).getCell(2, 1).value, "anchored");
assert.equal(document.resolve(comment.id).author, "Reviewer");
assert.match(document.help("document.addTable").ndjson, /Word-style table/);
assert.match(document.help("document.addListItem").ndjson, /numbering definitions/);
assert.match(document.help("document.addHeader").ndjson, /DOCX header/);

const preview = await document.render();
assert.equal(preview.type, "image/svg+xml");
const svg = await preview.text();
assert.match(svg, /Research memo/);
assert.match(svg, /DOCX styles/);
assert.match(svg, /Confidential research memo/);
assert.match(svg, /Render and verify/);

const docx = await DocumentFile.exportDocx(document);
assert.equal(docx.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.docx`);
await docx.save(out);
const loaded = await DocumentFile.importDocx(await FileBlob.load(out));
const loadedInspect = loaded.inspect({ kind: "paragraph,table,comment,listItem,header,footer", maxChars: 10000 }).ndjson;
assert.match(loadedInspect, /clean-room DOCX facade/);
assert.match(loadedInspect, /DOCX styles/);
assert.match(loadedInspect, /anchored/);
assert.match(loadedInspect, /Check this heading/);
assert.match(loadedInspect, /Use real numbering definitions/);
assert.match(loadedInspect, /Render and verify/);
assert.match(loadedInspect, /Confidential research memo/);
assert.match(loadedInspect, /Page footer/);
console.log("document smoke ok");
