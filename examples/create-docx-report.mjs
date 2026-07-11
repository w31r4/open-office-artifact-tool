import os from "node:os";
import path from "node:path";

import { DocumentModel, DocumentFile } from "open-office-artifact-tool";

const outputDir = process.env.OUTPUT_DIR || path.join(os.tmpdir(), "open-office-artifact-examples");

const document = DocumentModel.create();
document.styles.add("ExecutiveSummary", { name: "Executive Summary", fontSize: 28, bold: true });
document.addHeader("Clean-room quarterly report");
document.addParagraph("Quarterly readiness", { styleId: "ExecutiveSummary" });
document.addParagraph("This DOCX was generated with public WordprocessingML building blocks.");
document.addListItem("Revenue trend is stable", { level: 0 });
document.addListItem("Native image/section/tracked-change facades are available", { level: 0 });
document.addTable({ name: "summary-table", values: [["Metric", "Value"], ["Revenue", "$12M"], ["Retention", "94%"]] });
document.addHyperlink("Project repository", "https://github.com/w31r4/open-office-artifact-tool");
document.addCitation("Source: internal clean-room research", { url: "https://example.com/source" });
document.addComment(document.blocks[1].id, "Review title wording before publishing.");
document.addInsertion("Added tracked insertion", { author: "Analyst" });
document.addSection({ orientation: "portrait", pageSize: { widthTwips: 12240, heightTwips: 15840 } });

const file = await DocumentFile.exportDocx(document);
await file.save(path.join(outputDir, "docx-report.docx"));
console.log(document.inspect({ kind: "paragraph,table,comment", maxChars: 4000 }).ndjson);
console.log(`saved ${path.join(outputDir, "docx-report.docx")}`);
