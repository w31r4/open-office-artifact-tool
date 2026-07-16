import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import {
  PdfArtifact,
  PdfFile,
  verifyArtifact,
} from "open-office-artifact-tool";
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";

const STATUS_MARK = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

export function buildPdf() {
  const pdf = PdfArtifact.create({
    metadata: { title: "Release readiness scorecard", language: "en-US" },
    pages: [{
      text: "Release readiness scorecard\nDecision-ready evidence for the launch gate",
      width: 612,
      height: 792,
    }],
  });

  const decisionHeading = pdf.addText("Decision", {
    id: "decision-heading",
    bbox: [72, 145, 180, 24],
    fontSize: 18,
    bold: true,
    color: "#123B5D",
    headingLevel: 2,
  });
  const decision = pdf.addText("Approve the controlled rollout after final native-render QA.", {
    id: "decision-recommendation",
    bbox: [72, 177, 440, 18],
    fontSize: 11,
    color: "#172033",
  });
  const evidence = pdf.addTable({
    id: "readiness-evidence",
    name: "readiness-evidence",
    values: [
      ["Gate", "Owner", "Status"],
      ["Semantic model", "Artifact Platform", "Pass"],
      ["Package install", "Release Engineering", "Pass"],
      ["Native render", "Release QA", "Pending"],
    ],
    bbox: [72, 225, 468, 112],
  });
  const trend = pdf.addChart({
    id: "readiness-trend",
    name: "readiness-trend",
    title: "Readiness by gate",
    alt: "Bar chart showing readiness increasing from 76 percent for the model to 96 percent for native render.",
    chartType: "bar",
    categories: ["Model", "Package", "Render"],
    series: [{ name: "Readiness", values: [76, 90, 96], color: "#0F766E" }],
    bbox: [72, 380, 468, 180],
  });
  const statusMark = pdf.addImage({
    id: "verified-status-mark",
    name: "verified-status-mark",
    dataUrl: STATUS_MARK,
    alt: "Verified release evidence status mark.",
    bbox: [72, 610, 28, 28],
  });
  const footer = pdf.addText("Generated and verified with open-office-artifact-tool", {
    id: "verification-footer",
    bbox: [112, 615, 360, 14],
    fontSize: 9,
    color: "#486779",
  });

  const page = pdf.pages[0];
  page.setReadingOrder([
    `${page.id}/text`,
    decisionHeading,
    decision,
    evidence,
    trend,
    statusMark,
    footer,
  ]);
  return pdf;
}

export async function createPdf(outputPath, options = {}) {
  const authored = buildPdf();
  const authoredReport = verifyArtifact(authored);
  assert.equal(authoredReport.ok, true, authoredReport.ndjson || JSON.stringify(authoredReport.issues));

  const firstFile = await PdfFile.exportPdf(authored, {
    title: "Release readiness scorecard",
    language: "en-US",
  });
  const imported = await PdfFile.importPdf(firstFile);
  const evidence = imported.pages.flatMap((page) => page.tables)
    .find((table) => table.name === "readiness-evidence");
  assert.ok(evidence, "The readiness evidence table must survive PDF model round-trip");
  evidence.getCell(3, 2).value = "Verified";

  const finalFile = await PdfFile.exportPdf(imported, {
    title: "Release readiness scorecard",
    language: "en-US",
  });
  const finalPdf = await PdfFile.importPdf(finalFile);
  const verification = verifyArtifact(finalPdf);
  assert.equal(verification.ok, true, verification.ndjson || JSON.stringify(verification.issues));
  assert.equal(finalPdf.pages[0].tables[0].getCell(3, 2).value, "Verified");

  const inspection = finalPdf.inspect({
    kind: "page,text,textItem,readingOrder,table,tableCell,image,chart",
    maxChars: 24_000,
  });
  for (const expected of ["Release readiness scorecard", "Verified", "readiness-trend", "verified-status-mark"]) {
    assert.match(inspection.ndjson, new RegExp(expected));
  }
  const fileInspection = await PdfFile.inspectPdf(finalFile, { maxChars: 24_000 });
  assert.equal(fileInspection.summary.tagged, true);
  assert.equal(fileInspection.summary.tableStructures, 1);
  assert.equal(fileInspection.summary.figures, 2);
  assert.equal(fileInspection.summary.missingFigureAltTexts, 0);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await finalFile.save(outputPath);

  const renderedPages = [];
  if (options.renderDir) {
    const renderDir = path.resolve(options.renderDir);
    await fs.mkdir(renderDir, { recursive: true });
    const renderer = createPopplerRenderer({ dpi: options.dpi ?? 144, timeoutMs: options.timeoutMs ?? 60_000 });
    for (let pageIndex = 0; pageIndex < finalPdf.pages.length; pageIndex += 1) {
      const png = await finalPdf.render({ source: "pdf", format: "png", pageIndex, renderer });
      const pagePath = path.join(renderDir, `page-${pageIndex + 1}.png`);
      await png.save(pagePath);
      renderedPages.push({ page: pageIndex + 1, path: pagePath, bytes: png.bytes.length });
    }
  }

  return { pdf: finalPdf, file: finalFile, inspection, fileInspection, verification, renderedPages };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "release-readiness-scorecard.pdf");
  const renderDir = process.argv[3] ? path.resolve(process.argv[3]) : undefined;
  const result = await createPdf(outputPath, { renderDir });
  console.log(JSON.stringify({
    outputPath,
    bytes: result.file.bytes.length,
    pages: result.pdf.pages.length,
    tagged: result.fileInspection.summary.tagged,
    renderedPages: result.renderedPages.length,
    verified: result.verification.ok,
  }));
}
