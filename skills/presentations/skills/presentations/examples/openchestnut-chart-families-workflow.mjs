import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, Presentation, PresentationFile, renderArtifact } from "open-office-artifact-tool";
import { playwrightRenderer } from "open-office-artifact-tool/renderers/playwright";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function requiredPath(value, name) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${name} must be a non-empty path.`);
  return path.resolve(value);
}

function chartByName(presentation, name) {
  const matches = presentation.slides.getItem(0).charts.items.filter((chart) => chart.name === name);
  if (matches.length !== 1) throw new Error(`Expected exactly one chart named ${JSON.stringify(name)}; found ${matches.length}.`);
  return matches[0];
}

async function chartXmlInventory(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((entry) => /(?:^|\/)charts\/chart\d+\.xml$/i.test(entry)).sort();
  const charts = await Promise.all(paths.map(async (partPath) => ({ partPath, xml: await zip.file(partPath).async("text") })));
  const typeFor = (xml) => ["area", "doughnut", "scatter", "bubble"].find((type) => new RegExp(`<c:${type}Chart>`).test(xml));
  return charts.map(({ partPath, xml }) => ({
    partPath,
    type: typeFor(xml),
    sha256: sha256(Buffer.from(xml)),
    showPercent: /<c:showPercent val="1"\s*\/>/.test(xml),
    numericX: /<c:xVal>[\s\S]*<c:yVal>/.test(xml),
    bubbleSizes: /<c:bubbleSize>/.test(xml),
  }));
}

function createChartFamilyDeck() {
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const slide = presentation.slides.add({ name: "Native chart families" });
  slide.charts.add("area", {
    name: "area-family",
    title: "Regional trajectory",
    position: { left: 40, top: 35, width: 570, height: 300 },
    categories: ["Q1", "Q2", "Q3"],
    series: [{ name: "Revenue", values: [42, 53, 68], fill: "#0EA5E9", line: { fill: "#0369A1", width: 1.5 } }],
    xAxis: { title: "Quarter" },
    yAxis: { title: "Revenue", min: 0, max: 80, majorUnit: 20 },
    legend: false,
  });
  slide.charts.add("doughnut", {
    name: "doughnut-family",
    title: "Regional mix",
    position: { left: 670, top: 35, width: 570, height: 300 },
    categories: ["North", "Central", "South"],
    series: [{ name: "Share", values: [52, 31, 17] }],
    dataLabels: { showCategoryName: true, showPercent: true, position: "outsideEnd" },
    legend: true,
  });
  slide.charts.add("scatter", {
    name: "scatter-family",
    title: "Reach relationship",
    position: { left: 40, top: 370, width: 570, height: 300 },
    series: [{ name: "Portfolio", xValues: [10, 20, 34], values: [35, 68, 84], marker: { symbol: "diamond", size: 8, fill: "#8B5CF6", line: { fill: "#6D28D9", width: 1 } } }],
    xAxis: { title: "Reach", min: 0, max: 40, majorUnit: 10 },
    yAxis: { title: "Return", min: 0, max: 100, majorUnit: 20 },
    legend: false,
  });
  slide.charts.add("bubble", {
    name: "bubble-family",
    title: "Opportunity map",
    position: { left: 670, top: 370, width: 570, height: 300 },
    series: [{ name: "Opportunity", xValues: [10, 20, 34], values: [35, 68, 84], bubbleSizes: [4, 9, 16], fill: "#F97316", line: { fill: "#C2410C", width: 1 } }],
    xAxis: { title: "Reach", min: 0, max: 40, majorUnit: 10 },
    yAxis: { title: "Return", min: 0, max: 100, majorUnit: 20 },
    legend: false,
  });
  return presentation;
}

export async function createAndEditChartFamilyDeck({ outputPath, previewPath, auditPath }) {
  const finalOutput = requiredPath(outputPath, "outputPath");
  const finalPreview = requiredPath(previewPath, "previewPath");
  const finalAudit = requiredPath(auditPath, "auditPath");
  if (new Set([finalOutput, finalPreview, finalAudit]).size !== 3) throw new Error("outputPath, previewPath, and auditPath must be distinct.");
  const temporaryOutput = `${finalOutput}.tmp-${process.pid}-${Date.now()}`;
  const temporaryPreview = `${finalPreview}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAudit = `${finalAudit}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([finalOutput, finalPreview, finalAudit].map((entry) => fs.mkdir(path.dirname(entry), { recursive: true })));
  try {
    const authored = createChartFamilyDeck();
    const authoredVerification = authored.verify({ visualQa: true });
    if (!authoredVerification.ok) throw new Error(`Authored chart deck verification failed: ${authoredVerification.ndjson}`);
    const firstExport = await PresentationFile.exportPptx(authored);
    const firstBytes = new Uint8Array(await firstExport.arrayBuffer());
    const firstInventory = await chartXmlInventory(firstBytes);
    if (JSON.stringify(firstInventory.map((entry) => entry.type)) !== JSON.stringify(["area", "doughnut", "scatter", "bubble"])) throw new Error("First PPTX export did not contain the four canonical native chart families.");
    if (!firstInventory[1].showPercent || !firstInventory[2].numericX || !firstInventory[3].numericX || !firstInventory[3].bubbleSizes) throw new Error("First PPTX export is missing required native chart semantics.");

    const imported = await PresentationFile.importPptx(new FileBlob(firstBytes, { type: PPTX_MIME, name: "chart-families-source.pptx" }));
    chartByName(imported, "area-family").series[0].values[1] = 57;
    chartByName(imported, "doughnut-family").series[0].values = [49, 33, 18];
    chartByName(imported, "scatter-family").series[0].xValues[1] = 22;
    chartByName(imported, "bubble-family").series[0].bubbleSizes[1] = 12;
    const editedExport = await PresentationFile.exportPptx(imported);
    await editedExport.save(temporaryOutput);
    const outputBytes = await fs.readFile(temporaryOutput);
    const outputInventory = await chartXmlInventory(outputBytes);

    const roundTrip = await PresentationFile.importPptx(new FileBlob(outputBytes, { type: PPTX_MIME, name: path.basename(finalOutput) }));
    const chartTypes = roundTrip.slides.getItem(0).charts.items.map((chart) => chart.chartType);
    if (JSON.stringify(chartTypes) !== JSON.stringify(["area", "doughnut", "scatter", "bubble"])) throw new Error("Second import changed chart-family order or type.");
    if (chartByName(roundTrip, "area-family").series[0].values[1] !== 57 || chartByName(roundTrip, "doughnut-family").series[0].values[1] !== 33 || chartByName(roundTrip, "scatter-family").series[0].xValues[1] !== 22 || chartByName(roundTrip, "bubble-family").series[0].bubbleSizes[1] !== 12) throw new Error("Second import lost a requested chart edit.");
    const verification = roundTrip.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Round-trip chart deck verification failed: ${verification.ndjson}`);
    const inspect = roundTrip.inspect({ kind: "slide,chart", maxChars: 20_000 });
    for (const type of chartTypes) if (!inspect.ndjson.includes(`"chartType":"${type}"`)) throw new Error(`Inspect evidence is missing ${type}.`);
    const preview = await renderArtifact(roundTrip, {
      slide: roundTrip.slides.getItem(0),
      format: "png",
      renderer: playwrightRenderer,
      viewport: { width: 1280, height: 720 },
    });
    const previewBytes = new Uint8Array(await preview.arrayBuffer());
    if (previewBytes.byteLength < 1_000) throw new Error("Chart-family model render is unexpectedly empty.");
    await fs.writeFile(temporaryPreview, previewBytes);

    const audit = {
      schema: "open-office-artifact-tool.pptx-audit.v1",
      status: "succeeded",
      provider: { actual: "open-chestnut", silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: { type: "greenfield-native-chart-family-author-edit", chartTypes, editedFields: ["area.values", "doughnut.values", "scatter.xValues", "bubble.bubbleSizes"] },
      source: { kind: "in-memory-presentation", firstExportSha256: sha256(firstBytes), bytes: firstBytes.byteLength },
      output: { path: finalOutput, sha256: sha256(outputBytes), bytes: outputBytes.byteLength },
      preview: { path: finalPreview, sha256: sha256(previewBytes), bytes: previewBytes.byteLength, renderer: "model-svg+playwright" },
      validation: { verify: { ok: true }, inspect: { ok: true, chartCount: chartTypes.length }, package: { ok: true, charts: outputInventory }, reimport: { ok: true, chartTypes } },
      warnings: [],
    };
    await fs.writeFile(temporaryAudit, JSON.stringify(audit, null, 2));
    await fs.rename(temporaryOutput, finalOutput);
    await fs.rename(temporaryPreview, finalPreview);
    await fs.rename(temporaryAudit, finalAudit);
    return { outputPath: finalOutput, previewPath: finalPreview, auditPath: finalAudit, audit };
  } catch (error) {
    await Promise.all([temporaryOutput, temporaryPreview, temporaryAudit].map((entry) => fs.rm(entry, { force: true })));
    throw error;
  }
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const [outputPath = "output/chart-families.pptx", previewPath = "output/chart-families.png", auditPath = "output/chart-families.audit.json"] = process.argv.slice(2);
  const result = await createAndEditChartFamilyDeck({ outputPath, previewPath, auditPath });
  console.log(JSON.stringify({ outputPath: result.outputPath, previewPath: result.previewPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256 }));
}
