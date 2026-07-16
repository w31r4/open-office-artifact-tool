import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { PdfArtifact, PdfFile, verifyArtifact } from "open-office-artifact-tool";
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";

const PAGE_COUNT = 6;
const CJK_FONT_CANDIDATES = [
  process.env.OPEN_OFFICE_CJK_FONT,
  "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
  "/Library/Fonts/Arial Unicode.ttf",
  "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttf",
  "/usr/share/fonts/truetype/noto/NotoSansSC-Regular.ttf",
].filter(Boolean);

async function existingFile(candidates) {
  for (const candidate of candidates) {
    const absolute = path.resolve(candidate);
    if (await fs.access(absolute).then(() => true, () => false)) return absolute;
  }
  return undefined;
}

async function fileEvidence(filePath) {
  const absolute = path.resolve(filePath);
  const bytes = await fs.readFile(absolute);
  return { path: absolute, bytes: bytes.byteLength, sha256: createHash("sha256").update(bytes).digest("hex") };
}

async function packageVersion() {
  try {
    const entry = fileURLToPath(import.meta.resolve("open-office-artifact-tool"));
    const packageJson = JSON.parse(await fs.readFile(path.resolve(path.dirname(entry), "../package.json"), "utf8"));
    return String(packageJson.version || "unknown");
  } catch {
    return "unknown";
  }
}

function addHeading(page, id, text, level, top) {
  return page.addText(text, { id, bbox: [72, top, 468, level === 2 ? 28 : 22], fontSize: level === 2 ? 20 : 15, bold: true, color: level === 2 ? "#123B5D" : "#28536B", headingLevel: level });
}

function addParagraph(page, id, text, top, width = 468) {
  return page.addText(text, { id, bbox: [72, top, width, 20], fontSize: 11, color: "#172033" });
}

function addPageArtifacts(page, pageNumber, title) {
  page.addText(title, { id: `running-header-${pageNumber}`, bbox: [72, 28, 380, 12], fontSize: 9, color: "#486779", artifact: true });
  page.addText(`第 ${pageNumber} 页，共 ${PAGE_COUNT} 页`, { id: `running-footer-${pageNumber}`, bbox: [72, 754, 180, 12], fontSize: 9, color: "#486779", artifact: true });
}

export function buildAccessibleBoardReport(data) {
  assert.equal(typeof data?.title, "string", "report data needs title");
  assert.equal(typeof data?.language, "string", "report data needs language");
  assert.ok(Array.isArray(data.sections) && data.sections.length >= 2, "report data needs summary and risk sections");
  const metrics = data.sections[0]?.metrics;
  const risks = data.sections[1]?.rows;
  assert.ok(Array.isArray(metrics) && metrics.length >= 2, "summary metrics are required");
  assert.ok(Array.isArray(risks) && risks.length >= 2, "at least two risk rows are required for the cross-page table");
  assert.ok(String(data.figureAlt || "").trim(), "figureAlt is required");

  const pdf = PdfArtifact.create({
    metadata: { title: data.title, language: data.language },
    pages: Array.from({ length: PAGE_COUNT }, (_, index) => ({ id: `board-page-${index + 1}`, width: 612, height: 792, text: index === 0 ? data.title : "" })),
  });
  pdf.pages.forEach((page, index) => addPageArtifacts(page, index + 1, data.title));

  const [cover, summary, riskOne, riskTwo, validation, conclusion] = pdf.pages;
  const subtitle = addParagraph(cover, "cover-subtitle", "董事会可访问性与交付就绪评估", 150);
  const purpose = addParagraph(cover, "cover-purpose", "本报告将语义模型、机器规则和人工判断分开记录。", 190);
  cover.setReadingOrder([`${cover.id}/text`, subtitle, purpose]);

  const summaryH2 = addHeading(summary, "summary-h2", data.sections[0].heading || "执行摘要", 2, 86);
  const summaryH3 = addHeading(summary, "summary-h3", "格式通过率", 3, 132);
  const categories = metrics.slice(1).map((row) => String(row[0]));
  const values = metrics.slice(1).map((row) => Number(row[1]) * 100);
  const chart = summary.addChart({
    id: "format-pass-rate-chart",
    title: "四种文件格式通过率",
    alt: data.figureAlt,
    chartType: "bar",
    categories,
    series: [{ name: "通过率（%）", values, color: "#0F766E" }],
    bbox: [72, 190, 468, 260],
  });
  const chartNote = addParagraph(summary, "summary-chart-note", "图表替代文本来自输入数据，并作为 Figure 语义输出。", 480);
  summary.setReadingOrder([summaryH2, summaryH3, chart, chartNote]);

  const riskH2 = addHeading(riskOne, "risks-h2", data.sections[1].heading || "风险与缓解", 2, 86);
  const riskH3 = addHeading(riskOne, "risks-h3", "优先风险清单", 3, 132);
  const riskIntro = addParagraph(riskOne, "risks-intro", "下表是一个逻辑表格，跨两页保持同一 semanticId。", 170);
  const splitAt = Math.max(1, Math.ceil(risks.length / 2));
  const tableHeader = ["风险", "级别", "缓解措施"];
  const riskPartOne = riskOne.addTable({
    id: "risk-register-part-1",
    semanticId: "risk-register",
    name: "风险登记表",
    values: [tableHeader, ...risks.slice(0, splitAt)],
    bbox: [72, 540, 468, Math.max(72, (splitAt + 1) * 30)],
  });
  riskOne.setReadingOrder([riskH2, riskH3, riskIntro, riskPartOne]);

  const riskPartTwo = riskTwo.addTable({
    id: "risk-register-part-2",
    semanticId: "risk-register",
    name: "风险登记表（续）",
    values: [tableHeader, ...risks.slice(splitAt)],
    bbox: [72, 72, 468, Math.max(72, (risks.length - splitAt + 1) * 30)],
  });
  const mitigationH3 = addHeading(riskTwo, "mitigation-h3", "缓解责任", 3, 260);
  const mitigationText = addParagraph(riskTwo, "mitigation-text", "高风险项必须 fail closed；视觉漂移通过最终逐页渲染复核。", 306);
  riskTwo.setReadingOrder([riskPartTwo, mitigationH3, mitigationText]);

  const validationH2 = addHeading(validation, "validation-h2", "验证分层", 2, 86);
  const modeledH3 = addHeading(validation, "modeled-h3", "Modeled verify", 3, 142);
  const modeledText = addParagraph(validation, "modeled-text", "验证本项目建模的标题、表格、Figure、链接、artifact 与阅读顺序不变量。", 182);
  const machineH3 = addHeading(validation, "machine-h3", "veraPDF 机器检查", 3, 236);
  const machineText = addParagraph(validation, "machine-text", "若环境具备 veraPDF，则单独运行并保留版本、退出状态和报告；它不等于人工认证。", 276);
  const humanH3 = addHeading(validation, "human-h3", "PDF/UA 人工判断", 3, 330);
  const humanText = addParagraph(validation, "human-text", "阅读意图、替代文本质量、链接目的、对比度和辅助技术体验仍需人工确认。", 370);
  validation.setReadingOrder([validationH2, modeledH3, modeledText, machineH3, machineText, humanH3, humanText]);

  const conclusionH2 = addHeading(conclusion, "conclusion-h2", "董事会结论", 2, 86);
  const conclusionH3 = addHeading(conclusion, "conclusion-h3", "交付边界", 3, 142);
  const conclusionText = addParagraph(conclusion, "conclusion-text", "本文件证明的是可审计的 tagged authoring 工作流，不声明完整 PDF/UA 认证。", 184);
  const guidance = conclusion.addLink({ id: "wai-guidance-link", text: "W3C 无障碍标准与指南", url: "https://www.w3.org/WAI/standards-guidelines/", bbox: [72, 244, 260, 20] });
  conclusion.setReadingOrder([conclusionH2, conclusionH3, conclusionText, guidance]);

  return pdf;
}

function runVeraPdf(outputPath) {
  const version = spawnSync("verapdf", ["--version"], { encoding: "utf8" });
  if (version.error?.code === "ENOENT") return { available: false, status: "not-run", claim: "No veraPDF machine validation was performed." };
  if (version.status !== 0) return { available: false, status: "probe-failed", exitCode: version.status, stderr: version.stderr.trim(), claim: "No veraPDF machine validation was performed." };
  const validation = spawnSync("verapdf", ["--format", "json", outputPath], { encoding: "utf8", maxBuffer: 16 * 1024 * 1024 });
  return {
    available: true,
    version: (version.stdout || version.stderr).trim(),
    status: validation.status === 0 ? "completed" : "completed-with-findings",
    exitCode: validation.status,
    report: validation.stdout.trim(),
    stderr: validation.stderr.trim(),
    claim: "veraPDF machine-rule evidence only; not a complete PDF/UA certification.",
  };
}

export async function createAccessibleBoardReport(inputPath, outputPath, auditPath, options = {}) {
  const input = path.resolve(inputPath);
  const output = path.resolve(outputPath);
  const auditOutput = path.resolve(auditPath);
  const data = JSON.parse(await fs.readFile(input, "utf8"));
  const font = options.font || await existingFile(CJK_FONT_CANDIDATES);
  if (!font) throw new Error("No standalone CJK TrueType font found. Set OPEN_OFFICE_CJK_FONT to a .ttf file; TTC/OTF inputs are not supported by this writer.");

  const pdf = buildAccessibleBoardReport(data);
  const modeled = verifyArtifact(pdf, { maxChars: 32_000 });
  assert.equal(modeled.ok, true, modeled.ndjson || JSON.stringify(modeled.issues));
  const file = await PdfFile.exportPdf(pdf, { title: data.title, language: data.language, font, maxFontBytes: 32 * 1024 * 1024 });
  await fs.mkdir(path.dirname(output), { recursive: true });
  await file.save(output);
  const fileInspection = await PdfFile.inspectPdf(file, { maxChars: 32_000 });
  assert.equal(fileInspection.summary.pages, PAGE_COUNT);
  assert.equal(fileInspection.summary.tagged, true);
  assert.equal(fileInspection.summary.language, data.language);
  assert.equal(fileInspection.summary.tableStructures, 1);
  assert.ok(fileInspection.summary.headingLevels.H1 >= 1 && fileInspection.summary.headingLevels.H2 >= 1 && fileInspection.summary.headingLevels.H3 >= 1);
  assert.ok(fileInspection.summary.figures >= 1 && fileInspection.summary.missingFigureAltTexts === 0);
  assert.ok(fileInspection.summary.linkAnnotations >= 1 && fileInspection.summary.linkStructParents >= 1);
  assert.ok(fileInspection.summary.artifacts >= PAGE_COUNT * 2);

  const renderDir = path.resolve(options.renderDir || path.join(path.dirname(auditOutput), "../tmp/accessibility-render"));
  await fs.mkdir(renderDir, { recursive: true });
  const renderer = createPopplerRenderer({ dpi: options.dpi ?? 144, timeoutMs: options.timeoutMs ?? 60_000 });
  const renderedPages = [];
  for (let pageIndex = 0; pageIndex < PAGE_COUNT; pageIndex += 1) {
    const png = await renderer({ input: file, inputType: "application/pdf", outputType: "image/png", format: "png", artifactKind: "pdf", options: { pageIndex } });
    assert.ok(png.bytes.length > 100, `Poppler page ${pageIndex + 1} render is unexpectedly empty`);
    const pagePath = path.join(renderDir, `page-${pageIndex + 1}.png`);
    await png.save(pagePath);
    renderedPages.push({ page: pageIndex + 1, path: pagePath, bytes: png.bytes.length });
  }

  const audit = {
    schema: "open-office-artifact-tool.pdf-audit.v1",
    status: "succeeded",
    source: await fileEvidence(input),
    output: await fileEvidence(output),
    provider: { actual: "artifact-tool", version: await packageVersion(), silentFallback: false },
    savePolicy: { strategy: "rewrite" },
    preflight: { probeCompleted: true, planCompleted: true, font, poppler: "required" },
    operation: { type: "create-accessible-report", pages: PAGE_COUNT },
    validation: {
      modeledVerify: { status: "passed", scope: "PdfArtifact modeled invariants", issues: modeled.issues.length },
      fileInspect: { status: "passed", summary: fileInspection.summary },
      poppler: { status: "passed", renderer: "pdftoppm", pages: renderedPages },
      veraPdfMachine: runVeraPdf(output),
      humanPdfUa: {
        status: "required",
        claim: "No complete PDF/UA certification is claimed.",
        checkpoints: ["author intent and heading hierarchy", "reading order", "table interpretation", "alternative-text quality", "link purpose", "color contrast", "assistive-technology behavior"],
      },
    },
  };
  await fs.mkdir(path.dirname(auditOutput), { recursive: true });
  await fs.writeFile(auditOutput, `${JSON.stringify(audit, null, 2)}\n`, "utf8");
  return { pdf, file, audit, fileInspection, modeled, renderedPages };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const inputPath = path.resolve(process.argv[2] || "inputs/report-data.json");
  const outputPath = path.resolve(process.argv[3] || "outputs/readiness-report.pdf");
  const auditPath = path.resolve(process.argv[4] || "outputs/audit.json");
  const result = await createAccessibleBoardReport(inputPath, outputPath, auditPath);
  console.log(JSON.stringify({ outputPath, auditPath, bytes: result.file.bytes.length, pages: result.pdf.pages.length, tagged: result.fileInspection.summary.tagged, modeledVerify: result.modeled.ok, veraPdf: result.audit.validation.veraPdfMachine.status }));
}
