import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";

import {
  DocumentFile,
  DocumentModel,
  FileBlob,
  verifyArtifact,
  visualQaArtifact,
} from "open-office-artifact-tool";
import { createLibreOfficeRenderer } from "open-office-artifact-tool/renderers/libreoffice";
import { createPlaywrightRenderer } from "open-office-artifact-tool/renderers/playwright";
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PREVIEW_EXTENSION = { svg: "svg", png: "png", webp: "webp", jpeg: "jpg", jpg: "jpg", pdf: "pdf" };

function xmlEscape(value) {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}

function packageCommentsXml(comments = []) {
  const body = comments.map((comment) => `<w:comment w:id="${xmlEscape(comment.commentId)}" w:author="${xmlEscape(comment.author || "User")}" w:initials="${xmlEscape(comment.initials || "")}"${comment.date ? ` w:date="${xmlEscape(comment.date)}"` : ""}><w:p><w:r><w:t>${xmlEscape(comment.text || "")}</w:t></w:r></w:p></w:comment>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${body}</w:comments>`;
}

function commandExists(command) {
  const result = spawnSync(process.platform === "win32" ? "where" : "which", [command], {
    encoding: "utf8",
    shell: false,
  });
  return result.status === 0;
}

export function nativeDocumentRenderStatus() {
  const commands = { soffice: commandExists("soffice"), pdftoppm: commandExists("pdftoppm"), pdfinfo: commandExists("pdfinfo") };
  return { available: Object.values(commands).every(Boolean), commands };
}

function modelRenderer(format, options = {}) {
  if (format === "svg") return undefined;
  return createPlaywrightRenderer({ viewport: options.viewport, deviceScaleFactor: options.deviceScaleFactor ?? 1, timeout: options.timeout ?? 30_000 });
}

function addFixtureBlock(document, block = {}) {
  const config = { ...block };
  switch (block.kind) {
    case "paragraph": return document.addParagraph(block.text || "", config);
    case "listItem": return document.addListItem(block.text || "", config);
    case "table": return document.addTable(config);
    case "hyperlink": return document.addHyperlink(block.text || "", block.url, config);
    case "field": return document.addField(block.instruction, block.display, config);
    case "citation": return document.addCitation(block.text || "", block.metadata || {}, config);
    case "image": return document.addImage(config);
    case "section": return document.addSection(config);
    case "insertion": return document.addInsertion(block.text || "", config);
    case "deletion": return document.addDeletion(block.text || "", config);
    default: throw new Error(`Unsupported document fixture block kind: ${block.kind}`);
  }
}

export function createDocumentFromFixture(fixture = {}) {
  const document = DocumentModel.create({ name: fixture.name || "Fixture document", blocks: [] });
  if (fixture.designPreset) document.applyDesignPreset(fixture.designPreset, fixture.designOptions || {});
  for (const [id, style] of Object.entries(fixture.styles || {})) document.styles.add(id, style);
  for (const header of fixture.headers || []) document.addHeader(header.text || "", header);
  for (const footer of fixture.footers || []) document.addFooter(footer.text || "", footer);
  const byName = new Map();
  for (const block of fixture.blocks || []) {
    const created = addFixtureBlock(document, block);
    if (block.name) byName.set(block.name, created);
  }
  for (const comment of fixture.comments || []) {
    const target = byName.get(comment.targetName) || comment.targetId;
    assert.ok(target, `Missing document fixture comment target ${comment.targetName || comment.targetId}`);
    document.addComment(target, comment.text || "", comment);
  }
  for (const expected of fixture.expectInspect || []) {
    assert.match(document.inspect({ kind: expected.kind || "document,paragraph,listItem,table,comment,header,footer,hyperlink,field,citation,image,section,change,style", maxChars: 20_000 }).ndjson, new RegExp(expected.pattern));
  }
  return document;
}

function pdfPageCount(pdfPath) {
  const result = spawnSync("pdfinfo", [pdfPath], { encoding: "utf8" });
  if (result.status !== 0) throw new Error(`pdfinfo failed for ${pdfPath}: ${result.stderr || result.stdout}`);
  const pages = Number(/^Pages:\s+(\d+)/m.exec(result.stdout)?.[1]);
  if (!Number.isInteger(pages) || pages < 1) throw new Error(`pdfinfo did not report a valid page count for ${pdfPath}.`);
  return pages;
}

async function optionalBaseline(baselinePath) {
  if (!baselinePath) return undefined;
  try { return await FileBlob.load(baselinePath); } catch (error) { if (error.code === "ENOENT") return undefined; throw error; }
}

async function nativeBaselineFiles(baselineDir) {
  if (!baselineDir) return [];
  try {
    return (await fs.readdir(baselineDir))
      .filter((name) => /^native-page-\d+\.png$/.test(name))
      .map((name) => path.join(baselineDir, name));
  } catch (error) {
    if (error.code === "ENOENT") return [];
    throw error;
  }
}

async function renderNativePages(docxBlob, outputDir, options = {}) {
  const pdf = await createLibreOfficeRenderer({ timeoutMs: options.nativeTimeout ?? 60_000 })({
    input: docxBlob,
    inputType: DOCX_MIME,
    outputType: "application/pdf",
    format: "pdf",
    artifactKind: "document",
  });
  const pdfPath = path.join(outputDir, "native-render.pdf");
  await pdf.save(pdfPath);
  const pageCount = pdfPageCount(pdfPath);
  const pagesDir = path.join(outputDir, "pages");
  await fs.mkdir(pagesDir, { recursive: true });
  const poppler = createPopplerRenderer({ dpi: options.dpi ?? 150, timeoutMs: options.nativeTimeout ?? 60_000 });
  const pages = [];
  const qaLines = [];
  const baselineDir = options.baselineDir;
  if (options.writeBaseline && baselineDir) {
    await fs.mkdir(baselineDir, { recursive: true });
    await Promise.all((await nativeBaselineFiles(baselineDir)).map((filePath) => fs.unlink(filePath)));
  }
  const existingBaselines = options.writeBaseline ? [] : await nativeBaselineFiles(baselineDir);
  for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
    const png = await poppler({ input: pdf, inputType: "application/pdf", outputType: "image/png", format: "png", artifactKind: "document", pageIndex });
    const pagePath = path.join(pagesDir, `page-${pageIndex + 1}.png`);
    await png.save(pagePath);
    const baselinePath = baselineDir ? path.join(baselineDir, `native-page-${pageIndex + 1}.png`) : undefined;
    const baseline = options.writeBaseline ? undefined : await optionalBaseline(baselinePath);
    const qa = await visualQaArtifact({ render: () => png }, { baseline, pixelDiff: Boolean(baseline), minBytes: options.minBytes ?? 100, maxChars: options.maxChars ?? 16_000, pixelThreshold: options.pixelThreshold, diffAlignment: options.diffAlignment, diffPalette: options.diffPalette, pixelRegistration: options.pixelRegistration });
    if (options.writeBaseline && baselinePath) await png.save(baselinePath);
    const diffPath = qa.diffBlob ? path.join(outputDir, "diffs", `native-page-${pageIndex + 1}.png`) : undefined;
    if (diffPath) { await fs.mkdir(path.dirname(diffPath), { recursive: true }); await qa.diffBlob.save(diffPath); }
    qaLines.push(qa.ndjson);
    pages.push({ page: pageIndex + 1, path: pagePath, diffPath, baselinePath, baselineCompared: Boolean(baseline), bytes: png.bytes.length, hash: qa.summary.hash, pixelDiff: qa.summary.pixelDiff, ok: qa.ok });
  }
  const qaPath = path.join(outputDir, "native-visual-qa.ndjson");
  const baselinePageCount = baselineDir && !options.writeBaseline ? existingBaselines.length : undefined;
  const pageCountMatches = baselinePageCount == null || baselinePageCount === 0 || baselinePageCount === pageCount;
  if (!pageCountMatches) qaLines.push(JSON.stringify({ kind: "visualPageCountDiff", artifactKind: "document", severity: "warning", pageCount, baselinePageCount }));
  await fs.writeFile(qaPath, `${qaLines.filter(Boolean).join("\n")}\n`, "utf8");
  return { status: "passed", ok: pageCountMatches && pages.every((page) => page.ok), pdfPath, qaPath, pageCount, baselinePageCount, pageCountMatches, pages };
}

export async function verifyDocumentFile(inputPath, options = {}) {
  const absoluteInput = path.resolve(inputPath);
  const outputDir = path.resolve(options.outputDir || path.join(path.dirname(absoluteInput), `${path.basename(absoluteInput, path.extname(absoluteInput))}-qa`));
  await fs.mkdir(outputDir, { recursive: true });
  const loaded = await FileBlob.load(absoluteInput);
  const docxBlob = new FileBlob(loaded.bytes, { type: DOCX_MIME, name: path.basename(absoluteInput) });
  const document = await DocumentFile.importDocx(docxBlob, { preferNative: options.preferNative === true });
  const maxChars = options.maxChars ?? 20_000;
  const inspect = document.inspect({ kind: options.inspectKind || "document,paragraph,listItem,table,comment,header,footer,hyperlink,field,citation,image,section,change,style,layout", maxChars });
  const packageInspect = await DocumentFile.inspectDocx(docxBlob, { includeText: options.includePackageText === true, maxChars });
  const verify = verifyArtifact(document, { visualQa: true, maxChars });
  const layoutBlob = await document.render({ format: "layout" });
  const previewFormat = String(options.previewFormat || "svg").toLowerCase();
  const previewExtension = PREVIEW_EXTENSION[previewFormat];
  if (!previewExtension) throw new Error(`Unsupported document model preview format: ${previewFormat}`);
  const baselineDir = options.baselineDir ? path.resolve(options.baselineDir) : undefined;
  const modelBaselinePath = baselineDir ? path.join(baselineDir, `model-preview.${previewExtension}`) : undefined;
  const modelBaseline = options.writeBaseline ? undefined : await optionalBaseline(modelBaselinePath);
  const visualQa = await visualQaArtifact(document, {
    format: previewFormat,
    renderer: modelRenderer(previewFormat, options),
    baseline: modelBaseline,
    pixelDiff: Boolean(modelBaseline && ["png", "webp", "jpeg", "jpg"].includes(previewFormat)),
    pixelThreshold: options.pixelThreshold,
    diffAlignment: options.diffAlignment,
    diffPalette: options.diffPalette,
    pixelRegistration: options.pixelRegistration,
    minBytes: options.minBytes ?? 20,
    maxChars,
  });
  if (options.writeBaseline && modelBaselinePath) {
    await fs.mkdir(baselineDir, { recursive: true });
    await visualQa.blob.save(modelBaselinePath);
  }
  const paths = {
    inspect: path.join(outputDir, "inspect.ndjson"),
    packageInspect: path.join(outputDir, "package-inspect.ndjson"),
    verify: path.join(outputDir, "verify.ndjson"),
    layout: path.join(outputDir, "layout.json"),
    visualQa: path.join(outputDir, "model-visual-qa.ndjson"),
    preview: path.join(outputDir, `model-preview.${previewExtension}`),
    summary: path.join(outputDir, "summary.json"),
  };
  if (visualQa.diffBlob) paths.diff = path.join(outputDir, "model-diff.png");
  const verifyNdjson = verify.ndjson || JSON.stringify({ kind: "verificationSummary", artifactKind: "document", ok: verify.ok, issues: verify.issues?.length || 0 });
  await Promise.all([
    fs.writeFile(paths.inspect, inspect.ndjson, "utf8"),
    fs.writeFile(paths.packageInspect, packageInspect.ndjson, "utf8"),
    fs.writeFile(paths.verify, `${verifyNdjson}\n`, "utf8"),
    fs.writeFile(paths.layout, await layoutBlob.text(), "utf8"),
    fs.writeFile(paths.visualQa, visualQa.ndjson, "utf8"),
    visualQa.blob.save(paths.preview),
    ...(paths.diff ? [visualQa.diffBlob.save(paths.diff)] : []),
  ]);

  const requestedNative = String(options.nativeRender ?? "auto").toLowerCase();
  const nativeStatus = nativeDocumentRenderStatus();
  let nativeRender = { status: "skipped", reason: "native render disabled" };
  if (requestedNative !== "off" && requestedNative !== "false") {
    if (nativeStatus.available) nativeRender = await renderNativePages(docxBlob, outputDir, { ...options, baselineDir });
    else if (requestedNative === "required" || requestedNative === "true") throw new Error(`Native document render requires soffice, pdftoppm, and pdfinfo: ${JSON.stringify(nativeStatus.commands)}`);
    else nativeRender = { status: "skipped", reason: "native render commands unavailable", commands: nativeStatus.commands };
  }
  const summary = {
    input: absoluteInput,
    outputDir,
    packageOk: packageInspect.ok,
    previewFormat,
    baselineDir,
    writeBaseline: Boolean(options.writeBaseline),
    modelBaselinePath,
    modelBaselineCompared: Boolean(modelBaseline),
    modelPixelDiff: visualQa.summary.pixelDiff,
    verifyOk: verify.ok,
    visualQaOk: visualQa.ok,
    renderHash: visualQa.summary.hash,
    nativeRender,
    files: paths,
  };
  await fs.writeFile(paths.summary, `${JSON.stringify(summary, null, 2)}\n`, "utf8");
  if (options.failOnIssues !== false && (!packageInspect.ok || !verify.ok || !visualQa.ok || (nativeRender.status === "passed" && nativeRender.ok === false))) {
    throw new Error(`Document QA failed: package=${packageInspect.ok}, semantic=${verify.ok}, modelVisual=${visualQa.ok}, native=${nativeRender.status}. See ${outputDir}`);
  }
  return { document, inspect, packageInspect, verify, visualQa, layoutBlob, summary };
}

export async function runDocumentFixture(fixturePath, options = {}) {
  const absoluteFixture = path.resolve(fixturePath);
  const fixture = JSON.parse(await fs.readFile(absoluteFixture, "utf8"));
  const outputDir = path.resolve(options.outputDir || path.join("tmp", "document-skill", fixture.name || "fixture"));
  await fs.mkdir(outputDir, { recursive: true });
  const document = createDocumentFromFixture(fixture);
  const docxPath = path.join(outputDir, fixture.outputName || `${fixture.name || "document"}.docx`);
  let docx = await DocumentFile.exportDocx(document);
  if (fixture.packageComments?.length) {
    const partPath = fixture.packageCommentsPart || "word/review/fixture-comments.xml";
    docx = await DocumentFile.patchDocx(docx, [{
      path: partPath,
      xml: packageCommentsXml(fixture.packageComments),
      recipe: {
        kind: "comments",
        source: "word/document.xml",
        id: fixture.packageCommentsRelationshipId,
        sourceReference: { anchors: fixture.packageComments.map(({ commentId, target }) => ({ commentId, target })) },
      },
    }]);
  }
  await docx.save(docxPath);
  const qa = await verifyDocumentFile(docxPath, {
    outputDir: path.join(outputDir, "qa"),
    previewFormat: options.previewFormat || fixture.qa?.previewFormat || "svg",
    nativeRender: options.nativeRender ?? fixture.qa?.nativeRender ?? "auto",
    baselineDir: options.baselineDir,
    writeBaseline: options.writeBaseline,
    pixelThreshold: options.pixelThreshold,
    diffAlignment: options.diffAlignment,
    diffPalette: options.diffPalette,
    pixelRegistration: options.pixelRegistration,
    inspectKind: fixture.qa?.inspectKind,
    maxChars: fixture.qa?.maxChars,
    preferNative: Boolean(fixture.packageComments?.length),
  });
  return { fixture, docxPath, qa };
}
