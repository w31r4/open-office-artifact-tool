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
import {
  loadVisualBaseline,
  prepareNumberedVisualBaselines,
  runPngVisualQa,
  visualBaselineCountResult,
} from "../../shared/visual-baselines.mjs";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PREVIEW_EXTENSION = { svg: "svg", png: "png", webp: "webp", jpeg: "jpg", jpg: "jpg", pdf: "pdf" };

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
    case "toc": return document.addTableOfContents(config);
    case "citation": return document.addCitation(block.text || "", block.metadata || {}, config);
    case "image": return document.addImage(config);
    case "section": return document.addSection(config);
    case "insertion": return document.addInsertion(block.text || "", config);
    case "deletion": return document.addDeletion(block.text || "", config);
    default: throw new Error(`Unsupported document fixture block kind: ${block.kind}`);
  }
}

export function createDocumentFromFixture(fixture = {}) {
  const settings = fixture.settings || {};
  const unsupportedSettings = Object.keys(settings).filter((key) => !new Set(["evenAndOddHeaders", "updateFields", "trackRevisions"]).has(key));
  if (unsupportedSettings.length) {
    throw new Error(`Document fixture settings are limited to evenAndOddHeaders, updateFields, and trackRevisions; imported ${unsupportedSettings.join(", ")} semantics are read-only.`);
  }
  const document = DocumentModel.create({
    name: fixture.name || "Fixture document",
    defaultRunStyle: fixture.defaultRunStyle || {},
    settings,
    bibliography: fixture.bibliography || {},
    blocks: [],
  });
  for (const [id, style] of Object.entries(fixture.styles || {})) document.styles.add(id, style);
  for (const source of fixture.bibliographySources || []) document.addBibliographySource(source);
  const byName = new Map();
  for (const block of fixture.blocks || []) {
    const created = addFixtureBlock(document, block);
    if (block.name) byName.set(block.name, created);
  }
  for (const settings of fixture.sectionSettings || []) document.setSectionSettings(settings.sectionIndex, settings);
  for (const header of fixture.headers || []) document.addHeader(header.text || "", header);
  for (const footer of fixture.footers || []) document.addFooter(footer.text || "", footer);
  for (const comment of fixture.comments || []) {
    const target = byName.get(comment.targetName) || comment.targetId;
    assert.ok(target, `Missing document fixture comment target ${comment.targetName || comment.targetId}`);
    document.addComment(target, comment.text || "", comment);
  }
  for (const expected of fixture.expectInspect || []) {
    assert.match(document.inspect({ kind: expected.kind || "document,paragraph,listItem,table,comment,header,footer,hyperlink,field,citation,bibliographySource,image,section,style", maxChars: 20_000 }).ndjson, new RegExp(expected.pattern));
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
  const baselineSet = await prepareNumberedVisualBaselines(baselineDir, "native-page", options);
  for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
    const png = await poppler({ input: pdf, inputType: "application/pdf", outputType: "image/png", format: "png", artifactKind: "document", pageIndex });
    const pagePath = path.join(pagesDir, `page-${pageIndex + 1}.png`);
    await png.save(pagePath);
    const baselinePath = baselineDir ? path.join(baselineDir, `native-page-${pageIndex + 1}.png`) : undefined;
    const diffPath = path.join(outputDir, "diffs", `native-page-${pageIndex + 1}.png`);
    const qa = await runPngVisualQa({ render: () => png }, { baselinePath, diffPath, writeBaseline: options.writeBaseline, minBytes: options.minBytes ?? 100, maxChars: options.maxChars ?? 16_000, pixelThreshold: options.pixelThreshold, diffAlignment: options.diffAlignment, diffPalette: options.diffPalette, pixelRegistration: options.pixelRegistration });
    qaLines.push(qa.ndjson);
    pages.push({ page: pageIndex + 1, path: pagePath, diffPath: qa.diffPath, baselinePath, baselineCompared: Boolean(qa.summary.baselineHash), bytes: png.bytes.length, hash: qa.summary.hash, pixelDiff: qa.summary.pixelDiff, ok: qa.ok });
  }
  const qaPath = path.join(outputDir, "native-visual-qa.ndjson");
  const { baselinePageCount, pageCountMatches, issue } = visualBaselineCountResult(baselineSet, pageCount, { artifactKind: "document", baselineKind: "native" });
  if (issue) qaLines.push(issue);
  await fs.writeFile(qaPath, `${qaLines.filter(Boolean).join("\n")}\n`, "utf8");
  return { status: "passed", ok: pageCountMatches && pages.every((page) => page.ok), pdfPath, qaPath, pageCount, baselinePageCount, pageCountMatches, pages };
}

export async function verifyDocumentFile(inputPath, options = {}) {
  const absoluteInput = path.resolve(inputPath);
  const outputDir = path.resolve(options.outputDir || path.join(path.dirname(absoluteInput), `${path.basename(absoluteInput, path.extname(absoluteInput))}-qa`));
  await fs.mkdir(outputDir, { recursive: true });
  const loaded = await FileBlob.load(absoluteInput);
  const docxBlob = new FileBlob(loaded.bytes, { type: DOCX_MIME, name: path.basename(absoluteInput) });
  const document = await DocumentFile.importDocx(docxBlob);
  const maxChars = options.maxChars ?? 20_000;
  const inspect = document.inspect({ kind: options.inspectKind || "document,paragraph,listItem,table,comment,header,footer,hyperlink,field,citation,bibliographySource,image,section,style,layout", maxChars });
  const packageInspect = await DocumentFile.inspectDocx(docxBlob, { includeText: options.includePackageText === true, maxChars });
  const verify = verifyArtifact(document, { visualQa: true, maxChars });
  const layoutBlob = await document.render({ format: "layout" });
  const previewFormat = String(options.previewFormat || "svg").toLowerCase();
  const previewExtension = PREVIEW_EXTENSION[previewFormat];
  if (!previewExtension) throw new Error(`Unsupported document model preview format: ${previewFormat}`);
  const baselineDir = options.baselineDir ? path.resolve(options.baselineDir) : undefined;
  const modelBaselinePath = baselineDir ? path.join(baselineDir, `model-preview.${previewExtension}`) : undefined;
  const modelBaseline = await loadVisualBaseline(modelBaselinePath, options);
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
  const imported = await DocumentFile.importDocx(docx);
    for (const edit of fixture.edits || []) {
      if (edit.kind === "materializeFields") {
        const result = imported.materializeFields(edit.options || {});
        if (Object.prototype.hasOwnProperty.call(edit, "expectUpdated")) assert.equal(result.updated, edit.expectUpdated);
        if (Object.prototype.hasOwnProperty.call(edit, "expectSeqFields")) assert.equal(result.seqFields, edit.expectSeqFields);
        if (Object.prototype.hasOwnProperty.call(edit, "expectRefFields")) assert.equal(result.refFields, edit.expectRefFields);
        if (Object.prototype.hasOwnProperty.call(edit, "expectSkippedPageReferences")) assert.equal(result.skippedPageReferences, edit.expectSkippedPageReferences);
        continue;
      }
      if (edit.kind === "paragraph") {
        const paragraph = imported.blocks.find((block) => block.kind === "paragraph" && (!edit.matchText || block.text === edit.matchText));
        assert.ok(paragraph, `Missing source-bound paragraph fixture target ${edit.matchText || "(unspecified)"}.`);
        if (Object.prototype.hasOwnProperty.call(edit, "text")) {
          paragraph.text = String(edit.text);
          if (paragraph.runs.length === 1) paragraph.runs[0].text = paragraph.text;
        }
        if (edit.paragraphFormat) paragraph.paragraphFormat = { ...(paragraph.paragraphFormat || {}), ...structuredClone(edit.paragraphFormat) };
        if (edit.runStyle) {
          assert.equal(paragraph.runs.length, 1, "Run-style fixture edits require one modeled run.");
          paragraph.runs[0].style = { ...paragraph.runs[0].style, ...structuredClone(edit.runStyle) };
        }
        continue;
      }
      if (edit.kind === "paragraphRuns") {
        const paragraph = imported.blocks.find((block) => block.kind === "paragraph" && (!edit.matchText || block.text === edit.matchText));
        assert.ok(paragraph, `Missing source-bound paragraph-runs fixture target ${edit.matchText || "(unspecified)"}.`);
        for (const change of edit.runs || []) {
          const run = paragraph.runs[Number(change.index)];
          assert.ok(run, `Missing source-bound paragraph run ${change.index}.`);
          if (Object.prototype.hasOwnProperty.call(change, "expectInstruction")) assert.equal(run.inlineField?.instruction, change.expectInstruction);
          if (Object.prototype.hasOwnProperty.call(change, "expectBookmarkName")) assert.equal(run.inlineField?.bookmarkName, change.expectBookmarkName);
          if (Object.prototype.hasOwnProperty.call(change, "text")) run.text = String(change.text);
        }
        paragraph.text = paragraph.runs.map((run) => String(run.text || "")).join("");
        continue;
      }
      if (edit.kind === "contentControls") {
        const result = imported.fillContentControls(edit.values || {}, { strict: edit.strict !== false });
        if (Object.prototype.hasOwnProperty.call(edit, "expectUpdated")) assert.equal(result.updated, edit.expectUpdated);
        continue;
      }
      if (edit.kind === "checkboxContentControls") {
        const result = imported.setCheckboxContentControls(edit.values || {}, { strict: edit.strict !== false });
        if (Object.prototype.hasOwnProperty.call(edit, "expectUpdated")) assert.equal(result.updated, edit.expectUpdated);
        continue;
      }
      if (edit.kind === "dropdownContentControls") {
        const result = imported.setDropdownContentControls(edit.values || {}, { strict: edit.strict !== false });
        if (Object.prototype.hasOwnProperty.call(edit, "expectUpdated")) assert.equal(result.updated, edit.expectUpdated);
        continue;
      }
      if (edit.kind === "image") {
        const image = imported.blocks.find((block) => block.kind === "image" && (!edit.matchAlt || block.alt === edit.matchAlt));
        assert.ok(image, `Missing source-bound image fixture target ${edit.matchAlt || "(unspecified)"}.`);
        for (const field of ["alt", "widthPx", "heightPx", "dataUrl"]) {
          if (Object.prototype.hasOwnProperty.call(edit, field)) image[field] = edit[field];
        }
        continue;
      }
      if (edit.kind === "section") {
        const sections = imported.blocks.filter((block) => block.kind === "section");
        const section = sections[Number(edit.index || 0)];
        assert.ok(section, `Missing source-bound section fixture index ${edit.index || 0}.`);
        for (const field of ["breakType", "orientation"]) {
          if (Object.prototype.hasOwnProperty.call(edit, field)) section[field] = edit[field];
        }
        if (edit.pageSize) section.pageSize = { ...section.pageSize, ...structuredClone(edit.pageSize) };
        if (edit.margins) section.margins = { ...section.margins, ...structuredClone(edit.margins) };
        continue;
      }
      if (edit.kind === "hyperlink") {
        const hyperlink = imported.blocks.find((block) => block.kind === "hyperlink" && (!edit.matchText || block.text === edit.matchText));
        assert.ok(hyperlink, `Missing source-bound hyperlink fixture target ${edit.matchText || "(unspecified)"}.`);
        if (Object.prototype.hasOwnProperty.call(edit, "text")) hyperlink.text = String(edit.text);
        if (Object.prototype.hasOwnProperty.call(edit, "url")) {
          hyperlink.url = String(edit.url || "");
          hyperlink.anchor = undefined;
        }
        if (Object.prototype.hasOwnProperty.call(edit, "anchor")) {
          hyperlink.anchor = String(edit.anchor || "") || undefined;
          hyperlink.url = "";
        }
        if (Object.prototype.hasOwnProperty.call(edit, "tooltip")) hyperlink.tooltip = edit.tooltip;
        if (Object.prototype.hasOwnProperty.call(edit, "history")) hyperlink.history = edit.history !== false;
        continue;
      }
      if (edit.kind === "field") {
        const field = imported.blocks.find((block) => block.kind === "field" && (!edit.matchInstruction || block.instruction === edit.matchInstruction));
        assert.ok(field, `Missing source-bound field fixture target ${edit.matchInstruction || "(unspecified)"}.`);
        if (Object.prototype.hasOwnProperty.call(edit, "instruction")) field.instruction = String(edit.instruction);
        if (Object.prototype.hasOwnProperty.call(edit, "display")) field.display = String(edit.display);
        continue;
      }
      if (edit.kind === "settings") {
        imported.setSettings(edit.values || {});
        continue;
      }
      if (edit.kind === "citation") {
        const citation = imported.blocks.find((block) => block.kind === "citation" && (!edit.matchTag || block.metadata?.tag === edit.matchTag));
        assert.ok(citation, `Missing source-bound citation fixture target ${edit.matchTag || "(unspecified)"}.`);
        if (Object.prototype.hasOwnProperty.call(edit, "text")) citation.text = String(edit.text);
        continue;
      }
      if (edit.kind === "bibliographySource") {
        const source = imported.bibliographySources.find((item) => !edit.matchTag || item.tag === edit.matchTag);
        assert.ok(source, `Missing source-bound bibliography fixture target ${edit.matchTag || "(unspecified)"}.`);
        for (const [field, value] of Object.entries(edit.fields || {})) source[field] = structuredClone(value);
        if (Array.isArray(edit.authors)) source.authors = structuredClone(edit.authors);
        if (Object.prototype.hasOwnProperty.call(edit, "corporateAuthor")) source.corporateAuthor = String(edit.corporateAuthor || "") || undefined;
        continue;
      }
      if (edit.kind === "tableCell") {
        const table = imported.blocks.find((block) => block.kind === "table" && (!edit.matchText || block.values?.some((row) => row.some((value) => String(value) === edit.matchText))));
        assert.ok(table, `Missing source-bound table fixture target ${edit.matchText || "(unspecified)"}.`);
        assert.ok(Number.isInteger(edit.row) && edit.row >= 0 && edit.row < table.values.length, `Invalid source-bound table row ${edit.row}.`);
        assert.ok(Number.isInteger(edit.column) && edit.column >= 0 && edit.column < table.values[edit.row].length, `Invalid source-bound table column ${edit.column}.`);
        const cell = table.getCell(edit.row, edit.column);
        if (Object.prototype.hasOwnProperty.call(edit, "expectColumnSpan")) assert.equal(cell.columnSpan, edit.expectColumnSpan, `Unexpected OpenChestnut column span for ${edit.matchText || cell.id}.`);
        if (Object.prototype.hasOwnProperty.call(edit, "expectRowSpan")) assert.equal(cell.rowSpan, edit.expectRowSpan, `Unexpected OpenChestnut row span for ${edit.matchText || cell.id}.`);
        if (Object.prototype.hasOwnProperty.call(edit, "expectVerticalMerge")) assert.equal(cell.verticalMerge, edit.expectVerticalMerge, `Unexpected OpenChestnut vertical merge for ${edit.matchText || cell.id}.`);
        if (Object.prototype.hasOwnProperty.call(edit, "expectEditable")) assert.equal(cell.editable, edit.expectEditable, `Unexpected OpenChestnut editability for ${edit.matchText || cell.id}.`);
        table.values[edit.row][edit.column] = String(edit.value ?? "");
        continue;
      }
      if (edit.kind === "tableFormatting") {
        const table = imported.blocks.find((block) => block.kind === "table" && (!edit.matchText || block.values?.some((row) => row.some((value) => String(value) === edit.matchText))));
        assert.ok(table, `Missing source-bound table-formatting fixture target ${edit.matchText || "(unspecified)"}.`);
        for (const field of ["widthDxa", "indentDxa", "columnWidthsDxa", "cellMarginsDxa", "borderColor", "borderSize", "headerFill"]) {
          if (Object.prototype.hasOwnProperty.call(edit, field)) table[field] = structuredClone(edit[field]);
        }
        continue;
      }
      if (edit.kind === "listItem") {
        const listItem = imported.blocks.find((block) => block.kind === "listItem" && (!edit.matchText || block.text === edit.matchText));
        assert.ok(listItem, `Missing source-bound list-item fixture target ${edit.matchText || "(unspecified)"}.`);
        if (Object.prototype.hasOwnProperty.call(edit, "expectNumberingStyleId")) {
          assert.equal(listItem.numberingStyleId, edit.expectNumberingStyleId, `Unexpected OpenChestnut numbering style for ${edit.matchText || listItem.id}.`);
        }
        if (Object.prototype.hasOwnProperty.call(edit, "expectLevel")) {
          assert.equal(listItem.level, edit.expectLevel, `Unexpected OpenChestnut numbering level for ${edit.matchText || listItem.id}.`);
        }
        if (Object.prototype.hasOwnProperty.call(edit, "text")) listItem.text = String(edit.text);
        if (Object.prototype.hasOwnProperty.call(edit, "numberFormat")) listItem.numberFormat = String(edit.numberFormat);
        if (Object.prototype.hasOwnProperty.call(edit, "start")) listItem.start = Number(edit.start);
        if (Object.prototype.hasOwnProperty.call(edit, "levelText")) listItem.levelText = String(edit.levelText);
        if (Object.prototype.hasOwnProperty.call(edit, "listType")) listItem.listType = String(edit.listType);
        continue;
      }
      if (edit.kind === "comment") {
        const comment = imported.comments.find((item) =>
          (!edit.matchText || item.text === edit.matchText) &&
          (!edit.matchAuthor || item.author === edit.matchAuthor));
        assert.ok(comment, `Missing source-bound classic comment fixture target ${edit.matchText || edit.matchAuthor || "(unspecified)"}.`);
        if (Object.prototype.hasOwnProperty.call(edit, "text")) comment.text = String(edit.text);
        if (Object.prototype.hasOwnProperty.call(edit, "author")) comment.author = String(edit.author);
        if (Object.prototype.hasOwnProperty.call(edit, "initials")) comment.initials = edit.initials == null ? undefined : String(edit.initials);
        if (Object.prototype.hasOwnProperty.call(edit, "date")) comment.date = edit.date == null ? undefined : String(edit.date);
        continue;
      }
      throw new Error(`Unsupported document fixture edit kind ${edit.kind}.`);
  }
  docx = await DocumentFile.exportDocx(imported);
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
  });
  return { fixture, docxPath, qa };
}
