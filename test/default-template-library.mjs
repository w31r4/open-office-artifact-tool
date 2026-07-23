import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { crc32 } from "node:zlib";

import {
  DocumentFile,
  FileBlob,
  PresentationFile,
  SpreadsheetFile,
} from "../src/index.mjs";
import { materializeTemplate } from "../skills/default-template-library/scripts/materialize-template.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const libraryRoot = path.join(repoRoot, "skills", "default-template-library");
const integrity = JSON.parse(await fs.readFile(path.join(libraryRoot, "integrity.json"), "utf8"));
const manifest = JSON.parse(await fs.readFile(path.join(libraryRoot, "manifest.json"), "utf8"));
const plugin = JSON.parse(await fs.readFile(path.join(libraryRoot, ".codex-plugin", "plugin.json"), "utf8"));

const TEMPLATES = [
  ["artifact-template-design-report", "Design Report", "document", ".docx"],
  ["artifact-template-experiment-analysis", "Experiment Analysis", "document", ".docx"],
  ["artifact-template-investment-committee-memo", "Investment Committee Memo", "document", ".docx"],
  ["artifact-template-legal-memorandum", "Legal Memorandum", "document", ".docx"],
  ["artifact-template-minimal-letterhead", "Minimal Letterhead", "document", ".docx"],
  ["artifact-template-strategy-memorandum", "Strategy Memorandum", "document", ".docx"],
  ["artifact-template-system-design", "System Design", "document", ".docx"],
  ["artifact-template-business-review", "Business Review", "presentation", ".pptx"],
  ["artifact-template-market-trends-report", "Market Trends Report", "presentation", ".pptx"],
  ["artifact-template-operating-review", "Operating Review", "presentation", ".pptx"],
  ["artifact-template-project-kickoff", "Project Kickoff", "presentation", ".pptx"],
  ["artifact-template-simple-dark-mode", "Simple Dark Mode", "presentation", ".pptx"],
  ["artifact-template-simple-light-mode", "Simple Light Mode", "presentation", ".pptx"],
  ["artifact-template-team-alignment", "Team Alignment", "presentation", ".pptx"],
  ["artifact-template-analytics-dashboard", "Analytics Dashboard", "spreadsheet", ".xlsx"],
  ["artifact-template-financial-budget", "Financial Budget", "spreadsheet", ".xlsx"],
  ["artifact-template-operating-calendar", "Operating Calendar", "spreadsheet", ".xlsx"],
  ["artifact-template-project-tracker", "Project Tracker", "spreadsheet", ".xlsx"],
  ["artifact-template-sales-pipeline", "Sales Pipeline", "spreadsheet", ".xlsx"],
  ["artifact-template-three-statement-forecast", "Three-Statement Forecast", "spreadsheet", ".xlsx"],
];

const SPREADSHEET_RECALCULATION_SENTINELS = new Map([
  ["artifact-template-analytics-dashboard", { sheet: "Dashboard", address: "B4" }],
  ["artifact-template-financial-budget", { sheet: "Summary", address: "E8" }],
  ["artifact-template-operating-calendar", { sheet: "Annual", address: "C2" }],
  ["artifact-template-project-tracker", { sheet: "Project Plan", address: "M9" }],
  ["artifact-template-sales-pipeline", [
    { sheet: "Sales Pipeline", address: "B10" },
    { sheet: "Sales Pipeline", address: "H6" },
  ]],
  ["artifact-template-three-statement-forecast", { sheet: "Exec Sum", address: "C7" }],
]);

const SPREADSHEET_SOURCE_EMPTY_FORMULA_CELL_COUNTS = new Map([
  ["artifact-template-analytics-dashboard", 0],
  ["artifact-template-financial-budget", 0],
  ["artifact-template-operating-calendar", 323],
  ["artifact-template-project-tracker", 556],
  ["artifact-template-sales-pipeline", 0],
  ["artifact-template-three-statement-forecast", 0],
]);

const SPREADSHEET_FORMULA_CELL_COUNTS = new Map([
  ["artifact-template-analytics-dashboard", 239],
  ["artifact-template-financial-budget", 465],
  ["artifact-template-operating-calendar", 803],
  ["artifact-template-project-tracker", 655],
  ["artifact-template-sales-pipeline", 83],
  ["artifact-template-three-statement-forecast", 2770],
]);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function yamlValue(text, key) {
  const value = text.match(new RegExp(`^\\s{2}${key}:\\s*(.+)$`, "mu"))?.[1]?.trim();
  if (!value) return undefined;
  if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
    try { return JSON.parse(value); } catch { return value.slice(1, -1); }
  }
  return value;
}

function hasValidPngStructure(bytes) {
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  if (bytes.length < signature.length + 12 || !bytes.subarray(0, signature.length).equals(signature)) return false;
  let offset = signature.length;
  let sawIhdr = false;
  while (offset + 12 <= bytes.length) {
    const length = bytes.readUInt32BE(offset);
    const dataStart = offset + 8;
    const dataEnd = dataStart + length;
    const checksumEnd = dataEnd + 4;
    if (dataEnd < dataStart || checksumEnd > bytes.length) return false;
    const type = bytes.toString("ascii", offset + 4, offset + 8);
    const data = bytes.subarray(dataStart, dataEnd);
    if ((crc32(Buffer.concat([Buffer.from(type, "ascii"), data])) >>> 0) !== bytes.readUInt32BE(dataEnd)) return false;
    if (!sawIhdr && (type !== "IHDR" || length !== 13)) return false;
    sawIhdr ||= type === "IHDR";
    offset = checksumEnd;
    if (type === "IEND") return sawIhdr && length === 0 && offset === bytes.length;
  }
  return false;
}

async function walkFiles(root) {
  const files = [];
  const entries = await fs.readdir(root, { withFileTypes: true });
  for (const entry of entries) {
    const target = path.join(root, entry.name);
    const stat = await fs.lstat(target);
    assert.equal(stat.isSymbolicLink(), false, `template library must not contain symbolic links: ${target}`);
    if (stat.isDirectory()) files.push(...await walkFiles(target));
    else if (stat.isFile()) files.push(target);
    else assert.fail(`template library must contain only regular files and directories: ${target}`);
  }
  return files.sort();
}

function updateAggregate(hash, relativePath, bytes) {
  hash.update(relativePath, "utf8");
  hash.update("\0", "utf8");
  hash.update(bytes);
  hash.update("\0", "utf8");
}

function commandAvailable(command) {
  const probe = spawnSync(command, ["--version"], { encoding: "utf8" });
  return !probe.error && probe.status === 0;
}

async function assertNativeRender(sourcePath, outputDirectory) {
  const profile = path.join(outputDirectory, "profile");
  await fs.mkdir(outputDirectory, { recursive: true });
  const converted = spawnSync("soffice", [
    `-env:UserInstallation=${pathToFileURL(profile).href}`,
    "--headless",
    "--convert-to", "pdf",
    "--outdir", outputDirectory,
    sourcePath,
  ], { encoding: "utf8" });
  assert.equal(converted.status, 0, `LibreOffice could not render ${sourcePath}\n${converted.stdout}\n${converted.stderr}`);
  const pdfPath = path.join(outputDirectory, `${path.basename(sourcePath, path.extname(sourcePath))}.pdf`);
  const info = spawnSync("pdfinfo", [pdfPath], { encoding: "utf8" });
  assert.equal(info.status, 0, `Poppler could not inspect ${pdfPath}\n${info.stdout}\n${info.stderr}`);
  const pages = Number(info.stdout.match(/^Pages:\s*([1-9]\d*)/m)?.[1]);
  assert.ok(Number.isInteger(pages), `native render needs at least one page: ${sourcePath}`);
  const prefix = path.join(outputDirectory, "page");
  const raster = spawnSync("pdftoppm", ["-png", "-f", "1", "-l", String(pages), pdfPath, prefix], { encoding: "utf8" });
  assert.equal(raster.status, 0, `Poppler could not rasterize ${pdfPath}\n${raster.stdout}\n${raster.stderr}`);
  for (let page = 1; page <= pages; page += 1) {
    assert.ok((await fs.stat(`${prefix}-${page}.png`)).size > 0, `native raster is empty: ${sourcePath} page ${page}`);
  }
}

function workbookXmlWithForcedCalculation(workbookXml) {
  const calcPr = '<calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>';
  const start = workbookXml.indexOf("<calcPr");
  const end = start < 0 ? -1 : workbookXml.indexOf("/>", start);
  return end >= 0
    ? `${workbookXml.slice(0, start)}${calcPr}${workbookXml.slice(end + 2)}`
    : workbookXml.replace("</workbook>", `${calcPr}</workbook>`);
}

async function assertNativeSpreadsheetCalculation(templateId, sourcePath, outputDirectory) {
  const inputDirectory = path.join(outputDirectory, "input");
  const nativeDirectory = path.join(outputDirectory, "native");
  const profile = path.join(outputDirectory, "profile");
  await Promise.all([fs.mkdir(inputDirectory, { recursive: true }), fs.mkdir(nativeDirectory, { recursive: true }), fs.mkdir(profile, { recursive: true })]);
  const inputPath = path.join(inputDirectory, "forced-recalculation.xlsx");
  const sourceZip = await JSZip.loadAsync(await fs.readFile(sourcePath));
  const workbookXml = await sourceZip.file("xl/workbook.xml")?.async("text");
  assert.ok(workbookXml, `Spreadsheet source must include xl/workbook.xml: ${sourcePath}`);
  sourceZip.file("xl/workbook.xml", workbookXmlWithForcedCalculation(workbookXml));
  await fs.writeFile(inputPath, await sourceZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));

  const converted = spawnSync("soffice", [
    `-env:UserInstallation=${pathToFileURL(profile).href}`,
    "--headless",
    "--convert-to", "xlsx:Calc MS Excel 2007 XML",
    "--outdir", nativeDirectory,
    inputPath,
  ], { encoding: "utf8", timeout: 60_000 });
  assert.equal(converted.status, 0, `LibreOffice could not force-recalculate ${sourcePath}\n${converted.stdout}\n${converted.stderr}`);
  const nativePath = path.join(nativeDirectory, "forced-recalculation.xlsx");
  assert.ok((await fs.stat(nativePath)).size > 0, `LibreOffice did not save a recalculated workbook: ${sourcePath}`);

  const [model, native] = await Promise.all([
    SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath)),
    SpreadsheetFile.importXlsx(await FileBlob.load(nativePath)),
  ]);
  model.recalculate();
  let formulaCells = 0;
  for (const sheet of model.worksheets.items) {
    const nativeSheet = native.worksheets.getItem(sheet.name);
    assert.ok(nativeSheet, `LibreOffice recalculation must retain worksheet ${sheet.name}: ${sourcePath}`);
    for (const [address, cell] of sheet.store.entries()) {
      if (!cell.formula) continue;
      formulaCells += 1;
      const nativeCell = nativeSheet.store.get(address);
      assert.ok(nativeCell?.formula, `LibreOffice recalculation must retain formula ${sheet.name}!${address}: ${sourcePath}`);
      assert.equal(sameFormulaValue(cell.value, nativeCell.value), true, `Spreadsheet model calculation must match LibreOffice at ${sheet.name}!${address}: ${sourcePath}`);
    }
  }
  assert.equal(formulaCells, SPREADSHEET_FORMULA_CELL_COUNTS.get(templateId), `Spreadsheet native formula inventory must remain pinned: ${sourcePath}`);
}

function sameFormulaValue(left, right) {
  if (typeof left === "number" && typeof right === "number") return Math.abs(left - right) <= Math.max(1e-9, Math.abs(left) * 1e-10);
  return left === right;
}

function hasCachedFormulaValue(value) {
  return value !== null && value !== undefined && value !== "";
}

function assertSpreadsheetTemplateCalculation(templateId, workbook, sourcePath) {
  const configured = SPREADSHEET_RECALCULATION_SENTINELS.get(templateId);
  const sentinels = Array.isArray(configured) ? configured : [configured];
  assert.ok(sentinels.every(Boolean), `Spreadsheet template needs a model-calculation sentinel: ${sourcePath}`);
  const cells = sentinels.map((sentinel) => {
    const sheet = workbook.worksheets.getItem(sentinel.sheet);
    assert.ok(sheet, `Spreadsheet template sentinel sheet is missing: ${sourcePath}`);
    const cell = sheet.store.get(sentinel.address);
    assert.ok(cell.formula, `Spreadsheet template sentinel must retain a formula: ${sourcePath}`);
    return { ...sentinel, cell, cachedValue: cell.value };
  });
  const cachedFormulaCells = [];
  const uncachedFormulaCells = [];
  for (const candidateSheet of workbook.worksheets.items) {
    for (const [address, candidate] of candidateSheet.store.entries()) {
      if (candidate.formula && hasCachedFormulaValue(candidate.value))
        cachedFormulaCells.push({ sheet: candidateSheet.name, address, cachedValue: candidate.value });
      if (candidate.formula && !hasCachedFormulaValue(candidate.value))
        uncachedFormulaCells.push({ sheet: candidateSheet.name, address });
    }
  }
  assert.ok(cachedFormulaCells.length, `Spreadsheet template must retain at least one cached formula value: ${sourcePath}`);
  assert.equal(uncachedFormulaCells.length, SPREADSHEET_SOURCE_EMPTY_FORMULA_CELL_COUNTS.get(templateId), `Spreadsheet template source-empty formula inventory must remain pinned: ${sourcePath}`);
  workbook.recalculate();
  for (const { address, cell, cachedValue } of cells)
    assert.equal(sameFormulaValue(cell.value, cachedValue), true, `Spreadsheet template model calculation must match its cached sentinel ${address}: ${sourcePath}`);
  for (const { sheet: sheetName, address, cachedValue } of cachedFormulaCells) {
    const actual = workbook.worksheets.getItem(sheetName).store.get(address).value;
    assert.equal(sameFormulaValue(actual, cachedValue), true, `Spreadsheet template model calculation must match cached ${sheetName}!${address}: ${sourcePath}`);
  }
  for (const { sheet: sheetName, address } of uncachedFormulaCells) {
    const actual = workbook.worksheets.getItem(sheetName).store.get(address).value;
    assert.equal(actual, "", `Spreadsheet template source-empty formula must retain its intentional empty-text branch at ${sheetName}!${address}: ${sourcePath}`);
  }
  const errors = [];
  for (const candidateSheet of workbook.worksheets.items) {
    for (const [address, candidate] of candidateSheet.store.entries()) {
      if (candidate.formula && /^#(?:NAME\?|VALUE!|REF!|DIV\/0!|NUM!|N\/A|CYCLE!|SPILL!)$/.test(String(candidate.value))) errors.push(`${candidateSheet.name}!${address}`);
    }
  }
  assert.deepEqual(errors, [], `Spreadsheet template model calculation must not leave formula errors: ${sourcePath}`);
}

async function assertPublicOfficeRoundTrip(templateId, kind, sourcePath) {
  const source = await FileBlob.load(sourcePath);
  if (kind === "document") {
    const imported = await DocumentFile.importDocx(source);
    const exported = await DocumentFile.exportDocx(imported);
    const reimported = await DocumentFile.importDocx(exported);
    assert.equal(reimported.blocks.length, imported.blocks.length, `Document facade round trip: ${sourcePath}`);
    return exported;
  }
  if (kind === "presentation") {
    const imported = await PresentationFile.importPptx(source);
    const exported = await PresentationFile.exportPptx(imported);
    const reimported = await PresentationFile.importPptx(exported);
    assert.equal(reimported.slides.items.length, imported.slides.items.length, `Presentation facade round trip: ${sourcePath}`);
    return exported;
  }
  if (kind === "spreadsheet") {
    const imported = await SpreadsheetFile.importXlsx(source);
    assertSpreadsheetTemplateCalculation(templateId, imported, sourcePath);
    const exported = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
    const reimported = await SpreadsheetFile.importXlsx(exported);
    assert.equal(reimported.worksheets.items.length, imported.worksheets.items.length, `Spreadsheet facade round trip: ${sourcePath}`);
    return exported;
  }
  assert.fail(`Unknown retained template kind: ${kind}`);
}

function presentationPlaceholderIdentity(shape) {
  const snapshot = structuredClone(shape.layoutJson());
  delete snapshot.text;
  delete snapshot.paragraphs;
  return snapshot;
}

function presentationTextFormatting(shape) {
  const paragraphs = structuredClone(shape.text.paragraphs);
  for (const paragraph of paragraphs) for (const run of paragraph.runs || []) {
    if (Object.hasOwn(run, "text")) run.text = "";
    if (run.field) run.field.text = "";
  }
  return paragraphs;
}

async function assertPublicPresentationPlaceholderTextEdit(sourcePath) {
  const source = await FileBlob.load(sourcePath);
  const presentation = await PresentationFile.importPptx(source);
  const target = presentation.slides.items
    .flatMap((slide) => slide.shapes.items)
    .find((shape) => shape.placeholder && shape.text.value.trim());
  assert.ok(target, `Presentation template must expose a visible slide placeholder: ${sourcePath}`);
  assert.equal(target.placeholder.textEditable, true, `Presentation template placeholder must advertise its verified text capability: ${sourcePath}`);
  const id = target.id;
  const identity = presentationPlaceholderIdentity(target);
  const formatting = presentationTextFormatting(target);
  const replacement = `${target.text.value} · Agent QA`;
  // Agent workflows commonly use TextFrame.set(). Imported styled titles must
  // preserve their native formatting while replacing only the characters.
  target.text.set(replacement);
  const exported = await PresentationFile.exportPptx(presentation);
  const reimported = await PresentationFile.importPptx(exported);
  const roundTrip = reimported.slides.items.flatMap((slide) => slide.shapes.items).find((shape) => shape.id === id);
  assert.ok(roundTrip, `Edited placeholder identity must survive reimport: ${sourcePath}`);
  assert.equal(roundTrip.placeholder.textEditable, true, `Reimported placeholder must re-prove its text capability: ${sourcePath}`);
  assert.equal(roundTrip.text.value, replacement, `Edited placeholder text must survive reimport: ${sourcePath}`);
  assert.deepEqual(presentationPlaceholderIdentity(roundTrip), identity, `Placeholder geometry/style/identity must stay source-bound: ${sourcePath}`);
  assert.deepEqual(presentationTextFormatting(roundTrip), formatting, `Placeholder paragraph/run formatting must stay source-bound: ${sourcePath}`);
  return exported;
}

function documentParagraphFormattingIdentity(block) {
  return {
    styleId: block.styleId,
    paragraphFormat: structuredClone(block.paragraphFormat),
    runStyles: block.runs.map((run) => structuredClone(run.style)),
  };
}

async function assertPublicDocumentTextEdit(sourcePath) {
  const source = await FileBlob.load(sourcePath);
  const document = await DocumentFile.importDocx(source);
  const editable = document.blocks.find((block) =>
    block.kind === "paragraph" && block.textEditable && block.runs.length === 1 && block.text.trim());
  const patchable = document.blocks.find((block) =>
    block.kind === "paragraph" && block.textPatchable && block.text.trim());
  const target = editable || patchable;
  if (target) {
    const identity = documentParagraphFormattingIdentity(target);
    const replacement = `${target.text} · Agent QA`;
    const range = document.resolve(`${target.id}/text`);
    assert.ok(range, `Document template must resolve its advertised paragraph text range: ${sourcePath}`);
    if (target.textEditable) range.text = replacement;
    else {
      assert.throws(() => { range.text = replacement; }, /source-bound.*replace/i, `Source-bound paragraph assignment must fail early: ${sourcePath}`);
      range.replace(target.text, replacement);
    }
    const exported = await DocumentFile.exportDocx(document);
    const reimported = await DocumentFile.importDocx(exported);
    const roundTrip = reimported.blocks.find((block) => block.id === target.id);
    assert.ok(roundTrip, `Edited paragraph identity must survive reimport: ${sourcePath}`);
    assert.equal(roundTrip.text, replacement, `Edited paragraph text must survive reimport: ${sourcePath}`);
    assert.deepEqual(documentParagraphFormattingIdentity(roundTrip), identity, `Paragraph/run formatting must stay source-bound: ${sourcePath}`);
    return exported;
  }

  const table = document.blocks.find((block) => block.kind === "table" && block.cells.some((cell) => cell.textPatchable));
  assert.ok(table, `Document template must expose a paragraph or table-cell source text capability: ${sourcePath}`);
  const cellRecord = table.cells.find((cell) => cell.textPatchable && table.getCell(cell.row, cell.column).value.includes("[Greeting]")) || table.cells.find((cell) => cell.textPatchable);
  const cell = table.getCell(cellRecord.row, cellRecord.column);
  assert.equal(cell.textPatchable, true, `Complex document cell must advertise textPatchable: ${sourcePath}`);
  if (!cell.editable) assert.throws(() => { cell.value = "Unsafe whole-cell replacement"; }, /whole-cell replacement/i, `Complex cell assignment must fail early: ${sourcePath}`);
  const search = cell.value.includes("[Greeting]") ? "[Greeting]" : cell.value;
  const replacement = search === "[Greeting]" ? "Hello · Agent QA" : `${search} · Agent QA`;
  const range = document.resolve(`${cell.id}/text`);
  assert.ok(range, `Document template must resolve its advertised table-cell text range: ${sourcePath}`);
  range.replace(search, replacement);
  const tableInspection = JSON.parse(document.inspect({ kind: "table", target: table.id }).ndjson.trim());
  assert.equal(tableInspection.pendingTextPatches, 1, `Parent table inspection must expose the pending patch: ${sourcePath}`);
  assert.equal(tableInspection.values[cellRecord.row][cellRecord.column].includes(replacement), true, `Parent table inspection must project patched cell text: ${sourcePath}`);
  const exported = await DocumentFile.exportDocx(document);
  const reimported = await DocumentFile.importDocx(exported);
  const roundTrip = reimported.resolve(cell.id);
  assert.ok(roundTrip, `Edited table-cell identity must survive reimport: ${sourcePath}`);
  assert.equal(roundTrip.value.includes(replacement), true, `Edited table-cell text must survive reimport: ${sourcePath}`);
  assert.equal(roundTrip.textPatchable, true, `Reimported table cell must re-prove textPatchable: ${sourcePath}`);
  return exported;
}

assert.equal(plugin.name, "default-template-library");
assert.equal(plugin.license, "MIT");
assert.equal(plugin.skills, "./skills/");
assert.equal(manifest.schemaVersion, 1);
assert.equal(manifest.name, "default-template-library");
assert.deepEqual(manifest.skills, TEMPLATES.map(([id]) => `skills/${id}`));
assert.deepEqual(integrity.source, {
  repository: "https://github.com/w31r4/office-artifact-tool",
  commit: "256cb31bfe0a07b3cef0051b6b159342be381378",
  license: "MIT",
  copyright: "Copyright (c) 2026 w31r4",
});
assert.match(await fs.readFile(path.join(libraryRoot, "LICENSE.md"), "utf8"), /MIT License\n\nCopyright \(c\) 2026 w31r4/);
assert.equal(integrity.assets.length, TEMPLATES.length * 2);

const expectedFiles = new Set([
  ".codex-plugin/plugin.json",
  "LICENSE.md",
  "README.md",
  "manifest.json",
  "integrity.json",
  "assets/icon.svg",
  "scripts/materialize-template.mjs",
]);
const aggregate = crypto.createHash("sha256");
let totalBytes = 0;
for (const [id, displayName, kind, extension] of [...TEMPLATES].sort((left, right) => left[0].localeCompare(right[0]))) {
  const skillRoot = path.join(libraryRoot, "skills", id);
  const referencePath = `assets/reference${extension}`;
  for (const relativePath of ["SKILL.md", "artifact-template.json", "agents/agent.yaml", "assets/preview.png", referencePath]) {
    expectedFiles.add(path.posix.join("skills", id, relativePath));
  }
  const [skillText, agentText, sidecarText, previewBytes, referenceBytes] = await Promise.all([
    fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8"),
    fs.readFile(path.join(skillRoot, "agents", "agent.yaml"), "utf8"),
    fs.readFile(path.join(skillRoot, "artifact-template.json"), "utf8"),
    fs.readFile(path.join(skillRoot, "assets", "preview.png")),
    fs.readFile(path.join(skillRoot, referencePath)),
  ]);
  assert.equal(skillText.match(/^name:\s*(.+)$/mu)?.[1]?.trim(), id, `${id} frontmatter name`);
  assert.equal(yamlValue(agentText, "display_name"), displayName, `${id} display name`);
  assert.equal(yamlValue(agentText, "icon_large"), "./assets/preview.png", `${id} preview icon`);
  assert.deepEqual(JSON.parse(sidecarText), { schemaVersion: 1, kind, reference: referencePath, preview: "assets/preview.png" }, `${id} sidecar`);
  assert.equal(hasValidPngStructure(previewBytes), true, `${id} preview PNG structure`);
  assert.equal(referenceBytes.subarray(0, 4).equals(Buffer.from([0x50, 0x4b, 0x03, 0x04])), true, `${id} Office reference ZIP signature`);

  for (const [role, relativePath, bytes] of [["preview", "assets/preview.png", previewBytes], ["reference", referencePath, referenceBytes]]) {
    const assetPath = path.posix.join("skills", id, relativePath);
    const record = integrity.assets.find((asset) => asset.templateId === id && asset.role === role);
    assert.deepEqual({ templateId: id, role, path: assetPath, bytes: bytes.length, sha256: sha256(bytes) }, record, `${id} ${role} integrity`);
    updateAggregate(aggregate, assetPath, bytes);
    totalBytes += bytes.length;
  }
}
assert.equal(aggregate.digest("hex"), integrity.assetAggregateSha256, "binary aggregate SHA-256");
assert.ok(totalBytes <= 32 * 1024 * 1024, `template binary budget exceeded: ${totalBytes}`);
assert.ok(integrity.assets.filter((asset) => asset.role === "preview").every((asset) => asset.bytes <= 512 * 1024), "preview budget exceeded");
assert.ok(integrity.assets.filter((asset) => asset.role === "reference").every((asset) => asset.bytes <= 8 * 1024 * 1024), "Office reference budget exceeded");
const actualFiles = (await walkFiles(libraryRoot)).map((file) => path.relative(libraryRoot, file).split(path.sep).join("/"));
assert.deepEqual(actualFiles, [...expectedFiles].sort(), "template library canonical file inventory");

const sourceRoot = process.env.OFFICE_TEMPLATE_SOURCE_ROOT;
if (sourceRoot) {
  for (const asset of integrity.assets) {
    const [sourceBytes, targetBytes] = await Promise.all([fs.readFile(path.join(sourceRoot, asset.path)), fs.readFile(path.join(libraryRoot, asset.path))]);
    assert.deepEqual(targetBytes, sourceBytes, `source byte match: ${asset.path}`);
  }
}

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-template-library-"));
try {
  const materialized = [];
  for (const [id, , kind, extension] of TEMPLATES) {
    const output = path.join(temporary, `${id}${extension}`);
    const audit = path.join(temporary, `${id}.audit.json`);
    const result = await materializeTemplate({ templateId: id, outputPath: output, auditPath: audit });
    const reference = integrity.assets.find((asset) => asset.templateId === id && asset.role === "reference");
    assert.equal(result.audit.operation, "materialize-retained-reference");
    assert.equal(result.audit.source.sha256, reference.sha256);
    assert.equal(result.audit.output.sha256, reference.sha256);
    assert.equal((await fs.readFile(output)).equals(await fs.readFile(path.join(libraryRoot, reference.path))), true, `${id} materialized bytes`);
    await assert.rejects(materializeTemplate({ templateId: id, outputPath: output, auditPath: audit }), /already exists/);
    materialized.push({ id, kind, output });
  }

  const roundTripped = [];
  for (const { id, kind, output } of materialized) {
    const exported = await assertPublicOfficeRoundTrip(id, kind, output);
    const roundTripOutput = path.join(temporary, `${id}-openchestnut${path.extname(output)}`);
    await exported.save(roundTripOutput);
    roundTripped.push({ id, output: roundTripOutput });
  }

  const editedPresentations = [];
  for (const { id, kind, output } of materialized.filter((item) => item.kind === "presentation")) {
    const exported = await assertPublicPresentationPlaceholderTextEdit(output);
    const editedOutput = path.join(temporary, `${id}-placeholder-edit.pptx`);
    await exported.save(editedOutput);
    editedPresentations.push({ id, output: editedOutput });
  }

  const editedDocuments = [];
  for (const { id, kind, output } of materialized.filter((item) => item.kind === "document")) {
    const exported = await assertPublicDocumentTextEdit(output);
    const editedOutput = path.join(temporary, `${id}-text-edit.docx`);
    await exported.save(editedOutput);
    editedDocuments.push({ id, output: editedOutput });
  }

  const structuredPresentation = materialized.find((item) => item.id === "artifact-template-market-trends-report");
  assert.ok(structuredPresentation, "Market Trends retained template must be materialized");
  const topologyProbe = await PresentationFile.importPptx(await FileBlob.load(structuredPresentation.output));
  const topologyTarget = topologyProbe.slides.items.flatMap((slide) => slide.shapes.items)
    .find((shape) => shape.placeholder && shape.text.value.includes("\n"));
  assert.ok(topologyTarget, "Market Trends template must retain a multi-line placeholder title");
  topologyTarget.text.set(topologyTarget.text.value.replace("\n", " "));
  await assert.rejects(
    () => PresentationFile.exportPptx(topologyProbe),
    (error) => error?.code === "presentation_text_topology_changed",
    "imported placeholder text.set must fail closed when it changes the source line-break topology",
  );

  const financialBudget = materialized.find((item) => item.id === "artifact-template-financial-budget");
  assert.ok(financialBudget, "Financial Budget retained template must be materialized");
  const partialSharedFormulaWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(financialBudget.output));
  partialSharedFormulaWorkbook.worksheets.getItem("Op Build").getRange("C24").values = [[42]];
  await assert.rejects(
    () => SpreadsheetFile.exportXlsx(partialSharedFormulaWorkbook, { recalculate: false }),
    (error) => error?.code === "unsupported_cell_formula_edit",
    "partial native shared-formula edits must fail closed through the public facade",
  );

  if (commandAvailable("soffice")) {
    for (const spreadsheet of materialized.filter((item) => item.kind === "spreadsheet"))
      await assertNativeSpreadsheetCalculation(spreadsheet.id, spreadsheet.output, path.join(temporary, "native-calculation", spreadsheet.id));
  }

  if (["soffice", "pdfinfo", "pdftoppm"].every(commandAvailable)) {
    const rendered = path.join(temporary, "native-render");
    for (const { id, output } of materialized) await assertNativeRender(output, path.join(rendered, id, "source"));
    for (const { id, output } of roundTripped) await assertNativeRender(output, path.join(rendered, id, "openchestnut"));
    for (const { id, output } of editedDocuments) await assertNativeRender(output, path.join(rendered, id, "text-edit"));
    for (const { id, output } of editedPresentations) await assertNativeRender(output, path.join(rendered, id, "placeholder-edit"));
  }
} finally {
  await fs.rm(temporary, { force: true, recursive: true });
}

console.log("default template library integrity and materialization smoke ok");
