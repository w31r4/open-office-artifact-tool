import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), "open-office-artifact-tool-pack-"));

try {
  const packed = run("npm", ["pack", repoRoot, "--json", "--ignore-scripts", "--pack-destination", temporary], repoRoot);
  const report = JSON.parse(packed.stdout)[0];
  const tarball = path.join(temporary, report.filename);
  assert.ok(fs.existsSync(tarball), `npm pack did not create ${tarball}`);
  const dependencyTarballs = packProductionDependencies(temporary);
  // Exercise a real npm install without making a release gate depend on the
  // registry. Optional renderer peers are intentionally outside this core
  // OpenChestnut/PDF probe and remain covered by package metadata tests.
  run("npm", [
    "install", "--offline", "--ignore-scripts", "--no-audit", "--no-fund",
    "--omit=dev", "--legacy-peer-deps", "--no-save", tarball, ...dependencyTarballs,
  ], temporary);

  const probe = String.raw`
    import {
      DocumentFile, DocumentModel, PdfArtifact, PdfFile,
      Presentation, PresentationFile, SpreadsheetFile, Workbook,
    } from "open-office-artifact-tool";

    try {
      await import("open-office-artifact-tool/codecs/openxml-wasm");
      process.exit(40);
    } catch (error) {
      if (error?.code !== "ERR_PACKAGE_PATH_NOT_EXPORTED") process.exit(41);
    }

    const workbook = Workbook.create();
    const sheet = workbook.worksheets.add("Packaged");
    sheet.getRange("A1:B2").values = [["Label", "Value"], ["clean install", 7]];
    const xlsx = await SpreadsheetFile.exportXlsx(workbook);
    if (xlsx.metadata.codec !== "open-chestnut" || xlsx.bytes[0] !== 0x50 || xlsx.bytes[1] !== 0x4b) process.exit(1);
    const importedWorkbook = await SpreadsheetFile.importXlsx(xlsx);
    if (importedWorkbook.worksheets.getItem("Packaged").getRange("B2").values[0][0] !== 7) process.exit(2);
    const xlsx2 = await SpreadsheetFile.exportXlsx(importedWorkbook, { recalculate: false });
    if ((await SpreadsheetFile.importXlsx(xlsx2)).worksheets.getItem("Packaged").getRange("A2").values[0][0] !== "clean install") process.exit(3);

    const document = DocumentModel.create({ paragraphs: ["clean install DOCX"] });
    const docx = await DocumentFile.exportDocx(document);
    if (docx.metadata.codec !== "open-chestnut" || docx.bytes[0] !== 0x50 || docx.bytes[1] !== 0x4b) process.exit(10);
    const importedDocument = await DocumentFile.importDocx(docx);
    if (importedDocument.blocks[0].text !== "clean install DOCX") process.exit(11);
    if ((await DocumentFile.importDocx(await DocumentFile.exportDocx(importedDocument))).blocks[0].text !== "clean install DOCX") process.exit(12);

    const presentation = Presentation.create();
    presentation.slides.add({ name: "Packaged" }).shapes.add({
      name: "Title", geometry: "roundRect", text: "clean install PPTX",
      position: { left: 40, top: 40, width: 520, height: 80 },
    });
    const pptx = await PresentationFile.exportPptx(presentation);
    if (pptx.metadata.codec !== "open-chestnut" || pptx.bytes[0] !== 0x50 || pptx.bytes[1] !== 0x4b) process.exit(20);
    const importedPresentation = await PresentationFile.importPptx(pptx);
    if (importedPresentation.slides.getItem(0).shapes.items[0].text.value !== "clean install PPTX") process.exit(21);
    if ((await PresentationFile.importPptx(await PresentationFile.exportPptx(importedPresentation))).slides.count !== 1) process.exit(22);

    const pdf = PdfArtifact.create({ pages: [{ text: "clean install PDF" }] });
    const pdfFile = await PdfFile.exportPdf(pdf);
    if (pdfFile.bytes[0] !== 0x25 || pdfFile.bytes[1] !== 0x50 || pdfFile.bytes[2] !== 0x44 || pdfFile.bytes[3] !== 0x46) process.exit(30);
    const inspection = await PdfFile.inspectPdf(pdfFile);
    if (inspection.summary.pages !== 1 || !inspection.summary.tagged) process.exit(31);
    const importedPdf = await PdfFile.importPdf(pdfFile);
    if (!importedPdf.extractText().includes("clean install PDF")) process.exit(32);
  `;

  run(process.execPath, ["--input-type=module", "-e", probe], temporary, {
    PATH: process.platform === "win32" ? "C:\\Windows\\System32" : "/usr/bin:/bin",
  });
} finally {
  fs.rmSync(temporary, { force: true, recursive: true });
}

console.log("OpenChestnut and PDF clean-install package smoke ok");

function run(command, args, cwd, environment = {}) {
  const result = spawnSync(command, args, {
    cwd,
    encoding: "utf8",
    env: { ...process.env, ...environment },
    shell: false,
  });
  assert.equal(result.status, 0, `${command} ${args.join(" ")} failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
  return result;
}

function packProductionDependencies(temporary) {
  const lock = JSON.parse(fs.readFileSync(path.join(repoRoot, "package-lock.json"), "utf8"));
  const destination = path.join(temporary, "dependency-tarballs");
  fs.mkdirSync(destination, { recursive: true });
  return Object.entries(lock.packages || {})
    .filter(([location, metadata]) => location.startsWith("node_modules/") && !metadata.dev && !metadata.optional && !metadata.peer)
    .map(([location]) => {
      const source = path.join(repoRoot, location);
      assert.ok(fs.existsSync(source), `npm ci production dependency is missing: ${location}`);
      const packed = run("npm", ["pack", source, "--json", "--ignore-scripts", "--pack-destination", destination], repoRoot);
      const report = JSON.parse(packed.stdout)[0];
      return path.join(destination, report.filename);
    });
}
