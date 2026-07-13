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

  run("npm", ["install", "--ignore-scripts", "--no-audit", "--no-fund", "--omit=dev", tarball], temporary);
  const probe = `
    import { DocumentModel, Presentation, Workbook } from "open-office-artifact-tool";
    import { exportDocxWithOpenXmlWasm, exportPptxWithOpenXmlWasm, exportXlsxWithOpenXmlWasm, importDocxWithOpenXmlWasm, importPptxWithOpenXmlWasm, importXlsxWithOpenXmlWasm } from "open-office-artifact-tool/codecs/openxml-wasm";
    const workbook = Workbook.create({ dateSystem: "1904" });
    const sheet = workbook.worksheets.add("Packaged");
    sheet.getRange("A1:B1").values = [["clean install", 7]];
    const file = await exportXlsxWithOpenXmlWasm(workbook);
    const imported = await importXlsxWithOpenXmlWasm(file);
    if (file.bytes[0] !== 0x50 || file.bytes[1] !== 0x4b) process.exit(1);
    if (imported.worksheets.getItem("Packaged").getRange("A1:B1").values[0][1] !== 7) process.exit(2);
    const document = DocumentModel.create({ paragraphs: ["clean install DOCX"] });
    const docx = await exportDocxWithOpenXmlWasm(document);
    const importedDocument = await importDocxWithOpenXmlWasm(docx);
    if (docx.bytes[0] !== 0x50 || docx.bytes[1] !== 0x4b) process.exit(3);
    if (importedDocument.blocks[0].text !== "clean install DOCX") process.exit(4);
    const presentation = Presentation.create();
    presentation.slides.add({ name: "Packaged" }).shapes.add({ name: "Title", text: [{ bulletCharacter: "•", bulletFont: "Georgia", bulletColor: "#2563EB", bulletSizePercent: 1.25, runs: ["clean install PPTX"] }], position: { left: 40, top: 40, width: 640, height: 80 } });
    const pptx = await exportPptxWithOpenXmlWasm(presentation);
    const importedPresentation = await importPptxWithOpenXmlWasm(pptx);
    if (pptx.bytes[0] !== 0x50 || pptx.bytes[1] !== 0x4b) process.exit(5);
    if (importedPresentation.slides.getItem(0).shapes.items[0].text.value !== "clean install PPTX") process.exit(6);
    const marker = importedPresentation.slides.getItem(0).shapes.items[0].text.paragraphs[0];
    if (marker.bulletCharacter !== "•" || marker.bulletFont !== "Georgia" || marker.bulletColor.toLowerCase() !== "#2563eb" || marker.bulletSizePercent !== 1.25) process.exit(7);
  `;
  run(process.execPath, ["--input-type=module", "-e", probe], temporary, {
    PATH: process.platform === "win32" ? "C:\\Windows\\System32" : "/usr/bin:/bin",
  });
} finally {
  fs.rmSync(temporary, { force: true, recursive: true });
}

console.log("openxml wasm clean-install package smoke ok");

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
