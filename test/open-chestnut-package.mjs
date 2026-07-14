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
    import { exportDocxWithOpenChestnut, exportPptxWithOpenChestnut, exportXlsxWithOpenChestnut, importDocxWithOpenChestnut, importPptxWithOpenChestnut, importXlsxWithOpenChestnut } from "open-office-artifact-tool/codecs/open-chestnut";
    import { exportXlsxWithOpenXmlWasm } from "open-office-artifact-tool/codecs/openxml-wasm";
    if (exportXlsxWithOpenXmlWasm !== exportXlsxWithOpenChestnut) process.exit(9);
    const workbook = Workbook.create({ dateSystem: "1904" });
    const sheet = workbook.worksheets.add("Packaged");
    sheet.getRange("A1:D2").values = [["Label", "Date", "Status", "Value"], ["clean install", 45853, "ready", 7]];
    const packagedTable = sheet.tables.add({ range: "A1:D2", name: "PackagedTable", style: "TableStyleMedium4" });
    packagedTable.columnDefinitions = [{ name: "Label" }, { name: "Date" }, { name: "Status" }, { name: "Value", calculatedColumnFormula: "=LEN([@Label])" }];
    packagedTable.filters = [
      { columnIndex: 0, kind: "icon", iconSet: "3Arrows", iconId: 0 },
      { columnIndex: 1, kind: "values", values: [], includeBlank: false, calendarType: "gregorian", dateGroups: [{ grouping: "day", year: 2026, month: 7, day: 15 }] },
      { columnIndex: 2, kind: "dynamic", type: "today", value: 45853, maxValue: 45854 },
      { columnIndex: 3, kind: "top10", top: true, percent: false, value: 5, filterValue: 7 },
    ];
    packagedTable.sortState = { reference: "A2:D2", caseSensitive: false, conditions: [{ reference: "D2:D2", descending: true, kind: "icon", iconSet: "5Rating", iconId: 4 }] };
    const file = await exportXlsxWithOpenChestnut(workbook);
    const imported = await importXlsxWithOpenChestnut(file);
    if (file.bytes[0] !== 0x50 || file.bytes[1] !== 0x4b) process.exit(1);
    if (imported.worksheets.getItem("Packaged").getRange("A1:D2").values[1][3] !== 7) process.exit(2);
    const importedTable = imported.worksheets.getItem("Packaged").tables.getItemOrNullObject("PackagedTable");
    if (importedTable.style !== "TableStyleMedium4") process.exit(11);
    if (importedTable.columnDefinitions[3].calculatedColumnFormula !== "=LEN([@Label])") process.exit(12);
    if (importedTable.filters[0]?.iconSet !== "3Arrows" || importedTable.filters[0]?.iconId !== 0 || importedTable.filters[1]?.dateGroups[0]?.day !== 15 || importedTable.filters[2]?.type !== "today" || importedTable.filters[3]?.filterValue !== 7) process.exit(13);
    if (!importedTable.sortState?.conditions[0]?.descending || importedTable.sortState.conditions[0]?.iconSet !== "5Rating" || importedTable.sortState.conditions[0]?.iconId !== 4) process.exit(14);
    const legacyFile = await exportXlsxWithOpenXmlWasm(workbook);
    if (legacyFile.metadata.codec !== "open-chestnut") process.exit(10);
    const document = DocumentModel.create({ paragraphs: ["clean install DOCX"] });
    const docx = await exportDocxWithOpenChestnut(document);
    const importedDocument = await importDocxWithOpenChestnut(docx);
    if (docx.bytes[0] !== 0x50 || docx.bytes[1] !== 0x4b) process.exit(3);
    if (importedDocument.blocks[0].text !== "clean install DOCX") process.exit(4);
    const presentation = Presentation.create();
    const packagedSlide = presentation.slides.add({ name: "Packaged" });
    packagedSlide.shapes.add({ name: "Title", text: [{ bulletCharacter: "•", bulletFont: "Georgia", bulletColor: "#2563EB", bulletSizePercent: 1.25, runs: [{ text: "clean install PPTX", link: { uri: "https://example.com/packaged", tooltip: "Packaged link" } }] }], position: { left: 40, top: 40, width: 640, height: 80 } });
    const pptx = await exportPptxWithOpenChestnut(presentation);
    const importedPresentation = await importPptxWithOpenChestnut(pptx);
    if (pptx.bytes[0] !== 0x50 || pptx.bytes[1] !== 0x4b) process.exit(5);
    if (importedPresentation.slides.getItem(0).shapes.items[0].text.value !== "clean install PPTX") process.exit(6);
    const marker = importedPresentation.slides.getItem(0).shapes.items[0].text.paragraphs[0];
    if (marker.bulletCharacter !== "•" || marker.bulletFont !== "Georgia" || marker.bulletColor.toLowerCase() !== "#2563eb" || marker.bulletSizePercent !== 1.25) process.exit(7);
    if (marker.runs[0].link?.uri !== "https://example.com/packaged" || marker.runs[0].link?.tooltip !== "Packaged link") process.exit(8);
  `;
  run(process.execPath, ["--input-type=module", "-e", probe], temporary, {
    PATH: process.platform === "win32" ? "C:\\Windows\\System32" : "/usr/bin:/bin",
  });
} finally {
  fs.rmSync(temporary, { force: true, recursive: true });
}

console.log("OpenChestnut clean-install package smoke ok");

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
