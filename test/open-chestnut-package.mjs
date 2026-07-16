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
    import { DocumentFile, DocumentModel, Presentation, PresentationFile, SpreadsheetFile, Workbook } from "open-office-artifact-tool";
    import JSZip from "jszip";
    import { exportXlsxWithOpenChestnut } from "open-office-artifact-tool/codecs/open-chestnut";
    import { exportXlsxWithOpenXmlWasm } from "open-office-artifact-tool/codecs/openxml-wasm";
    if (exportXlsxWithOpenXmlWasm !== exportXlsxWithOpenChestnut) process.exit(9);
    const workbook = Workbook.create({ dateSystem: "1904" });
    const sheet = workbook.worksheets.add("Packaged");
    sheet.getRange("A1:F2").values = [["Label", "Date", "Status", "Value", "Fill", "Font"], ["clean install", 45853, "ready", 7, 1, 2]];
    const packagedTable = sheet.tables.add({ range: "A1:F2", name: "PackagedTable", style: "TableStyleMedium4" });
    const selectedSheet = workbook.worksheets.add("Selected");
    selectedSheet.getRange("A1").values = [["active"]];
    workbook.worksheets.setActiveWorksheet(selectedSheet);
    workbook.worksheets.setSelectedWorksheets([sheet, selectedSheet]);
    workbook.windows.add({ activeWorksheet: sheet, selectedWorksheets: [sheet] });
    packagedTable.columnDefinitions = [{ name: "Label" }, { name: "Date" }, { name: "Status" }, { name: "Value", calculatedColumnFormula: "=LEN([@Label])" }, { name: "Fill" }, { name: "Font" }];
    packagedTable.filters = [
      { columnIndex: 0, kind: "icon", iconSet: "3Arrows", iconId: 0 },
      { columnIndex: 1, kind: "values", values: [], includeBlank: false, calendarType: "gregorian", dateGroups: [{ grouping: "day", year: 2026, month: 7, day: 15 }] },
      { columnIndex: 2, kind: "dynamic", type: "today", value: 45853, maxValue: 45854 },
      { columnIndex: 3, kind: "top10", top: true, percent: false, value: 5, filterValue: 7 },
      { columnIndex: 4, kind: "color", target: "cell", color: "#E11D48" },
      { columnIndex: 5, kind: "color", target: "font", color: { theme: 4, tint: -0.25 } },
    ];
    packagedTable.sortState = { reference: "A2:F2", caseSensitive: false, conditions: [
      { reference: "D2:D2", descending: true, kind: "icon", iconSet: "5Rating", iconId: 4 },
      { reference: "E2:E2", descending: false, kind: "color", target: "cell", color: "#E11D48" },
      { reference: "F2:F2", descending: true, kind: "color", target: "font", color: { theme: 4, tint: -0.25 } },
    ] };
    const file = await SpreadsheetFile.exportXlsx(workbook, { codec: "open-chestnut" });
    const imported = await SpreadsheetFile.importXlsx(file, { codec: "open-chestnut" });
    if (file.bytes[0] !== 0x50 || file.bytes[1] !== 0x4b) process.exit(1);
    if (imported.worksheets.getItem("Packaged").getRange("A1:F2").values[1][3] !== 7) process.exit(2);
    if (imported.worksheets.getActiveWorksheet().name !== "Selected") process.exit(17);
    if (imported.worksheets.getSelectedWorksheets().map((item) => item.name).join(",") !== "Packaged,Selected") process.exit(18);
    if (imported.windows.count !== 2 || imported.windows.getItemAt(1).getActiveWorksheet().name !== "Packaged") process.exit(19);
    const importedTable = imported.worksheets.getItem("Packaged").tables.getItemOrNullObject("PackagedTable");
    if (importedTable.style !== "TableStyleMedium4") process.exit(11);
    if (importedTable.columnDefinitions[3].calculatedColumnFormula !== "=LEN([@Label])") process.exit(12);
    if (importedTable.filters[0]?.iconSet !== "3Arrows" || importedTable.filters[0]?.iconId !== 0 || importedTable.filters[1]?.dateGroups[0]?.day !== 15 || importedTable.filters[2]?.type !== "today" || importedTable.filters[3]?.filterValue !== 7) process.exit(13);
    if (importedTable.filters[4]?.target !== "cell" || importedTable.filters[4]?.color !== "#E11D48" || importedTable.filters[5]?.target !== "font" || importedTable.filters[5]?.color?.theme !== 4 || importedTable.filters[5]?.color?.tint !== -0.25) process.exit(15);
    if (!importedTable.sortState?.conditions[0]?.descending || importedTable.sortState.conditions[0]?.iconSet !== "5Rating" || importedTable.sortState.conditions[0]?.iconId !== 4) process.exit(14);
    if (importedTable.sortState.conditions[1]?.target !== "cell" || importedTable.sortState.conditions[1]?.color !== "#E11D48" || importedTable.sortState.conditions[2]?.target !== "font" || importedTable.sortState.conditions[2]?.color?.theme !== 4 || importedTable.sortState.conditions[2]?.color?.tint !== -0.25) process.exit(16);
    const legacyFile = await exportXlsxWithOpenXmlWasm(workbook);
    if (legacyFile.metadata.codec !== "open-chestnut") process.exit(10);
    const document = DocumentModel.create({ paragraphs: ["clean install DOCX"] });
    const docx = await DocumentFile.exportDocx(document, { codec: "open-chestnut" });
    const importedDocument = await DocumentFile.importDocx(docx, { codec: "open-chestnut" });
    if (docx.bytes[0] !== 0x50 || docx.bytes[1] !== 0x4b) process.exit(3);
    if (importedDocument.blocks[0].text !== "clean install DOCX") process.exit(4);
    const presentation = Presentation.create();
    const packagedSlide = presentation.slides.add({ name: "Packaged" });
    packagedSlide.shapes.add({ name: "Title", text: [{ bulletCharacter: "•", bulletFont: "Georgia", bulletColor: "#2563EB", bulletSizePercent: 1.25, runs: [{ text: "clean install PPTX", link: { uri: "https://example.com/packaged", tooltip: "Packaged link" } }] }], position: { left: 40, top: 40, width: 640, height: 80 }, transform: { rotationDegrees: 4, flipHorizontal: false, flipVertical: true } });
    packagedSlide.images.add({ name: "Packaged image", alt: "clean install image", position: { left: 40, top: 160, width: 64, height: 64 }, fit: "stretch", transform: { rotationDegrees: -2, flipHorizontal: false }, dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=" });
    const pptx = await PresentationFile.exportPptx(presentation, { codec: "open-chestnut" });
    const importedPresentation = await PresentationFile.importPptx(pptx, { codec: "open-chestnut" });
    if (pptx.bytes[0] !== 0x50 || pptx.bytes[1] !== 0x4b) process.exit(5);
    if (importedPresentation.slides.getItem(0).shapes.items[0].text.value !== "clean install PPTX") process.exit(6);
    if (importedPresentation.slides.getItem(0).shapes.items[0].transform?.rotationDegrees !== 4 || importedPresentation.slides.getItem(0).shapes.items[0].transform?.flipHorizontal !== false || importedPresentation.slides.getItem(0).shapes.items[0].transform?.flipVertical !== true) process.exit(25);
    if (importedPresentation.slides.getItem(0).images.items[0]?.alt !== "clean install image" || importedPresentation.slides.getItem(0).images.items[0]?.transform?.rotationDegrees !== -2 || importedPresentation.slides.getItem(0).images.items[0]?.transform?.flipHorizontal !== false) process.exit(26);
    const marker = importedPresentation.slides.getItem(0).shapes.items[0].text.paragraphs[0];
    if (marker.bulletCharacter !== "•" || marker.bulletFont !== "Georgia" || marker.bulletColor.toLowerCase() !== "#2563eb" || marker.bulletSizePercent !== 1.25) process.exit(7);
    if (marker.runs[0].link?.uri !== "https://example.com/packaged" || marker.runs[0].link?.tooltip !== "Packaged link") process.exit(8);
    const template = Presentation.create({
      master: {
        id: "master/packaged",
        name: "Packaged Master",
        placeholders: [{ type: "title", idx: 0, name: "Packaged Prompt", position: { left: 40, top: 30, width: 640, height: 80 }, transform: { rotationDegrees: 1, flipHorizontal: true, flipVertical: false }, text: "before" }],
      },
    });
    template.master.placeholders[0].position = undefined;
    template.master.placeholders[0].transform = undefined;
    const packagedLayout = template.layouts.add({ id: "layout/packaged", name: "Packaged Layout", type: "blank", masterId: "master/packaged", placeholders: [{ type: "body", idx: 2, name: "Packaged Layout Prompt", position: { left: 50, top: 180, width: 600, height: 90 }, transform: { rotationDegrees: 3, flipHorizontal: false }, text: "layout before" }] });
    template.slides.add({ name: "Template", layoutId: "layout/packaged" }).applyLayout(packagedLayout);
    const templateSource = await PresentationFile.exportPptx(template, { codec: "javascript" });
    const inheritedTemplateZip = await JSZip.loadAsync(templateSource.bytes);
    const inheritedTemplateXml = await inheritedTemplateZip.file("ppt/slides/slide1.xml").async("text");
    const inheritedPlaceholderIndex = inheritedTemplateXml.indexOf('idx="2"');
    const inheritedShapeStart = inheritedTemplateXml.lastIndexOf("<p:sp", inheritedPlaceholderIndex);
    const inheritedFrameStart = inheritedTemplateXml.indexOf("<a:xfrm", inheritedShapeStart);
    const inheritedFrameEnd = inheritedTemplateXml.indexOf("</a:xfrm>", inheritedFrameStart) + "</a:xfrm>".length;
    if (inheritedPlaceholderIndex < 0 || inheritedShapeStart < 0 || inheritedFrameStart < 0 || inheritedFrameEnd < "</a:xfrm>".length) process.exit(24);
    inheritedTemplateZip.file("ppt/slides/slide1.xml", inheritedTemplateXml.slice(0, inheritedFrameStart) + inheritedTemplateXml.slice(inheritedFrameEnd));
    const inheritedTemplate = await PresentationFile.importPptx(await inheritedTemplateZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { codec: "open-chestnut" });
    const inheritedTemplateShape = inheritedTemplate.slides.items[0].shapes.items.find((shape) => shape.placeholder?.idx === 2);
    if (inheritedTemplateShape.position.left !== 50 || inheritedTemplateShape.position.top !== 180 || inheritedTemplateShape.position.width !== 600 || inheritedTemplateShape.position.height !== 90 || inheritedTemplateShape.transform?.rotationDegrees !== 3 || inheritedTemplateShape.transform?.flipHorizontal !== false || inheritedTemplateShape.placeholder?.geometrySource !== "layout") process.exit(22);
    const inheritedTemplateRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(inheritedTemplate, { codec: "open-chestnut" }), { codec: "open-chestnut" });
    const inheritedTemplateRoundTripShape = inheritedTemplateRoundTrip.slides.items[0].shapes.items.find((shape) => shape.placeholder?.idx === 2);
    if (inheritedTemplateRoundTripShape?.position.left !== 50 || inheritedTemplateRoundTripShape.transform?.rotationDegrees !== 3 || inheritedTemplateRoundTripShape.transform?.flipHorizontal !== false || inheritedTemplateRoundTripShape.placeholder?.geometrySource !== "layout") process.exit(23);
    const importedTemplate = await PresentationFile.importPptx(templateSource, { codec: "open-chestnut" });
    const packagedPlaceholder = importedTemplate.master.placeholders[0];
    const packagedLayoutPlaceholder = importedTemplate.layouts.items[0].placeholders[0];
    if (packagedPlaceholder.position !== undefined || packagedPlaceholder.transform !== undefined || packagedLayoutPlaceholder.transform?.flipHorizontal !== false) process.exit(20);
    packagedPlaceholder.position = { left: 48, top: 36, width: 620, height: 72 };
    packagedPlaceholder.transform = { rotationDegrees: -15, flipHorizontal: false, flipVertical: true };
    packagedLayoutPlaceholder.position = undefined;
    packagedLayoutPlaceholder.transform = undefined;
    const editedTemplate = await PresentationFile.exportPptx(importedTemplate, { codec: "open-chestnut" });
    const roundTripTemplate = await PresentationFile.importPptx(editedTemplate, { codec: "open-chestnut" });
    if (roundTripTemplate.master.placeholders[0].position.left !== 48 || roundTripTemplate.master.placeholders[0].transform?.rotationDegrees !== -15 || roundTripTemplate.master.placeholders[0].transform?.flipHorizontal !== false || roundTripTemplate.master.placeholders[0].transform?.flipVertical !== true || roundTripTemplate.layouts.items[0].placeholders[0].position !== undefined) process.exit(21);
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
