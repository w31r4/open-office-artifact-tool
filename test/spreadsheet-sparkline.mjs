import assert from "node:assert/strict";
import JSZip from "jszip";

import { SpreadsheetFile, Workbook } from "../src/index.mjs";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Forecast Data");
sheet.getRange("A1:F5").values = [
  [1, 2, 3, 4, 5, 6],
  [6, 4, 5, 2, 3, 1],
  [-3, 2, -1, 4, -2, 5],
  [1, 2, 3, 4, 5, 6],
  [6, 5, 4, 3, 2, 1],
];

const line = sheet.sparklineGroups.add({
  type: "line",
  targetRange: "G1:G2",
  sourceData: "A1:F2",
  dateAxisRange: "A1:F1",
  seriesColor: "#0EA5E9",
  negativeColor: "#DC2626",
  axisColor: "#64748B",
  markersColor: "#334155",
  firstMarkerColor: "#16A34A",
  lastMarkerColor: "#9333EA",
  highMarkerColor: "#22C55E",
  lowMarkerColor: "#EF4444",
  markers: { show: true, high: true, low: true, first: true, last: true, negative: true },
  axis: { showAxis: true, manualMin: -5, manualMax: 10, minMode: 2, maxMode: 2 },
  lineWeight: 2.25,
  displayEmptyCellsAs: 2,
  displayHidden: true,
});
const columns = sheet.getRange("H1:H2").sparklines.add("column", sheet.getRange("A4:F5"), {
  seriesColor: { theme: 4, tint: 0.1 },
  negativeColor: "#B91C1C",
});
const stacked = sheet.sparklineGroups.add({ type: "stacked", locationRange: "I1", sourceData: "A3:F3", displayEmptyCellsAs: "zero" });
const horizontal = sheet.sparklineGroups.add({
  type: "line",
  targetRange: "J1:K1",
  sourceData: "A1:B5",
  axis: { minMode: "group", maxMode: "group", rightToLeft: true },
});

assert.equal(line.sparklineCount, 2);
assert.equal(line.locationRange.address, "G1:G2");
assert.equal(columns.sparklineCount, 2);
assert.equal(stacked.sparklineCount, 1);
assert.equal(horizontal.sparklineCount, 2);
assert.equal(sheet.sparklines, sheet.sparklineGroups);
assert.equal(sheet.sparklineGroups.getAll().length, 4);

const preview = await workbook.render({ format: "svg", sheetName: sheet.name, range: "A1:K5" });
const svg = await preview.text();
assert.ok((svg.match(/<polyline\b/g) || []).length >= 2, "Each line sparkline target should render independently.");
assert.ok((svg.match(/<circle\b/g) || []).length >= 2, "Configured line markers should render.");

const file = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
const zip = await JSZip.loadAsync(file.bytes);
const worksheetXml = await zip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(worksheetXml, /uri="\{05C60535-1F16-4FD2-B633-F4F36F0B64E0\}"/i);
assert.equal((worksheetXml.match(/<x14:sparklineGroup\b/g) || []).length, 4);
assert.equal((worksheetXml.match(/<x14:sparkline\b/g) || []).length, 7);
assert.match(worksheetXml, /Forecast Data(?:'?)!A1:F1/);
assert.match(worksheetXml, /<xne:sqref[^>]*>G2<\/xne:sqref>/);

const imported = await SpreadsheetFile.importXlsx(file);
const importedSheet = imported.worksheets.getItem("Forecast Data");
assert.equal(importedSheet.sparklineGroups.items.length, 4);
assert.deepEqual(importedSheet.sparklineGroups.items.map((group) => group.type), ["line", "column", "stacked", "line"]);
assert.equal(importedSheet.sparklineGroups.items[0].targetRange.address, "G1:G2");
assert.equal(importedSheet.sparklineGroups.items[0].sourceData.sheetName, "Forecast Data");
assert.equal(importedSheet.sparklineGroups.items[0].sourceData.address, "A1:F2");
assert.equal(importedSheet.sparklineGroups.items[0].dateAxisRange.address, "A1:F1");
assert.deepEqual(importedSheet.sparklineGroups.items[0].markers, { show: true, high: true, low: true, first: true, last: true, negative: true });
assert.equal(importedSheet.sparklineGroups.items[2].displayEmptyCellsAs, 3);
assert.equal(importedSheet.sparklineGroups.items[3].targetRange.address, "J1:K1");
assert.equal(importedSheet.sparklineGroups.items[3].sourceData.address, "A1:B5");
assert.equal(importedSheet.sparklineGroups.items[3].axis.minMode, 1);
assert.equal(importedSheet.sparklineGroups.items[3].axis.rightToLeft, true);

importedSheet.sparklineGroups.items[0].seriesColor = "#F97316";
importedSheet.sparklineGroups.items[0].axis.manualMax = 12;
const edited = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
const reimported = await SpreadsheetFile.importXlsx(edited);
assert.equal(reimported.worksheets.getItem("Forecast Data").sparklineGroups.items[0].seriesColor, "#F97316");
assert.equal(reimported.worksheets.getItem("Forecast Data").sparklineGroups.items[0].axis.manualMax, 12);

const invalid = Workbook.create();
const invalidSheet = invalid.worksheets.add("Invalid");
invalidSheet.sparklineGroups.add({ targetRange: "D1:E2", sourceData: "A1:C4" });
await assert.rejects(() => SpreadsheetFile.exportXlsx(invalid, { recalculate: false }), /one-dimensional target range/i);

const missingSource = Workbook.create();
const missingSheet = missingSource.worksheets.add("Only Sheet");
missingSheet.sparklineGroups.add({ targetRange: "D1", sourceData: "Missing!A1:C1" });
await assert.rejects(() => SpreadsheetFile.exportXlsx(missingSource, { recalculate: false }), /sourceData.*worksheet|Missing/i);

const importedForTopology = await SpreadsheetFile.importXlsx(file);
importedForTopology.worksheets.getItem("Forecast Data").sparklineGroups.items.pop();
await assert.rejects(() => SpreadsheetFile.exportXlsx(importedForTopology, { recalculate: false }), /cannot remove or reorder.*source-bound/i);

console.log("spreadsheet sparkline tests passed");
