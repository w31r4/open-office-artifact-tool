import assert from "node:assert/strict";

import { SpreadsheetFile, Workbook } from "../src/index.mjs";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet 1");
workbook.worksheets.add("Other Sheet").getRange("A1").values = [[9]];
sheet.getRange("A1:C4").values = [
  [1, 2, null],
  [3, 4, null],
  [null, null, null],
  [7, null, 9],
];

const r1c1 = sheet.getRange("B2:C3");
r1c1.formulasR1C1 = [
  ["=RC[-1]*2", "=R[-1]C+1"],
  ["=R[-1]C", "=R1C1+R1C1"],
];
assert.deepEqual(r1c1.formulas, [
  ["=A2*2", "=C1+1"],
  ["=B2", "=$A$1+$A$1"],
]);
assert.deepEqual(r1c1.formulasR1C1, [
  ["=RC[-1]*2", "=R[-1]C+1"],
  ["=R[-1]C", "=R1C1+R1C1"],
]);
assert.deepEqual(r1c1.displayFormulas, r1c1.formulas);
assert.deepEqual(r1c1.formulaInfos[0][0], {
  kind: "stored",
  formula: "=A2*2",
  display: "=A2*2",
  isEditable: true,
});
assert.equal(sheet.getRange("A1").formulas[0][0], "");
sheet.getRange("D2").formulas = [["=\"A1\"&'Other Sheet'!$A$1"]];
assert.equal(sheet.getRange("D2").formulasR1C1[0][0], "=\"A1\"&'Other Sheet'!R1C1");
sheet.getRange("D3").formulasR1C1 = [["='Other Sheet'!R1C1+RC[-1]"]];
assert.equal(sheet.getRange("D3").formulas[0][0], "='Other Sheet'!$A$1+C3");
sheet.getRange("D4").formulas = [["=SUM($A1,B$2,$C$3,A1:B2,\"A1\")"]];
assert.equal(sheet.getRange("D4").formulasR1C1[0][0], "=SUM(R[-3]C1,R2C[-2],R3C3,R[-3]C[-3]:R[-2]C[-2],\"A1\")");
sheet.getRange("D4").formulasR1C1 = [[sheet.getRange("D4").formulasR1C1[0][0]]];
assert.equal(sheet.getRange("D4").formulas[0][0], "=SUM($A1,B$2,$C$3,A1:B2,\"A1\")");
assert.throws(() => { sheet.getRange("A1").formulasR1C1 = [["=R[-1]C"]]; }, /outside worksheet/);

sheet.getRange("T2").formulas = [["=SEQUENCE(2,2,10,1)"]];
const spill = sheet.getRange("T2:U3");
assert.deepEqual(spill.values, [[10, 11], [12, 13]]);
assert.deepEqual(spill.displayFormulas, [
  ["=SEQUENCE(2,2,10,1)", "=SEQUENCE(2,2,10,1)"],
  ["=SEQUENCE(2,2,10,1)", "=SEQUENCE(2,2,10,1)"],
]);
assert.deepEqual(spill.formulaInfos[0][1], {
  kind: "projected",
  source: "spill",
  display: "=SEQUENCE(2,2,10,1)",
  anchor: "Sheet 1!T2",
  ref: "Sheet 1!T2:U3",
  isEditable: false,
});

const written = sheet.getRange("E2").write([[10, 20], [30, "=E3+10"]]);
assert.equal(written.address, "E2:F3");
assert.equal(written.rowIndex, 1);
assert.equal(written.columnIndex, 4);
assert.equal(written.rowCount, 2);
assert.equal(written.columnCount, 2);
assert.deepEqual(written.values, [[10, 20], [30, 40]]);
assert.deepEqual(written.formulas, [["", ""], ["", "=E3+10"]]);
assert.equal(sheet.getRange("H2").writeValues([50, 60, 70]), undefined);
assert.deepEqual(sheet.getRange("H2:J2").values, [[50, 60, 70]]);
assert.deepEqual(sheet.getRange("E5").write({ formulas: [["=E2*3", "=F2*4"]] }).formulas, [["=E2*3", "=F2*4"]]);
assert.deepEqual(sheet.getRange("E6").write({ formulasR1C1: [["=R[-1]C+1", "=R[-1]C+1"]] }).formulas, [["=E5+1", "=F5+1"]]);
assert.throws(
  () => sheet.getRange("E7").write({ values: [[1]], formulas: [["=1"]] }),
  /expects exactly one field/,
);

const copySource = sheet.getRange("W4:X4");
copySource.values = [[5, null]];
copySource.getCell(0, 1).formulas = [["=W4*2"]];
copySource.format = { fill: "#00FF00", font: { bold: true } };
const copied = sheet.getRange("W6:X6");
assert.equal(copied.copyFrom(copySource, "all"), undefined);
assert.deepEqual(copied.values, [[5, 10]]);
assert.deepEqual(copied.formulas, [["", "=W6*2"]]);
assert.equal(copied.format.fill, "#00FF00");
const tiled = sheet.getRange("Z7:AC8");
tiled.copyFrom(copySource, "all");
assert.deepEqual(tiled.values, [[5, 10, 5, 10], [5, 10, 5, 10]]);
assert.deepEqual(tiled.formulas, [
  ["", "=Z7*2", "", "=AB7*2"],
  ["", "=Z8*2", "", "=AB8*2"],
]);
assert.equal(copySource.copyTo(sheet.getRange("W10:X10"), "values"), undefined);
assert.deepEqual(sheet.getRange("W10:X10").formulas, [["", ""]]);
assert.throws(() => sheet.getRange("Z10:AB10").copyFrom(sheet.getRange("E2:F3")), /Source is 2x2, destination is 1x3/);

const clearRange = sheet.getRange("Q2:R2");
clearRange.values = [[1, 2]];
clearRange.format = { fill: "#FF0000", font: { bold: true } };
assert.equal(clearRange.clear({ applyTo: "contents" }), undefined);
assert.deepEqual(clearRange.values, [[null, null]]);
assert.equal(clearRange.format.fill, "#FF0000");
clearRange.values = [[3, 4]];
clearRange.clear({ applyTo: "formats" });
assert.deepEqual(clearRange.values, [[3, 4]]);
assert.equal(clearRange.format.fill, undefined);
clearRange.clear({ applyTo: "all" });
assert.deepEqual(clearRange.values, [[null, null]]);
assert.throws(() => clearRange.clear({ applyTo: "validation" }), /contents, formats, or all/);

const navigation = sheet.getRange("B2:C3");
assert.equal(navigation.offset(1, 2).address, "D3:E4");
assert.equal(navigation.resize(3, 4).address, "B2:E4");
assert.equal(navigation.getRow(1).address, "B3:C3");
assert.equal(navigation.getColumn(1).address, "C2:C3");
assert.equal(sheet.getRange("B2:D4").getRangeByIndexes(1, 1, 2, 2).address, "C3:D4");
assert.equal(navigation.getCell(1, 1).address, "C3");
assert.equal(sheet.getRange("A1").getCurrentRegion().address, "A1:F6");
assert.throws(() => sheet.getRange("A1").offset(-1, 0), /outside the worksheet/);
assert.throws(() => navigation.getRangeByIndexes(1, 1, 2, 2), /outside the bounds/);

sheet.getRange("M20").format = { fill: "#0000FF" };
assert.equal(sheet.getUsedRange(false).address.endsWith("M20"), false);
assert.equal(sheet.getUsedRange(false).rowCount, 20);
assert.ok(sheet.getUsedRange(false).columnCount >= 29);
assert.ok(sheet.getUsedRange(true).rowCount < 20);
const numberFormats = sheet.getRange("A22:A24");
numberFormats.setNumberFormat([["0"], ["0.00"], ["0.0%"]]);
assert.equal(sheet.getRange("A22").format.numberFormat, "0");
assert.equal(sheet.getRange("A23").format.numberFormat, "0.00");
assert.equal(sheet.getRange("A24").format.numberFormat, "0.0%");

const conditionalWorkbook = Workbook.create();
const conditionalSheet = conditionalWorkbook.worksheets.add("Status");
conditionalSheet.getRange("D2:D3").values = [["Review"], ["Done"]];
const containsText = conditionalSheet.getRange("D2:D3").conditionalFormats.add("containsText", {
  text: "Review",
  format: { fill: "#FEF3C7" },
});
assert.equal(containsText.formula, 'NOT(ISERROR(SEARCH("Review",D2)))');
assert.throws(
  () => conditionalSheet.getRange("D2:D3").conditionalFormats.add("containsText", { format: { fill: "#FEF3C7" } }),
  /requires text/,
);
const conditionalRoundTrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(conditionalWorkbook));
assert.equal(conditionalRoundTrip.worksheets.getItem("Status").conditionalFormattings.items[0].formula, 'NOT(ISERROR(SEARCH("Review",D2)))');

const chartWorkbook = Workbook.create();
const chartData = chartWorkbook.worksheets.add("Chart Data");
chartData.getRange("D2:E5").values = [
  ["Aug 2026", 1250],
  ["Sep 2026", 1325],
  ["Oct 2026", 1404.5],
  ["Nov 2026", 1488.77],
];
chartData.getRange("A2:A5").formulas = [["=D2"], ["=D3"], ["=D4"], ["=D5"]];
chartData.getRange("B2:B5").formulas = [["=E2"], ["=E3"], ["=E4"], ["=E5"]];
assert.deepEqual(chartData.getRange("A2:A5").values, [["Aug 2026"], ["Sep 2026"], ["Oct 2026"], ["Nov 2026"]]);
const dashboard = chartWorkbook.worksheets.add("Dashboard");
const directSeriesChart = dashboard.charts.add("line", { title: "Formula-backed trend", hasLegend: true });
const directSeries = directSeriesChart.series.add("Revenue");
directSeries.categoryFormula = "'Chart Data'!$A$2:$A$5";
directSeries.formula = "'Chart Data'!$B$2:$B$5";
directSeriesChart.setPosition("A1", "G14");
directSeriesChart.xAxis = { textStyle: { fontSize: 10 } };
const chartPreview = await chartWorkbook.render({ sheetName: "Dashboard", autoCrop: "all", format: "svg" });
assert.match(await chartPreview.text(), /<polyline[^>]+data-series-index="0"/);
assert.match(await chartPreview.text(), /Aug 2026/);
const chartRoundTrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(chartWorkbook));
const importedDirectChart = chartRoundTrip.worksheets.getItem("Dashboard").charts.items[0];
assert.deepEqual(importedDirectChart.categories, ["Aug 2026", "Sep 2026", "Oct 2026", "Nov 2026"]);
assert.deepEqual(importedDirectChart.series.items[0].values, [1250, 1325, 1404.5, 1488.77]);

console.log("spreadsheet range compatibility tests passed");
