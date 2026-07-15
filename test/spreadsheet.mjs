import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { FileBlob, SpreadsheetFile, Workbook } from "../src/index.mjs";
import { exportXlsxWithOpenChestnut, importXlsxWithOpenChestnut } from "../src/codecs/open-chestnut.mjs";
import { parsePivotCacheDefinition, parsePivotTableDefinition } from "../src/spreadsheet/ooxml-pivots.mjs";
import { evaluatePivotFormula } from "../src/spreadsheet/pivot-formulas.mjs";
import { pivotItemVisible } from "../src/spreadsheet/pivot-filters.mjs";
import { normalizePivotGroupFields, pivotGroupValue } from "../src/spreadsheet/pivot-groups.mjs";

assert.equal(evaluatePivotFormula("=2+3*4", {}, []), 14);
assert.equal(evaluatePivotFormula("=('Revenue'-'Cost')/2", { Revenue: 15, Cost: 9 }), 3);
assert.equal(evaluatePivotFormula("='Owner''s Revenue'*10%", { "Owner's Revenue": 80 }), 8);
assert.equal(evaluatePivotFormula("=Revenue/0", { Revenue: 10 }), "#DIV/0!");
assert.equal(evaluatePivotFormula("=-(Revenue/0)%", { Revenue: 10 }), "#DIV/0!");
assert.equal(evaluatePivotFormula("=SUM(Revenue,Cost*2)", { Revenue: 15, Cost: 4 }), 23);
assert.equal(evaluatePivotFormula("=AVERAGE(MIN(Revenue,Cost),MAX(Revenue,Cost),ABS(-3))", { Revenue: 15, Cost: 9 }), 9);
assert.equal(evaluatePivotFormula("=ROUND(-1.25,1)", {}), -1.3);
assert.equal(evaluatePivotFormula("=ROUND(1.005,2)", {}), 1.01);
assert.equal(Object.is(evaluatePivotFormula("=ROUND(-0.1,0)", {}), -0), false);
assert.equal(evaluatePivotFormula("=ROUND(Revenue/0,2)", { Revenue: 10 }), "#DIV/0!");
assert.equal(evaluatePivotFormula("=IF(Cost=0,0,Revenue/Cost)", { Revenue: 10, Cost: 0 }), 0);
assert.equal(evaluatePivotFormula("=IF(Cost<>0,Revenue/Cost,0)", { Revenue: 10, Cost: 4 }), 2.5);
assert.equal(evaluatePivotFormula("=IF(TRUE,1,Revenue/0)", { Revenue: 10 }), 1);
assert.equal(evaluatePivotFormula("=IF(FALSE,Revenue/0)", { Revenue: 10 }), false);
assert.equal(evaluatePivotFormula('=IF(Revenue>=Cost,"profit","loss")', { Revenue: 10, Cost: 4 }), "profit");
assert.equal(evaluatePivotFormula('=IFERROR(Revenue/Cost,"n/a")', { Revenue: 10, Cost: 0 }), "n/a");
assert.equal(evaluatePivotFormula("=IFERROR(Revenue/Cost,Cost/0)", { Revenue: 10, Cost: 4 }), 2.5);
assert.equal(evaluatePivotFormula("=IFERROR(1E308*1E308,0)", {}), 0);
assert.equal(evaluatePivotFormula("=IFERROR(SUM(1E308,1E308),0)", {}), 0);
assert.equal(evaluatePivotFormula('=IFERROR(Revenue/Cost,"say ""n/a""")', { Revenue: 10, Cost: 0 }), 'say "n/a"');
assert.equal(evaluatePivotFormula('=IFERROR("#review","fallback")', {}), "#review");
assert.equal(evaluatePivotFormula('=IFERROR("#N/A","fallback")', {}), "#N/A");
assert.equal(evaluatePivotFormula('=IFERROR("#"&"N/A","fallback")', {}), "#N/A");
assert.equal(evaluatePivotFormula("=IF(Revenue/0,1,2)", { Revenue: 10 }), "#DIV/0!");
assert.equal(evaluatePivotFormula("=PRODUCT(Revenue,Cost,2)", { Revenue: 3, Cost: 4 }), 24);
assert.equal(evaluatePivotFormula("=PRODUCT(1E308,1E308)", {}), "#NUM!");
assert.equal(evaluatePivotFormula("=POWER(3,2)", {}), 9);
assert.equal(evaluatePivotFormula("=POWER(0,-1)", {}), "#DIV/0!");
assert.equal(evaluatePivotFormula("=POWER(-1,0.5)", {}), "#NUM!");
assert.equal(evaluatePivotFormula("=SQRT(81)", {}), 9);
assert.equal(evaluatePivotFormula("=SQRT(-1)", {}), "#NUM!");
assert.equal(evaluatePivotFormula("=MOD(-3,2)", {}), 1);
assert.equal(evaluatePivotFormula("=MOD(3,-2)", {}), -1);
assert.equal(evaluatePivotFormula("=MOD(3,0)", {}), "#DIV/0!");
assert.equal(evaluatePivotFormula("=SIGN(-4)+SIGN(0)+SIGN(5)", {}), 0);
assert.equal(evaluatePivotFormula("=INT(-1.2)", {}), -2);
assert.equal(evaluatePivotFormula("=AND(Revenue>Cost,Cost>0)", { Revenue: 10, Cost: 4 }), true);
assert.equal(evaluatePivotFormula("=OR(FALSE,0,Revenue>Cost)", { Revenue: 10, Cost: 4 }), true);
assert.equal(evaluatePivotFormula("=NOT(Revenue=Cost)", { Revenue: 10, Cost: 4 }), true);
assert.equal(evaluatePivotFormula('=AND(TRUE,"invalid")', {}), "#VALUE!");
assert.equal(evaluatePivotFormula("=NA()", {}), "#N/A");
assert.equal(evaluatePivotFormula("=IFNA(NA(),7)", {}), 7);
assert.equal(evaluatePivotFormula("=IFNA(1,Revenue/0)", { Revenue: 10 }), 1);
assert.equal(evaluatePivotFormula("=IFNA(Revenue/0,7)", { Revenue: 10 }), "#DIV/0!");
assert.equal(evaluatePivotFormula("=ISERROR(Revenue/0)", { Revenue: 10 }), true);
assert.equal(evaluatePivotFormula("=ISERROR(NA())", {}), true);
assert.equal(evaluatePivotFormula('=ISERROR("#N/A")', {}), false);
assert.equal(evaluatePivotFormula("=ISERROR(Revenue)", { Revenue: 10 }), false);
assert.equal(evaluatePivotFormula("=ISNUMBER(Revenue)", { Revenue: 10 }), true);
assert.equal(evaluatePivotFormula('=ISNUMBER("10")', {}), false);
assert.equal(evaluatePivotFormula('=ISTEXT("#review")', {}), true);
assert.equal(evaluatePivotFormula('=ISTEXT("#N/A")', {}), true);
assert.equal(evaluatePivotFormula("=ISTEXT(Revenue/0)", { Revenue: 10 }), false);
assert.throws(() => evaluatePivotFormula("=ABS(Revenue,Cost)", { Revenue: 1, Cost: 2 }), /exactly one argument/);
assert.throws(() => evaluatePivotFormula("=SUM()", {}), /at least one argument/);
assert.throws(() => evaluatePivotFormula(`=SUM(${Array(33).fill("1").join(",")})`, {}), /exceeds 32 arguments/);
assert.throws(() => evaluatePivotFormula("=ROUND(1,16)", {}), /integer from -15 to 15/);
assert.throws(() => evaluatePivotFormula("=IF(Revenue>0)", { Revenue: 1 }), /requires 2 or 3 arguments/);
assert.throws(() => evaluatePivotFormula("=IFERROR(Revenue)", { Revenue: 1 }), /exactly two arguments/);
assert.throws(() => evaluatePivotFormula("=NOT(TRUE,FALSE)", {}), /exactly one argument/);
assert.throws(() => evaluatePivotFormula("=AND()", {}), /at least one argument/);
assert.throws(() => evaluatePivotFormula("=NA(1)", {}), /exactly 0 arguments/);
assert.throws(() => evaluatePivotFormula('=IFERROR(Revenue,"unterminated)', { Revenue: 1 }), /unterminated string constant/);
assert.throws(() => evaluatePivotFormula("=LOG(Revenue)", { Revenue: 1 }), /unsupported function LOG/);
assert.equal(pivotItemVisible([{ field: "Date", type: "dateEqual", value1: "1904-01-01" }], "Date", 0, "1904"), true);
assert.equal(pivotItemVisible([{ field: "Date", type: "dateEqual", value1: "1900-03-01" }], "Date", 61, "1900"), true);
assert.equal(pivotItemVisible([{ field: "Date", type: "dateEqual", value1: "1900-02-29" }], "Date", 60, "1900"), true);
for (const [type, current, expected] of [
  ["yesterday", "2026-07-11", true], ["today", "2026-07-12", true], ["tomorrow", "2026-07-13", true],
  ["lastWeek", "2026-07-01", true], ["thisWeek", "2026-07-06", true], ["nextWeek", "2026-07-19", true],
  ["lastMonth", "2026-06-30", true], ["thisMonth", "2026-07-31", true], ["nextMonth", "2026-08-01", true],
  ["lastQuarter", "2026-06-30", true], ["thisQuarter", "2026-09-30", true], ["nextQuarter", "2026-10-01", true],
  ["lastYear", "2025-12-31", true], ["thisYear", "2026-01-01", true], ["nextYear", "2027-01-01", true],
  ["yearToDate", "2026-07-12", true], ["yearToDate", "2026-07-13", false], ["thisWeek", "2026-07-05", false],
]) assert.equal(pivotItemVisible([{ field: "Date", type, asOf: "2026-07-12" }], "Date", current), expected, `${type} ${current}`);
assert.equal(pivotGroupValue({ groupBy: "years" }, 0, "1904"), "1904");
assert.equal(pivotGroupValue({ groupBy: "months" }, 60, "1900"), "Feb");
assert.equal(pivotGroupValue({ groupBy: "quarters" }, new Date("2026-07-15T00:00:00Z")), "Q3");
assert.equal(pivotGroupValue({ groupBy: "range", range: { startNum: 0, endNum: 29, groupInterval: 10 } }, 18), "10-19");
assert.equal(pivotGroupValue({ groupBy: "range", range: { startNum: 0, endNum: 29, groupInterval: 10 } }, 30), ">29");
assert.equal(pivotGroupValue({ groupBy: "discrete", groups: [{ name: "Coasts", items: ["East", "West"] }] }, "West"), "Coasts");
assert.equal(pivotGroupValue({ groupBy: "discrete", groups: [{ name: "Coasts", items: ["East", "West"] }] }, "Central"), "Central");
assert.equal(pivotGroupValue({ groupBy: "hours", range: { groupInterval: 1 } }, 0.5, "1904"), "12");
assert.equal(pivotGroupValue({ groupBy: "minutes", range: { groupInterval: 15 } }, "2026-07-12T09:37:45Z"), "30-44");
assert.equal(pivotGroupValue({ groupBy: "seconds", range: { groupInterval: 10 } }, "2026-07-12T09:37:45Z"), "40-49");
assert.equal(pivotGroupValue({ groupBy: "days", range: { groupInterval: 7 } }, "2026-07-12"), "8-14");
assert.deepEqual(normalizePivotGroupFields([{ name: "Auto Band", sourceField: "Score", groupBy: "range", range: { groupInterval: 10 } }], ["Score"], false, { Score: [5, 25] })[0].range, { autoStart: true, autoEnd: true, startNum: 5, endNum: 25, groupInterval: 10 });
assert.deepEqual(normalizePivotGroupFields([{ name: "Auto Band", sourceField: "Score", groupBy: "range", range: { groupInterval: 10 } }], ["Score"], false, { Score: [null, "", 5, 25] })[0].range, { autoStart: true, autoEnd: true, startNum: 5, endNum: 25, groupInterval: 10 });
assert.equal(pivotGroupValue({ groupBy: "range", range: { startNum: 0, endNum: 29, groupInterval: 10 } }, ""), null);
assert.equal(pivotGroupValue({ groupBy: "years" }, Number.MAX_VALUE), null);
for (const [type, current, expected] of [
  ["dateEqual", "2026-02-01", true],
  ["dateNotEqual", "2026-02-01", false],
  ["dateOlderThan", "2026-01-31", true],
  ["dateOlderThan", "2026-02-01", false],
  ["dateOlderThanOrEqual", "2026-02-01", true],
  ["dateNewerThan", "2026-02-02", true],
  ["dateNewerThan", "2026-02-01", false],
  ["dateNewerThanOrEqual", "2026-02-01", true],
  ["dateBetween", "2026-02-28", true],
  ["dateNotBetween", "2026-02-28", false],
  ["dateNotBetween", "2026-04-01", true],
]) {
  assert.equal(pivotItemVisible([{ field: "Date", type, value1: "2026-02-01", value2: "2026-03-31" }], "Date", current), expected, `${type} ${current}`);
}
assert.equal(pivotItemVisible([{ field: "Date", type: "dateEqual", value1: "2026-02-01" }], "Date", "not-a-date"), false);
const groupedCacheContract = parsePivotCacheDefinition(`<pivotCacheDefinition><cacheSource><worksheetSource ref="A1:A2" sheet="S"/></cacheSource><cacheFields count="2"><cacheField name="Source"><sharedItems count="1"><s v="A"/></sharedItems></cacheField><cacheField name="Grouped" databaseField="0"><fieldGroup base="0"/></cacheField></cacheFields></pivotCacheDefinition>`);
assert.deepEqual(groupedCacheContract.sourceFields, ["Source"]);
assert.deepEqual(groupedCacheContract.calculatedFields, []);
const customDateFilterContract = parsePivotTableDefinition(`<pivotTableDefinition name="DatePivot" cacheId="3"><location ref="D4:E6"/><pivotFields count="2"><pivotField/><pivotField axis="axisRow"><items count="2"><item x="0" h="1"/><item t="default"/></items></pivotField></pivotFields><rowFields count="1"><field x="1"/></rowFields><filters count="1"><filter fld="1" type="dateBetween" id="1"><autoFilter ref="A1"><filterColumn colId="0"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="2026-02-01"/><customFilter operator="lessThanOrEqual" val="2026-03-31"/></customFilters></filterColumn></autoFilter></filter></filters></pivotTableDefinition>`, {
  fields: ["Revenue", "OrderDate"],
  items: [[], ["2026-01-31"]],
});
assert.deepEqual(customDateFilterContract.rowFields, ["OrderDate"]);
assert.deepEqual(customDateFilterContract.filters, [{ field: "OrderDate", type: "dateBetween", value1: "2026-02-01", value2: "2026-03-31", useWholeDay: true }]);
const mutableDateSystemBook = Workbook.create();
const mutableDateSystemSheet = mutableDateSystemBook.worksheets.add("Dates");
mutableDateSystemSheet.getRange("A1:B2").values = [["Date", "Amount"], [0, 10]];
const mutableDateSystemPivot = mutableDateSystemSheet.pivotTables.add({ sourceRange: "A1:B2", targetRange: "D1", rowFields: ["Date"], valueFields: [{ field: "Amount" }], filters: [{ field: "Date", type: "dateEqual", value1: "1904-01-01" }] });
assert.deepEqual(mutableDateSystemPivot.computedValues(), [["Date", "sum of Amount"]]);
mutableDateSystemBook.setDateSystem("1904");
assert.deepEqual(mutableDateSystemPivot.computedValues(), [["Date", "sum of Amount"], [0, 10]]);

const calculationBook = Workbook.create({ calculation: { mode: "autoNoTable", calculateOnSave: false, fullCalculationOnLoad: true, forceFullCalculation: true, iteration: { enabled: true, maxIterations: 100, maxChange: 0.001 }, fullPrecision: false } });
calculationBook.worksheets.add("Calculation").getRange("A1").formulas = [["=1+1"]];
assert.deepEqual(calculationBook.calculation, {
  mode: "automaticExceptTables",
  calculateOnSave: false,
  fullCalculationOnLoad: true,
  forceFullCalculation: true,
  iteration: { enabled: true, maxIterations: 100, maxChange: 0.001 },
  fullPrecision: false,
});
assert.match(calculationBook.help("workbook.setCalculation").ndjson, /automaticExceptTables/);
assert.match(calculationBook.inspect({ kind: "workbook" }).ndjson, /"calculation":\{"mode":"automaticExceptTables"/);
const calculationXlsx = await SpreadsheetFile.exportXlsx(calculationBook);
const calculationZip = await JSZip.loadAsync(calculationXlsx.bytes);
assert.match(await calculationZip.file("xl/workbook.xml").async("text"), /<calcPr calcMode="autoNoTable" calcOnSave="0" fullCalcOnLoad="1" forceFullCalc="1" iterate="1" iterateCount="100" iterateDelta="0.001" fullPrecision="0"\/>/);
calculationZip.remove("customXml/open-office-artifact.json");
const calculationRoundtrip = await SpreadsheetFile.importXlsx(new FileBlob(await calculationZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: calculationXlsx.type }));
assert.deepEqual(calculationRoundtrip.calculation, calculationBook.calculation);
assert.throws(() => Workbook.create({ calculation: { mode: "automatic", iteration: { maxIterations: 0 } } }), /maximum calculation iterations/);
calculationRoundtrip.calculation.iteration.maxChange = 0;
assert.ok(calculationRoundtrip.verify().issues.some((issue) => issue.type === "invalidCalculation"));

const visibilityBook = Workbook.create();
const visibleSheet = visibilityBook.worksheets.add({ name: "Visible", visibility: "visible" });
const hiddenSheet = visibilityBook.worksheets.add("Hidden", { visibility: "hidden" });
const veryHiddenSheet = visibilityBook.worksheets.add("Internal", { visibility: "very-hidden" });
visibleSheet.getRange("A1").values = [["shown"]];
hiddenSheet.getRange("A1").values = [["hidden"]];
veryHiddenSheet.getRange("A1").values = [["internal"]];
assert.equal(visibilityBook.worksheets.getActiveWorksheet(), visibleSheet);
assert.deepEqual(visibilityBook.worksheets.items.map((item) => item.visibility), ["visible", "hidden", "veryHidden"]);
assert.match(visibilityBook.inspect({ kind: "sheet" }).ndjson, /"name":"Internal"[^\n]*"visibility":"veryHidden"/);
assert.equal(veryHiddenSheet.layoutJson().view.visibility, "veryHidden");
assert.match(visibilityBook.help("worksheet.visibility").ndjson, /at least one sheet must remain visible/i);
assert.throws(() => { hiddenSheet.visibility = "collapsed"; }, /worksheet visibility/i);
const visibilityXlsx = await SpreadsheetFile.exportXlsx(visibilityBook);
const visibilityZip = await JSZip.loadAsync(visibilityXlsx.bytes);
const visibilityWorkbookXml = await visibilityZip.file("xl/workbook.xml").async("text");
assert.match(visibilityWorkbookXml, /<workbookView activeTab="0"\s*\/>/);
assert.match(visibilityWorkbookXml, /<sheet name="Hidden" sheetId="2" state="hidden"/);
assert.match(visibilityWorkbookXml, /<sheet name="Internal" sheetId="3" state="veryHidden"/);
visibilityZip.remove("customXml/open-office-artifact.json");
const visibilityRoundtrip = await SpreadsheetFile.importXlsx(new FileBlob(await visibilityZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: visibilityXlsx.type }));
assert.deepEqual(visibilityRoundtrip.worksheets.items.map((item) => item.visibility), ["visible", "hidden", "veryHidden"]);
assert.equal(visibilityRoundtrip.worksheets.getActiveWorksheet().name, "Visible");
visibleSheet.visibility = "hidden";
assert.ok(visibilityBook.verify().issues.some((issue) => issue.type === "noVisibleSheets"));
await assert.rejects(() => SpreadsheetFile.exportXlsx(visibilityBook), /at least one visible worksheet/i);
assert.throws(() => visibilityBook.worksheets.getActiveWorksheet(), /no visible worksheets/i);

const activeWorksheetBook = Workbook.create();
const firstActiveCandidate = activeWorksheetBook.worksheets.add("First");
const selectedActiveWorksheet = activeWorksheetBook.worksheets.add("Selected");
const inactiveHiddenWorksheet = activeWorksheetBook.worksheets.add("Hidden", { visibility: "hidden" });
assert.equal(activeWorksheetBook.worksheets.setActiveWorksheet("Selected"), selectedActiveWorksheet);
assert.equal(activeWorksheetBook.worksheets.getActiveWorksheet(), selectedActiveWorksheet);
assert.deepEqual(activeWorksheetBook.worksheets.getSelectedWorksheets(), [selectedActiveWorksheet]);
assert.deepEqual(activeWorksheetBook.worksheets.setSelectedWorksheets([selectedActiveWorksheet, firstActiveCandidate]), [firstActiveCandidate, selectedActiveWorksheet]);
assert.equal(activeWorksheetBook.worksheets.getActiveWorksheet(), selectedActiveWorksheet);
assert.deepEqual(activeWorksheetBook.worksheets.getSelectedWorksheets(), [firstActiveCandidate, selectedActiveWorksheet]);
assert.match(activeWorksheetBook.inspect({ kind: "workbook" }).ndjson, /"activeSheet":"Selected","selectedSheets":\["First","Selected"\]/);
assert.equal(activeWorksheetBook.layoutJson().activeSheet, "Selected");
assert.deepEqual(activeWorksheetBook.layoutJson().selectedSheets, ["First", "Selected"]);
assert.match(activeWorksheetBook.help("workbook.worksheets.setActiveWorksheet").ndjson, /zero-based position.*activeTab/i);
assert.match(activeWorksheetBook.help("workbook.worksheets.setSelectedWorksheets").ndjson, /tabSelected/i);
assert.throws(() => activeWorksheetBook.worksheets.setActiveWorksheet(inactiveHiddenWorksheet), /must be visible/i);
assert.throws(() => activeWorksheetBook.worksheets.setSelectedWorksheets([inactiveHiddenWorksheet]), /must all be visible/i);
assert.throws(() => activeWorksheetBook.worksheets.setSelectedWorksheets([firstActiveCandidate, firstActiveCandidate]), /duplicates/i);
assert.throws(() => { selectedActiveWorksheet.visibility = "hidden"; }, /select another active worksheet first/i);
assert.throws(() => { firstActiveCandidate.visibility = "hidden"; }, /selected worksheet/i);
const activeWorksheetXlsx = await SpreadsheetFile.exportXlsx(activeWorksheetBook);
const activeWorksheetZip = await JSZip.loadAsync(activeWorksheetXlsx.bytes);
assert.match(await activeWorksheetZip.file("xl/workbook.xml").async("text"), /<workbookView activeTab="1"\s*\/>/);
assert.match(await activeWorksheetZip.file("xl/worksheets/sheet1.xml").async("text"), /<sheetView workbookViewId="0" tabSelected="1">/);
assert.match(await activeWorksheetZip.file("xl/worksheets/sheet2.xml").async("text"), /<sheetView workbookViewId="0" tabSelected="1">/);
activeWorksheetZip.remove("customXml/open-office-artifact.json");
const activeWorksheetRoundtrip = await SpreadsheetFile.importXlsx(new FileBlob(await activeWorksheetZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: activeWorksheetXlsx.type }));
assert.equal(activeWorksheetRoundtrip.worksheets.getActiveWorksheet().name, "Selected");
assert.deepEqual(activeWorksheetRoundtrip.worksheets.getSelectedWorksheets().map((sheet) => sheet.name), ["First", "Selected"]);
activeWorksheetBook.worksheets.setActiveWorksheet(firstActiveCandidate);
selectedActiveWorksheet.visibility = "hidden";
assert.equal(activeWorksheetBook.worksheets.getActiveWorksheet(), firstActiveCandidate);
assert.deepEqual(activeWorksheetBook.worksheets.getSelectedWorksheets(), [firstActiveCandidate]);

const multiWindowBook = Workbook.create();
const multiWindowSummary = multiWindowBook.worksheets.add("Summary");
const multiWindowDetail = multiWindowBook.worksheets.add("Detail");
const multiWindowReview = multiWindowBook.worksheets.add("Review");
multiWindowBook.worksheets.setActiveWorksheet(multiWindowDetail);
multiWindowBook.worksheets.setSelectedWorksheets([multiWindowSummary, multiWindowDetail]);
const reviewWindow = multiWindowBook.windows.add({
  activeWorksheet: multiWindowReview,
  selectedWorksheets: [multiWindowDetail, multiWindowReview],
});
assert.equal(multiWindowBook.windows.count, 2);
assert.equal(multiWindowBook.windows.getItemAt(0).getActiveWorksheet(), multiWindowDetail);
assert.equal(reviewWindow.getActiveWorksheet(), multiWindowReview);
assert.deepEqual(reviewWindow.getSelectedWorksheets(), [multiWindowDetail, multiWindowReview]);
assert.match(multiWindowBook.inspect({ kind: "workbookWindow" }).ndjson, /"kind":"workbookWindow","id":"workbook-window\/2","index":1,"activeWorksheet":"Review","selectedWorksheets":\["Detail","Review"\]/);
assert.match(multiWindowBook.help("workbook.windows.add").ndjson, /matching workbookView and one sheetView per worksheet/i);
assert.throws(() => { multiWindowReview.visibility = "hidden"; }, /window 1 active worksheet/i);
assert.throws(() => reviewWindow.setSelectedWorksheets([]), /at least one visible worksheet/i);
const multiWindowXlsx = await SpreadsheetFile.exportXlsx(multiWindowBook);
const multiWindowZip = await JSZip.loadAsync(multiWindowXlsx.bytes);
const multiWindowWorkbookXml = await multiWindowZip.file("xl/workbook.xml").async("text");
assert.match(multiWindowWorkbookXml, /<bookViews><workbookView activeTab="1"\s*\/><workbookView activeTab="2"\s*\/><\/bookViews>/);
assert.match(await multiWindowZip.file("xl/worksheets/sheet1.xml").async("text"), /<sheetView workbookViewId="0" tabSelected="1">[\s\S]*<sheetView workbookViewId="1">/);
assert.match(await multiWindowZip.file("xl/worksheets/sheet2.xml").async("text"), /<sheetView workbookViewId="0" tabSelected="1">[\s\S]*<sheetView workbookViewId="1" tabSelected="1">/);
assert.match(await multiWindowZip.file("xl/worksheets/sheet3.xml").async("text"), /<sheetView workbookViewId="0">[\s\S]*<sheetView workbookViewId="1" tabSelected="1">/);
multiWindowZip.remove("customXml/open-office-artifact.json");
const multiWindowRoundtrip = await SpreadsheetFile.importXlsx(new FileBlob(await multiWindowZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: multiWindowXlsx.type }));
assert.equal(multiWindowRoundtrip.windows.count, 2);
assert.deepEqual(multiWindowRoundtrip.windows.toJSON().map((window) => ({ activeWorksheet: window.activeWorksheet, selectedWorksheets: window.selectedWorksheets })), [
  { activeWorksheet: "Detail", selectedWorksheets: ["Summary", "Detail"] },
  { activeWorksheet: "Review", selectedWorksheets: ["Detail", "Review"] },
]);

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet1");
sheet.getRange("A1:C3").values = [["A", "B", "Sum"], [2, 3, null], [5, 7, null]];
sheet.getRange("C2:C3").formulas = [["=A2+B2"], ["=A3+B3"]];
sheet.getRange("A1:D1").format = { fill: "#0f172a", font: { bold: true, color: "#ffffff", name: "Aptos", size: 12 }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: { style: "thin", color: "#334155" } };
sheet.getRange("B2:C3").setFormat({ numberFormat: "#,##0", fill: "sky-100" });
workbook.recalculate();
assert.equal(sheet.freezePanes.freezeRows(1), sheet.freezePanes);
assert.equal(sheet.freezePanes.freezeColumns(2), sheet.freezePanes);
assert.deepEqual(sheet.freezePanes.toJSON(), { rows: 1, columns: 2, frozen: true, topLeftCell: "C2", activePane: "bottomRight" });
sheet.freezePanes.freezeRows(0);
assert.deepEqual(sheet.freezePanes.toJSON(), { rows: 0, columns: 2, frozen: true, topLeftCell: "C1", activePane: "topRight" });
sheet.freezePanes.unfreeze().freezeRows(1).freezeColumns(2);
sheet.showGridLines = false;
assert.throws(() => sheet.freezePanes.freezeRows(-1), /integer from 0/);
assert.throws(() => sheet.freezePanes.freezeColumns(1.5), /integer from 0/);
assert.match(workbook.help("worksheet.freezePanes").ndjson, /Freeze a leading row count/);
const sheetViewInspect = workbook.inspect({ kind: "sheet" }).ndjson;
assert.match(sheetViewInspect, /"showGridLines":false/);
assert.match(sheetViewInspect, /"freezePanes":\{"rows":1,"columns":2,"frozen":true,"topLeftCell":"C2","activePane":"bottomRight"\}/);
assert.deepEqual(sheet.layoutJson().view.freezePanes, { rows: 1, columns: 2, frozen: true, topLeftCell: "C2", activePane: "bottomRight" });
const invalidFreezeBook = Workbook.create();
const invalidFreezeSheet = invalidFreezeBook.worksheets.add("InvalidFreeze");
invalidFreezeSheet.getRange("A1").values = [[1]];
invalidFreezeSheet.freezePanes._rows = -1;
assert.ok(invalidFreezeBook.verify().issues.some((issue) => issue.type === "invalidFrozenRows"));
await assert.rejects(() => SpreadsheetFile.exportXlsx(invalidFreezeBook), /frozen row count must be an integer/);
const unfrozenBook = Workbook.create();
const unfrozenSheet = unfrozenBook.worksheets.add("Unfrozen");
unfrozenSheet.getRange("A1").values = [[1]];
unfrozenSheet.freezePanes.freezeRows(1).freezeColumns(1).unfreeze();
const unfrozenZip = await JSZip.loadAsync(new Uint8Array(await (await SpreadsheetFile.exportXlsx(unfrozenBook)).arrayBuffer()));
assert.doesNotMatch(await unfrozenZip.file("xl/worksheets/sheet1.xml").async("text"), /<sheetViews>/);

const dimensionBook = Workbook.create();
const dimensionSheet = dimensionBook.worksheets.add("Dimensions");
dimensionSheet.getRange("A1:C3").values = [["ID", "Long descriptive heading", "Hidden"], [1, "first\nsecond\nthird", "x"], [2, "short", "y"]];
dimensionSheet.getRange("A1:A3").format.columnWidth = 18;
dimensionSheet.getRange("B1:B3").format.columnWidthPx = 120;
dimensionSheet.getRange("A1:C1").format.rowHeight = 24;
dimensionSheet.getRange("A2:C2").format.rowHeightPx = 30;
dimensionSheet.getRange("C1:C3").format.columnHidden = true;
dimensionSheet.getRange("A3:C3").format.rowHidden = true;
dimensionSheet.getRange("A1:C1").format.fill = "#123456";
assert.equal(dimensionSheet.getRange("A1").format.fill, "#123456");
assert.equal(dimensionSheet.getRange("A1:A3").format.columnWidth, 18);
assert.ok(Math.abs(dimensionSheet.getRange("B1:B3").format.columnWidthPx - 120) <= 1);
assert.equal(dimensionSheet.getRange("A1:C1").format.rowHeight, 24);
assert.ok(Math.abs(dimensionSheet.getRange("A2:C2").format.rowHeightPx - 30) < 0.01);
assert.equal(dimensionSheet.getRange("C1:C3").format.columnHidden, true);
assert.equal(dimensionSheet.getRange("A3:C3").format.rowHidden, true);
assert.throws(() => { dimensionSheet.getRange("A1").format.columnWidth = 0; }, /column width must be greater than 0/);
assert.throws(() => { dimensionSheet.getRange("A1").format.rowHeightPx = -1; }, /row height must be greater than 0/);
dimensionSheet.getRange("A1:B3").format.autofitColumns();
dimensionSheet.getRange("A1:B3").format.autofitRows();
assert.ok(dimensionSheet.getRange("B1:B3").format.columnWidthPx > dimensionSheet.getRange("A1:A3").format.columnWidthPx);
assert.ok(dimensionSheet.getRange("A2:C2").format.rowHeight > 15);
const dimensionInspect = dimensionBook.inspect({ kind: "dimension" }).ndjson;
assert.match(dimensionInspect, /"kind":"column"/);
assert.match(dimensionInspect, /"bestFit":true/);
assert.match(dimensionInspect, /"column":"C"[\s\S]*"hidden":true/);
assert.match(dimensionInspect, /"kind":"row"[\s\S]*"row":3[\s\S]*"hidden":true/);
const dimensionLayout = dimensionSheet.layoutJson();
assert.equal(dimensionLayout.cells.find((cell) => cell.address === "C1").bbox[2], 0);
assert.equal(dimensionLayout.cells.find((cell) => cell.address === "A3").bbox[3], 0);
assert.equal(dimensionLayout.cells.find((cell) => cell.address === "C1").hidden, true);
const dimensionXlsx = await SpreadsheetFile.exportXlsx(dimensionBook);
const dimensionZip = await JSZip.loadAsync(new Uint8Array(await dimensionXlsx.arrayBuffer()));
const dimensionXml = await dimensionZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(dimensionXml, /<cols>[\s\S]*<col min="1" max="1" width="[^"]+" customWidth="1" bestFit="1"\/>/);
assert.match(dimensionXml, /<col min="3" max="3" hidden="1"\/>/);
assert.match(dimensionXml, /<row r="1" ht="[^"]+" customHeight="1">/);
assert.match(dimensionXml, /<row r="3"[^>]*hidden="1"[^>]*>/);
const dimensionRoundtripSheet = (await SpreadsheetFile.importXlsx(dimensionXlsx)).worksheets.getItem("Dimensions");
assert.equal(dimensionRoundtripSheet.getRange("C1:C3").format.columnHidden, true);
assert.equal(dimensionRoundtripSheet.getRange("A3:C3").format.rowHidden, true);
assert.ok(dimensionRoundtripSheet.getRange("B1:B3").format.columnWidthPx > dimensionRoundtripSheet.getRange("A1:A3").format.columnWidthPx);
const thirdPartyDimensionXml = dimensionXml
  .replace(/<cols>[\s\S]*?<\/cols>/, '<cols><col min="2" max="3" width="20" customWidth="1" hidden="1" bestFit="1"/></cols>')
  .replace('</sheetData>', '<row r="4" ht="33" customHeight="1" hidden="1"/></sheetData>');
dimensionZip.file("xl/worksheets/sheet1.xml", thirdPartyDimensionXml);
const thirdPartyDimensionBook = await SpreadsheetFile.importXlsx(new FileBlob(await dimensionZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: dimensionXlsx.type }));
const thirdPartyDimensionSheet = thirdPartyDimensionBook.worksheets.getItem("Dimensions");
assert.equal(thirdPartyDimensionSheet.getRange("B1:C3").format.columnHidden, true);
assert.equal(thirdPartyDimensionSheet.getRange("B1:C3").format.columnWidthPx, 140);
assert.equal(thirdPartyDimensionSheet.getRange("A4:C4").format.rowHeight, 33);
assert.equal(thirdPartyDimensionSheet.getRange("A4:C4").format.rowHidden, true);
const thirdPartyDimensionRoundtripZip = await JSZip.loadAsync(new Uint8Array(await (await SpreadsheetFile.exportXlsx(thirdPartyDimensionBook)).arrayBuffer()));
assert.match(await thirdPartyDimensionRoundtripZip.file("xl/worksheets/sheet1.xml").async("text"), /<col min="2" max="3" width="20" customWidth="1" hidden="1" bestFit="1"\/>/);
const invalidDimensionBook = Workbook.create();
const invalidDimensionSheet = invalidDimensionBook.worksheets.add("InvalidDimensions");
invalidDimensionSheet.getRange("A1").values = [[1]];
invalidDimensionSheet.columnDimensions.set(0, { width: 300 });
assert.ok(invalidDimensionBook.verify().issues.some((issue) => issue.type === "invalidColumnWidth"));
await assert.rejects(() => SpreadsheetFile.exportXlsx(invalidDimensionBook), /column 1 width must be greater than 0/);

const mergeFillBook = Workbook.create();
const mergeSheet = mergeFillBook.worksheets.add("Merged");
mergeSheet.getRange("A1:F2").values = [["Header", "discard", "discard", "North", "discard", "discard"], ["Left", "discard", "discard", "South", "discard", "discard"]];
const mergedHeaderRange = mergeSheet.getRange("A1:C1");
assert.equal(mergedHeaderRange.merge(), mergedHeaderRange);
assert.deepEqual(mergeSheet.getRange("A1:C1").values, [["Header", null, null]]);
mergeSheet.getRange("D1:F2").merge(true);
assert.deepEqual(mergeSheet.mergedRanges, ["A1:C1", "D1:F1", "D2:F2"]);
assert.deepEqual(mergeSheet.getRange("D1:F2").values, [["North", null, null], ["South", null, null]]);
assert.throws(() => mergeSheet.getRange("B1:D1").merge(), /overlaps existing merged range/);
mergeSheet.getRange("A1:B1").unmerge();
assert.deepEqual(mergeSheet.mergedRanges, ["D1:F1", "D2:F2"]);
mergeSheet.mergeCells("A1:C1");
mergeSheet.mergeCells("J10:K11");
mergeSheet.getRange("M1").merge();
assert.equal(mergeSheet.mergedRanges.includes("M1:M1"), false);
assert.equal(mergeSheet.usedBounds().right, 10);
assert.equal(mergeSheet.usedBounds().bottom, 10);
const mergeInspect = mergeFillBook.inspect({ kind: "sheet,merge" }).ndjson;
assert.match(mergeInspect, /"mergedRanges":4/);
assert.match(mergeInspect, /"kind":"mergedCell"[\s\S]*"range":"J10:K11"/);
const mergeLayout = mergeSheet.layoutJson();
assert.equal(mergeLayout.merges.length, 4);
assert.equal(mergeLayout.cells.find((cell) => cell.address === "A1").colSpan, 3);
assert.equal(mergeLayout.cells.find((cell) => cell.address === "B1").mergedParent, "A1");
assert.equal(mergeLayout.cells.find((cell) => cell.address === "B1").bbox[2], 0);
assert.equal((mergeSheet.toSvg().match(/>Header<\/text>/g) || []).length, 1);

const dataSheet = mergeFillBook.worksheets.add("Data Sheet");
dataSheet.getRange("A1:A3").values = [[1], [2], [3]];
const fillSheet = mergeFillBook.worksheets.add("Fill");
fillSheet.getRange("A1:B3").values = [[10, 2], [20, 3], [30, 4]];
fillSheet.getRange("C1").formulas = [["=A1+$B$1+B$1+$A1+'Data Sheet'!A1"]];
fillSheet.getRange("C1").format = { fill: "#abcdef", font: { bold: true } };
const filledFormulaRange = fillSheet.getRange("C1:C3");
assert.equal(filledFormulaRange.fillDown(), filledFormulaRange);
assert.deepEqual(fillSheet.getRange("C1:C3").formulas.flat(), ["=A1+$B$1+B$1+$A1+'Data Sheet'!A1", "=A2+$B$1+B$1+$A2+'Data Sheet'!A2", "=A3+$B$1+B$1+$A3+'Data Sheet'!A3"]);
assert.deepEqual(fillSheet.getRange("C1:C3").values.flat(), [25, 46, 67]);
assert.equal(fillSheet.getRange("C3").format.fill, "#abcdef");
assert.equal(fillSheet.getRange("C3").format.font.bold, true);
fillSheet.getRange("D1").formulas = [["=IF(A1=\"A1\",A1,0)"]];
fillSheet.getRange("D1:D3").fillDown();
assert.equal(fillSheet.getRange("D2").formulas[0][0], '=IF(A2="A1",A2,0)');
fillSheet.getRange("A5").formulas = [["=A1+$A1+A$1+$A$1"]];
fillSheet.getRange("A5").format = { fill: "#fedcba" };
fillSheet.getRange("A5:C5").fillRight();
assert.deepEqual(fillSheet.getRange("A5:C5").formulas.flat(), ["=A1+$A1+A$1+$A$1", "=B1+$A1+B$1+$A$1", "=C1+$A1+C$1+$A$1"]);
assert.equal(fillSheet.getRange("C5").format.fill, "#fedcba");
fillSheet.getRange("A7").values = [["repeat"]];
fillSheet.getRange("A7:C7").fillRight();
assert.deepEqual(fillSheet.getRange("A7:C7").values, [["repeat", "repeat", "repeat"]]);
fillSheet.getRange("E1:F1").merge();
assert.throws(() => fillSheet.getRange("E1:E3").fillDown(), /intersects merged cells/);

const mergeFillXlsx = await SpreadsheetFile.exportXlsx(mergeFillBook);
const mergeFillZip = await JSZip.loadAsync(new Uint8Array(await mergeFillXlsx.arrayBuffer()));
const mergeSheetXml = await mergeFillZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(mergeSheetXml, /<mergeCells count="4"><mergeCell ref="A1:C1"\/><mergeCell ref="D1:F1"\/><mergeCell ref="D2:F2"\/><mergeCell ref="J10:K11"\/><\/mergeCells>/);
const mergeFillRoundtrip = await SpreadsheetFile.importXlsx(mergeFillXlsx);
assert.deepEqual(mergeFillRoundtrip.worksheets.getItem("Merged").mergedRanges, ["A1:C1", "D1:F1", "D2:F2", "J10:K11"]);
assert.deepEqual(mergeFillRoundtrip.worksheets.getItem("Fill").getRange("C1:C3").formulas, fillSheet.getRange("C1:C3").formulas);
assert.equal(mergeFillRoundtrip.worksheets.getItem("Fill").getRange("C3").format.fill, "#ABCDEF");
const overlappingMergeXml = mergeSheetXml.replace(/<mergeCells[\s\S]*?<\/mergeCells>/, '<mergeCells count="2"><mergeCell ref="A1:C1"/><mergeCell ref="B1:D1"/></mergeCells>');
mergeFillZip.file("xl/worksheets/sheet1.xml", overlappingMergeXml);
const overlappingMergeBook = await SpreadsheetFile.importXlsx(new FileBlob(await mergeFillZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: mergeFillXlsx.type }));
assert.ok(overlappingMergeBook.verify().issues.some((issue) => issue.type === "overlappingMergedRanges"));
await assert.rejects(() => SpreadsheetFile.exportXlsx(overlappingMergeBook), /merged range B1:D1 overlaps A1:C1/);

const inspect = workbook.inspect({ kind: "workbook,sheet,table,formula", range: "A1:C3", include: "values,formulas" });
assert.match(inspect.ndjson, /"value":5|"values":/);
assert.match(inspect.ndjson, /"formula":"=A2\+B2"/);
const styleInspect = workbook.inspect({ kind: "style", range: "A1:D3", maxChars: 8000 }).ndjson;
assert.match(styleInspect, /"kind":"style"/);
assert.match(styleInspect, /"numberFormat":"#,##0"/);
assert.match(styleInspect, /"fill":"sky-100"/);
assert.match(styleInspect, /"alignment":\{"horizontal":"center"/);
assert.match(styleInspect, /"border":\{"style":"thin"/);
assert.match(workbook.help("range.format").ndjson, /cell style/);
assert.match(workbook.help("fx.PMT").ndjson, /fx.PMT/);
assert.match(workbook.help("range.dataValidation").ndjson, /validation rule/);
assert.match(workbook.help("range.conditionalFormats.add").ndjson, /conditional formatting/);
assert.match(workbook.help("workbook.comments.addThread").ndjson, /threaded comments/);
const statusRange = sheet.getRange("D2:D3");
statusRange.values = [["Not Started"], ["In Progress"]];
statusRange.dataValidation = { rule: { type: "list", values: ["Not Started", "In Progress", "Done"] } };
const cf = sheet.getRange("C2:C3").conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 10, format: { fill: "green" } });
const customCf = sheet.getRange("A2:B3").conditionalFormats.addCustom("=A2<B2", { fill: "sky-100" });
workbook.comments.setSelf({ displayName: "Analyst" });
const rootCommentId = "{11111111-1111-4111-8111-111111111111}";
const replyCommentId = "{22222222-2222-4222-8222-222222222222}";
const analystPersonId = "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}";
const reviewerPersonId = "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}";
const thread = workbook.comments.addThread({ cell: sheet.getRange("C2") }, "Formula checks revenue sum.", {
  comment: { id: rootCommentId, personId: analystPersonId, person: { id: analystPersonId, displayName: "Analyst", userId: "analyst@example.com", providerId: "None" }, date: "2026-07-12T09:00:00.000Z" },
});
thread.addReply("Reviewed by model.", { id: replyCommentId, parentId: rootCommentId, personId: reviewerPersonId, person: { id: reviewerPersonId, displayName: "Reviewer", userId: "reviewer@example.com", providerId: "None" }, author: "Reviewer", date: "2026-07-12T09:05:00.000Z" }).resolve();
const invalidThreadWorkbook = Workbook.create();
const invalidThreadSheet = invalidThreadWorkbook.worksheets.add("Invalid");
invalidThreadWorkbook.comments.addThread({ cell: invalidThreadSheet.getRange("A1") }, "Invalid native identity", { comment: { id: "not-a-guid" } });
assert.ok(invalidThreadWorkbook.verify().issues.some((issue) => issue.type === "threadedCommentMetadataInvalid"));
await assert.rejects(() => SpreadsheetFile.exportXlsx(invalidThreadWorkbook), /brace-delimited GUID/);

const tasksTable = sheet.tables.add("A1:D3", true, "TasksTable");
tasksTable.style = "TableStyleMedium4";
tasksTable.showBandedColumns = true;
tasksTable.showFirstColumn = true;
tasksTable.showLastColumn = true;
tasksTable.showRowStripes = false;
tasksTable.rows.add(null, [["8", "13", "21", "Done"]]);
sheet.getRange("E2").formulas = [["=SUM(TasksTable[Sum])"]];
workbook.recalculate();
assert.equal(sheet.getRange("E2").values[0][0], 38);
assert.deepEqual(tasksTable.getDataRows()[0].slice(0, 3), [2, 3, 5]);
assert.equal(tasksTable.getHeaderRowRange().values[0][0], "A");

const chartSource = sheet.getRange("F1:G4");
chartSource.values = [["Month", "Revenue"], ["Jan", 100], ["Feb", 120], ["Mar", 130]];
const colorScaleCf = chartSource.getRange("G2:G4").conditionalFormats.addColorScale({ colors: ["#fee2e2", "#fef3c7", "#22c55e"] });
const revenueName = workbook.definedNames.add("RevenueData", "Sheet1!G2:G4", { comment: "Revenue data body", hidden: false });
const localRevenueName = workbook.definedNames.add("LocalRevenueData", "Sheet1!G2:G4", { scope: "Sheet1", hidden: true });
sheet.getRange("E3").formulas = [["=SUM(RevenueData)"]];
workbook.recalculate();
assert.equal(sheet.getRange("E3").values[0][0], 350);
const chartFromRange = sheet.charts.add("line", chartSource);
chartFromRange.title = "Revenue Trend";
chartFromRange.hasLegend = false;
chartFromRange.setPosition("I1", "M10");
const chartFromConfig = sheet.charts.add("bar", { name: "ScoresChart", title: "Scores", categories: ["A", "B"], series: [{ name: "Score", values: [9, 7] }], position: { left: 40, top: 220, width: 240, height: 160 } });
const revenuePivot = sheet.pivotTables.add({
  name: "RevenuePivot",
  sourceRange: "F1:G4",
  targetRange: "N1:O4",
  rows: ["Month"],
  values: [{ field: "Revenue", summarizeBy: "sum", name: "Revenue sum" }],
});
assert.deepEqual(revenuePivot.computedValues(), [["Month", "Revenue sum"], ["Jan", 100], ["Feb", 120], ["Mar", 130]]);
sheet.getRange("P1:U7").values = [
  ["Region", "Quarter", "Product", "Revenue", "Cost", "OrderDate"],
  ["East", "Q1", "Core", 10, 6, "2026-01-15"],
  ["East", "Q2", "Core", 20, 12, "2026-02-01"],
  ["West", "Q1", "Core", 30, 20, "2026-02-20"],
  ["West", "Q2", "Legacy", 40, 25, "2026-03-15"],
  ["East", "Q1", "Legacy", 5, 3, "2026-03-05"],
  ["West", "Q1", "Legacy", 7, 4, "2026-04-01"],
];
const regionalPivot = sheet.pivotTables.add({
  name: "RegionalPivot",
  sourceRange: "P1:U7",
  targetRange: "W1:Z4",
  rowFields: ["Region", "OrderDate"],
  columnFields: ["Quarter"],
  valueFields: [{ field: "Revenue", summarizeBy: "sum", name: "Revenue total" }, { field: "Profit", name: "Profit total" }],
  calculatedFields: [{ name: "Profit", formula: "=([Revenue] - [Cost]) * 100%" }],
  filters: [
    { field: "Quarter", include: ["Q1"] },
    { field: "OrderDate", type: "dateBetween", value1: "2026-02-01", value2: "2026-03-31" },
  ],
  refreshPolicy: { refreshOnLoad: false, saveData: true, enableRefresh: false, invalid: true, missingItemsLimit: 3, refreshedBy: "QA Agent", refreshedDateIso: "2026-07-12T00:00:00Z" },
});
assert.deepEqual(regionalPivot.computedValues(), [["Region", "OrderDate", "Q1 — Revenue total", "Q1 — Profit total"], ["West", "2026-02-20", 30, 10], ["East", "2026-03-05", 5, 2]]);
assert.deepEqual(regionalPivot.calculatedFields, [{ name: "Profit", formula: "=('Revenue'-'Cost')*100%", numFmtId: 0, references: ["Revenue", "Cost"] }]);
assert.deepEqual(regionalPivot.filters, [{ field: "Quarter", include: ["Q1"] }, { field: "OrderDate", type: "dateBetween", value1: "2026-02-01", value2: "2026-03-31", useWholeDay: true }]);
assert.equal(regionalPivot.refreshPolicy.saveData, true);
assert.match(regionalPivot.inspectRecord().columnFields.join(","), /Quarter/);
assert.match(JSON.stringify(regionalPivot.layoutJson()), /refreshPolicy/);
const calendarPivot = sheet.pivotTables.add({
  name: "CalendarPivot",
  sourceRange: "P1:U7",
  targetRange: "W8:Z12",
  groupFields: [
    { name: "Order Year", sourceField: "OrderDate", groupBy: "years" },
    { name: "Order Quarter", sourceField: "OrderDate", groupBy: "quarters" },
    { name: "Order Month", sourceField: "OrderDate", groupBy: "months" },
  ],
  rowFields: ["Order Year", "Order Quarter"],
  columnFields: ["Order Month"],
  valueFields: [{ field: "Revenue", summarizeBy: "sum", name: "Revenue by month" }],
  filters: [{ field: "Order Month", exclude: ["Jan"] }],
});
assert.deepEqual(calendarPivot.groupFields, [
  { name: "Order Year", sourceField: "OrderDate", groupBy: "years" },
  { name: "Order Quarter", sourceField: "OrderDate", groupBy: "quarters", parent: "Order Year" },
  { name: "Order Month", sourceField: "OrderDate", groupBy: "months", parent: "Order Quarter" },
]);
assert.deepEqual(calendarPivot.computedValues(), [["Order Year", "Order Quarter", "Feb", "Mar", "Apr"], ["2026", "Q1", 50, 45, 0], ["2026", "Q2", 0, 0, 7]]);
assert.match(calendarPivot.inspectRecord().groupFields[2].parent, /Order Quarter/);
assert.equal(workbook.verify().issues.some((issue) => issue.code === "pivotFieldMissing" && issue.id === calendarPivot.id), false);
const groupingBook = Workbook.create();
const groupingSheet = groupingBook.worksheets.add("Grouping");
groupingSheet.getRange("A1:C7").values = [
  ["Region", "Score", "Amount"],
  ["East", 5, 10],
  ["West", 12, 20],
  ["Central", 18, 30],
  ["East", 25, 40],
  ["West", 35, 50],
  ["Other", 45, 60],
];
const numericPivot = groupingSheet.pivotTables.add({
  name: "NumericGrouping",
  sourceRange: "A1:C7",
  targetRange: "E1:F8",
  groupFields: [{ name: "Score Band", sourceField: "Score", groupBy: "range", range: { autoStart: false, autoEnd: false, startNum: 0, endNum: 39, groupInterval: 10 } }],
  rowFields: ["Score Band"],
  valueFields: [{ field: "Amount", summarizeBy: "sum" }],
});
const discretePivot = groupingSheet.pivotTables.add({
  name: "DiscreteGrouping",
  sourceRange: "A1:C7",
  targetRange: "H1:I6",
  groupFields: [{ name: "Region Cluster", sourceField: "Region", groupBy: "discrete", groups: [{ name: "Coasts", items: ["East", "West"] }, { name: "Core", items: ["Central"] }] }],
  rowFields: ["Region Cluster"],
  valueFields: [{ field: "Amount", summarizeBy: "sum" }],
});
assert.deepEqual(numericPivot.computedValues(), [["Score Band", "sum of Amount"], ["0-9", 10], ["10-19", 50], ["20-29", 40], ["30-39", 50], [">39", 60]]);
assert.deepEqual(discretePivot.computedValues(), [["Region Cluster", "sum of Amount"], ["Coasts", 120], ["Core", 30], ["Other", 60]]);
const groupingXlsx = await SpreadsheetFile.exportXlsx(groupingBook);
const groupingZip = await JSZip.loadAsync(new Uint8Array(await groupingXlsx.arrayBuffer()));
const numericCacheXml = await groupingZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text");
const discreteCacheXml = await groupingZip.file("xl/pivotCache/pivotCacheDefinition2.xml").async("text");
assert.match(numericCacheXml, /<fieldGroup base="1"><rangePr groupBy="range" autoStart="0" autoEnd="0" startNum="0" endNum="39" groupInterval="10"\/><groupItems count="5"><s v="0-9"\/><s v="10-19"\/><s v="20-29"\/><s v="30-39"\/><s v="&gt;39"\/><\/groupItems><\/fieldGroup>/);
assert.match(discreteCacheXml, /<fieldGroup base="0"><discretePr count="4"><x v="0"\/><x v="0"\/><x v="1"\/><x v="2"\/><\/discretePr><groupItems count="3"><s v="Coasts"\/><s v="Core"\/><s v="Other"\/><\/groupItems><\/fieldGroup>/);
groupingZip.remove("customXml/open-office-artifact.json");
const nativeGroupingBook = await SpreadsheetFile.importXlsx(new FileBlob(await groupingZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: groupingXlsx.type }));
const nativeNumericPivot = nativeGroupingBook.resolve("NumericGrouping");
const nativeDiscretePivot = nativeGroupingBook.resolve("DiscreteGrouping");
assert.deepEqual(nativeNumericPivot.computedValues(), numericPivot.computedValues());
assert.deepEqual(nativeNumericPivot.groupFields[0].range, { autoStart: false, autoEnd: false, startNum: 0, endNum: 39, groupInterval: 10 });
assert.deepEqual(nativeDiscretePivot.computedValues(), discretePivot.computedValues());
assert.deepEqual(nativeDiscretePivot.groupFields[0].groups, [{ name: "Coasts", items: ["East", "West"] }, { name: "Core", items: ["Central"] }]);
const nativeGroupingRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(nativeGroupingBook));
assert.deepEqual(nativeGroupingRoundtrip.resolve("NumericGrouping").computedValues(), numericPivot.computedValues());
assert.deepEqual(nativeGroupingRoundtrip.resolve("DiscreteGrouping").computedValues(), discretePivot.computedValues());
const timeBook = Workbook.create({ dateSystem: "1904" });
const timeSheet = timeBook.worksheets.add("Time Groups");
timeSheet.getRange("A1:B6").values = [
  ["Timestamp", "Amount"],
  ["2026-07-01T09:07:05Z", 10],
  ["2026-07-08T09:22:15Z", 20],
  [new Date("2026-07-15T10:37:25Z"), 30],
  ["2026-07-22T10:52:35Z", 40],
  [0.5, 50],
];
const timePivot = timeSheet.pivotTables.add({
  name: "TimeGrouping",
  sourceRange: "A1:B6",
  targetRange: "D1:I8",
  groupFields: [
    { name: "Day Band", sourceField: "Timestamp", groupBy: "days", range: { groupInterval: 7 } },
    { name: "Hour", sourceField: "Timestamp", groupBy: "hours" },
    { name: "Minute Band", sourceField: "Timestamp", groupBy: "minutes", range: { groupInterval: 15 } },
    { name: "Second Band", sourceField: "Timestamp", groupBy: "seconds", range: { groupInterval: 10 } },
  ],
  rowFields: ["Day Band", "Hour", "Minute Band", "Second Band"],
  valueFields: [{ field: "Amount", summarizeBy: "sum" }],
});
assert.deepEqual(timePivot.groupFields.map(({ name, parent }) => [name, parent]), [["Day Band", undefined], ["Hour", "Day Band"], ["Minute Band", "Hour"], ["Second Band", "Minute Band"]]);
assert.deepEqual(timePivot.computedValues()[1], ["1-7", "09", "00-14", "00-09", 10]);
assert.deepEqual(timePivot.computedValues().at(-1), ["1-7", "12", "00-14", "00-09", 50]);
const timeXlsx = await SpreadsheetFile.exportXlsx(timeBook);
const timeZip = await JSZip.loadAsync(new Uint8Array(await timeXlsx.arrayBuffer()));
const timeCacheXml = await timeZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text");
assert.match(timeCacheXml, /<d v="2026-07-01T09:07:05"\/>/);
assert.match(timeCacheXml, /<rangePr groupBy="days" autoStart="1" autoEnd="1" groupInterval="7"\/>/);
assert.match(timeCacheXml, /<rangePr groupBy="hours" autoStart="1" autoEnd="1" groupInterval="1"\/>/);
assert.match(timeCacheXml, /<rangePr groupBy="minutes" autoStart="1" autoEnd="1" groupInterval="15"\/>/);
assert.match(timeCacheXml, /<rangePr groupBy="seconds" autoStart="1" autoEnd="1" groupInterval="10"\/>/);
timeZip.remove("customXml/open-office-artifact.json");
const nativeTimeBook = await SpreadsheetFile.importXlsx(new FileBlob(await timeZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: timeXlsx.type }));
assert.deepEqual(nativeTimeBook.resolve("TimeGrouping").computedValues(), timePivot.computedValues());
const timeRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(nativeTimeBook));
assert.deepEqual(timeRoundtrip.resolve("TimeGrouping").computedValues(), timePivot.computedValues());
const relativeBook = Workbook.create();
const relativeSheet = relativeBook.worksheets.add("Relative Filters");
relativeSheet.getRange("A1:B6").values = [["Date", "Amount"], ["2026-06-30", 5], ["2026-07-01", 10], ["2026-07-12", 20], ["2026-07-31", 30], ["2026-08-01", 40]];
const relativePivot = relativeSheet.pivotTables.add({
  name: "RelativeDatePivot",
  sourceRange: "A1:B6",
  targetRange: "D1:E6",
  rowFields: ["Date"],
  valueFields: [{ field: "Amount", summarizeBy: "sum" }],
  filters: [{ field: "Date", type: "thisMonth", asOf: "2026-07-12" }],
});
assert.deepEqual(relativePivot.filters, [{ field: "Date", type: "thisMonth", asOf: "2026-07-12", useWholeDay: true }]);
assert.deepEqual(relativePivot.computedValues(), [["Date", "sum of Amount"], ["2026-07-01", 10], ["2026-07-12", 20], ["2026-07-31", 30]]);
const relativeXlsx = await SpreadsheetFile.exportXlsx(relativeBook);
const relativeZip = await JSZip.loadAsync(new Uint8Array(await relativeXlsx.arrayBuffer()));
assert.match(await relativeZip.file("xl/pivotTables/pivotTable1.xml").async("text"), /<filters count="1"><filter fld="0" type="thisMonth" id="1"\/><\/filters>/);
relativeZip.remove("customXml/open-office-artifact.json");
const nativeRelativeBook = await SpreadsheetFile.importXlsx(new FileBlob(await relativeZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: relativeXlsx.type }), { relativeDateAsOf: "2026-07-12" });
assert.deepEqual(nativeRelativeBook.resolve("RelativeDatePivot").filters, relativePivot.filters);
assert.deepEqual(nativeRelativeBook.resolve("RelativeDatePivot").computedValues(), relativePivot.computedValues());
const relativeRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(nativeRelativeBook), { relativeDateAsOf: "2026-07-12" });
assert.deepEqual(relativeRoundtrip.resolve("RelativeDatePivot").computedValues(), relativePivot.computedValues());
const preciseDateBook = Workbook.create();
const preciseDateSheet = preciseDateBook.worksheets.add("Precise Dates");
preciseDateSheet.getRange("A1:B5").values = [
  ["Timestamp", "Amount"],
  ["2026-07-12T08:30:00Z", 10],
  ["2026-07-12T12:00:00Z", 20],
  ["2026-07-12T17:45:00Z", 30],
  ["2026-07-13T00:00:00Z", 40],
];
const preciseDatePivot = preciseDateSheet.pivotTables.add({
  name: "PreciseDatePivot",
  sourceRange: "A1:B5",
  targetRange: "D1:E4",
  rowFields: ["Timestamp"],
  filters: [{ field: "Timestamp", type: "dateBetween", value1: "2026-07-12T09:00:00Z", value2: "2026-07-12T17:00:00Z", useWholeDay: false }],
  valueFields: [{ field: "Amount", summarizeBy: "sum" }],
});
const wholeDayPivot = preciseDateSheet.pivotTables.add({
  name: "WholeDayPivot",
  sourceRange: "A1:B5",
  targetRange: "G1:H6",
  rowFields: ["Timestamp"],
  filters: [{ field: "Timestamp", type: "dateEqual", value1: "2026-07-12", useWholeDay: true }],
  valueFields: [{ field: "Amount", summarizeBy: "sum" }],
});
assert.deepEqual(preciseDatePivot.filters, [{ field: "Timestamp", type: "dateBetween", value1: "2026-07-12T09:00:00", value2: "2026-07-12T17:00:00", useWholeDay: false }]);
assert.deepEqual(preciseDatePivot.computedValues(), [["Timestamp", "sum of Amount"], ["2026-07-12T12:00:00Z", 20]]);
assert.deepEqual(wholeDayPivot.computedValues(), [["Timestamp", "sum of Amount"], ["2026-07-12T08:30:00Z", 10], ["2026-07-12T12:00:00Z", 20], ["2026-07-12T17:45:00Z", 30]]);
const offsetPrecisePivot = preciseDateSheet.pivotTables.add({
  name: "OffsetPrecisePivot",
  sourceRange: "A1:B5",
  targetRange: "J1:K3",
  rowFields: ["Timestamp"],
  filters: [{ field: "Timestamp", type: "dateEqual", value1: "2026-07-12T10:30:00+02:00", useWholeDay: false }],
  valueFields: [{ field: "Amount", summarizeBy: "sum" }],
});
assert.equal(offsetPrecisePivot.filters[0].value1, "2026-07-12T08:30:00");
assert.deepEqual(offsetPrecisePivot.computedValues(), [["Timestamp", "sum of Amount"], ["2026-07-12T08:30:00Z", 10]]);
const preciseDateXlsx = await SpreadsheetFile.exportXlsx(preciseDateBook);
const preciseDateZip = await JSZip.loadAsync(new Uint8Array(await preciseDateXlsx.arrayBuffer()));
const precisePivotXml = await preciseDateZip.file("xl/pivotTables/pivotTable1.xml").async("text");
assert.match(precisePivotXml, /xmlns:x14="http:\/\/schemas\.microsoft\.com\/office\/spreadsheetml\/2010\/11\/main"/);
assert.match(precisePivotXml, /mc:Ignorable="x14"/);
assert.match(precisePivotXml, /stringValue1="2026-07-12T09:00:00" stringValue2="2026-07-12T17:00:00"/);
assert.match(precisePivotXml, /<ext uri="\{0605FD5F-26C8-4aeb-8148-2DB25E43C511\}"><x14:pivotFilter useWholeDay="0"\/><\/ext>/);
preciseDateZip.remove("customXml/open-office-artifact.json");
const nativePreciseDateBook = await SpreadsheetFile.importXlsx(new FileBlob(await preciseDateZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: preciseDateXlsx.type }));
assert.deepEqual(nativePreciseDateBook.resolve("PreciseDatePivot").filters, preciseDatePivot.filters);
assert.deepEqual(nativePreciseDateBook.resolve("PreciseDatePivot").computedValues(), preciseDatePivot.computedValues());
const nativePreciseRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(nativePreciseDateBook));
assert.deepEqual(nativePreciseRoundtrip.resolve("PreciseDatePivot").filters, preciseDatePivot.filters);
assert.deepEqual(nativePreciseRoundtrip.resolve("PreciseDatePivot").computedValues(), preciseDatePivot.computedValues());
const alternatePrefixPreciseZip = await JSZip.loadAsync(new Uint8Array(await preciseDateXlsx.arrayBuffer()));
alternatePrefixPreciseZip.remove("customXml/open-office-artifact.json");
alternatePrefixPreciseZip.file("xl/pivotTables/pivotTable1.xml", precisePivotXml.replaceAll("xmlns:x14=", "xmlns:p14=").replaceAll("x14:pivotFilter", "p14:pivotFilter").replace('mc:Ignorable="x14"', 'mc:Ignorable="p14"'));
const alternatePrefixPrecise = await SpreadsheetFile.importXlsx(new FileBlob(await alternatePrefixPreciseZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: preciseDateXlsx.type }));
assert.equal(alternatePrefixPrecise.resolve("PreciseDatePivot").filters[0].useWholeDay, false);
const functionBook = Workbook.create();
const functionSheet = functionBook.worksheets.add("Calculated Functions");
functionSheet.getRange("A1:C4").values = [["Region", "Revenue", "Cost"], ["East", 10, 4], ["East", 20, 8], ["West", 5, 0]];
assert.equal(evaluatePivotFormula('LEN("A😀")'), 2);
assert.equal(evaluatePivotFormula('LEN("❤️")'), 2);
assert.equal(evaluatePivotFormula('LEFT("A😀B",2)'), "A😀");
assert.equal(evaluatePivotFormula('RIGHT("A😀B",2)'), "😀B");
assert.equal(evaluatePivotFormula('MID("A😀BC",2,2)'), "😀B");
assert.equal(evaluatePivotFormula('LOWER("MiXeD")'), "mixed");
assert.equal(evaluatePivotFormula('UPPER("MiXeD")'), "MIXED");
assert.equal(evaluatePivotFormula('TRIM("  alpha   beta  ")'), "alpha beta");
assert.equal(evaluatePivotFormula('TRIM("  alpha   beta  ")'), "alpha  beta");
assert.equal(evaluatePivotFormula('LEFT("abc",-1)'), "#VALUE!");
assert.equal(evaluatePivotFormula('RIGHT("abc",0)'), "");
assert.equal(evaluatePivotFormula('MID("abc",0,1)'), "#VALUE!");
assert.equal(evaluatePivotFormula('MID("abc",5,2)'), "");
assert.throws(() => evaluatePivotFormula('MID("abc",1)'), /requires exactly 3 arguments/);
assert.equal(evaluatePivotFormula("DATE(1904,1,1)"), 1462);
assert.equal(evaluatePivotFormula("DATE(1904,1,1)", {}, [], { dateSystem: "1904" }), 0);
assert.equal(evaluatePivotFormula("DATE(1900,2,29)"), 60);
assert.equal(evaluatePivotFormula("DATE(1900,2,29)", {}, [], { dateSystem: "1904" }), "#NUM!");
assert.equal(evaluatePivotFormula("YEAR(60)"), 1900);
assert.equal(evaluatePivotFormula("MONTH(60)"), 2);
assert.equal(evaluatePivotFormula("DAY(60)"), 29);
assert.equal(evaluatePivotFormula("YEAR(0)"), 1900);
assert.equal(evaluatePivotFormula("MONTH(0)"), 1);
assert.equal(evaluatePivotFormula("DAY(0)"), 0);
assert.equal(evaluatePivotFormula("YEAR(DATE(2026,3,31))"), 2026);
assert.equal(evaluatePivotFormula("MONTH(DATE(2026,3,31))"), 3);
assert.equal(evaluatePivotFormula("DAY(DATE(2026,3,31))"), 31);
assert.equal(evaluatePivotFormula("DAY(EDATE(DATE(2024,1,31),1))"), 29);
assert.equal(evaluatePivotFormula("DAY(EDATE(DATE(2023,1,31),1))"), 28);
assert.equal(evaluatePivotFormula("DAY(EDATE(DATE(2024,3,31),-1.9))"), 29);
assert.equal(evaluatePivotFormula("DAY(EOMONTH(DATE(2024,2,10),0))"), 29);
assert.equal(evaluatePivotFormula("EOMONTH(DATE(1900,2,1),0)"), 60);
assert.equal(evaluatePivotFormula("EOMONTH(DATE(1904,2,1),0)", {}, [], { dateSystem: "1904" }), 59);
assert.equal(evaluatePivotFormula("DAYS(DATE(2024,3,1),DATE(2024,2,28))"), 2);
assert.equal(evaluatePivotFormula("DAYS(61,59)"), 2);
assert.equal(evaluatePivotFormula("DAYS(2958465.5,2958465)"), 0);
assert.equal(evaluatePivotFormula("WEEKDAY(DATE(2026,3,31))"), 3);
assert.equal(evaluatePivotFormula("WEEKDAY(DATE(2026,3,31),2)"), 2);
assert.equal(evaluatePivotFormula("WEEKDAY(DATE(2026,3,31),3)"), 1);
assert.equal(evaluatePivotFormula("WEEKDAY(DATE(2026,3,31),12)"), 1);
assert.equal(evaluatePivotFormula("WEEKDAY(DATE(2026,3,31),17)"), 3);
assert.equal(evaluatePivotFormula("WEEKDAY(60)"), 5);
assert.equal(evaluatePivotFormula("WEEKDAY(61)"), 5);
assert.equal(evaluatePivotFormula("WEEKDAY(DATE(1904,1,1))", {}, [], { dateSystem: "1904" }), 6);
assert.equal(evaluatePivotFormula("WEEKDAY(DATE(2026,3,31),4)"), "#NUM!");
assert.equal(evaluatePivotFormula("TIME(12,0,0)"), 0.5);
assert.equal(evaluatePivotFormula("TIME(27,0,0)"), 0.125);
assert.equal(evaluatePivotFormula("TIME(0,750,0)"), 12.5 / 24);
assert.equal(evaluatePivotFormula("TIME(0,0,2000)"), 2000 / 86400);
assert.equal(evaluatePivotFormula("TIME(0.9,0,0)"), 0);
assert.equal(evaluatePivotFormula("HOUR(TIME(16,48,10))"), 16);
assert.equal(evaluatePivotFormula("MINUTE(TIME(16,48,10))"), 48);
assert.equal(evaluatePivotFormula("SECOND(TIME(16,48,10))"), 10);
assert.equal(evaluatePivotFormula("HOUR(DATE(2026,3,31)+TIME(18,45,30))"), 18);
assert.equal(evaluatePivotFormula('HOUR("6:45 PM")'), 18);
assert.equal(evaluatePivotFormula('MINUTE("2026-03-31 18:45:30")'), 45);
assert.equal(evaluatePivotFormula('SECOND("2026-03-31 18:45:30")'), 30);
assert.equal(evaluatePivotFormula('HOUR("not-a-date 18:45:30")'), "#VALUE!");
assert.equal(evaluatePivotFormula("HOUR(-0.5)"), "#VALUE!");
assert.equal(evaluatePivotFormula("TIME(-1,0,0)"), "#NUM!");
assert.equal(evaluatePivotFormula("TIME(32768,0,0)"), "#NUM!");
assert.equal(evaluatePivotFormula("NETWORKDAYS(DATE(2026,3,30),DATE(2026,4,5))"), 5);
assert.equal(evaluatePivotFormula("NETWORKDAYS(DATE(2026,4,5),DATE(2026,3,30))"), -5);
assert.equal(evaluatePivotFormula("NETWORKDAYS(DATE(2026,3,30),DATE(2026,4,5),DATE(2026,4,1))"), 4);
assert.equal(evaluatePivotFormula("NETWORKDAYS.INTL(DATE(2026,3,30),DATE(2026,4,5),11)"), 6);
assert.equal(evaluatePivotFormula('NETWORKDAYS.INTL(DATE(2026,3,30),DATE(2026,4,5),"0000011")'), 5);
assert.equal(evaluatePivotFormula('NETWORKDAYS.INTL(DATE(2026,3,30),DATE(2026,4,5),"1111111")'), 0);
assert.equal(evaluatePivotFormula('NETWORKDAYS.INTL(DATE(2026,3,30),DATE(2026,4,5),"000011")'), "#VALUE!");
assert.equal(evaluatePivotFormula('NETWORKDAYS.INTL(DATE(2026,3,30),DATE(2026,4,5)," 0000011")'), "#VALUE!");
assert.equal(evaluatePivotFormula("NETWORKDAYS.INTL(DATE(2026,3,30),DATE(2026,4,5),0)"), "#NUM!");
assert.equal(evaluatePivotFormula("DAYS(WORKDAY(DATE(2026,3,30),5),DATE(2026,3,30))"), 7);
assert.equal(evaluatePivotFormula("DAYS(WORKDAY(DATE(2026,3,30),5,DATE(2026,4,3)),DATE(2026,3,30))"), 8);
assert.equal(evaluatePivotFormula("DAYS(DATE(2026,3,30),WORKDAY(DATE(2026,3,30),-1))"), 3);
assert.equal(evaluatePivotFormula("WORKDAY(DATE(2026,4,4),0)"), evaluatePivotFormula("DATE(2026,4,4)"));
assert.equal(evaluatePivotFormula("DAYS(WORKDAY.INTL(DATE(2026,3,30),5,11),DATE(2026,3,30))"), 5);
assert.equal(evaluatePivotFormula('DAYS(WORKDAY.INTL(DATE(2026,3,30),5,"0000011"),DATE(2026,3,30))'), 7);
assert.equal(evaluatePivotFormula('WORKDAY.INTL(DATE(2026,3,30),5,"1111111")'), "#VALUE!");
assert.equal(evaluatePivotFormula("NETWORKDAYS(57,62)"), 6);
assert.equal(evaluatePivotFormula("NETWORKDAYS(DATE(1904,1,1),DATE(1904,1,4))", {}, [], { dateSystem: "1904" }), 2);
assert.equal(evaluatePivotFormula("WORKDAY(DATE(9999,12,31),1)"), "#NUM!");
assert.equal(evaluatePivotFormula("EDATE(-1,1)"), "#NUM!");
assert.equal(evaluatePivotFormula("DAYS(2958466,1)"), "#NUM!");
assert.equal(evaluatePivotFormula("DATE(10000,1,1)"), "#NUM!");
assert.equal(evaluatePivotFormula("YEAR(-1)"), "#NUM!");
assert.throws(() => evaluatePivotFormula("DATE(2026,1)", {}, [], { dateSystem: "1900" }), /requires exactly 3 arguments/);
assert.throws(() => evaluatePivotFormula("DATE(2026,1,1)", {}, [], { dateSystem: "unix" }), /dateSystem must be 1900 or 1904/);
const functionPivot = functionSheet.pivotTables.add({
  name: "FunctionPivot",
  sourceRange: "A1:C4",
  targetRange: "E1:L4",
  rowFields: ["Region"],
  calculatedFields: [
    { name: "Margin Ratio", formula: "=ROUND(ABS([Revenue]-[Cost])/MAX([Cost],1),2)" },
    { name: "Guarded Margin", formula: "=IF([Cost]=0,0,ROUND(([Revenue]-[Cost])/[Cost],2))" },
    { name: "Safe Ratio", formula: '=IFERROR([Revenue]/[Cost],"n/a")' },
    { name: "Valid Margin", formula: "=AND(ISNUMBER([Revenue]),NOT([Revenue]<[Cost]))" },
    { name: "Margin Distance", formula: "=IFNA(SQRT(POWER([Revenue]-[Cost],2)),0)" },
    { name: "Revenue Modulus", formula: "=MOD(INT([Revenue]),MAX(SIGN([Cost]),1)+2)" },
    { name: "Scaled Margin", formula: "=PRODUCT([Revenue]-[Cost],2)" },
    { name: "Text Contract", formula: '=IF(AND(LEN(TRIM("  ok  "))=2,UPPER(LEFT("pass",1))="P",LOWER(RIGHT("OK",1))="k",MID("margin",2,2)="ar"),[Revenue]-[Cost],0)' },
  ],
  valueFields: [
    { field: "Margin Ratio", name: "Rounded margin" },
    { field: "Guarded Margin", name: "Guarded margin" },
    { field: "Safe Ratio", name: "Safe ratio" },
    { field: "Valid Margin", name: "Valid margin" },
    { field: "Margin Distance", name: "Margin distance" },
    { field: "Revenue Modulus", name: "Revenue modulus" },
    { field: "Scaled Margin", name: "Scaled margin" },
    { field: "Text Contract", name: "Text contract" },
  ],
});
assert.deepEqual(functionPivot.computedValues(), [["Region", "Rounded margin", "Guarded margin", "Safe ratio", "Valid margin", "Margin distance", "Revenue modulus", "Scaled margin", "Text contract"], ["East", 1.5, 1.5, 2.5, true, 18, 0, 36, 18], ["West", 5, 0, "n/a", true, 5, 2, 10, 5]]);
const functionXlsx = await SpreadsheetFile.exportXlsx(functionBook);
const functionZip = await JSZip.loadAsync(new Uint8Array(await functionXlsx.arrayBuffer()));
const functionPivotCacheXml = await functionZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text");
assert.match(functionPivotCacheXml, /formula="ROUND\(ABS\('Revenue'-'Cost'\)\/MAX\('Cost',1\),2\)"/);
assert.match(functionPivotCacheXml, /formula="IF\('Cost'=0,0,ROUND\(\('Revenue'-'Cost'\)\/'Cost',2\)\)"/);
assert.match(functionPivotCacheXml, /formula="IFERROR\('Revenue'\/'Cost',&quot;n\/a&quot;\)"/);
assert.match(functionPivotCacheXml, /formula="AND\(ISNUMBER\('Revenue'\),NOT\('Revenue'&lt;'Cost'\)\)"/);
assert.match(functionPivotCacheXml, /formula="IFNA\(SQRT\(POWER\('Revenue'-'Cost',2\)\),0\)"/);
assert.match(functionPivotCacheXml, /formula="MOD\(INT\('Revenue'\),MAX\(SIGN\('Cost'\),1\)\+2\)"/);
assert.match(functionPivotCacheXml, /formula="PRODUCT\('Revenue'-'Cost',2\)"/);
assert.match(functionPivotCacheXml, /formula="IF\(AND\(LEN\(TRIM\(&quot;  ok  &quot;\)\)=2,UPPER\(LEFT\(&quot;pass&quot;,1\)\)=&quot;P&quot;,LOWER\(RIGHT\(&quot;OK&quot;,1\)\)=&quot;k&quot;,MID\(&quot;margin&quot;,2,2\)=&quot;ar&quot;\),'Revenue'-'Cost',0\)"/);
functionZip.remove("customXml/open-office-artifact.json");
const nativeFunctionBook = await SpreadsheetFile.importXlsx(new FileBlob(await functionZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: functionXlsx.type }));
assert.deepEqual(nativeFunctionBook.resolve("FunctionPivot").computedValues(), functionPivot.computedValues());
assert.deepEqual(nativeFunctionBook.resolve("FunctionPivot").calculatedFields, functionPivot.calculatedFields);
assert.deepEqual((await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(nativeFunctionBook))).resolve("FunctionPivot").computedValues(), functionPivot.computedValues());
assert.match(JSON.stringify(regionalPivot.layoutJson()), /calculatedFields/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Missing"] }), /not present in the source headers/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], columnFields: ["Region"] }), /both a row and column field/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], filters: { Quarter: ["Q1"] } }), /must also be a row or column field/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], filters: { Region: { include: ["East"], exclude: ["West"] } } }), /exactly one of include or exclude/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], filters: { Region: ["North"] } }), /unknown item North/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], refreshPolicy: { saveData: "yes" } }), /saveData must be a boolean/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], refreshPolicy: { refreshedDateIso: "today" } }), /XML date-time string/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: [{ name: "Revenue", formula: "Cost" }] }), /must be unique/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: [{ name: "Profit", formula: "Revenue - Missing" }] }), /unknown source field Missing/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: [{ name: "Profit", formula: "Revenue +" }] }), /ended unexpectedly/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: [{ name: "Profit", formula: "Revenue-Cost", numFmtId: -1 }] }), /non-negative integer/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: [{ name: "Profit", formula: `Revenue+${"1+".repeat(2050)}1` }] }), /4096 characters/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: Array.from({ length: 129 }, (_, index) => ({ name: `C${index}`, formula: "Revenue" })) }), /exceeds 128 fields/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", rowFields: ["OrderDate"], filters: [{ field: "OrderDate", type: "dateBetween", value1: "2026-03-01" }] }), /value2 must be an ISO date/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", rowFields: ["OrderDate"], filters: [{ field: "OrderDate", type: "dateEqual", value1: "2026-02-30" }] }), /value1 must be an ISO date/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", rowFields: ["OrderDate"], filters: [{ field: "OrderDate", type: "dateBetween", value1: "2026-04-01", value2: "2026-03-01" }] }), /must not be after/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", rowFields: ["OrderDate"], filters: [{ field: "OrderDate", type: "today", value1: "2026-03-01" }] }), /cannot define absolute date values/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", rowFields: ["OrderDate"], filters: [{ field: "OrderDate", type: "thisMonth", asOf: "not-a-date" }] }), /asOf must be an ISO date/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", rowFields: ["OrderDate"], filters: [{ field: "OrderDate", type: "thisDecade" }] }), /unsupported type thisDecade/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", rowFields: ["OrderDate"], filters: [{ field: "OrderDate", type: "dateEqual", value1: "2026-03-01T25:00:00", useWholeDay: false }] }), /value1 must be an ISO date-time/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", rowFields: ["OrderDate"], filters: [{ field: "OrderDate", type: "today", useWholeDay: false }] }), /relative date filter OrderDate requires useWholeDay=true/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: {}, rowFields: ["Region"] }), /groupFields must be an array/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: Array.from({ length: 129 }, (_, index) => ({ name: `Year ${index}`, sourceField: "OrderDate", groupBy: "years" })), rowFields: ["Region"] }), /exceeds 128 fields/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Region", sourceField: "OrderDate", groupBy: "years" }], rowFields: ["Region"] }), /must not replace a source field/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Year", sourceField: "Missing", groupBy: "years" }], rowFields: ["Year"] }), /unknown source field Missing/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Week", sourceField: "OrderDate", groupBy: "weeks" }], rowFields: ["Week"] }), /supported calendar\/time level/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Minute", sourceField: "OrderDate", groupBy: "minutes", range: { groupInterval: 0 } }], rowFields: ["Minute"] }), /integer from 1 to 32767/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Month 1", sourceField: "OrderDate", groupBy: "months" }, { name: "Month 2", sourceField: "OrderDate", groupBy: "months" }], rowFields: ["Month 1"] }), /must not repeat a groupBy/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Period", sourceField: "OrderDate", groupBy: "months" }], rowFields: ["Period"], calculatedFields: [{ name: "Period", formula: "Revenue" }] }), /must not replace a group field/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Year", sourceField: "OrderDate", groupBy: "years" }], rowFields: ["Year"], filters: [{ field: "Year", type: "dateEqual", value1: "2026-01-01" }] }), /must target a source date field/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Month", sourceField: "OrderDate", groupBy: "months" }], rowFields: ["Month"], filters: [{ field: "Month", include: ["Not a month"] }] }), /references unknown item Not a month/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Band", sourceField: "Revenue", groupBy: "range", range: { groupInterval: 0 } }], rowFields: ["Band"] }), /positive groupInterval/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Cluster", sourceField: "Region", groupBy: "discrete", groups: [{ name: "Coasts", items: ["East", "Missing"] }] }], rowFields: ["Cluster"] }), /unknown source item Missing/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Cluster", sourceField: "Region", groupBy: "discrete", groups: [{ name: "One", items: ["East"] }, { name: "Two", items: ["East"] }] }], rowFields: ["Cluster"] }), /only one group/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:U7", targetRange: "W8", groupFields: [{ name: "Cluster", sourceField: "Region", groupBy: "discrete", groups: [{ name: "West", items: ["East"] }] }], rowFields: ["Cluster"] }), /conflicts with an ungrouped source item/);
assert.equal(workbook.resolve(revenuePivot.id).name, "RevenuePivot");
assert.match(workbook.help("sheet.pivotTables.add").ndjson, /pivot table facade/);
assert.match(workbook.help("sheet.pivotTables.add").ndjson, /dateBetween/);
assert.match(workbook.help("sheet.pivotTables.add").ndjson, /groupFields/);
const image = sheet.images.add({
  name: "LogoImage",
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  alt: "Logo placeholder",
  anchor: { from: { row: 8, col: 0 }, extent: { widthPx: 120, heightPx: 80 } },
});
const sparklineGroup = sheet.sparklineGroups.add({
  type: "line",
  targetRange: "H2:H2",
  sourceData: chartSource.getRange("G2:G4"),
  seriesColor: "#0284c7",
  markers: { show: true },
});
const sparklineAlias = sheet.getRange("H3:H3").sparklines.add("column", chartSource.getRange("G2:G4"), { seriesColor: "#f97316" });
assert.equal(image.alt, "Logo placeholder");
assert.deepEqual(sparklineGroup.values(), [100, 120, 130]);
assert.equal(sparklineAlias.type, "column");
assert.match(workbook.help("sheet.images.add").ndjson, /worksheet image/);
assert.match(workbook.help("sheet.sparklineGroups.add").ndjson, /sparklines/);
assert.equal(chartFromRange.series.items[0].name, "Revenue");
assert.match(chartFromRange.series.items[0].formula, /G2:G4/);
assert.equal(chartFromConfig.series.getItemAt(0).values[1], 7);

const metadataInspect = workbook.inspect({ kind: "dataValidation,conditionalFormat,thread,table,pivotTable,drawing", maxChars: 16000 }).ndjson;
assert.match(metadataInspect, /"kind":"dataValidation"/);
assert.match(metadataInspect, /"type":"list"/);
assert.match(metadataInspect, /"kind":"conditionalFormat"/);
assert.match(metadataInspect, /"ruleType":"cellIs"/);
assert.match(metadataInspect, /"ruleType":"expression"/);
assert.match(metadataInspect, /"ruleType":"colorScale"/);
assert.match(metadataInspect, /fee2e2/);
assert.match(metadataInspect, /"kind":"thread"/);
assert.match(metadataInspect, /Formula checks revenue sum/);
assert.match(metadataInspect, /TasksTable/);
assert.match(metadataInspect, /RevenuePivot/);
assert.match(metadataInspect, /Revenue sum/);
assert.match(metadataInspect, /Revenue Trend/);
assert.match(metadataInspect, /ScoresChart/);
assert.match(metadataInspect, /LogoImage/);
assert.match(metadataInspect, /"kind":"sparkline"/);
const definedNameInspect = workbook.inspect({ kind: "definedName", target: revenueName.id, maxChars: 4000 }).ndjson;
assert.match(definedNameInspect, /RevenueData/);
assert.match(definedNameInspect, /Sheet1!G2:G4/);
assert.match(definedNameInspect, /"hidden":false/);
assert.equal(localRevenueName.scope, "Sheet1");
assert.equal(localRevenueName.hidden, true);
assert.equal(workbook.resolve(revenueName.id).name, "RevenueData");
assert.equal(workbook.resolve("RevenueData").refersTo, "Sheet1!G2:G4");
assert.match(workbook.help("workbook.definedNames.add").ndjson, /defined name/);
const targetedImageInspect = workbook.inspect({ kind: "drawing", target: image.id, maxChars: 4000 }).ndjson;
assert.match(targetedImageInspect, /LogoImage/);
assert.doesNotMatch(targetedImageInspect, /Revenue Trend/);
const targetedFormulaInspect = workbook.inspect({ kind: "formula", target: "Sheet1!C2", maxChars: 4000 }).ndjson;
assert.match(targetedFormulaInspect, /=A2\+B2/);
assert.doesNotMatch(targetedFormulaInspect, /=A3\+B3/);
const formulaContextInspect = workbook.inspect({ kind: "formula", target: "Sheet1!C2", after: 1, maxChars: 4000 }).ndjson;
assert.match(formulaContextInspect, /=A2\+B2/);
assert.match(formulaContextInspect, /=A3\+B3/);
const shapedFormulaInspect = workbook.inspect({ kind: "formula", target: "Sheet1!E2", include: "formula,value,precedents", exclude: "dependents", maxChars: 4000 }).ndjson;
assert.match(shapedFormulaInspect, /TasksTable\[Sum\]/);
assert.match(shapedFormulaInspect, /"value":38/);
assert.match(shapedFormulaInspect, /"precedents"/);
assert.doesNotMatch(shapedFormulaInspect, /"dependents"/);
assert.equal(workbook.resolve(thread.id).resolved, true);
assert.equal(workbook.resolve(cf.id).operator, "greaterThan");
assert.equal(workbook.resolve(customCf.id).formula, "=A2<B2");
const computedCellIsStyle = workbook.inspect({ kind: "computedStyle", target: "Sheet1!C3", maxChars: 4000 }).ndjson;
assert.match(computedCellIsStyle, new RegExp(cf.id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.match(computedCellIsStyle, /"fill":"green"/);
const computedExpressionStyle = workbook.inspect({ kind: "computedStyle", target: "Sheet1!A2", maxChars: 4000 }).ndjson;
assert.match(computedExpressionStyle, new RegExp(customCf.id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.match(computedExpressionStyle, /"fill":"sky-100"/);
const colorScaleStyle = workbook.inspect({ kind: "computedStyle", target: "Sheet1!G4", maxChars: 4000 }).ndjson;
assert.match(colorScaleStyle, new RegExp(colorScaleCf.id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.match(colorScaleStyle, /"fill":"#22c55e"/);
assert.equal(workbook.resolve(tasksTable.id).name, "TasksTable");
assert.equal(workbook.resolve(chartFromRange.id).title, "Revenue Trend");
assert.equal(workbook.resolve(image.id).alt, "Logo placeholder");
assert.equal(workbook.resolve(sparklineGroup.id).type, "line");
assert.equal(workbook.resolve(sparklineAlias.id).type, "column");
const trace = workbook.trace("Sheet1!C2");
assert.equal(trace.tree.address, "C2");
assert.equal(trace.tree.value, 5);
assert.deepEqual(trace.tree.precedents.map((node) => node.address), ["A2", "B2"]);
assert.match(trace.ndjson, /"precedents":\["Sheet1!A2","Sheet1!B2"\]/);
const structuredTrace = workbook.trace("Sheet1!E2");
assert.equal(structuredTrace.tree.value, 38);
assert.deepEqual(structuredTrace.tree.precedents.map((node) => node.address), ["C2", "C3", "C4"]);
assert.match(workbook.inspect({ kind: "formula", target: "Sheet1!E2", maxChars: 8000 }).ndjson, /TasksTable\[Sum\]/);
assert.match(workbook.inspect({ kind: "formula", target: "Sheet1!E3", maxChars: 8000 }).ndjson, /RevenueData/);
assert.match(workbook.inspect({ kind: "formulaNode", target: "Sheet1!E3", maxChars: 12000 }).ndjson, /Sheet1!G4/);
assert.match(workbook.inspect({ kind: "formulaNode", target: "Sheet1!E2", maxChars: 12000 }).ndjson, /Sheet1!C4/);
assert.match(workbook.help("workbook.structuredReferences").ndjson, /TableName\[Column\]/);
assert.match(workbook.help("workbook.sharedArrayFormulas").ndjson, /shared formulas/);
assert.match(workbook.help("workbook.trace").ndjson, /precedent tree/);
assert.match(workbook.help("workbook.formulaGraph").ndjson, /dependency graph/);

const structuredBook = Workbook.create();
const structuredSheet = structuredBook.worksheets.add("Structured");
structuredSheet.getRange("A1:C4").values = [["Region", "Revenue", "Cost"], ["East", 10, 4], ["West", 5, 2], ["Total", 15, 6]];
const structuredTable = structuredSheet.tables.add({ range: "A1:C4", name: "SalesTable", showTotals: true });
structuredTable.showTotals = true;
structuredSheet.getRange("D1:E4").formulas = [
  ["=TEXTJOIN(\"|\",TRUE,SalesTable[#Headers])", "=TEXTJOIN(\"|\",TRUE,SalesTable[[#Headers],[Region]:[Cost]])"],
  ["=SUM(SalesTable[[#Data],[Revenue]])", "=SUM(SalesTable[[#Data],[Revenue]:[Cost]])"],
  ["=SUM(SalesTable[[#Totals],[Revenue]])", "=TEXTJOIN(\"|\",TRUE,SalesTable[[#Data],[Region],[Cost]])"],
  ["=SUM(SalesTable[[#All],[Revenue]])", "=SUM(SalesTable[[#Totals],[Revenue]:[Cost]])"],
];
structuredBook.recalculate();
assert.deepEqual(structuredSheet.getRange("D1:E4").values, [["Region|Revenue|Cost", "Region|Revenue|Cost"], [15, 21], [15, "East|4|West|2"], [30, 21]]);
const structuredNode = structuredBook.inspect({ kind: "formulaNode", target: "Structured!D2", maxChars: 12000 }).ndjson;
assert.match(structuredNode, /SalesTable\[\[#Data\],\[Revenue\]\]/);
assert.match(structuredNode, /Structured!B2/);
assert.match(structuredNode, /Structured!B3/);
assert.doesNotMatch(structuredNode, /Structured!B4/);
const structuredRangeNode = structuredBook.inspect({ kind: "formulaNode", target: "Structured!E2", maxChars: 12000 }).ndjson;
assert.match(structuredRangeNode, /SalesTable\[\[#Data\],\[Revenue\]:\[Cost\]\]/);
assert.match(structuredRangeNode, /Structured!B2/);
assert.match(structuredRangeNode, /Structured!C3/);
assert.doesNotMatch(structuredRangeNode, /Structured!A2/);
const structuredUnionNode = structuredBook.inspect({ kind: "formulaNode", target: "Structured!E3", maxChars: 12000 }).ndjson;
assert.match(structuredUnionNode, /SalesTable\[\[#Data\],\[Region\],\[Cost\]\]/);
assert.match(structuredUnionNode, /Structured!A2/);
assert.match(structuredUnionNode, /Structured!C3/);
assert.doesNotMatch(structuredUnionNode, /Structured!B2/);
assert.match(structuredBook.help("workbook.structuredReferences").ndjson, /\[First\]:\[Last\]/);
structuredSheet.getRange("F1:F4").formulas = [
  ["=SUM(SalesTable[[Region]:[Revenue]] SalesTable[[Revenue]:[Cost]])"],
  ["=SUM(SalesTable[[#All],[Revenue]] SalesTable[[#Data],[Revenue]])"],
  ["=SUM(SalesTable[Region] SalesTable[Cost])"],
  ["=TEXTJOIN(\"|\",TRUE,SalesTable[[Region]:[Cost]] SalesTable[[Revenue]:[Cost]] SalesTable[[Region]:[Revenue]])"],
];
structuredSheet.getRange("G1:G2").formulas = [["=SUM(A1:B3 B2:C4)"], ["=SUM(A1:A2 C1:C2)"]];
structuredBook.recalculate();
assert.deepEqual(structuredSheet.getRange("F1:F4").values.flat(), [15, 15, "#NULL!", "10|5"]);
assert.deepEqual(structuredSheet.getRange("G1:G2").values.flat(), [15, "#NULL!"]);
const structuredIntersectionTrace = structuredBook.trace("Structured!F1");
assert.deepEqual(structuredIntersectionTrace.tree.precedents.map((node) => node.address), ["B2", "B3"]);
const structuredIntersectionNode = structuredBook.inspect({ kind: "formulaNode", target: "Structured!F1", maxChars: 12000 }).ndjson;
assert.match(structuredIntersectionNode, /SalesTable\[\[Region\]:\[Revenue\]\] SalesTable\[\[Revenue\]:\[Cost\]\]/);
assert.match(structuredIntersectionNode, /Structured!B2/);
assert.match(structuredIntersectionNode, /Structured!B3/);
assert.doesNotMatch(structuredIntersectionNode, /Structured!A2/);
assert.doesNotMatch(structuredIntersectionNode, /Structured!C2/);
assert.deepEqual(structuredBook.trace("Structured!G1").tree.precedents.map((node) => node.address), ["B2", "B3"]);
const structuredIntersectionXlsx = await SpreadsheetFile.exportXlsx(structuredBook);
const structuredIntersectionZip = await JSZip.loadAsync(new Uint8Array(await structuredIntersectionXlsx.arrayBuffer()));
assert.match(await structuredIntersectionZip.file("xl/worksheets/sheet1.xml").async("text"), /SalesTable\[\[Region\]:\[Revenue\]\] SalesTable\[\[Revenue\]:\[Cost\]\]/);
structuredIntersectionZip.remove("customXml/open-office-artifact.json");
const structuredIntersectionNative = await SpreadsheetFile.importXlsx(new FileBlob(await structuredIntersectionZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: structuredIntersectionXlsx.type }));
assert.deepEqual(structuredIntersectionNative.worksheets.getItem("Structured").getRange("F1:F4").values.flat(), [15, 15, "#NULL!", "10|5"]);
assert.deepEqual(structuredIntersectionNative.trace("Structured!F1").tree.precedents.map((node) => node.address), ["B2", "B3"]);

const escapedStructuredBook = Workbook.create();
const escapedStructuredSheet = escapedStructuredBook.worksheets.add("Escaped");
escapedStructuredSheet.getRange("A1:G4").values = [
  ["Revenue", "Cost", "#Items", "@Rate", "Bracket[Value]", "Owner's", "Net"],
  [10, 4, 2, 0.1, 7, 1, null],
  [5, 2, 3, 0.2, 8, 2, null],
  [12, 6, 1, 0.3, 9, 3, null],
];
escapedStructuredSheet.tables.add({ range: "A1:G4", name: "Specials" });
escapedStructuredSheet.getRange("G2:G4").formulas = [
  ["=[Revenue]-[Cost]+['#Items]+[Owner''s]"],
  ["=Specials[@Revenue]-Specials[@Cost]+Specials[@['@Rate]]"],
  ["=SUM(Specials[[#This Row],[Revenue]:[Bracket'[Value']]])"],
];
escapedStructuredSheet.getRange("H2").formulas = [["=Specials[@Revenue]"]];
escapedStructuredBook.recalculate();
assert.deepEqual(escapedStructuredSheet.getRange("G2:G4").values.flat(), [9, 3.2, 28.3]);
assert.equal(escapedStructuredSheet.getRange("H2").values[0][0], "#VALUE!");
const escapedStructuredG2 = escapedStructuredBook.inspect({ kind: "formulaNode", target: "Escaped!G2", maxChars: 12000 }).ndjson;
assert.match(escapedStructuredG2, /Escaped!A2/);
assert.match(escapedStructuredG2, /Escaped!B2/);
assert.match(escapedStructuredG2, /Escaped!C2/);
assert.match(escapedStructuredG2, /Escaped!F2/);
assert.doesNotMatch(escapedStructuredG2, /Escaped!A3/);
const escapedStructuredTrace = escapedStructuredBook.trace("Escaped!G4");
assert.deepEqual(escapedStructuredTrace.tree.precedents.map((node) => node.address), ["A4", "B4", "C4", "D4", "E4"]);
const escapedStructuredXlsx = await SpreadsheetFile.exportXlsx(escapedStructuredBook);
const escapedStructuredZip = await JSZip.loadAsync(new Uint8Array(await escapedStructuredXlsx.arrayBuffer()));
assert.match(await escapedStructuredZip.file("xl/worksheets/sheet1.xml").async("text"), /Specials\[\[#This Row\],\[Revenue\]:\[Bracket'\[Value'\]\]\]/);
escapedStructuredZip.file("xl/custom/native-table.xml", await escapedStructuredZip.file("xl/tables/table1.xml").async("text"));
escapedStructuredZip.remove("xl/tables/table1.xml");
escapedStructuredZip.file("xl/worksheets/_rels/sheet1.xml.rels", (await escapedStructuredZip.file("xl/worksheets/_rels/sheet1.xml.rels").async("text")).replace("../tables/table1.xml", "../custom/native-table.xml"));
escapedStructuredZip.file("[Content_Types].xml", (await escapedStructuredZip.file("[Content_Types].xml").async("text")).replace("/xl/tables/table1.xml", "/xl/custom/native-table.xml"));
escapedStructuredZip.remove("customXml/open-office-artifact.json");
const escapedStructuredNative = await SpreadsheetFile.importXlsx(new FileBlob(await escapedStructuredZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: escapedStructuredXlsx.type }));
assert.deepEqual(escapedStructuredNative.worksheets.getItem("Escaped").getRange("G2:G4").values.flat(), [9, 3.2, 28.3]);
assert.deepEqual(escapedStructuredNative.worksheets.getItem("Escaped").tables.getItemOrNullObject("Specials").columnNames, ["Revenue", "Cost", "#Items", "@Rate", "Bracket[Value]", "Owner's", "Net"]);
assert.deepEqual(escapedStructuredNative.trace("Escaped!G3").tree.precedents.map((node) => node.address), ["A3", "B3", "D3"]);
assert.match(escapedStructuredNative.help("workbook.structuredReferences").ndjson, /#This Row/);

const graphBook = Workbook.create();
const inputsSheet = graphBook.worksheets.add("Inputs");
inputsSheet.getRange("A1:A3").values = [[2], [4], [6]];
const calcSheet = graphBook.worksheets.add("Calc");
calcSheet.getRange("A1:A5").formulas = [["=SUM(Inputs!A1:A3)"], ["=AVERAGE(Inputs!A1:A3)"], ["=MIN(Inputs!A1:A3)"], ["=MAX(Inputs!A1:A3)"], ["=COUNT(Inputs!A1:A3)"]];
graphBook.recalculate();
assert.deepEqual(calcSheet.getRange("A1:A5").values.flat(), [12, 4, 2, 6, 3]);
const formulaGraph = graphBook.formulaGraph();
assert.equal(formulaGraph.nodes.length, 5);
assert.ok(formulaGraph.edges.some((edge) => edge.from === "Calc!A1" && edge.to === "Inputs!A3"));
assert.match(graphBook.inspect({ kind: "formulaNode,formulaEdge", maxChars: 12000 }).ndjson, /"kind":"formulaEdge"/);
assert.match(graphBook.inspect({ kind: "formula", maxChars: 12000 }).ndjson, /"precedents":\["Inputs!A1","Inputs!A2","Inputs!A3"\]/);

const catalogBook = Workbook.create();
const catalogSheet = catalogBook.worksheets.add("Catalog");
catalogSheet.getRange("A1:D4").values = [["Name", "Score", "Region", "Code"], ["Alpha", 10, "East", "AX-100"], ["Beta", 15, "West", "BX-200"], ["Gamma", 20, "East", "CX-300"]];
catalogSheet.getRange("F1:F39").formulas = [
  ["=IF(B2>9,\"ok\",\"bad\")"],
  ["=ROUND(1.234,2)"],
  ["=COUNTIF(C2:C4,\"East\")"],
  ["=SUMIF(C2:C4,\"East\",B2:B4)"],
  ["=VLOOKUP(\"Beta\",A2:B4,2,FALSE)"],
  ["=XLOOKUP(\"Gamma\",A2:A4,B2:B4,\"missing\")"],
  ["=CONCAT(LEFT(A2,2),\"-\",RIGHT(D2,3))"],
  ["=TEXTJOIN(\"/\",TRUE,A2:A3)"],
  ["=LEN(TRIM(\"  padded  \"))"],
  ["=UPPER(LOWER(\"MiXeD\"))"],
  ["=AND(B2>5,B3>10)"],
  ["=OR(B2>50,B3>10)"],
  ["=NOT(B2<5)"],
  ["=ABS(-4)"],
  ["=CEILING(7,3)"],
  ["=FLOOR(7,3)"],
  ["=INDEX(A2:D4,2,4)"],
  ["=MATCH(\"Gamma\",A2:A4,0)"],
  ["=INDEX(B2:B4,MATCH(\"Beta\",A2:A4,0),1)"],
  ["=MATCH(16,B2:B4,1)"],
  ["=COUNTIFS(C2:C4,\"East\",B2:B4,\">=10\")"],
  ["=SUMIFS(B2:B4,C2:C4,\"East\",B2:B4,\">10\")"],
  ["=SUMPRODUCT(B2:B4,B2:B4)"],
  ["=HLOOKUP(\"Region\",A1:D4,3,FALSE)"],
  ["=IFERROR(XLOOKUP(\"Missing\",A2:A4,B2:B4),\"not found\")"],
  ["=ISNUMBER(B2)"],
  ["=ISTEXT(A2)"],
  ["=ISBLANK(E2)"],
  ["=ISERROR(INDEX(A1:A1,5))"],
  ["=AVERAGEIF(C2:C4,\"East\",B2:B4)"],
  ["=AVERAGEIFS(B2:B4,C2:C4,\"East\",B2:B4,\">10\")"],
  ["=XMATCH(\"Beta\",A2:A4)"],
  ["=XMATCH(\"A*\",A2:A4,2)"],
  ["=XMATCH(\"East\",C2:C4,0,-1)"],
  ["=XMATCH(16,B2:B4,-1)"],
  ["=XMATCH(16,B2:B4,1)"],
  ["=XMATCH(\"Missing\",A2:A4)"],
  ["=XMATCH(15,B2:B4,0,2)"],
  ["=XMATCH(15,B2:B4,0,7)"],
];
catalogBook.recalculate();
assert.deepEqual(catalogSheet.getRange("F1:F39").values.flat(), ["ok", 1.23, 2, 30, 15, 20, "Al-100", "Alpha/Beta", 6, "MIXED", true, true, true, 4, 9, 6, "BX-200", 3, 15, 2, 2, 20, 725, "West", "not found", true, true, true, true, 15, 20, 2, 1, 3, 2, 3, "#N/A", 2, "#VALUE!"]);
assert.match(catalogBook.help("fx.XLOOKUP").ndjson, /lookup/);
assert.match(catalogBook.help("fx.INDEX").ndjson, /1-based row/);
assert.match(catalogBook.help("fx.MATCH").ndjson, /1-based position/);
assert.match(catalogBook.help("fx.XMATCH").ndjson, /reverse search/);
assert.match(catalogBook.help("fx.COUNTIFS").ndjson, /multiple criteria/);
assert.match(catalogBook.help("fx.SUMIFS").ndjson, /all supplied criteria/);
assert.match(catalogBook.help("fx.SUMPRODUCT").ndjson, /corresponding numeric values/);
assert.match(catalogBook.help("fx.HLOOKUP").ndjson, /first row/);
assert.match(catalogBook.help("fx.IFERROR").ndjson, /fallback value/);
assert.match(catalogBook.help("fx.ISERROR").ndjson, /recognized formula error/);
assert.match(catalogBook.help("fx.AVERAGEIFS").ndjson, /all supplied criteria/);
assert.match(catalogBook.help("fx.TEXTJOIN").ndjson, /delimiter/);
assert.match(catalogBook.inspect({ kind: "formula", maxChars: 20000 }).ndjson, /XLOOKUP/);

const criteriaBook = Workbook.create();
const criteriaSheet = criteriaBook.worksheets.add("Criteria");
criteriaSheet.getRange("A1:D8").values = [
  ["Alpha", 1, "East", 1],
  ["ALPINE", 2, "EAST", "#N/A"],
  ["Beta", 3, "West", 3],
  ["A*literal", 4, "East", 4],
  ["A?literal", 5, "East", 5],
  ["", 6, "East", 6],
  [10, 7, "East", 7],
  ["10", 8, "East", 8],
];
criteriaSheet.getRange("F1:F15").formulas = [
  ["=COUNTIF(A1:A8,\"alp*\")"],
  ["=COUNTIF(A1:A8,\"A~*literal\")"],
  ["=COUNTIF(A1:A8,\"A~?literal\")"],
  ["=COUNTIF(A1:A8,\"*\")"],
  ["=COUNTIF(A1:A8,\"10\")"],
  ["=SUMIF(A1:A8,\"alp*\",B1:B8)"],
  ["=SUMIFS(B1:B8,A1:A8,\"a*\",C1:C8,\"east\")"],
  ["=AVERAGEIF(A1:A8,\"alp*\",B1:B8)"],
  ["=COUNTIFS(A1:A2,\"*\",C1:C3,\"east\")"],
  ["=SUMIFS(B1:B2,A1:A3,\"*\")"],
  ["=SUMIF(C1:C8,\"east\",D1:D8)"],
  ["=COUNTIFS(A1:B2,\"*\",A1:D1,\"*\")"],
  ["=SUMIFS(D1:D8,C1:C8,\"east\")"],
  ["=AVERAGEIF(C1:C8,\"east\",D1:D8)"],
  ["=AVERAGEIFS(D1:D8,C1:C8,\"east\")"],
];
criteriaBook.recalculate();
assert.deepEqual(criteriaSheet.getRange("F1:F15").values.flat(), [2, 1, 1, 6, 2, 3, 12, 1.5, "#VALUE!", "#VALUE!", "#N/A", "#VALUE!", "#N/A", "#N/A", "#N/A"]);
assert.match(criteriaBook.help("fx.SUMIFS").ndjson, /wildcard/);
const criteriaRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(criteriaBook));
assert.deepEqual(criteriaRoundtrip.worksheets.getItem("Criteria").getRange("F1:F15").values, criteriaSheet.getRange("F1:F15").values);

const dateBook = Workbook.create();
const dateSheet = dateBook.worksheets.add("Dates");
dateSheet.getRange("A1:A2").values = [[45292], [45299]];
dateSheet.getRange("B1:B32").formulas = [
  ["=DATE(1900,2,29)"],
  ["=YEAR(60)"],
  ["=MONTH(60)"],
  ["=DAY(60)"],
  ["=DATE(2024,2,29)"],
  ["=DATE(2024,13,1)"],
  ["=DATE(2024,1,0)"],
  ["=EDATE(DATE(2024,1,31),1)"],
  ["=EOMONTH(DATE(2024,2,10),0)"],
  ["=DAYS(DATE(2024,3,1),DATE(2024,2,28))"],
  ["=WEEKDAY(DATE(2024,1,1),1)"],
  ["=WEEKDAY(DATE(2024,1,1),2)"],
  ["=WEEKDAY(60,1)"],
  ["=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,10))"],
  ["=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,10),A1)"],
  ["=NETWORKDAYS(DATE(2024,1,10),DATE(2024,1,1),A1)"],
  ["=WORKDAY(DATE(2024,1,5),1)"],
  ["=WORKDAY(DATE(2024,1,5),1,A2)"],
  ["=WORKDAY(DATE(2024,1,8),-1)"],
  ["=DATE(\"bad\",1,1)"],
  ["=YEAR(-1)"],
  ["=WEEKDAY(DATE(2024,1,1),9)"],
  ["=DATE(0,1,1)"],
  ["=EDATE(60,1)"],
  ["=EDATE(60,-1)"],
  ["=EOMONTH(60,0)"],
  ["=DAYS(61,59)"],
  ["=WEEKDAY(61,1)"],
  ["=WEEKDAY(DATE(2024,1,1),3)"],
  ["=WEEKDAY(DATE(2024,1,1),12)"],
  ["=NETWORKDAYS(DATE(2024,1,6),DATE(2024,1,7))"],
  ["=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,2),\"bad\")"],
];
dateBook.recalculate();
assert.deepEqual(dateSheet.getRange("B1:B32").values.flat(), [60, 1900, 2, 29, 45351, 45658, 45291, 45351, 45351, 2, 2, 1, 5, 8, 7, -7, 45299, 45300, 45296, "#VALUE!", "#NUM!", "#NUM!", 1, 89, 29, 60, 2, 5, 0, 7, 0, "#VALUE!"]);
assert.match(dateBook.help("fx.DATE").ndjson, /1900 serial-60 compatibility/);
assert.match(dateBook.help("fx.NETWORKDAYS").ndjson, /optional holidays/);
const dateRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(dateBook));
assert.deepEqual(dateRoundtrip.worksheets.getItem("Dates").getRange("B1:B32").values, dateSheet.getRange("B1:B32").values);

const dateTimeTextBook = Workbook.create();
const dateTimeTextSheet = dateTimeTextBook.worksheets.add("DateTimeText");
dateTimeTextSheet.getRange("A1:A3").values = [["2024-02-29"], ["6:45 PM"], ["1,234.50"]];
dateTimeTextSheet.getRange("B1:B32").formulas = [
  ["=DATEVALUE(A1)"],
  ["=DATEVALUE(\"29-Feb-2024\")"],
  ["=DATEVALUE(\"February 29, 2024 6:45 PM\")"],
  ["=DATEVALUE(\"2024-02-30\")"],
  ["=DATEVALUE(\"2/29/2024\")"],
  ["=DATEVALUE(\"1900-02-29\")"],
  ["=TIME(27,0,0)"],
  ["=TIME(0,750,0)"],
  ["=TIME(0,0,2000)"],
  ["=TIME(-1,0,0)"],
  ["=TIME(32768,0,0)"],
  ["=TIMEVALUE(A2)"],
  ["=TIMEVALUE(\"22-Aug-2008 6:35 AM\")"],
  ["=TIMEVALUE(\"25:00\")"],
  ["=HOUR(0.78125)"],
  ["=MINUTE(\"12:45:18 PM\")"],
  ["=SECOND(\"12:45:18 PM\")"],
  ["=HOUR(DATEVALUE(\"2024-01-01\"))"],
  ["=VALUE(A3)"],
  ["=VALUE(\"(1,234.50)\")"],
  ["=VALUE(\"12.5%\")"],
  ["=VALUE(\"1.25E3\")"],
  ["=VALUE(\"not-a-number\")"],
  ["=VALUE(42)"],
  ["=TIME(1,2)"],
  ["=DATEVALUE(45351)"],
  ["=TIMEVALUE(0.5)"],
  ["=HOUR(-0.25)"],
  ["=TIMEVALUE(\"2024-02-30 6:45 PM\")"],
  ["=HOUR(\"not-a-date 6:45 PM\")"],
  ["=VALUE(\"(12.5%)\")"],
  ["=VALUE(\"(-1)\")"],
];
dateTimeTextBook.recalculate();
assert.deepEqual(dateTimeTextSheet.getRange("B1:B32").values.flat(), [45351, 45351, 45351, "#VALUE!", "#VALUE!", 60, 0.125, 750 / 1440, 2000 / 86400, "#NUM!", "#NUM!", 0.78125, (6 * 3600 + 35 * 60) / 86400, "#VALUE!", 18, 45, 18, 0, 1234.5, -1234.5, 0.125, 1250, "#VALUE!", 42, "#VALUE!", "#VALUE!", "#VALUE!", "#VALUE!", "#VALUE!", "#VALUE!", -0.125, "#VALUE!"]);
assert.match(dateTimeTextBook.help("fx.DATEVALUE").ndjson, /ambiguous locale-numeric dates/);
assert.match(dateTimeTextBook.help("fx.TIME").ndjson, /wrapping at 24 hours/);
assert.match(dateTimeTextBook.help("fx.TIMEVALUE").ndjson, /12-hour or 24-hour/);
assert.match(dateTimeTextBook.help("fx.HOUR").ndjson, /0 through 23/);
assert.match(dateTimeTextBook.help("fx.VALUE").ndjson, /accounting parentheses/);
const dateTimeTextRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(dateTimeTextBook));
assert.deepEqual(dateTimeTextRoundtrip.worksheets.getItem("DateTimeText").getRange("B1:B32").values, dateTimeTextSheet.getRange("B1:B32").values);

const date1904Book = Workbook.create({ dateSystem: "1904" });
assert.equal(date1904Book.dateSystem, "1904");
assert.equal(Workbook.create({ date1904: true }).dateSystem, "1904");
assert.equal(Workbook.create().setDateSystem(true).dateSystem, "1904");
assert.throws(() => Workbook.create({ dateSystem: "unix" }), /expected 1900 or 1904/);
const date1904Sheet = date1904Book.worksheets.add("Dates1904");
date1904Sheet.getRange("A1:A20").formulas = [
  ["=DATE(1904,1,1)"],
  ["=DATE(1904,1,2)"],
  ["=DATE(2024,2,29)"],
  ["=YEAR(0)"],
  ["=MONTH(0)"],
  ["=DAY(0)"],
  ["=EDATE(0,1)"],
  ["=EOMONTH(0,0)"],
  ["=DAYS(DATE(2024,3,1),DATE(2024,2,28))"],
  ["=WEEKDAY(0,1)"],
  ["=WEEKDAY(0,2)"],
  ["=NETWORKDAYS(DATE(1904,1,1),DATE(1904,1,10))"],
  ["=WORKDAY(0,1)"],
  ["=DATE(1900,1,1)"],
  ["=DATE(9999,12,31)"],
  ["=DATE(1904,2,29)"],
  ["=DATEVALUE(\"1904-01-01\")"],
  ["=DATEVALUE(\"2024-02-29\")"],
  ["=DATEVALUE(\"1900-02-29\")"],
  ["=HOUR(TIME(27,0,0))"],
];
date1904Book.recalculate();
assert.deepEqual(date1904Sheet.getRange("A1:A20").values.flat(), [0, 1, 43889, 1904, 1, 1, 31, 30, 2, 6, 5, 6, 3, "#NUM!", 2957003, 59, 0, 43889, "#VALUE!", 3]);
const date1904Inspect = date1904Book.inspect({ kind: "workbook" });
assert.match(date1904Inspect.ndjson, /"dateSystem":"1904"/);
assert.match(date1904Inspect.ndjson, /"date1904":true/);
assert.match(date1904Book.help("workbook.setDateSystem").ndjson, /workbookPr/);
const date1904Xlsx = await SpreadsheetFile.exportXlsx(date1904Book);
const date1904Zip = await JSZip.loadAsync(new Uint8Array(await date1904Xlsx.arrayBuffer()));
const date1904WorkbookXml = await date1904Zip.file("xl/workbook.xml").async("text");
assert.match(date1904WorkbookXml, /<workbookPr date1904="1"\/>/);
assert.match(await date1904Zip.file("customXml/open-office-artifact.json").async("text"), /"dateSystem": "1904"/);
const date1904Roundtrip = await SpreadsheetFile.importXlsx(date1904Xlsx);
assert.equal(date1904Roundtrip.dateSystem, "1904");
assert.deepEqual(date1904Roundtrip.worksheets.getItem("Dates1904").getRange("A1:A20").values, date1904Sheet.getRange("A1:A20").values);
const date1904TrueXlsx = await SpreadsheetFile.patchXlsx(date1904Xlsx, [{ path: "xl/workbook.xml", xml: date1904WorkbookXml.replace('date1904="1"', 'date1904="true"') }]);
assert.equal((await SpreadsheetFile.importXlsx(date1904TrueXlsx)).dateSystem, "1904");
const date1900PatchedXlsx = await SpreadsheetFile.patchXlsx(date1904Xlsx, [{ path: "xl/workbook.xml", xml: date1904WorkbookXml.replace('date1904="1"', 'date1904="0"') }]);
const date1900PatchedBook = await SpreadsheetFile.importXlsx(date1900PatchedXlsx);
assert.equal(date1900PatchedBook.dateSystem, "1900");
assert.equal(date1900PatchedBook.worksheets.getItem("Dates1904").getRange("A3").values[0][0], 45351);
const invalidDateSystemBook = Workbook.create();
invalidDateSystemBook.worksheets.add("Invalid").getRange("A1").values = [[1]];
invalidDateSystemBook.dateSystem = "invalid";
assert.ok(invalidDateSystemBook.verify().issues.some((issue) => issue.type === "invalidDateSystem"));
await assert.rejects(() => SpreadsheetFile.exportXlsx(invalidDateSystemBook), /expected 1900 or 1904/);

const intlDateBook = Workbook.create();
const intlDateSheet = intlDateBook.worksheets.add("IntlDates");
intlDateSheet.getRange("A1:A2").values = [[45293], [-1]];
intlDateSheet.getRange("B1:B26").formulas = [
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7))"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),2)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),11)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),\"0000011\")"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),\"0010001\")"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),\"1111111\")"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,7),DATE(2024,1,1),1)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),1,A1)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),\"bad\")"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),0)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),1.5)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),1,A2)"],
  ["=WORKDAY.INTL(DATE(2024,1,1),5)"],
  ["=WORKDAY.INTL(DATE(2024,1,1),1,2)"],
  ["=WORKDAY.INTL(DATE(2024,1,1),3,\"0010001\")"],
  ["=WORKDAY.INTL(DATE(2024,1,1),1,1,A1)"],
  ["=WORKDAY.INTL(DATE(2024,1,8),-5)"],
  ["=WORKDAY.INTL(DATE(2024,1,6),0)"],
  ["=WORKDAY.INTL(DATE(2024,1,1),1,\"1111111\")"],
  ["=WORKDAY.INTL(DATE(2024,1,1),1,8)"],
  ["=WORKDAY.INTL(DATE(2024,1,1),1,1,A2)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),17)"],
  ["=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))"],
  ["=WORKDAY.INTL(DATE(2012,1,1),30,17)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),\"0000000\")"],
  ["=WORKDAY.INTL(DATE(2024,1,1),7,\"0000000\")"],
];
intlDateBook.recalculate();
assert.deepEqual(intlDateSheet.getRange("B1:B26").values.flat(), [5, 5, 6, 5, 5, 0, -5, 4, "#VALUE!", "#NUM!", "#NUM!", "#NUM!", 45299, 45293, 45296, 45294, 45292, 45297, "#VALUE!", "#NUM!", "#NUM!", 6, 22, 40944, 7, 45299]);
assert.match(intlDateBook.help("fx.NETWORKDAYS.INTL").ndjson, /seven-character custom weekend/);
assert.match(intlDateBook.help("fx.WORKDAY.INTL").ndjson, /numbered or Monday-first/);

const intl1904Book = Workbook.create({ dateSystem: "1904" });
const intl1904Sheet = intl1904Book.worksheets.add("Intl1904");
intl1904Sheet.getRange("A1:A3").formulas = [
  ["=NETWORKDAYS.INTL(0,6,7)"],
  ["=WORKDAY.INTL(0,1,7)"],
  ["=WORKDAY.INTL(0,1,11)"],
];
intl1904Book.recalculate();
assert.deepEqual(intl1904Sheet.getRange("A1:A3").values.flat(), [5, 2, 1]);
const intl1904Roundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(intl1904Book));
assert.equal(intl1904Roundtrip.dateSystem, "1904");
assert.deepEqual(intl1904Roundtrip.worksheets.getItem("Intl1904").getRange("A1:A3").values, [[5], [2], [1]]);

const formulaEdgeBook = Workbook.create();
const formulaEdgeSheet = formulaEdgeBook.worksheets.add("FormulaEdges");
formulaEdgeSheet.getRange("A1:B3").values = [[1, 10], [2, 20], [3, 30]];
formulaEdgeSheet.getRange("D1:D15").formulas = [
  ["=SUMPRODUCT(A1:B2,A1:A4)"],
  ["=AVERAGEIF(A1:A3,\">10\",B1:B3)"],
  ["=IFERROR(AVERAGEIF(A1:A3,\">10\",B1:B3),\"empty\")"],
  ["=HLOOKUP(1,A1:B3,5,FALSE)"],
  ["=ISERROR(#N/A)"],
  ["=ISTEXT(\"\")"],
  ["=ISBLANK(\"\")"],
  ["=NOT(TRUE)"],
  ["=NA()"],
  ["=IFNA(NA(),\"missing\")"],
  ["=IFNA(#VALUE!,\"missing\")"],
  ["=ISNA(NA())"],
  ["=ISNA(#REF!)"],
  ["=ISERR(#REF!)"],
  ["=ISERR(NA())"],
];
formulaEdgeBook.recalculate();
assert.deepEqual(formulaEdgeSheet.getRange("D1:D15").values.flat(), ["#VALUE!", "#DIV/0!", "empty", "#REF!", true, true, false, false, "#N/A", "missing", "#VALUE!", true, false, true, false]);

const statisticsBook = Workbook.create();
const statisticsSheet = statisticsBook.worksheets.add("Statistics");
statisticsSheet.getRange("A1:A6").values = [[10], [20], [20], [40], ["text"], [null]];
statisticsSheet.getRange("D1:D15").formulas = [
  ["=MEDIAN(A1:A6)"],
  ["=MODE.SNGL(A1:A6)"],
  ["=LARGE(A1:A6,2)"],
  ["=SMALL(A1:A6,2)"],
  ["=RANK.EQ(20,A1:A6)"],
  ["=RANK.EQ(20,A1:A6,1)"],
  ["=ROUND(1250,-2)"],
  ["=ROUND(-1250,-2)"],
  ["=ROUNDUP(12.341,2)"],
  ["=ROUNDUP(-1250,-2)"],
  ["=ROUNDDOWN(12.349,2)"],
  ["=ROUNDDOWN(-1250,-2)"],
  ["=LARGE(A1:A6,9)"],
  ["=MODE.SNGL(1,2,3)"],
  ["=MEDIAN(A1:A6,#REF!)"],
];
statisticsBook.recalculate();
assert.deepEqual(statisticsSheet.getRange("D1:D15").values.flat(), [20, 20, 20, 20, 2, 2, 1300, -1300, 12.35, -1300, 12.34, -1200, "#NUM!", "#N/A", "#REF!"]);
statisticsSheet.getRange("E1:E2").formulas = [["=ROUND(1.005,2)"], ["=MODE.SNGL(3,3,2,2)"]];
statisticsBook.recalculate();
assert.deepEqual(statisticsSheet.getRange("E1:E2").values.flat(), [1.01, 2]);
assert.match(statisticsBook.help("fx.MEDIAN").ndjson, /middle numeric value/);
assert.match(statisticsBook.help("fx.MODE.SNGL").ndjson, /most frequently/);
assert.match(statisticsBook.help("fx.RANK.EQ").ndjson, /descending by default/);
assert.match(statisticsBook.help("fx.ROUNDUP").ndjson, /away from zero/);
assert.match(statisticsBook.inspect({ kind: "formula", target: "Statistics!D1", maxChars: 4000 }).ndjson, /MEDIAN/);

const spillBook = Workbook.create();
const spillSheet = spillBook.worksheets.add("Spill");
spillSheet.getRange("A1").formulas = [["=SEQUENCE(2,3,10,2)"]];
spillBook.recalculate();
assert.deepEqual(spillSheet.getRange("A1:C2").values, [[10, 12, 14], [16, 18, 20]]);
const spillInspect = spillBook.inspect({ kind: "formula", target: "Spill!A1", maxChars: 8000 }).ndjson;
assert.match(spillInspect, /"spillRange":"A1:C2"/);
assert.match(spillInspect, /"spillValues":\[\[10,12,14\],\[16,18,20\]\]/);
const spillLayout = spillBook.layoutJson({ sheetName: "Spill", target: "Spill!B2" });
assert.equal(spillLayout.sheets[0].cells[0].spillParent, "Spill!A1");
spillSheet.getRange("E1").formulas = [["=TRANSPOSE(A1:C2)"]];
spillBook.recalculate();
assert.deepEqual(spillSheet.getRange("E1:F3").values, [[10, 16], [12, 18], [14, 20]]);
spillSheet.getRange("H1:J4").values = [["Item", "Region", "Score"], ["Alpha", "East", 10], ["Beta", "West", 15], ["Gamma", "East", 20]];
spillSheet.getRange("L1").formulas = [["=FILTER(H2:J4,I2:I4=\"East\")"]];
spillSheet.getRange("P1").formulas = [["=UNIQUE(I2:I4)"]];
spillSheet.getRange("R1").formulas = [["=SORT(H2:J4,3,-1)"]];
spillSheet.getRange("V1").formulas = [["=TAKE(H2:J4,2,-2)"]];
spillSheet.getRange("Y1").formulas = [["=DROP(H2:J4,1,1)"]];
spillSheet.getRange("AB1").formulas = [["=CHOOSECOLS(H2:J4,3,1)"]];
spillSheet.getRange("AE1").formulas = [["=CHOOSEROWS(H2:J4,3,1)"]];
spillSheet.getRange("AI1").formulas = [["=TAKE(H2:J4,,2)"]];
spillSheet.getRange("AL1").formulas = [["=DROP(H2:J4,,1)"]];
spillSheet.getRange("AO1:AP1").formulas = [["=TAKE(H2:J4,0)", "=CHOOSECOLS(H2:J4,0)"]];
spillSheet.getRange("AR1").formulas = [["=TOCOL(H2:I3)"]];
spillSheet.getRange("AT1").formulas = [["=TOROW(H2:I3,0,TRUE)"]];
spillSheet.getRange("AY1").formulas = [["=WRAPROWS(H2:H4,2)"]];
spillSheet.getRange("BB1").formulas = [["=WRAPCOLS(H2:H4,2,\"pad\")"]];
spillSheet.getRange("BE1:BF1").formulas = [["=WRAPROWS(H2:I3,2)", "=WRAPCOLS(H2:H4,0)"]];
spillSheet.getRange("H6:I7").values = [[1, null], ["#N/A", 4]];
spillSheet.getRange("BH1").formulas = [["=TOCOL(H6:I7,3)"]];
spillSheet.getRange("BJ1").formulas = [["=HSTACK(H2:H4,I2:I3)"]];
spillSheet.getRange("BM1").formulas = [["=VSTACK(H2:I3,H4:H4)"]];
spillSheet.getRange("BQ1").formulas = [["=EXPAND(H2:I3,3,3,\"pad\")"]];
spillSheet.getRange("BU1").formulas = [["=EXPAND(H2:I3,1,1)"]];
spillSheet.getRange("BW1").formulas = [["=EXPAND(H2:I3,,3)"]];
spillBook.recalculate();
assert.deepEqual(spillSheet.getRange("L1:N2").values, [["Alpha", "East", 10], ["Gamma", "East", 20]]);
assert.deepEqual(spillSheet.getRange("P1:P2").values, [["East"], ["West"]]);
assert.deepEqual(spillSheet.getRange("R1:T3").values, [["Gamma", "East", 20], ["Beta", "West", 15], ["Alpha", "East", 10]]);
assert.deepEqual(spillSheet.getRange("V1:W2").values, [["East", 10], ["West", 15]]);
assert.deepEqual(spillSheet.getRange("Y1:Z2").values, [["West", 15], ["East", 20]]);
assert.deepEqual(spillSheet.getRange("AB1:AC3").values, [[10, "Alpha"], [15, "Beta"], [20, "Gamma"]]);
assert.deepEqual(spillSheet.getRange("AE1:AG2").values, [["Gamma", "East", 20], ["Alpha", "East", 10]]);
assert.deepEqual(spillSheet.getRange("AI1:AJ3").values, [["Alpha", "East"], ["Beta", "West"], ["Gamma", "East"]]);
assert.deepEqual(spillSheet.getRange("AL1:AM3").values, [["East", 10], ["West", 15], ["East", 20]]);
assert.deepEqual(spillSheet.getRange("AO1:AP1").values, [["#CALC!", "#VALUE!"]]);
assert.deepEqual(spillSheet.getRange("AR1:AR4").values, [["Alpha"], ["East"], ["Beta"], ["West"]]);
assert.deepEqual(spillSheet.getRange("AT1:AW1").values, [["Alpha", "Beta", "East", "West"]]);
assert.deepEqual(spillSheet.getRange("AY1:AZ2").values, [["Alpha", "Beta"], ["Gamma", "#N/A"]]);
assert.deepEqual(spillSheet.getRange("BB1:BC2").values, [["Alpha", "Gamma"], ["Beta", "pad"]]);
assert.deepEqual(spillSheet.getRange("BE1:BF1").values, [["#VALUE!", "#NUM!"]]);
assert.deepEqual(spillSheet.getRange("BH1:BH2").values, [[1], [4]]);
assert.deepEqual(spillSheet.getRange("BJ1:BK3").values, [["Alpha", "East"], ["Beta", "West"], ["Gamma", "#N/A"]]);
assert.deepEqual(spillSheet.getRange("BM1:BN3").values, [["Alpha", "East"], ["Beta", "West"], ["Gamma", "#N/A"]]);
assert.deepEqual(spillSheet.getRange("BQ1:BS3").values, [["Alpha", "East", "pad"], ["Beta", "West", "pad"], ["pad", "pad", "pad"]]);
assert.equal(spillSheet.getRange("BU1").values[0][0], "#VALUE!");
assert.deepEqual(spillSheet.getRange("BW1:BY2").values, [["Alpha", "East", "#N/A"], ["Beta", "West", "#N/A"]]);
assert.match(spillBook.inspect({ kind: "formula", target: "Spill!L1", maxChars: 8000 }).ndjson, /"spillRange":"L1:N2"/);
assert.match(spillBook.help("fx.SEQUENCE").ndjson, /dynamic array/);
assert.match(spillBook.help("fx.TRANSPOSE").ndjson, /spillRange/);
assert.match(spillBook.help("fx.FILTER").ndjson, /include array/);
assert.match(spillBook.help("fx.UNIQUE").ndjson, /unique rows/);
assert.match(spillBook.help("fx.SORT").ndjson, /column index/);
assert.match(spillBook.help("fx.TAKE").ndjson, /start or end/);
assert.match(spillBook.help("fx.DROP").ndjson, /spill the remainder/);
assert.match(spillBook.help("fx.CHOOSECOLS").ndjson, /Select and reorder/);
assert.match(spillBook.help("fx.CHOOSEROWS").ndjson, /Select and reorder/);
assert.match(spillBook.help("fx.TOCOL").ndjson, /one spilled column/);
assert.match(spillBook.help("fx.TOROW").ndjson, /one spilled row/);
assert.match(spillBook.help("fx.WRAPROWS").ndjson, /requested width/);
assert.match(spillBook.help("fx.WRAPCOLS").ndjson, /requested height/);
assert.match(spillBook.help("fx.HSTACK").ndjson, /horizontally/);
assert.match(spillBook.help("fx.VSTACK").ndjson, /vertically/);
assert.match(spillBook.help("fx.EXPAND").ndjson, /requested row/);
const spillXlsx = await SpreadsheetFile.exportXlsx(spillBook);
const spillZip = await JSZip.loadAsync(new Uint8Array(await spillXlsx.arrayBuffer()));
const spillWorksheetXml = await spillZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(spillWorksheetXml, /<f t="array" ref="A1:C2">SEQUENCE\(2,3,10,2\)<\/f>/);
assert.match(spillWorksheetXml, /<f t="array" ref="E1:F3">TRANSPOSE\(A1:C2\)<\/f>/);
assert.match(spillWorksheetXml, /<f t="array" ref="L1:N2">FILTER\(H2:J4,I2:I4="East"\)<\/f>/);
assert.match(spillWorksheetXml, /<f t="array" ref="V1:W2">TAKE\(H2:J4,2,-2\)<\/f>/);
const blockedSpillBook = Workbook.create();
const blockedSheet = blockedSpillBook.worksheets.add("Blocked");
blockedSheet.getRange("C1").values = [["blocked"]];
blockedSheet.getRange("A1").formulas = [["=SEQUENCE(1,3)"]];
blockedSpillBook.recalculate();
assert.equal(blockedSheet.getRange("A1").values[0][0], "#SPILL!");
assert.match(blockedSpillBook.verify().ndjson, /formulaError/);

const cycleBook = Workbook.create();
const cycleSheet = cycleBook.worksheets.add("Cycle");
cycleSheet.getRange("A1").formulas = [["=B1+1"]];
cycleSheet.getRange("B1").formulas = [["=A1+1"]];
cycleBook.recalculate();
assert.equal(cycleSheet.getRange("A1").values[0][0], "#CYCLE!");
const cycleGraph = cycleBook.formulaGraph();
assert.equal(cycleGraph.cycles.length, 1);
assert.match(cycleBook.inspect({ kind: "formulaGraph,formulaCycle" }).ndjson, /formulaCycle/);
assert.match(cycleBook.verify().ndjson, /formulaCycle/);

const missingSheetBook = Workbook.create();
missingSheetBook.worksheets.add("Main").getRange("A1").formulas = [["=Missing!A1+1"]];
assert.match(missingSheetBook.verify().ndjson, /missingFormulaSheet/);

const preview = await workbook.render({ sheetName: "Sheet1", range: "A1:C3" });
assert.equal(preview.type, "image/svg+xml");
const previewSvg = await preview.text();
assert.match(previewSvg, /<svg/);
assert.match(previewSvg, /TasksTable/);
assert.match(previewSvg, /Revenue Trend/);
assert.match(previewSvg, /Logo placeholder/);
assert.match(previewSvg, /RevenuePivot/);
assert.match(previewSvg, /Revenue sum/);
assert.match(previewSvg, /polyline/);
const layoutBlob = await workbook.render({ format: "layout", sheetName: "Sheet1", range: "A1:C3" });
assert.equal(layoutBlob.type, "application/vnd.open-office-artifact.layout+json");
const layout = JSON.parse(await layoutBlob.text());
assert.equal(layout.kind, "workbookLayout");
assert.equal(layout.sheets[0].name, "Sheet1");
assert.equal(layout.sheets[0].bounds.address, "A1:C3");
assert.ok(layout.sheets[0].cells.some((cell) => cell.address === "C2" && cell.formula === "=A2+B2"));
const c3LayoutCell = layout.sheets[0].cells.find((cell) => cell.address === "C3");
assert.equal(c3LayoutCell.computedStyle.fill, "green");
assert.deepEqual(c3LayoutCell.conditionalFormats.map((item) => item.id), [cf.id]);
const a2LayoutCell = layout.sheets[0].cells.find((cell) => cell.address === "A2");
assert.equal(a2LayoutCell.computedStyle.fill, "sky-100");
assert.deepEqual(a2LayoutCell.conditionalFormats.map((item) => item.id), [customCf.id]);
assert.match(previewSvg, /fill="#22c55e"/);
const scalePreviewSvg = await (await workbook.render({ sheetName: "Sheet1", range: "G2:G4" })).text();
assert.match(scalePreviewSvg, /fill="#22c55e"/);
const scaleLayout = workbook.layoutJson({ sheetName: "Sheet1", range: "G2:G4", target: "Sheet1!G4" });
assert.equal(scaleLayout.sheets[0].cells[0].computedStyle.fill, "#22c55e");
assert.ok(layout.sheets[0].tables.some((table) => table.name === "TasksTable"));
assert.ok(layout.sheets[0].pivots.some((pivot) => pivot.name === "RevenuePivot" && pivot.values.some((row) => row.includes("Revenue sum"))));
assert.ok(layout.sheets[0].charts.some((chart) => chart.title === "Revenue Trend"));
assert.ok(layout.sheets[0].images.some((item) => item.alt === "Logo placeholder"));
assert.ok(layout.sheets[0].sparklines.some((item) => item.targetRange === "H2:H2"));
const cellLayoutBlob = await workbook.render({ format: "layout", sheetName: "Sheet1", range: "A1:C3", target: "Sheet1!C2" });
assert.equal(cellLayoutBlob.metadata.target, "Sheet1!C2");
const cellLayout = JSON.parse(await cellLayoutBlob.text());
assert.deepEqual(cellLayout.sheets[0].cells.map((cell) => cell.address), ["C2"]);
assert.equal(cellLayout.sheets[0].tables.length, 0);
const imageLayout = workbook.layoutJson({ sheetName: "Sheet1", range: "A1:C3", target: image.id });
assert.deepEqual(imageLayout.sheets[0].images.map((item) => item.id), [image.id]);
assert.equal(imageLayout.sheets[0].cells.length, 0);
const chartSearchLayout = workbook.layoutJson({ sheetName: "Sheet1", search: "Revenue Trend" });
assert.deepEqual(chartSearchLayout.sheets[0].charts.map((item) => item.id), [chartFromRange.id]);
const workbookContextLayout = workbook.layoutJson({ sheetName: "Sheet1", target: image.id, before: 1 });
assert.deepEqual(workbookContextLayout.sheets[0].charts.map((item) => item.id), [chartFromConfig.id]);
assert.deepEqual(workbookContextLayout.sheets[0].images.map((item) => item.id), [image.id]);
const sheetTargetLayout = workbook.layoutJson({ target: "Sheet1" });
assert.equal(sheetTargetLayout.sheets.length, 1);
assert.ok(sheetTargetLayout.sheets[0].cells.length > 0);
assert.match(workbook.help("workbook.layoutJson").ndjson, /target\/search context slicing/);

const delimitedBook = Workbook.create();
const delimitedSheet = delimitedBook.worksheets.add("Data");
delimitedSheet.getRange("A1:D3").values = [
  ["Region", "Amount", "Note", "Detail"],
  ["North, East", 42, 'He said "ok"', "line one\nline two"],
  ["=literal", null, "plain", "tab\there"],
];
delimitedSheet.getRange("B3").formulas = [["=B2*2"]];
delimitedBook.recalculate();
const csv = await SpreadsheetFile.exportCsv(delimitedBook, { sheetName: "Data" });
assert.equal(csv.type, "text/csv");
assert.equal(csv.metadata.rows, 3);
assert.equal(csv.metadata.columns, 4);
assert.match(await csv.text(), /"North, East",42,"He said ""ok""","line one\nline two"/);
assert.doesNotMatch(await csv.text(), /=B2\*2/);
const csvInspect = await SpreadsheetFile.inspectDelimited(csv);
assert.equal(csvInspect.summary.rows, 3);
assert.equal(csvInspect.summary.columns, 4);
assert.equal(csvInspect.summary.quotedCells, 3);
assert.equal(csvInspect.summary.formulaLikeCells, 1);
const csvImported = await SpreadsheetFile.importCsv(csv, { sheetName: "Imported", coerceTypes: true });
assert.deepEqual(csvImported.worksheets.getItem("Imported").getRange("A2:D3").values, [
  ["North, East", 42, 'He said "ok"', "line one\nline two"],
  ["=literal", 84, "plain", "tab\there"],
]);
const formulaCsv = await SpreadsheetFile.exportCsv(delimitedBook, { sheetName: "Data", range: "A2:B3", formulas: true, includeBom: true, lineEnding: "\n" });
assert.equal(formulaCsv.bytes[0], 0xef);
assert.match(await formulaCsv.text(), /=B2\*2/);
const tsv = await SpreadsheetFile.exportTsv(delimitedBook, { sheetName: "Data", includeBom: true });
assert.equal(tsv.type, "text/tab-separated-values");
assert.equal((await SpreadsheetFile.inspectDelimited(tsv)).summary.hasBom, true);
assert.equal((await SpreadsheetFile.inspectDelimited(tsv)).summary.columns, 4);
const tsvImported = await SpreadsheetFile.importTsv(tsv, { coerceTypes: true });
assert.equal(tsvImported.worksheets.getItemAt(0).getRange("D3").values[0][0], "tab\there");
await assert.rejects(() => SpreadsheetFile.importCsv('a,"unterminated'), /unterminated quoted field/);
await assert.rejects(() => SpreadsheetFile.importCsv('a,"quoted"junk'), /unexpected content after a closing quote/);
await assert.rejects(() => SpreadsheetFile.inspectDelimited(csv, { maxBytes: 4 }), /exceeds maxBytes/);
await assert.rejects(() => SpreadsheetFile.exportCsv(delimitedBook, { maxBytes: 4 }), /exceeds maxBytes/);
await assert.rejects(() => SpreadsheetFile.exportCsv(delimitedBook, { maxRows: 2 }), /exceeds maxRows/);
await assert.rejects(() => SpreadsheetFile.exportCsv(delimitedBook, { maxColumns: 3 }), /exceeds maxColumns/);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
assert.equal(xlsx.type, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
const xlsxInspect = await SpreadsheetFile.inspectXlsx(xlsx, { includeText: true, maxChars: 16000 });
assert.equal(xlsxInspect.records[0].kind, "xlsxPackage");
assert.equal(xlsxInspect.ok, true);
assert.deepEqual(xlsxInspect.issues, []);
assert.equal(xlsxInspect.records[0].sheets, 1);
assert.ok(xlsxInspect.records[0].uncompressedBytes > 0);
assert.ok(xlsxInspect.records[0].relationshipReferences > 0);
assert.equal(xlsxInspect.records[0].relationshipReferenceIssues, 0);
assert.equal(xlsxInspect.records[0].semanticValidation, true);
assert.equal(xlsxInspect.records[0].semanticIssues, 0);
assert.ok(xlsxInspect.parts.some((part) => part.path === "xl/workbook.xml" && part.contentType.includes("spreadsheetml.sheet.main+xml")));
const xlsxReferenceZip = await JSZip.loadAsync(new Uint8Array(await xlsx.arrayBuffer()));
const xlsxSheetXml = await xlsxReferenceZip.file("xl/worksheets/sheet1.xml").async("text");
const singleQuotedContentTypesXml = (await xlsxReferenceZip.file("[Content_Types].xml").async("text")).replace(/="([^"]*)"/g, "='$1'");
const singleQuotedWorkbookRelsXml = (await xlsxReferenceZip.file("xl/_rels/workbook.xml.rels").async("text")).replace(/="([^"]*)"/g, "='$1'");
const singleQuotedXlsx = await SpreadsheetFile.patchXlsx(xlsx, [
  { path: "[Content_Types].xml", xml: singleQuotedContentTypesXml },
  { path: "xl/_rels/workbook.xml.rels", xml: singleQuotedWorkbookRelsXml },
]);
assert.equal((await SpreadsheetFile.inspectXlsx(singleQuotedXlsx)).ok, true);
const brokenXlsxReferenceXml = xlsxSheetXml.replace(/<\/worksheet>\s*$/, '<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdMissingSourceReference"/></worksheet>');
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/sheet1.xml", xml: brokenXlsxReferenceXml }]), /invalid OOXML package.*relationshipReferenceIdNotFound/);
const invalidReferenceXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/sheet1.xml", xml: brokenXlsxReferenceXml }], { validateResult: false });
const invalidReferenceXlsxInspect = await SpreadsheetFile.inspectXlsx(invalidReferenceXlsx);
assert.equal(invalidReferenceXlsxInspect.ok, false);
assert.ok(invalidReferenceXlsxInspect.issues.some((issue) => issue.type === "relationshipReferenceIdNotFound" && issue.path === "xl/worksheets/sheet1.xml" && issue.referenceAttribute === "r:id"));
assert.ok(xlsxReferenceZip.file("xl/media/image1.png"));
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/media/image1.png", remove: true }]), /invalid OOXML package.*relationshipReferenceIdNotFound/);
const patchedXlsx = await SpreadsheetFile.patchXlsx(xlsx, { "customXml/review.json": { json: { status: "ok" } } });
assert.equal(patchedXlsx.type, xlsx.type);
assert.equal(patchedXlsx.metadata.patchedParts, 1);
assert.equal(patchedXlsx.metadata.validated, true);
assert.equal(patchedXlsx.metadata.validationIssues, 0);
assert.match((await SpreadsheetFile.inspectXlsx(patchedXlsx, { includeText: true, maxChars: 16000 })).ndjson, /review\.json/);
const recipeImageBytes = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=", "base64");
const recipeImageXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{
  path: "xl/media/review.png",
  bytes: recipeImageBytes,
  recipe: { kind: "image", source: "xl/worksheets/sheet1.xml", id: "rIdReviewImage" },
}]);
assert.equal(recipeImageXlsx.metadata.recipesApplied, 1);
const recipeImageInspect = await SpreadsheetFile.inspectXlsx(recipeImageXlsx);
assert.equal(recipeImageInspect.ok, true);
assert.ok(recipeImageInspect.parts.some((part) => part.path === "xl/media/review.png" && part.contentType === "image/png"));
const recipeImageZip = await JSZip.loadAsync(new Uint8Array(await recipeImageXlsx.arrayBuffer()));
assert.match(await recipeImageZip.file("xl/worksheets/_rels/sheet1.xml.rels").async("text"), /Id="rIdReviewImage"[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/image"[^>]*Target="\.\.\/media\/review\.png"/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/media/unsupported.png", bytes: recipeImageBytes, recipe: { kind: "image", source: "xl/worksheets/sheet1.xml", sourceReference: true } }]), /root element wsDr/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/collision.xml", xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>', recipe: { kind: "worksheet", source: "xl/workbook.xml", id: "rId1", sourceReference: { name: "Collision" } } }]), /relationship Id rId1.*already targets/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/duplicate-id.xml", xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>', recipe: { kind: "worksheet", source: "xl/workbook.xml", sourceReference: { name: "Duplicate Id", sheetId: 1 } } }]), /sheetId 1 already exists/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/duplicate-name.xml", xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>', recipe: { kind: "worksheet", source: "xl/workbook.xml", sourceReference: { name: "Sheet1" } } }]), /name Sheet1 already exists/);
const recipeTableXml = '<?xml version="1.0" encoding="UTF-8"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="99" name="ReviewTable" displayName="ReviewTable" ref="A1:B2" totalsRowShown="0"><autoFilter ref="A1:B2"/><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
const recipeTableXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/tables/review.xml", xml: recipeTableXml, recipe: { kind: "table", source: "xl/worksheets/sheet1.xml", id: "rIdReviewTable", sourceReference: true } }]);
assert.equal(recipeTableXlsx.metadata.sourceReferencesUpdated, 1);
const recipeTableZip = await JSZip.loadAsync(new Uint8Array(await recipeTableXlsx.arrayBuffer()));
assert.match(await recipeTableZip.file("xl/worksheets/sheet1.xml").async("text"), /<tableParts count="2">[\s\S]*<tablePart r:id="rIdReviewTable"\/>/);
const removedRecipeTableXlsx = await SpreadsheetFile.patchXlsx(recipeTableXlsx, [{ path: "xl/tables/review.xml", remove: true, recipe: { kind: "table", source: "xl/worksheets/sheet1.xml", id: "rIdReviewTable", sourceReference: true } }]);
assert.equal((await SpreadsheetFile.inspectXlsx(removedRecipeTableXlsx)).ok, true);
const removedRecipeTableZip = await JSZip.loadAsync(new Uint8Array(await removedRecipeTableXlsx.arrayBuffer()));
assert.match(await removedRecipeTableZip.file("xl/worksheets/sheet1.xml").async("text"), /<tableParts count="1">/);
assert.doesNotMatch(await removedRecipeTableZip.file("xl/worksheets/sheet1.xml").async("text"), /rIdReviewTable/);
const recipeWorksheetXml = '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>';
const recipeWorksheetXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/sheetReview.xml", xml: recipeWorksheetXml, recipe: { kind: "worksheet", source: "xl/workbook.xml", sourceReference: { name: "Review Data" } } }]);
assert.equal(recipeWorksheetXlsx.metadata.sourceReferencesUpdated, 1);
const recipeWorksheetZip = await JSZip.loadAsync(new Uint8Array(await recipeWorksheetXlsx.arrayBuffer()));
const recipeWorksheetRels = await recipeWorksheetZip.file("xl/_rels/workbook.xml.rels").async("text");
const recipeWorksheetRelId = /<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="worksheets\/sheetReview\.xml"/.exec(recipeWorksheetRels)?.[1];
assert.ok(recipeWorksheetRelId);
assert.match(await recipeWorksheetZip.file("xl/workbook.xml").async("text"), new RegExp(`<sheet name="Review Data" sheetId="2" r:id="${recipeWorksheetRelId}"/>`));
assert.ok((await SpreadsheetFile.importXlsx(recipeWorksheetXlsx)).worksheets.getItem("Review Data"));
const removedRecipeWorksheetXlsx = await SpreadsheetFile.patchXlsx(recipeWorksheetXlsx, [{ path: "xl/worksheets/sheetReview.xml", remove: true, recipe: { kind: "worksheet", source: "xl/workbook.xml", sourceReference: { name: "Review Data" } } }]);
assert.equal((await SpreadsheetFile.inspectXlsx(removedRecipeWorksheetXlsx)).ok, true);
assert.equal((await SpreadsheetFile.importXlsx(removedRecipeWorksheetXlsx)).worksheets.getItem("Review Data"), undefined);
assert.doesNotMatch(await (await JSZip.loadAsync(new Uint8Array(await removedRecipeWorksheetXlsx.arrayBuffer()))).file("xl/workbook.xml").async("text"), /Review Data/);
const replacementXlsx = await SpreadsheetFile.patchXlsx(xlsx, { "customXml/open-office-artifact.json": { json: { replaced: true } } }, { maxParts: xlsxInspect.parts.length });
assert.equal(replacementXlsx.metadata.patchedParts, 1);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "customXml/new-part.json", json: {} }], { maxParts: xlsxInspect.parts.length }), /would create .* maxParts/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "../evil.xml", text: "bad" }]), /Unsafe XLSX part path/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "customXml/unknown.xml", xml: "<x/>", recipe: "not-a-real-part" }]), /OOXML part recipe.*unsupported/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "customXml/large.txt", text: "12345" }], { maxPatchBytes: 4 }), /exceeds maxPatchBytes/);
await assert.rejects(() => SpreadsheetFile.inspectXlsx(xlsx, { maxParts: 1 }), /maxParts/);
await assert.rejects(() => SpreadsheetFile.inspectXlsx(xlsx, { maxPartBytes: 1 }), /maxPartBytes/);
const safelyRemovedStylesXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/styles.xml", remove: true }]);
assert.equal(safelyRemovedStylesXlsx.metadata.contentTypesUpdated, 1);
assert.equal(safelyRemovedStylesXlsx.metadata.relationshipsUpdated, 1);
assert.equal((await SpreadsheetFile.inspectXlsx(safelyRemovedStylesXlsx)).ok, true);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/styles.xml", remove: true }], { syncRelationships: false }), /invalid OOXML package.*relationshipTargetNotFound/);
const brokenRelationshipXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/styles.xml", remove: true }], { syncRelationships: false, validateResult: false });
assert.equal(brokenRelationshipXlsx.metadata.validated, false);
const brokenRelationshipInspect = await SpreadsheetFile.inspectXlsx(brokenRelationshipXlsx);
assert.equal(brokenRelationshipInspect.ok, false);
assert.ok(brokenRelationshipInspect.issues.some((issue) => issue.type === "relationshipTargetNotFound" && issue.target === "xl/styles.xml"));
const brokenRelationshipZip = await JSZip.loadAsync(new Uint8Array(await safelyRemovedStylesXlsx.arrayBuffer()));
assert.doesNotMatch(await brokenRelationshipZip.file("[Content_Types].xml").async("text"), /PartName="\/xl\/styles\.xml"/);
const xlsxBytes = new Uint8Array(await xlsx.arrayBuffer());
const zip = await JSZip.loadAsync(xlsxBytes);
const tablePartNames = Object.keys(zip.files).filter((name) => /^xl\/tables\/table\d+\.xml$/.test(name));
assert.equal(tablePartNames.length, 1);
const tableXml = await zip.file(tablePartNames[0]).async("text");
assert.match(tableXml, /displayName="TasksTable"/);
assert.match(tableXml, /ref="A1:D4"/);
assert.match(tableXml, /<tableColumns count="4">/);
assert.match(tableXml, /showFirstColumn="1"/);
assert.match(tableXml, /showLastColumn="1"/);
assert.match(tableXml, /showRowStripes="0"/);
const columnSortWorkbook = Workbook.create();
const columnSortSheet = columnSortWorkbook.worksheets.add("ColumnSort");
columnSortSheet.getRange("A1:B2").values = [["Plan", "Actual"], [1, 2]];
columnSortSheet.sortState = {
  reference: "A1:B2",
  caseSensitive: true,
  sortMethod: "pinYin",
  columnSort: true,
  conditions: [{ reference: "A2:B2", descending: true, customList: "Actual,Plan" }, { reference: "A1:B1", descending: false }],
};
const columnSortXlsx = await SpreadsheetFile.exportXlsx(columnSortWorkbook);
const columnSortZip = await JSZip.loadAsync(new Uint8Array(await columnSortXlsx.arrayBuffer()));
const columnSortXml = await columnSortZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(columnSortXml, /<sortState ref="A1:B2" caseSensitive="1" sortMethod="pinYin" columnSort="1">/);
assert.match(columnSortXml, /<sortCondition ref="A2:B2" descending="1" customList="Actual,Plan"\/>/);
assert.deepEqual((await SpreadsheetFile.importXlsx(columnSortXlsx)).worksheets.getItem("ColumnSort").sortState, columnSortSheet.sortState);
const formulaTableWorkbook = Workbook.create();
const formulaTableSheet = formulaTableWorkbook.worksheets.add("FormulaTable");
formulaTableSheet.getRange("A1:C4").values = [["Product", "Units", "Revenue"], ["North", 2, 4], ["South", 3, 6], ["Total", 2.5, 10]];
formulaTableSheet.getRange("C2:C3").formulas = [["=B2*2"], ["=B3*2"]];
formulaTableSheet.getRange("B4:C4").formulas = [["=AVERAGE(B2:B3)", "=SUBTOTAL(109,C2:C3)"]];
formulaTableSheet.tables.add({
  range: "A1:C4",
  name: "FormulaTable",
  showTotals: true,
  filters: [
    { columnIndex: 0, kind: "values", values: ["North"], includeBlank: true },
    { columnIndex: 1, kind: "custom", matchAll: true, criteria: [{ operator: "greaterThanOrEqual", value: "2" }, { operator: "lessThanOrEqual", value: "3" }] },
  ],
  sortState: {
    reference: "A2:C3",
    caseSensitive: true,
    sortMethod: "stroke",
    conditions: [{ reference: "C2:C3", descending: true, customList: "6,4" }, { reference: "A2:A3", descending: false }],
  },
  columnDefinitions: [
    { name: "Product", totalsRowFunction: "none", totalsRowLabel: "Total" },
    { name: "Units", totalsRowFunction: "average" },
    { name: "Revenue", calculatedColumnFormula: "=[@Units]*2", totalsRowFunction: "custom", totalsRowFormula: "=SUBTOTAL(109,[Revenue])" },
  ],
});
const formulaTableXlsx = await SpreadsheetFile.exportXlsx(formulaTableWorkbook);
const formulaTableZip = await JSZip.loadAsync(new Uint8Array(await formulaTableXlsx.arrayBuffer()));
const formulaTableXml = await formulaTableZip.file("xl/tables/table1.xml").async("text");
assert.match(formulaTableXml, /totalsRowLabel="Total"/);
assert.match(formulaTableXml, /totalsRowFunction="average"/);
assert.match(formulaTableXml, /<calculatedColumnFormula>\[@Units\]\*2<\/calculatedColumnFormula>/);
assert.match(formulaTableXml, /<totalsRowFormula>SUBTOTAL\(109,\[Revenue\]\)<\/totalsRowFormula>/);
assert.match(formulaTableXml, /<filterColumn colId="0"><filters blank="1"><filter val="North"\/><\/filters><\/filterColumn>/);
assert.match(formulaTableXml, /<customFilters and="1"><customFilter operator="greaterThanOrEqual" val="2"\/><customFilter operator="lessThanOrEqual" val="3"\/><\/customFilters>/);
assert.match(formulaTableXml, /<sortState ref="A2:C3" caseSensitive="1" sortMethod="stroke"><sortCondition ref="C2:C3" descending="1" customList="6,4"\/><sortCondition ref="A2:A3"\/><\/sortState>/);
const formulaTableImported = await SpreadsheetFile.importXlsx(formulaTableXlsx);
const importedFormulaTable = formulaTableImported.worksheets.getItem("FormulaTable").tables.getItemOrNullObject("FormulaTable");
assert.deepEqual(importedFormulaTable.columnDefinitions, [
  { name: "Product", calculatedColumnFormula: "", calculatedColumnFormulaArray: false, totalsRowFunction: "none", totalsRowLabel: "Total", totalsRowFormula: "", totalsRowFormulaArray: false },
  { name: "Units", calculatedColumnFormula: "", calculatedColumnFormulaArray: false, totalsRowFunction: "average", totalsRowLabel: "", totalsRowFormula: "", totalsRowFormulaArray: false },
  { name: "Revenue", calculatedColumnFormula: "=[@Units]*2", calculatedColumnFormulaArray: false, totalsRowFunction: "custom", totalsRowLabel: "", totalsRowFormula: "=SUBTOTAL(109,[Revenue])", totalsRowFormulaArray: false },
]);
assert.deepEqual(importedFormulaTable.filters, [
  { columnIndex: 0, kind: "values", values: ["North"], includeBlank: true },
  { columnIndex: 1, kind: "custom", matchAll: true, criteria: [{ operator: "greaterThanOrEqual", value: "2" }, { operator: "lessThanOrEqual", value: "3" }] },
]);
assert.deepEqual(importedFormulaTable.sortState, {
  reference: "A2:C3",
  caseSensitive: true,
  sortMethod: "stroke",
  conditions: [{ reference: "C2:C3", descending: true, customList: "6,4" }, { reference: "A2:A3", descending: false }],
});
importedFormulaTable.sortState.columnSort = false;
await assert.rejects(SpreadsheetFile.exportXlsx(formulaTableImported), /AutoFilter sortState cannot define columnSort/i);
const advancedFilterWorkbook = Workbook.create();
const advancedFilterSheet = advancedFilterWorkbook.worksheets.add("AdvancedFilters");
advancedFilterSheet.getRange("A1:C3").values = [["Date", "Status", "Score"], [45853, "ready", 95], [45854, "pending", 80]];
const javascriptAdvancedFilterTable = advancedFilterSheet.tables.add({ range: "A1:C3", name: "AdvancedFilterTable" });
javascriptAdvancedFilterTable.filters = [
  { columnIndex: 0, kind: "values", values: [], includeBlank: false, calendarType: "gregorian", dateGroups: [{ grouping: "day", year: 2026, month: 7, day: 15 }] },
  { columnIndex: 1, kind: "dynamic", type: "today", value: 45853, maxValue: 45854 },
  { columnIndex: 2, kind: "top10", top: true, percent: true, value: 10, filterValue: 95 },
];
const advancedFilterXlsx = await SpreadsheetFile.exportXlsx(advancedFilterWorkbook);
const advancedFilterZip = await JSZip.loadAsync(new Uint8Array(await advancedFilterXlsx.arrayBuffer()));
const advancedFilterXml = await advancedFilterZip.file("xl/tables/table1.xml").async("text");
assert.match(advancedFilterXml, /<filters calendarType="gregorian"><dateGroupItem year="2026" month="7" day="15" dateTimeGrouping="day"\/><\/filters>/);
assert.match(advancedFilterXml, /<dynamicFilter type="today" val="45853" maxVal="45854"\/>/);
assert.match(advancedFilterXml, /<top10 top="1" percent="1" val="10" filterVal="95"\/>/);
assert.deepEqual((await SpreadsheetFile.importXlsx(advancedFilterXlsx)).worksheets.getItem("AdvancedFilters").tables.items[0].filters, javascriptAdvancedFilterTable.filters);
javascriptAdvancedFilterTable.filters[0].values = ["2026-07-15"];
await assert.rejects(SpreadsheetFile.exportXlsx(advancedFilterWorkbook), /cannot mix exact values and grouped dates/i);
const iconRuleWorkbook = Workbook.create();
const iconRuleSheet = iconRuleWorkbook.worksheets.add("IconRules");
iconRuleSheet.getRange("A1:B3").values = [["Trend", "Rating"], [1, 5], [2, 3]];
const javascriptIconTable = iconRuleSheet.tables.add({ range: "A1:B3", name: "IconRuleTable" });
javascriptIconTable.filters = [
  { columnIndex: 0, kind: "icon", iconSet: "3Arrows", iconId: 0 },
  { columnIndex: 1, kind: "icon", iconSet: "3Flags" },
];
javascriptIconTable.sortState = {
  reference: "A2:B3",
  caseSensitive: false,
  conditions: [
    { reference: "B2:B3", descending: true, kind: "icon", iconSet: "5Rating", iconId: 4 },
    { reference: "A2:A3", descending: false, kind: "icon", iconSet: "3Symbols2" },
  ],
};
const iconRuleXlsx = await SpreadsheetFile.exportXlsx(iconRuleWorkbook);
const iconRuleZip = await JSZip.loadAsync(new Uint8Array(await iconRuleXlsx.arrayBuffer()));
const iconRuleXml = await iconRuleZip.file("xl/tables/table1.xml").async("text");
assert.match(iconRuleXml, /<iconFilter iconSet="3Arrows" iconId="0"\/>/);
assert.match(iconRuleXml, /<iconFilter iconSet="3Flags"\/>/);
assert.match(iconRuleXml, /<sortCondition ref="B2:B3" descending="1" sortBy="icon" iconSet="5Rating" iconId="4"\/>/);
assert.match(iconRuleXml, /<sortCondition ref="A2:A3" sortBy="icon" iconSet="3Symbols2"\/>/);
const javascriptIconImported = await SpreadsheetFile.importXlsx(iconRuleXlsx);
assert.deepEqual(javascriptIconImported.worksheets.getItem("IconRules").tables.items[0].filters, javascriptIconTable.filters);
assert.deepEqual(javascriptIconImported.worksheets.getItem("IconRules").tables.items[0].sortState, javascriptIconTable.sortState);
const colorRuleWorkbook = Workbook.create();
const colorRuleSheet = colorRuleWorkbook.worksheets.add("ColorRules");
colorRuleSheet.getRange("A1:B3").values = [["Fill", "Font"], [1, 5], [2, 3]];
const javascriptColorTable = colorRuleSheet.tables.add({ range: "A1:B3", name: "ColorRuleTable" });
javascriptColorTable.filters = [
  { columnIndex: 0, kind: "color", target: "cell", color: "#E11D48" },
  { columnIndex: 1, kind: "color", target: "font", color: { theme: 4, tint: -0.25 } },
];
javascriptColorTable.sortState = {
  reference: "A2:B3",
  caseSensitive: false,
  conditions: [
    { reference: "B2:B3", descending: true, kind: "color", target: "font", color: { theme: 4, tint: -0.25 } },
    { reference: "A2:A3", descending: false, kind: "color", target: "cell", color: "#E11D48" },
  ],
};
const colorRuleXlsx = await SpreadsheetFile.exportXlsx(colorRuleWorkbook);
const colorRuleZip = await JSZip.loadAsync(new Uint8Array(await colorRuleXlsx.arrayBuffer()));
const colorRuleXml = await colorRuleZip.file("xl/tables/table1.xml").async("text");
const colorStylesXml = await colorRuleZip.file("xl/styles.xml").async("text");
assert.match(colorRuleXml, /<colorFilter dxfId="0" cellColor="1"\/>/);
assert.match(colorRuleXml, /<colorFilter dxfId="1" cellColor="0"\/>/);
assert.match(colorRuleXml, /<sortCondition ref="B2:B3" descending="1" sortBy="fontColor" dxfId="1"\/>/);
assert.match(colorRuleXml, /<sortCondition ref="A2:A3" sortBy="cellColor" dxfId="0"\/>/);
assert.ok(colorStylesXml.includes('<dxfs count="2"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFE11D48"/><bgColor indexed="64"/></patternFill></fill></dxf><dxf><font><color theme="4" tint="-0.25"/></font></dxf></dxfs>'));
const javascriptColorImported = await SpreadsheetFile.importXlsx(colorRuleXlsx);
assert.deepEqual(javascriptColorImported.worksheets.getItem("ColorRules").tables.items[0].filters, javascriptColorTable.filters);
assert.deepEqual(javascriptColorImported.worksheets.getItem("ColorRules").tables.items[0].sortState, javascriptColorTable.sortState);
const invalidColorRuleWorkbook = Workbook.create();
const invalidColorRuleSheet = invalidColorRuleWorkbook.worksheets.add("InvalidColor");
invalidColorRuleSheet.getRange("A1:A2").values = [["Value"], [1]];
invalidColorRuleSheet.tables.add({ range: "A1:A2", name: "InvalidColorTable", filters: [{ columnIndex: 0, kind: "color", target: "background", color: "#E11D48" }] });
await assert.rejects(SpreadsheetFile.exportXlsx(invalidColorRuleWorkbook), /color target must be 'cell' or 'font'/i);
const pivotPartNames = Object.keys(zip.files).filter((name) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(name));
assert.equal(pivotPartNames.length, 3);
const pivotTableXml = await zip.file(pivotPartNames[0]).async("text");
assert.match(pivotTableXml, /<pivotTableDefinition/);
assert.match(pivotTableXml, /name="RevenuePivot"/);
assert.match(pivotTableXml, /cacheId="1"/);
assert.match(pivotTableXml, /<rowFields count="1"><field x="0"\/><\/rowFields>/);
assert.match(pivotTableXml, /<dataField name="Revenue sum" fld="1" subtotal="sum"\/>/);
const regionalPivotTableXml = await zip.file(pivotPartNames[1]).async("text");
assert.match(regionalPivotTableXml, /name="RegionalPivot"/);
assert.match(regionalPivotTableXml, /<colFields count="1"><field x="1"\/><\/colFields>/);
assert.match(regionalPivotTableXml, /<pivotField axis="axisCol" multipleItemSelectionAllowed="1" showAll="0"><items count="3"><item x="0"\/><item x="1" h="1"\/><item t="default"\/><\/items><\/pivotField>/);
assert.match(regionalPivotTableXml, /<dataField name="Profit total" fld="6" subtotal="sum"\/>/);
assert.match(regionalPivotTableXml, /<filters count="1"><filter fld="5" type="dateBetween" id="1" stringValue1="2026-02-01" stringValue2="2026-03-31">/);
assert.match(regionalPivotTableXml, /<customFilters and="1"><customFilter operator="greaterThanOrEqual" val="2026-02-01"\/><customFilter operator="lessThanOrEqual" val="2026-03-31"\/><\/customFilters>/);
const calendarPivotTableXml = await zip.file(pivotPartNames[2]).async("text");
assert.match(calendarPivotTableXml, /name="CalendarPivot"/);
assert.match(calendarPivotTableXml, /<rowFields count="2"><field x="6"\/><field x="7"\/><\/rowFields>/);
assert.match(calendarPivotTableXml, /<colFields count="1"><field x="8"\/><\/colFields>/);
assert.match(calendarPivotTableXml, /<pivotField axis="axisCol" multipleItemSelectionAllowed="1" showAll="0"><items count="5"><item x="0" h="1"\/><item x="1"\/><item x="2"\/><item x="3"\/><item t="default"\/><\/items><\/pivotField>/);
const pivotCacheXml = await zip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text");
assert.match(pivotCacheXml, /<pivotCacheDefinition/);
assert.match(pivotCacheXml, /r:id="rId1"/);
assert.match(pivotCacheXml, /<worksheetSource ref="F1:G4" sheet="Sheet1"\/>/);
assert.match(pivotCacheXml, /<cacheField name="Month"/);
assert.match(pivotCacheXml, /<cacheField name="Revenue"/);
const pivotCacheRecordsXml = await zip.file("xl/pivotCache/pivotCacheRecords1.xml").async("text");
assert.match(pivotCacheRecordsXml, /<pivotCacheRecords[^>]*count="3"/);
assert.match(pivotCacheRecordsXml, /<s v="Jan"\/><n v="100"\/>/);
assert.match(pivotCacheRecordsXml, /<s v="Mar"\/><n v="130"\/>/);
const pivotCacheDefinitionRelsXml = await zip.file("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels").async("text");
assert.match(pivotCacheDefinitionRelsXml, /relationships\/pivotCacheRecords/);
assert.match(pivotCacheDefinitionRelsXml, /Target="pivotCacheRecords1\.xml"/);
const regionalPivotCacheXml = await zip.file("xl/pivotCache/pivotCacheDefinition2.xml").async("text");
assert.match(regionalPivotCacheXml, /refreshOnLoad="0" saveData="1" enableRefresh="0" invalid="1" missingItemsLimit="3" refreshedBy="QA Agent" refreshedDateIso="2026-07-12T00:00:00Z"/);
assert.match(regionalPivotCacheXml, /<cacheFields count="7">/);
assert.match(regionalPivotCacheXml, /<cacheField name="Profit" formula="\('Revenue'-'Cost'\)\*100%" databaseField="0" numFmtId="0"><sharedItems containsNumber="1" count="0"\/><\/cacheField>/);
const calendarPivotCacheXml = await zip.file("xl/pivotCache/pivotCacheDefinition3.xml").async("text");
assert.match(calendarPivotCacheXml, /<cacheFields count="9">/);
assert.match(calendarPivotCacheXml, /<cacheField name="OrderDate" numFmtId="0"><sharedItems containsDate="1" count="6"><d v="2026-01-15T00:00:00"\/>/);
assert.match(calendarPivotCacheXml, /<cacheField name="Order Year" databaseField="0" numFmtId="0"><fieldGroup base="5"><rangePr groupBy="years" autoStart="1" autoEnd="1"\/><groupItems count="1"><s v="2026"\/><\/groupItems><\/fieldGroup><\/cacheField>/);
assert.match(calendarPivotCacheXml, /<cacheField name="Order Quarter" databaseField="0" numFmtId="0"><fieldGroup base="5" par="6"><rangePr groupBy="quarters" autoStart="1" autoEnd="1"\/><groupItems count="2"><s v="Q1"\/><s v="Q2"\/><\/groupItems><\/fieldGroup><\/cacheField>/);
assert.match(calendarPivotCacheXml, /<cacheField name="Order Month" databaseField="0" numFmtId="0"><fieldGroup base="5" par="7"><rangePr groupBy="months" autoStart="1" autoEnd="1"\/><groupItems count="4"><s v="Jan"\/><s v="Feb"\/><s v="Mar"\/><s v="Apr"\/><\/groupItems><\/fieldGroup><\/cacheField>/);
assert.match(await zip.file("xl/pivotCache/pivotCacheRecords3.xml").async("text"), /<d v="2026-01-15T00:00:00"\/>/);
const pivotTableDefinitionRelsXml = await zip.file("xl/pivotTables/_rels/pivotTable1.xml.rels").async("text");
assert.match(pivotTableDefinitionRelsXml, /relationships\/pivotCacheDefinition/);
assert.match(pivotTableDefinitionRelsXml, /Target="\.\.\/pivotCache\/pivotCacheDefinition1\.xml"/);
const worksheetXml = await zip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(worksheetXml, /<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"\/><\/sheetView><\/sheetViews><sheetData>/);
assert.match(worksheetXml, /<conditionalFormatting sqref="C2:C3">/);
assert.match(worksheetXml, /<cfRule type="cellIs"[^>]*dxfId="0"/);
assert.match(worksheetXml, /<conditionalFormatting sqref="G2:G4"><cfRule type="colorScale"/);
assert.match(worksheetXml, /<color rgb="FF22C55E"\/>/);
assert.match(worksheetXml, /<cfvo type="percentile" val="50"\/>/);
assert.match(worksheetXml, /<dataValidations count="1">/);
assert.match(worksheetXml, /<dataValidation type="list"[^>]*sqref="D2:D3"/);
assert.match(worksheetXml, /<c r="A1" t="s" s="\d+"><v>\d+<\/v><\/c>/);
assert.match(worksheetXml, /<c r="B2" s="\d+"><v>3<\/v><\/c>/);
const sharedStringsXml = await zip.file("xl/sharedStrings.xml").async("text");
assert.match(sharedStringsXml, /<sst/);
assert.match(sharedStringsXml, /<t>Month<\/t>/);
assert.match(sharedStringsXml, /<t>Not Started<\/t>/);
const workbookXml = await zip.file("xl/workbook.xml").async("text");
assert.doesNotMatch(workbookXml, /<workbookPr\b[^>]*date1904=/);
assert.match(workbookXml, /<definedNames>/);
assert.match(workbookXml, /name="RevenueData"/);
assert.match(workbookXml, /name="RevenueData" comment="Revenue data body" hidden="0"/);
assert.match(workbookXml, /name="LocalRevenueData" localSheetId="0" hidden="1"/);
assert.match(workbookXml, /Sheet1!G2:G4/);
assert.match(workbookXml, /<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><pivotCache cacheId="2" r:id="rId\d+"\/><pivotCache cacheId="3" r:id="rId\d+"\/><\/pivotCaches>/);
const stylesXml = await zip.file("xl/styles.xml").async("text");
assert.match(stylesXml, /<dxfs count="2">/);
assert.match(stylesXml, /<fgColor rgb="FF22C55E"/);
assert.match(stylesXml, /<fgColor rgb="FFE0F2FE"/);
assert.match(stylesXml, /<numFmt numFmtId="164" formatCode="#,##0"/);
assert.match(stylesXml, /<fgColor rgb="FF0F172A"/);
assert.match(stylesXml, /<fgColor rgb="FFE0F2FE"/);
assert.match(stylesXml, /<alignment horizontal="center" vertical="center" wrapText="1"\/>/);
assert.match(stylesXml, /<left style="thin"><color rgb="FF334155"\/><\/left>/);
assert.match(stylesXml, /applyAlignment="1"/);
assert.match(stylesXml, /applyBorder="1"/);
assert.match(stylesXml, /<b\/>/);
assert.match(worksheetXml, /<tableParts count="1">/);
assert.match(worksheetXml, /<tablePart r:id="rId1"\/>/);
assert.match(worksheetXml, /<x14:sparklineGroups/);
assert.match(worksheetXml, /<x14:sparklineGroup type="line"/);
assert.match(worksheetXml, /<xm:f>Sheet1!G2:G4<\/xm:f>/);
assert.match(worksheetXml, /<xm:sqref>H2:H2<\/xm:sqref>/);
assert.match(worksheetXml, /<x14:sparklineGroup type="column"/);
assert.match(worksheetXml, /<xm:sqref>H3<\/xm:sqref>/);
const thirdPartyViewXml = worksheetXml
  .replace('showGridLines="0"', 'showGridLines="1"')
  .replace('<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>', '<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozenSplit"/>');
const thirdPartyViewXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/sheet1.xml", xml: thirdPartyViewXml }]);
const thirdPartyViewSheet = (await SpreadsheetFile.importXlsx(thirdPartyViewXlsx)).worksheets.getItem("Sheet1");
assert.equal(thirdPartyViewSheet.showGridLines, true);
assert.deepEqual(thirdPartyViewSheet.freezePanes.toJSON(), { rows: 2, columns: 1, frozen: true, topLeftCell: "B3", activePane: "bottomRight" });
const splitOnlyViewXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/sheet1.xml", xml: thirdPartyViewXml.replace('state="frozenSplit"', 'state="split"') }]);
assert.equal((await SpreadsheetFile.importXlsx(splitOnlyViewXlsx)).worksheets.getItem("Sheet1").freezePanes.frozen, false);
const worksheetRelsXml = await zip.file("xl/worksheets/_rels/sheet1.xml.rels").async("text");
assert.match(worksheetRelsXml, /Target="\.\.\/tables\/table1\.xml"/);
assert.match(worksheetRelsXml, /Target="\.\.\/drawings\/drawing1\.xml"/);
assert.match(worksheetRelsXml, /Target="\.\.\/threadedComments\/threadedComment1\.xml"/);
assert.match(worksheetRelsXml, /Target="\.\.\/pivotTables\/pivotTable1\.xml"/);
assert.match(worksheetRelsXml, /relationships\/pivotTable/);
assert.match(worksheetRelsXml, /relationships\/threadedComment/);
const drawingXml = await zip.file("xl/drawings/drawing1.xml").async("text");
assert.match(drawingXml, /<xdr:oneCellAnchor>/);
assert.match(drawingXml, /name="LogoImage"/);
assert.match(drawingXml, /descr="Logo placeholder"/);
const drawingRelsXml = await zip.file("xl/drawings/_rels/drawing1.xml.rels").async("text");
assert.match(drawingRelsXml, /Target="\.\.\/media\/image1\.png"/);
assert.match(drawingRelsXml, /Target="\.\.\/charts\/chart1\.xml"/);
const drawingChartXml = await zip.file("xl/drawings/drawing1.xml").async("text");
assert.match(drawingChartXml, /<xdr:graphicFrame>/);
assert.match(drawingChartXml, /<c:chart r:id="rId2"\/>/);
const chartXml = await zip.file("xl/charts/chart1.xml").async("text");
assert.match(chartXml, /<c:chartSpace/);
assert.match(chartXml, /Revenue Trend/);
assert.match(chartXml, /<c:lineChart>/);
assert.match(chartXml, /<c:v>130<\/c:v>/);
const threadedCommentsXml = await zip.file("xl/threadedComments/threadedComment1.xml").async("text");
assert.match(threadedCommentsXml, /<ThreadedComments/);
assert.match(threadedCommentsXml, /ref="C2"/);
assert.match(threadedCommentsXml, /Formula checks revenue sum/);
assert.match(threadedCommentsXml, /Reviewed by model/);
assert.match(threadedCommentsXml, /done="1"/);
assert.match(threadedCommentsXml, new RegExp(`id="${rootCommentId.replace(/[{}]/g, "\\$&")}"`));
assert.match(threadedCommentsXml, new RegExp(`id="${replyCommentId.replace(/[{}]/g, "\\$&")}" parentId="${rootCommentId.replace(/[{}]/g, "\\$&")}"`));
assert.match(threadedCommentsXml, /dT="2026-07-12T09:00:00\.000Z"/);
const personsXml = await zip.file("xl/persons/person.xml").async("text");
assert.match(personsXml, /<personList/);
assert.match(personsXml, /displayName="Analyst"/);
assert.match(personsXml, /displayName="Reviewer"/);
assert.match(personsXml, /userId="analyst@example\.com"/);
const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels").async("text");
assert.match(workbookRelsXml, /Target="persons\/person.xml"/);
assert.match(workbookRelsXml, /relationships\/person/);
assert.match(workbookRelsXml, /Target="pivotCache\/pivotCacheDefinition1\.xml"/);
assert.match(workbookRelsXml, /relationships\/pivotCacheDefinition/);
const mediaBytes = await zip.file("xl/media/image1.png").async("uint8array");
assert.ok(mediaBytes.byteLength > 10);
const contentTypesXml = await zip.file("[Content_Types].xml").async("text");
assert.match(contentTypesXml, /Default Extension="png" ContentType="image\/png"/);
assert.match(contentTypesXml, /PartName="\/xl\/drawings\/drawing1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.drawing\+xml"/);
assert.match(contentTypesXml, /table1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.table\+xml"/);
assert.match(contentTypesXml, /chart1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.drawingml\.chart\+xml"/);
assert.match(contentTypesXml, /pivotTable1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.pivotTable\+xml"/);
assert.match(contentTypesXml, /pivotCacheDefinition1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.pivotCacheDefinition\+xml"/);
assert.match(contentTypesXml, /pivotCacheRecords1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.pivotCacheRecords\+xml"/);
assert.match(contentTypesXml, /threadedComment1\.xml" ContentType="application\/vnd\.ms-excel\.threadedcomments\+xml"/);
assert.match(contentTypesXml, /person\.xml" ContentType="application\/vnd\.ms-excel\.person\+xml"/);
assert.match(contentTypesXml, /sharedStrings\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sharedStrings\+xml"/);
const nativeOnlyZip = await JSZip.loadAsync(xlsxBytes);
nativeOnlyZip.remove("customXml/open-office-artifact.json");
const nativeOnlyWorkbook = await SpreadsheetFile.importXlsx(new FileBlob(await nativeOnlyZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
const nativeOnlySheet = nativeOnlyWorkbook.worksheets.getItem("Sheet1");
assert.equal(nativeOnlySheet.images.items.length, 1);
assert.equal(nativeOnlySheet.charts.items.length, 2);
assert.equal(nativeOnlySheet.pivotTables.items.length, 3);
const nativeOnlyPivot = nativeOnlySheet.pivotTables.getItemOrNullObject("RevenuePivot");
assert.deepEqual(nativeOnlyPivot.computedValues(), [["Month", "Revenue sum"], ["Jan", 100], ["Feb", 120], ["Mar", 130]]);
assert.equal(nativeOnlyWorkbook.resolve("RevenuePivot"), nativeOnlyPivot);
assert.match(nativeOnlyWorkbook.inspect({ kind: "pivotTable", target: "RevenuePivot", maxChars: 4000 }).ndjson, /Revenue sum/);
const nativeOnlyRegionalPivot = nativeOnlySheet.pivotTables.getItemOrNullObject("RegionalPivot");
assert.deepEqual(nativeOnlyRegionalPivot.computedValues(), [["Region", "OrderDate", "Q1 — Revenue total", "Q1 — Profit total"], ["West", "2026-02-20", 30, 10], ["East", "2026-03-05", 5, 2]]);
assert.deepEqual(nativeOnlyRegionalPivot.columnFields, ["Quarter"]);
assert.deepEqual(nativeOnlyRegionalPivot.calculatedFields, [{ name: "Profit", formula: "=('Revenue'-'Cost')*100%", numFmtId: 0, references: ["Revenue", "Cost"] }]);
assert.deepEqual(nativeOnlyRegionalPivot.filters, [{ field: "Quarter", exclude: ["Q2"] }, { field: "OrderDate", type: "dateBetween", value1: "2026-02-01", value2: "2026-03-31", useWholeDay: true }]);
assert.deepEqual(nativeOnlyRegionalPivot.refreshPolicy, { refreshOnLoad: false, saveData: true, enableRefresh: false, invalid: true, missingItemsLimit: 3, refreshedBy: "QA Agent", refreshedDateIso: "2026-07-12T00:00:00Z" });
const nativeOnlyRegionalRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(nativeOnlyWorkbook));
assert.deepEqual(nativeOnlyRegionalRoundtrip.resolve("RegionalPivot").computedValues(), [["Region", "OrderDate", "Q1 — Revenue total", "Q1 — Profit total"], ["West", "2026-02-20", 30, 10], ["East", "2026-03-05", 5, 2]]);
assert.deepEqual(nativeOnlyRegionalRoundtrip.resolve("RegionalPivot").filters, [{ field: "Quarter", exclude: ["Q2"] }, { field: "OrderDate", type: "dateBetween", value1: "2026-02-01", value2: "2026-03-31", useWholeDay: true }]);
const nativeOnlyCalendarPivot = nativeOnlySheet.pivotTables.getItemOrNullObject("CalendarPivot");
assert.deepEqual(nativeOnlyCalendarPivot.groupFields.map(({ name, sourceField, groupBy, parent }) => ({ name, sourceField, groupBy, ...(parent ? { parent } : {}) })), calendarPivot.groupFields);
assert.deepEqual(nativeOnlyCalendarPivot.computedValues(), calendarPivot.computedValues());
assert.deepEqual(nativeOnlyCalendarPivot.filters, [{ field: "Order Month", exclude: ["Jan"] }]);
const nativeOnlyCalendarRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(nativeOnlyWorkbook));
assert.deepEqual(nativeOnlyCalendarRoundtrip.resolve("CalendarPivot").groupFields.map(({ name, sourceField, groupBy, parent }) => ({ name, sourceField, groupBy, ...(parent ? { parent } : {}) })), calendarPivot.groupFields);
assert.deepEqual(nativeOnlyCalendarRoundtrip.resolve("CalendarPivot").computedValues(), calendarPivot.computedValues());
const unsupportedCalculatedZip = await JSZip.loadAsync(xlsxBytes);
unsupportedCalculatedZip.remove("customXml/open-office-artifact.json");
unsupportedCalculatedZip.file("xl/pivotCache/pivotCacheDefinition2.xml", regionalPivotCacheXml.replace(/formula="[^"]+"/, "formula=\"LOG('Revenue')\""));
const unsupportedCalculatedWorkbook = await SpreadsheetFile.importXlsx(new FileBlob(await unsupportedCalculatedZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
const unsupportedCalculatedPivot = unsupportedCalculatedWorkbook.resolve("RegionalPivot");
assert.equal(unsupportedCalculatedPivot.calculatedFields[0].supported, false);
assert.equal(unsupportedCalculatedPivot.computedValues()[1][3], "#NAME?");
assert.match(unsupportedCalculatedWorkbook.verify().ndjson, /pivotCalculatedFieldUnsupported/);
const unsupportedCalculatedRoundtripZip = await JSZip.loadAsync(new Uint8Array(await (await SpreadsheetFile.exportXlsx(unsupportedCalculatedWorkbook)).arrayBuffer()));
assert.match(await unsupportedCalculatedRoundtripZip.file("xl/pivotCache/pivotCacheDefinition2.xml").async("text"), /formula="LOG\('Revenue'\)"/);
const unsupportedGroupZip = await JSZip.loadAsync(xlsxBytes);
unsupportedGroupZip.remove("customXml/open-office-artifact.json");
unsupportedGroupZip.file("xl/pivotCache/pivotCacheDefinition3.xml", calendarPivotCacheXml.replace('groupBy="months"', 'groupBy="weeks"'));
const unsupportedGroupWorkbook = await SpreadsheetFile.importXlsx(new FileBlob(await unsupportedGroupZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
const unsupportedGroupPivot = unsupportedGroupWorkbook.resolve("CalendarPivot");
assert.equal(unsupportedGroupPivot.groupFields.find((field) => field.name === "Order Month").supported, false);
assert.match(unsupportedGroupPivot.computedValues()[0].join(","), /#NAME\?/);
assert.match(unsupportedGroupWorkbook.verify().ndjson, /pivotGroupFieldUnsupported/);
const unsupportedGroupRoundtripZip = await JSZip.loadAsync(new Uint8Array(await (await SpreadsheetFile.exportXlsx(unsupportedGroupWorkbook)).arrayBuffer()));
assert.match(await unsupportedGroupRoundtripZip.file("xl/pivotCache/pivotCacheDefinition3.xml").async("text"), /groupBy="weeks"/);
const nativeOnlyImage = nativeOnlySheet.images.getItemOrNullObject("LogoImage");
assert.equal(nativeOnlyImage.alt, "Logo placeholder");
assert.match(nativeOnlyImage.dataUrl, /^data:image\/png;base64,/);
const nativeOnlyRevenueChart = nativeOnlySheet.charts.items.find((chart) => chart.title === "Revenue Trend");
assert.equal(nativeOnlyRevenueChart.type, "line");
assert.deepEqual(nativeOnlyRevenueChart.categories, ["Jan", "Feb", "Mar"]);
assert.deepEqual(nativeOnlyRevenueChart.series.items[0].values, [100, 120, 130]);
assert.ok(nativeOnlyRevenueChart.position.width > 100 && nativeOnlyRevenueChart.position.height > 80);
assert.equal(nativeOnlyWorkbook.resolve(nativeOnlyImage.id), nativeOnlyImage);
assert.match(nativeOnlyWorkbook.inspect({ kind: "drawing", target: nativeOnlyImage.id, maxChars: 4000 }).ndjson, /LogoImage/);
assert.equal(nativeOnlyWorkbook.worksheets.getItem("Sheet1").showGridLines, false);
assert.deepEqual(nativeOnlyWorkbook.worksheets.getItem("Sheet1").freezePanes.toJSON(), { rows: 1, columns: 2, frozen: true, topLeftCell: "C2", activePane: "bottomRight" });
const nativeSparklineInspect = nativeOnlyWorkbook.inspect({ kind: "sparkline", maxChars: 12000 }).ndjson;
assert.match(nativeSparklineInspect, /"kind":"sparkline"/);
assert.match(nativeSparklineInspect, /"type":"line"/);
assert.match(nativeSparklineInspect, /"targetRange":"H2:H2"/);
assert.match(nativeSparklineInspect, /"sourceData":"G2:G4"/);
assert.match(nativeSparklineInspect, /"type":"column"/);
const nativeThreadInspect = nativeOnlyWorkbook.inspect({ kind: "thread", maxChars: 12000 }).ndjson;
assert.match(nativeThreadInspect, /"kind":"thread"/);
assert.match(nativeThreadInspect, /Formula checks revenue sum/);
assert.match(nativeThreadInspect, /Reviewed by model/);
assert.match(nativeThreadInspect, /"resolved":true/);
assert.match(nativeThreadInspect, /11111111-1111-4111-8111-111111111111/);
assert.match(nativeThreadInspect, /BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB/);
const nativeThread = nativeOnlyWorkbook.comments.threads.find((item) => item.target.address === "C2");
assert.equal(nativeThread?.comments[0].id, rootCommentId);
assert.equal(nativeThread?.comments[0].date, "2026-07-12T09:00:00.000Z");
assert.equal(nativeThread?.comments[0].person.userId, "analyst@example.com");
assert.equal(nativeThread?.comments[1].id, replyCommentId);
assert.equal(nativeThread?.comments[1].parentId, rootCommentId);
assert.equal(nativeThread?.comments[1].author, "Reviewer");
const nativeThreadRoundtripZip = await JSZip.loadAsync(new Uint8Array(await (await SpreadsheetFile.exportXlsx(nativeOnlyWorkbook)).arrayBuffer()));
assert.match(await nativeThreadRoundtripZip.file("xl/threadedComments/threadedComment1.xml").async("text"), /id="\{22222222-2222-4222-8222-222222222222\}" parentId="\{11111111-1111-4111-8111-111111111111\}"/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/threadedComments/threadedComment1.xml", xml: threadedCommentsXml.replace(replyCommentId, "{NOT-A-GUID}") }]), /xlsxThreadedCommentIdInvalid/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/threadedComments/threadedComment1.xml", xml: threadedCommentsXml.replace(`parentId="${rootCommentId}"`, 'parentId="{DEADBEEF-DEAD-4EAD-8EAD-DEADBEEFDEAD}"') }]), /xlsxThreadedCommentParentNotFound/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/persons/person.xml", xml: personsXml.replace(analystPersonId, "not-a-guid") }]), /xlsxPersonIdInvalid/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/threadedComments/threadedComment1.xml", xml: threadedCommentsXml.replace(`ref="C2" dT="2026-07-12T09:05:00.000Z"`, `ref="D2" dT="2026-07-12T09:05:00.000Z"`) }]), /xlsxThreadedCommentParentRefMismatch/);
const secondRootId = "{33333333-3333-4333-8333-333333333333}";
const alternatePeopleXml = personsXml
  .replace(`<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">`, `<people:personList xmlns:people="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">`)
  .replaceAll("<person ", "<people:person ")
  .replace("</personList>", "</people:personList>");
const alternateThreadsXml = threadedCommentsXml
  .replace(`<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">`, `<tc:ThreadedComments xmlns:tc="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">`)
  .replaceAll("<threadedComment ", "<tc:threadedComment ")
  .replaceAll("</threadedComment>", "</tc:threadedComment>")
  .replaceAll("<text>", "<tc:text>")
  .replaceAll("</text>", "</tc:text>")
  .replace("</ThreadedComments>", `<tc:threadedComment ref="C2" dT="2026-07-12T09:10:00.000Z" personId="${analystPersonId}" id="${secondRootId}" done="0"><tc:text>Independent review on the same cell.</tc:text></tc:threadedComment></tc:ThreadedComments>`);
const relocatedThreadZip = await JSZip.loadAsync(xlsxBytes);
relocatedThreadZip.remove("customXml/open-office-artifact.json");
const metadataFreeThreadXlsx = new FileBlob(await relocatedThreadZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type });
const relocatedThreadXlsx = await SpreadsheetFile.patchXlsx(metadataFreeThreadXlsx, [
  { path: "xl/persons/person.xml", remove: true },
  { path: "xl/threadedComments/threadedComment1.xml", remove: true },
  { path: "xl/collaboration/people.xml", xml: alternatePeopleXml, recipe: { kind: "person", source: "xl/workbook.xml", id: "rIdRelocatedPeople" } },
  { path: "xl/collaboration/sheet1-threads.xml", xml: alternateThreadsXml, recipe: { kind: "threadedComments", source: "xl/worksheets/sheet1.xml", id: "rIdRelocatedThreads" } },
]);
const relocatedThreadInspect = await SpreadsheetFile.inspectXlsx(relocatedThreadXlsx);
assert.equal(relocatedThreadInspect.ok, true, JSON.stringify(relocatedThreadInspect.issues));
assert.ok(relocatedThreadInspect.parts.some((part) => part.path === "xl/collaboration/people.xml" && part.contentType === "application/vnd.ms-excel.person+xml"));
assert.ok(relocatedThreadInspect.parts.some((part) => part.path === "xl/collaboration/sheet1-threads.xml" && part.contentType === "application/vnd.ms-excel.threadedcomments+xml"));
const relocatedThreadWorkbook = await SpreadsheetFile.importXlsx(relocatedThreadXlsx);
const sameCellThreads = relocatedThreadWorkbook.comments.threads.filter((item) => item.target.address === "C2");
assert.equal(sameCellThreads.length, 2);
assert.equal(sameCellThreads.find((item) => item.comments[0].id === rootCommentId)?.comments[1].parentId, rootCommentId);
assert.equal(sameCellThreads.find((item) => item.comments[0].id === secondRootId)?.comments[0].text, "Independent review on the same cell.");
const nativeStyleInspect = nativeOnlyWorkbook.inspect({ kind: "style", range: "A1:C3", maxChars: 12000 }).ndjson;
assert.match(nativeStyleInspect, /"kind":"style"/);
assert.match(nativeStyleInspect, /"numberFormat":"#,##0"/);
assert.match(nativeStyleInspect, /"alignment":\{"horizontal":"center"/);
assert.match(nativeStyleInspect, /"border":\{"style":"thin"/);
const relocatedDrawingZip = await JSZip.loadAsync(xlsxBytes);
relocatedDrawingZip.remove("customXml/open-office-artifact.json");
const relocatedDrawingXml = (await relocatedDrawingZip.file("xl/drawings/drawing1.xml").async("text")).replace(
  /<xdr:oneCellAnchor>((?:(?!<\/xdr:oneCellAnchor>)[\s\S])*?name="Chart 1"(?:(?!<\/xdr:oneCellAnchor>)[\s\S])*?)<\/xdr:oneCellAnchor>/,
  (_anchor, body) => `<xdr:twoCellAnchor editAs="twoCell">${body.replace(/<xdr:ext[^>]*\/>/, '<xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>')}</xdr:twoCellAnchor>`,
).replace(
  /<xdr:oneCellAnchor>((?:(?!<\/xdr:oneCellAnchor>)[\s\S])*?name="ScoresChart"(?:(?!<\/xdr:oneCellAnchor>)[\s\S])*?)<\/xdr:oneCellAnchor>/,
  (_anchor, body) => `<xdr:absoluteAnchor>${body.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, '<xdr:pos x="1905000" y="2857500"/>')}</xdr:absoluteAnchor>`,
);
assert.match(relocatedDrawingXml, /<xdr:twoCellAnchor editAs="twoCell">/);
assert.match(relocatedDrawingXml, /<xdr:absoluteAnchor>/);
relocatedDrawingZip.file("xl/custom/visuals/workbook-drawing.xml", relocatedDrawingXml);
relocatedDrawingZip.remove("xl/drawings/drawing1.xml");
const relocatedDrawingRelationships = (await relocatedDrawingZip.file("xl/drawings/_rels/drawing1.xml.rels").async("text"))
  .replace("../media/image1.png", "../../assets/logo.png")
  .replace("../charts/chart1.xml", "chart-revenue.xml")
  .replace("../charts/chart2.xml", "chart-scores.xml");
relocatedDrawingZip.file("xl/custom/visuals/_rels/workbook-drawing.xml.rels", relocatedDrawingRelationships);
relocatedDrawingZip.remove("xl/drawings/_rels/drawing1.xml.rels");
relocatedDrawingZip.file("xl/assets/logo.png", await relocatedDrawingZip.file("xl/media/image1.png").async("uint8array"));
relocatedDrawingZip.remove("xl/media/image1.png");
const relocatedRevenueChartXml = (await relocatedDrawingZip.file("xl/charts/chart1.xml").async("text"))
  .replace(/<c:tx><c:v>([\s\S]*?)<\/c:v><\/c:tx>/, '<c:tx><c:strRef><c:f>Sheet1!$G$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>$1</c:v></c:pt></c:strCache></c:strRef></c:tx>')
  .replace(/<c:cat><c:strLit>([\s\S]*?)<\/c:strLit><\/c:cat>/, '<c:cat><c:strRef><c:f>Sheet1!$F$2:$F$4</c:f><c:strCache>$1</c:strCache></c:strRef></c:cat>')
  .replace(/<c:val><c:numLit>([\s\S]*?)<\/c:numLit><\/c:val>/, '<c:val><c:numRef><c:f>Sheet1!$G$2:$G$4</c:f><c:numCache>$1</c:numCache></c:numRef></c:val>');
relocatedDrawingZip.file("xl/custom/visuals/chart-revenue.xml", relocatedRevenueChartXml);
relocatedDrawingZip.file("xl/custom/visuals/chart-scores.xml", await relocatedDrawingZip.file("xl/charts/chart2.xml").async("text"));
relocatedDrawingZip.remove("xl/charts/chart1.xml");
relocatedDrawingZip.remove("xl/charts/chart2.xml");
relocatedDrawingZip.file("xl/worksheets/_rels/sheet1.xml.rels", (await relocatedDrawingZip.file("xl/worksheets/_rels/sheet1.xml.rels").async("text")).replace("../drawings/drawing1.xml", "../custom/visuals/workbook-drawing.xml"));
relocatedDrawingZip.file("[Content_Types].xml", (await relocatedDrawingZip.file("[Content_Types].xml").async("text"))
  .replace("/xl/drawings/drawing1.xml", "/xl/custom/visuals/workbook-drawing.xml")
  .replace("/xl/charts/chart1.xml", "/xl/custom/visuals/chart-revenue.xml")
  .replace("/xl/charts/chart2.xml", "/xl/custom/visuals/chart-scores.xml"));
const relocatedDrawingBlob = new FileBlob(await relocatedDrawingZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type });
const relocatedDrawingInspect = await SpreadsheetFile.inspectXlsx(relocatedDrawingBlob);
assert.equal(relocatedDrawingInspect.ok, true);
assert.deepEqual(relocatedDrawingInspect.issues, []);
const relocatedDrawingBook = await SpreadsheetFile.importXlsx(relocatedDrawingBlob);
const relocatedDrawingSheet = relocatedDrawingBook.worksheets.getItem("Sheet1");
assert.equal(relocatedDrawingSheet.images.items.length, 1);
assert.equal(relocatedDrawingSheet.charts.items.length, 2);
assert.equal(relocatedDrawingSheet.images.items[0].alt, "Logo placeholder");
const relocatedRevenueChart = relocatedDrawingSheet.charts.items.find((chart) => chart.title === "Revenue Trend");
const relocatedScoreChart = relocatedDrawingSheet.charts.items.find((chart) => chart.title === "Scores");
assert.deepEqual(relocatedRevenueChart.series.items[0].values, [100, 120, 130]);
assert.equal(relocatedRevenueChart.series.items[0].formula, "Sheet1!$G$2:$G$4");
assert.equal(relocatedRevenueChart.series.items[0].categoryFormula, "Sheet1!$F$2:$F$4");
assert.ok(relocatedRevenueChart.position.width > 100 && relocatedRevenueChart.position.height > 80);
assert.deepEqual(relocatedScoreChart.position, { left: 240, top: 340, width: 240, height: 160 });
assert.match(relocatedDrawingSheet.toSvg(), /Revenue Trend/);
assert.match(relocatedDrawingSheet.toSvg(), /data:image\/png;base64/);
assert.ok(relocatedDrawingSheet.layoutJson().charts.some((chart) => chart.title === "Revenue Trend"));
assert.ok(relocatedDrawingBook.verify().issues.every((issue) => !["emptyChart", "emptyImage", "chartDataMismatch", "invalidImageDataUrl"].includes(issue.type)));
const wasmRelocatedDrawingBook = await importXlsxWithOpenChestnut(relocatedDrawingBlob);
wasmRelocatedDrawingBook.worksheets.getItem("Sheet1").getRange("A1").values = [["WASM preserved"]];
const wasmRelocatedDrawingRoundtrip = await exportXlsxWithOpenChestnut(wasmRelocatedDrawingBook, { recalculate: false });
const wasmRelocatedDrawingInspect = await SpreadsheetFile.inspectXlsx(wasmRelocatedDrawingRoundtrip);
assert.equal(wasmRelocatedDrawingInspect.ok, true, JSON.stringify(wasmRelocatedDrawingInspect.issues));
assert.ok(wasmRelocatedDrawingInspect.parts.some((part) => part.path === "xl/custom/visuals/workbook-drawing.xml"));
assert.ok(wasmRelocatedDrawingInspect.parts.some((part) => part.path === "xl/custom/visuals/chart-revenue.xml"));
assert.ok(wasmRelocatedDrawingInspect.parts.some((part) => part.path === "xl/assets/logo.png"));
const wasmRelocatedDrawingImported = await SpreadsheetFile.importXlsx(wasmRelocatedDrawingRoundtrip);
assert.equal(wasmRelocatedDrawingImported.worksheets.getItem("Sheet1").getRange("A1").values[0][0], "WASM preserved");
assert.equal(wasmRelocatedDrawingImported.worksheets.getItem("Sheet1").images.items.length, 1);
assert.equal(wasmRelocatedDrawingImported.worksheets.getItem("Sheet1").charts.items.length, 2);
const relocatedDrawingRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(relocatedDrawingBook));
assert.equal(relocatedDrawingRoundtrip.worksheets.getItem("Sheet1").images.items.length, 1);
assert.equal(relocatedDrawingRoundtrip.worksheets.getItem("Sheet1").charts.items.length, 2);
const styleFidelityZip = await JSZip.loadAsync(xlsxBytes);
styleFidelityZip.remove("customXml/open-office-artifact.json");
const assignStyle = (xml, address, styleIndex) => xml.replace(new RegExp(`<c r="${address}"([^>]*)>`), (match, attrs) => `<c r="${address}"${attrs.replace(/\s+s="[^"]*"/, "")} s="${styleIndex}">`);
let styleFidelityXml = assignStyle(worksheetXml, "A2", 1);
styleFidelityXml = assignStyle(styleFidelityXml, "B2", 2);
styleFidelityXml = assignStyle(styleFidelityXml, "C2", 3);
styleFidelityZip.file("xl/worksheets/sheet1.xml", styleFidelityXml);
styleFidelityZip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b val="false"/><i val="true"/><u val="double"/><strike val="true"/><sz val="12"/><color rgb="80123456"/><name val="Liberation Sans"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="7FFDE68A"/></patternFill></fill></fills><borders count="2"><border/><border diagonalUp="false"><left style="thin"><color rgb="40334155"/></left><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="4" fontId="1" fillId="2" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"><alignment horizontal="center" vertical="center" textRotation="30" wrapText="true" indent="1" shrinkToFit="true" readingOrder="2"/><protection locked="false" hidden="true"/></xf></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1" applyNumberFormat="0" applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/></cellXfs><dxfs count="0"/></styleSheet>`);
const styleFidelityBook = await SpreadsheetFile.importXlsx(new FileBlob(await styleFidelityZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
const styleFidelitySheet = styleFidelityBook.worksheets.getItem("Sheet1");
assert.equal(styleFidelitySheet.getRange("A2").format.numberFormat, "mm-dd-yy");
assert.equal(styleFidelitySheet.getRange("B2").format.numberFormat, "0.00%");
assert.deepEqual(styleFidelitySheet.getRange("C2").format.font, { bold: false, italic: true, underline: "double", strike: true, color: "#123456", size: 12, name: "Liberation Sans" });
assert.equal(styleFidelitySheet.getRange("C2").format.fill, "#FDE68A");
assert.deepEqual(styleFidelitySheet.getRange("C2").format.alignment, { horizontal: "center", vertical: "center", wrapText: true, shrinkToFit: true, textRotation: 30, indent: 1, readingOrder: 2 });
assert.deepEqual(styleFidelitySheet.getRange("C2").format.protection, { locked: false, hidden: true });
const styleFidelityLayout = styleFidelitySheet.layoutJson({ range: "A2:C2" });
assert.deepEqual(styleFidelityLayout.cells.map((cell) => cell.displayValue), ["01-02-00", "300.00%", "5.00"]);
const styleFidelitySvg = styleFidelitySheet.toSvg();
assert.match(styleFidelitySvg, />300\.00%<\/text>/);
assert.match(styleFidelitySvg, /text-decoration="underline line-through"/);
assert.match(styleFidelitySvg, /transform="rotate\(-30 /);
const styleFidelityRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(styleFidelityBook));
assert.deepEqual(styleFidelityRoundtrip.worksheets.getItem("Sheet1").getRange("C2").format.alignment, styleFidelitySheet.getRange("C2").format.alignment);
assert.deepEqual(styleFidelityRoundtrip.worksheets.getItem("Sheet1").getRange("C2").format.protection, { locked: false, hidden: true });
const themeStyleZip = await JSZip.loadAsync(xlsxBytes);
themeStyleZip.remove("customXml/open-office-artifact.json");
themeStyleZip.file("xl/worksheets/sheet1.xml", assignStyle(worksheetXml, "A2", 1));
themeStyleZip.file("xl/_rels/workbook.xml.rels", workbookRelsXml.replace(/<Relationship\b(?=[^>]*Type="[^"]*\/theme")[^>]*\/>/, `<Relationship Id="rTheme99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="appearance/custom-theme.xml"/>`));
themeStyleZip.file("xl/appearance/custom-theme.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Clean room theme"><a:themeElements><a:clrScheme name="Clean room"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="336699"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>`);
themeStyleZip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><sz val="11"/><color theme="4" tint="0.5"/><name val="Aptos"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor indexed="0"/></patternFill></fill></fills><borders count="2"><border/><border diagonalUp="1"><left style="thin"><color theme="4" tint="-0.25"/></left><right style="medium"><color indexed="0"/></right><top/><bottom style="double"><color rgb="FF112233"/></bottom><diagonal style="dashed"><color auto="1"/></diagonal></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs><colors><indexedColors><rgbColor rgb="FFABCDEF"/></indexedColors></colors></styleSheet>`);
const themeStyleBook = await SpreadsheetFile.importXlsx(new FileBlob(await themeStyleZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
const themeStyleRange = themeStyleBook.worksheets.getItem("Sheet1").getRange("A2");
assert.deepEqual(themeStyleRange.format.font.color, { theme: 4, tint: 0.5, resolved: "#8CB3D9" });
assert.deepEqual(themeStyleRange.format.fill, { patternType: "solid", foreground: { indexed: 0, resolved: "#ABCDEF" } });
assert.deepEqual(themeStyleRange.format.border, {
  left: { style: "thin", color: { theme: 4, tint: -0.25, resolved: "#264D73" } },
  right: { style: "medium", color: { indexed: 0, resolved: "#ABCDEF" } },
  bottom: { style: "double", color: "#112233" },
  diagonal: { style: "dashed", color: { auto: true, resolved: "#000000" } },
  diagonalUp: true,
});
assert.match(themeStyleBook.worksheets.getItem("Sheet1").toSvg(), /fill="#ABCDEF"/);
const themeStyleRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(themeStyleBook));
assert.deepEqual(themeStyleRoundtrip.worksheets.getItem("Sheet1").getRange("A2").format.font.color, themeStyleRange.format.font.color);
assert.deepEqual(themeStyleRoundtrip.worksheets.getItem("Sheet1").getRange("A2").format.border, themeStyleRange.format.border);
const patternBook = Workbook.create({ theme: { name: "Pattern Theme", colors: { accent1: "#336699", accent2: "#C0504D" } } });
const patternSheet = patternBook.worksheets.add("Patterns");
patternSheet.getRange("A1:B2").values = [["Pattern", "Theme"], ["Grid", "Tint"]];
patternSheet.getRange("A1:B2").format = {
  fill: { patternType: "darkGrid", foreground: { theme: 4, tint: 0.25 }, background: "#FFF7ED" },
  font: { bold: true, color: { theme: 4, tint: -0.25 } },
  border: { style: "thin", color: { theme: 5 } },
};
const patternXlsx = await SpreadsheetFile.exportXlsx(patternBook);
const patternZip = await JSZip.loadAsync(new Uint8Array(await patternXlsx.arrayBuffer()));
const patternStylesXml = await patternZip.file("xl/styles.xml").async("text");
assert.match(patternStylesXml, /<patternFill patternType="darkGrid"><fgColor theme="4" tint="0.25"\/><bgColor rgb="FFFFF7ED"\/><\/patternFill>/);
assert.match(patternStylesXml, /<color theme="4" tint="-0.25"\/>/);
assert.match(patternStylesXml, /<color theme="5"\/>/);
const patternThemeXml = await patternZip.file("xl/theme/theme1.xml").async("text");
assert.match(patternThemeXml, /<a:accent1><a:srgbClr val="336699"\/><\/a:accent1>/);
patternZip.file("xl/theme/theme1.xml", patternThemeXml.replaceAll("<a:", "<d:").replaceAll("</a:", "</d:").replace("xmlns:a=", "xmlns:d="));
patternZip.remove("customXml/open-office-artifact.json");
const importedPatternBook = await SpreadsheetFile.importXlsx(new FileBlob(await patternZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
assert.equal(importedPatternBook.theme.name, "Pattern Theme");
assert.equal(importedPatternBook.theme.colors.accent1, "#336699");
assert.deepEqual(importedPatternBook.worksheets.getItem("Patterns").getRange("A1").format.fill, {
  patternType: "darkGrid",
  foreground: { theme: 4, tint: 0.25, resolved: "#538CC6" },
  background: "#FFF7ED",
});
assert.deepEqual(importedPatternBook.worksheets.getItem("Patterns").getRange("A1").format.font.color, { theme: 4, tint: -0.25, resolved: "#264D73" });
const patternSvg = importedPatternBook.worksheets.getItem("Patterns").toSvg();
assert.match(patternSvg, /<pattern id="fill-0-0"/);
assert.match(patternSvg, /fill="url\(#fill-0-0\)"/);
assert.match(patternSvg, /stroke="#538CC6"/);
assert.match(importedPatternBook.inspect({ kind: "theme,style", maxChars: 8000 }).ndjson, /Pattern Theme/);
const patternRoundtripZip = await JSZip.loadAsync(new Uint8Array(await (await SpreadsheetFile.exportXlsx(importedPatternBook)).arrayBuffer()));
assert.match(await patternRoundtripZip.file("xl/styles.xml").async("text"), /patternType="darkGrid"><fgColor theme="4" tint="0.25"/);
for (const invalidStyle of [
  { fill: { patternType: "unsupported", foreground: "#FFFFFF" } },
  { fill: { patternType: "solid" } },
  { font: { color: { theme: 12 } } },
  { border: { style: "thin", color: { theme: 4, tint: 1.1 } } },
]) {
  const invalidBook = Workbook.create();
  invalidBook.worksheets.add("Invalid").getRange("A1").format = invalidStyle;
  await assert.rejects(() => SpreadsheetFile.exportXlsx(invalidBook), /patternType|requires foreground|theme color index|tint must be/);
}
const nativeColorScaleInspect = nativeOnlyWorkbook.inspect({ kind: "conditionalFormat,computedStyle", target: "Sheet1!G4", maxChars: 12000 }).ndjson;
assert.match(nativeColorScaleInspect, /"ruleType":"colorScale"/);
assert.match(nativeColorScaleInspect, /"fill":"#22c55e"/);
const pivotRoundtripWorkbook = await SpreadsheetFile.importXlsx(xlsx);
assert.deepEqual(pivotRoundtripWorkbook.resolve("RevenuePivot").computedValues()[1], ["Jan", 100]);
assert.match(pivotRoundtripWorkbook.inspect({ kind: "pivotTable", target: "RevenuePivot", maxChars: 8000 }).ndjson, /Revenue sum/);
const sharedFormulaZip = await JSZip.loadAsync(xlsxBytes);
sharedFormulaZip.remove("customXml/open-office-artifact.json");
const sharedFormulaXml = worksheetXml
  .replace(/<c r="C2"([^>]*)>([\s\S]*?)<f>A2\+B2<\/f>([\s\S]*?)<\/c>/, `<c r="C2"$1>$2<f t="shared" si="0" ref="C2:C3">A2+B2</f>$3</c>`)
  .replace(/<c r="C3"([^>]*)>([\s\S]*?)<f>A3\+B3<\/f>([\s\S]*?)<\/c>/, `<c r="C3"$1>$2<f t="shared" si="0"></f>$3</c>`)
  .replace(/<c r="E3"([^>]*)>([\s\S]*?)<f>SUM\(RevenueData\)<\/f>([\s\S]*?)<\/c>/, `<c r="E3"$1>$2<f t="array" ref="E3:E4">SUM(G2:G4)</f>$3</c>`);
assert.match(sharedFormulaXml, /<f t="shared" si="0" ref="C2:C3">A2\+B2<\/f>/);
assert.match(sharedFormulaXml, /<f t="array" ref="E3:E4">SUM\(G2:G4\)<\/f>/);
sharedFormulaZip.file("xl/worksheets/sheet1.xml", sharedFormulaXml);
const sharedFormulaWorkbook = await SpreadsheetFile.importXlsx(new FileBlob(await sharedFormulaZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
const sharedFormulaInspect = sharedFormulaWorkbook.inspect({ kind: "formula", target: "Sheet1!C3", maxChars: 8000 }).ndjson;
assert.match(sharedFormulaInspect, /"formula":"=A3\+B3"/);
assert.match(sharedFormulaInspect, /"formulaType":"shared"/);
assert.match(sharedFormulaInspect, /"sharedIndex":0/);
assert.match(sharedFormulaInspect, /"sharedRef":"C2:C3"/);
const arrayFormulaInspect = sharedFormulaWorkbook.inspect({ kind: "formula", target: "Sheet1!E3", maxChars: 8000 }).ndjson;
assert.match(arrayFormulaInspect, /"formula":"=SUM\(G2:G4\)"/);
assert.match(arrayFormulaInspect, /"formulaType":"array"/);
assert.match(arrayFormulaInspect, /"arrayRef":"E3:E4"/);
assert.equal(sharedFormulaWorkbook.worksheets.getItem("Sheet1").getRange("C3").values[0][0], 12);
assert.equal(sharedFormulaWorkbook.worksheets.getItem("Sheet1").getRange("E3").values[0][0], 350);
const sharedFormulaRoundtrip = await SpreadsheetFile.exportXlsx(sharedFormulaWorkbook);
const sharedFormulaRoundtripZip = await JSZip.loadAsync(new Uint8Array(await sharedFormulaRoundtrip.arrayBuffer()));
const sharedFormulaRoundtripXml = await sharedFormulaRoundtripZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(sharedFormulaRoundtripXml, /<f t="shared" si="0" ref="C2:C3">A2\+B2<\/f>/);
assert.match(sharedFormulaRoundtripXml, /<c r="C3"[^>]*>[\s\S]*<f t="shared" si="0"><\/f>/);
assert.match(sharedFormulaRoundtripXml, /<f t="array" ref="E3:E4">SUM\(G2:G4\)<\/f>/);
const sharedFormulaRoundtripImport = await SpreadsheetFile.importXlsx(sharedFormulaRoundtrip);
assert.match(sharedFormulaRoundtripImport.inspect({ kind: "formula", target: "Sheet1!C3", maxChars: 8000 }).ndjson, /"formulaType":"shared"/);
assert.match(sharedFormulaRoundtripImport.inspect({ kind: "formula", target: "Sheet1!E3", maxChars: 8000 }).ndjson, /"formulaType":"array"/);
const nativeRuleInspect = nativeOnlyWorkbook.inspect({ kind: "dataValidation,conditionalFormat", maxChars: 12000 }).ndjson;
assert.match(nativeRuleInspect, /"kind":"dataValidation"/);
assert.match(nativeRuleInspect, /"range":"D2:D3"/);
assert.match(nativeRuleInspect, /Not Started/);
assert.match(nativeRuleInspect, /"kind":"conditionalFormat"/);
assert.match(nativeRuleInspect, /"ruleType":"cellIs"/);
assert.match(nativeRuleInspect, /"ruleType":"expression"/);
const nativeComputedStyle = nativeOnlyWorkbook.inspect({ kind: "computedStyle", target: "Sheet1!C3", maxChars: 8000 }).ndjson;
assert.match(nativeComputedStyle, /"kind":"computedStyle"/);
assert.match(nativeComputedStyle, /"fill":"#22C55E"/);
assert.match(nativeOnlyWorkbook.inspect({ kind: "table", range: "A1:D4" }).ndjson, /"A"/);
assert.match(nativeOnlyWorkbook.inspect({ kind: "definedName", target: "RevenueData" }).ndjson, /Sheet1!G2:G4/);
assert.equal(nativeOnlyWorkbook.definedNames.getItem("RevenueData").comment, "Revenue data body");
assert.equal(nativeOnlyWorkbook.definedNames.getItem("RevenueData").hidden, false);
assert.equal(nativeOnlyWorkbook.definedNames.getItem("LocalRevenueData", "Sheet1").hidden, true);
assert.equal(nativeOnlyWorkbook.worksheets.getItem("Sheet1").getRange("E3").values[0][0], 350);
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.xlsx`);
await xlsx.save(out);
const loaded = await SpreadsheetFile.importXlsx(await FileBlob.load(out));
assert.equal(loaded.worksheets.getItem("Sheet1").showGridLines, false);
assert.deepEqual(loaded.worksheets.getItem("Sheet1").freezePanes.toJSON(), { rows: 1, columns: 2, frozen: true, topLeftCell: "C2", activePane: "bottomRight" });
const loadedTasksTable = loaded.worksheets.getItem("Sheet1").tables.getItemOrNullObject("TasksTable");
assert.equal(loadedTasksTable.showFirstColumn, true);
assert.equal(loadedTasksTable.showLastColumn, true);
assert.equal(loadedTasksTable.showRowStripes, false);
const roundtrip = loaded.inspect({ kind: "table,formula", range: "A1:C3" }).ndjson;
assert.match(roundtrip, /"values":\[\["A","B","Sum"\],\[2,3,5\],\[5,7,12\]\]/);
assert.match(roundtrip, /"formula":"=A2\+B2"/);
const roundtripMetadata = loaded.inspect({ kind: "dataValidation,conditionalFormat,thread,table,pivotTable,drawing", maxChars: 16000 }).ndjson;
assert.match(roundtripMetadata, /"kind":"dataValidation"/);
assert.match(roundtripMetadata, /"kind":"conditionalFormat"/);
assert.match(roundtripMetadata, /"kind":"thread"/);
assert.match(roundtripMetadata, /TasksTable/);
assert.match(roundtripMetadata, /RevenuePivot/);
assert.match(roundtripMetadata, /Revenue sum/);
assert.match(roundtripMetadata, /Revenue Trend/);
assert.match(roundtripMetadata, /ScoresChart/);
assert.match(roundtripMetadata, /LogoImage/);
assert.match(roundtripMetadata, /"kind":"sparkline"/);
assert.match(roundtripMetadata, /Reviewed by model/);
assert.match(loaded.inspect({ kind: "definedName", target: "RevenueData" }).ndjson, /Sheet1!G2:G4/);
assert.equal(loaded.definedNames.getItem("RevenueData").hidden, false);
assert.equal(loaded.definedNames.getItem("LocalRevenueData", "Sheet1").hidden, true);
assert.equal(loaded.worksheets.getItem("Sheet1").getRange("E3").values[0][0], 350);
const loadedThreadId = JSON.parse(roundtripMetadata.split("\n").find((line) => line.includes('"kind":"thread"'))).id;
assert.equal(loaded.resolve(loadedThreadId).resolved, true);
assert.deepEqual(loaded.trace("Sheet1!C3").tree.precedents.map((node) => node.address), ["A3", "B3"]);
const refreshOnlyBook = Workbook.create();
const refreshOnlySheet = refreshOnlyBook.worksheets.add("Source");
refreshOnlySheet.getRange("A1:B3").values = [["Team", "Score"], ["Red", 4], ["Blue", 6]];
refreshOnlySheet.pivotTables.add({
  name: "RefreshOnlyPivot",
  sourceRange: "A1:B3",
  targetRange: "D1:E3",
  rowFields: ["Team"],
  valueFields: [{ field: "Score", summarizeBy: "sum" }],
  refreshPolicy: { saveData: false, refreshOnLoad: true },
});
const refreshOnlyXlsx = await SpreadsheetFile.exportXlsx(refreshOnlyBook);
assert.equal((await SpreadsheetFile.inspectXlsx(refreshOnlyXlsx)).ok, true);
const refreshOnlyZip = await JSZip.loadAsync(new Uint8Array(await refreshOnlyXlsx.arrayBuffer()));
assert.equal(refreshOnlyZip.file("xl/pivotCache/pivotCacheRecords1.xml"), null);
assert.equal(refreshOnlyZip.file("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"), null);
assert.doesNotMatch(await refreshOnlyZip.file("[Content_Types].xml").async("text"), /pivotCacheRecords/);
assert.doesNotMatch(await refreshOnlyZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text"), /r:id=/);
assert.match(await refreshOnlyZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text"), /refreshOnLoad="1" saveData="0"/);
refreshOnlyZip.remove("customXml/open-office-artifact.json");
const refreshOnlyNative = await SpreadsheetFile.importXlsx(new FileBlob(await refreshOnlyZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: refreshOnlyXlsx.type }));
const importedRefreshOnlyPivot = refreshOnlyNative.resolve("RefreshOnlyPivot");
assert.deepEqual(importedRefreshOnlyPivot.computedValues(), [["Team", "sum of Score"], ["Red", 4], ["Blue", 6]]);
assert.equal(importedRefreshOnlyPivot.refreshPolicy.saveData, false);
console.log("spreadsheet smoke ok");
