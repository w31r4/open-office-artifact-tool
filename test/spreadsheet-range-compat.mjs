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

const arithmeticWorkbook = Workbook.create();
const arithmeticSheet = arithmeticWorkbook.worksheets.add("Arithmetic");
arithmeticSheet.getRange("A1:B2").values = [[-1000, -1000], [2.7e-11, 0]];
arithmeticSheet.getRange("C1:C3").formulas = [["=A1-B1"], ["=A2-B2"], ["=1E-7"]];
assert.equal(arithmeticSheet.getRange("C1").values[0][0], 0);
assert.equal(arithmeticSheet.getRange("C2").values[0][0], 2.7e-11);
assert.equal(arithmeticSheet.getRange("C3").values[0][0], 1e-7);

const financialWorkbook = Workbook.create();
const financialSheet = financialWorkbook.worksheets.add("Cash Flows");
financialSheet.getRange("A1:A3").values = [[-100], [60], [60]];
financialSheet.getRange("B1:B3").formulas = [
  ["=DATE(2017,1,1)"],
  ["=DATE(2018,1,1)"],
  ["=DATE(2019,1,1)"],
];
financialSheet.getRange("C1:C6").formulas = [
  ["=NPV(0.1,A1:A3)"],
  ["=XNPV(0.1,A1:A3,B1:B3)"],
  ["=IRR(A1:A3)"],
  ["=XIRR(A1:A3,B1:B3)"],
  ["=PMT(0.008333333333333333,12,1000)"],
  ["=PMT(0,2,100)"],
];
const financialValues = financialSheet.getRange("C1:C6").values.flat();
assert.ok(Math.abs(financialValues[0] - 3.7565740045078897) < 1e-10);
assert.ok(Math.abs(financialValues[1] - 4.132231404958681) < 1e-10);
assert.ok(Math.abs(financialValues[2] - 0.1306623862918075) < 1e-10);
assert.ok(Math.abs(financialValues[3] - 0.1306623862918075) < 1e-10);
assert.ok(Math.abs(financialValues[4] + 87.91588723000957) < 1e-10);
assert.equal(financialValues[5], -50);
financialSheet.getRange("P1:R12").formulas = Array.from({ length: 12 }, (_, index) => [
  "=PMT(0.01,12,100000)",
  `=IPMT(0.01,${index + 1},12,100000)`,
  `=PPMT(0.01,${index + 1},12,100000)`,
]);
const amortizationSchedule = financialSheet.getRange("P1:R12").values;
assert.equal(amortizationSchedule[0][1], -1000);
assert.ok(Math.abs(amortizationSchedule[0][2] + 7884.878867834168) < 1e-10);
assert.ok(Math.abs(amortizationSchedule[1][1] + 921.1512113216584) < 1e-10);
for (const [payment, interest, principal] of amortizationSchedule) {
  assert.ok(Math.abs(payment - interest - principal) < 1e-10);
}
assert.ok(Math.abs(100000 + amortizationSchedule.reduce((balance, [, , principal]) => balance + principal, 0)) < 1e-8);
financialSheet.getRange("S1:S11").formulas = [
  ["=PMT(0.1,2,1000,0,1)"],
  ["=IPMT(0.1,1,2,1000,0,1)"],
  ["=IPMT(0.1,2,2,1000,0,1)"],
  ["=PPMT(0.1,2,2,1000,0,1)"],
  ["=IPMT(0,1,2,100)"],
  ["=PPMT(0,1,2,100)"],
  ["=IPMT(0.1,0,2,1000)"],
  ["=IPMT(0.1,2.5,2,1000)"],
  ["=IPMT(0.1,3,2,1000)"],
  ["=PPMT(0.1,1,2,1000,0,2)"],
  ["=IPMT(0.1,1,2)"],
];
const duePaymentValues = financialSheet.getRange("S1:S11").values.flat();
assert.ok(Math.abs(duePaymentValues[0] + 523.8095238095239) < 1e-10);
assert.equal(duePaymentValues[1], 0);
assert.ok(Math.abs(duePaymentValues[2] + 52.38095238095239) < 1e-10);
assert.ok(Math.abs(duePaymentValues[3] + 471.42857142857144) < 1e-10);
assert.equal(duePaymentValues[4], 0);
assert.equal(duePaymentValues[5], -50);
assert.deepEqual(duePaymentValues.slice(6), ["#NUM!", "#NUM!", "#NUM!", "#NUM!", "#VALUE!"]);
financialSheet.getRange("U1:U8").formulas = [
  ["=RATE(12,-8884.878867834168,100000)"],
  ["=RATE(2,-523.8095238095239,1000,0,1)"],
  ["=RATE(12,-100,1200)"],
  ["=RATE(12,-800,8000,-1000,0)"],
  ["=RATE(10000,-10,1000)"],
  ["=RATE(12,-100,1200,0,2)"],
  ["=RATE(12,-100,1200,0,0,-1)"],
  ["=RATE(12,100,1000)"],
];
const rateValues = financialSheet.getRange("U1:U8").values.flat();
assert.ok(Math.abs(rateValues[0] - 0.01) < 1e-10);
assert.ok(Math.abs(rateValues[1] - 0.1) < 1e-10);
assert.ok(Math.abs(rateValues[2]) < 1e-10);
assert.ok(Math.abs(rateValues[3] - 0.0426446765858793) < 1e-10);
assert.deepEqual(rateValues.slice(4), ["#NUM!", "#NUM!", "#NUM!", "#NUM!"]);
financialSheet.getRange("V1:V3").formulas = [
  ["=SLN(1000,100,5)"],
  ["=SLN(-1000,100,5)"],
  ["=SLN(1000,100,0)"],
];
assert.deepEqual(financialSheet.getRange("V1:V3").values, [[180], [-220], ["#DIV/0!"]]);
financialSheet.getRange("W1:W10").formulas = [
  ["=DB(1000,100,5,1)"],
  ["=DB(1000,100,5,2)"],
  ["=DB(1000,100,5,5)"],
  ["=DB(1000,100,5,1,6)"],
  ["=DB(1000,100,5,6,6)"],
  ["=DB(1000,100,5,6)"],
  ["=DB(1000,100,5,1,0)"],
  ["=DB(1000,100,5,0)"],
  ["=DB(1000,1200,5,1)"],
  ["=DB(0,0,5,1)"],
];
const dbValues = financialSheet.getRange("W1:W10").values.flat();
assert.equal(dbValues[0], 369);
assert.equal(dbValues[1], 232.839);
assert.ok(Math.abs(dbValues[2] - 58.498375128849) < 1e-10);
assert.equal(dbValues[3], 184.5);
assert.ok(Math.abs(dbValues[4] - 23.8527124587882) < 1e-10);
assert.deepEqual(dbValues.slice(5, 9), ["#NUM!", "#NUM!", "#NUM!", "#NUM!"]);
assert.equal(dbValues[9], 0);
financialSheet.getRange("X1:X10").formulas = [
  ["=DDB(1000,100,5,1)"],
  ["=DDB(1000,100,5,2)"],
  ["=DDB(1000,100,5,5)"],
  ["=DDB(1000,100,5,1,1.5)"],
  ["=DDB(1000,100,5,5,1.5)"],
  ["=DDB(1000,100,5,6)"],
  ["=DDB(1000,100,5,0)"],
  ["=DDB(1000,100,5,1,0)"],
  ["=DDB(1000,1200,5,1)"],
  ["=DDB(0,0,5,1)"],
];
const ddbValues = financialSheet.getRange("X1:X10").values.flat();
assert.equal(ddbValues[0], 400);
assert.equal(ddbValues[1], 240);
assert.ok(Math.abs(ddbValues[2] - 29.6) < 1e-10);
assert.equal(ddbValues[3], 300);
assert.ok(Math.abs(ddbValues[4] - 72.03) < 1e-10);
assert.deepEqual(ddbValues.slice(5, 9), ["#NUM!", "#NUM!", "#NUM!", "#NUM!"]);
assert.equal(ddbValues[9], 0);
financialSheet.getRange("Y1:Y17").formulas = [
  ["=PV(0.01,12,-8884.878867834168)"],
  ["=FV(0.01,12,-8884.878867834168,100000)"],
  ["=NPER(0.01,-8884.878867834168,100000)"],
  ["=PV(0,2,-50)"],
  ["=FV(0,2,-50,100)"],
  ["=NPER(0,-50,100)"],
  ["=PV(0.1,2,-523.8095238095239,0,1)"],
  ["=FV(0.1,2,-523.8095238095239,1000,1)"],
  ["=NPER(0.1,-523.8095238095239,1000,0,1)"],
  ["=NPER(0.1,100,100)"],
  ["=NPER(0.1,-100,0)"],
  ["=PV(0.01,0,-1)"],
  ["=FV(0.01,0,-1)"],
  ["=NPER(0,0,100)"],
  ["=PV(0.1,2,-1,0,2)"],
  ["=FV(0.1,2,-1,0,2)"],
  ["=NPER(0.1,-1,1,0,2)"],
];
const inversePaymentValues = financialSheet.getRange("Y1:Y17").values.flat();
assert.ok(Math.abs(inversePaymentValues[0] - 100000) < 1e-8);
assert.ok(Math.abs(inversePaymentValues[1]) < 1e-8);
assert.ok(Math.abs(inversePaymentValues[2] - 12) < 1e-10);
assert.equal(inversePaymentValues[3], 100);
assert.ok(Math.abs(inversePaymentValues[4]) < 1e-10);
assert.equal(inversePaymentValues[5], 2);
assert.ok(Math.abs(inversePaymentValues[6] - 1000) < 1e-8);
assert.ok(Math.abs(inversePaymentValues[7]) < 1e-8);
assert.ok(Math.abs(inversePaymentValues[8] - 2) < 1e-10);
assert.ok(Math.abs(inversePaymentValues[9] + 1) < 1e-10);
assert.equal(inversePaymentValues[10], 0);
assert.deepEqual(inversePaymentValues.slice(11), ["#NUM!", "#NUM!", "#NUM!", "#NUM!", "#NUM!", "#NUM!"]);
financialSheet.getRange("Z1:Z15").formulas = [
  ["=CUMIPMT(0.01,12,100000,1,12,0)"],
  ["=CUMPRINC(0.01,12,100000,1,12,0)"],
  ["=CUMIPMT(0.01,12,100000,1,1,0)"],
  ["=CUMPRINC(0.01,12,100000,1,1,0)"],
  ["=CUMIPMT(0.1,2,1000,1,2,1)"],
  ["=CUMPRINC(0.1,2,1000,1,2,1)"],
  ["=CUMIPMT(0,12,100000,1,12,0)"],
  ["=CUMPRINC(0.01,12,0,1,12,0)"],
  ["=CUMIPMT(0.01,12,100000,0,1,0)"],
  ["=CUMIPMT(0.01,12,100000,2,1,0)"],
  ["=CUMIPMT(0.01,12,100000,1,13,0)"],
  ["=CUMIPMT(0.01,12,100000,1,12,2)"],
  ["=CUMIPMT(0.01,12,100000,1,1.5,0)"],
  ["=CUMPRINC(0.01,10000,100000,1,10000,0)"],
  ["=CUMIPMT(0.01,12,100000,1,12)"],
];
const cumulativePaymentValues = financialSheet.getRange("Z1:Z15").values.flat();
assert.ok(Math.abs(cumulativePaymentValues[0] + 6618.54641401005) < 1e-8);
assert.ok(Math.abs(cumulativePaymentValues[1] + 100000) < 1e-8);
assert.equal(cumulativePaymentValues[2], -1000);
assert.ok(Math.abs(cumulativePaymentValues[3] + 7884.878867834168) < 1e-8);
assert.ok(Math.abs(cumulativePaymentValues[4] + 52.3809523809524) < 1e-10);
assert.ok(Math.abs(cumulativePaymentValues[5] + 995.238095238095) < 1e-10);
assert.deepEqual(cumulativePaymentValues.slice(6, 14), Array.from({ length: 8 }, () => "#NUM!"));
assert.equal(cumulativePaymentValues[14], "#VALUE!");
financialSheet.getRange("AA1:AA8").values = [[-100000], [30000], [35000], [40000], [45000], [-100], [0], [100]];
financialSheet.getRange("AB1:AB11").formulas = [
  ["=MIRR(AA1:AA5,0.1,0.12)"],
  ["=MIRR(AA1:AA5,0,0)"],
  ["=MIRR(AA1:AA5,-0.5,-0.25)"],
  ["=MIRR(AA1:AA5,-1,0.1)"],
  ["=MIRR(AA1:AA5,0.1,-1)"],
  ["=MIRR(AA6:AA8,0.1,0.1)"],
  ["=MIRR(AA7:AA8,0.1,0.1)"],
  ["=MIRR(AA3:AA5,0.1,0.1)"],
  ["=MIRR(AA1:AA5,0.1,0.1,1)"],
  ["=MIRR(AA1:AA5,0.1)"],
  ["=MIRR(AA1:AA5,\"bad\",0.1)"],
];
const mirrValues = financialSheet.getRange("AB1:AB11").values.flat();
assert.ok(Math.abs(mirrValues[0] - 0.151560419415717) < 1e-12);
assert.ok(Math.abs(mirrValues[1] - 0.106681919700321) < 1e-12);
assert.ok(Math.abs(mirrValues[2] - 0.0178743975830735) < 1e-12);
assert.equal(mirrValues[5], 0);
assert.deepEqual(mirrValues.slice(3, 5), ["#NUM!", "#NUM!"]);
assert.deepEqual(mirrValues.slice(6, 8), ["#NUM!", "#NUM!"]);
assert.deepEqual(mirrValues.slice(8), ["#VALUE!", "#VALUE!", "#VALUE!"]);
financialSheet.getRange("G1:G2").values = [[-100], [110]];
financialSheet.getRange("H1:H2").formulas = [
  ["=DATE(2017,1,1)"],
  ["=DATE(2017,7,2)"],
];
financialSheet.getRange("I1:I2").formulas = [
  ["=XNPV(0.1,G1:G2,H1:H2)"],
  ["=XIRR(G1:G2,H1:H2)"],
];
const irregularDateReturns = financialSheet.getRange("I1:I2").values.flat();
assert.ok(Math.abs(irregularDateReturns[0] - 4.894579157536754) < 1e-10);
assert.ok(Math.abs(irregularDateReturns[1] - 0.2106338215370842) < 1e-10);
financialSheet.getRange("J1:J3").values = [[-100], [230], [-132]];
financialSheet.getRange("K1:K2").formulas = [["=IRR(J1:J3,0.1)"], ["=IRR(J1:J3,0.2)"]];
assert.ok(Math.abs(financialSheet.getRange("K1").values[0][0] - 0.1) < 1e-10);
assert.ok(Math.abs(financialSheet.getRange("K2").values[0][0] - 0.2) < 1e-9);
financialSheet.getRange("D1:D5").formulas = [
  ["=NPV(-1,A1:A3)"],
  ["=IRR(A1:A2, -1)"],
  ["=IRR(B1:B3)"],
  ["=XIRR(A1:A3,B1:B2)"],
  ["=XNPV(0.1,A1:A3,B1:B2)"],
];
assert.deepEqual(financialSheet.getRange("D1:D5").values, [["#NUM!"], ["#NUM!"], ["#NUM!"], ["#VALUE!"], ["#VALUE!"]]);
financialSheet.getRange("E1:E2").values = [[-100], ["not a cash flow"]];
financialSheet.getRange("F1").formulas = [["=IRR(E1:E2)"]];
assert.equal(financialSheet.getRange("F1").values[0][0], "#VALUE!");
financialSheet.getRange("M1:M2").values = [[new Date("2017-01-01T00:00:00.000Z")], [new Date("2017-07-02T00:00:00.000Z")]];
financialSheet.getRange("N1:N2").values = [[-100], [110]];
financialSheet.getRange("O1:O2").formulas = [["=XNPV(0.1,N1:N2,M1:M2)"], ["=XIRR(N1:N2,M1:M2)"]];
const dateObjectReturns = financialSheet.getRange("O1:O2").values.flat();
assert.ok(Math.abs(dateObjectReturns[0] - 4.894579157536754) < 1e-10);
assert.ok(Math.abs(dateObjectReturns[1] - 0.2106338215370842) < 1e-10);
const financialRoundTrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(financialWorkbook, { recalculate: false }));
assert.deepEqual(financialRoundTrip.worksheets.getItem("Cash Flows").getRange("C1:C6").formulas, financialSheet.getRange("C1:C6").formulas);
assert.deepEqual(financialRoundTrip.worksheets.getItem("Cash Flows").getRange("P1:R12").formulas, financialSheet.getRange("P1:R12").formulas);
assert.deepEqual(financialRoundTrip.worksheets.getItem("Cash Flows").getRange("U1:U8").formulas, financialSheet.getRange("U1:U8").formulas);
assert.deepEqual(financialRoundTrip.worksheets.getItem("Cash Flows").getRange("V1:X10").formulas, financialSheet.getRange("V1:X10").formulas);
assert.deepEqual(financialRoundTrip.worksheets.getItem("Cash Flows").getRange("Y1:Z15").formulas, financialSheet.getRange("Y1:Z15").formulas);
assert.deepEqual(financialRoundTrip.worksheets.getItem("Cash Flows").getRange("AB1:AB11").formulas, financialSheet.getRange("AB1:AB11").formulas);
const financialLimitWorkbook = Workbook.create();
const financialLimitSheet = financialLimitWorkbook.worksheets.add("Cash flow limit");
financialLimitSheet.getRange("A1:A10001").values = Array.from({ length: 10_001 }, (_, index) => [index === 0 ? -100 : 1]);
financialLimitSheet.getRange("B1").formulas = [["=NPV(0.1,A1:A10001)"]];
financialLimitSheet.getRange("C1").formulas = [["=MIRR(A1:A10001,0.1,0.1)"]];
assert.equal(financialLimitSheet.getRange("B1").values[0][0], "#NUM!");
assert.equal(financialLimitSheet.getRange("C1").values[0][0], "#NUM!");

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
