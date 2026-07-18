import { Workbook } from "open-office-artifact-tool";

import { sectionRange, titleRange } from "./common.mjs";

function headerFormat() {
  return { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
}

function configureColumns(sheet, widths) {
  for (const [column, width] of Object.entries(widths)) sheet.getRange(`${column}1:${column}40`).format.columnWidthPx = width;
}

export function buildFinancialBudget() {
  const workbook = Workbook.create({ calculation: { mode: "automatic", fullCalculationOnLoad: true } });
  const assumptions = workbook.worksheets.add("Assumptions");
  const plan = workbook.worksheets.add("Monthly Plan");
  const summary = workbook.worksheets.add("Budget Summary");
  for (const sheet of [assumptions, plan, summary]) sheet.showGridLines = false;

  titleRange(assumptions, "A1:D1", "Financial Budget — Assumptions", 4);
  sectionRange(assumptions, "A3:D3", "Planning inputs", 4);
  assumptions.getRange("A4:D8").values = [
    ["Input", "Value", "Owner", "Purpose"],
    ["Starting revenue", 120000, "Commercial", "Baseline"],
    ["Monthly growth", 0.04, "Commercial", "Revenue change"],
    ["Gross margin", 0.62, "Finance", "Profit target"],
    ["Fixed operating cost", 41000, "Operations", "Cost floor"],
  ];
  assumptions.getRange("A4:D4").format = headerFormat();
  assumptions.getRange("B5:B8").format = { fill: "#FFF2CC", font: { bold: true, color: "#1D4ED8" } };
  assumptions.getRange("B6:B7").format.numberFormat = "0.0%";
  assumptions.getRange("B5:B5").format.numberFormat = "$#,##0;[Red]($#,##0);-";
  assumptions.getRange("B8:B8").format.numberFormat = "$#,##0;[Red]($#,##0);-";
  configureColumns(assumptions, { A: 185, B: 120, C: 140, D: 190 });
  assumptions.freezePanes.freezeRows(4);

  titleRange(plan, "A1:E1", "Financial Budget — Monthly Plan", 5);
  plan.getRange("A3:E3").values = [["Month", "Revenue", "Gross profit", "Operating result", "Variance"]];
  plan.getRange("A3:E3").format = headerFormat();
  plan.getRange("A4:A9").values = [["Month 1"], ["Month 2"], ["Month 3"], ["Month 4"], ["Month 5"], ["Month 6"]];
  plan.getRange("B4").formulas = [["='Assumptions'!$B$5"]];
  plan.getRange("B5").formulas = [["=B4*(1+'Assumptions'!$B$6)"]];
  plan.getRange("B5:B9").fillDown();
  plan.getRange("C4:C9").formulas = [["=B4*'Assumptions'!$B$7"], ["=B5*'Assumptions'!$B$7"], ["=B6*'Assumptions'!$B$7"], ["=B7*'Assumptions'!$B$7"], ["=B8*'Assumptions'!$B$7"], ["=B9*'Assumptions'!$B$7"]];
  plan.getRange("D4:D9").formulas = [["=C4-'Assumptions'!$B$8"], ["=C5-'Assumptions'!$B$8"], ["=C6-'Assumptions'!$B$8"], ["=C7-'Assumptions'!$B$8"], ["=C8-'Assumptions'!$B$8"], ["=C9-'Assumptions'!$B$8"]];
  plan.getRange("E4:E9").formulas = [["=D4-$B$12"], ["=D5-$B$12"], ["=D6-$B$12"], ["=D7-$B$12"], ["=D8-$B$12"], ["=D9-$B$12"]];
  plan.getRange("A11:E11").values = [["Total / target", null, null, null, null]];
  plan.getRange("A11:E11").format = { fill: "#D9F0EE", font: { bold: true, color: "#0F3D4C" } };
  plan.getRange("B11:D11").formulas = [["=SUM(B4:B9)", "=SUM(C4:C9)", "=SUM(D4:D9)"]];
  plan.getRange("B12").formulas = [["=D11/6"]];
  plan.getRange("A12").values = [["Monthly target"]];
  plan.getRange("A12:B12").format = { fill: "#F8FAFC", font: { bold: true, color: "#0F3D4C" } };
  plan.getRange("A13").values = [["Fixed cost"]];
  plan.getRange("B13").formulas = [["='Assumptions'!$B$8"]];
  plan.getRange("A13:B13").format = { fill: "#F8FAFC", font: { bold: true, color: "#0F3D4C" } };
  plan.getRange("B4:E13").format.numberFormat = "$#,##0;[Red]($#,##0);-";
  configureColumns(plan, { A: 105, B: 130, C: 130, D: 135, E: 125 });
  plan.freezePanes.freezeRows(3);

  titleRange(summary, "A1:D1", "Financial Budget — Summary", 4);
  summary.getRange("A3:D3").values = [["Metric", "Value", "Check", "Status"]];
  summary.getRange("A3:D3").format = headerFormat();
  summary.getRange("A4:D7").values = [
    ["Six-month revenue", null, "Positive total", null],
    ["Six-month operating result", null, "Positive total", null],
    ["Average monthly result", null, "Meets target", null],
    ["Model status", null, "No failed checks", null],
  ];
  summary.getRange("B4:B7").formulas = [["='Monthly Plan'!$B$11"], ["='Monthly Plan'!$D$11"], ["='Monthly Plan'!$B$12"], ["=COUNTIF(D4:D6,\"CHECK\")"]];
  summary.getRange("D4:D6").formulas = [["=IF(B4>0,\"OK\",\"CHECK\")"], ["=IF(B5>0,\"OK\",\"CHECK\")"], ["=IF(B6>0,\"OK\",\"CHECK\")"]];
  summary.getRange("D7").formulas = [["=IF(B7=0,\"OK\",\"CHECK\")"]];
  summary.getRange("B4:B6").format = { fill: "#F0FDF4", font: { bold: true, color: "#166534" }, numberFormat: "$#,##0;[Red]($#,##0);-" };
  summary.getRange("B7:D7").format = { fill: "#F8FAFC", font: { bold: true, color: "#0F3D4C" } };
  summary.getRange("D4:D7").format = { fill: "#DCFCE7", font: { bold: true, color: "#166534" }, alignment: { horizontal: "center" } };
  summary.getRange("D4:D7").conditionalFormats.add("containsText", { text: "CHECK", format: { fill: "#FEE2E2", font: { bold: true, color: "#B91C1C" } } });
  configureColumns(summary, { A: 220, B: 145, C: 145, D: 90 });
  summary.freezePanes.freezeRows(3);
  workbook.worksheets.setActiveWorksheet("Budget Summary");
  workbook.recalculate();
  return workbook;
}

export function buildProjectTracker() {
  const workbook = Workbook.create({ calculation: { mode: "automatic", fullCalculationOnLoad: true } });
  const summary = workbook.worksheets.add("Project Summary");
  const plan = workbook.worksheets.add("Work Plan");
  const risks = workbook.worksheets.add("Risks and Decisions");
  for (const sheet of [summary, plan, risks]) sheet.showGridLines = false;

  titleRange(summary, "A1:D1", "Project Tracker — Summary", 4);
  summary.getRange("A3:D3").values = [["Metric", "Value", "Signal", "Status"]];
  summary.getRange("A3:D3").format = headerFormat();
  summary.getRange("A4:D8").values = [
    ["Completed deliverables", null, "At least two", null],
    ["In-progress deliverables", null, "Named owner", null],
    ["At-risk deliverables", null, "No unowned risk", null],
    ["Open decisions", null, "Dated owner", null],
    ["Tracker status", null, "No missing control", null],
  ];
  summary.getRange("B4:B8").formulas = [
    ["=COUNTIF('Work Plan'!$D$4:$D$9,\"Done\")"],
    ["=COUNTIF('Work Plan'!$D$4:$D$9,\"In progress\")"],
    ["=COUNTIF('Work Plan'!$E$4:$E$9,\"At risk\")"],
    ["=COUNTIF('Risks and Decisions'!$E$4:$E$7,\"Open\")"],
    ["=COUNTIF(D4:D7,\"CHECK\")"],
  ];
  summary.getRange("D4:D7").formulas = [
    ["=IF(B4>=2,\"OK\",\"CHECK\")"],
    ["=IF(B5>=1,\"OK\",\"CHECK\")"],
    ["=IF(B6<=1,\"OK\",\"CHECK\")"],
    ["=IF(B7<=2,\"OK\",\"CHECK\")"],
  ];
  summary.getRange("D8").formulas = [["=IF(B8=0,\"OK\",\"CHECK\")"]];
  summary.getRange("B4:B8").format = { fill: "#EFF6FF", font: { bold: true, color: "#1D4ED8" }, alignment: { horizontal: "center" } };
  summary.getRange("D4:D8").format = { fill: "#DCFCE7", font: { bold: true, color: "#166534" }, alignment: { horizontal: "center" } };
  summary.getRange("D4:D8").conditionalFormats.add("containsText", { text: "CHECK", format: { fill: "#FEE2E2", font: { bold: true, color: "#B91C1C" } } });
  configureColumns(summary, { A: 220, B: 88, C: 165, D: 92 });
  summary.freezePanes.freezeRows(3);

  titleRange(plan, "A1:E1", "Project Tracker — Work Plan", 5);
  plan.getRange("A3:E3").values = [["ID", "Deliverable", "Owner", "Status", "Health"]];
  plan.getRange("A3:E3").format = headerFormat();
  plan.getRange("A4:D9").values = [
    ["D-01", "Confirm acceptance criteria", "Product", "Done"],
    ["D-02", "Validate source boundary", "Engineering", "Done"],
    ["D-03", "Build bounded release", "Engineering", "In progress"],
    ["D-04", "Prepare rollback", "Operations", "In progress"],
    ["D-05", "Run pilot", "Product", "Not started"],
    ["D-06", "Publish decision record", "Decision owner", "Not started"],
  ];
  plan.getRange("E4:E9").formulas = [
    ["=IF(D4=\"Done\",\"On track\",\"On track\")"],
    ["=IF(D5=\"Done\",\"On track\",\"On track\")"],
    ["=IF(D6=\"In progress\",\"Watch\",\"At risk\")"],
    ["=IF(D7=\"In progress\",\"At risk\",\"At risk\")"],
    ["=IF(D8=\"Not started\",\"Watch\",\"On track\")"],
    ["=IF(D9=\"Not started\",\"Watch\",\"On track\")"],
  ];
  sectionRange(plan, "A11:E11", "Delivery evidence", 5);
  plan.getRange("A12:E12").values = [["ID", "Start", "Due", "Progress", "Completion evidence"]];
  plan.getRange("A12:E12").format = headerFormat();
  plan.getRange("A13:C18").values = [
    ["D-01", "Week 1", "Week 1"],
    ["D-02", "Week 1", "Week 2"],
    ["D-03", "Week 2", "Week 3"],
    ["D-04", "Week 2", "Week 3"],
    ["D-05", "Week 3", "Week 4"],
    ["D-06", "Week 4", "Week 4"],
  ];
  plan.getRange("D13:D18").formulas = [
    ["=IF(D4=\"Done\",1,IF(D4=\"In progress\",0.5,0))"],
    ["=IF(D5=\"Done\",1,IF(D5=\"In progress\",0.5,0))"],
    ["=IF(D6=\"Done\",1,IF(D6=\"In progress\",0.5,0))"],
    ["=IF(D7=\"Done\",1,IF(D7=\"In progress\",0.5,0))"],
    ["=IF(D8=\"Done\",1,IF(D8=\"In progress\",0.5,0))"],
    ["=IF(D9=\"Done\",1,IF(D9=\"In progress\",0.5,0))"],
  ];
  plan.getRange("E13:E18").values = [
    ["Acceptance approved"],
    ["Contract checks passed"],
    ["Implementation reviewed"],
    ["Rollback drill recorded"],
    ["Pilot outcome recorded"],
    ["Decision record published"],
  ];
  plan.getRange("D13:D18").format.numberFormat = "0%";
  plan.getRange("D4:D9").conditionalFormats.add("containsText", { text: "Done", format: { fill: "#DCFCE7", font: { bold: true, color: "#166534" } } });
  plan.getRange("E4:E9").conditionalFormats.add("containsText", { text: "At risk", format: { fill: "#FEE2E2", font: { bold: true, color: "#B91C1C" } } });
  configureColumns(plan, { A: 60, B: 185, C: 105, D: 90, E: 195 });
  plan.freezePanes.freezeRows(3);

  titleRange(risks, "A1:E1", "Project Tracker — Risks and Decisions", 5);
  risks.getRange("A3:E3").values = [["ID", "Type", "Description", "Owner", "Status"]];
  risks.getRange("A3:E3").format = headerFormat();
  risks.getRange("A4:E7").values = [
    ["R-01", "Risk", "Dependency date may slip", "Delivery owner", "Open"],
    ["R-02", "Risk", "Evidence may be incomplete", "Product", "Mitigated"],
    ["D-01", "Decision", "Approve scope boundary", "Decision owner", "Open"],
    ["D-02", "Decision", "Confirm release checkpoint", "Operations", "Closed"],
  ];
  risks.getRange("E4:E7").conditionalFormats.add("containsText", { text: "Open", format: { fill: "#FEF3C7", font: { bold: true, color: "#92400E" } } });
  sectionRange(risks, "A9:E9", "Next actions", 5);
  risks.getRange("A10:B10").values = [["ID", "Next action"]];
  risks.getRange("A10:B10").format = headerFormat();
  risks.getRange("A11:B14").values = [
    ["R-01", "Confirm fallback by Week 2"],
    ["R-02", "Review signal definition"],
    ["D-01", "Decide at operating review"],
    ["D-02", "Record in rollout plan"],
  ];
  configureColumns(risks, { A: 60, B: 130, C: 215, D: 130, E: 90 });
  risks.freezePanes.freezeRows(3);
  workbook.worksheets.setActiveWorksheet("Project Summary");
  workbook.recalculate();
  return workbook;
}
