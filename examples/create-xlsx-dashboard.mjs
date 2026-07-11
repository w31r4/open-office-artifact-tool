import os from "node:os";
import path from "node:path";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

const outputDir = process.env.OUTPUT_DIR || path.join(os.tmpdir(), "open-office-artifact-examples");

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Dashboard");
sheet.getRange("A1:D4").values = [
  ["Month", "Revenue", "Target", "Status"],
  ["Jan", 100, 95, "Done"],
  ["Feb", 120, 110, "Done"],
  ["Mar", 130, 125, "In Progress"],
];
sheet.getRange("E2:E4").formulas = [["=B2-C2"], ["=B3-C3"], ["=B4-C4"]];
sheet.tables.add("A1:E4", true, "RevenueTable");
sheet.charts.add("line", sheet.getRange("A1:B4"), { name: "RevenueChart", title: "Revenue trend" }).setPosition("G1", "L12");
sheet.sparklineGroups.add({ type: "line", targetRange: "F2:F2", sourceData: sheet.getRange("B2:B4"), markers: { show: true } });
sheet.getRange("D2:D4").dataValidation = { rule: { type: "list", values: ["Not Started", "In Progress", "Done"] } };
sheet.getRange("E2:E4").conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 0, format: { fill: "green" } });
workbook.comments.setSelf({ displayName: "Analyst" });
workbook.comments.addThread({ cell: sheet.getRange("E4") }, "Variance depends on latest monthly target.").addReply("Checked by agent.");
workbook.recalculate();

const file = await SpreadsheetFile.exportXlsx(workbook);
await file.save(path.join(outputDir, "xlsx-dashboard.xlsx"));
console.log(workbook.inspect({ kind: "formula,table,chart,sparkline,thread", maxChars: 6000 }).ndjson);
console.log(`saved ${path.join(outputDir, "xlsx-dashboard.xlsx")}`);
