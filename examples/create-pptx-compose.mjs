import os from "node:os";
import path from "node:path";

import { box, column, paragraph, Presentation, PresentationFile, row } from "open-office-artifact-tool";

const outputDir = process.env.OUTPUT_DIR || path.join(os.tmpdir(), "open-office-artifact-examples");

const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const slide = presentation.slides.add({ name: "Overview" });
const [headline] = slide.compose(
  column({ name: "content-frame", width: "fill", height: "fill", gap: 20, padding: { x: 28, y: 24 } }, [
    paragraph({ name: "headline", className: "text-slate-950 text-4xl font-bold" }, ["Clean-room PPTX compose"]),
    row({ name: "kpis", width: "fill", height: 120, gap: 16 }, [
      box({ name: "kpi-a", fill: "slate-50", padding: { x: 16, y: 12 } }, [paragraph({ name: "kpi-a-label" }, ["Native tables"])]),
      box({ name: "kpi-b", fill: "sky-50", padding: { x: 16, y: 12 } }, [paragraph({ name: "kpi-b-label" }, ["SVG previews"])]),
    ]),
  ]),
  { frame: { left: 80, top: 96, width: 680, height: 260 } },
);
const table = slide.tables.add({ name: "roadmap-table", position: { left: 820, top: 120, width: 320, height: 140 }, values: [["Area", "Status"], ["Notes", "Partial"], ["Connectors", "Partial"]], styleOptions: { headerRow: true } });
slide.charts.add("bar", { name: "coverage-chart", title: "Coverage", position: { left: 820, top: 300, width: 320, height: 200 }, categories: ["DOCX", "XLSX", "PPTX", "PDF"], series: [{ name: "Done", values: [65, 70, 68, 60] }] });
slide.connectors.add({ name: "headline-to-table", from: headline, to: table, line: { fill: "#0284c7", width: 2, endArrow: "triangle" } });
slide.addNotes("Mention that this deck uses clean-room compose/layout APIs.");
slide.comments.addThread(headline, "Check headline tone before sending.").resolve();

const file = await PresentationFile.exportPptx(presentation);
await file.save(path.join(outputDir, "pptx-compose.pptx"));
console.log(presentation.inspect({ kind: "slide,textbox,table,chart,connector,notes,comment", maxChars: 6000 }).ndjson);
console.log(`saved ${path.join(outputDir, "pptx-compose.pptx")}`);
