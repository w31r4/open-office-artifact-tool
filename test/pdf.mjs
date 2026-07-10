import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { FileBlob, PdfArtifact, PdfFile } from "../src/index.mjs";

const pdf = PdfArtifact.create({
  pages: [
    {
      text: "PDF research artifact\nSection 1 summary",
      tables: [
        {
          name: "metrics-table",
          values: [["Metric", "Value"], ["Revenue", "$12M"], ["Retention", "94%"]],
          bbox: [72, 180, 360, 96],
        },
      ],
    },
    { text: "Second page notes" },
  ],
});

pdf.addTable({ name: "inline-table", values: [["A", "B"], ["1", "2"]], bbox: [72, 310, 240, 60] });
assert.match(pdf.inspect({ kind: "page,text,table", maxChars: 8000 }).ndjson, /metrics-table/);
assert.match(pdf.inspect({ kind: "table", search: "Retention" }).ndjson, /94%/);
assert.match(pdf.extractText(), /Second page notes/);
assert.equal(pdf.extractTables().length, 2);
assert.deepEqual(pdf.extractTables()[0].values[1], ["Revenue", "$12M"]);

const preview = await pdf.render({ pageIndex: 0 });
assert.equal(preview.type, "image/svg+xml");
const svg = await preview.text();
assert.match(svg, /PDF research artifact/);
assert.match(svg, /Revenue/);

const blob = await PdfFile.exportPdf(pdf);
assert.equal(blob.type, "application/pdf");
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.pdf`);
await blob.save(out);
const loaded = await PdfFile.importPdf(await FileBlob.load(out));
const loadedInspect = loaded.inspect({ kind: "page,text,table", maxChars: 8000 }).ndjson;
assert.match(loadedInspect, /PDF research artifact/);
assert.match(loadedInspect, /metrics-table/);
assert.match(loadedInspect, /inline-table/);
assert.deepEqual(loaded.extractTables()[0].values[2], ["Retention", "94%"]);
console.log("pdf smoke ok");
