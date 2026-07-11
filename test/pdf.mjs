import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { createPdfjsParser } from "open-office-artifact-tool/pdf/pdfjs";
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

const inlineTable = pdf.addTable({ name: "inline-table", values: [["A", "B"], ["1", "2"]], bbox: [72, 310, 240, 60] });
const image = pdf.addImage({
  name: "logo-image",
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  alt: "Report logo",
  bbox: [360, 72, 120, 80],
});
assert.match(pdf.inspect({ kind: "page,text,table,image", maxChars: 8000 }).ndjson, /metrics-table/);
assert.match(pdf.inspect({ kind: "image", search: "logo" }).ndjson, /Report logo/);
assert.equal(image.alt, "Report logo");
assert.equal(pdf.resolve(pdf.id), pdf);
assert.equal(pdf.resolve(pdf.pages[0].id), pdf.pages[0]);
assert.equal(pdf.resolve(`${pdf.pages[0].id}/text`).text, pdf.pages[0].text);
assert.equal(pdf.resolve(inlineTable.id), inlineTable);
assert.equal(pdf.resolve(image.id), image);
assert.equal(pdf.resolve("missing/pdf-id"), undefined);
assert.match(pdf.inspect({ kind: "table", search: "Retention" }).ndjson, /94%/);
assert.match(pdf.extractText(), /Second page notes/);
assert.equal(pdf.extractTables().length, 2);
assert.deepEqual(pdf.extractTables()[0].values[1], ["Revenue", "$12M"]);
assert.match(pdf.help("pdf.addImage").ndjson, /image region/);

const preview = await pdf.render({ pageIndex: 0 });
assert.equal(preview.type, "image/svg+xml");
const svg = await preview.text();
assert.match(svg, /PDF research artifact/);
assert.match(svg, /Revenue/);
assert.match(svg, /Report logo/);
assert.match(svg, /<image href="data:image\/png;base64/);
assert.equal(pdf.verify().ok, true);

const blob = await PdfFile.exportPdf(pdf);
assert.equal(blob.type, "application/pdf");
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.pdf`);
await blob.save(out);
const loaded = await PdfFile.importPdf(await FileBlob.load(out));
const loadedInspect = loaded.inspect({ kind: "page,text,table,image", maxChars: 8000 }).ndjson;
assert.match(loadedInspect, /PDF research artifact/);
assert.match(loadedInspect, /metrics-table/);
assert.match(loadedInspect, /inline-table/);
assert.match(loadedInspect, /Report logo/);
assert.deepEqual(loaded.extractTables()[0].values[2], ["Retention", "94%"]);

const parsed = await PdfFile.importPdf(new FileBlob(new Uint8Array([0x25, 0x50, 0x44, 0x46]), { type: "application/pdf" }), {
  parserName: "unit-parser",
  parser: async ({ inputType, bytes }) => {
    assert.equal(inputType, "application/pdf");
    assert.ok(bytes.byteLength > 0);
    return {
      parser: "unit-parser",
      pages: [
        {
          width: 300,
          height: 200,
          text: "Parsed arbitrary PDF\nMetric | Value\nRevenue | 12",
          textItems: [
            { text: "Parsed", bbox: [10, 20, 42, 12] },
            { text: "arbitrary PDF", bbox: [58, 20, 96, 12] },
          ],
          regions: [{ kind: "textLine", label: "Parsed arbitrary PDF", bbox: [10, 20, 144, 12] }],
          tables: [{ name: "parsed-table", values: [["Metric", "Value"], ["Revenue", "12"]], bbox: [10, 48, 180, 44] }],
          images: [{ name: "parsed-image", alt: "Extracted raster", prompt: "image placeholder", bbox: [200, 40, 60, 60] }],
        },
      ],
    };
  },
});
assert.equal(parsed.metadata.parser, "unit-parser");
const parsedInspect = parsed.inspect({ kind: "page,text,textItem,region,table,image", maxChars: 8000 }).ndjson;
assert.match(parsedInspect, /Parsed arbitrary PDF/);
assert.match(parsedInspect, /"kind":"textItem"/);
assert.match(parsedInspect, /"kind":"region"/);
assert.match(parsedInspect, /parsed-table/);
assert.match(parsedInspect, /parsed-image/);
const parsedPage = parsed.pages[0];
assert.equal(parsed.resolve(parsedPage.textItems[0].id).text, "Parsed");
assert.equal(parsed.resolve(parsedPage.regions[0].id).label, "Parsed arbitrary PDF");
assert.equal(parsed.resolve(parsedPage.tables[0].id).name, "parsed-table");
assert.equal(parsed.resolve(parsedPage.images[0].id).alt, "Extracted raster");
assert.match(parsed.help("pdf.resolve").ndjson, /stable PDF artifact IDs/);
assert.equal(parsed.verify().ok, true);
assert.match(parsed.help("createPdfjsParser").ndjson, /PDF\.js parser adapter/);

const pdfjsParser = createPdfjsParser();
assert.equal(typeof pdfjsParser, "function");
try {
  await PdfFile.importPdf(blob, { parser: pdfjsParser, preferParser: true });
} catch (error) {
  if (!/pdfjs-dist|PDF\.js parser requires/.test(String(error?.message || error))) throw error;
}
console.log("pdf smoke ok");
