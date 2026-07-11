import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { createPdfjsParser } from "open-office-artifact-tool/pdf/pdfjs";
import { FileBlob, PdfArtifact, PdfFile, renderArtifact } from "../src/index.mjs";

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
const chart = pdf.addChart({
  name: "pipeline-chart",
  title: "Pipeline by quarter",
  chartType: "bar",
  categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "Pipeline", values: [8, 12, 18], color: "#2563eb" }],
  bbox: [72, 420, 360, 150],
});
assert.match(pdf.inspect({ kind: "page,text,table,image,chart", maxChars: 10000 }).ndjson, /pipeline-chart/);
assert.match(pdf.inspect({ kind: "page,text,table,image,chart", maxChars: 10000 }).ndjson, /metrics-table/);
assert.match(pdf.inspect({ kind: "image", search: "logo" }).ndjson, /Report logo/);
assert.equal(image.alt, "Report logo");
assert.equal(pdf.resolve(pdf.id), pdf);
assert.equal(pdf.resolve(pdf.pages[0].id), pdf.pages[0]);
assert.equal(pdf.resolve(`${pdf.pages[0].id}/text`).text, pdf.pages[0].text);
assert.equal(pdf.resolve(inlineTable.id), inlineTable);
assert.equal(pdf.resolve(image.id), image);
assert.equal(pdf.resolve(chart.id), chart);
assert.equal(pdf.resolve("missing/pdf-id"), undefined);
assert.match(pdf.inspect({ kind: "table", search: "Retention" }).ndjson, /94%/);
const targetedPdfInspect = pdf.inspect({ kind: "table,image", target: image.id, maxChars: 4000 }).ndjson;
assert.match(targetedPdfInspect, /Report logo/);
assert.doesNotMatch(targetedPdfInspect, /inline-table/);
const shapedPdfInspect = pdf.inspect({ kind: "image", target: image.id, include: "alt,bbox", exclude: "hasDataUrl", maxChars: 4000 }).ndjson;
assert.match(shapedPdfInspect, /Report logo/);
assert.match(shapedPdfInspect, /"bbox"/);
assert.doesNotMatch(shapedPdfInspect, /"hasDataUrl"/);
assert.match(pdf.extractText(), /Second page notes/);
const pdfContextInspect = pdf.inspect({ kind: "page,text", target: pdf.pages[0].id, after: 2, maxChars: 4000 }).ndjson;
assert.match(pdfContextInspect, /PDF research artifact/);
assert.match(pdfContextInspect, /Second page notes/);
assert.equal(pdf.extractTables().length, 2);
assert.deepEqual(pdf.extractTables()[0].values[1], ["Revenue", "$12M"]);
assert.match(pdf.help("pdf.addImage").ndjson, /image region/);
assert.match(pdf.help("pdf.addChart").ndjson, /chart region/);

const preview = await pdf.render({ pageIndex: 0 });
assert.equal(preview.type, "image/svg+xml");
const svg = await preview.text();
assert.match(svg, /PDF research artifact/);
assert.match(svg, /Revenue/);
assert.match(svg, /Report logo/);
assert.match(svg, /Pipeline by quarter/);
assert.match(svg, /Q1/);
assert.match(svg, /<image href="data:image\/png;base64/);
const layoutBlob = await pdf.render({ format: "layout", pageIndex: 0 });
assert.equal(layoutBlob.type, "application/vnd.open-office-artifact.layout+json");
const layout = JSON.parse(await layoutBlob.text());
assert.equal(layout.kind, "pdfLayout");
assert.equal(layout.pages.length, 1);
assert.equal(layout.pages[0].kind, "pdfPageLayout");
assert.equal(layout.pages[0].tables.length, 2);
assert.ok(layout.pages[0].images.some((item) => item.alt === "Report logo"));
assert.ok(layout.pages[0].charts.some((item) => item.title === "Pipeline by quarter"));
const imageLayoutBlob = await pdf.render({ format: "layout", target: image.id });
assert.equal(imageLayoutBlob.metadata.target, image.id);
const imageLayout = JSON.parse(await imageLayoutBlob.text());
assert.deepEqual(imageLayout.pages.map((page) => page.page), [1]);
assert.deepEqual(imageLayout.pages[0].images.map((item) => item.id), [image.id]);
assert.equal(imageLayout.pages[0].tables.length, 0);
assert.equal(imageLayout.pages[0].charts.length, 0);
assert.equal("text" in imageLayout.pages[0], false);
const tableSearchLayout = pdf.layoutJson({ search: "Retention" });
assert.deepEqual(tableSearchLayout.pages[0].tables.map((item) => item.name), ["metrics-table"]);
assert.equal(tableSearchLayout.slice.matchedRecords, 1);
const pdfContextLayout = pdf.layoutJson({ target: image.id, before: 1 });
assert.deepEqual(pdfContextLayout.pages[0].tables.map((item) => item.id), [inlineTable.id]);
assert.deepEqual(pdfContextLayout.pages[0].images.map((item) => item.id), [image.id]);
const chartLayout = pdf.layoutJson({ target: chart.id });
assert.deepEqual(chartLayout.pages[0].charts.map((item) => item.id), [chart.id]);
assert.equal(chartLayout.pages[0].tables.length, 0);
const pdfPageLayout = pdf.layoutJson({ target: pdf.pages[1].id });
assert.deepEqual(pdfPageLayout.pages.map((page) => page.id), [pdf.pages[1].id]);
assert.equal(pdfPageLayout.pages[0].text.text, "Second page notes");
assert.match(pdf.help("pdf.layoutJson").ndjson, /target\/search context slicing/);
assert.equal(pdf.verify().ok, true);

const blob = await PdfFile.exportPdf(pdf);
assert.equal(blob.type, "application/pdf");
let pdfRendererSawInput = false;
const renderedPdfPng = await renderArtifact(pdf, {
  format: "png",
  source: "pdf",
  renderer: async ({ input, inputType, outputType, artifactKind }) => {
    pdfRendererSawInput = true;
    assert.equal(artifactKind, "pdf");
    assert.equal(inputType, "application/pdf");
    assert.equal(outputType, "image/png");
    assert.match(await input.text(), /^%PDF/);
    return new FileBlob(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a]), { type: outputType, metadata: { renderer: "mock-poppler" } });
  },
});
assert.equal(pdfRendererSawInput, true);
assert.equal(renderedPdfPng.type, "image/png");
assert.equal(renderedPdfPng.metadata.renderSource, "pdf");
const pdfSourceBlob = await pdf.render({ format: "pdf", source: "pdf" });
assert.equal(pdfSourceBlob.type, "application/pdf");
assert.match(await pdfSourceBlob.text(), /^%PDF/);
assert.match(pdf.help("pdf.render").ndjson, /source: 'pdf'/);
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.pdf`);
await blob.save(out);
const loaded = await PdfFile.importPdf(await FileBlob.load(out));
const loadedInspect = loaded.inspect({ kind: "page,text,table,image,chart", maxChars: 10000 }).ndjson;
assert.match(loadedInspect, /PDF research artifact/);
assert.match(loadedInspect, /metrics-table/);
assert.match(loadedInspect, /inline-table/);
assert.match(loadedInspect, /Report logo/);
assert.match(loadedInspect, /Pipeline by quarter/);
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
          images: [{ name: "parsed-image", alt: "Extracted raster", bytes: [0x89, 0x50, 0x4e, 0x47], contentType: "image/png", bbox: [200, 40, 60, 60] }],
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
assert.match(parsed.resolve(parsedPage.images[0].id).dataUrl, /^data:image\/png;base64,/);
assert.equal(parsed.resolve(parsedPage.images[0].id).prompt, undefined);
const parsedImageLayout = parsed.layoutJson({ target: parsedPage.images[0].id });
assert.equal(parsedImageLayout.pages[0].images[0].hasDataUrl, true);
assert.match(parsed.help("PdfFile.importPdf").ndjson, /image bytes/);
assert.match(parsed.help("pdf.resolve").ndjson, /stable PDF artifact IDs/);
assert.equal(parsed.verify().ok, true);
const badGeometryPdf = PdfArtifact.create({
  pages: [{
    width: 100,
    height: 100,
    text: "Bad geometry",
    textItems: [{ id: "txt/bad", text: "outside", bbox: [95, 95, 20, 20] }],
    regions: [{ id: "rg/bad", kind: "textLine", label: "outside", bbox: [-1, 0, 20, 20] }],
    images: [{ id: "im/bad-data", dataUrl: "data:not-base64", bbox: [10, 10, 20, 20] }],
    charts: [{ id: "ch/bad", title: "Bad chart", categories: ["A"], series: [{ name: "Bad", values: [Number.NaN] }], bbox: [80, 80, 50, 50] }],
  }],
});
const badGeometryIssues = badGeometryPdf.verify({ maxChars: 4000 }).ndjson;
assert.match(badGeometryIssues, /textItemOutOfBounds/);
assert.match(badGeometryIssues, /regionOutOfBounds/);
assert.match(badGeometryIssues, /invalidImageDataUrl/);
assert.match(badGeometryIssues, /chartNonNumericData/);
assert.match(badGeometryIssues, /chartOutOfBounds/);
assert.match(parsed.help("createPdfjsParser").ndjson, /PDF\.js parser adapter/);

const pdfjsParser = createPdfjsParser();
assert.equal(typeof pdfjsParser, "function");
try {
  await PdfFile.importPdf(blob, { parser: pdfjsParser, preferParser: true });
} catch (error) {
  if (!/pdfjs-dist|PDF\.js parser requires/.test(String(error?.message || error))) throw error;
}
console.log("pdf smoke ok");
