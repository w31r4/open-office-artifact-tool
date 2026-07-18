import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { inflateSync } from "node:zlib";
import mupdf from "mupdf";
import sharp from "sharp";
import { createPdfjsParser } from "open-office-artifact-tool/pdf/pdfjs";
import { MUPDF_VERSION, createMuPdfParser, parsePdfWithMuPdf, renderPdfWithMuPdf } from "open-office-artifact-tool/pdf/mupdf";
import { FileBlob, PdfArtifact, PdfFile, renderArtifact } from "../src/index.mjs";
import { PdfArtifact as PdfArtifactModule, PdfFile as PdfFileModule } from "../src/pdf/index.mjs";

assert.equal(PdfArtifact, PdfArtifactModule, "the root package must re-export the PDF domain class without wrapping it");
assert.equal(PdfFile, PdfFileModule, "the root package must re-export the PDF file facade without wrapping it");

const lazyMuPdfProbe = spawnSync(process.execPath, ["--input-type=module", "-e", [
  "const loaded=()=>Object.keys(globalThis).some((key)=>key.startsWith('$libmupdf'));",
  "const api=await import('./src/index.mjs');",
  "if(loaded()) throw new Error('root import initialized MuPDF');",
  "for(const method of ['renderPdf','editPdf']){try{await api.PdfFile[method](new Uint8Array(4),{limits:{maxBytes:3}})}catch(error){if(!/exceeds maxBytes/.test(error.message)) throw error}}",
  "if(loaded()) throw new Error('budget rejection initialized MuPDF');",
  "const file=await api.PdfFile.exportPdf(api.PdfArtifact.create({text:'lazy MuPDF probe'}));",
  "await api.PdfFile.inspectPdf(file);",
  "if(!loaded()) throw new Error('first PDF operation did not initialize MuPDF');",
].join("\n")], { cwd: path.resolve(import.meta.dirname, ".."), encoding: "utf8" });
assert.equal(lazyMuPdfProbe.status, 0, `MuPDF lazy-load probe failed\n${lazyMuPdfProbe.stdout}\n${lazyMuPdfProbe.stderr}`);
await assert.rejects(PdfFile.inspectPdf(new Uint8Array(4), { limits: { maxBytes: 3 } }), /exceeds maxBytes/);
await assert.rejects(PdfFile.importPdf(new Uint8Array(4), { limits: { maxBytes: 3 } }), /exceeds maxBytes/);

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

const inlineTable = pdf.addTable({
  id: "inline-table-id",
  name: "inline-table",
  values: [["Evidence", "", "Status"], ["A", "B", ""], ["1", "2", "Pass"]],
  cells: [
    { row: 0, column: 0, columnSpan: 2, role: "TH", scope: "Column" },
    { row: 0, column: 2, rowSpan: 2, role: "TH", scope: "Column" },
    { row: 1, column: 0, role: "TH", scope: "Column" },
    { row: 1, column: 1, role: "TH", scope: "Column" },
    { row: 2, column: 2, headers: ["inline-table-id/cell/1/3"] },
  ],
  bbox: [72, 310, 300, 90],
});
const image = pdf.addImage({
  name: "logo-image",
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  alt: "Report logo",
  bbox: [360, 72, 120, 80],
});
const chart = pdf.addChart({
  name: "pipeline-chart",
  title: "Pipeline by quarter",
  alt: "Bar chart showing pipeline rising from 8 in Q1 to 18 in Q3.",
  chartType: "bar",
  categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "Pipeline", values: [8, 12, 18], color: "#2563eb" }],
  bbox: [72, 420, 360, 150],
});
const positionedText = pdf.addText("Positioned KPI", { bbox: [72, 150, 120, 14], fontSize: 13, color: "#7c3aed", fontName: "Helvetica", headingLevel: 2 });
const firstPage = pdf.pages[0];
const firstPageReadingOrder = [
  `${firstPage.id}/text`,
  image,
  positionedText,
  firstPage.tables[0],
  inlineTable,
  chart,
];
assert.equal(firstPage.setReadingOrder(firstPageReadingOrder), firstPage);
assert.deepEqual(firstPage.readingOrder, firstPageReadingOrder.map((item) => typeof item === "string" ? item : item.id));
const readingOrderInspect = pdf.inspect({ kind: "readingOrder", maxChars: 12000 });
const readingOrderRecords = readingOrderInspect.ndjson.trim().split("\n").map(JSON.parse);
assert.deepEqual(readingOrderRecords.map((record) => record.targetId), firstPage.readingOrder.concat(`${pdf.pages[1].id}/text`));
assert.ok(readingOrderRecords.slice(0, firstPage.readingOrder.length).every((record) => record.explicit && record.valid));
assert.equal(pdf.resolve(readingOrderRecords[0].id).targetId, `${firstPage.id}/text`);
assert.throws(() => firstPage.setReadingOrder("not-an-array"), /readingOrder must be an array/);
firstPage.setReadingOrder(firstPageReadingOrder);
assert.match(pdf.inspect({ kind: "page,text,table,image,chart,textItem", maxChars: 12000 }).ndjson, /Positioned KPI/);
assert.match(pdf.inspect({ kind: "page,text,table,image,chart,textItem", maxChars: 12000 }).ndjson, /pipeline-chart/);
assert.match(pdf.inspect({ kind: "page,text,table,image,chart,textItem", maxChars: 12000 }).ndjson, /metrics-table/);
assert.match(pdf.inspect({ kind: "image", search: "logo" }).ndjson, /Report logo/);
assert.equal(image.alt, "Report logo");
assert.equal(pdf.resolve(pdf.id), pdf);
assert.equal(pdf.resolve(pdf.pages[0].id), pdf.pages[0]);
assert.equal(pdf.resolve(`${pdf.pages[0].id}/text`).text, pdf.pages[0].text);
assert.equal(pdf.resolve(inlineTable.id), inlineTable);
const spanningHeader = inlineTable.getCell(0, 0);
assert.equal(spanningHeader.columnSpan, 2);
assert.equal(inlineTable.getCell(0, 1).id, spanningHeader.id);
assert.equal(inlineTable.getCell(0, 2).rowSpan, 2);
assert.deepEqual(inlineTable.getCell(2, 2).headers, ["inline-table-id/cell/1/3"]);
assert.deepEqual(inlineTable.getCell(2, 0).effectiveHeaders, ["inline-table-id/cell/1/1", "inline-table-id/cell/2/1"]);
assert.equal(pdf.resolve(spanningHeader.id).id, spanningHeader.id);
const editableSpanTable = PdfArtifact.create({ text: "Editable span" }).addTable({ id: "editable-span", values: [["A", "B"], ["1", "2"]] });
editableSpanTable.getCell(0, 0).columnSpan = 2;
assert.equal(editableSpanTable.getCell(0, 1).id, "editable-span/cell/1/1");
assert.equal(editableSpanTable.page.artifact.verify().ok, true);
assert.equal(pdf.resolve(image.id), image);
assert.equal(pdf.resolve(chart.id), chart);
assert.equal(pdf.resolve(positionedText.id).text, "Positioned KPI");
assert.equal(pdf.resolve(positionedText.id).headingLevel, 2);
assert.equal(pdf.resolve("missing/pdf-id"), undefined);
assert.match(pdf.inspect({ kind: "table", search: "Retention" }).ndjson, /94%/);
assert.match(pdf.inspect({ kind: "tableCell", target: spanningHeader.id }).ndjson, /"columnSpan":2/);
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
assert.equal(pdf.extractTables()[1].cells.filter((cell) => cell.role === "TH").length, 4);
assert.match(pdf.help("pdf.addImage").ndjson, /image region/);
assert.match(pdf.help("pdf.addChart").ndjson, /chart region/);
assert.match(pdf.help("pdf.addText").ndjson, /positioned PDF text/);
assert.match(pdf.help("pdf.addFlowText").ndjson, /automatically append pages/);
assert.match(pdf.help("PdfFile.exportPdf").ndjson, /Table\/TR\/TH\/TD hierarchy/);
assert.match(pdf.help("PdfFile.exportPdf").ndjson, /Unicode TrueType embedding/);

const flowPdf = PdfArtifact.create({ pages: [{ width: 240, height: 120, text: "" }] });
const flowText = [
  "Automatic pagination keeps agent-authored reports inside the requested content box while preserving inspectable positioned lines.",
  "AveryLongUnbrokenTokenThatMustBeSplitAcrossSeveralLinesWithoutOverflowingThePageWidth",
  "The final paragraph proves that page creation continues after both ordinary wrapping and a long token.",
].join("\n");
const flow = flowPdf.addFlowText(flowText, { margins: 20, fontSize: 10, lineHeight: 13, paragraphGap: 5 });
assert.ok(flowPdf.pages.length > 1);
assert.equal(flow.items.length, flow.lineCount);
assert.equal(flow.pageIds.length, flowPdf.pages.length);
assert.equal(flowPdf.resolve(flow.items[0].id).flowId, flow.id);
assert.ok(flow.items.every((item) => item.bbox[0] >= 20 && item.bbox[1] >= 20 && item.bbox[0] + item.bbox[2] <= 220 && item.bbox[1] + item.bbox[3] <= 100));
assert.ok(flow.items.some((item) => item.text.startsWith("AveryLong")));
assert.equal(flowPdf.verify().ok, true);
const flowBlob = await PdfFile.exportPdf(flowPdf);
assert.equal((await PdfFile.inspectPdf(flowBlob)).summary.pages, flowPdf.pages.length);
const flowRoundtrip = await PdfFile.importPdf(flowBlob);
assert.equal(flowRoundtrip.pages.length, flowPdf.pages.length);
assert.equal(flowRoundtrip.pages.flatMap((page) => page.textItems).filter((item) => item.flowId === flow.id).length, flow.lineCount);

const preview = await pdf.render({ pageIndex: 0 });
assert.equal(preview.type, "image/svg+xml");
const svg = await preview.text();
assert.match(svg, /PDF research artifact/);
assert.match(svg, /Positioned KPI/);
assert.match(svg, /data-text-item-id=/);
assert.match(svg, /#7c3aed/);
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
assert.ok(layout.pages[0].textItems.some((item) => item.text === "Positioned KPI" && item.color === "#7c3aed" && item.headingLevel === 2));
assert.ok(layout.pages[0].images.some((item) => item.alt === "Report logo"));
assert.ok(layout.pages[0].charts.some((item) => item.title === "Pipeline by quarter"));
assert.deepEqual(layout.pages[0].readingOrder.map((item) => item.targetId), firstPage.readingOrder);
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
assert.equal(blob.metadata.tagged, true);
assert.equal(blob.metadata.language, "en-US");
assert.equal(blob.metadata.title, "PDF research artifact");

// Arbitrary imported PDFs use the required MuPDF.js dependency by default,
// while the package root keeps loading it lazily until a PDF operation asks.
const ownedPdfText = Buffer.from(blob.bytes).toString("latin1");
const arbitraryPdfBytes = Buffer.from(ownedPdfText.replace(/%OPEN_OFFICE_ARTIFACT [A-Za-z0-9+/=]+/, (match) => `%${"X".repeat(match.length - 1)}`), "latin1");
const arbitraryPdf = new FileBlob(arbitraryPdfBytes, { type: "application/pdf" });
const mupdfParsed = await PdfFile.importPdf(arbitraryPdf);
assert.equal(mupdfParsed.metadata.parser, "mupdf");
assert.equal(mupdfParsed.metadata.provider, "mupdf");
assert.equal(mupdfParsed.metadata.providerVersion, MUPDF_VERSION);
assert.equal(mupdfParsed.pages.length, 2);
assert.match(mupdfParsed.extractText(), /PDF research artifact/);
assert.ok(mupdfParsed.pages[0].textItems.some((item) => item.text.includes("PDF research artifact")));
assert.ok(mupdfParsed.pages[0].images.some((item) => item.dataUrl?.startsWith("data:image\/png;base64,")));
assert.equal(mupdfParsed.pages[0].images[0].transform.length, 6);
const mupdfRawWithoutImages = await parsePdfWithMuPdf({ bytes: arbitraryPdf.bytes, options: { includeImages: false } });
assert.equal(mupdfRawWithoutImages.pages[0].rotation, 0);
assert.equal(mupdfRawWithoutImages.pages[0].images.length, 0);
await assert.rejects(parsePdfWithMuPdf(arbitraryPdf.bytes, { limits: { maxPages: Number.NaN } }), /limit maxPages must be a positive finite number/);
const mupdfInspect = await PdfFile.inspectPdf(arbitraryPdf, { maxChars: 20_000 });
assert.equal(mupdfInspect.summary.nativeProvider, "mupdf");
assert.equal(mupdfInspect.summary.nativeProviderVersion, MUPDF_VERSION);
assert.equal(mupdfInspect.summary.pages, 2);
assert.equal(mupdfInspect.records.filter((record) => record.kind === "mupdfPage").length, 2);
const mupdfPng = await PdfFile.renderPdf(arbitraryPdf, { page: 1, dpi: 72 });
assert.equal(mupdfPng.type, "image/png");
assert.deepEqual([...mupdfPng.bytes.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
assert.equal(mupdfPng.metadata.provider, "mupdf");
await assert.rejects(renderPdfWithMuPdf(arbitraryPdf.bytes, { dpi: 72, limits: { maxRenderPixels: 1 } }), /exceeds maxRenderPixels/);
const mupdfIncremental = await PdfFile.editPdf(arbitraryPdf, {
  savePolicy: "incremental",
  operations: [{ type: "add_text_annotation", page: 1, bbox: [40, 40, 24, 24], text: "Agent review" }],
});
assert.equal(mupdfIncremental.metadata.savePolicy, "incremental");
assert.equal(Buffer.from(mupdfIncremental.bytes.subarray(0, arbitraryPdf.bytes.length)).equals(Buffer.from(arbitraryPdf.bytes)), true);
assert.equal((await PdfFile.inspectPdf(mupdfIncremental)).records.find((record) => record.kind === "mupdfPage" && record.page === 1).annotations, 1);
const mupdfCropped = await PdfFile.editPdf(arbitraryPdf, {
  savePolicy: "incremental",
  operations: [{ type: "set_page_crop", page: 1, bbox: [72, 72, 468, 648] }],
});
assert.equal(mupdfCropped.metadata.savePolicy, "incremental");
assert.equal(Buffer.from(mupdfCropped.bytes.subarray(0, arbitraryPdf.bytes.length)).equals(Buffer.from(arbitraryPdf.bytes)), true);
assert.deepEqual(mupdfCropped.metadata.operations[0], {
  type: "set_page_crop",
  page: 1,
  bbox: [72, 72, 468, 648],
  mediaBox: [0, 0, 612, 792],
  contentRemoved: false,
});
const croppedRecord = (await PdfFile.inspectPdf(mupdfCropped)).records.find((record) => record.kind === "mupdfPage" && record.page === 1);
assert.deepEqual(croppedRecord.mediaBox, [0, 0, 612, 792]);
assert.deepEqual(croppedRecord.cropBox, [72, 72, 468, 648]);
assert.deepEqual(croppedRecord.bbox, [0, 0, 468, 648]);
const croppedRender = await PdfFile.renderPdf(mupdfCropped, { page: 1, dpi: 72 });
assert.equal(croppedRender.metadata.width, 468);
assert.equal(croppedRender.metadata.height, 648);
const mupdfRotated = await PdfFile.editPdf(arbitraryPdf, {
  savePolicy: "incremental",
  operations: [{ type: "rotate_page", page: 1, rotation: 90 }],
});
assert.equal(Buffer.from(mupdfRotated.bytes.subarray(0, arbitraryPdf.bytes.length)).equals(Buffer.from(arbitraryPdf.bytes)), true);
assert.deepEqual(mupdfRotated.metadata.operations[0], {
  type: "rotate_page",
  page: 1,
  rotation: 90,
  previousRotation: 0,
  contentRemoved: false,
});
const rotatedRecord = (await PdfFile.inspectPdf(mupdfRotated)).records.find((record) => record.kind === "mupdfPage" && record.page === 1);
assert.equal(rotatedRecord.rotation, 90);
assert.deepEqual(rotatedRecord.mediaBox, [0, 0, 612, 792]);
assert.deepEqual(rotatedRecord.cropBox, [0, 0, 612, 792]);
const rotatedRender = await PdfFile.renderPdf(mupdfRotated, { page: 1, dpi: 72 });
assert.equal(rotatedRender.metadata.width, 792);
assert.equal(rotatedRender.metadata.height, 612);
const mupdfRotationRestored = await PdfFile.editPdf(mupdfRotated, {
  savePolicy: "incremental",
  operations: [{ type: "rotate_page", page: 1, rotation: 0 }],
});
assert.equal(Buffer.from(mupdfRotationRestored.bytes.subarray(0, mupdfRotated.bytes.length)).equals(Buffer.from(mupdfRotated.bytes)), true);
assert.equal((await PdfFile.inspectPdf(mupdfRotationRestored)).records.find((record) => record.kind === "mupdfPage" && record.page === 1).rotation, 0);
await assert.rejects(PdfFile.editPdf(arbitraryPdf, {
  operations: [{ type: "rotate_page", page: 1, rotation: 45 }],
}), /rotate_page rotation must be 0, 90, 180, or 270/);
const mupdfCropRestored = await PdfFile.editPdf(mupdfCropped, {
  savePolicy: "incremental",
  operations: [{ type: "set_page_crop", page: 1, bbox: [0, 0, 612, 792] }],
});
assert.equal(Buffer.from(mupdfCropRestored.bytes.subarray(0, mupdfCropped.bytes.length)).equals(Buffer.from(mupdfCropped.bytes)), true);
const restoredRecord = (await PdfFile.inspectPdf(mupdfCropRestored)).records.find((record) => record.kind === "mupdfPage" && record.page === 1);
assert.deepEqual(restoredRecord.cropBox, [0, 0, 612, 792]);
await assert.rejects(PdfFile.editPdf(arbitraryPdf, {
  operations: [{ type: "set_page_crop", page: 1, bbox: [-1, 0, 613, 792] }],
}), /crop rectangle must fit fully inside the page MediaBox/);
const offsetMediaNativeDocument = new mupdf.PDFDocument(arbitraryPdf.bytes);
const offsetMediaNativePage = offsetMediaNativeDocument.loadPage(0);
offsetMediaNativePage.setPageBox("MediaBox", [10, 20, 622, 812]);
offsetMediaNativePage.update();
const offsetMediaNativeOutput = offsetMediaNativeDocument.saveToBuffer("garbage=2,compress=yes");
const offsetMediaPdf = new FileBlob(new Uint8Array(offsetMediaNativeOutput.asUint8Array()), { type: "application/pdf" });
offsetMediaNativeOutput.destroy();
offsetMediaNativePage.destroy();
offsetMediaNativeDocument.destroy();
const offsetMediaRecord = (await PdfFile.inspectPdf(offsetMediaPdf)).records.find((record) => record.kind === "mupdfPage" && record.page === 1);
const offsetMediaCrop = [
  offsetMediaRecord.mediaBox[0] + 72,
  offsetMediaRecord.mediaBox[1] + 72,
  offsetMediaRecord.mediaBox[2] - 144,
  offsetMediaRecord.mediaBox[3] - 144,
];
const offsetMediaCropped = await PdfFile.editPdf(offsetMediaPdf, {
  operations: [{ type: "set_page_crop", page: 1, bbox: offsetMediaCrop }],
});
assert.deepEqual((await PdfFile.inspectPdf(offsetMediaCropped)).records.find((record) => record.kind === "mupdfPage" && record.page === 1).cropBox, offsetMediaCrop);
const rotatedNativeDocument = new mupdf.PDFDocument(arbitraryPdf.bytes);
const rotatedNativePage = rotatedNativeDocument.loadPage(0);
const rotatedNativeObject = rotatedNativePage.getObject();
rotatedNativeObject.put("Rotate", 90);
rotatedNativePage.update();
const rotatedNativeOutput = rotatedNativeDocument.saveToBuffer("garbage=2,compress=yes");
const rotatedPdf = new FileBlob(new Uint8Array(rotatedNativeOutput.asUint8Array()), { type: "application/pdf" });
rotatedNativeOutput.destroy();
rotatedNativeObject.destroy();
rotatedNativePage.destroy();
rotatedNativeDocument.destroy();
await assert.rejects(PdfFile.editPdf(rotatedPdf, {
  operations: [{ type: "set_page_crop", page: 1, bbox: [72, 72, 468, 648] }],
}), /set_page_crop supports only unrotated pages/);
const mupdfRedacted = await PdfFile.editPdf(arbitraryPdf, {
  savePolicy: "rewrite",
  operations: [{ type: "redact_text", page: 2, term: "Second page notes" }],
});
assert.doesNotMatch((await PdfFile.importPdf(mupdfRedacted)).extractText(), /Second page notes/);
assert.doesNotMatch(Buffer.from(mupdfRedacted.bytes).toString("latin1"), /Second page notes/);
await assert.rejects(PdfFile.editPdf(arbitraryPdf, { savePolicy: "incremental", operations: [{ type: "redact_text", page: 2, term: "Second page notes" }] }), /redaction cannot save incrementally.*prior revisions/is);
await assert.rejects(PdfFile.editPdf(arbitraryPdf, { savePolicy: "incremental", operations: [{ type: "delete_page", page: 2 }] }), /destructive operation delete_page cannot save incrementally.*prior revisions/is);
await assert.rejects(PdfFile.editPdf(arbitraryPdf, { operations: [{ type: "replace_text", page: 1, term: "PDF", replacement: "Document" }] }), /Unsupported MuPDF edit operation/);
const signatureLiteralPdf = await PdfFile.exportPdf(PdfArtifact.create({ text: "Literal /ByteRange [ text is not a signature" }));
const signatureLiteralEdited = await PdfFile.editPdf(signatureLiteralPdf, { operations: [{ type: "add_text_annotation", page: 1, text: "Not signed" }] });
assert.equal(signatureLiteralEdited.metadata.signedInput, false);
assert.equal(typeof createMuPdfParser(), "function");
const fileInspect = await PdfFile.inspectPdf(blob, { maxChars: 12000 });
assert.equal(fileInspect.summary.version, "1.7");
assert.equal(fileInspect.summary.pages, 2);
assert.equal(fileInspect.summary.hasEmbeddedModel, true);
assert.equal(fileInspect.summary.hasEof, true);
assert.equal(fileInspect.summary.tagged, true);
assert.equal(fileInspect.summary.language, "en-US");
assert.equal(fileInspect.summary.tableStructures, 2);
assert.equal(fileInspect.summary.tableRows, 6);
assert.equal(fileInspect.summary.tableHeaders, 6);
assert.equal(fileInspect.summary.tableDataCells, 7);
assert.equal(fileInspect.summary.tableCellIds, 13);
assert.equal(fileInspect.summary.rowSpans, 1);
assert.equal(fileInspect.summary.columnSpans, 1);
assert.equal(fileInspect.summary.headerAssociations, 7);
assert.equal(fileInspect.summary.structureRoles.Table, 2);
assert.equal(fileInspect.summary.structureRoles.TR, 6);
assert.equal(fileInspect.summary.structureRoles.TH, 6);
assert.equal(fileInspect.summary.structureRoles.TD, 7);
assert.equal(fileInspect.summary.headingLevels.H1, 2);
assert.equal(fileInspect.summary.headingLevels.H2, 1);
assert.equal(fileInspect.summary.headings, 3);
assert.equal(fileInspect.summary.figures, 2);
assert.equal(fileInspect.summary.figureAltTexts, 2);
assert.equal(fileInspect.summary.missingFigureAltTexts, 0);
assert.equal(fileInspect.summary.artifacts, 0);
assert.deepEqual(fileInspect.summary.readingOrderIds, firstPage.readingOrder.concat(`${pdf.pages[1].id}/text`));
assert.equal(fileInspect.summary.readingOrderItems, firstPage.readingOrder.length + 1);
assert.equal(fileInspect.summary.structureElements, fileInspect.summary.markedContentItems + fileInspect.summary.tableStructures + fileInspect.summary.tableRows);
assert.match(fileInspect.ndjson, /"type":"Page"/);
const taggedText = await blob.text();
assert.match(taggedText, /\/Type \/StructTreeRoot/);
assert.match(taggedText, /\/S \/Table/);
assert.match(taggedText, /\/S \/TR/);
assert.match(taggedText, /\/S \/TH/);
assert.match(taggedText, /\/S \/TD/);
assert.match(taggedText, /\/S \/H2/);
assert.match(taggedText, /\/A << \/O \/Table \/Scope \/Column >>/);
assert.match(taggedText, /\/ColSpan 2/);
assert.match(taggedText, /\/RowSpan 2/);
assert.match(taggedText, /\/ID \(inline-table-id\/cell\/1\/3\)/);
assert.match(taggedText, /\/Headers \[\(inline-table-id\/cell\/1\/3\)\]/);
assert.match(taggedText, /\/Alt \(Report logo\)/);
assert.match(taggedText, /\/Alt \(Bar chart showing pipeline rising from 8 in Q1 to 18 in Q3\.\)/);
const taggedContentStream = [...taggedText.matchAll(/stream\n([\s\S]*?)endstream/g)].map((match) => match[1]).find((content) => /\/H1 << \/MCID/.test(content));
assert.ok(taggedContentStream.indexOf("/H1") < taggedContentStream.indexOf(" Do"), "explicit logical order must not change visual paint order");

const readingOrderRoundtrip = await PdfFile.importPdf(blob);
assert.deepEqual(readingOrderRoundtrip.pages[0].readingOrder, firstPage.readingOrder);
assert.equal(readingOrderRoundtrip.pages[0].textItems.find((item) => item.id === positionedText.id).headingLevel, 2);
assert.deepEqual(readingOrderRoundtrip.layoutJson({ page: 1 }).pages[0].readingOrder.map((item) => item.targetId), firstPage.readingOrder);

assert.throws(() => PdfArtifact.create({ pages: [{ textItems: [{ text: "Bad heading", headingLevel: 0 }] }] }), /headingLevel must be an integer from 1 through 6/);
assert.throws(() => PdfArtifact.create({ pages: [{ textItems: [{ text: "Bad heading", headingLevel: 7 }] }] }), /headingLevel must be an integer from 1 through 6/);
assert.throws(() => PdfArtifact.create({ pages: [{ textItems: [{ text: "Bad heading", headingLevel: 1.5 }] }] }), /headingLevel must be an integer from 1 through 6/);
const skippedHeadingPdf = PdfArtifact.create({ pages: [{ id: "heading-skip-page", text: "Report title", textItems: [{ id: "heading-skip-h3", text: "Skipped subsection", headingLevel: 3 }], readingOrder: ["heading-skip-page/text", "heading-skip-h3"] }] });
assert.match(skippedHeadingPdf.verify().ndjson, /headingLevelSkipped/);
const crossPageHeadingPdf = PdfArtifact.create({ pages: [
  { id: "heading-page-1", text: "Document title", readingOrder: ["heading-page-1/text"] },
  { id: "heading-page-2", text: "Continued section", textItems: [{ id: "heading-page-2-h2", text: "Continued section", headingLevel: 2 }], readingOrder: ["heading-page-2-h2"] },
] });
assert.equal(crossPageHeadingPdf.verify().ok, true);

const invalidReadingOrder = PdfArtifact.create({
  pages: [{
    id: "invalid-reading-page",
    text: "Invalid order",
    tables: [{ id: "missing-reading-table", values: [["A"]] }],
    readingOrder: ["invalid-reading-page/text", "invalid-reading-page/text", "unknown-target"],
  }],
});
const invalidReadingReport = invalidReadingOrder.verify();
assert.equal(invalidReadingReport.ok, false);
assert.match(invalidReadingReport.ndjson, /readingOrderDuplicate/);
assert.match(invalidReadingReport.ndjson, /readingOrderUnknown/);
assert.match(invalidReadingReport.ndjson, /readingOrderMissing/);
await assert.rejects(() => PdfFile.exportPdf(invalidReadingOrder), /Invalid PDF reading order.*readingOrderDuplicate/);

const unicodeReadingOrderId = PdfArtifact.create({ pages: [{ id: "页面(1)", text: "ASCII body", readingOrder: ["页面(1)/text"] }] });
const unicodeReadingOrderInspect = await PdfFile.inspectPdf(await PdfFile.exportPdf(unicodeReadingOrderId));
assert.deepEqual(unicodeReadingOrderInspect.summary.readingOrderIds, ["页面(1)/text"]);

const decorativePdf = PdfArtifact.create({
  pages: [{
    id: "decorative-page",
    text: "Decorative content stays visual",
    images: [{ id: "decorative-accent", decorative: true, dataUrl: image.dataUrl, bbox: [500, 40, 20, 20] }],
    readingOrder: ["decorative-page/text"],
  }],
});
assert.equal(decorativePdf.verify().ok, true);
assert.deepEqual(decorativePdf.pages[0].readingOrderRecords(0).map((record) => record.targetId), ["decorative-page/text"]);
const decorativeBlob = await PdfFile.exportPdf(decorativePdf);
const decorativeInspect = await PdfFile.inspectPdf(decorativeBlob);
assert.equal(decorativeInspect.summary.figures, 0);
assert.equal(decorativeInspect.summary.figureAltTexts, 0);
assert.equal(decorativeInspect.summary.artifacts, 1);
assert.match(await decorativeBlob.text(), /\/Artifact BMC/);
assert.equal((await PdfFile.importPdf(decorativeBlob)).pages[0].images[0].decorative, true);

const accessibleLinkPdf = PdfArtifact.create({
  metadata: { title: "Accessible links", language: "en-US" },
  pages: [{ id: "accessible-link-page", text: "Accessible link report" }],
});
const headerArtifact = accessibleLinkPdf.addText("Board accessibility report", { id: "running-header", bbox: [72, 28, 300, 12], fontSize: 9, artifact: true });
const footerArtifact = accessibleLinkPdf.addText("Page 1 of 1", { id: "running-footer", bbox: [72, 752, 120, 12], fontSize: 9, artifact: true });
const guidanceLink = accessibleLinkPdf.addLink({ id: "accessibility-guidance", text: "W3C PDF accessibility guidance", url: "https://www.w3.org/WAI/", bbox: [72, 150, 240, 18] });
accessibleLinkPdf.pages[0].setReadingOrder(["accessible-link-page/text", guidanceLink]);
assert.equal(accessibleLinkPdf.verify().ok, true);
assert.equal(accessibleLinkPdf.resolve(guidanceLink.id), guidanceLink);
assert.equal(accessibleLinkPdf.resolve(headerArtifact.id), headerArtifact);
assert.equal(accessibleLinkPdf.resolve(footerArtifact.id), footerArtifact);
assert.deepEqual(accessibleLinkPdf.pages[0].readingOrderRecords(0).map((record) => record.targetId), ["accessible-link-page/text", guidanceLink.id]);
assert.match(accessibleLinkPdf.inspect({ kind: "link" }).ndjson, /W3C PDF accessibility guidance/);
assert.equal(accessibleLinkPdf.layoutJson().pages[0].links[0].url, "https://www.w3.org/WAI/");
const accessibleLinkBlob = await PdfFile.exportPdf(accessibleLinkPdf);
const accessibleLinkInspect = await PdfFile.inspectPdf(accessibleLinkBlob);
assert.equal(accessibleLinkInspect.summary.artifacts, 2);
assert.equal(accessibleLinkInspect.summary.linkAnnotations, 1);
assert.equal(accessibleLinkInspect.summary.uriActions, 1);
assert.equal(accessibleLinkInspect.summary.linkStructParents, 1);
assert.equal(accessibleLinkInspect.summary.structureRoles.Link, 1);
const accessibleLinkText = await accessibleLinkBlob.text();
assert.match(accessibleLinkText, /\/Subtype \/Link/);
assert.match(accessibleLinkText, /\/S \/URI/);
assert.match(accessibleLinkText, /\/StructParent \d+/);
assert.match(accessibleLinkText, /\/S \/Link/);
assert.match(accessibleLinkText, /\/Type \/OBJR/);
assert.equal((await PdfFile.importPdf(accessibleLinkBlob)).pages[0].links[0].url, "https://www.w3.org/WAI/");
const arbitraryLinkBytes = Buffer.from(Buffer.from(accessibleLinkBlob.bytes).toString("latin1").replace(/%OPEN_OFFICE_ARTIFACT [A-Za-z0-9+/=]+/, (match) => `%${"X".repeat(match.length - 1)}`), "latin1");
const mupdfLinkParsed = await PdfFile.importPdf(new FileBlob(arbitraryLinkBytes, { type: "application/pdf" }));
assert.equal(mupdfLinkParsed.metadata.parser, "mupdf");
assert.equal(mupdfLinkParsed.pages[0].links[0].url, "https://www.w3.org/WAI/");

const invalidLinkPdf = PdfArtifact.create({ pages: [{ text: "Unsafe link" }] });
invalidLinkPdf.addLink({ text: "click here", url: "javascript:alert(1)" });
assert.match(invalidLinkPdf.verify().ndjson, /genericLinkText/);
assert.match(invalidLinkPdf.verify().ndjson, /unsafeLinkProtocol/);
const artifactHeadingPdf = PdfArtifact.create({ pages: [{ text: "Artifact heading", textItems: [{ id: "artifact-h2", text: "Not semantic", bbox: [72, 28, 200, 12], headingLevel: 2, artifact: true }] }] });
assert.match(artifactHeadingPdf.verify().ndjson, /artifactHeading/);

const crossPageTablePdf = PdfArtifact.create({
  metadata: { title: "Cross-page table", language: "en-US" },
  pages: [
    {
      id: "cross-table-page-1",
      text: "Risk register",
      tables: [{ id: "risk-register-part-1", semanticId: "risk-register", name: "Risk register", values: [["Risk", "Owner"], ["Supply", "Operations"]], bbox: [72, 620, 468, 72] }],
      readingOrder: ["cross-table-page-1/text", "risk-register-part-1"],
    },
    {
      id: "cross-table-page-2",
      text: "Mitigation details",
      textItems: [{ id: "mitigation-h2", text: "Mitigation details", headingLevel: 2, bbox: [72, 220, 300, 24] }],
      tables: [{ id: "risk-register-part-2", semanticId: "risk-register", name: "Risk register continued", values: [["Risk", "Owner"], ["Security", "Engineering"]], bbox: [72, 72, 468, 72] }],
      readingOrder: ["risk-register-part-2", "mitigation-h2"],
    },
  ],
});
assert.equal(crossPageTablePdf.verify().ok, true);
assert.equal(crossPageTablePdf.layoutJson().pages[0].tables[0].semanticId, "risk-register");
const crossPageTableBlob = await PdfFile.exportPdf(crossPageTablePdf);
const crossPageTableInspect = await PdfFile.inspectPdf(crossPageTableBlob);
assert.equal(crossPageTableInspect.summary.tableStructures, 1);
assert.equal(crossPageTableInspect.summary.tableRows, 4);
assert.equal(crossPageTableInspect.summary.tableHeaders, 4);
assert.equal(crossPageTableInspect.summary.tableDataCells, 4);
assert.deepEqual(crossPageTableInspect.summary.readingOrderIds, ["cross-table-page-1/text", "risk-register", "mitigation-h2"]);
const crossPageTableText = await crossPageTableBlob.text();
assert.equal([...crossPageTableText.matchAll(/\/S \/Table\b/g)].length, 1);
assert.match(crossPageTableText, /\/ID \(risk-register\)/);
assert.equal((await PdfFile.importPdf(crossPageTableBlob)).pages[1].tables[0].semanticId, "risk-register");

const invalidCrossPageColumns = PdfArtifact.create({ pages: [
  { id: "columns-page-1", text: "Columns", tables: [{ id: "columns-part-1", semanticId: "columns-table", values: [["A", "B"]] }], readingOrder: ["columns-page-1/text", "columns-part-1"] },
  { id: "columns-page-2", text: "", tables: [{ id: "columns-part-2", semanticId: "columns-table", values: [["A", "B", "C"]] }], readingOrder: ["columns-part-2"] },
] });
assert.match(invalidCrossPageColumns.verify().ndjson, /crossPageTableColumnMismatch/);
await assert.rejects(() => PdfFile.exportPdf(invalidCrossPageColumns), /Invalid cross-page PDF table.*crossPageTableColumnMismatch/);
const invalidCrossPageOrder = PdfArtifact.create({ pages: [
  { id: "order-page-1", text: "Order", tables: [{ id: "order-part-1", semanticId: "order-table", values: [["A"]] }], textItems: [{ id: "after-table", text: "Interleaved", bbox: [72, 300, 100, 12] }], readingOrder: ["order-page-1/text", "order-part-1", "after-table"] },
  { id: "order-page-2", text: "", tables: [{ id: "order-part-2", semanticId: "order-table", values: [["B"]] }], readingOrder: ["order-part-2"] },
] });
assert.match(invalidCrossPageOrder.verify().ndjson, /crossPageTableInterleaving/);

const missingAltPdf = PdfArtifact.create({ text: "Missing alt" });
missingAltPdf.addImage({ dataUrl: image.dataUrl });
assert.match(missingAltPdf.verify().ndjson, /missingFigureAltText/);
const genericAltPdf = PdfArtifact.create({ text: "Generic alt" });
genericAltPdf.addChart({ title: "Trend", alt: "Chart", categories: ["A"], series: [{ values: [1] }] });
assert.match(genericAltPdf.verify().ndjson, /genericFigureAltText/);

const unicodePdf = PdfArtifact.create({ metadata: { title: "Unicode résumé", language: "el-GR" }, text: "Unicode résumé\nПривет κόσμος café\nA A\u00a0A" });
await assert.rejects(() => PdfFile.exportPdf(unicodePdf), /provide PdfFile\.exportPdf.*font/);
await assert.rejects(() => PdfFile.exportPdf(PdfArtifact.create({ text: "ASCII" }), { font: Uint8Array.from([0x74, 0x74, 0x63, 0x66]) }), /collections \(\.ttc\) are not supported/);
await assert.rejects(() => PdfFile.exportPdf(PdfArtifact.create({ text: "ASCII" }), { font: Uint8Array.from([0, 1, 0, 0]) }), /Truncated TrueType/);
const liberationFontPath = path.resolve("node_modules/pdfjs-dist/standard_fonts/LiberationSans-Regular.ttf");
await assert.rejects(() => PdfFile.exportPdf(unicodePdf, { font: liberationFontPath, maxFontBytes: 1024 }), /exceeds maxFontBytes/);
await assert.rejects(() => PdfFile.exportPdf(PdfArtifact.create({ text: "中文" }), { font: liberationFontPath }), /does not contain U\+4E2D/);
const unicodeBlob = await PdfFile.exportPdf(unicodePdf, { font: liberationFontPath });
assert.equal(unicodeBlob.metadata.embeddedFont, "LiberationSans-Regular");
assert.equal(unicodeBlob.metadata.fontSubset, true);
const unicodeFullFontBlob = await PdfFile.exportPdf(unicodePdf, { font: liberationFontPath, subsetFont: false });
assert.equal(unicodeFullFontBlob.metadata.fontSubset, false);
assert.ok(unicodeBlob.bytes.length < unicodeFullFontBlob.bytes.length * 0.4, `subset PDF should be materially smaller (${unicodeBlob.bytes.length} vs ${unicodeFullFontBlob.bytes.length})`);
const unicodeInspect = await PdfFile.inspectPdf(unicodeBlob);
assert.equal(unicodeInspect.summary.embeddedFonts, 1);
assert.equal(unicodeInspect.summary.subsetFonts, 1);
assert.equal(unicodeInspect.summary.toUnicodeMaps, 1);
assert.equal((await PdfFile.inspectPdf(unicodeFullFontBlob)).summary.subsetFonts, 0);
const unicodeBytesText = await unicodeBlob.text();
assert.match(unicodeBytesText, /\/Subtype \/Type0/);
assert.match(unicodeBytesText, /\/Subtype \/CIDFontType2/);
assert.match(unicodeBytesText, /\/CIDToGIDMap \d+ 0 R/);
assert.match(unicodeBytesText, /beginbfchar/);
assert.match(unicodeBytesText, /<0020>/);
assert.match(unicodeBytesText, /<00A0>/);
const fontFileObject = Number(/\/FontFile2 (\d+) 0 R/.exec(unicodeBytesText)?.[1]);
assert.ok(fontFileObject > 0);
const unicodeBuffer = Buffer.from(unicodeBlob.bytes);
const fontObjectStart = unicodeBuffer.indexOf(Buffer.from(`${fontFileObject} 0 obj\n`, "ascii"));
const fontStreamStartMarker = unicodeBuffer.indexOf(Buffer.from("stream\n", "ascii"), fontObjectStart);
const fontHeader = unicodeBuffer.subarray(fontObjectStart, fontStreamStartMarker).toString("ascii");
const compressedFontLength = Number(/\/Length (\d+)/.exec(fontHeader)?.[1]);
const subsetFontBytes = inflateSync(unicodeBuffer.subarray(fontStreamStartMarker + 7, fontStreamStartMarker + 7 + compressedFontLength));
const paddedSubsetFont = Buffer.alloc(Math.ceil(subsetFontBytes.length / 4) * 4);
subsetFontBytes.copy(paddedSubsetFont);
let subsetChecksum = 0;
for (let offset = 0; offset < paddedSubsetFont.length; offset += 4) subsetChecksum = (subsetChecksum + paddedSubsetFont.readUInt32BE(offset)) >>> 0;
assert.equal(subsetChecksum, 0xb1b0afba);
assert.ok(subsetFontBytes.length < (await fs.stat(liberationFontPath)).size * 0.5);
const unicodeParsed = await PdfFile.importPdf(unicodeBlob, { parser: createPdfjsParser(), preferParser: true, parserName: "pdfjs" });
assert.match(unicodeParsed.extractText(), /Привет κόσμος café/);

const localCjkFont = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf";
if (await fs.access(localCjkFont).then(() => true, () => false)) {
  const cjkPdf = PdfArtifact.create({ metadata: { title: "中文报告", language: "zh-CN" }, text: "中文报告\n数据验证通过" });
  const cjkBlob = await PdfFile.exportPdf(cjkPdf, { font: localCjkFont, maxFontBytes: 32 * 1024 * 1024 });
  const cjkFontBytes = (await fs.stat(localCjkFont)).size;
  assert.ok(cjkBlob.bytes.length < cjkFontBytes * 0.1, `CJK subset PDF should be under 10% of the source font (${cjkBlob.bytes.length} vs ${cjkFontBytes})`);
  const cjkParsed = await PdfFile.importPdf(cjkBlob, { parser: createPdfjsParser(), preferParser: true, parserName: "pdfjs" });
  assert.match(cjkParsed.extractText(), /中文报告/);
  assert.match(cjkParsed.extractText(), /数据验证通过/);
}
const untaggedBlob = await PdfFile.exportPdf(PdfArtifact.create({ text: "Deliberately untagged fixture" }), { tagged: false, title: "Untagged fixture" });
const untaggedInspect = await PdfFile.inspectPdf(untaggedBlob);
assert.equal(untaggedBlob.metadata.tagged, false);
assert.equal(untaggedInspect.summary.tagged, false);
assert.equal(untaggedInspect.summary.structureElements, 0);
assert.doesNotMatch(await untaggedBlob.text(), /\/StructTreeRoot/);
const jpegPdf = PdfArtifact.create({ text: "JPEG export" });
jpegPdf.addImage({
  name: "jpeg-mark",
  alt: "Green JPEG mark",
  dataUrl: "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwgIyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCj/wAARCAACAAIDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAABQf/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCQAHTd/9k=",
  bbox: [72, 100, 160, 80],
  fit: "cover",
});
const jpegBlob = await PdfFile.exportPdf(jpegPdf);
const jpegBytesText = await jpegBlob.text();
assert.match(jpegBytesText, /\/Filter \/DCTDecode/);
assert.match(jpegBytesText, /\/Width 2 \/Height 2/);
assert.match(jpegBytesText, /\/ColorSpace \/DeviceRGB/);
assert.match(jpegBytesText, /72 612 160 80 re W n 160 0 0 160 72 572 cm \/Im1 Do/);
const limitedPdfjs = await PdfFile.importPdf(jpegBlob, { parser: createPdfjsParser({ maxImagePixels: 1 }), preferParser: true });
assert.equal(limitedPdfjs.pages[0].images.length, 1);
assert.equal(limitedPdfjs.pages[0].images[0].dataUrl, undefined);
assert.match(limitedPdfjs.pages[0].images[0].prompt, /maxImagePixels/);
const corruptJpeg = PdfArtifact.create({ text: "Corrupt JPEG" });
corruptJpeg.addImage({ dataUrl: "data:image/jpeg;base64,/9j/2Q==", bbox: [72, 100, 80, 80] });
await assert.rejects(() => PdfFile.exportPdf(corruptJpeg), /Unable to embed JPEG image/);
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
assert.equal(loaded.pages[0].tables[1].getCell(0, 0).columnSpan, 2);
assert.equal(loaded.resolve("inline-table-id/cell/1/3").rowSpan, 2);
assert.deepEqual(loaded.pages[0].tables[1].getCell(2, 2).headers, ["inline-table-id/cell/1/3"]);

const invalidTablePdf = PdfArtifact.create({ text: "Invalid table semantics" });
invalidTablePdf.addTable({
  id: "invalid-table",
  values: [["A", "B"], ["1", "2"]],
  cells: [
    { row: 0, column: 0, columnSpan: 2 },
    { row: 0, column: 1, rowSpan: 2 },
    { row: 1, column: 0, headers: ["missing-header"] },
  ],
});
const invalidTableVerification = invalidTablePdf.verify({ maxChars: 8000 });
assert.equal(invalidTableVerification.ok, false);
assert.match(invalidTableVerification.ndjson, /overlappingSpan/);
assert.match(invalidTableVerification.ndjson, /missingHeader/);

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
assert.match(parsed.help("PdfFile.importPdf").ndjson, /image bytes|raster placements/i);
assert.match(parsed.help("pdf.resolve").ndjson, /stable PDF artifact IDs/);
assert.equal(parsed.verify().ok, true);

const geometryParsed = await PdfFile.importPdf(new FileBlob(new Uint8Array([0x25, 0x50, 0x44, 0x46]), { type: "application/pdf" }), {
  parserName: "geometry-parser",
  parser: async () => ({
    parser: "geometry-parser",
    pages: [{
      width: 320,
      height: 220,
      text: "Metric Value\nRevenue 12\nRetention 94%",
      textItems: [
        { text: "Metric", bbox: [24, 40, 42, 12] },
        { text: "Value", bbox: [150, 40, 38, 12] },
        { text: "Revenue", bbox: [24, 62, 54, 12] },
        { text: "12", bbox: [150, 62, 16, 12] },
        { text: "Retention", bbox: [24, 84, 62, 12] },
        { text: "94%", bbox: [150, 84, 24, 12] },
      ],
    }],
  }),
});
const geometryTables = geometryParsed.extractTables();
assert.equal(geometryTables.length, 1);
assert.deepEqual(geometryTables[0].values, [["Metric", "Value"], ["Revenue", "12"], ["Retention", "94%"]]);
assert.equal(geometryTables[0].name, "geometry-table-1");
const geometryInspect = geometryParsed.inspect({ kind: "table", maxChars: 8000 }).ndjson;
assert.match(geometryInspect, /textGeometry/);
assert.match(geometryParsed.help("PdfFile.importPdf").ndjson, /text geometry|page geometry.*positioned text/s);

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
const mockMaskOps = {
  save: 1,
  restore: 2,
  transform: 3,
  setFillRGBColor: 4,
  paintImageMaskXObject: 5,
  paintSolidColorImageMask: 6,
};
const mockMaskParser = createPdfjsParser({
  pdfjs: {
    OPS: mockMaskOps,
    getDocument: () => ({
      promise: Promise.resolve({
        numPages: 1,
        getPage: async () => ({
          objs: { get: (id) => id === "mask-1" ? { width: 2, height: 1, data: new Uint8Array([0x40]) } : undefined },
          getViewport: () => ({ width: 100, height: 100, convertToViewportRectangle: ([left, bottom, right, top]) => [left, 100 - top, right, 100 - bottom] }),
          getTextContent: async () => ({ items: [] }),
          getOperatorList: async () => ({
            fnArray: [1, 4, 3, 5, 2, 1, 4, 3, 6, 2],
            argsArray: [[], ["#ff0000"], [20, 0, 0, 10, 10, 20], [{ width: 2, height: 1, data: "mask-1" }], [], [], [0, 1, 0], [8, 0, 0, 8, 50, 30], [], []],
          }),
        }),
      }),
      destroy: async () => undefined,
    }),
  },
});
const maskImported = await PdfFile.importPdf(new FileBlob(new Uint8Array([0x25, 0x50, 0x44, 0x46]), { type: "application/pdf" }), { parser: mockMaskParser, preferParser: true });
assert.equal(maskImported.pages[0].images.length, 2);
const [stencilMask, solidMask] = maskImported.pages[0].images;
assert.equal(stencilMask.isMask, true);
assert.equal(stencilMask.fillColor, "#ff0000");
assert.deepEqual(stencilMask.bbox, [10, 70, 20, 10]);
assert.equal(stencilMask.pixelWidth, 2);
assert.equal(stencilMask.pixelHeight, 1);
assert.equal(stencilMask.sourceObject, "mask-1");
const stencilPixels = await sharp(Buffer.from(stencilMask.dataUrl.split(",")[1], "base64")).raw().toBuffer({ resolveWithObject: true });
assert.deepEqual([...stencilPixels.data], [255, 0, 0, 255, 255, 0, 0, 0]);
assert.equal(solidMask.isMask, true);
assert.equal(solidMask.fillColor, "#00ff00");
assert.deepEqual([...((await sharp(Buffer.from(solidMask.dataUrl.split(",")[1], "base64")).raw().toBuffer()))], [0, 255, 0, 255]);
assert.match(maskImported.inspect({ kind: "image", maxChars: 4000 }).ndjson, /"isMask":true/);
assert.equal(maskImported.layoutJson().pages[0].images[0].fillColor, "#ff0000");
try {
  const pdfjsImported = await PdfFile.importPdf(blob, { parser: pdfjsParser, preferParser: true });
  const pdfjsImage = pdfjsImported.pages.flatMap((page) => page.images).find((item) => item.dataUrl);
  assert.ok(pdfjsImage, "Expected PDF.js to extract embedded image pixels");
  assert.match(pdfjsImage.dataUrl, /^data:image\/png;base64,/);
  assert.ok(pdfjsImage.bbox[0] >= 0 && pdfjsImage.bbox[1] >= 0 && pdfjsImage.bbox[2] > 0 && pdfjsImage.bbox[3] > 0);
  assert.equal(pdfjsImported.verify().ok, true);
} catch (error) {
  if (!/pdfjs-dist|PDF\.js parser requires/.test(String(error?.message || error))) throw error;
}
console.log("pdf smoke ok");
