import fs from "node:fs/promises";
import path from "node:path";
import { deflateSync } from "node:zlib";

import { queryHelpRecords } from "../help/index.mjs";
import { decoder, toUint8Array } from "../shared/binary.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { aid } from "../shared/ids.mjs";
import { imageContentTypeFromExtension, imageDataFromDataUrl } from "../shared/images.mjs";
import { filterInspectRecords, inspectRecordMatchesTarget, inspectTargetTokens, ndjson, normalizeKinds, verificationIssue, verificationResult } from "../shared/inspection.mjs";
import { decodePngRgba } from "../shared/png.mjs";
import { fileBlobFromRenderOutput, LAYOUT_MIME, renderTypeForOptions } from "../shared/render-output.mjs";
import { attrEscape, xmlEscape } from "../shared/xml.mjs";
import { inspectPdfFigureAccessibility, normalizePdfFigureAccessibility, normalizePdfHeadingLevel, pdfFigureAccessibilityIssue, pdfHeadingNestingIssues } from "./accessibility.mjs";
import { analyzePdfReadingOrder, inspectPdfReadingOrderIds, normalizePdfReadingOrder, pdfPageBodyTextLines, pdfReadingOrderInspectRecords, resolvePdfReadingOrder } from "./reading-order.mjs";
import { normalizePdfTableGrid, pdfTableCellBBox, serializePdfTableCells } from "./table-grid.mjs";

const PDF_MIME = "application/pdf";

class PdfTableCell {
  constructor(table, row, column) {
    this.table = table;
    this.row = row;
    this.column = column;
  }

  _override() {
    let config = this.table.cellConfigs.find((cell) => Number(cell.row) === this.row && Number(cell.column ?? cell.col) === this.column);
    if (!config) { config = { row: this.row, column: this.column }; this.table.cellConfigs.push(config); }
    return config;
  }

  _record() { return normalizePdfTableGrid(this.table).occupied[this.row]?.[this.column]; }
  _setSpan(property, value) {
    const config = this._override();
    config[property] = value;
    const rowSpan = Math.max(1, Number(property === "rowSpan" ? value : config.rowSpan) || 1);
    const columnSpan = Math.max(1, Number(property === "columnSpan" ? value : config.columnSpan ?? config.colSpan) || 1);
    this.table.cellConfigs = this.table.cellConfigs.filter((cell) => {
      const row = Number(cell.row);
      const column = Number(cell.column ?? cell.col);
      if (row === this.row && column === this.column) return true;
      return row < this.row || row >= this.row + rowSpan || column < this.column || column >= this.column + columnSpan;
    });
  }
  get id() { return this._record()?.id; }
  get value() { return this._record()?.value; }
  set value(value) { while (this.table.values.length <= this.row) this.table.values.push([]); this.table.values[this.row][this.column] = value; this._override().value = value; }
  get rowSpan() { return this._record()?.rowSpan || 1; }
  set rowSpan(value) { this._setSpan("rowSpan", value); }
  get columnSpan() { return this._record()?.columnSpan || 1; }
  set columnSpan(value) { this._setSpan("columnSpan", value); }
  get role() { return this._record()?.role; }
  set role(value) { this._override().role = value; }
  get scope() { return this._record()?.scope; }
  set scope(value) { this._override().scope = value; }
  get headers() { return [...(this._record()?.headers || [])]; }
  set headers(value) { this._override().headers = Array.isArray(value) ? [...value] : []; }
  get effectiveHeaders() { return [...(this._record()?.effectiveHeaders || [])]; }
  get bbox() { const record = this._record(); return record ? pdfTableCellBBox(this.table, record) : undefined; }
  toJSON() { return serializePdfTableCells(this.table).find((cell) => cell.row === this.row && cell.column === this.column); }
}

class PdfTable {
  constructor(page, config = {}) {
    this.page = page;
    this.id = config.id || aid("ptb");
    this.name = config.name || "";
    this.values = (config.values || [[]]).map((row) => [...row]);
    this.cellConfigs = (config.cells || config.cellConfigs || []).map((cell) => ({ ...cell, headers: Array.isArray(cell.headers) ? [...cell.headers] : undefined }));
    for (const cell of this.cellConfigs) {
      const row = Number(cell.row);
      const column = Number(cell.column ?? cell.col);
      if (Number.isInteger(row) && Number.isInteger(column) && row >= 0 && column >= 0 && cell.value !== undefined) {
        while (this.values.length <= row) this.values.push([]);
        this.values[row][column] = cell.value;
      }
    }
    this.bbox = config.bbox || [72, 140, 468, Math.max(24, this.values.length * 24)];
    this.source = config.source;
  }

  grid() { return normalizePdfTableGrid(this); }
  getCell(row, column) {
    const owner = this.grid().occupied[Number(row)]?.[Number(column)];
    return owner ? new PdfTableCell(this, owner.row, owner.column) : undefined;
  }
  cellInspectRecords(pageIndex) { return serializePdfTableCells(this).map((cell) => ({ kind: "tableCell", page: pageIndex + 1, tableId: this.id, ...cell })); }
  inspectRecord(pageIndex) { const grid = this.grid(); return { kind: "table", id: this.id, page: pageIndex + 1, name: this.name || undefined, rows: grid.rows, cols: grid.columns, cells: grid.cells.length, bbox: this.bbox, values: this.values, source: this.source || undefined }; }
  toJSON() {
    const cells = this.grid().cells.map(({ id, row, column, rowSpan, columnSpan, role, scope, headers, value }) => ({ id, row, column, rowSpan, columnSpan, role, scope, headers, value }));
    return { id: this.id, name: this.name, values: this.values, cells, bbox: this.bbox, source: this.source };
  }
}

class PdfImage {
  constructor(page, config = {}) {
    this.page = page;
    this.id = config.id || aid("pim");
    this.name = config.name || "";
    this.dataUrl = config.dataUrl;
    this.uri = config.uri;
    this.prompt = config.prompt;
    const accessibility = normalizePdfFigureAccessibility(config);
    this.alt = accessibility.alt;
    this.decorative = accessibility.decorative;
    this.bbox = config.bbox || [72, 280, Number(config.width || 180), Number(config.height || 120)];
    this.fit = config.fit || "contain";
    this.isMask = Boolean(config.isMask);
    this.fillColor = config.fillColor;
    this.pixelWidth = config.pixelWidth;
    this.pixelHeight = config.pixelHeight;
    this.sourceObject = config.sourceObject;
    this.sourceOperator = config.sourceOperator;
  }

  inspectRecord(pageIndex) { return { kind: "image", id: this.id, page: pageIndex + 1, name: this.name || undefined, alt: this.alt, decorative: this.decorative, uri: this.uri, prompt: this.prompt, bbox: this.bbox, fit: this.fit, hasDataUrl: Boolean(this.dataUrl), isMask: this.isMask || undefined, fillColor: this.fillColor, pixelWidth: this.pixelWidth, pixelHeight: this.pixelHeight, sourceObject: this.sourceObject, sourceOperator: this.sourceOperator }; }
  toJSON() { return { id: this.id, name: this.name, dataUrl: this.dataUrl, uri: this.uri, prompt: this.prompt, alt: this.alt, decorative: this.decorative, bbox: this.bbox, fit: this.fit, isMask: this.isMask || undefined, fillColor: this.fillColor, pixelWidth: this.pixelWidth, pixelHeight: this.pixelHeight, sourceObject: this.sourceObject, sourceOperator: this.sourceOperator }; }
}

class PdfChart {
  constructor(page, config = {}) {
    this.page = page;
    this.id = config.id || aid("pch");
    this.name = config.name || "";
    this.title = config.title || config.name || "Chart";
    const accessibility = normalizePdfFigureAccessibility(config);
    this.alt = accessibility.alt;
    this.decorative = accessibility.decorative;
    this.chartType = config.chartType || config.type || "bar";
    this.categories = (config.categories || config.labels || []).map((item) => String(item ?? ""));
    const sourceSeries = config.series || [{ name: config.seriesName || "Series 1", values: config.values || config.data || [] }];
    this.series = sourceSeries.map((series, index) => ({
      name: series.name || `Series ${index + 1}`,
      values: (series.values || series.data || []).map((value) => Number(value)),
      color: series.color || ["#2563eb", "#16a34a", "#f97316", "#9333ea"][index % 4],
    }));
    this.bbox = config.bbox || [72, 430, 468, 180];
  }

  inspectRecord(pageIndex) { return { kind: "chart", id: this.id, page: pageIndex + 1, name: this.name || undefined, title: this.title, alt: this.alt, decorative: this.decorative, chartType: this.chartType, categories: this.categories, series: this.series, bbox: this.bbox }; }
  toJSON() { return { id: this.id, name: this.name, title: this.title, alt: this.alt, decorative: this.decorative, chartType: this.chartType, categories: this.categories, series: this.series, bbox: this.bbox }; }
}

class PdfPage {
  constructor(artifact, config = {}) {
    this.artifact = artifact;
    this.id = config.id || aid("pg");
    this.text = String(config.text || "");
    this.width = config.width || 612;
    this.height = config.height || 792;
    this.tables = (config.tables || []).map((table) => new PdfTable(this, table));
    this.images = (config.images || []).map((image) => new PdfImage(this, image));
    this.charts = (config.charts || []).map((chart) => new PdfChart(this, chart));
    this.textItems = (config.textItems || []).map((item, index) => this.normalizeTextItem(item, index));
    this.regions = (config.regions || []).map((region, index) => ({ id: region.id || `${this.id}/rg/${index + 1}`, kind: region.kind || "region", bbox: region.bbox || [0, 0, this.width, this.height], label: region.label }));
    this.readingOrder = normalizePdfReadingOrder(config.readingOrder);
  }

  normalizeTextItem(item = {}, index = this.textItems?.length || 0) {
    const bbox = pdfTextItemBBox(item);
    return { id: item.id || `${this.id}/txt/${index + 1}`, text: String(item.text ?? item.str ?? ""), bbox, fontName: item.fontName || item.fontFamily, fontSize: item.fontSize || item.size, color: item.color, bold: Boolean(item.bold), italic: Boolean(item.italic), headingLevel: normalizePdfHeadingLevel(item.headingLevel), dir: item.dir, flowId: item.flowId, paragraphIndex: item.paragraphIndex, lineIndex: item.lineIndex };
  }

  addText(textOrConfig = "", config = {}) {
    const item = typeof textOrConfig === "object" ? textOrConfig : { ...config, text: textOrConfig };
    const normalized = this.normalizeTextItem(item, this.textItems.length);
    this.textItems.push(normalized);
    if (!this.text) this.text = normalized.text;
    else if (!this.text.includes(normalized.text)) this.text += `\n${normalized.text}`;
    return normalized;
  }

  addTable(config = {}) { const table = new PdfTable(this, config); this.tables.push(table); return table; }
  addImage(config = {}) { const image = new PdfImage(this, config); this.images.push(image); return image; }
  addChart(config = {}) { const chart = new PdfChart(this, config); this.charts.push(chart); return chart; }
  setReadingOrder(order) { this.readingOrder = normalizePdfReadingOrder(order); return this; }
  inspectRecord(index) { const readingOrder = analyzePdfReadingOrder(this); return { kind: "page", id: this.id, page: index + 1, width: this.width, height: this.height, textPreview: this.text.slice(0, 300), textChars: this.text.length, textItems: this.textItems.length, tables: this.tables.length, images: this.images.length, charts: this.charts.length, regions: this.regions.length, readingOrderItems: readingOrder.declaredIds.length, explicitReadingOrder: readingOrder.explicit, validReadingOrder: readingOrder.valid }; }
  textRecord(index) { return { kind: "text", id: `${this.id}/text`, page: index + 1, text: this.text, textChars: this.text.length, textItems: this.textItems.length }; }
  textItemRecords(index) { return this.textItems.map((item) => ({ kind: "textItem", page: index + 1, ...item })); }
  regionRecords(index) { return this.regions.map((region) => ({ ...region, kind: "region", regionKind: region.kind || "region", page: index + 1 })); }
  readingOrderRecords(index) { return pdfReadingOrderInspectRecords(this, index); }
  toJSON() { return { id: this.id, text: this.text, width: this.width, height: this.height, textItems: this.textItems, regions: this.regions, tables: this.tables.map((table) => table.toJSON()), images: this.images.map((image) => image.toJSON()), charts: this.charts.map((chart) => chart.toJSON()), readingOrder: this.readingOrder }; }
}

function pdfTextWidth(text, fontSize = 12) {
  let units = 0;
  for (const char of String(text || "")) {
    if (char === " ") units += 0.28;
    else if (/[ilI1.,'`!|:;]/.test(char)) units += 0.28;
    else if (/[mwMW@#%&]/.test(char)) units += 0.88;
    else if (char.codePointAt(0) > 0xff) units += 1;
    else if (/[A-Z]/.test(char)) units += 0.64;
    else units += 0.54;
  }
  return units * fontSize;
}

function splitPdfFlowToken(token, maxWidth, fontSize) {
  const chunks = [];
  let current = "";
  for (const char of String(token)) {
    if (current && pdfTextWidth(current + char, fontSize) > maxWidth) { chunks.push(current); current = char; }
    else current += char;
  }
  if (current) chunks.push(current);
  return chunks;
}

function wrapPdfFlowParagraph(text, maxWidth, fontSize) {
  const words = String(text || "").trim().split(/\s+/).filter(Boolean).flatMap((word) => pdfTextWidth(word, fontSize) > maxWidth ? splitPdfFlowToken(word, maxWidth, fontSize) : [word]);
  if (!words.length) return [];
  const lines = [];
  let current = "";
  for (const word of words) {
    const candidate = current ? `${current} ${word}` : word;
    if (current && pdfTextWidth(candidate, fontSize) > maxWidth) { lines.push(current); current = word; }
    else current = candidate;
  }
  if (current) lines.push(current);
  return lines;
}

function pdfMargins(value) {
  if (Number.isFinite(Number(value))) return { top: Number(value), right: Number(value), bottom: Number(value), left: Number(value) };
  const numberOr = (candidate, fallback) => Number.isFinite(Number(candidate)) ? Number(candidate) : fallback;
  return { top: numberOr(value?.top, 72), right: numberOr(value?.right, 72), bottom: numberOr(value?.bottom, 72), left: numberOr(value?.left, 72) };
}

function pdfLayoutRecordsForPage(pageLayout, pageArrayIndex) {
  const pageRecord = { kind: pageLayout.kind, id: pageLayout.id, page: pageLayout.page, width: pageLayout.width, height: pageLayout.height, unit: pageLayout.unit, textChars: pageLayout.text?.textChars || 0, textPreview: String(pageLayout.text?.text || "").slice(0, 300), tables: pageLayout.tables.length, images: pageLayout.images.length, charts: pageLayout.charts.length, regions: pageLayout.regions.length, textItems: pageLayout.textItems.length };
  const entries = [{ pageArrayIndex, collection: "page", record: pageRecord }];
  if (pageLayout.text) entries.push({ pageArrayIndex, collection: "text", record: pageLayout.text });
  pageLayout.textItems.forEach((record, itemIndex) => entries.push({ pageArrayIndex, collection: "textItems", itemIndex, record }));
  pageLayout.regions.forEach((record, itemIndex) => entries.push({ pageArrayIndex, collection: "regions", itemIndex, record }));
  pageLayout.tables.forEach((record, itemIndex) => entries.push({ pageArrayIndex, collection: "tables", itemIndex, record: { ...record, textPreview: record.values.map((row) => row.map((cell) => String(cell ?? "")).join(" ")).join(" ") } }));
  pageLayout.images.forEach((record, itemIndex) => entries.push({ pageArrayIndex, collection: "images", itemIndex, record }));
  pageLayout.charts.forEach((record, itemIndex) => entries.push({ pageArrayIndex, collection: "charts", itemIndex, record }));
  pageLayout.readingOrder.forEach((record, itemIndex) => entries.push({ pageArrayIndex, collection: "readingOrder", itemIndex, record }));
  return entries;
}

function pdfLayoutSlice(layout, options = {}) {
  const targets = inspectTargetTokens(options);
  const search = String(options.search || options.searchTerm || "").trim().toLowerCase();
  if (!targets.length && !search) return layout;
  const before = Math.max(0, Number(options.before ?? options.contextBefore ?? options.context ?? 0) || 0);
  const after = Math.max(0, Number(options.after ?? options.contextAfter ?? options.context ?? 0) || 0);
  const targetsArtifact = targets.some((target) => target === layout.id || target === "pdfLayout");
  if (targetsArtifact && !search) return { ...layout, slice: { targets, before, after, matchedPages: layout.pages.length, returnedPages: layout.pages.length } };
  const entries = layout.pages.flatMap((pageLayout, pageArrayIndex) => pdfLayoutRecordsForPage(pageLayout, pageArrayIndex));
  const matchingEntryIndexes = [];
  entries.forEach((entry, index) => {
    const matchesSearch = !search || JSON.stringify(entry.record).toLowerCase().includes(search);
    const matchesTarget = !targets.length || targetsArtifact || inspectRecordMatchesTarget(entry.record, targets);
    if (matchesSearch && matchesTarget) matchingEntryIndexes.push(index);
  });
  const keepEntries = new Set();
  for (const index of matchingEntryIndexes) {
    for (let i = Math.max(0, index - before); i <= Math.min(entries.length - 1, index + after); i += 1) keepEntries.add(i);
  }
  const keepByPage = new Map();
  const ensurePageKeep = (pageArrayIndex) => {
    if (!keepByPage.has(pageArrayIndex)) keepByPage.set(pageArrayIndex, { full: false, text: false, textItems: new Set(), regions: new Set(), tables: new Set(), images: new Set(), charts: new Set(), readingOrder: new Set() });
    return keepByPage.get(pageArrayIndex);
  };
  for (const entryIndex of keepEntries) {
    const entry = entries[entryIndex];
    const keep = ensurePageKeep(entry.pageArrayIndex);
    if (entry.collection === "page") keep.full = true;
    else if (entry.collection === "text") keep.text = true;
    else keep[entry.collection].add(entry.itemIndex);
  }
  const pages = layout.pages.map((pageLayout, pageArrayIndex) => {
    const keep = keepByPage.get(pageArrayIndex);
    if (!keep) return undefined;
    if (keep.full) return pageLayout;
    return {
      ...pageLayout,
      text: keep.text ? pageLayout.text : undefined,
      textItems: pageLayout.textItems.filter((_, index) => keep.textItems.has(index)),
      regions: pageLayout.regions.filter((_, index) => keep.regions.has(index)),
      tables: pageLayout.tables.filter((_, index) => keep.tables.has(index)),
      images: keep.images ? pageLayout.images.filter((_, index) => keep.images.has(index)) : [],
      charts: keep.charts ? pageLayout.charts.filter((_, index) => keep.charts.has(index)) : [],
      readingOrder: keep.readingOrder ? pageLayout.readingOrder.filter((_, index) => keep.readingOrder.has(index)) : [],
    };
  }).filter(Boolean);
  const matchedPages = new Set(matchingEntryIndexes.map((index) => entries[index]?.pageArrayIndex).filter((index) => index != null));
  return { ...layout, pages, slice: { targets, search: search || undefined, before, after, matchedPages: matchedPages.size, returnedPages: pages.length, matchedRecords: matchingEntryIndexes.length } };
}

function pdfRenderUsesPdfSource(options = {}) {
  const source = String(options.source || options.inputFormat || options.renderSource || "").trim().toLowerCase();
  return source === "pdf" || options.pdf === true || options.usePdf === true || options.poppler === true;
}

async function renderPdfArtifactFromPdf(artifact, options = {}) {
  const pdf = await PdfFile.exportPdf(artifact);
  pdf.metadata = { ...(pdf.metadata || {}), artifactKind: "pdf", format: "pdf", renderSource: "pdf" };
  const desiredType = renderTypeForOptions(options, pdf.type);
  const format = String(options.format || "pdf").trim().toLowerCase();
  if (!format || format === "pdf" || desiredType === pdf.type) return pdf;
  const adapter = options.renderer || options.rasterRenderer || options.renderAdapter;
  if (typeof adapter !== "function") return pdf;
  const converted = await adapter({ input: pdf, source: pdf, inputType: pdf.type, outputType: desiredType, format: options.format, artifactKind: "pdf", options: { ...options, source: "pdf" } });
  const blob = await fileBlobFromRenderOutput(converted, desiredType, { artifactKind: "pdf", format: options.format, renderedFrom: pdf.type, renderSource: "pdf" });
  if (!blob.type || blob.type === "application/octet-stream") blob.type = desiredType;
  return blob;
}

export class PdfArtifact {
  constructor(options = {}) {
    this.id = options.id || aid("pdf");
    this.metadata = options.metadata || {};
    const pages = options.pages || [{ text: options.text || "", tables: options.tables || [], images: options.images || [], charts: options.charts || [] }];
    this.pages = pages.map((page) => new PdfPage(this, page));
  }

  static create(options = {}) { return new PdfArtifact(options); }
  addPage(config = {}) { const page = new PdfPage(this, config); this.pages.push(page); return page; }
  addText(textOrConfig = "", config = {}) { const pageIndex = Number((typeof textOrConfig === "object" ? textOrConfig.pageIndex ?? textOrConfig.page : config.pageIndex ?? config.page) ?? 0); return (this.pages[pageIndex] || this.pages[0] || this.addPage()).addText(textOrConfig, config); }
  addFlowText(text, config = {}) {
    const flowId = config.id || aid("flow");
    let pageIndex = Math.max(0, Math.trunc(Number(config.pageIndex ?? 0) || 0));
    let page = this.pages[pageIndex] || this.addPage({ width: config.pageWidth, height: config.pageHeight });
    const margins = pdfMargins(config.margins ?? config.margin);
    const fontSize = Math.max(1, Number(config.fontSize ?? 12));
    const lineHeight = Math.max(fontSize, Number(config.lineHeight ?? fontSize * 1.35));
    const paragraphGap = Math.max(0, Number(config.paragraphGap ?? lineHeight * 0.45));
    const left = Number(config.left ?? margins.left);
    const top = Number(config.top ?? margins.top);
    const contentWidth = Math.max(1, Number(config.width ?? (page.width - left - margins.right)));
    const bottom = Math.max(0, Number(config.bottom ?? margins.bottom));
    const existingBottom = Math.max(top, ...page.textItems.map((item) => Number(item.bbox?.[1] || 0) + Number(item.bbox?.[3] || 0) + paragraphGap));
    let cursor = existingBottom;
    const items = [];
    const pageIndexes = new Set();
    const newPage = () => {
      page = this.addPage({ width: page.width, height: page.height });
      pageIndex = this.pages.length - 1;
      cursor = margins.top;
    };
    const ensureLineSpace = () => { if (cursor + lineHeight > page.height - bottom) newPage(); };
    const paragraphs = String(text ?? "").split(/\r?\n/);
    let globalLineIndex = 0;
    paragraphs.forEach((paragraph, paragraphIndex) => {
      const lines = wrapPdfFlowParagraph(paragraph, contentWidth, fontSize);
      if (!lines.length) { cursor += paragraphGap; return; }
      lines.forEach((line) => {
        ensureLineSpace();
        const item = page.addText(line, { bbox: [left, cursor, Math.min(contentWidth, Math.max(1, pdfTextWidth(line, fontSize))), lineHeight], fontName: config.fontName || "Helvetica", fontSize, color: config.color, bold: config.bold, italic: config.italic, flowId, paragraphIndex, lineIndex: globalLineIndex++ });
        items.push(item);
        pageIndexes.add(pageIndex);
        cursor += lineHeight;
      });
      cursor += paragraphGap;
    });
    const indexes = [...pageIndexes];
    return { id: flowId, items, pageIds: indexes.map((index) => this.pages[index].id), pageIndexes: indexes, startPageIndex: indexes[0] ?? pageIndex, endPageIndex: indexes.at(-1) ?? pageIndex, lineCount: items.length };
  }
  addTable(config = {}) { return (this.pages[0] || this.addPage()).addTable(config); }
  addImage(config = {}) { const pageIndex = Number(config.pageIndex ?? config.page ?? 0); return (this.pages[pageIndex] || this.pages[0] || this.addPage()).addImage(config); }
  addChart(config = {}) { const pageIndex = Number(config.pageIndex ?? config.page ?? 0); return (this.pages[pageIndex] || this.pages[0] || this.addPage()).addChart(config); }
  extractText(options = {}) { const pages = options.page == null ? this.pages : [this.pages[Number(options.page) - 1]].filter(Boolean); return pages.map((page) => page.text).join("\n\n"); }
  extractTables(options = {}) { const pages = options.page == null ? this.pages : [this.pages[Number(options.page) - 1]].filter(Boolean); return pages.flatMap((page, index) => page.tables.map((table) => ({ page: options.page || index + 1, id: table.id, name: table.name, values: table.values, cells: serializePdfTableCells(table), bbox: table.bbox }))); }
  resolve(id) {
    if (id === this.id) return this;
    for (const [pageIndex, page] of this.pages.entries()) {
      if (id === page.id) return page;
      if (id === `${page.id}/text`) return { kind: "text", id, page: pageIndex + 1, text: page.text, textChars: page.text.length, pageObject: page };
      const readingOrderRecord = page.readingOrderRecords(pageIndex).find((record) => record.id === id);
      if (readingOrderRecord) return readingOrderRecord;
      const textItem = page.textItems.find((item) => item.id === id);
      if (textItem) return textItem;
      const region = page.regions.find((item) => item.id === id);
      if (region) return region;
      const table = page.tables.find((item) => item.id === id);
      if (table) return table;
      for (const candidate of page.tables) {
        const cell = candidate.grid().byId.get(id);
        if (cell) return candidate.getCell(cell.row, cell.column);
      }
      const image = page.images.find((item) => item.id === id);
      if (image) return image;
      const chart = page.charts.find((item) => item.id === id);
      if (chart) return chart;
    }
    return undefined;
  }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["page", "text", "table", "image", "chart", "readingOrder"]);
    const records = [];
    this.pages.forEach((page, index) => {
      if (kinds.has("page")) records.push(page.inspectRecord(index));
      if (kinds.has("text")) records.push(page.textRecord(index));
      if (kinds.has("textItem")) records.push(...page.textItemRecords(index));
      if (kinds.has("region")) records.push(...page.regionRecords(index));
      if (kinds.has("table")) records.push(...page.tables.map((table) => table.inspectRecord(index)));
      if (kinds.has("tableCell")) records.push(...page.tables.flatMap((table) => table.cellInspectRecords(index)));
      if (kinds.has("image")) records.push(...page.images.map((image) => image.inspectRecord(index)));
      if (kinds.has("chart")) records.push(...page.charts.map((chart) => chart.inspectRecord(index)));
      if (kinds.has("readingOrder")) records.push(...page.readingOrderRecords(index));
    });
    return ndjson(filterInspectRecords(records, options), options.maxChars ?? Infinity);
  }

  verify(options = {}) {
    const issues = [];
    const tableCellIds = new Set();
    const headingSequence = [];
    if (this.pages.length === 0) issues.push(verificationIssue("pdf", "noPages", "PDF artifact has no pages."));
    this.pages.forEach((page, pageIndex) => {
      const readingOrderAnalysis = analyzePdfReadingOrder(page);
      for (const error of readingOrderAnalysis.errors) issues.push(verificationIssue("pdf", error.code, `PDF page ${pageIndex + 1} reading order: ${error.message}`, { page: pageIndex + 1, id: error.id }));
      const headingById = new Map(page.textItems.filter((item) => item.headingLevel != null).map((item) => [item.id, item.headingLevel]));
      if (readingOrderAnalysis.valid) headingSequence.push(...resolvePdfReadingOrder(page).flatMap((entry) => entry.kind === "text" ? [{ id: entry.id, page: pageIndex + 1, level: 1 }] : headingById.has(entry.id) ? [{ id: entry.id, page: pageIndex + 1, level: headingById.get(entry.id) }] : []));
      if (!page.text.trim() && page.tables.length === 0 && page.images.length === 0 && page.charts.length === 0) issues.push(verificationIssue("pdf", "emptyPage", `PDF page ${pageIndex + 1} has no modeled text, tables, images, or charts.`, { page: pageIndex + 1 }));
      if (page.textItems.length && !page.text.trim()) issues.push(verificationIssue("pdf", "textExtractionMismatch", `PDF page ${pageIndex + 1} has positioned text items but no extracted page text.`, { severity: "warning", page: pageIndex + 1 }));
      if (page.textItems.some((item) => !item.text)) issues.push(verificationIssue("pdf", "emptyTextItem", `PDF page ${pageIndex + 1} contains an empty positioned text item.`, { severity: "warning", page: pageIndex + 1 }));
      if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(page.text)) issues.push(verificationIssue("pdf", "textExtractionControlChars", `PDF page ${pageIndex + 1} extracted text contains control characters.`, { page: pageIndex + 1 }));
      if (page.width <= 0 || page.height <= 0) issues.push(verificationIssue("pdf", "invalidPageGeometry", `PDF page ${pageIndex + 1} has invalid page geometry.`, { page: pageIndex + 1, width: page.width, height: page.height }));
      if (/[\u2010-\u2015]/.test(page.text)) issues.push(verificationIssue("pdf", "unicodeDash", `PDF page ${pageIndex + 1} contains a Unicode dash; use ASCII hyphen for compatibility.`, { page: pageIndex + 1 }));
      for (const item of page.textItems) {
        const [left, top, width, height] = item.bbox || [];
        if (!Number.isFinite(left) || !Number.isFinite(top) || !Number.isFinite(width) || !Number.isFinite(height) || left < 0 || top < 0 || width < 0 || height < 0 || left + width > page.width || top + height > page.height) {
          issues.push(verificationIssue("pdf", "textItemOutOfBounds", `PDF text item ${item.id} extends outside page ${pageIndex + 1}.`, { severity: "warning", page: pageIndex + 1, id: item.id, bbox: item.bbox }));
        }
      }
      for (const region of page.regions) {
        const [left, top, width, height] = region.bbox || [];
        if (!Number.isFinite(left) || !Number.isFinite(top) || !Number.isFinite(width) || !Number.isFinite(height) || left < 0 || top < 0 || width <= 0 || height <= 0 || left + width > page.width || top + height > page.height) {
          issues.push(verificationIssue("pdf", "regionOutOfBounds", `PDF layout region ${region.id} extends outside page ${pageIndex + 1}.`, { severity: "warning", page: pageIndex + 1, id: region.id, bbox: region.bbox, regionKind: region.kind }));
        }
      }
      for (const table of page.tables) {
        if (!table.values.length || !table.values[0]?.length) issues.push(verificationIssue("pdf", "emptyTable", `PDF table ${table.id} on page ${pageIndex + 1} has no cells.`, { page: pageIndex + 1, id: table.id }));
        const tableGrid = table.grid();
        for (const cell of tableGrid.cells) {
          if (tableCellIds.has(cell.id)) issues.push(verificationIssue("pdf", "duplicateCellId", `PDF table cell ID ${cell.id} is duplicated across the artifact.`, { page: pageIndex + 1, id: cell.id, tableId: table.id }));
          else tableCellIds.add(cell.id);
        }
        for (const error of tableGrid.errors) issues.push(verificationIssue("pdf", error.code, `PDF table ${table.id}: ${error.message}`, { page: pageIndex + 1, id: error.id || table.id, tableId: table.id, row: error.row, column: error.column, headerId: error.headerId }));
        const [left, top, width, height] = table.bbox || [];
        if (left < 0 || top < 0 || width <= 0 || height <= 0 || left + width > page.width || top + height > page.height) {
          issues.push(verificationIssue("pdf", "tableOutOfBounds", `PDF table ${table.id} extends outside page ${pageIndex + 1}.`, { page: pageIndex + 1, id: table.id, bbox: table.bbox }));
        }
      }
      for (const image of page.images) {
        const accessibilityIssue = pdfFigureAccessibilityIssue(image, "PDF image");
        if (accessibilityIssue) issues.push(verificationIssue("pdf", accessibilityIssue.code, accessibilityIssue.message, { page: pageIndex + 1, id: image.id, figureKind: "image" }));
        if (!image.dataUrl && !image.uri && !image.prompt) issues.push(verificationIssue("pdf", "emptyImage", `PDF image ${image.id} on page ${pageIndex + 1} has no dataUrl, uri, or prompt.`, { page: pageIndex + 1, id: image.id }));
        if (image.dataUrl && !imageDataFromDataUrl(image.dataUrl)) issues.push(verificationIssue("pdf", "invalidImageDataUrl", `PDF image ${image.id} on page ${pageIndex + 1} has an unsupported data URL.`, { page: pageIndex + 1, id: image.id }));
        const [left, top, width, height] = image.bbox || [];
        if (left < 0 || top < 0 || width <= 0 || height <= 0 || left + width > page.width || top + height > page.height) {
          issues.push(verificationIssue("pdf", "imageOutOfBounds", `PDF image ${image.id} extends outside page ${pageIndex + 1}.`, { page: pageIndex + 1, id: image.id, bbox: image.bbox }));
        }
      }
      for (const chart of page.charts) {
        const accessibilityIssue = pdfFigureAccessibilityIssue(chart, "PDF chart");
        if (accessibilityIssue) issues.push(verificationIssue("pdf", accessibilityIssue.code, accessibilityIssue.message, { page: pageIndex + 1, id: chart.id, figureKind: "chart" }));
        if (!chart.categories.length || !chart.series.length || chart.series.every((series) => !series.values.length)) issues.push(verificationIssue("pdf", "emptyChart", `PDF chart ${chart.id} on page ${pageIndex + 1} has no categories or series data.`, { page: pageIndex + 1, id: chart.id }));
        for (const series of chart.series) {
          if (series.values.some((value) => !Number.isFinite(value))) issues.push(verificationIssue("pdf", "chartNonNumericData", `PDF chart ${chart.id} contains non-numeric series values.`, { page: pageIndex + 1, id: chart.id, series: series.name }));
        }
        const [left, top, width, height] = chart.bbox || [];
        if (left < 0 || top < 0 || width <= 0 || height <= 0 || left + width > page.width || top + height > page.height) {
          issues.push(verificationIssue("pdf", "chartOutOfBounds", `PDF chart ${chart.id} extends outside page ${pageIndex + 1}.`, { page: pageIndex + 1, id: chart.id, bbox: chart.bbox }));
        }
      }
    });
    for (const error of pdfHeadingNestingIssues(headingSequence)) issues.push(verificationIssue("pdf", error.code, `PDF heading sequence: ${error.message}`, { page: error.page, id: error.id, headingLevel: error.level, previousHeadingLevel: error.previousLevel }));
    return verificationResult("pdf", issues, options);
  }

  help(query = "*", options = {}) { return ndjson(queryHelpRecords("pdf", query, options), options.maxChars ?? Infinity); }

  layoutJson(options = {}) {
    const pageNumber = options.page != null ? Number(options.page) : options.pageIndex != null ? Number(options.pageIndex) + 1 : undefined;
    const selectedPages = pageNumber ? [this.pages[pageNumber - 1]].filter(Boolean) : this.pages;
    const layout = {
      kind: "pdfLayout",
      id: this.id,
      pageCount: this.pages.length,
      metadata: this.metadata,
      pages: selectedPages.map((page) => {
        const pageIndex = this.pages.indexOf(page);
        const pageNumber = pageIndex + 1;
        return {
          kind: "pdfPageLayout",
          id: page.id,
          page: pageNumber,
          width: page.width,
          height: page.height,
          unit: "pt",
          text: { kind: "text", id: `${page.id}/text`, page: pageNumber, text: page.text, textChars: page.text.length, bbox: [0, 0, page.width, page.height] },
          textItems: page.textItems.map((item) => ({ kind: "textItem", page: pageNumber, ...item })),
          regions: page.regions.map((region) => ({ kind: "region", regionKind: region.kind || "region", page: pageNumber, ...region })),
          tables: page.tables.map((table) => ({ kind: "table", id: table.id, page: pageNumber, name: table.name || undefined, values: table.values, cells: serializePdfTableCells(table), bbox: table.bbox, source: table.source || undefined })),
          images: page.images.map((image) => ({ kind: "image", id: image.id, page: pageNumber, name: image.name || undefined, alt: image.alt, decorative: image.decorative, bbox: image.bbox, fit: image.fit, hasDataUrl: Boolean(image.dataUrl), uri: image.uri, prompt: image.prompt, isMask: image.isMask || undefined, fillColor: image.fillColor, pixelWidth: image.pixelWidth, pixelHeight: image.pixelHeight, sourceObject: image.sourceObject, sourceOperator: image.sourceOperator })),
          charts: page.charts.map((chart) => ({ kind: "chart", id: chart.id, page: pageNumber, name: chart.name || undefined, title: chart.title, alt: chart.alt, decorative: chart.decorative, chartType: chart.chartType, categories: chart.categories, series: chart.series, bbox: chart.bbox })),
          readingOrder: page.readingOrderRecords(pageIndex),
        };
      }),
    };
    return pdfLayoutSlice(layout, options);
  }

  async render(options = {}) {
    const format = String(options.format || "").trim().toLowerCase();
    if (format === "layout" || format === LAYOUT_MIME) return new FileBlob(JSON.stringify(this.layoutJson(options), null, 2), { type: LAYOUT_MIME, metadata: { artifactKind: "pdf", format: "layout", page: options.page, pageIndex: options.pageIndex, target: options.target ?? options.targetId ?? options.id ?? options.anchor, search: options.search ?? options.searchTerm } });
    if (pdfRenderUsesPdfSource(options)) return renderPdfArtifactFromPdf(this, options);
    return new FileBlob(pdfPageSvg(this.pages[options.pageIndex || 0] || new PdfPage(this)), { type: "image/svg+xml" });
  }
  toJSON() { return { id: this.id, metadata: this.metadata, pages: this.pages.map((page) => page.toJSON()) }; }
}

export class PdfFile {
  static async inspectPdf(blobOrBuffer, options = {}) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const text = decoder.decode(bytes);
    const version = /^%PDF-(\d+\.\d+)/.exec(text)?.[1];
    const pages = [...text.matchAll(/\/Type\s*\/Page\b/g)].length;
    const objects = [...text.matchAll(/\b\d+\s+\d+\s+obj\b/g)].length;
    const readingOrderIds = inspectPdfReadingOrderIds(text);
    const figureAccessibility = inspectPdfFigureAccessibility(text);
    const structureRoles = {};
    for (const match of text.matchAll(/\/Type\s*\/StructElem\b[\s\S]*?\/S\s*\/([A-Za-z0-9]+)/g)) structureRoles[match[1]] = (structureRoles[match[1]] || 0) + 1;
    const headingLevels = Object.fromEntries(Array.from({ length: 6 }, (_, index) => [`H${index + 1}`, structureRoles[`H${index + 1}`] || 0]));
    const records = [
      { kind: "pdfFile", bytes: bytes.byteLength, version, pages, objects, hasEmbeddedModel: /%OPEN_OFFICE_ARTIFACT [A-Za-z0-9+/=]+/.test(text), hasEof: /%%EOF\s*$/.test(text), tagged: /\/StructTreeRoot\s+\d+\s+0\s+R/.test(text) && /\/MarkInfo\s*<<[^>]*\/Marked\s+true/.test(text), language: /\/Lang\s*\(([^)]*)\)/.exec(text)?.[1], embeddedFonts: [...text.matchAll(/\/Subtype\s*\/Type0\b/g)].length, subsetFonts: new Set([...text.matchAll(/\/BaseFont\s*\/([A-Z]{6}\+[A-Za-z0-9_-]+)/g)].map((match) => match[1])).size, toUnicodeMaps: [...text.matchAll(/\/ToUnicode\s+\d+\s+0\s+R/g)].length, structureElements: [...text.matchAll(/\/Type\s*\/StructElem\b/g)].length, structureRoles, headings: Object.values(headingLevels).reduce((sum, count) => sum + count, 0), headingLevels, readingOrderIds, readingOrderItems: readingOrderIds.length, ...figureAccessibility, tableStructures: structureRoles.Table || 0, tableRows: structureRoles.TR || 0, tableHeaders: structureRoles.TH || 0, tableDataCells: structureRoles.TD || 0, tableCellIds: [...text.matchAll(/\/S\s*\/(?:TH|TD)\b[\s\S]*?\/ID\s*(?:\([^)]*\)|<[A-Fa-f0-9]+>)/g)].length, rowSpans: [...text.matchAll(/\/RowSpan\s+[2-9]\d*/g)].length, columnSpans: [...text.matchAll(/\/ColSpan\s+[2-9]\d*/g)].length, headerAssociations: [...text.matchAll(/\/Headers\s*\[[^\]]+\]/g)].length, markedContentItems: [...text.matchAll(/\/MCID\s+\d+/g)].length },
      ...[...text.matchAll(/(\d+)\s+0\s+obj\s*<<([\s\S]*?)>>/g)].slice(0, Math.max(0, Number(options.maxObjects ?? 200) || 0)).map((match) => ({ kind: "pdfObject", object: Number(match[1]), type: /\/Type\s*\/([A-Za-z0-9]+)/.exec(match[2])?.[1], subtype: /\/Subtype\s*\/([A-Za-z0-9]+)/.exec(match[2])?.[1], stream: /\bstream\b/.test(match[0]) })),
    ];
    return { records, summary: records[0], ...ndjson(records, options.maxChars ?? Infinity) };
  }

  static async exportPdf(artifact, options = {}) {
    const language = options.language || options.lang || artifact.metadata?.language || artifact.metadata?.lang || "en-US";
    const title = options.title || artifact.metadata?.title || String(artifact.pages?.[0]?.text || "").split(/\r?\n/).find(Boolean) || "Office artifact";
    const embeddedFont = await resolvePdfEmbeddedFont(options.font ?? options.fontBytes ?? options.unicodeFont, options);
    const exportOptions = { ...options, language, title, embeddedFont };
    return new FileBlob(buildMinimalPdf(artifact, exportOptions), { type: PDF_MIME, metadata: { tagged: options.tagged !== false, language, title, embeddedFont: embeddedFont?.name, fontSubset: embeddedFont ? options.subsetFont !== false : undefined } });
  }

  static async importPdf(blobOrBuffer, options = {}) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const text = decoder.decode(bytes);
    const metadata = /%OPEN_OFFICE_ARTIFACT ([A-Za-z0-9+/=]+)/.exec(text)?.[1];
    if (metadata && options.preferParser !== true) return PdfArtifact.create(JSON.parse(Buffer.from(metadata, "base64").toString("utf8")));
    const parser = options.parser || options.parseAdapter || options.adapter;
    if (typeof parser === "function") {
      const parsed = await parser({ input: new FileBlob(bytes, { type: options.inputType || blobOrBuffer?.type || PDF_MIME }), bytes, inputType: options.inputType || blobOrBuffer?.type || PDF_MIME, artifactKind: "pdf", options });
      return pdfArtifactFromParserOutput(parsed, { parser: options.parserName || parsed?.parser || parsed?.metadata?.parser || "custom" });
    }
    const strings = [...text.matchAll(/\(([^()]*)\)\s*Tj/g)].map((m) => m[1].replaceAll("\\)", ")").replaceAll("\\(", "("));
    const plain = strings.join("\n");
    const tableRows = strings.filter((line) => line.includes("|")).map((line) => line.split("|").map((cell) => cell.trim()));
    const images = strings.filter((line) => /^\[Image: .+\]$/.test(line)).map((line, index) => ({ name: `extracted-image-${index + 1}`, alt: line.replace(/^\[Image: |\]$/g, ""), bbox: [72, 280 + index * 140, 180, 120] }));
    return PdfArtifact.create({ metadata: { parser: "heuristic" }, pages: [{ text: plain, tables: tableRows.length ? [{ name: "extracted-table", values: tableRows }] : [], images }] });
  }
}

function pdfImageDataUrl(image = {}) {
  if (image.dataUrl) return image.dataUrl;
  const base64 = image.base64 || image.base64Data || image.dataBase64;
  const contentType = image.contentType || image.mimeType || image.mime || image.type || imageDataFromDataUrl(image.dataUrl)?.contentType || imageContentTypeFromExtension(image.extension || image.ext || image.format || "png");
  if (typeof base64 === "string" && base64.trim()) return `data:${contentType};base64,${base64.replace(/^data:[^,]+,/, "")}`;
  const bytes = image.bytes || image.data || image.buffer || image.uint8Array || image.binary;
  if (bytes == null) return undefined;
  try {
    const payload = Array.isArray(bytes) ? Uint8Array.from(bytes) : toUint8Array(bytes);
    return `data:${contentType};base64,${Buffer.from(payload).toString("base64")}`;
  } catch {
    return undefined;
  }
}

function pdfTextItemText(item = {}) {
  return String(item.text ?? item.str ?? item.value ?? "").trim();
}

function pdfTextItemBBox(item = {}) {
  const bbox = item.bbox || item.bounds || item.rect;
  if (Array.isArray(bbox) && bbox.length >= 4) return bbox.slice(0, 4).map(Number);
  return [Number(item.x || item.left || 0), Number(item.y ?? item.top ?? 0), Number(item.width || 0), Number(item.height || 0)];
}

function pdfBboxForTextItems(items = []) {
  const boxes = items.map((item) => item.bbox || pdfTextItemBBox(item)).filter((bbox) => bbox.every(Number.isFinite));
  if (!boxes.length) return [0, 0, 0, 0];
  const left = Math.min(...boxes.map((bbox) => bbox[0]));
  const top = Math.min(...boxes.map((bbox) => bbox[1]));
  const right = Math.max(...boxes.map((bbox) => bbox[0] + Math.max(0, bbox[2])));
  const bottom = Math.max(...boxes.map((bbox) => bbox[1] + Math.max(0, bbox[3])));
  return [left, top, Math.max(1, right - left), Math.max(1, bottom - top)];
}

function medianNumber(values = []) {
  const sorted = values.filter(Number.isFinite).sort((a, b) => a - b);
  if (!sorted.length) return 0;
  return sorted[Math.floor(sorted.length / 2)];
}

function reconstructPdfTablesFromTextGeometry(page = {}, pageIndex = 0) {
  const rawItems = page.textItems || page.items || [];
  const items = rawItems.map((item, index) => ({ id: item.id || `txt/${pageIndex + 1}/${index + 1}`, text: pdfTextItemText(item), bbox: pdfTextItemBBox(item) })).filter((item) => item.text && item.bbox.every(Number.isFinite));
  if (items.length < 4) return [];
  const rowTolerance = Number(page.tableRowTolerance || 6);
  const rows = [];
  for (const item of items.sort((a, b) => (a.bbox[1] - b.bbox[1]) || (a.bbox[0] - b.bbox[0]))) {
    const centerY = item.bbox[1] + item.bbox[3] / 2;
    let row = rows.find((candidate) => Math.abs(candidate.centerY - centerY) <= rowTolerance);
    if (!row) { row = { centerY, items: [] }; rows.push(row); }
    row.items.push(item);
    row.centerY = medianNumber(row.items.map((entry) => entry.bbox[1] + entry.bbox[3] / 2));
  }
  const candidateRows = rows.map((row) => ({ ...row, items: row.items.sort((a, b) => a.bbox[0] - b.bbox[0]) })).filter((row) => row.items.length >= 2);
  if (candidateRows.length < 2) return [];
  const counts = new Map();
  for (const row of candidateRows) counts.set(row.items.length, (counts.get(row.items.length) || 0) + 1);
  const [columnCount, rowCount] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0] || [0, 0];
  if (columnCount < 2 || rowCount < 2) return [];
  const tableRows = candidateRows.filter((row) => row.items.length === columnCount);
  const columns = Array.from({ length: columnCount }, (_, col) => medianNumber(tableRows.map((row) => row.items[col]?.bbox[0])));
  const aligned = tableRows.every((row) => row.items.every((item, col) => Math.abs(item.bbox[0] - columns[col]) <= Math.max(16, item.bbox[2] * 0.4)));
  if (!aligned) return [];
  return [{ name: `geometry-table-${pageIndex + 1}`, values: tableRows.map((row) => row.items.map((item) => item.text)), bbox: pdfBboxForTextItems(tableRows.flatMap((row) => row.items)), source: "textGeometry" }];
}

function pdfArtifactFromParserOutput(parsed, metadata = {}) {
  if (parsed instanceof PdfArtifact) {
    parsed.metadata = { ...(parsed.metadata || {}), ...metadata };
    return parsed;
  }
  const source = parsed?.data || parsed?.document || parsed || {};
  const pages = (source.pages || []).map((page, index) => ({
    id: page.id,
    text: page.text ?? (page.lines || []).map((line) => typeof line === "string" ? line : line.text).join("\n"),
    width: Number(page.width || page.pageWidth || page.geometry?.width || 612),
    height: Number(page.height || page.pageHeight || page.geometry?.height || 792),
    textItems: page.textItems || page.items || [],
    regions: page.regions || [],
    tables: ((page.tables || []).length ? (page.tables || []) : reconstructPdfTablesFromTextGeometry(page, index)).map((table, tableIndex) => ({ id: table.id, name: table.name || `parsed-table-${index + 1}-${tableIndex + 1}`, values: table.values || table.rows || [], cells: table.cells || table.cellConfigs, bbox: table.bbox || table.bounds || [72, 140 + tableIndex * 120, 468, 96], source: table.source })),
    images: (page.images || []).map((image, imageIndex) => ({ name: image.name || `parsed-image-${index + 1}-${imageIndex + 1}`, alt: image.alt || image.altText || image.name || "Parsed raster content", decorative: image.decorative, dataUrl: pdfImageDataUrl(image), uri: image.uri, prompt: image.prompt, bbox: image.bbox || image.bounds || [72, 280 + imageIndex * 140, 180, 120], isMask: image.isMask, fillColor: image.fillColor, pixelWidth: image.pixelWidth, pixelHeight: image.pixelHeight, sourceObject: image.sourceObject, sourceOperator: image.sourceOperator })),
    charts: (page.charts || []).map((chart, chartIndex) => ({ name: chart.name || `parsed-chart-${index + 1}-${chartIndex + 1}`, title: chart.title || chart.name || `Parsed chart ${chartIndex + 1}`, alt: chart.alt || chart.altText || chart.title || chart.name || `Parsed chart ${chartIndex + 1}`, decorative: chart.decorative, chartType: chart.chartType || chart.type || "bar", categories: chart.categories || chart.labels || [], series: chart.series || [{ name: chart.seriesName || "Series 1", values: chart.values || chart.data || [] }], bbox: chart.bbox || chart.bounds || [72, 430 + chartIndex * 180, 468, 160] })),
  }));
  return PdfArtifact.create({ metadata: { ...metadata, ...(source.metadata || parsed?.metadata || {}) }, pages: pages.length ? pages : [{ text: "" }] });
}

function pdfChartSvg(chart) {
  const [left, top, width, height] = chart.bbox || [72, 430, 468, 180];
  const padding = { left: 36, right: 12, top: 28, bottom: 28 };
  const plotLeft = left + padding.left;
  const plotTop = top + padding.top;
  const plotWidth = Math.max(1, width - padding.left - padding.right);
  const plotHeight = Math.max(1, height - padding.top - padding.bottom);
  const values = chart.series.flatMap((series) => series.values).filter(Number.isFinite);
  const max = Math.max(1, ...values);
  const categories = chart.categories.length ? chart.categories : Array.from({ length: Math.max(0, ...chart.series.map((series) => series.values.length)) }, (_, index) => String(index + 1));
  const axis = `<rect x="${left}" y="${top}" width="${width}" height="${height}" fill="#ffffff" stroke="#cbd5e1"/><text x="${left + 10}" y="${top + 18}" font-family="Helvetica" font-size="12" font-weight="700" fill="#111827">${xmlEscape(chart.title)}</text><line x1="${plotLeft}" y1="${plotTop + plotHeight}" x2="${plotLeft + plotWidth}" y2="${plotTop + plotHeight}" stroke="#94a3b8"/><line x1="${plotLeft}" y1="${plotTop}" x2="${plotLeft}" y2="${plotTop + plotHeight}" stroke="#94a3b8"/>`;
  if (/^line$/i.test(chart.chartType)) {
    const lines = chart.series.map((series) => {
      const points = series.values.map((value, index) => {
        const x = plotLeft + (categories.length <= 1 ? plotWidth / 2 : (index / Math.max(1, categories.length - 1)) * plotWidth);
        const y = plotTop + plotHeight - (Math.max(0, Number(value) || 0) / max) * plotHeight;
        return `${x},${y}`;
      }).join(" ");
      return `<polyline points="${points}" fill="none" stroke="${attrEscape(series.color)}" stroke-width="2"/>`;
    }).join("");
    return `${axis}${lines}`;
  }
  const groupWidth = plotWidth / Math.max(1, categories.length);
  const barWidth = Math.max(2, groupWidth / Math.max(1, chart.series.length) - 4);
  const bars = [];
  chart.series.forEach((series, seriesIndex) => {
    series.values.forEach((value, index) => {
      const barHeight = (Math.max(0, Number(value) || 0) / max) * plotHeight;
      const x = plotLeft + index * groupWidth + seriesIndex * (barWidth + 3) + 3;
      const y = plotTop + plotHeight - barHeight;
      bars.push(`<rect x="${x}" y="${y}" width="${barWidth}" height="${barHeight}" fill="${attrEscape(series.color)}"/>`);
    });
  });
  const labels = categories.slice(0, 8).map((category, index) => `<text x="${plotLeft + index * groupWidth + 3}" y="${top + height - 8}" font-family="Helvetica" font-size="9" fill="#475569">${xmlEscape(category)}</text>`).join("");
  return `${axis}${bars.join("")}${labels}`;
}

function pdfPageSvg(page) {
  const width = page.width || 612;
  const height = page.height || 792;
  const lines = pdfPageBodyTextLines(page);
  const positionedText = (page.textItems || []).map((item) => {
    const [left, top, itemWidth, itemHeight] = item.bbox || [72, 72, 0, 14];
    const fontSize = Math.max(6, Number(item.fontSize || itemHeight || 12));
    return `<text x="${left}" y="${top + fontSize}" font-family="${xmlEscape(item.fontName || "Helvetica")}" font-size="${fontSize}" font-weight="${item.bold ? "700" : "400"}" font-style="${item.italic ? "italic" : "normal"}" fill="${xmlEscape(item.color || "#111827")}" data-text-item-id="${attrEscape(item.id || "")}">${xmlEscape(item.text || "")}</text>`;
  }).join("");
  const lineText = lines.map((line, index) => `<text x="72" y="${96 + index * (index ? 22 : 30)}" font-family="Helvetica" font-size="${index === 0 ? 24 : 14}" font-weight="${index === 0 ? "700" : "400"}" fill="${index === 0 ? "#0f172a" : "#334155"}">${xmlEscape(line)}</text>`).join("");
  const text = `${lineText}${positionedText}`;
  const tables = (page.tables || []).map((table) => {
    return serializePdfTableCells(table).map((cell) => {
      const [x, y, width, height] = cell.bbox;
      const textY = y + Math.max(11, (height - 11) / 2 + 11);
      return `<g data-table-cell-id="${attrEscape(cell.id)}" data-row-span="${cell.rowSpan}" data-column-span="${cell.columnSpan}"><rect x="${x}" y="${y}" width="${width}" height="${height}" fill="${cell.role === "TH" ? "#f1f5f9" : "#ffffff"}" stroke="#cbd5e1"/><text x="${x + 5}" y="${textY}" font-family="Helvetica" font-size="11" fill="#111827">${xmlEscape(cell.value)}</text></g>`;
    }).join("");
  }).join("");
  const images = (page.images || []).map((image) => {
    const [left, top, imageWidth, imageHeight] = image.bbox || [72, 280, 180, 120];
    if (image.dataUrl) return `<image href="${attrEscape(image.dataUrl)}" x="${left}" y="${top}" width="${imageWidth}" height="${imageHeight}" preserveAspectRatio="xMidYMid meet"/><text x="${left}" y="${top + imageHeight + 14}" font-family="Helvetica" font-size="10" fill="#475569">${xmlEscape(image.alt || image.name || "image")}</text>`;
    return `<rect x="${left}" y="${top}" width="${imageWidth}" height="${imageHeight}" fill="#fef3c7" stroke="#f59e0b"/><text x="${left + 8}" y="${top + 18}" font-family="Helvetica" font-size="11" fill="#92400e">${xmlEscape(image.alt || image.prompt || image.uri || image.name || "image")}</text>`;
  }).join("");
  const charts = (page.charts || []).map(pdfChartSvg).join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="white"/>${text}${tables}${images}${charts}</svg>`;
}

function escapePdfString(text) {
  return String(text).replace(/[^\x20-\x7e]/g, "?").replaceAll("\\", "\\\\").replaceAll("(", "\\(").replaceAll(")", "\\)");
}

function pdfUtf16Hex(text, bom = true) {
  const units = [...String(text)].flatMap((character) => {
    const codePoint = character.codePointAt(0);
    if (codePoint <= 0xffff) return [codePoint];
    const adjusted = codePoint - 0x10000;
    return [0xd800 + (adjusted >> 10), 0xdc00 + (adjusted & 0x3ff)];
  });
  return `${bom ? "feff" : ""}${units.map((unit) => unit.toString(16).padStart(4, "0")).join("")}`.toUpperCase();
}

function pdfStringToken(text) {
  const value = String(text ?? "");
  return /[^\x20-\x7e]/.test(value) ? `<${pdfUtf16Hex(value)}>` : `(${escapePdfString(value)})`;
}

function pdfFontRead(view, method, offset, bytes) {
  if (!Number.isInteger(offset) || offset < 0 || offset + bytes > view.byteLength) throw new Error("Truncated TrueType font table.");
  return view[method](offset, false);
}

function parsePdfTrueTypeFont(input, options = {}) {
  const bytes = toUint8Array(input);
  const configuredMax = Number(options.maxFontBytes ?? 16 * 1024 * 1024);
  const maxBytes = Number.isFinite(configuredMax) && configuredMax > 0 ? Math.max(1024, Math.floor(configuredMax)) : 16 * 1024 * 1024;
  if (bytes.byteLength > maxBytes) throw new Error(`TrueType font exceeds maxFontBytes (${bytes.byteLength} > ${maxBytes}).`);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (offset) => pdfFontRead(view, "getUint16", offset, 2);
  const i16 = (offset) => pdfFontRead(view, "getInt16", offset, 2);
  const u32 = (offset) => pdfFontRead(view, "getUint32", offset, 4);
  const scaler = u32(0);
  if (scaler === 0x74746366) throw new Error("TrueType collections (.ttc) are not supported; provide a standalone .ttf font.");
  if (scaler !== 0x00010000 && scaler !== 0x74727565) throw new Error("Only glyf-based TrueType .ttf fonts are supported for PDF embedding.");
  const tables = new Map();
  const tableCount = u16(4);
  for (let index = 0; index < tableCount; index += 1) {
    const record = 12 + index * 16;
    if (record + 16 > bytes.byteLength) throw new Error("Truncated TrueType table directory.");
    const tag = String.fromCharCode(bytes[record], bytes[record + 1], bytes[record + 2], bytes[record + 3]);
    const offset = u32(record + 8);
    const length = u32(record + 12);
    if (offset + length > bytes.byteLength) throw new Error(`TrueType table ${tag} exceeds the font bytes.`);
    tables.set(tag, { offset, length });
  }
  const required = (tag) => {
    const table = tables.get(tag);
    if (!table) throw new Error(`TrueType font is missing required ${tag} table.`);
    return table;
  };
  const head = required("head");
  const hhea = required("hhea");
  const maxp = required("maxp");
  const hmtx = required("hmtx");
  const cmap = required("cmap");
  const glyf = required("glyf");
  const loca = required("loca");
  if (head.length < 54 || hhea.length < 36 || maxp.length < 6) throw new Error("TrueType font has truncated metrics tables.");
  const unitsPerEm = u16(head.offset + 18);
  if (!unitsPerEm) throw new Error("TrueType font has invalid unitsPerEm.");
  const numberOfGlyphs = u16(maxp.offset + 4);
  const numberOfHMetrics = u16(hhea.offset + 34);
  if (!numberOfGlyphs || !numberOfHMetrics || numberOfHMetrics > numberOfGlyphs) throw new Error("TrueType font has invalid glyph metrics.");
  if (numberOfHMetrics * 4 > hmtx.length) throw new Error("TrueType hmtx table is truncated.");
  const locaFormat = i16(head.offset + 50);
  if (locaFormat !== 0 && locaFormat !== 1) throw new Error("TrueType head table has an unsupported indexToLocFormat.");
  const locaEntryBytes = locaFormat ? 4 : 2;
  if ((numberOfGlyphs + 1) * locaEntryBytes > loca.length) throw new Error("TrueType loca table is truncated.");
  const glyphOffsets = Array.from({ length: numberOfGlyphs + 1 }, (_, glyph) => locaFormat ? u32(loca.offset + glyph * 4) : u16(loca.offset + glyph * 2) * 2);
  for (let glyph = 0; glyph < numberOfGlyphs; glyph += 1) if (glyphOffsets[glyph] > glyphOffsets[glyph + 1] || glyphOffsets[glyph + 1] > glyf.length) throw new Error("TrueType loca offsets are invalid.");
  const advances = Array(numberOfGlyphs);
  let lastAdvance = 0;
  for (let glyph = 0; glyph < numberOfGlyphs; glyph += 1) {
    if (glyph < numberOfHMetrics) lastAdvance = u16(hmtx.offset + glyph * 4);
    advances[glyph] = lastAdvance;
  }
  if (cmap.length < 4) throw new Error("TrueType cmap table is truncated.");
  const cmapCount = u16(cmap.offset + 2);
  let chosen;
  for (let index = 0; index < cmapCount; index += 1) {
    const record = cmap.offset + 4 + index * 8;
    if (record + 8 > cmap.offset + cmap.length) throw new Error("TrueType cmap encoding records are truncated.");
    const platform = u16(record);
    const encoding = u16(record + 2);
    const relativeOffset = u32(record + 4);
    if (relativeOffset + 4 > cmap.length) continue;
    const subtableOffset = cmap.offset + relativeOffset;
    const format = u16(subtableOffset);
    if (format === 12 && relativeOffset + 8 > cmap.length) continue;
    const length = format === 12 ? u32(subtableOffset + 4) : format === 4 ? u16(subtableOffset + 2) : 0;
    if (!length || relativeOffset + length > cmap.length) continue;
    const score = format === 12 ? 100 + (platform === 3 && encoding === 10 ? 10 : 0) : format === 4 ? 50 + (platform === 3 ? 5 : 0) : 0;
    if (score && (!chosen || score > chosen.score)) chosen = { format, offset: subtableOffset, length, score };
  }
  if (!chosen) throw new Error("TrueType font lacks a supported Unicode cmap format 4 or 12.");
  let glyphForCodePoint;
  if (chosen.format === 12) {
    const groupCount = u32(chosen.offset + 12);
    if (16 + groupCount * 12 > chosen.length) throw new Error("TrueType cmap format 12 groups are truncated.");
    const groups = Array.from({ length: groupCount }, (_, index) => {
      const offset = chosen.offset + 16 + index * 12;
      return { start: u32(offset), end: u32(offset + 4), glyph: u32(offset + 8) };
    });
    glyphForCodePoint = (codePoint) => {
      let low = 0;
      let high = groups.length - 1;
      while (low <= high) {
        const middle = (low + high) >> 1;
        const group = groups[middle];
        if (codePoint < group.start) high = middle - 1;
        else if (codePoint > group.end) low = middle + 1;
        else return group.glyph + codePoint - group.start;
      }
      return 0;
    };
  } else {
    const segmentCount = u16(chosen.offset + 6) / 2;
    if (!Number.isInteger(segmentCount) || segmentCount < 1 || 16 + segmentCount * 8 > chosen.length) throw new Error("TrueType cmap format 4 segments are invalid.");
    const endCodes = chosen.offset + 14;
    const startCodes = endCodes + segmentCount * 2 + 2;
    const deltas = startCodes + segmentCount * 2;
    const rangeOffsets = deltas + segmentCount * 2;
    glyphForCodePoint = (codePoint) => {
      if (codePoint > 0xffff) return 0;
      for (let segment = 0; segment < segmentCount; segment += 1) {
        const end = u16(endCodes + segment * 2);
        if (codePoint > end) continue;
        const start = u16(startCodes + segment * 2);
        if (codePoint < start) return 0;
        const delta = i16(deltas + segment * 2);
        const rangeOffsetAddress = rangeOffsets + segment * 2;
        const rangeOffset = u16(rangeOffsetAddress);
        if (!rangeOffset) return (codePoint + delta) & 0xffff;
        const glyphAddress = rangeOffsetAddress + rangeOffset + (codePoint - start) * 2;
        if (glyphAddress + 2 > chosen.offset + chosen.length) throw new Error("TrueType cmap glyph index exceeds its subtable.");
        let glyph = u16(glyphAddress);
        if (glyph) glyph = (glyph + delta) & 0xffff;
        return glyph;
      }
      return 0;
    };
  }
  const scale = (value) => Math.round(value / unitsPerEm * 1000);
  return {
    bytes,
    tables,
    numberOfGlyphs,
    glyphOffsets,
    glyphTableOffset: glyf.offset,
    name: String(options.name || "OpenOfficeArtifactEmbedded").replace(/[^A-Za-z0-9_-]/g, "") || "OpenOfficeArtifactEmbedded",
    unitsPerEm,
    advances,
    glyphForCodePoint,
    bbox: [i16(head.offset + 36), i16(head.offset + 38), i16(head.offset + 40), i16(head.offset + 42)].map(scale),
    ascent: scale(i16(hhea.offset + 4)),
    descent: scale(i16(hhea.offset + 6)),
  };
}

async function resolvePdfEmbeddedFont(source, options = {}) {
  if (source == null) return undefined;
  let bytes;
  let name;
  if (source instanceof FileBlob) { bytes = source.bytes; name = source.name; }
  else if (typeof source === "string") { bytes = await fs.readFile(source); name = path.basename(source, path.extname(source)); }
  else if (ArrayBuffer.isView(source) || source instanceof ArrayBuffer || Array.isArray(source)) bytes = source;
  else if (typeof source === "object") {
    name = source.name;
    if (source.path) bytes = await fs.readFile(source.path);
    else if (source.bytes != null || source.data != null) bytes = source.bytes ?? source.data;
    else if (source.base64) bytes = Buffer.from(String(source.base64).replace(/^data:[^,]+,/, ""), "base64");
  }
  if (bytes == null) throw new Error("PDF font must be a .ttf path, FileBlob, byte array, ArrayBuffer, or { path|bytes|base64 } object.");
  return parsePdfTrueTypeFont(bytes, { maxFontBytes: options.maxFontBytes, name });
}

function pdfFontTableBytes(font, tag) {
  const table = font.tables.get(tag);
  return table ? font.bytes.slice(table.offset, table.offset + table.length) : undefined;
}

function pdfCompositeGlyphs(font, glyph) {
  const start = font.glyphOffsets[glyph];
  const end = font.glyphOffsets[glyph + 1];
  if (end - start < 10) return [];
  const bytes = font.bytes;
  const base = font.glyphTableOffset + start;
  const view = new DataView(bytes.buffer, bytes.byteOffset + base, end - start);
  if (view.getInt16(0, false) >= 0) return [];
  const dependencies = [];
  let offset = 10;
  let flags;
  do {
    if (offset + 4 > view.byteLength) throw new Error(`Composite TrueType glyph ${glyph} is truncated.`);
    flags = view.getUint16(offset, false);
    const component = view.getUint16(offset + 2, false);
    if (component >= font.numberOfGlyphs) throw new Error(`Composite TrueType glyph ${glyph} references invalid glyph ${component}.`);
    dependencies.push(component);
    offset += 4;
    offset += flags & 0x0001 ? 4 : 2;
    if (flags & 0x0008) offset += 2;
    else if (flags & 0x0040) offset += 4;
    else if (flags & 0x0080) offset += 8;
    if (offset > view.byteLength) throw new Error(`Composite TrueType glyph ${glyph} has truncated component arguments.`);
  } while (flags & 0x0020);
  return dependencies;
}

function pdfSubsetGlyphClosure(fontState) {
  const included = new Set([0, ...[...fontState.used.values()].map((mapping) => mapping.glyph)]);
  const pending = [...included];
  while (pending.length) {
    const glyph = pending.pop();
    for (const dependency of pdfCompositeGlyphs(fontState.font, glyph)) if (!included.has(dependency)) { included.add(dependency); pending.push(dependency); }
  }
  return included;
}

function pdfSubsetCmap(fontState) {
  const entries = [...fontState.used.values()].map(({ codePoint, glyph }) => ({ codePoint, glyph })).sort((a, b) => a.codePoint - b.codePoint);
  const groups = [];
  for (const entry of entries) {
    const previous = groups.at(-1);
    if (previous && entry.codePoint === previous.end + 1 && entry.glyph === previous.glyph + entry.codePoint - previous.start) previous.end = entry.codePoint;
    else groups.push({ start: entry.codePoint, end: entry.codePoint, glyph: entry.glyph });
  }
  const subtableLength = 16 + groups.length * 12;
  const bytes = new Uint8Array(12 + subtableLength);
  const view = new DataView(bytes.buffer);
  view.setUint16(0, 0, false);
  view.setUint16(2, 1, false);
  view.setUint16(4, 3, false);
  view.setUint16(6, 10, false);
  view.setUint32(8, 12, false);
  view.setUint16(12, 12, false);
  view.setUint16(14, 0, false);
  view.setUint32(16, subtableLength, false);
  view.setUint32(20, 0, false);
  view.setUint32(24, groups.length, false);
  groups.forEach((group, index) => {
    const offset = 28 + index * 12;
    view.setUint32(offset, group.start, false);
    view.setUint32(offset + 4, group.end, false);
    view.setUint32(offset + 8, group.glyph, false);
  });
  return bytes;
}

function pdfSfntChecksum(bytes) {
  const paddedLength = Math.ceil(bytes.byteLength / 4) * 4;
  const padded = paddedLength === bytes.byteLength ? bytes : (() => { const copy = new Uint8Array(paddedLength); copy.set(bytes); return copy; })();
  const view = new DataView(padded.buffer, padded.byteOffset, padded.byteLength);
  let sum = 0;
  for (let offset = 0; offset < padded.byteLength; offset += 4) sum = (sum + view.getUint32(offset, false)) >>> 0;
  return sum;
}

function pdfBuildSfnt(tables) {
  const entries = [...tables.entries()].filter(([, bytes]) => bytes).sort(([a], [b]) => a.localeCompare(b));
  const tableCount = entries.length;
  const largestPower = 2 ** Math.floor(Math.log2(tableCount));
  const directoryBytes = 12 + tableCount * 16;
  let totalBytes = directoryBytes;
  for (const [, bytes] of entries) totalBytes += Math.ceil(bytes.byteLength / 4) * 4;
  const output = new Uint8Array(totalBytes);
  const view = new DataView(output.buffer);
  view.setUint32(0, 0x00010000, false);
  view.setUint16(4, tableCount, false);
  view.setUint16(6, largestPower * 16, false);
  view.setUint16(8, Math.log2(largestPower), false);
  view.setUint16(10, tableCount * 16 - largestPower * 16, false);
  let tableOffset = directoryBytes;
  let headOffset;
  entries.forEach(([tag, bytes], index) => {
    const record = 12 + index * 16;
    for (let char = 0; char < 4; char += 1) output[record + char] = tag.charCodeAt(char);
    view.setUint32(record + 4, pdfSfntChecksum(bytes), false);
    view.setUint32(record + 8, tableOffset, false);
    view.setUint32(record + 12, bytes.byteLength, false);
    output.set(bytes, tableOffset);
    if (tag === "head") headOffset = tableOffset;
    tableOffset += Math.ceil(bytes.byteLength / 4) * 4;
  });
  if (headOffset == null) throw new Error("Subset TrueType font is missing head table.");
  view.setUint32(headOffset + 8, (0xb1b0afba - pdfSfntChecksum(output)) >>> 0, false);
  return output;
}

function pdfSubsetTrueTypeFont(fontState) {
  const { font } = fontState;
  const included = pdfSubsetGlyphClosure(fontState);
  const loca = new Uint8Array((font.numberOfGlyphs + 1) * 4);
  const locaView = new DataView(loca.buffer);
  const glyphParts = [];
  let glyphBytes = 0;
  for (let glyph = 0; glyph < font.numberOfGlyphs; glyph += 1) {
    locaView.setUint32(glyph * 4, glyphBytes, false);
    if (!included.has(glyph)) continue;
    const start = font.glyphOffsets[glyph];
    const end = font.glyphOffsets[glyph + 1];
    if (end > start) {
      const data = font.bytes.slice(font.glyphTableOffset + start, font.glyphTableOffset + end);
      glyphParts.push({ offset: glyphBytes, data });
      glyphBytes += Math.ceil(data.byteLength / 4) * 4;
    }
  }
  locaView.setUint32(font.numberOfGlyphs * 4, glyphBytes, false);
  const glyf = new Uint8Array(glyphBytes);
  for (const part of glyphParts) glyf.set(part.data, part.offset);
  const head = pdfFontTableBytes(font, "head").slice();
  new DataView(head.buffer, head.byteOffset, head.byteLength).setUint32(8, 0, false);
  new DataView(head.buffer, head.byteOffset, head.byteLength).setInt16(50, 1, false);
  const tables = new Map([
    ["cmap", pdfSubsetCmap(fontState)],
    ["glyf", glyf],
    ["head", head],
    ["hhea", pdfFontTableBytes(font, "hhea")],
    ["hmtx", pdfFontTableBytes(font, "hmtx")],
    ["loca", loca],
    ["maxp", pdfFontTableBytes(font, "maxp")],
  ]);
  for (const tag of ["OS/2", "cvt ", "fpgm", "gasp", "name", "post", "prep"]) if (font.tables.has(tag)) tables.set(tag, pdfFontTableBytes(font, tag));
  const bytes = pdfBuildSfnt(tables);
  fontState.subsetGlyphs = included.size;
  fontState.fontProgramBytes = bytes.byteLength;
  return bytes;
}

function pdfSubsetFontName(fontState) {
  let hash = 2166136261;
  for (const mapping of fontState.used.values()) {
    hash ^= mapping.codePoint;
    hash = Math.imul(hash, 16777619) >>> 0;
    hash ^= mapping.glyph;
    hash = Math.imul(hash, 16777619) >>> 0;
  }
  let prefix = "";
  for (let index = 0; index < 6; index += 1) { prefix += String.fromCharCode(65 + hash % 26); hash = Math.floor(hash / 26); }
  return `${prefix}+${fontState.font.name}`;
}

function pdfNumber(value) {
  const number = Number(value) || 0;
  return String(Math.round(number * 1000) / 1000);
}

function pdfRgb(color, fallback = "#111827") {
  const raw = String(color || fallback).replace(/^#/, "");
  const normalized = /^[A-Fa-f0-9]{3}$/.test(raw) ? raw.split("").map((ch) => ch + ch).join("") : /^[A-Fa-f0-9]{6}$/.test(raw) ? raw : String(fallback).replace(/^#/, "");
  return [0, 2, 4].map((index) => Number.parseInt(normalized.slice(index, index + 2), 16) / 255);
}

function pdfColorCommand(color, stroke = false) {
  return `${pdfRgb(color).map(pdfNumber).join(" ")} ${stroke ? "RG" : "rg"}`;
}

function pdfEmbeddedCid(fontState, codePoint) {
  const glyph = fontState.font.glyphForCodePoint(codePoint);
  if (!glyph || glyph > 0xffff || glyph >= fontState.font.advances.length) throw new Error(`Embedded PDF font ${fontState.font.name} does not contain U+${codePoint.toString(16).toUpperCase().padStart(4, "0")}.`);
  let cid = fontState.cidByCodePoint.get(codePoint);
  if (cid == null) {
    cid = fontState.used.size + 1;
    if (cid > 0xffff) throw new Error("Embedded PDF font uses more than 65,535 distinct Unicode code points.");
    fontState.cidByCodePoint.set(codePoint, cid);
    fontState.used.set(cid, { codePoint, glyph });
  }
  return cid;
}

function pdfTextOperand(text, fontState) {
  const value = String(text ?? "");
  if (!fontState) {
    if (/[^\x20-\x7e]/.test(value)) throw new Error("PDF text contains non-ASCII characters; provide PdfFile.exportPdf(..., { font }) with a Unicode TrueType .ttf font.");
    return `(${escapePdfString(value)})`;
  }
  const cids = [...value].map((character) => pdfEmbeddedCid(fontState, character.codePointAt(0)));
  return `<${cids.map((cid) => cid.toString(16).padStart(4, "0")).join("").toUpperCase()}>`;
}

function pdfExportTextWidth(text, fontSize, fontState) {
  if (!fontState) return pdfTextWidth(text, fontSize);
  const units = [...String(text ?? "")].reduce((sum, character) => {
    const cid = pdfEmbeddedCid(fontState, character.codePointAt(0));
    const glyph = fontState.used.get(cid).glyph;
    return sum + (fontState.font.advances[glyph] || fontState.font.unitsPerEm);
  }, 0);
  return units / fontState.font.unitsPerEm * fontSize;
}

function pdfTextCommand(page, text, left, top, options = {}) {
  const fontSize = Math.max(1, Number(options.fontSize || 12));
  const baseline = page.height - Number(top) - fontSize;
  return `BT /${options.bold ? "F2" : "F1"} ${pdfNumber(fontSize)} Tf ${pdfColorCommand(options.color || "#111827")} ${pdfNumber(left)} ${pdfNumber(baseline)} Td ${pdfTextOperand(text, options.fontState)} Tj ET`;
}

function pdfLineCommand(page, x1, top1, x2, top2, options = {}) {
  return `q ${pdfColorCommand(options.color || "#94a3b8", true)} ${pdfNumber(options.width || 1)} w ${pdfNumber(x1)} ${pdfNumber(page.height - top1)} m ${pdfNumber(x2)} ${pdfNumber(page.height - top2)} l S Q`;
}

function pdfRectCommand(page, bbox, options = {}) {
  const [left, top, width, height] = bbox.map(Number);
  const operator = options.fill === false ? "S" : options.stroke === false ? "f" : "B";
  return `q ${pdfColorCommand(options.fillColor || "#ffffff")} ${pdfColorCommand(options.strokeColor || "#cbd5e1", true)} ${pdfNumber(options.lineWidth || 1)} w ${pdfNumber(left)} ${pdfNumber(page.height - top - height)} ${pdfNumber(width)} ${pdfNumber(height)} re ${operator} Q`;
}

function pdfFitText(text, width, fontSize, fontState) {
  const value = String(text ?? "");
  const available = Math.max(1, Number(width));
  if (pdfExportTextWidth(value, fontSize, fontState) <= available) return value;
  const chars = [...value];
  while (chars.length && pdfExportTextWidth(`${chars.join("")}...`, fontSize, fontState) > available) chars.pop();
  return chars.length ? `${chars.join("")}...` : "...";
}

function pdfPageTextCommands(page, fontState) {
  const lines = pdfPageBodyTextLines(page);
  return lines.map((line, index) => pdfTextCommand(page, line, 72, 72 + index * (index ? 22 : 30), { fontSize: index === 0 ? 24 : 14, bold: index === 0, color: index === 0 ? "#0f172a" : "#334155", fontState }));
}

function pdfPositionedTextCommands(page, fontState) {
  return (page.textItems || []).filter((item) => item.text).map((item) => {
    const [left, top, width, height] = item.bbox || [72, 72, 120, 14];
    const fontSize = Math.max(6, Number(item.fontSize || height || 12));
    return pdfTextCommand(page, pdfFitText(item.text, width || page.width - left, fontSize, fontState), left, top, { fontSize, color: item.color || "#111827", bold: Boolean(item.bold), fontState });
  });
}

function pdfTableSemanticRows(page, table, fontState) {
  const grid = table.grid();
  return Array.from({ length: Math.max(1, grid.rows) }, (_, row) => ({
    role: "TR",
    children: grid.cells.filter((cell) => cell.row === row).map((cell) => {
      const bbox = pdfTableCellBBox(table, cell);
      const fontSize = Math.max(8, Math.min(12, bbox[3] * 0.32));
      return {
        role: cell.role,
        structureId: cell.id,
        tableAttributes: {
          scope: cell.scope,
          rowSpan: cell.rowSpan,
          columnSpan: cell.columnSpan,
          headers: cell.role === "TD" || cell.headers.length ? cell.effectiveHeaders : [],
        },
        commands: [
          pdfRectCommand(page, bbox, { fillColor: cell.role === "TH" ? "#e2e8f0" : row % 2 ? "#f8fafc" : "#ffffff", strokeColor: "#94a3b8", lineWidth: 0.75 }),
          pdfTextCommand(page, pdfFitText(cell.value, bbox[2] - 12, fontSize, fontState), bbox[0] + 6, bbox[1] + Math.max(5, (bbox[3] - fontSize) / 2), { fontSize, bold: cell.role === "TH", color: "#0f172a", fontState }),
        ],
      };
    }),
  }));
}

function pdfChartCommands(page, chart, fontState) {
  const [left, top, width, height] = chart.bbox.map(Number);
  const commands = [pdfRectCommand(page, chart.bbox, { fillColor: "#ffffff", strokeColor: "#cbd5e1", lineWidth: 1 }), pdfTextCommand(page, chart.title, left + 10, top + 10, { fontSize: 13, bold: true, fontState })];
  const plot = { left: left + 42, top: top + 40, width: Math.max(1, width - 62), height: Math.max(1, height - 76) };
  commands.push(pdfLineCommand(page, plot.left, plot.top + plot.height, plot.left + plot.width, plot.top + plot.height));
  commands.push(pdfLineCommand(page, plot.left, plot.top, plot.left, plot.top + plot.height));
  const values = chart.series.flatMap((series) => series.values).filter(Number.isFinite);
  const max = Math.max(1, ...values);
  const categories = chart.categories.length ? chart.categories : Array.from({ length: Math.max(0, ...chart.series.map((series) => series.values.length)) }, (_, index) => String(index + 1));
  if (/^line$/i.test(chart.chartType)) {
    chart.series.forEach((series) => {
      const points = series.values.map((value, index) => ({ x: plot.left + (categories.length <= 1 ? plot.width / 2 : index / Math.max(1, categories.length - 1) * plot.width), top: plot.top + plot.height - Math.max(0, Number(value) || 0) / max * plot.height }));
      for (let index = 1; index < points.length; index += 1) commands.push(pdfLineCommand(page, points[index - 1].x, points[index - 1].top, points[index].x, points[index].top, { color: series.color, width: 2 }));
    });
  } else {
    const groupWidth = plot.width / Math.max(1, categories.length);
    const barWidth = Math.max(2, groupWidth / Math.max(1, chart.series.length) * 0.72);
    chart.series.forEach((series, seriesIndex) => series.values.forEach((value, index) => {
      const barHeight = Math.max(0, Number(value) || 0) / max * plot.height;
      const x = plot.left + index * groupWidth + seriesIndex * barWidth + groupWidth * 0.1;
      commands.push(pdfRectCommand(page, [x, plot.top + plot.height - barHeight, Math.max(1, barWidth - 2), barHeight], { fillColor: series.color, stroke: false }));
    }));
  }
  categories.slice(0, 8).forEach((category, index) => commands.push(pdfTextCommand(page, pdfFitText(category, plot.width / Math.max(1, categories.length), 8, fontState), plot.left + index * (plot.width / Math.max(1, categories.length)), top + height - 18, { fontSize: 8, color: "#475569", fontState })));
  chart.series.slice(0, 4).forEach((series, index) => {
    commands.push(pdfRectCommand(page, [left + width - 94, top + 12 + index * 14, 8, 8], { fillColor: series.color, stroke: false }));
    commands.push(pdfTextCommand(page, pdfFitText(series.name, 74, 8, fontState), left + width - 82, top + 10 + index * 14, { fontSize: 8, color: "#334155", fontState }));
  });
  return commands;
}

function jpegImageInfo(bytes) {
  if (bytes?.[0] !== 0xff || bytes?.[1] !== 0xd8) throw new Error("not a JPEG file");
  const sofMarkers = new Set([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf]);
  let offset = 2;
  while (offset + 3 < bytes.byteLength) {
    while (bytes[offset] === 0xff) offset += 1;
    const marker = bytes[offset++];
    if (marker === 0xd9 || marker === 0xda) break;
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    if (offset + 1 >= bytes.byteLength) break;
    const length = (bytes[offset] << 8) | bytes[offset + 1];
    if (length < 2 || offset + length > bytes.byteLength) throw new Error("truncated JPEG segment");
    if (sofMarkers.has(marker)) {
      if (length < 8) throw new Error("invalid JPEG frame header");
      const height = (bytes[offset + 3] << 8) | bytes[offset + 4];
      const width = (bytes[offset + 5] << 8) | bytes[offset + 6];
      const components = bytes[offset + 7];
      if (!width || !height || !components) throw new Error("invalid JPEG dimensions");
      return { width, height, components };
    }
    offset += length;
  }
  throw new Error("JPEG frame header not found");
}

function pdfImageAsset(image, objectId, resourceName) {
  const data = imageDataFromDataUrl(image.dataUrl);
  if (!data || !["image/png", "image/jpeg"].includes(data.contentType)) return undefined;
  if (data.contentType === "image/jpeg") {
    try {
      const decoded = jpegImageInfo(data.bytes);
      return {
        image,
        objectId,
        resourceName,
        width: decoded.width,
        height: decoded.height,
        colorSpace: decoded.components === 1 ? "/DeviceGray" : decoded.components === 4 ? "/DeviceCMYK" : "/DeviceRGB",
        filter: "/DCTDecode",
        streamBytes: data.bytes,
      };
    } catch (error) {
      throw new Error(`Unable to embed JPEG image ${image.name || image.id || "image"}: ${error.message}`);
    }
  }
  try {
    const decoded = decodePngRgba(data.bytes);
    const rgb = new Uint8Array(decoded.width * decoded.height * 3);
    for (let pixel = 0; pixel < decoded.width * decoded.height; pixel += 1) {
      const source = pixel * 4;
      const target = pixel * 3;
      const alpha = decoded.pixels[source + 3] / 255;
      rgb[target] = Math.round(decoded.pixels[source] * alpha + 255 * (1 - alpha));
      rgb[target + 1] = Math.round(decoded.pixels[source + 1] * alpha + 255 * (1 - alpha));
      rgb[target + 2] = Math.round(decoded.pixels[source + 2] * alpha + 255 * (1 - alpha));
    }
    return { image, objectId, resourceName, width: decoded.width, height: decoded.height, colorSpace: "/DeviceRGB", filter: "/FlateDecode", streamBytes: deflateSync(rgb) };
  } catch (error) {
    throw new Error(`Unable to embed PNG image ${image.name || image.id || "image"}: ${error.message}`);
  }
}

function pdfImageCommands(page, image, asset, fontState) {
  const [left, top, width, height] = image.bbox.map(Number);
  if (!asset) return [pdfRectCommand(page, image.bbox, { fillColor: "#fef3c7", strokeColor: "#f59e0b" }), pdfTextCommand(page, pdfFitText(image.alt || image.prompt || image.uri || image.name || "image", width - 16, 10, fontState), left + 8, top + 8, { fontSize: 10, color: "#92400e", fontState })];
  const sourceRatio = asset.width / asset.height;
  const frameRatio = width / height;
  const cover = image.fit === "cover";
  const drawWidth = cover
    ? (sourceRatio > frameRatio ? height * sourceRatio : width)
    : (sourceRatio > frameRatio ? width : height * sourceRatio);
  const drawHeight = cover
    ? (sourceRatio > frameRatio ? height : width / sourceRatio)
    : (sourceRatio > frameRatio ? width / sourceRatio : height);
  const x = left + (width - drawWidth) / 2;
  const drawTop = top + (height - drawHeight) / 2;
  const draw = `${pdfNumber(drawWidth)} 0 0 ${pdfNumber(drawHeight)} ${pdfNumber(x)} ${pdfNumber(page.height - drawTop - drawHeight)} cm /${asset.resourceName} Do`;
  if (cover) return [`q ${pdfNumber(left)} ${pdfNumber(page.height - top - height)} ${pdfNumber(width)} ${pdfNumber(height)} re W n ${draw} Q`];
  return [`q ${draw} Q`];
}

function pdfSemanticPlan(page, assetByImage, fontState) {
  const contentGroups = [];
  const bodyId = `${page.id}/text`;
  pdfPageTextCommands(page, fontState).forEach((command, index) => contentGroups.push({ role: index === 0 ? "H1" : "P", commands: [command], sourceId: bodyId, structureId: index === 0 ? bodyId : undefined }));
  const positionedItems = page.textItems.filter((item) => String(item.text || "").trim());
  pdfPositionedTextCommands(page, fontState).forEach((command, index) => {
    const item = positionedItems[index];
    contentGroups.push({ role: item.headingLevel ? `H${item.headingLevel}` : "P", commands: [command], sourceId: item.id, structureId: item.id });
  });
  page.tables.forEach((table) => contentGroups.push({ role: "Table", title: table.name || "Data table", children: pdfTableSemanticRows(page, table, fontState), sourceId: table.id, structureId: table.id }));
  page.images.forEach((image) => contentGroups.push({ role: "Figure", alt: image.alt, artifact: image.decorative, commands: pdfImageCommands(page, image, assetByImage.get(image), fontState), sourceId: image.id, structureId: image.id }));
  page.charts.forEach((chart) => contentGroups.push({ role: "Figure", alt: chart.alt, artifact: chart.decorative, commands: pdfChartCommands(page, chart, fontState), sourceId: chart.id, structureId: chart.id }));
  const semanticContentGroups = contentGroups.filter((group) => !group.artifact);
  let mcid = 0;
  for (const leaf of pdfSemanticLeaves(semanticContentGroups)) leaf.mcid = mcid++;
  const groupsBySourceId = new Map();
  for (const group of semanticContentGroups) {
    if (!groupsBySourceId.has(group.sourceId)) groupsBySourceId.set(group.sourceId, []);
    groupsBySourceId.get(group.sourceId).push(group);
  }
  const semanticGroups = resolvePdfReadingOrder(page).flatMap((entry) => groupsBySourceId.get(entry.id) || []);
  return { contentGroups, semanticGroups };
}

function pdfSemanticNodes(groups) {
  return groups.flatMap((group) => [group, ...pdfSemanticNodes(group.children || [])]);
}

function pdfSemanticLeaves(groups) {
  return groups.flatMap((group) => Array.isArray(group.children) ? pdfSemanticLeaves(group.children) : [group]);
}

function pdfMarkedContent(group) {
  if (group.artifact) return `/Artifact BMC\n${group.commands.join("\n")}\nEMC`;
  if (Array.isArray(group.children)) return group.children.map(pdfMarkedContent).join("\n");
  return `/${group.role} << /MCID ${group.mcid} >> BDC\n${group.commands.join("\n")}\nEMC`;
}

function pdfToUnicodeCmap(fontState) {
  const mappings = [...fontState.used.entries()].sort((a, b) => a[0] - b[0]);
  const chunks = [];
  for (let offset = 0; offset < mappings.length; offset += 100) {
    const slice = mappings.slice(offset, offset + 100);
    chunks.push(`${slice.length} beginbfchar\n${slice.map(([cid, mapping]) => `<${cid.toString(16).padStart(4, "0").toUpperCase()}> <${pdfUtf16Hex(String.fromCodePoint(mapping.codePoint), false)}> `).join("\n")}\nendbfchar`);
  }
  return `/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /${fontState.font.name}-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n${chunks.join("\n")}\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend`;
}

function pdfEmbeddedFontObjectBodies(fontState, ids) {
  const { font } = fontState;
  const fontProgram = fontState.subset ? pdfSubsetTrueTypeFont(fontState) : font.bytes;
  const pdfFontName = fontState.subset ? pdfSubsetFontName(fontState) : font.name;
  const mappings = [...fontState.used.entries()].sort((a, b) => a[0] - b[0]);
  const widths = mappings.map(([cid, mapping]) => `${cid} [${Math.round((font.advances[mapping.glyph] || font.unitsPerEm) / font.unitsPerEm * 1000)}]`).join(" ");
  const cmap = Buffer.from(pdfToUnicodeCmap(fontState), "ascii");
  const compressedFont = deflateSync(fontProgram);
  const cidToGid = new Uint8Array((mappings.at(-1)?.[0] + 1 || 1) * 2);
  for (const [cid, mapping] of mappings) {
    cidToGid[cid * 2] = mapping.glyph >> 8;
    cidToGid[cid * 2 + 1] = mapping.glyph & 0xff;
  }
  const compressedCidToGid = deflateSync(cidToGid);
  return new Map([
    [ids.type0, Buffer.from(`<< /Type /Font /Subtype /Type0 /BaseFont /${pdfFontName} /Encoding /Identity-H /DescendantFonts [${ids.cid} 0 R] /ToUnicode ${ids.toUnicode} 0 R >>`, "ascii")],
    [ids.cid, Buffer.from(`<< /Type /Font /Subtype /CIDFontType2 /BaseFont /${pdfFontName} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${ids.descriptor} 0 R /DW 1000 /W [${widths}] /CIDToGIDMap ${ids.cidToGid} 0 R >>`, "ascii")],
    [ids.descriptor, Buffer.from(`<< /Type /FontDescriptor /FontName /${pdfFontName} /Flags 4 /FontBBox [${font.bbox.join(" ")}] /ItalicAngle 0 /Ascent ${font.ascent} /Descent ${font.descent} /CapHeight ${font.ascent} /StemV 80 /FontFile2 ${ids.file} 0 R >>`, "ascii")],
    [ids.file, Buffer.concat([Buffer.from(`<< /Length ${compressedFont.byteLength} /Length1 ${fontProgram.byteLength} /Filter /FlateDecode >>\nstream\n`, "ascii"), compressedFont, Buffer.from("\nendstream", "ascii")])],
    [ids.toUnicode, Buffer.concat([Buffer.from(`<< /Length ${cmap.byteLength} >>\nstream\n`, "ascii"), cmap, Buffer.from("\nendstream", "ascii")])],
    [ids.cidToGid, Buffer.concat([Buffer.from(`<< /Length ${compressedCidToGid.byteLength} /Filter /FlateDecode >>\nstream\n`, "ascii"), compressedCidToGid, Buffer.from("\nendstream", "ascii")])],
  ]);
}

function buildMinimalPdf(artifact, options = {}) {
  const pages = artifact.pages.length ? artifact.pages : [new PdfPage(artifact)];
  const metadata = Buffer.from(JSON.stringify(artifact.toJSON()), "utf8").toString("base64");
  const tagged = options.tagged !== false;
  const language = String(options.language || options.lang || artifact.metadata?.language || artifact.metadata?.lang || "en-US");
  const title = String(options.title || artifact.metadata?.title || String(pages[0]?.text || "").split(/\r?\n/).find(Boolean) || "Office artifact");
  const fontState = options.embeddedFont ? { font: options.embeddedFont, subset: options.subsetFont !== false, used: new Map(), cidByCodePoint: new Map() } : undefined;
  let nextObjectId = 3;
  let embeddedFontIds;
  if (fontState) embeddedFontIds = { type0: nextObjectId++, cid: nextObjectId++, descriptor: nextObjectId++, file: nextObjectId++, toUnicode: nextObjectId++, cidToGid: nextObjectId++ };
  else nextObjectId = 5;
  const plans = pages.map((page) => {
    const imageAssets = [];
    for (const image of page.images) {
      const asset = pdfImageAsset(image, nextObjectId, `Im${imageAssets.length + 1}`);
      if (asset) { imageAssets.push(asset); nextObjectId += 1; }
    }
    const pageObjectId = nextObjectId++;
    const contentObjectId = nextObjectId++;
    const assetByImage = new Map(imageAssets.map((asset) => [asset.image, asset]));
    const semanticPlan = pdfSemanticPlan(page, assetByImage, fontState);
    return { page, imageAssets, assetByImage, pageObjectId, contentObjectId, ...semanticPlan };
  });
  const structTreeRootObjectId = tagged ? nextObjectId++ : undefined;
  const parentTreeObjectId = tagged ? nextObjectId++ : undefined;
  if (tagged) for (const plan of plans) for (const group of pdfSemanticNodes(plan.semanticGroups)) group.structureObjectId = nextObjectId++;
  const infoObjectId = nextObjectId++;
  const objects = new Map();
  const taggedCatalog = tagged ? ` /StructTreeRoot ${structTreeRootObjectId} 0 R /MarkInfo << /Marked true >> /Lang ${pdfStringToken(language)}` : "";
  objects.set(1, Buffer.from(`<< /Type /Catalog /Pages 2 0 R${taggedCatalog} /ViewerPreferences << /DisplayDocTitle true >> >>`, "ascii"));
  objects.set(2, Buffer.from(`<< /Type /Pages /Kids [${plans.map((plan) => `${plan.pageObjectId} 0 R`).join(" ")}] /Count ${plans.length} >>`, "ascii"));
  if (fontState) for (const [objectId, body] of pdfEmbeddedFontObjectBodies(fontState, embeddedFontIds)) objects.set(objectId, body);
  else {
    objects.set(3, Buffer.from("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>", "ascii"));
    objects.set(4, Buffer.from("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>", "ascii"));
  }
  for (const [pageIndex, plan] of plans.entries()) {
    const { page, imageAssets, contentGroups } = plan;
    const commands = tagged ? contentGroups.map(pdfMarkedContent) : pdfSemanticLeaves(contentGroups).flatMap((group) => group.commands);
    const content = Buffer.from(`${commands.join("\n")}\n`, "ascii");
    const xobjects = imageAssets.length ? `/XObject << ${imageAssets.map((asset) => `/${asset.resourceName} ${asset.objectId} 0 R`).join(" ")} >>` : "";
    const structParents = tagged ? ` /StructParents ${pageIndex}` : "";
    const fontResources = fontState ? `/F1 ${embeddedFontIds.type0} 0 R /F2 ${embeddedFontIds.type0} 0 R` : "/F1 3 0 R /F2 4 0 R";
    objects.set(plan.pageObjectId, Buffer.from(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pdfNumber(page.width || 612)} ${pdfNumber(page.height || 792)}] /Resources << /Font << ${fontResources} >> ${xobjects} >> /Contents ${plan.contentObjectId} 0 R${structParents} >>`, "ascii"));
    objects.set(plan.contentObjectId, Buffer.concat([Buffer.from(`<< /Length ${content.byteLength} >>\nstream\n`, "ascii"), content, Buffer.from("endstream", "ascii")]));
    for (const asset of imageAssets) objects.set(asset.objectId, Buffer.concat([Buffer.from(`<< /Type /XObject /Subtype /Image /Width ${asset.width} /Height ${asset.height} /ColorSpace ${asset.colorSpace} /BitsPerComponent 8 /Filter ${asset.filter} /Length ${asset.streamBytes.byteLength} >>\nstream\n`, "ascii"), asset.streamBytes, Buffer.from("\nendstream", "ascii")]));
  }
  if (tagged) {
    const structureRefs = plans.flatMap((plan) => plan.semanticGroups.map((group) => `${group.structureObjectId} 0 R`));
    objects.set(structTreeRootObjectId, Buffer.from(`<< /Type /StructTreeRoot /K [${structureRefs.join(" ")}] /ParentTree ${parentTreeObjectId} 0 R /ParentTreeNextKey ${plans.length} >>`, "ascii"));
    const parentTreeNums = plans.map((plan, pageIndex) => `${pageIndex} [${pdfSemanticLeaves(plan.semanticGroups).sort((left, right) => left.mcid - right.mcid).map((group) => `${group.structureObjectId} 0 R`).join(" ")}]`).join(" ");
    objects.set(parentTreeObjectId, Buffer.from(`<< /Nums [${parentTreeNums}] >>`, "ascii"));
    for (const plan of plans) {
      const writeStructureGroup = (group, parentObjectId) => {
        const alt = group.alt ? ` /Alt ${pdfStringToken(group.alt)}` : "";
        const titleEntry = group.title ? ` /T ${pdfStringToken(group.title)}` : "";
        const tableAttributeEntries = group.tableAttributes ? [
          "/O /Table",
          group.tableAttributes.scope ? `/Scope /${group.tableAttributes.scope}` : "",
          group.tableAttributes.rowSpan > 1 ? `/RowSpan ${group.tableAttributes.rowSpan}` : "",
          group.tableAttributes.columnSpan > 1 ? `/ColSpan ${group.tableAttributes.columnSpan}` : "",
          group.tableAttributes.headers?.length ? `/Headers [${group.tableAttributes.headers.map(pdfStringToken).join(" ")}]` : "",
        ].filter(Boolean) : [];
        const attributes = tableAttributeEntries.length ? ` /A << ${tableAttributeEntries.join(" ")} >>` : "";
        const structureId = group.structureId ? ` /ID ${pdfStringToken(group.structureId)}` : "";
        const kid = Array.isArray(group.children) ? `[${group.children.map((child) => `${child.structureObjectId} 0 R`).join(" ")}]` : group.mcid;
        objects.set(group.structureObjectId, Buffer.from(`<< /Type /StructElem /S /${group.role} /P ${parentObjectId} 0 R /Pg ${plan.pageObjectId} 0 R /K ${kid}${alt}${titleEntry}${structureId}${attributes} >>`, "ascii"));
        for (const child of group.children || []) writeStructureGroup(child, group.structureObjectId);
      };
      for (const group of plan.semanticGroups) writeStructureGroup(group, structTreeRootObjectId);
    }
  }
  objects.set(infoObjectId, Buffer.from(`<< /Title ${pdfStringToken(title)} /Creator (open-office-artifact-tool) /Producer (open-office-artifact-tool clean-room PDF writer) >>`, "ascii"));
  const header = Buffer.from(`%PDF-1.7\n%OPEN_OFFICE_ARTIFACT ${metadata}\n`, "ascii");
  const chunks = [header];
  const offsets = Array(nextObjectId).fill(0);
  let byteLength = header.byteLength;
  for (let objectId = 1; objectId < nextObjectId; objectId += 1) {
    const body = objects.get(objectId);
    if (!body) throw new Error(`Missing PDF object ${objectId}.`);
    offsets[objectId] = byteLength;
    const object = Buffer.concat([Buffer.from(`${objectId} 0 obj\n`, "ascii"), body, Buffer.from("\nendobj\n", "ascii")]);
    chunks.push(object);
    byteLength += object.byteLength;
  }
  const xrefOffset = byteLength;
  let xref = `xref\n0 ${nextObjectId}\n0000000000 65535 f \n`;
  for (let objectId = 1; objectId < nextObjectId; objectId += 1) xref += `${String(offsets[objectId]).padStart(10, "0")} 00000 n \n`;
  xref += `trailer\n<< /Size ${nextObjectId} /Root 1 0 R /Info ${infoObjectId} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`;
  chunks.push(Buffer.from(xref, "ascii"));
  return new Uint8Array(Buffer.concat(chunks));
}
