import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PDF_MIME = "application/pdf";
const LAYOUT_MIME = "application/vnd.open-office-artifact.layout+json";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
let nextId = 1;

function aid(prefix) {
  return `${prefix}/${(nextId++).toString(36).padStart(4, "0")}`;
}

function xmlEscape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function attrEscape(value) {
  return xmlEscape(value).replaceAll('"', "&quot;");
}

function stripTags(value) {
  return String(value ?? "").replace(/<[^>]+>/g, "");
}

function decodeXml(value) {
  return String(value ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function columnNumberToLabel(index) {
  let n = index + 1;
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function columnLabelToNumber(label) {
  let n = 0;
  for (const ch of String(label).toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

function parseCellAddress(address) {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(String(address).trim());
  if (!match) throw new Error(`Invalid cell address: ${address}`);
  return { col: columnLabelToNumber(match[1]), row: Number(match[2]) - 1 };
}

function makeCellAddress(row, col) {
  return `${columnNumberToLabel(col)}${row + 1}`;
}

function parseRangeAddress(address) {
  const raw = String(address || "A1").trim();
  const withoutSheet = raw.includes("!") ? raw.slice(raw.lastIndexOf("!") + 1) : raw;
  const [startRaw, endRaw = startRaw] = withoutSheet.split(":");
  const start = parseCellAddress(startRaw);
  const end = parseCellAddress(endRaw);
  const top = Math.min(start.row, end.row);
  const left = Math.min(start.col, end.col);
  const bottom = Math.max(start.row, end.row);
  const right = Math.max(start.col, end.col);
  return { top, left, bottom, right, rowCount: bottom - top + 1, colCount: right - left + 1 };
}

function rangeToAddress(bounds) {
  const start = makeCellAddress(bounds.top, bounds.left);
  const end = makeCellAddress(bounds.bottom, bounds.right);
  return start === end ? start : `${start}:${end}`;
}

function ndjson(records, maxChars = Infinity) {
  const lines = records.map((record) => JSON.stringify(record));
  let text = lines.join("\n");
  let truncated = false;
  if (text.length > maxChars) {
    truncated = true;
    const kept = [];
    let chars = 0;
    for (const line of lines) {
      if (chars + line.length + 1 > maxChars) break;
      kept.push(line);
      chars += line.length + 1;
    }
    kept.push(JSON.stringify({ kind: "notice", message: `Truncated: omitted ${lines.length - kept.length} lines. Increase maxChars or narrow query.` }));
    text = kept.join("\n");
  }
  return { ndjson: text, truncated };
}

function normalizeKinds(kind, fallback) {
  if (!kind) return new Set(fallback);
  return new Set(String(kind).split(",").map((k) => k.trim()).filter(Boolean));
}

function toUint8Array(data) {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  if (typeof data === "string") return encoder.encode(data);
  throw new Error("Unsupported binary payload");
}

export class FileBlob {
  constructor(data, options = {}) {
    this.bytes = toUint8Array(data ?? new Uint8Array());
    this.type = options.type || "application/octet-stream";
  }

  async arrayBuffer() {
    return this.bytes.buffer.slice(this.bytes.byteOffset, this.bytes.byteOffset + this.bytes.byteLength);
  }

  async text() {
    return decoder.decode(this.bytes);
  }

  async save(filePath) {
    await fs.mkdir(path.dirname(path.resolve(filePath)), { recursive: true });
    await fs.writeFile(filePath, this.bytes);
    return filePath;
  }

  static async load(filePath, options = {}) {
    const bytes = await fs.readFile(filePath);
    return new FileBlob(bytes, options);
  }
}

class WorksheetCollection {
  constructor(workbook) {
    this.workbook = workbook;
    this.items = [];
  }

  add(name = `Sheet${this.items.length + 1}`) {
    const worksheet = new Worksheet(this.workbook, name);
    this.items.push(worksheet);
    return worksheet;
  }

  getItem(nameOrIndex) {
    if (typeof nameOrIndex === "number") return this.items[nameOrIndex];
    return this.items.find((sheet) => sheet.name === nameOrIndex);
  }

  getOrAdd(name) {
    return this.getItem(name) ?? this.add(name);
  }

  getItemAt(index) {
    return this.items[index];
  }

  getActiveWorksheet() {
    if (!this.items[0]) throw new Error("Workbook has no worksheets. Add one first.");
    return this.items[0];
  }

  [Symbol.iterator]() {
    return this.items[Symbol.iterator]();
  }
}

class CellStore {
  constructor() {
    this.cells = new Map();
  }

  get(address) {
    if (!this.cells.has(address)) this.cells.set(address, { value: null, formula: null, style: {} });
    return this.cells.get(address);
  }

  entries() {
    return [...this.cells.entries()];
  }
}

export class Workbook {
  constructor() {
    this.id = aid("wb");
    this.worksheets = new WorksheetCollection(this);
    this.comments = { setSelf() {}, addThread: () => ({ id: aid("th"), replies: [] }) };
  }

  static create() {
    return new Workbook();
  }

  static async fromCSV(text, options = {}) {
    const workbook = Workbook.create();
    await workbook.fromCSV(text, options);
    return workbook;
  }

  async fromCSV(text, options = {}) {
    const sheet = this.worksheets.add(options.sheetName || `Sheet${this.worksheets.items.length + 1}`);
    const rows = String(text).trim().split(/\r?\n/).map((line) => line.split(",").map((cell) => cell.trim()));
    sheet.getRangeByIndexes(0, 0, rows.length, rows[0]?.length || 1).values = rows;
    return this;
  }

  recalculate() {
    for (const sheet of this.worksheets) sheet.recalculate();
  }

  inspect(options = {}) {
    this.recalculate();
    const kinds = normalizeKinds(options.kind, ["workbook", "sheet", "table", "formula"]);
    const records = [];
    if (kinds.has("workbook")) records.push({ kind: "workbook", id: this.id, sheets: this.worksheets.items.length });
    for (const sheet of this.worksheets) {
      if (kinds.has("sheet")) records.push({ kind: "sheet", id: sheet.id, name: sheet.name, rows: sheet.usedBounds().rowCount, cols: sheet.usedBounds().colCount });
      if (kinds.has("table") || kinds.has("region")) records.push(sheet.tableRecord(options));
      if (kinds.has("formula")) records.push(...sheet.formulaRecords(options));
      if (kinds.has("match")) records.push(...sheet.matchRecords(options));
    }
    return ndjson(records.filter(Boolean), options.maxChars ?? Infinity);
  }

  trace(reference, options = {}) {
    this.recalculate();
    const root = parseWorkbookReference(this, reference);
    const maxDepth = options.maxDepth ?? 8;
    const seen = new Set();
    const build = (sheet, address, depth) => {
      const key = `${sheet.name}!${address}`;
      const cell = sheet.store.get(address);
      const node = {
        kind: "trace",
        sheet: sheet.name,
        address,
        value: cell.value,
        formula: cell.formula || undefined,
        depth,
        circular: seen.has(key) || undefined,
        precedents: [],
      };
      if (!cell.formula || depth >= maxDepth || seen.has(key)) return node;
      seen.add(key);
      for (const ref of formulaReferences(cell.formula)) {
        const targetSheet = ref.sheetName ? this.worksheets.getItem(ref.sheetName) : sheet;
        if (!targetSheet) {
          node.precedents.push({ kind: "trace", sheet: ref.sheetName, address: ref.address, missing: true, depth: depth + 1, precedents: [] });
          continue;
        }
        node.precedents.push(build(targetSheet, ref.address, depth + 1));
      }
      seen.delete(key);
      return node;
    };
    const tree = build(root.sheet, root.address, 0);
    const flat = [];
    const visit = (node) => {
      flat.push({ kind: "trace", sheet: node.sheet, address: node.address, value: node.value, formula: node.formula, depth: node.depth, missing: node.missing, circular: node.circular, precedents: node.precedents.map((p) => `${p.sheet}!${p.address}`) });
      node.precedents.forEach(visit);
    };
    visit(tree);
    return { tree, ...ndjson(flat, options.maxChars ?? Infinity) };
  }

  resolve(id) {
    if (id === this.id) return this;
    for (const sheet of this.worksheets) if (sheet.id === id) return sheet;
    return undefined;
  }

  help(query = "*", options = {}) {
    const q = String(query).toLowerCase();
    const catalog = [
      { kind: "api", name: "Workbook.create", summary: "Create an empty workbook; add worksheets before editing." },
      { kind: "api", name: "SpreadsheetFile.importXlsx", summary: "Load an XLSX file into a Workbook facade." },
      { kind: "api", name: "SpreadsheetFile.exportXlsx", summary: "Serialize a Workbook facade to an XLSX FileBlob." },
      { kind: "api", name: "worksheet.getRange", summary: "Select an A1 range for values, formulas, formatting, merge, fill, and copy operations." },
      { kind: "api", name: "workbook.inspect", summary: "Emit bounded NDJSON records for workbook, sheets, tables, formulas, matches, and styles." },
      { kind: "api", name: "workbook.render", summary: "Return a lightweight SVG preview for a sheet or range in the current clean-room MVP." },
      { kind: "api", name: "workbook.trace", summary: "Return a formula precedent tree and bounded NDJSON trace for a target cell." },
      { kind: "formula", name: "fx.SUM", category: "math-trig", examples: ["=SUM(A1:A10)"] },
      { kind: "formula", name: "fx.PMT", category: "financial", examples: ["=PMT(rate,nper,pv)"], notes: ["Catalog entry only in MVP; full financial formula evaluation is roadmap."] },
    ];
    const records = catalog.filter((item) => q === "*" || item.name.toLowerCase().includes(q.replace("fx.", "")) || item.summary?.toLowerCase().includes(q));
    return ndjson(records, options.maxChars ?? Infinity);
  }

  async render(options = {}) {
    const sheet = this.worksheets.getItem(options.sheetName) || this.worksheets.getActiveWorksheet();
    return new FileBlob(sheet.toSvg(options), { type: "image/svg+xml" });
  }
}

export class Worksheet {
  constructor(workbook, name) {
    this.workbook = workbook;
    this.id = aid("ws");
    this.name = name;
    this.store = new CellStore();
    this.charts = [];
    this.shapes = [];
    this.images = [];
    this.tables = [];
    this.freezePanes = { freezeRows() {}, freezeColumns() {}, unfreeze() {} };
    this.showGridLines = true;
  }

  getRange(address) {
    return new Range(this, parseRangeAddress(address));
  }

  getRangeByIndexes(startRow, startCol, rowCount, colCount) {
    return new Range(this, { top: startRow, left: startCol, bottom: startRow + rowCount - 1, right: startCol + colCount - 1, rowCount, colCount });
  }

  getCell(row, col) {
    return this.getRangeByIndexes(row, col, 1, 1);
  }

  getUsedRange() {
    return new Range(this, this.usedBounds());
  }

  deleteAllDrawings() {
    this.charts = [];
    this.shapes = [];
    this.images = [];
  }

  usedBounds() {
    const coords = this.store.entries().map(([address]) => parseCellAddress(address));
    if (coords.length === 0) return { top: 0, left: 0, bottom: 0, right: 0, rowCount: 1, colCount: 1 };
    const top = Math.min(...coords.map((c) => c.row));
    const left = Math.min(...coords.map((c) => c.col));
    const bottom = Math.max(...coords.map((c) => c.row));
    const right = Math.max(...coords.map((c) => c.col));
    return { top, left, bottom, right, rowCount: bottom - top + 1, colCount: right - left + 1 };
  }

  recalculate() {
    for (const [address, cell] of this.store.entries()) {
      if (cell.formula) cell.value = evaluateFormula(this, cell.formula, address);
    }
  }

  tableRecord(options = {}) {
    const bounds = options.range ? parseRangeAddress(options.range) : this.usedBounds();
    const range = new Range(this, bounds);
    return { kind: "table", sheet: this.name, address: rangeToAddress(bounds), rows: bounds.rowCount, cols: bounds.colCount, values: range.values };
  }

  formulaRecords(options = {}) {
    const bounds = options.range ? parseRangeAddress(options.range) : this.usedBounds();
    const records = [];
    for (const [address, cell] of this.store.entries()) {
      const coord = parseCellAddress(address);
      if (!cell.formula) continue;
      if (coord.row < bounds.top || coord.row > bounds.bottom || coord.col < bounds.left || coord.col > bounds.right) continue;
      records.push({ kind: "formula", sheet: this.name, address, formula: cell.formula, value: cell.value });
    }
    return records;
  }

  matchRecords(options = {}) {
    const term = options.searchTerm || options.search || "";
    if (!term) return [];
    const regex = options.options?.useRegex ? new RegExp(term) : null;
    const records = [];
    for (const [address, cell] of this.store.entries()) {
      const hay = `${cell.value ?? ""}\n${cell.formula ?? ""}`;
      const matched = regex ? regex.test(hay) : hay.includes(term);
      if (matched) records.push({ kind: "match", sheet: this.name, address, value: cell.value, formula: cell.formula });
    }
    return records;
  }

  toSvg() {
    this.recalculate();
    const bounds = this.usedBounds();
    const cellW = 96;
    const cellH = 28;
    const width = Math.max(320, bounds.colCount * cellW + 80);
    const height = Math.max(180, bounds.rowCount * cellH + 80);
    const rows = [];
    for (let r = bounds.top; r <= bounds.bottom; r++) {
      for (let c = bounds.left; c <= bounds.right; c++) {
        const address = makeCellAddress(r, c);
        const value = this.store.get(address).value ?? "";
        const x = 40 + (c - bounds.left) * cellW;
        const y = 40 + (r - bounds.top) * cellH;
        rows.push(`<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" fill="white" stroke="#d0d7de"/>`);
        rows.push(`<text x="${x + 6}" y="${y + 18}" font-family="Arial" font-size="13" fill="#24292f">${xmlEscape(value)}</text>`);
      }
    }
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="#f6f8fa"/>${rows.join("")}</svg>`;
  }
}

export class Range {
  constructor(worksheet, bounds) {
    this.worksheet = worksheet;
    this.bounds = bounds;
  }

  get values() {
    const out = [];
    for (let r = this.bounds.top; r <= this.bounds.bottom; r++) {
      const row = [];
      for (let c = this.bounds.left; c <= this.bounds.right; c++) row.push(this.worksheet.store.get(makeCellAddress(r, c)).value);
      out.push(row);
    }
    return out;
  }

  set values(matrix) {
    this.writeMatrix(matrix, "value");
  }

  get formulas() {
    const out = [];
    for (let r = this.bounds.top; r <= this.bounds.bottom; r++) {
      const row = [];
      for (let c = this.bounds.left; c <= this.bounds.right; c++) row.push(this.worksheet.store.get(makeCellAddress(r, c)).formula);
      out.push(row);
    }
    return out;
  }

  set formulas(matrix) {
    this.writeMatrix(matrix, "formula");
  }

  writeMatrix(matrix, field) {
    const rows = Array.isArray(matrix) ? matrix : [[matrix]];
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < (rows[r]?.length ?? 0); c++) {
        const address = makeCellAddress(this.bounds.top + r, this.bounds.left + c);
        const cell = this.worksheet.store.get(address);
        cell[field] = rows[r][c];
      }
    }
    this.worksheet.recalculate();
  }

  clear() {
    for (let r = this.bounds.top; r <= this.bounds.bottom; r++) for (let c = this.bounds.left; c <= this.bounds.right; c++) this.worksheet.store.cells.delete(makeCellAddress(r, c));
  }

  merge() {}
  unmerge() {}
  fillDown() {}
  fillRight() {}
  getCell(row, col) { return this.worksheet.getCell(this.bounds.top + row, this.bounds.left + col); }
  getRange(address) { return this.worksheet.getRange(address); }
  getOffsetRange(rowOffset, colOffset) { return this.worksheet.getRangeByIndexes(this.bounds.top + rowOffset, this.bounds.left + colOffset, this.bounds.rowCount, this.bounds.colCount); }
  getResizedRange(rowCount, colCount) { return this.worksheet.getRangeByIndexes(this.bounds.top, this.bounds.left, rowCount, colCount); }
}

function parseWorkbookReference(workbook, reference) {
  const raw = String(reference || "").trim();
  const bang = raw.lastIndexOf("!");
  let sheetName;
  let address = raw;
  if (bang !== -1) {
    sheetName = raw.slice(0, bang).replace(/^'|'$/g, "");
    address = raw.slice(bang + 1);
  }
  const sheet = sheetName ? workbook.worksheets.getItem(sheetName) : workbook.worksheets.getActiveWorksheet();
  if (!sheet) throw new Error(`Unknown worksheet in trace reference: ${reference}`);
  return { sheet, address: address.replaceAll("$", "").toUpperCase() };
}

function formulaReferences(formula) {
  const raw = String(formula || "");
  const refs = [];
  const rangeRegex = /(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]+\$?\d+)\s*:\s*(\$?[A-Za-z]+\$?\d+)/g;
  const consumed = [];
  for (const match of raw.matchAll(rangeRegex)) {
    consumed.push(match.index);
    const sheetName = match[1] || match[2] || undefined;
    const start = parseCellAddress(match[3].replaceAll("$", ""));
    const end = parseCellAddress(match[4].replaceAll("$", ""));
    for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
      for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
        refs.push({ sheetName, address: makeCellAddress(row, col) });
      }
    }
  }
  const cellRegex = /(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]+\$?\d+)/g;
  for (const match of raw.matchAll(cellRegex)) {
    if (consumed.some((index) => match.index >= index && match.index <= index + 64)) continue;
    refs.push({ sheetName: match[1] || match[2] || undefined, address: match[3].replaceAll("$", "").toUpperCase() });
  }
  const seen = new Set();
  return refs.filter((ref) => {
    const key = `${ref.sheetName || ""}!${ref.address}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function evaluateFormula(sheet, formula) {
  const raw = String(formula || "").trim();
  if (!raw.startsWith("=")) return raw;
  const expr = raw.slice(1);
  const sumMatch = /^SUM\(([^)]+)\)$/i.exec(expr);
  if (sumMatch) {
    const range = sheet.getRange(sumMatch[1]);
    return range.values.flat().reduce((acc, value) => acc + (Number(value) || 0), 0);
  }
  const safe = expr.replace(/\b\$?[A-Za-z]+\$?\d+\b/g, (ref) => Number(sheet.store.get(ref.replaceAll("$", "").toUpperCase()).value) || 0);
  if (!/^[0-9+\-*/().\s]+$/.test(safe)) return `#NAME?`;
  try {
    return Function(`"use strict"; return (${safe});`)();
  } catch {
    return "#VALUE!";
  }
}

export class SpreadsheetFile {
  static async exportXlsx(workbook) {
    workbook.recalculate();
    const zip = new JSZip();
    zip.file("[Content_Types].xml", xlsxContentTypes(workbook.worksheets.items.length));
    zip.file("_rels/.rels", relsXml([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "xl/workbook.xml" }]));
    zip.file("xl/workbook.xml", workbookXml(workbook));
    zip.file("xl/_rels/workbook.xml.rels", workbookRelsXml(workbook.worksheets.items.length));
    zip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>`);
    workbook.worksheets.items.forEach((sheet, index) => zip.file(`xl/worksheets/sheet${index + 1}.xml`, worksheetXml(sheet)));
    const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return new FileBlob(bytes, { type: XLSX_MIME });
  }

  static async importXlsx(blobOrBuffer) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const workbook = Workbook.create();
    const workbookText = await zip.file("xl/workbook.xml")?.async("text");
    const sheetNames = [...String(workbookText || "").matchAll(/<sheet[^>]*name="([^"]+)"[^>]*sheetId="(\d+)"/g)].map((m) => ({ name: decodeXml(m[1]), index: Number(m[2]) }));
    for (const { name, index } of sheetNames.length ? sheetNames : [{ name: "Sheet1", index: 1 }]) {
      const sheet = workbook.worksheets.add(name);
      const xml = await zip.file(`xl/worksheets/sheet${index}.xml`)?.async("text");
      if (xml) parseWorksheetXml(sheet, xml);
    }
    workbook.recalculate();
    return workbook;
  }
}

function xlsxContentTypes(sheetCount) {
  const sheets = Array.from({ length: sheetCount }, (_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets}</Types>`;
}

function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map((rel) => `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${rel.target}"/>`).join("")}</Relationships>`;
}

function workbookXml(workbook) {
  const sheets = workbook.worksheets.items.map((sheet, i) => `<sheet name="${attrEscape(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets}</sheets></workbook>`;
}

function workbookRelsXml(sheetCount) {
  const rels = Array.from({ length: sheetCount }, (_, i) => ({ id: `rId${i + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", target: `worksheets/sheet${i + 1}.xml` }));
  rels.push({ id: `rId${sheetCount + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", target: "styles.xml" });
  return relsXml(rels);
}

function worksheetXml(sheet) {
  const rows = new Map();
  for (const [address, cell] of sheet.store.entries()) {
    const { row, col } = parseCellAddress(address);
    if (!rows.has(row)) rows.set(row, []);
    rows.get(row).push({ address, col, cell });
  }
  const rowXml = [...rows.entries()].sort((a, b) => a[0] - b[0]).map(([row, cells]) => `<row r="${row + 1}">${cells.sort((a, b) => a.col - b.col).map(({ address, cell }) => cellXml(address, cell)).join("")}</row>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rowXml}</sheetData></worksheet>`;
}

function cellXml(address, cell) {
  const f = cell.formula ? `<f>${xmlEscape(String(cell.formula).replace(/^=/, ""))}</f>` : "";
  if (typeof cell.value === "number") return `<c r="${address}">${f}<v>${cell.value}</v></c>`;
  if (typeof cell.value === "boolean") return `<c r="${address}" t="b">${f}<v>${cell.value ? 1 : 0}</v></c>`;
  if (cell.value == null && !f) return "";
  if (cell.value == null) return `<c r="${address}">${f}</c>`;
  return `<c r="${address}" t="inlineStr">${f}<is><t>${xmlEscape(cell.value)}</t></is></c>`;
}

function parseWorksheetXml(sheet, xml) {
  for (const match of xml.matchAll(/<c\s+([^>]*)>([\s\S]*?)<\/c>/g)) {
    const attrs = match[1];
    const body = match[2];
    const address = /r="([^"]+)"/.exec(attrs)?.[1];
    if (!address) continue;
    const cell = sheet.store.get(address);
    const formula = /<f[^>]*>([\s\S]*?)<\/f>/.exec(body)?.[1];
    if (formula) cell.formula = `=${decodeXml(formula)}`;
    const text = /<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/is>/.exec(body)?.[1];
    const value = /<v[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1];
    if (text !== undefined) cell.value = decodeXml(text);
    else if (value !== undefined) cell.value = Number.isFinite(Number(value)) ? Number(value) : decodeXml(value);
  }
}

class SlideCollection {
  constructor(presentation) {
    this.presentation = presentation;
    this.items = [];
  }

  add(options = {}) {
    const slide = new Slide(this.presentation, options);
    this.items.push(slide);
    return slide;
  }

  getItem(index) { return this.items[index]; }
  get count() { return this.items.length; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

export class Presentation {
  constructor(options = {}) {
    this.id = aid("pr");
    this.slideSize = options.slideSize || { width: 1280, height: 720 };
    this.slides = new SlideCollection(this);
  }

  static create(options = {}) { return new Presentation(options); }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["deck", "slide", "textbox", "shape", "layout"]);
    const records = [];
    if (kinds.has("deck")) records.push({ kind: "deck", id: this.id, slides: this.slides.count });
    for (const slide of this.slides) records.push(...slide.inspectRecords(kinds));
    return ndjson(records, options.maxChars ?? Infinity);
  }

  resolve(id) {
    if (id === this.id) return this;
    for (const slide of this.slides) {
      if (slide.id === id) return slide;
      const found = slide.resolve(id);
      if (found) return found;
    }
    return undefined;
  }

  help(query = "*", options = {}) {
    const q = String(query).toLowerCase();
    const catalog = [
      { kind: "api", name: "Presentation.create", summary: "Create a deck with a default or explicit slide size." },
      { kind: "api", name: "presentation.inspect", summary: "Emit NDJSON for deck, slides, textboxes, shapes, tables, charts, images, notes, and layout." },
      { kind: "api", name: "presentation.resolve", summary: "Map stable inspect anchor IDs back to editable facade objects." },
      { kind: "api", name: "presentation.export", summary: "Export a slide preview, deck montage, or layout JSON." },
      { kind: "api", name: "slide.shapes.add", summary: "Add a shape/textbox with geometry, position, fill, line, and text." },
      { kind: "api", name: "slide.compose", summary: "Roadmap: compose-first layout tree and JSX runtime compatibility." },
      { kind: "api", name: "slide.charts.add", summary: "Roadmap: native chart facade compatible with agent chart workflows." },
    ];
    const records = catalog.filter((item) => q === "*" || item.name.toLowerCase().includes(q) || item.summary.toLowerCase().includes(q));
    return ndjson(records, options.maxChars ?? Infinity);
  }

  async export(options = {}) {
    const slide = options.slide || this.slides.getItem(0) || this.slides.add();
    if (options.format === "layout") return slide.export({ format: "layout" });
    return slide.export(options);
  }

  toProto() {
    return { id: this.id, slideSize: this.slideSize, slides: this.slides.items.map((slide) => slide.toProto()) };
  }
}

class ShapeCollection {
  constructor(slide) { this.slide = slide; this.items = []; }
  add(config = {}) { const shape = new Shape(this.slide, config); this.items.push(shape); return shape; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

export class Slide {
  constructor(presentation, options = {}) {
    this.presentation = presentation;
    this.id = aid("sl");
    this.name = options.name || "";
    this.shapes = new ShapeCollection(this);
    this.images = [];
    this.tables = [];
    this.charts = [];
    this.speakerNotes = { text: "" };
    this.background = { fill: "white" };
  }

  get index() { return this.presentation.slides.items.indexOf(this); }
  get frame() { return { left: 0, top: 0, ...this.presentation.slideSize }; }

  inspectRecords(kinds) {
    const records = [];
    if (kinds.has("layout")) records.push({ kind: "layout", layoutId: `${this.id}/layout`, name: "Blank", type: "blank" });
    if (kinds.has("slide")) records.push({ kind: "slide", id: this.id, slide: this.index + 1, title: this.title(), textShapes: this.shapes.items.filter((s) => s.text.value).length });
    for (const shape of this.shapes) {
      if (kinds.has("textbox") && shape.text.value) records.push(shape.inspectRecord("textbox"));
      else if (kinds.has("shape")) records.push(shape.inspectRecord("shape"));
    }
    return records;
  }

  title() { return this.shapes.items.find((shape) => shape.text.value)?.text.value || ""; }
  resolve(id) { return this.shapes.items.find((shape) => shape.id === id); }

  async export(options = {}) {
    if (options.format === "layout") return new FileBlob(JSON.stringify(this.layoutJson(), null, 2), { type: LAYOUT_MIME });
    return new FileBlob(this.toSvg(), { type: "image/svg+xml" });
  }

  layoutJson() {
    return { schema: "open-office-artifact.layout/v1", unit: "px", slide: { id: this.id, slide: this.index + 1, frame: this.frame }, elements: this.shapes.items.map((shape) => shape.layoutJson()) };
  }

  toSvg() {
    const { width, height } = this.presentation.slideSize;
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="${xmlEscape(this.background.fill || "white")}"/>${this.shapes.items.map((shape) => shape.toSvg()).join("")}</svg>`;
  }

  toProto() { return { id: this.id, shapes: this.shapes.items.map((shape) => shape.layoutJson()) }; }
  compose() { throw new Error("slide.compose is planned but not implemented in the clean-room MVP yet."); }
}

class TextFrame {
  constructor(text = "") { this.value = text; this.style = {}; }
  set(text) { this.value = String(text ?? ""); }
  replace(search, replacement) { this.value = this.value.replace(search, replacement); }
  toString() { return this.value; }
}

export class Shape {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = aid("sh");
    this.geometry = config.geometry || "rect";
    this.name = config.name || "";
    this.position = config.position || { left: 0, top: 0, width: 160, height: 80 };
    this.fill = config.fill || "transparent";
    this.line = config.line || { fill: "#334155", width: 1 };
    this._text = new TextFrame(config.text || "");
  }

  get text() { return this._text; }
  set text(value) { this._text.set(value); }

  inspectRecord(kind = "shape") {
    const p = this.position;
    return { kind, id: this.id, slide: this.slide.index + 1, name: this.name || undefined, text: this.text.value || undefined, textPreview: this.text.value || undefined, textChars: this.text.value.length || undefined, textLines: this.text.value ? this.text.value.split(/\r?\n/).length : undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px" };
  }

  layoutJson() { return { kind: this.text.value ? "textbox" : "shape", id: this.id, name: this.name, geometry: this.geometry, frame: this.position, text: this.text.value, style: { fill: this.fill, line: this.line, text: this.text.style } }; }

  toSvg() {
    const p = this.position;
    const fill = typeof this.fill === "string" ? this.fill : this.fill?.color || "transparent";
    const stroke = this.line?.fill || "#334155";
    const sw = this.line?.width ?? 1;
    const rect = this.geometry === "ellipse"
      ? `<ellipse cx="${p.left + p.width / 2}" cy="${p.top + p.height / 2}" rx="${p.width / 2}" ry="${p.height / 2}" fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}"/>`
      : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" rx="12" fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}"/>`;
    const text = this.text.value ? `<text x="${p.left + 12}" y="${p.top + 36}" font-family="Arial" font-size="${this.text.style.fontSize || 24}" font-weight="${this.text.style.bold ? "700" : "400"}" fill="${xmlEscape(this.text.style.color || "#0f172a")}">${xmlEscape(this.text.value)}</text>` : "";
    return rect + text;
  }
}

export class PresentationFile {
  static async exportPptx(presentation) {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", pptxContentTypes(presentation.slides.count));
    zip.file("_rels/.rels", relsXml([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "ppt/presentation.xml" }]));
    zip.file("ppt/presentation.xml", presentationXml(presentation));
    zip.file("ppt/_rels/presentation.xml.rels", relsXml(presentation.slides.items.map((_, i) => ({ id: `rId${i + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", target: `slides/slide${i + 1}.xml` }))));
    presentation.slides.items.forEach((slide, i) => zip.file(`ppt/slides/slide${i + 1}.xml`, slideXml(slide)));
    const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return new FileBlob(bytes, { type: PPTX_MIME });
  }

  static async importPptx(blobOrBuffer) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const presentation = Presentation.create();
    const slideFiles = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort();
    for (const file of slideFiles) {
      const slide = presentation.slides.add();
      parseSlideXml(slide, await zip.file(file).async("text"));
    }
    return presentation;
  }
}

function pptxContentTypes(slideCount) {
  const slides = Array.from({ length: slideCount }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${slides}</Types>`;
}

function presentationXml(presentation) {
  const ids = presentation.slides.items.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="12192000" cy="6858000"/><p:sldIdLst>${ids}</p:sldIdLst></p:presentation>`;
}

function slideXml(slide) {
  const shapes = slide.shapes.items.map((shape, index) => {
    const p = shape.position;
    const x = Math.round(p.left * 9525), y = Math.round(p.top * 9525), cx = Math.round(p.width * 9525), cy = Math.round(p.height * 9525);
    return `<p:sp><p:nvSpPr><p:cNvPr id="${index + 2}" name="${attrEscape(shape.name || shape.id)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${attrEscape(shape.geometry === "textbox" ? "rect" : shape.geometry)}"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(shape.text.value)}</a:t></a:r></a:p></p:txBody></p:sp>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld></p:sld>`;
}

function parseSlideXml(slide, xml) {
  for (const match of xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const part = match[0];
    const name = decodeXml(/<p:cNvPr[^>]*name="([^"]*)"/.exec(part)?.[1] || "");
    const text = [...part.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decodeXml(m[1])).join("");
    const off = /<a:off[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/.exec(part);
    const ext = /<a:ext[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(part);
    const shape = slide.shapes.add({ name, position: { left: off ? Number(off[1]) / 9525 : 0, top: off ? Number(off[2]) / 9525 : 0, width: ext ? Number(ext[1]) / 9525 : 160, height: ext ? Number(ext[2]) / 9525 : 80 } });
    shape.text = text;
  }
}

export class DocumentModel {
  constructor(options = {}) {
    this.id = aid("doc");
    this.name = options.name || "New document";
    this.paragraphs = options.paragraphs || ["Start writing here..."];
  }

  static create(options = {}) { return new DocumentModel(options); }
  toProto() { return { id: this.id, name: this.name, paragraphs: this.paragraphs.map((text, index) => ({ id: `p/${index + 1}`, text })) }; }
  inspect(options = {}) { return ndjson(this.paragraphs.map((text, index) => ({ kind: "paragraph", id: `p/${index + 1}`, text, textChars: text.length })), options.maxChars ?? Infinity); }
  help(query = "*") { return ndjson([{ kind: "api", name: "DocumentModel.create", summary: "Create a document with paragraphs." }, { kind: "api", name: "DocumentFile.exportDocx", summary: "Export a DocumentModel to a minimal DOCX package." }]); }
}

export class DocumentFile {
  static async exportDocx(document) {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
    zip.file("_rels/.rels", relsXml([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "word/document.xml" }]));
    const body = document.paragraphs.map((p) => `<w:p><w:r><w:t>${xmlEscape(p)}</w:t></w:r></w:p>`).join("");
    zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`);
    return new FileBlob(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: DOCX_MIME });
  }

  static async importDocx(blobOrBuffer) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("word/document.xml")?.async("text");
    const paragraphs = [...String(xml || "").matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((m) => decodeXml([...m[0].matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join(""))).filter(Boolean);
    return DocumentModel.create({ paragraphs: paragraphs.length ? paragraphs : [""] });
  }
}

export class PdfArtifact {
  constructor(options = {}) {
    this.id = aid("pdf");
    this.pages = options.pages || [{ text: options.text || "" }];
  }

  static create(options = {}) { return new PdfArtifact(options); }
  inspect(options = {}) { return ndjson(this.pages.map((page, i) => ({ kind: "page", id: `pg/${i + 1}`, page: i + 1, textPreview: page.text.slice(0, 300), textChars: page.text.length })), options.maxChars ?? Infinity); }
  async render(options = {}) { return new FileBlob(pdfPageSvg(this.pages[options.pageIndex || 0]?.text || ""), { type: "image/svg+xml" }); }
}

export class PdfFile {
  static async exportPdf(artifact) {
    return new FileBlob(buildMinimalPdf(artifact.pages.map((page) => page.text).join("\n\n")), { type: PDF_MIME });
  }

  static async importPdf(blobOrBuffer) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const text = decoder.decode(bytes);
    const strings = [...text.matchAll(/\(([^()]*)\)\s*Tj/g)].map((m) => m[1].replaceAll("\\)", ")").replaceAll("\\(", "("));
    return PdfArtifact.create({ pages: [{ text: strings.join("\n") }] });
  }
}

function pdfPageSvg(text) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="612" height="792" viewBox="0 0 612 792"><rect width="100%" height="100%" fill="white"/><text x="72" y="96" font-family="Helvetica" font-size="14" fill="#111827">${xmlEscape(text)}</text></svg>`;
}

function escapePdfString(text) {
  return String(text).replaceAll("\\", "\\\\").replaceAll("(", "\\(").replaceAll(")", "\\)");
}

function buildMinimalPdf(text) {
  const lines = String(text).split(/\r?\n/).filter(Boolean);
  const content = `BT\n/F1 18 Tf\n72 720 Td\n${lines.map((line, index) => `${index === 0 ? "" : "0 -24 Td\n"}(${escapePdfString(line)}) Tj`).join("\n")}\nET`;
  const objects = [
    `1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n`,
    `2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n`,
    `3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>\nendobj\n`,
    `4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n`,
    `5 0 obj\n<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}\nendstream\nendobj\n`,
  ];
  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  for (const obj of objects) { offsets.push(Buffer.byteLength(pdf)); pdf += obj; }
  const xref = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (let i = 1; i < offsets.length; i++) pdf += `${String(offsets[i]).padStart(10, "0")} 00000 n \n`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return encoder.encode(pdf);
}
