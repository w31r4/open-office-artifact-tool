import fs from "node:fs/promises";
import path from "node:path";
import { inflateSync } from "node:zlib";
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

function inspectTargetTokens(options = {}) {
  const raw = options.target ?? options.targetId ?? options.id ?? options.anchor;
  if (raw == null || raw === "") return [];
  const values = Array.isArray(raw) ? raw : [raw];
  const out = [];
  for (const value of values) {
    if (value == null) continue;
    if (typeof value === "object") {
      out.push(value.id, value.targetId, value.name, value.address, value.range, value.ref, value.sheetName && value.address ? `${value.sheetName}!${value.address}` : undefined);
    } else {
      out.push(...String(value).split(","));
    }
  }
  return out.map((value) => String(value ?? "").trim()).filter(Boolean);
}

function inspectRecordMatchesTarget(record, targets) {
  if (!targets.length) return true;
  if (!record) return false;
  const candidates = new Set();
  const add = (value) => { if (value != null && value !== "") candidates.add(String(value)); };
  for (const key of ["id", "targetId", "parentId", "layoutId", "name", "address", "range", "sheet", "slide", "page", "kind", "drawingType", "regionKind"]) add(record[key]);
  if (record.sheet && record.address) add(`${record.sheet}!${record.address}`);
  if (record.sheet && record.range) add(`${record.sheet}!${record.range}`);
  if (record.target) {
    add(record.target.id);
    add(record.target.address);
    add(record.target.range);
    if (record.target.sheetName && record.target.address) add(`${record.target.sheetName}!${record.target.address}`);
  }
  const haystack = JSON.stringify(record);
  return targets.some((target) => candidates.has(target) || haystack.includes(target));
}

function filterInspectRecords(records, options = {}) {
  const search = String(options.search || options.searchTerm || "").trim().toLowerCase();
  const targets = inspectTargetTokens(options);
  let filtered = records
    .filter(Boolean)
    .filter((record) => !search || JSON.stringify(record).toLowerCase().includes(search));
  if (targets.length) {
    const before = Math.max(0, Number(options.before ?? options.contextBefore ?? options.context ?? 0) || 0);
    const after = Math.max(0, Number(options.after ?? options.contextAfter ?? options.context ?? 0) || 0);
    if (before || after) {
      const keep = new Set();
      filtered.forEach((record, index) => {
        if (!inspectRecordMatchesTarget(record, targets)) return;
        for (let i = Math.max(0, index - before); i <= Math.min(filtered.length - 1, index + after); i += 1) keep.add(i);
      });
      filtered = filtered.filter((_, index) => keep.has(index));
    } else {
      filtered = filtered.filter((record) => inspectRecordMatchesTarget(record, targets));
    }
  }
  return shapeInspectRecords(filtered, options);
}

const INSPECT_CORE_FIELDS = new Set(["kind", "id", "sheet", "address", "range", "name", "page", "slide", "targetId", "parentId"]);
const INSPECT_FIELD_ALIASES = {
  values: ["values", "value"],
  value: ["value", "values"],
  formulas: ["formulas", "formula"],
  formula: ["formula", "formulas"],
  bbox: ["bbox", "bboxUnit"],
  text: ["text", "textPreview", "textChars"],
  style: ["style", "styleId"],
};

function inspectFieldList(value) {
  if (value == null || value === "") return [];
  const values = Array.isArray(value) ? value : String(value).split(",");
  return values.map((item) => String(item || "").trim()).filter(Boolean);
}

function expandedInspectFields(fields) {
  const out = new Set();
  for (const field of fields) {
    out.add(field);
    for (const alias of INSPECT_FIELD_ALIASES[field] || []) out.add(alias);
  }
  return out;
}

function shapeInspectRecord(record, options = {}) {
  const includeFields = expandedInspectFields(inspectFieldList(options.fields ?? options.includeFields ?? options.include));
  const excludeFields = expandedInspectFields(inspectFieldList(options.exclude ?? options.omit));
  if (!includeFields.size && !excludeFields.size) return record;
  const shaped = {};
  for (const [key, value] of Object.entries(record)) {
    const keepByInclude = !includeFields.size || includeFields.has(key) || INSPECT_CORE_FIELDS.has(key);
    const dropByExclude = excludeFields.has(key) && !INSPECT_CORE_FIELDS.has(key);
    if (keepByInclude && !dropByExclude) shaped[key] = value;
  }
  return shaped;
}

function shapeInspectRecords(records, options = {}) {
  return records.map((record) => shapeInspectRecord(record, options));
}

function verificationResult(artifactKind, issues, options = {}) {
  const result = {
    artifactKind,
    ok: issues.length === 0,
    issues,
    ...ndjson(issues, options.maxChars ?? Infinity),
  };
  return result;
}

function verificationIssue(artifactKind, type, message, details = {}) {
  return { kind: "verificationIssue", artifactKind, type, severity: details.severity || "error", message, ...details };
}

function inferArtifactKind(artifact) {
  if (artifact instanceof Workbook) return "workbook";
  if (artifact instanceof Presentation) return "presentation";
  if (artifact instanceof DocumentModel) return "document";
  if (artifact instanceof PdfArtifact) return "pdf";
  return "unknown";
}

export function verifyArtifact(artifact, options = {}) {
  if (!artifact || typeof artifact.verify !== "function") {
    return verificationResult("unknown", [verificationIssue("unknown", "unsupportedArtifact", "Artifact does not expose a verify() method.")], options);
  }
  return artifact.verify(options);
}

const RENDER_MIME_BY_FORMAT = {
  svg: "image/svg+xml",
  png: "image/png",
  webp: "image/webp",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  pdf: "application/pdf",
  layout: LAYOUT_MIME,
};

function renderTypeForOptions(options = {}, fallbackType = "application/octet-stream") {
  const format = String(options.format || "").trim().toLowerCase();
  if (!format) return fallbackType;
  return RENDER_MIME_BY_FORMAT[format] || (format.includes("/") ? format : fallbackType);
}

async function fileBlobFromRenderOutput(output, type, metadata = {}) {
  if (output instanceof FileBlob) {
    output.metadata = { ...(output.metadata || {}), ...metadata };
    return output;
  }
  if (output?.data !== undefined) return fileBlobFromRenderOutput(output.data, output.type || type, { ...metadata, ...(output.metadata || {}) });
  if (output?.arrayBuffer) return new FileBlob(new Uint8Array(await output.arrayBuffer()), { type: output.type || type, metadata });
  return new FileBlob(output instanceof Uint8Array || output instanceof ArrayBuffer || ArrayBuffer.isView(output) ? toUint8Array(output) : String(output ?? ""), { type, metadata });
}

function attachRenderMetadata(blob, artifactKind, options = {}, format = options.format || blob.type) {
  blob.metadata = {
    ...(blob.metadata || {}),
    artifactKind,
    format,
    page: options.page,
    pageIndex: options.pageIndex,
    slide: options.slide,
    sheetName: options.sheetName,
    range: options.range,
  };
  return blob;
}

function stableByteHash(bytes) {
  let hash = 2166136261;
  for (const byte of bytes || []) {
    hash ^= byte;
    hash = Math.imul(hash, 16777619) >>> 0;
  }
  return hash.toString(16).padStart(8, "0");
}

async function bytesForVisualBaseline(input) {
  if (input == null) return undefined;
  if (input instanceof FileBlob) return input.bytes;
  if (typeof input === "string") return new TextEncoder().encode(input);
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  if (typeof input.arrayBuffer === "function") return new Uint8Array(await input.arrayBuffer());
  return undefined;
}

function svgDimensions(svgText = "") {
  const open = String(svgText || "").match(/<svg\b[^>]*>/i)?.[0] || "";
  const width = Number.parseFloat(/\bwidth=["']([^"']+)/i.exec(open)?.[1] || "");
  const height = Number.parseFloat(/\bheight=["']([^"']+)/i.exec(open)?.[1] || "");
  const viewBox = /\bviewBox=["']([^"']+)["']/i.exec(open)?.[1]?.trim().split(/[\s,]+/).map(Number);
  return {
    width: Number.isFinite(width) ? width : Number.isFinite(viewBox?.[2]) ? viewBox[2] : undefined,
    height: Number.isFinite(height) ? height : Number.isFinite(viewBox?.[3]) ? viewBox[3] : undefined,
  };
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

function isPngBytes(bytes) {
  return PNG_SIGNATURE.every((byte, index) => bytes?.[index] === byte);
}

function paethPredictor(a, b, c) {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  if (pb <= pc) return b;
  return c;
}

function decodePngRgba(bytes) {
  if (!isPngBytes(bytes)) throw new Error("not a PNG file");
  let offset = 8;
  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  let compression = 0;
  let filterMethod = 0;
  let interlace = 0;
  const idat = [];
  while (offset + 12 <= bytes.byteLength) {
    const length = (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];
    const type = decoder.decode(bytes.slice(offset + 4, offset + 8));
    const dataStart = offset + 8;
    const dataEnd = dataStart + length;
    if (length < 0 || dataEnd + 4 > bytes.byteLength) throw new Error("truncated PNG chunk");
    const data = bytes.slice(dataStart, dataEnd);
    if (type === "IHDR") {
      width = (data[0] << 24) | (data[1] << 16) | (data[2] << 8) | data[3];
      height = (data[4] << 24) | (data[5] << 16) | (data[6] << 8) | data[7];
      bitDepth = data[8];
      colorType = data[9];
      compression = data[10];
      filterMethod = data[11];
      interlace = data[12];
    } else if (type === "IDAT") {
      idat.push(Buffer.from(data));
    } else if (type === "IEND") {
      break;
    }
    offset = dataEnd + 4;
  }
  if (!width || !height) throw new Error("PNG is missing IHDR geometry");
  if (bitDepth !== 8 || compression !== 0 || filterMethod !== 0 || interlace !== 0) throw new Error("only 8-bit non-interlaced PNGs are supported for pixel diff");
  const channels = { 0: 1, 2: 3, 4: 2, 6: 4 }[colorType];
  if (!channels) throw new Error(`unsupported PNG color type ${colorType}`);
  const rowBytes = width * channels;
  const inflated = inflateSync(Buffer.concat(idat));
  const expected = (rowBytes + 1) * height;
  if (inflated.byteLength < expected) throw new Error("PNG image data is truncated");
  const raw = new Uint8Array(width * height * channels);
  let inOffset = 0;
  for (let y = 0; y < height; y += 1) {
    const filter = inflated[inOffset++];
    const rowStart = y * rowBytes;
    const priorStart = (y - 1) * rowBytes;
    for (let x = 0; x < rowBytes; x += 1) {
      const left = x >= channels ? raw[rowStart + x - channels] : 0;
      const up = y > 0 ? raw[priorStart + x] : 0;
      const upLeft = y > 0 && x >= channels ? raw[priorStart + x - channels] : 0;
      const value = inflated[inOffset++];
      if (filter === 0) raw[rowStart + x] = value;
      else if (filter === 1) raw[rowStart + x] = (value + left) & 0xff;
      else if (filter === 2) raw[rowStart + x] = (value + up) & 0xff;
      else if (filter === 3) raw[rowStart + x] = (value + Math.floor((left + up) / 2)) & 0xff;
      else if (filter === 4) raw[rowStart + x] = (value + paethPredictor(left, up, upLeft)) & 0xff;
      else throw new Error(`unsupported PNG row filter ${filter}`);
    }
  }
  const rgba = new Uint8Array(width * height * 4);
  for (let i = 0, p = 0; p < width * height; p += 1, i += channels) {
    const out = p * 4;
    if (colorType === 0) {
      rgba[out] = raw[i];
      rgba[out + 1] = raw[i];
      rgba[out + 2] = raw[i];
      rgba[out + 3] = 255;
    } else if (colorType === 2) {
      rgba[out] = raw[i];
      rgba[out + 1] = raw[i + 1];
      rgba[out + 2] = raw[i + 2];
      rgba[out + 3] = 255;
    } else if (colorType === 4) {
      rgba[out] = raw[i];
      rgba[out + 1] = raw[i];
      rgba[out + 2] = raw[i];
      rgba[out + 3] = raw[i + 1];
    } else {
      rgba[out] = raw[i];
      rgba[out + 1] = raw[i + 1];
      rgba[out + 2] = raw[i + 2];
      rgba[out + 3] = raw[i + 3];
    }
  }
  return { width, height, pixels: rgba };
}

function comparePngPixels(bytes, baselineBytes, options = {}) {
  const threshold = Math.max(0, Number(options.pixelThreshold ?? options.threshold ?? 0));
  const actual = decodePngRgba(bytes);
  const expected = decodePngRgba(baselineBytes);
  const result = {
    format: "png",
    width: actual.width,
    height: actual.height,
    baselineWidth: expected.width,
    baselineHeight: expected.height,
    threshold,
    pixels: actual.width * actual.height,
    differentPixels: 0,
    mismatchRatio: 0,
    maxChannelDelta: 0,
    meanChannelDelta: 0,
  };
  if (actual.width !== expected.width || actual.height !== expected.height) {
    result.dimensionMismatch = true;
    result.differentPixels = Math.max(result.pixels, expected.width * expected.height);
    result.mismatchRatio = 1;
    result.changed = true;
    return result;
  }
  let changedPixels = 0;
  let channelDeltaSum = 0;
  for (let i = 0; i < actual.pixels.length; i += 4) {
    let pixelChanged = false;
    for (let c = 0; c < 4; c += 1) {
      const delta = Math.abs(actual.pixels[i + c] - expected.pixels[i + c]);
      channelDeltaSum += delta;
      if (delta > result.maxChannelDelta) result.maxChannelDelta = delta;
      if (delta > threshold) pixelChanged = true;
    }
    if (pixelChanged) changedPixels += 1;
  }
  result.differentPixels = changedPixels;
  result.mismatchRatio = result.pixels ? changedPixels / result.pixels : 0;
  result.meanChannelDelta = actual.pixels.length ? channelDeltaSum / actual.pixels.length : 0;
  result.changed = changedPixels > 0;
  return result;
}

export async function renderArtifact(artifact, options = {}) {
  const artifactKind = inferArtifactKind(artifact);
  if (!artifact || (typeof artifact.render !== "function" && typeof artifact.export !== "function")) {
    throw new Error("Artifact does not expose a render() or export() method.");
  }
  const renderer = typeof artifact.render === "function" ? artifact.render.bind(artifact) : artifact.export.bind(artifact);
  let blob = await renderer(options);
  if (!(blob instanceof FileBlob)) {
    blob = await fileBlobFromRenderOutput(blob, blob?.type || "application/octet-stream", { artifactKind, format: options.format || blob?.type || "unknown" });
  }
  const desiredType = renderTypeForOptions(options, blob.type);
  const wantsConversion = options.format && desiredType !== blob.type;
  if (wantsConversion) {
    const adapter = options.renderer || options.rasterRenderer || options.renderAdapter;
    if (typeof adapter !== "function") {
      throw new Error(`renderArtifact requested ${options.format} output, but no renderer adapter was provided.`);
    }
    const converted = await adapter({ input: blob, source: blob, inputType: blob.type, outputType: desiredType, format: options.format, artifactKind, options });
    blob = await fileBlobFromRenderOutput(converted, desiredType, { artifactKind, format: options.format, renderedFrom: blob.type });
    if (!blob.type || blob.type === "application/octet-stream") blob.type = desiredType;
  }
  return attachRenderMetadata(blob, artifactKind, options, options.format || blob.metadata?.format || blob.type);
}

export async function visualQaArtifact(artifact, options = {}) {
  const blob = await renderArtifact(artifact, options);
  const artifactKind = blob.metadata?.artifactKind || inferArtifactKind(artifact);
  const bytes = blob.bytes || new Uint8Array(await blob.arrayBuffer());
  const hash = stableByteHash(bytes);
  const issues = [];
  const summary = { kind: "visualQa", artifactKind, type: blob.type, format: blob.metadata?.format || options.format || blob.type, bytes: bytes.byteLength, hash };
  if (bytes.byteLength === 0) issues.push(verificationIssue(artifactKind, "emptyRender", "Rendered artifact is empty.", { severity: "error", type: blob.type }));
  if (options.minBytes != null && bytes.byteLength < Number(options.minBytes)) issues.push(verificationIssue(artifactKind, "renderTooSmall", `Rendered artifact has ${bytes.byteLength} bytes; expected at least ${options.minBytes}.`, { severity: "warning", bytes: bytes.byteLength, minBytes: Number(options.minBytes) }));
  if (options.maxBytes != null && bytes.byteLength > Number(options.maxBytes)) issues.push(verificationIssue(artifactKind, "renderTooLarge", `Rendered artifact has ${bytes.byteLength} bytes; expected at most ${options.maxBytes}.`, { severity: "warning", bytes: bytes.byteLength, maxBytes: Number(options.maxBytes) }));
  if (blob.type === "image/svg+xml") {
    const text = await blob.text();
    const dimensions = svgDimensions(text);
    summary.width = dimensions.width;
    summary.height = dimensions.height;
    if (!dimensions.width || !dimensions.height || dimensions.width <= 0 || dimensions.height <= 0) issues.push(verificationIssue(artifactKind, "invalidRenderGeometry", "SVG render is missing positive width/height geometry.", { severity: "error", dimensions }));
    if (!/<(text|image|rect|path|line|polyline|polygon|circle|ellipse)\b/i.test(text)) issues.push(verificationIssue(artifactKind, "blankSvgRender", "SVG render has no recognizable visible elements.", { severity: "warning" }));
  }
  const baselineBytes = await bytesForVisualBaseline(options.baseline || options.expected || options.baselineBlob);
  if (baselineBytes) {
    const baselineHash = stableByteHash(baselineBytes);
    summary.baselineHash = baselineHash;
    summary.changed = baselineHash !== hash;
    const pixelDiffEnabled = options.pixelDiff === true || typeof options.pixelDiff === "object";
    if (pixelDiffEnabled) {
      if (isPngBytes(bytes) && isPngBytes(baselineBytes)) {
        try {
          const pixelDiff = comparePngPixels(bytes, baselineBytes, typeof options.pixelDiff === "object" ? options.pixelDiff : options);
          summary.pixelDiff = pixelDiff;
          if (pixelDiff.changed && options.allowChange !== true && options.allowPixelChange !== true) {
            issues.push(verificationIssue(artifactKind, "visualPixelDiff", `Rendered PNG differs from the baseline in ${pixelDiff.differentPixels} pixels.`, { severity: options.diffSeverity || "warning", ...pixelDiff }));
          }
        } catch (error) {
          summary.pixelDiff = { skipped: true, reason: error.message };
        }
      } else {
        summary.pixelDiff = { skipped: true, reason: "pixelDiff currently supports PNG baselines only" };
      }
    }
    if (baselineHash !== hash && options.allowChange !== true) issues.push(verificationIssue(artifactKind, "visualDiff", "Rendered output differs from the supplied baseline.", { severity: options.diffSeverity || "warning", hash, baselineHash }));
  }
  const records = [summary, ...issues];
  return { artifactKind, ok: issues.length === 0, blob, summary, issues, ...ndjson(records, options.maxChars ?? Infinity) };
}

export const HELP_CATALOG = [
  { artifactKind: "workbook", kind: "api", name: "Workbook.create", summary: "Create an empty workbook; add worksheets before editing." },
  { artifactKind: "workbook", kind: "api", name: "SpreadsheetFile.importXlsx", summary: "Load an XLSX file into a Workbook facade." },
  { artifactKind: "workbook", kind: "api", name: "SpreadsheetFile.exportXlsx", summary: "Serialize a Workbook facade to an XLSX FileBlob." },
  { artifactKind: "workbook", kind: "api", name: "worksheet.getRange", summary: "Select an A1 range for values, formulas, formatting, merge, fill, and copy operations." },
  { artifactKind: "workbook", kind: "api", name: "workbook.inspect", summary: "Emit bounded NDJSON records for workbook, sheets, tables, formulas, matches, comments, validations, conditional formats, and drawings; narrow with search/target anchors and shape fields with include/exclude." },
  { artifactKind: "workbook", kind: "api", name: "workbook.render", summary: "Return a lightweight SVG preview for a sheet/range or layout JSON when called with { format: 'layout' }." },
  { artifactKind: "workbook", kind: "api", name: "workbook.layoutJson", summary: "Return workbook/worksheet layout JSON with cell, table, chart, image, sparkline, and rule bounding boxes in pixels." },
  { artifactKind: "workbook", kind: "api", name: "workbook.verify", summary: "Return bounded QA issues for sheets, formulas, tables, charts, and comments." },
  { artifactKind: "workbook", kind: "api", name: "workbook.trace", summary: "Return a formula precedent tree and bounded NDJSON trace for a target cell, with circular references flagged." },
  { artifactKind: "workbook", kind: "api", name: "workbook.formulaGraph", summary: "Return a dependency graph of formula nodes, edges, dependents, cycles, and formula errors for workbook QA." },
  { artifactKind: "workbook", kind: "formula", name: "workbook.structuredReferences", summary: "Evaluate a clean-room subset of Excel structured references such as TableName[Column] in formulas, expanding them to table data-body cell precedents." },
  { artifactKind: "workbook", kind: "api", name: "workbook.definedNames.add", summary: "Create a workbook or sheet-scoped defined name over an A1 range; exported as native workbook.xml definedName and usable in formulas such as SUM(RevenueData)." },
  { artifactKind: "workbook", kind: "api", name: "range.dataValidation", summary: "Assign a validation rule to a range or use sheet.dataValidations.add({ range, rule })." },
  { artifactKind: "workbook", kind: "api", name: "range.format", summary: "Assign basic cell style metadata such as fill, font, and numberFormat; XLSX export writes native styles.xml and cell style indexes." },
  { artifactKind: "workbook", kind: "api", name: "range.conditionalFormats.add", summary: "Add a conditional formatting rule to a range; addCustom(expression, format) creates expression rules." },
  { artifactKind: "workbook", kind: "api", name: "workbook.comments.addThread", summary: "Create threaded comments after comments.setSelf({ displayName }); resolve with wb.resolve('th/...')." },
  { artifactKind: "workbook", kind: "api", name: "sheet.tables.add", summary: "Create an inspectable worksheet table over an A1 range with rows.add, getDataRows, getHeaderRowRange, style, and visibility toggles." },
  { artifactKind: "workbook", kind: "api", name: "sheet.charts.add", summary: "Create an inspectable worksheet chart from a range or config; setData(range) infers categories and series formulas." },
  { artifactKind: "workbook", kind: "api", name: "sheet.images.add", summary: "Create an inspectable worksheet image placeholder from a data URL, URI, or prompt with 0-based cell anchors and pixel extents." },
  { artifactKind: "workbook", kind: "api", name: "sheet.sparklineGroups.add", summary: "Create line/column/stacked sparklines from sourceData into a targetRange; range.sparklines.add is a shorthand." },
  { artifactKind: "workbook", kind: "formula", name: "fx.SUM", category: "math-trig", summary: "Sum numeric values across arguments and ranges.", examples: ["=SUM(A1:A10)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.AVERAGE", category: "statistical", summary: "Average numeric values across arguments and ranges in the clean-room formula engine.", examples: ["=AVERAGE(A1:A10)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.MIN", category: "statistical", summary: "Return the minimum numeric value across arguments and ranges.", examples: ["=MIN(A1:A10)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.MAX", category: "statistical", summary: "Return the maximum numeric value across arguments and ranges.", examples: ["=MAX(A1:A10)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.COUNT", category: "statistical", summary: "Count numeric values across arguments and ranges.", examples: ["=COUNT(A1:A10)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.IF", category: "logical", summary: "Return one value when a condition is true and another when false.", examples: ["=IF(A1>0,\"ok\",\"bad\")"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.AND", category: "logical", summary: "Return TRUE when all conditions are true.", examples: ["=AND(A1>0,B1>0)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.OR", category: "logical", summary: "Return TRUE when any condition is true.", examples: ["=OR(A1>0,B1>0)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.ROUND", category: "math-trig", summary: "Round a numeric value to a fixed number of decimal places.", examples: ["=ROUND(A1,2)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.COUNTIF", category: "statistical", summary: "Count values in a range that match a criterion.", examples: ["=COUNTIF(A1:A10,\">0\")"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.SUMIF", category: "math-trig", summary: "Sum values whose corresponding criteria range entries match a criterion.", examples: ["=SUMIF(A1:A10,\"East\",B1:B10)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.VLOOKUP", category: "lookup-reference", summary: "Look up a value in the first column of a table range and return a value from another column.", examples: ["=VLOOKUP(\"Beta\",A2:B4,2,FALSE)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.XLOOKUP", category: "lookup-reference", summary: "Look up a value in one range and return the corresponding value from another range.", examples: ["=XLOOKUP(\"Gamma\",A2:A4,B2:B4,\"missing\")"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.TEXTJOIN", category: "text", summary: "Join text values with a delimiter and optional empty-value skipping.", examples: ["=TEXTJOIN(\"/\",TRUE,A1:A3)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.CONCAT", category: "text", summary: "Concatenate text values and ranges.", examples: ["=CONCAT(A1,\"-\",B1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.LEFT", category: "text", summary: "Return characters from the start of a text value.", examples: ["=LEFT(A1,3)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.RIGHT", category: "text", summary: "Return characters from the end of a text value.", examples: ["=RIGHT(A1,3)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.LEN", category: "text", summary: "Return the length of a text value.", examples: ["=LEN(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.MID", category: "text", summary: "Return characters from the middle of a text value.", examples: ["=MID(A1,2,3)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.UPPER", category: "text", summary: "Convert text to uppercase.", examples: ["=UPPER(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.LOWER", category: "text", summary: "Convert text to lowercase.", examples: ["=LOWER(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.TRIM", category: "text", summary: "Trim leading/trailing whitespace and collapse internal whitespace.", examples: ["=TRIM(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.ABS", category: "math-trig", summary: "Return the absolute value of a number.", examples: ["=ABS(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.INT", category: "math-trig", summary: "Round a number down to the nearest integer.", examples: ["=INT(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.CEILING", category: "math-trig", summary: "Round a number up to the nearest significance.", examples: ["=CEILING(A1,5)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.FLOOR", category: "math-trig", summary: "Round a number down to the nearest significance.", examples: ["=FLOOR(A1,5)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.PMT", category: "financial", summary: "Calculate a loan payment for constant payments and constant interest rate.", examples: ["=PMT(rate,nper,pv)"], notes: ["Catalog entry only in MVP; full financial formula evaluation is roadmap."] },

  { artifactKind: "presentation", kind: "api", name: "Presentation.create", summary: "Create a deck with a default or explicit slide size." },
  { artifactKind: "presentation", kind: "api", name: "presentation.inspect", summary: "Emit NDJSON for deck, slides, textboxes, shapes, tables, charts, images, notes, comments, and layout; narrow with search/target anchors and shape fields with include/exclude." },
  { artifactKind: "presentation", kind: "api", name: "presentation.resolve", summary: "Map stable inspect anchor IDs back to editable facade objects." },
  { artifactKind: "presentation", kind: "api", name: "presentation.export", summary: "Export a slide preview, deck montage, or layout JSON." },
  { artifactKind: "presentation", kind: "api", name: "presentation.validateLayout", summary: "Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow." },
  { artifactKind: "presentation", kind: "api", name: "presentation.verify", summary: "Return presentation QA issues for layout validation, placeholder/template fidelity, chart/data consistency, table shape, image data, and dangling comments." },
  { artifactKind: "presentation", kind: "api", name: "slide.shapes.add", summary: "Add a shape/textbox with geometry, position, fill, line, and text." },
  { artifactKind: "presentation", kind: "api", name: "slide.compose", summary: "Materialize a clean-room compose tree with row, column, grid, layers, box, paragraph, shape, table, chart, image, and rule nodes into editable slide objects." },
  { artifactKind: "presentation", kind: "api", name: "slide.autoLayout", summary: "Place existing shapes inside a frame using horizontal or vertical flow, gap, padding, and alignment options." },
  { artifactKind: "presentation", kind: "api", name: "slide.tables.add", summary: "Add an inspectable native-style table facade with rows, columns, values, cells, layout JSON, and SVG/PPTX placeholder output." },
  { artifactKind: "presentation", kind: "api", name: "slide.charts.add", summary: "Add an inspectable chart facade with chartType, title, categories, series, layout JSON, SVG preview, and PPTX placeholder output." },
  { artifactKind: "presentation", kind: "api", name: "slide.images.add", summary: "Add an inspectable image facade with alt text, prompt/URI/data URL metadata, fit, frame, layout JSON, SVG preview, and PPTX placeholder output." },
  { artifactKind: "presentation", kind: "api", name: "presentation.theme", summary: "Configure inspectable theme colors and major/minor fonts; export writes a real ppt/theme/theme1.xml part." },
  { artifactKind: "presentation", kind: "api", name: "presentation.layouts.add", summary: "Create a reusable slide layout with placeholders; export writes slideLayout and slideMaster parts for clean-room PPTX roundtrip." },
  { artifactKind: "presentation", kind: "api", name: "slide.applyLayout", summary: "Apply a slide layout to materialize editable placeholder shapes and preserve layout identity for inspect, verify, and PPTX export." },
  { artifactKind: "presentation", kind: "api", name: "slide.addNotes", summary: "Set speaker notes for a slide; exported as a PPTX notesSlide part and surfaced through inspect({ kind: 'notes' })." },
  { artifactKind: "presentation", kind: "api", name: "slide.comments.addThread", summary: "Attach threaded comments to slide elements; exported as PPTX comments parts and verified for dangling targets." },
  { artifactKind: "presentation", kind: "api", name: "slide.connectors.add", summary: "Add an inspectable connector line between points or element IDs with SVG preview, layout JSON, PPTX p:cxnSp export, and off-canvas QA." },
  { artifactKind: "presentation", kind: "api", name: "compose.column", summary: "Create a vertical compose container. Use width/height fill, hug, or fixed pixels; gap and padding are in pixels." },
  { artifactKind: "presentation", kind: "api", name: "compose.paragraph", summary: "Create an editable text block with name, className/style text tokens, and stable inspect output." },

  { artifactKind: "document", kind: "api", name: "DocumentModel.create", summary: "Create a document with paragraph, list, table, header/footer, style, and comment blocks." },
  { artifactKind: "document", kind: "api", name: "document.addParagraph", summary: "Append a styled paragraph block and return an inspectable/resolveable paragraph object." },
  { artifactKind: "document", kind: "api", name: "document.addListItem", summary: "Append a real numbered or bulleted list item backed by DOCX numbering definitions." },
  { artifactKind: "document", kind: "api", name: "document.addHeader", summary: "Add header text exported as a DOCX header part and referenced from section properties." },
  { artifactKind: "document", kind: "api", name: "document.addFooter", summary: "Add footer text exported as a DOCX footer part and referenced from section properties." },
  { artifactKind: "document", kind: "api", name: "document.addHyperlink", summary: "Append an external hyperlink backed by a DOCX relationship and w:hyperlink element." },
  { artifactKind: "document", kind: "api", name: "document.addField", summary: "Append a Word field block exported as w:fldSimple with instruction text such as PAGE, REF, PAGEREF, or TOC." },
  { artifactKind: "document", kind: "api", name: "document.addCitation", summary: "Append a citation block with visible text and structured metadata preserved through clean-room DOCX metadata." },
  { artifactKind: "document", kind: "api", name: "document.addImage", summary: "Append an inspectable image block; dataUrl images export as native DOCX media parts with DrawingML inline pictures." },
  { artifactKind: "document", kind: "api", name: "document.addSection", summary: "Append a DOCX section break with page size, orientation, margin, and break-type metadata backed by w:sectPr." },
  { artifactKind: "document", kind: "api", name: "document.addChange", summary: "Append a tracked insertion or deletion block backed by native DOCX w:ins/w:del revision markup." },
  { artifactKind: "document", kind: "api", name: "document.addInsertion", summary: "Append a tracked insertion with author/date metadata and native DOCX w:ins export." },
  { artifactKind: "document", kind: "api", name: "document.addDeletion", summary: "Append a tracked deletion with author/date metadata and native DOCX w:del/w:delText export." },
  { artifactKind: "document", kind: "api", name: "document.addTable", summary: "Append a Word-style table block with rows, columns, cell values, and style metadata." },
  { artifactKind: "document", kind: "api", name: "document.addComment", summary: "Attach a comment to a paragraph or table block using a stable target ID." },
  { artifactKind: "document", kind: "api", name: "document.applyDesignPreset", summary: "Apply a clean-room report or memo design preset that updates named styles for consistent DOCX export and SVG/layout previews." },
  { artifactKind: "document", kind: "api", name: "document.inspect", summary: "Emit bounded NDJSON for document blocks, comments, styles, headers/footers, and layout; narrow with search/target anchors and shape fields with include/exclude." },
  { artifactKind: "document", kind: "api", name: "document.layoutJson", summary: "Return page-aware layout JSON with block bounding boxes, page records, style IDs, and design preset metadata." },
  { artifactKind: "document", kind: "api", name: "document.verify", summary: "Return QA issues for fake lists, invalid links/citations, unknown styles, malformed tables, bad image dimensions/data URLs, section setup, dangling comments, visual layout overflow, and prose-like table cells." },
  { artifactKind: "document", kind: "api", name: "DocumentFile.exportDocx", summary: "Export DocumentModel to a DOCX package with document.xml, styles.xml, comments.xml, numbering.xml, header/footer parts, hyperlinks, fields, citations, and metadata." },

  { artifactKind: "pdf", kind: "api", name: "PdfArtifact.create", summary: "Create a modeled PDF artifact with pages, text, table regions, and image regions." },
  { artifactKind: "pdf", kind: "api", name: "pdf.addImage", summary: "Add a modeled PDF image region with dataUrl/URI/prompt metadata, alt text, and page-space bounding box." },
  { artifactKind: "pdf", kind: "api", name: "pdf.extractText", summary: "Extract modeled text across all pages or a selected page." },
  { artifactKind: "pdf", kind: "api", name: "pdf.extractTables", summary: "Extract modeled table values and bounding boxes across all pages or a selected page." },
  { artifactKind: "pdf", kind: "api", name: "pdf.inspect", summary: "Emit bounded NDJSON for pages, text, positioned text items, layout regions, tables, and images; narrow with search/target anchors and shape fields with include/exclude." },
  { artifactKind: "pdf", kind: "api", name: "pdf.resolve", summary: "Resolve stable PDF artifact IDs for pages, page text blocks, positioned text items, layout regions, tables, and images." },
  { artifactKind: "pdf", kind: "api", name: "pdf.render", summary: "Render a modeled PDF page to SVG or return page layout JSON when called with { format: 'layout' }." },
  { artifactKind: "pdf", kind: "api", name: "pdf.layoutJson", summary: "Return modeled PDF page layout JSON with page text, positioned text items, layout regions, tables, and images." },
  { artifactKind: "pdf", kind: "api", name: "pdf.verify", summary: "Return QA issues for empty pages, Unicode dashes, malformed tables, and out-of-bounds table boxes." },
  { artifactKind: "pdf", kind: "api", name: "PdfFile.exportPdf", summary: "Export a modeled PDF artifact to a minimal PDF with visible text/table rows and embedded clean-room metadata." },
  { artifactKind: "pdf", kind: "api", name: "PdfFile.importPdf", summary: "Import clean-room generated PDFs from metadata, use an injected parser adapter for arbitrary PDFs, or fall back to heuristic visible-text/table extraction." },
  { artifactKind: "pdf", kind: "api", name: "createPdfjsParser", summary: "Create an optional PDF.js parser adapter from open-office-artifact-tool/pdf/pdfjs to extract page geometry, positioned text, heuristic tables, and image placeholders." },

  { artifactKind: "shared", kind: "api", name: "verifyArtifact", summary: "Run an artifact's verify() method and return a bounded NDJSON QA report." },
  { artifactKind: "shared", kind: "api", name: "visualQaArtifact", summary: "Render an artifact, record deterministic render metadata/hash, validate empty or malformed render output, optionally compare against a baseline render, and compute PNG pixel-diff metrics when requested." },
  { artifactKind: "shared", kind: "api", name: "renderArtifact", summary: "Render an artifact through its render/export method, attach normalized FileBlob metadata, and optionally pass SVG output through a caller-provided renderer adapter for PNG/WebP/JPEG/PDF output." },
  { artifactKind: "shared", kind: "api", name: "createPlaywrightRenderer", summary: "Create an optional Playwright renderer adapter from open-office-artifact-tool/renderers/playwright for deterministic SVG/HTML to PNG, WebP, JPEG, or PDF conversion with network blocked by default." },
  { artifactKind: "shared", kind: "api", name: "createSharpRenderer", summary: "Create an optional sharp renderer adapter from open-office-artifact-tool/renderers/sharp for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG, WebP, or JPEG." },
  { artifactKind: "shared", kind: "api", name: "createCanvasRenderer", summary: "Create an optional node-canvas renderer adapter from open-office-artifact-tool/renderers/canvas for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG or JPEG." },
  { artifactKind: "shared", kind: "api", name: "createPopplerRenderer", summary: "Create a Poppler CLI renderer adapter from open-office-artifact-tool/renderers/poppler for application/pdf FileBlob page rasterization to PNG, PPM, or TIFF." },
  { artifactKind: "shared", kind: "api", name: "createLibreOfficeRenderer", summary: "Create a LibreOffice CLI renderer adapter from open-office-artifact-tool/renderers/libreoffice for DOCX/XLSX/PPTX/HTML/PDF FileBlob conversion, typically to PDF." },
  { artifactKind: "shared", kind: "api", name: "createNativeOfficeRenderer", summary: "Create a native Office renderer adapter from open-office-artifact-tool/native/office-bridge that calls a JSON stdin/stdout sidecar command with timeout, temp-file isolation, cleanup, and structured errors." },
  { artifactKind: "shared", kind: "api", name: "renderFileWithNativeOffice", summary: "Render or convert a DOCX/XLSX/PPTX/PDF FileBlob through a configured native Office bridge command, returning a FileBlob for PDF/PNG/WebP or other requested output." },
];

const HELP_DETAIL_OVERRIDES = {
  "workbook.inspect": {
    examples: ["workbook.inspect({ kind: 'formula', target: 'Sheet1!E2', include: 'formula,value,precedents' })"],
    options: ["kind", "search/searchTerm", "target/targetId/id/anchor", "before/after/context", "include/fields", "exclude/omit", "maxChars"],
    returns: "{ ndjson, truncated } bounded NDJSON records",
  },
  "presentation.inspect": {
    examples: ["presentation.inspect({ kind: 'image,comment', target: image.id, include: 'alt,bbox' })"],
    options: ["kind", "search", "target/targetId/id/anchor", "before/after/context", "include/fields", "exclude/omit", "maxChars"],
    returns: "{ ndjson, truncated } bounded NDJSON records",
  },
  "document.inspect": {
    examples: ["document.inspect({ kind: 'paragraph,comment', target: comment.id, maxChars: 4000 })"],
    options: ["kind", "search", "target/targetId/id/anchor", "before/after/context", "include/fields", "exclude/omit", "maxChars"],
    returns: "{ ndjson, truncated } bounded NDJSON records",
  },
  "pdf.inspect": {
    examples: ["pdf.inspect({ kind: 'image,table', target: image.id, include: 'alt,bbox' })"],
    options: ["kind", "search", "target/targetId/id/anchor", "before/after/context", "include/fields", "exclude/omit", "maxChars"],
    returns: "{ ndjson, truncated } bounded NDJSON records",
  },
  renderArtifact: {
    examples: ["await renderArtifact(document, { format: 'png', renderer: createPlaywrightRenderer() })"],
    options: ["format", "renderer/rasterRenderer/renderAdapter", "page/pageIndex", "slide", "sheetName", "range"],
    returns: "FileBlob with normalized render metadata",
  },
  visualQaArtifact: {
    examples: ["await visualQaArtifact(document, { baseline, pixelDiff: true, minBytes: 100 })"],
    options: ["baseline/expected/baselineBlob", "pixelDiff", "allowChange", "minBytes", "maxBytes", "maxChars"],
    returns: "{ ok, blob, summary, issues, ndjson }",
  },
  verifyArtifact: {
    examples: ["verifyArtifact(workbook, { maxChars: 12000 })"],
    options: ["maxChars"],
    returns: "{ artifactKind, ok, issues, ndjson, truncated }",
  },
  "workbook.definedNames.add": {
    examples: ["workbook.definedNames.add('RevenueData', 'Sheet1!G2:G4')", "sheet.getRange('E3').formulas = [['=SUM(RevenueData)']]"] ,
    options: ["name", "refersTo", "scope/sheetName", "comment"],
    returns: "DefinedName facade with id/name/refersTo/scope",
  },
  "workbook.structuredReferences": {
    examples: ["=SUM(TasksTable[Revenue])"],
    notes: ["Current clean-room subset supports TableName[Column] data-body references; #All/#Headers/#Totals forms remain roadmap."],
  },
  createPlaywrightRenderer: {
    examples: ["const renderer = createPlaywrightRenderer({ viewport: { width: 900, height: 1200 }, deviceScaleFactor: 1 })"],
    options: ["viewport", "deviceScaleFactor", "allowNetwork", "timeoutMs", "format"],
    returns: "renderer adapter function for renderArtifact(...)",
  },
};

for (const item of HELP_CATALOG) {
  const details = HELP_DETAIL_OVERRIDES[item.name];
  if (details) Object.assign(item, details);
}

export function helpArtifact(artifactOrKind = "*", query = "*", options = {}) {
  const artifactKind = typeof artifactOrKind === "string" ? artifactOrKind : inferArtifactKind(artifactOrKind);
  const q = String(query || "*").toLowerCase();
  const search = String(options.search || "").toLowerCase();
  const records = HELP_CATALOG.filter((item) => {
    const kindMatch = artifactKind === "*" || item.artifactKind === artifactKind || (artifactKind === "unknown" && item.artifactKind === "shared");
    if (!kindMatch) return false;
    const haystack = `${item.name}\n${item.summary}\n${item.category || ""}\n${(item.examples || []).join("\n")}\n${(item.options || item.params || []).join("\n")}\n${item.returns || ""}\n${JSON.stringify(item.schema || {})}\n${(item.notes || []).join("\n")}`.toLowerCase();
    const queryMatch = q === "*" || item.name.toLowerCase().includes(q.replace("fx.", "")) || haystack.includes(q);
    const searchMatch = !search || haystack.includes(search);
    return queryMatch && searchMatch;
  });
  return ndjson(records, options.maxChars ?? Infinity);
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

export function node(type, props = {}, children = []) {
  const normalizedProps = props && !Array.isArray(props) && typeof props === "object" ? props : {};
  const rawChildren = Array.isArray(props) || typeof props === "string" || typeof props === "number" ? props : children;
  return { type, props: normalizedProps, children: normalizeComposeChildren(rawChildren) };
}

export const row = (props = {}, children = []) => node("row", props, children);
export const column = (props = {}, children = []) => node("column", props, children);
export const grid = (props = {}, children = []) => node("grid", props, children);
export const layers = (props = {}, children = []) => node("layers", props, children);
export const box = (props = {}, children = []) => node("box", props, children);
export const paragraph = (props = {}, children = []) => node("paragraph", props, children);
export const run = (props = {}, children = []) => node("run", props, children);
export const shape = (props = {}, children = []) => node("shape", props, children);
export const image = (props = {}, children = []) => node("image", props, children);
export const table = (props = {}, children = []) => node("table", props, children);
export const chart = (props = {}, children = []) => node("chart", props, children);
export const rule = (props = {}, children = []) => node("rule", props, children);

function normalizeComposeChildren(children) {
  if (children == null || children === false) return [];
  if (!Array.isArray(children)) return [children];
  return children.flatMap((child) => normalizeComposeChildren(child));
}

function isComposeNode(value) {
  return value && typeof value === "object" && typeof value.type === "string" && Array.isArray(value.children);
}

function textFromComposeChildren(children) {
  return normalizeComposeChildren(children).map((child) => {
    if (typeof child === "string" || typeof child === "number") return String(child);
    if (isComposeNode(child)) return textFromComposeChildren(child.children);
    return "";
  }).join("");
}

function normalizePadding(padding = {}) {
  if (typeof padding === "number") return { top: padding, right: padding, bottom: padding, left: padding };
  return {
    top: padding.top ?? padding.y ?? 0,
    right: padding.right ?? padding.x ?? 0,
    bottom: padding.bottom ?? padding.y ?? 0,
    left: padding.left ?? padding.x ?? 0,
  };
}

function innerFrame(frame, padding) {
  return {
    left: frame.left + padding.left,
    top: frame.top + padding.top,
    width: Math.max(0, frame.width - padding.left - padding.right),
    height: Math.max(0, frame.height - padding.top - padding.bottom),
  };
}

function resolveColorToken(value, fallback = "transparent") {
  if (!value) return fallback;
  const raw = String(value);
  const base = raw.split("/")[0];
  const tokens = {
    white: "#ffffff",
    black: "#000000",
    transparent: "transparent",
    "slate-50": "#f8fafc",
    "slate-100": "#f1f5f9",
    "slate-200": "#e2e8f0",
    "slate-300": "#cbd5e1",
    "slate-500": "#64748b",
    "slate-600": "#475569",
    "slate-700": "#334155",
    "slate-900": "#0f172a",
    "slate-950": "#020617",
    "sky-50": "#f0f9ff",
    "sky-100": "#e0f2fe",
    "sky-500": "#0ea5e9",
    "sky-600": "#0284c7",
    accent1: "#156082",
    tx1: "#1f1f1f",
    bg1: "#ffffff",
  };
  return tokens[base] || raw;
}

function parseTextStyle(props = {}) {
  const style = {};
  const className = String(props.className || "");
  for (const token of className.split(/\s+/).filter(Boolean)) {
    if (token === "font-bold") style.bold = true;
    if (token === "font-semibold") style.bold = true;
    if (token.startsWith("text-")) {
      const sizeMap = { xs: 12, sm: 14, base: 16, lg: 18, xl: 20, "2xl": 24, "3xl": 30, "4xl": 36, "5xl": 48, "6xl": 60 };
      const key = token.slice(5);
      if (sizeMap[key]) style.fontSize = sizeMap[key];
      else if (/^\[\d+px\]$/.test(key)) style.fontSize = Number(key.slice(1, -3));
      else style.color = resolveColorToken(key, style.color);
    }
    if (token.startsWith("leading-")) {
      const key = token.slice(8);
      if (key === "tight") style.lineSpacing = 1.1;
      else if (key === "relaxed") style.lineSpacing = 1.35;
      else if (/^\[\d+(?:\.\d+)?\]$/.test(key)) style.lineSpacing = Number(key.slice(1, -1));
    }
  }
  if (typeof props.style === "string") {
    const font = /font:\s*(\d+)\s+(\d+)px/i.exec(props.style);
    if (font) { style.bold = Number(font[1]) >= 600; style.fontSize = Number(font[2]); }
    const color = /color:\s*([^;]+)/i.exec(props.style);
    if (color) style.color = resolveColorToken(color[1].trim(), style.color);
    const leading = /leading:\s*([\d.]+)/i.exec(props.style);
    if (leading) style.lineSpacing = Number(leading[1]);
  } else if (props.style && typeof props.style === "object") {
    Object.assign(style, props.style);
  }
  return style;
}

function styleFromClassName(className = "") {
  const style = {};
  for (const token of String(className).split(/\s+/).filter(Boolean)) {
    if (token.startsWith("bg-")) style.fill = resolveColorToken(token.slice(3), style.fill);
    if (token.startsWith("rounded-")) style.borderRadius = token;
  }
  return style;
}

function composeIntrinsicSize(composeNode) {
  const props = composeNode.props || {};
  const text = textFromComposeChildren(composeNode.children);
  const textStyle = parseTextStyle(props);
  const fontSize = textStyle.fontSize || 20;
  const lineCount = Math.max(1, text.split(/\r?\n/).length);
  return {
    width: typeof props.width === "number" ? props.width : Math.max(160, Math.min(720, text.length * fontSize * 0.55 + 24)),
    height: typeof props.height === "number" ? props.height : Math.max(36, lineCount * fontSize * (textStyle.lineSpacing || 1.2) + 18),
  };
}

function composeChildFrames(children, frame, direction, gap) {
  const mainSize = direction === "row" ? "width" : "height";
  const crossSize = direction === "row" ? "height" : "width";
  const fixed = children.reduce((sum, child) => {
    const prop = child.props?.[mainSize];
    if (typeof prop === "number") return sum + prop;
    if (prop === "hug") return sum + composeIntrinsicSize(child)[mainSize];
    return sum;
  }, 0);
  const fillCount = children.filter((child) => child.props?.[mainSize] === "fill" || child.props?.[mainSize] == null).length || 1;
  const available = Math.max(0, frame[mainSize] - fixed - gap * Math.max(0, children.length - 1));
  const fillSize = available / fillCount;
  let cursor = direction === "row" ? frame.left : frame.top;
  return children.map((child) => {
    const intrinsic = composeIntrinsicSize(child);
    const main = typeof child.props?.[mainSize] === "number" ? child.props[mainSize] : child.props?.[mainSize] === "hug" ? intrinsic[mainSize] : fillSize;
    const cross = typeof child.props?.[crossSize] === "number" ? child.props[crossSize] : child.props?.[crossSize] === "hug" ? intrinsic[crossSize] : frame[crossSize];
    const childFrame = direction === "row"
      ? { left: cursor, top: frame.top, width: main, height: cross }
      : { left: frame.left, top: cursor, width: cross, height: main };
    cursor += main + gap;
    return childFrame;
  });
}

function normalizeTrack(track) {
  if (typeof track === "number") return { mode: "fixed", value: track };
  if (typeof track === "string") return track === "fixed" ? { mode: "fixed", value: 0 } : { mode: "fr", value: 1 };
  return { mode: track?.mode || "fr", value: Number(track?.value ?? 1) };
}

function resolveGridTracks(total, tracks, fallbackCount, gap) {
  const normalized = (tracks?.length ? tracks : Array.from({ length: fallbackCount }, () => ({ mode: "fr", value: 1 }))).map(normalizeTrack);
  const fixed = normalized.reduce((sum, track) => track.mode === "fixed" ? sum + track.value : sum, 0);
  const fr = normalized.reduce((sum, track) => track.mode === "fr" ? sum + Math.max(0, track.value) : sum, 0) || 1;
  const available = Math.max(0, total - fixed - gap * Math.max(0, normalized.length - 1));
  return normalized.map((track) => track.mode === "fixed" ? track.value : available * Math.max(0, track.value) / fr);
}

function gridChildFrame(frame, columns, rows, columnGap, rowGap, columnIndex, rowIndex, columnSpan = 1, rowSpan = 1) {
  const left = frame.left + columns.slice(0, columnIndex).reduce((sum, value) => sum + value, 0) + columnGap * columnIndex;
  const top = frame.top + rows.slice(0, rowIndex).reduce((sum, value) => sum + value, 0) + rowGap * rowIndex;
  const width = columns.slice(columnIndex, columnIndex + columnSpan).reduce((sum, value) => sum + value, 0) + columnGap * Math.max(0, columnSpan - 1);
  const height = rows.slice(rowIndex, rowIndex + rowSpan).reduce((sum, value) => sum + value, 0) + rowGap * Math.max(0, rowSpan - 1);
  return { left, top, width, height };
}

function materializeComposeNode(slide, composeNode, frame) {
  if (typeof composeNode === "string" || typeof composeNode === "number") {
    return materializeComposeNode(slide, paragraph({}, [String(composeNode)]), frame);
  }
  if (!isComposeNode(composeNode)) return [];
  const props = composeNode.props || {};
  const children = normalizeComposeChildren(composeNode.children).filter((child) => child !== null && child !== undefined && child !== false);
  const type = composeNode.type;
  if (type === "row" || type === "column") {
    const pad = normalizePadding(props.padding);
    const inner = innerFrame(frame, pad);
    const childFrames = composeChildFrames(children.filter(isComposeNode), inner, type, Number(props.gap || 0));
    return children.filter(isComposeNode).flatMap((child, index) => materializeComposeNode(slide, child, childFrames[index]));
  }
  if (type === "grid") {
    const gridChildren = children.filter(isComposeNode);
    const pad = normalizePadding(props.padding);
    const inner = innerFrame(frame, pad);
    const columnGap = Number(props.columnGap ?? props.gap ?? 0);
    const rowGap = Number(props.rowGap ?? props.gap ?? 0);
    const fallbackColumns = Math.max(1, props.columns?.length || Math.ceil(Math.sqrt(gridChildren.length || 1)));
    const columns = resolveGridTracks(inner.width, props.columns, fallbackColumns, columnGap);
    const fallbackRows = Math.max(1, props.rows?.length || Math.ceil((gridChildren.length || 1) / columns.length));
    const rows = resolveGridTracks(inner.height, props.rows, fallbackRows, rowGap);
    return gridChildren.flatMap((child, index) => {
      const columnIndex = Math.min(columns.length - 1, Number(child.props?.column ?? child.props?.col ?? (index % columns.length)));
      const rowIndex = Math.min(rows.length - 1, Number(child.props?.row ?? Math.floor(index / columns.length)));
      const columnSpan = Math.min(columns.length - columnIndex, Math.max(1, Number(child.props?.columnSpan ?? 1)));
      const rowSpan = Math.min(rows.length - rowIndex, Math.max(1, Number(child.props?.rowSpan ?? 1)));
      return materializeComposeNode(slide, child, gridChildFrame(inner, columns, rows, columnGap, rowGap, columnIndex, rowIndex, columnSpan, rowSpan));
    });
  }
  if (type === "layers") {
    const pad = normalizePadding(props.padding);
    const inner = innerFrame(frame, pad);
    return children.filter(isComposeNode).flatMap((child) => materializeComposeNode(slide, child, {
      left: inner.left,
      top: inner.top,
      width: typeof child.props?.width === "number" ? child.props.width : inner.width,
      height: typeof child.props?.height === "number" ? child.props.height : inner.height,
    }));
  }
  if (type === "box") {
    const classStyle = styleFromClassName(props.className);
    const surface = slide.shapes.add({
      id: props.id,
      name: props.name,
      geometry: props.geometry || "roundRect",
      position: frame,
      fill: props.fill || classStyle.fill || "transparent",
      line: props.line || { fill: "transparent", width: 0 },
      borderRadius: props.borderRadius || classStyle.borderRadius,
    });
    const pad = normalizePadding(props.padding ?? { x: 0, y: 0 });
    return [surface, ...children.filter(isComposeNode).flatMap((child) => materializeComposeNode(slide, child, innerFrame(frame, pad)))];
  }
  if (type === "paragraph") {
    const shape = slide.shapes.add({
      id: props.id,
      name: props.name,
      geometry: "textbox",
      position: frame,
      fill: "transparent",
      line: { fill: "transparent", width: 0 },
      text: textFromComposeChildren(children),
    });
    shape.text.style = parseTextStyle(props);
    return [shape];
  }
  if (type === "shape") {
    const classStyle = styleFromClassName(props.className);
    const shape = slide.shapes.add({
      ...props,
      position: frame,
      fill: props.fill || classStyle.fill || "transparent",
      text: textFromComposeChildren(children) || props.text,
    });
    shape.text.style = parseTextStyle(props);
    return [shape];
  }
  if (type === "table") {
    return [slide.tables.add({ ...props, position: frame })];
  }
  if (type === "chart") {
    return [slide.charts.add(props.chartType || props.type || "bar", { ...props, position: frame })];
  }
  if (type === "image") {
    return [slide.images.add({ ...props, position: frame, alt: props.alt || textFromComposeChildren(children) || props.name })];
  }
  if (type === "rule") {
    const horizontal = (props.width ?? frame.width) >= (props.height ?? props.weight ?? 2);
    const shape = slide.shapes.add({
      id: props.id,
      name: props.name,
      geometry: "rect",
      position: { left: frame.left, top: frame.top, width: horizontal ? frame.width : Number(props.weight || 2), height: horizontal ? Number(props.weight || 2) : frame.height },
      fill: props.stroke || "#0f172a",
      line: { fill: props.stroke || "#0f172a", width: 0 },
    });
    return [shape];
  }
  return children.filter(isComposeNode).flatMap((child) => materializeComposeNode(slide, child, frame));
}

export class FileBlob {
  constructor(data, options = {}) {
    this.bytes = toUint8Array(data ?? new Uint8Array());
    this.type = options.type || "application/octet-stream";
    this.metadata = options.metadata || {};
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

function workbookRangeRef(rangeOrRef) {
  if (rangeOrRef instanceof Range) return { sheetName: rangeOrRef.worksheet.name, address: rangeToAddress(rangeOrRef.bounds) };
  if (typeof rangeOrRef === "string") {
    const bang = rangeOrRef.lastIndexOf("!");
    return bang === -1 ? { sheetName: undefined, address: rangeOrRef } : { sheetName: rangeOrRef.slice(0, bang).replace(/^'|'$/g, ""), address: rangeOrRef.slice(bang + 1) };
  }
  if (rangeOrRef?.address) return { sheetName: rangeOrRef.sheetName, address: rangeOrRef.address };
  if (rangeOrRef?.cell) return workbookRangeRef(rangeOrRef.cell);
  if (rangeOrRef?.range) return workbookRangeRef(rangeOrRef.range);
  return { sheetName: undefined, address: "A1" };
}

class CommentThread {
  constructor(workbook, target, text, author) {
    this.workbook = workbook;
    this.id = aid("th");
    this.target = workbookRangeRef(target);
    this.author = author || workbook.comments.self?.displayName || "User";
    this.comments = [{ author: this.author, text: String(text ?? "") }];
    this.resolved = false;
  }

  addReply(text) {
    this.comments.push({ author: this.workbook.comments.self?.displayName || this.author, text: String(text ?? "") });
    return this;
  }

  resolve() { this.resolved = true; return this; }
  reopen() { this.resolved = false; return this; }

  inspectRecord() {
    return { kind: "thread", id: this.id, sheet: this.target.sheetName, address: this.target.address, author: this.author, resolved: this.resolved, replies: Math.max(0, this.comments.length - 1), textPreview: this.comments.map((comment) => comment.text).join("\n").slice(0, 300) };
  }

  toJSON() { return { id: this.id, target: this.target, author: this.author, comments: this.comments, resolved: this.resolved }; }
}

class CommentsCollection {
  constructor(workbook) { this.workbook = workbook; this.self = undefined; this.threads = []; }
  setSelf(self) { this.self = { displayName: self?.displayName || "User" }; return this.self; }
  addThread(target, text) { const thread = new CommentThread(this.workbook, target, text, this.self?.displayName); this.threads.push(thread); return thread; }
  getItem(id) { return this.threads.find((thread) => thread.id === id); }
  toJSON() { return { self: this.self, threads: this.threads.map((thread) => thread.toJSON()) }; }
}

class DefinedName {
  constructor(workbook, config = {}) {
    this.workbook = workbook;
    this.id = config.id || aid("dn");
    this.name = config.name || "DefinedName";
    this.refersTo = config.refersTo || config.reference || config.range || "Sheet1!A1";
    this.scope = config.scope || config.sheetName || undefined;
    this.comment = config.comment;
  }
  inspectRecord() { return { kind: "definedName", id: this.id, name: this.name, refersTo: this.refersTo, scope: this.scope, comment: this.comment }; }
  toJSON() { return { id: this.id, name: this.name, refersTo: this.refersTo, scope: this.scope, comment: this.comment }; }
}

class DefinedNameCollection {
  constructor(workbook) { this.workbook = workbook; this.items = []; }
  add(nameOrConfig, refersTo, options = {}) {
    const config = typeof nameOrConfig === "object" ? nameOrConfig : { name: nameOrConfig, refersTo, ...options };
    const existing = this.getItem(config.name, config.scope || config.sheetName);
    if (existing) Object.assign(existing, new DefinedName(this.workbook, { ...existing.toJSON(), ...config }));
    else this.items.push(new DefinedName(this.workbook, config));
    return existing || this.items.at(-1);
  }
  getItem(name, scope) { return this.items.find((item) => item.name === name && (scope == null || item.scope === scope)); }
  getItemOrNullObject(name, scope) { return this.getItem(name, scope) || { isNullObject: true }; }
  delete(nameOrId) { this.items = this.items.filter((item) => item.id !== nameOrId && item.name !== nameOrId); }
  inspectRecords() { return this.items.map((item) => item.inspectRecord()); }
  toJSON() { return this.items.map((item) => item.toJSON()); }
}

class WorksheetRuleCollection {
  constructor(worksheet, kind) { this.worksheet = worksheet; this.kind = kind; this.items = []; }
  add(config = {}) { const record = { id: aid(this.kind === "dataValidation" ? "dv" : "cf"), kind: this.kind, sheet: this.worksheet.name, ...config }; this.items.push(record); return record; }
  deleteAll() { this.items = []; }
  clear() { this.deleteAll(); }
  inspectRecords() { return this.items.map((item) => ({ ...item })); }
  toJSON() { return this.items.map((item) => ({ ...item })); }
}

class WorksheetTableRowsFacade {
  constructor(table) { this.table = table; }
  add(index, rows) {
    const insertAt = index == null ? this.table.values.length : index;
    this.table.values.splice(insertAt, 0, ...rows.map((row) => [...row]));
    this.table.refreshDimensions();
    return this.table;
  }
}

class WorksheetTable {
  constructor(worksheet, rangeOrConfig, hasHeaders = true, name) {
    this.worksheet = worksheet;
    const config = typeof rangeOrConfig === "object" && !(rangeOrConfig instanceof Range) && !Array.isArray(rangeOrConfig) ? rangeOrConfig : {};
    const rangeInput = config.range || rangeOrConfig || "A1";
    const range = rangeInput instanceof Range ? rangeInput : worksheet.getRange(String(rangeInput));
    this.id = config.id || aid("tbl");
    this.name = config.name || name || `Table${worksheet.tables.items.length + 1}`;
    this.anchor = { top: range.bounds.top, left: range.bounds.left };
    this.range = rangeToAddress(range.bounds);
    this.hasHeaders = config.hasHeaders ?? hasHeaders;
    this.showHeaders = this.hasHeaders;
    this.showTotals = Boolean(config.showTotals);
    this.showBandedColumns = Boolean(config.showBandedColumns);
    this.showFilterButton = config.showFilterButton ?? true;
    this.style = config.style || "TableStyleMedium2";
    this.values = config.values ? config.values.map((row) => [...row]) : range.values.map((row) => [...row]);
    this.rows = new WorksheetTableRowsFacade(this);
    this.refreshDimensions();
  }

  refreshDimensions() {
    this.rowCount = this.values.length;
    this.columnCount = Math.max(0, ...this.values.map((row) => row.length));
    const bounds = {
      top: this.anchor.top,
      left: this.anchor.left,
      bottom: this.anchor.top + Math.max(1, this.rowCount) - 1,
      right: this.anchor.left + Math.max(1, this.columnCount) - 1,
    };
    this.range = rangeToAddress(bounds);
  }

  getDataRows() { return this.showHeaders ? this.values.slice(1) : this.values; }
  getHeaderRowRange() { return this.worksheet.getRange(this.range).getResizedRange(1, this.columnCount || 1); }
  delete() { this.worksheet.tables.items = this.worksheet.tables.items.filter((table) => table !== this); }

  inspectRecord() {
    return { kind: "table", id: this.id, sheet: this.worksheet.name, name: this.name, address: this.range, rows: this.rowCount, cols: this.columnCount, hasHeaders: this.hasHeaders, style: this.style, values: this.values };
  }

  toSvg(bounds) {
    const tableBounds = parseRangeAddress(this.range);
    const left = 40 + (tableBounds.left - bounds.left) * 96;
    const top = 40 + (tableBounds.top - bounds.top) * 28;
    const width = Math.max(96, this.columnCount * 96);
    const height = Math.max(28, this.rowCount * 28);
    return `<rect x="${left}" y="${top}" width="${width}" height="${height}" fill="none" stroke="#0ea5e9" stroke-width="2"/><text x="${left}" y="${Math.max(12, top - 6)}" font-family="Arial" font-size="11" fill="#0284c7">${xmlEscape(this.name)}</text>`;
  }

  toJSON() { return { id: this.id, name: this.name, range: this.range, hasHeaders: this.hasHeaders, showHeaders: this.showHeaders, showTotals: this.showTotals, showBandedColumns: this.showBandedColumns, showFilterButton: this.showFilterButton, style: this.style, values: this.values }; }
}

class WorksheetTableCollection {
  constructor(worksheet) { this.worksheet = worksheet; this.items = []; }
  add(rangeOrConfig, hasHeaders = true, name) { const table = new WorksheetTable(this.worksheet, rangeOrConfig, hasHeaders, name); this.items.push(table); return table; }
  getItemOrNullObject(name) { return this.items.find((table) => table.name === name) || { isNullObject: true }; }
  deleteAll() { this.items = []; }
  inspectRecords() { return this.items.map((table) => table.inspectRecord()); }
  toJSON() { return this.items.map((table) => table.toJSON()); }
}

class WorksheetChartSeriesCollection {
  constructor(chart) { this.chart = chart; this.items = []; }
  add(name, values = []) { const series = { name, values, categoryFormula: undefined, formula: undefined, fill: undefined }; this.items.push(series); return series; }
  getItemAt(index) { return this.items[index]; }
  toJSON() { return this.items.map((item) => ({ ...item })); }
}

class WorksheetChart {
  constructor(worksheet, chartType = "bar", sourceOrConfig = {}) {
    this.worksheet = worksheet;
    this.id = sourceOrConfig.id || aid("wch");
    this.type = chartType;
    this.name = sourceOrConfig.name || `Chart ${worksheet.charts.items.length + 1}`;
    this.title = sourceOrConfig.title || "";
    this.hasLegend = sourceOrConfig.hasLegend ?? true;
    this.categories = sourceOrConfig.categories || [];
    this.position = sourceOrConfig.position || { left: 420, top: 40, width: 360, height: 220 };
    this.series = new WorksheetChartSeriesCollection(this);
    if (sourceOrConfig.series) sourceOrConfig.series.forEach((series) => this.series.add(series.name, series.values || []));
    if (sourceOrConfig instanceof Range) this.setData(sourceOrConfig);
    else if (sourceOrConfig && sourceOrConfig.worksheet instanceof Worksheet) this.setData(sourceOrConfig);
  }

  setData(range) {
    const values = range.values;
    if (!values.length || !values[0]?.length) return this;
    const header = values[0];
    const dataRows = values.slice(1);
    this.categories = dataRows.map((row) => String(row[0] ?? ""));
    this.series.items = [];
    for (let column = 1; column < header.length; column++) {
      const series = this.series.add(String(header[column] || `Series ${column}`), dataRows.map((row) => Number(row[column]) || 0));
      const start = makeCellAddress(range.bounds.top + 1, range.bounds.left + column);
      const end = makeCellAddress(range.bounds.bottom, range.bounds.left + column);
      series.formula = `'${range.worksheet.name}'!${start}:${end}`;
      series.categoryFormula = `'${range.worksheet.name}'!${makeCellAddress(range.bounds.top + 1, range.bounds.left)}:${makeCellAddress(range.bounds.bottom, range.bounds.left)}`;
    }
    return this;
  }

  setPosition(topLeft, bottomRight) {
    const start = parseCellAddress(String(topLeft).replace(/^.*!/, ""));
    const end = parseCellAddress(String(bottomRight).replace(/^.*!/, ""));
    this.position = { left: 40 + start.col * 96, top: 40 + start.row * 28, width: Math.max(120, (end.col - start.col + 1) * 96), height: Math.max(80, (end.row - start.row + 1) * 28) };
    return this;
  }

  inspectRecord() { return { kind: "drawing", drawingType: "chart", id: this.id, sheet: this.worksheet.name, name: this.name, chartType: this.type, title: this.title, categories: this.categories, series: this.series.items.length, bbox: [this.position.left, this.position.top, this.position.width, this.position.height], bboxUnit: "px" }; }

  toSvg() {
    const p = this.position;
    const values = this.series.items[0]?.values || [];
    const max = Math.max(1, ...values.map((value) => Number(value) || 0));
    const plot = { left: p.left + 28, top: p.top + 36, width: Math.max(0, p.width - 44), height: Math.max(0, p.height - 62) };
    const barW = values.length ? plot.width / values.length * 0.65 : 0;
    const gap = values.length ? plot.width / values.length * 0.35 : 0;
    const bars = values.map((value, index) => {
      const h = plot.height * (Number(value) || 0) / max;
      return `<rect x="${plot.left + index * (barW + gap) + gap / 2}" y="${plot.top + plot.height - h}" width="${barW}" height="${h}" fill="#38bdf8"/>`;
    }).join("");
    return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#94a3b8"/><text x="${p.left + 8}" y="${p.top + 22}" font-family="Arial" font-size="13" font-weight="700" fill="#0f172a">${xmlEscape(this.title || this.name)}</text>${bars}`;
  }

  toJSON() { return { id: this.id, type: this.type, name: this.name, title: this.title, hasLegend: this.hasLegend, categories: this.categories, position: this.position, series: this.series.toJSON() }; }
}

class WorksheetChartCollection {
  constructor(worksheet) { this.worksheet = worksheet; this.items = []; }
  add(chartType, sourceOrConfig = {}) { const chart = new WorksheetChart(this.worksheet, chartType, sourceOrConfig); this.items.push(chart); return chart; }
  getItemOrNullObject(name) { return this.items.find((chart) => chart.name === name) || { isNullObject: true }; }
  deleteAll() { this.items = []; }
  inspectRecords() { return this.items.map((chart) => chart.inspectRecord()); }
  toJSON() { return this.items.map((chart) => chart.toJSON()); }
}

function worksheetAnchorFrame(anchor = {}) {
  const from = anchor.from || { row: 0, col: 0 };
  const extent = anchor.extent || {};
  return {
    left: 40 + Number(from.col || 0) * 96 + Number(anchor.colOffsetPx || from.colOffsetPx || 0),
    top: 40 + Number(from.row || 0) * 28 + Number(anchor.rowOffsetPx || from.rowOffsetPx || 0),
    width: Number(extent.widthPx || anchor.widthPx || 160),
    height: Number(extent.heightPx || anchor.heightPx || 120),
  };
}

class WorksheetImage {
  constructor(worksheet, config = {}) {
    this.worksheet = worksheet;
    this.id = config.id || aid("wim");
    this.name = config.name || `Image ${worksheet.images.items.length + 1}`;
    this.dataUrl = config.dataUrl;
    this.uri = config.uri;
    this.prompt = config.prompt;
    this.alt = config.alt || config.name || "image";
    this.anchor = config.anchor || { from: { row: 0, col: 0 }, extent: { widthPx: 160, heightPx: 120 } };
    this.fit = config.fit || "contain";
  }

  get position() { return worksheetAnchorFrame(this.anchor); }
  inspectRecord() { const p = this.position; return { kind: "drawing", drawingType: "image", id: this.id, sheet: this.worksheet.name, name: this.name, alt: this.alt, prompt: this.prompt, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px" }; }
  replace(config = {}) { Object.assign(this, config); return this; }
  toSvg() { const p = this.position; const image = this.dataUrl ? `<image href="${attrEscape(this.dataUrl)}" x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" preserveAspectRatio="xMidYMid meet"/>` : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#fef3c7" stroke="#f59e0b"/>`; return `${image}<text x="${p.left + 8}" y="${p.top + 20}" font-family="Arial" font-size="12" fill="#92400e">${xmlEscape(this.alt || this.prompt || this.name)}</text>`; }
  toJSON() { return { id: this.id, name: this.name, dataUrl: this.dataUrl, uri: this.uri, prompt: this.prompt, alt: this.alt, anchor: this.anchor, fit: this.fit }; }
}

class WorksheetImageCollection {
  constructor(worksheet) { this.worksheet = worksheet; this.items = []; }
  add(config = {}) { const image = new WorksheetImage(this.worksheet, config); this.items.push(image); return image; }
  getItemOrNullObject(name) { return this.items.find((image) => image.name === name) || { isNullObject: true }; }
  deleteAll() { this.items = []; }
  inspectRecords() { return this.items.map((image) => image.inspectRecord()); }
  toJSON() { return this.items.map((image) => image.toJSON()); }
}

class SparklineGroup {
  constructor(worksheet, config = {}) {
    this.worksheet = worksheet;
    this.id = config.id || aid("sp");
    this.type = config.type || "line";
    this.targetRange = workbookRangeRef(config.targetRange || "A1");
    this.sourceData = workbookRangeRef(config.sourceData || "A1");
    this.dateAxisRange = config.dateAxisRange ? workbookRangeRef(config.dateAxisRange) : undefined;
    this.seriesColor = config.seriesColor || "#0ea5e9";
    this.negativeColor = config.negativeColor;
    this.markers = config.markers || {};
    this.axis = config.axis || {};
    this.lineWeight = config.lineWeight ?? 1.5;
    this.displayHidden = Boolean(config.displayHidden);
    this.displayEmptyCellsAs = config.displayEmptyCellsAs;
  }

  delete() { this.worksheet.sparklineGroups.items = this.worksheet.sparklineGroups.items.filter((group) => group !== this); }
  inspectRecord() { return { kind: "sparkline", id: this.id, sheet: this.worksheet.name, type: this.type, targetRange: this.targetRange.address, sourceData: this.sourceData.address, dateAxisRange: this.dateAxisRange?.address, seriesColor: this.seriesColor }; }
  values() { const sourceSheet = this.sourceData.sheetName ? this.worksheet.workbook.worksheets.getItem(this.sourceData.sheetName) : this.worksheet; if (!sourceSheet) return []; return sourceSheet.getRange(this.sourceData.address).values.flat().map((value) => Number(value)).filter((value) => Number.isFinite(value)); }
  targetFrame(bounds) { const target = parseRangeAddress(this.targetRange.address); return { left: 40 + (target.left - bounds.left) * 96, top: 40 + (target.top - bounds.top) * 28, width: Math.max(96, target.colCount * 96), height: Math.max(28, target.rowCount * 28) }; }
  toSvg(bounds) { const p = this.targetFrame(bounds); const values = this.values(); if (!values.length) return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="none" stroke="#38bdf8" stroke-dasharray="3 2"/>`; const min = Math.min(...values); const max = Math.max(...values); const span = Math.max(1, max - min); const points = values.map((value, index) => `${p.left + (values.length === 1 ? p.width / 2 : index * p.width / (values.length - 1))},${p.top + p.height - ((value - min) / span) * p.height}`).join(" "); if (this.type === "column") { const barW = p.width / values.length * 0.7; return values.map((value, index) => { const h = ((value - Math.min(0, min)) / Math.max(1, max - Math.min(0, min))) * p.height; return `<rect x="${p.left + index * (p.width / values.length) + barW * 0.15}" y="${p.top + p.height - h}" width="${barW}" height="${h}" fill="${xmlEscape(this.seriesColor)}"/>`; }).join(""); } return `<polyline points="${points}" fill="none" stroke="${xmlEscape(this.seriesColor)}" stroke-width="${this.lineWeight}"/>`; }
  toJSON() { return { id: this.id, type: this.type, targetRange: this.targetRange, sourceData: this.sourceData, dateAxisRange: this.dateAxisRange, seriesColor: this.seriesColor, negativeColor: this.negativeColor, markers: this.markers, axis: this.axis, lineWeight: this.lineWeight, displayHidden: this.displayHidden, displayEmptyCellsAs: this.displayEmptyCellsAs }; }
}

class SparklineGroupCollection {
  constructor(worksheet) { this.worksheet = worksheet; this.items = []; }
  add(config = {}) { const group = new SparklineGroup(this.worksheet, config); this.items.push(group); return group; }
  deleteAll() { this.items = []; }
  inspectRecords() { return this.items.map((group) => group.inspectRecord()); }
  toJSON() { return this.items.map((group) => group.toJSON()); }
}

class RangeSparklineFacade {
  constructor(range) { this.range = range; }
  add(type, sourceData, config = {}) { return this.range.worksheet.sparklineGroups.add({ ...config, type, targetRange: this.range, sourceData }); }
}

function xlsxSheetNameForFormula(name) {
  const raw = String(name || "");
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(raw) ? raw : `'${raw.replaceAll("'", "''")}'`;
}

function xlsxQualifiedRangeRef(ref, defaultSheetName) {
  const normalized = workbookRangeRef(ref);
  const sheetName = normalized.sheetName || defaultSheetName;
  return `${xlsxSheetNameForFormula(sheetName)}!${normalized.address}`;
}

function xlsxColorRgb(color, fallback = "FF0EA5E9") {
  const raw = String(color || "").trim();
  const hex = raw.startsWith("#") ? raw.slice(1) : raw;
  if (/^[0-9a-fA-F]{6}$/.test(hex)) return `FF${hex.toUpperCase()}`;
  if (/^[0-9a-fA-F]{8}$/.test(hex)) return hex.toUpperCase();
  return fallback;
}

function sparklineGroupExtXml(sheet) {
  if (!sheet.sparklineGroups.items.length) return "";
  const groups = sheet.sparklineGroups.items.map((group) => {
    const attrs = [
      `type="${attrEscape(group.type || "line")}"`,
      group.displayHidden ? `displayHidden="1"` : "",
      group.displayEmptyCellsAs ? `displayEmptyCellsAs="${attrEscape(group.displayEmptyCellsAs)}"` : "",
      group.markers?.show ? `markers="1"` : "",
      group.axis?.show ? `displayXAxis="1"` : "",
      group.dateAxisRange ? `dateAxis="1"` : "",
      group.lineWeight != null ? `lineWeight="${Number(group.lineWeight) || 1}"` : "",
    ].filter(Boolean).join(" ");
    const dateAxis = group.dateAxisRange ? `<x14:dateAxisRange>${xmlEscape(xlsxQualifiedRangeRef(group.dateAxisRange, sheet.name))}</x14:dateAxisRange>` : "";
    const negative = group.negativeColor ? `<x14:colorNegative rgb="${xlsxColorRgb(group.negativeColor, "FFFF0000")}"/>` : "";
    return `<x14:sparklineGroup ${attrs}><x14:colorSeries rgb="${xlsxColorRgb(group.seriesColor)}"/>${negative}${dateAxis}<x14:sparklines><x14:sparkline><xm:f>${xmlEscape(xlsxQualifiedRangeRef(group.sourceData, sheet.name))}</xm:f><xm:sqref>${xmlEscape(group.targetRange.address)}</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>`;
  }).join("");
  return `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${groups}</x14:sparklineGroups></ext></extLst>`;
}

function parseXlsxBool(value) {
  return value === "1" || value === "true";
}

function parseSparklineGroupsXml(sheet, xml) {
  for (const match of String(xml || "").matchAll(/<x14:sparklineGroup\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroup>/g)) {
    const attrs = match[1] || "";
    const body = match[2] || "";
    const formula = decodeXml(/<xm:f[^>]*>([\s\S]*?)<\/xm:f>/.exec(body)?.[1] || "");
    const targetRange = decodeXml(/<xm:sqref[^>]*>([\s\S]*?)<\/xm:sqref>/.exec(body)?.[1] || "");
    if (!formula || !targetRange) continue;
    const sourceData = workbookRangeRef(formula.replace(/^'([^']*(?:''[^']*)*)'!/, (_, sheetName) => `${sheetName.replaceAll("''", "'")}!`));
    const type = /\btype="([^"]+)"/.exec(attrs)?.[1] || "line";
    const seriesColor = /<x14:colorSeries[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(body)?.[1];
    const negativeColor = /<x14:colorNegative[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(body)?.[1];
    const dateAxisText = decodeXml(/<x14:dateAxisRange[^>]*>([\s\S]*?)<\/x14:dateAxisRange>/.exec(body)?.[1] || "");
    sheet.sparklineGroups.add({
      type,
      targetRange,
      sourceData,
      dateAxisRange: dateAxisText || undefined,
      seriesColor: seriesColor ? `#${seriesColor}` : undefined,
      negativeColor: negativeColor ? `#${negativeColor}` : undefined,
      markers: { show: parseXlsxBool(/\bmarkers="([^"]+)"/.exec(attrs)?.[1]) },
      axis: { show: parseXlsxBool(/\bdisplayXAxis="([^"]+)"/.exec(attrs)?.[1]) },
      displayHidden: parseXlsxBool(/\bdisplayHidden="([^"]+)"/.exec(attrs)?.[1]),
      displayEmptyCellsAs: /\bdisplayEmptyCellsAs="([^"]+)"/.exec(attrs)?.[1],
      lineWeight: Number(/\blineWeight="([^"]+)"/.exec(attrs)?.[1] || 1.5),
    });
  }
}

class RangeConditionalFormatFacade {
  constructor(range) { this.range = range; }
  add(ruleType, config = {}) { return this.range.worksheet.conditionalFormattings.add({ range: rangeToAddress(this.range.bounds), ruleType, ...config }); }
  addCustom(expression, format = {}) { return this.add("expression", { formula: expression, format }); }
  deleteAll() { this.range.worksheet.conditionalFormattings.items = this.range.worksheet.conditionalFormattings.items.filter((item) => item.range !== rangeToAddress(this.range.bounds)); }
  clear() { this.deleteAll(); }
}

function safeRangeBounds(address) {
  try { return parseRangeAddress(address); } catch { return undefined; }
}

function workbookRangeTarget(workbook, defaultSheet, ref) {
  const normalized = workbookRangeRef(ref);
  const sheet = normalized.sheetName ? workbook.worksheets.getItem(normalized.sheetName) : defaultSheet;
  const bounds = safeRangeBounds(normalized.address);
  return { ...normalized, sheet, bounds };
}

function workbookRangeCount(workbook, defaultSheet, ref) {
  const target = workbookRangeTarget(workbook, defaultSheet, ref);
  return target.sheet && target.bounds ? target.bounds.rowCount * target.bounds.colCount : 0;
}

function workbookRangeValid(workbook, defaultSheet, ref) {
  const target = workbookRangeTarget(workbook, defaultSheet, ref);
  return Boolean(target.sheet && target.bounds);
}

function worksheetFrameForBounds(bounds) {
  return { left: 40 + bounds.left * 96, top: 40 + bounds.top * 28, width: Math.max(96, bounds.colCount * 96), height: Math.max(28, bounds.rowCount * 28) };
}

function workbookFrameIntersects(a, b) {
  const left = Math.max(a.left, b.left);
  const top = Math.max(a.top, b.top);
  const right = Math.min(a.left + a.width, b.left + b.width);
  const bottom = Math.min(a.top + a.height, b.top + b.height);
  return right > left && bottom > top;
}

export class Workbook {
  constructor() {
    this.id = aid("wb");
    this.worksheets = new WorksheetCollection(this);
    this.comments = new CommentsCollection(this);
    this.definedNames = new DefinedNameCollection(this);
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
    if (this._recalculating) return this._lastFormulaGraph;
    this._recalculating = true;
    try {
      const graph = buildWorkbookFormulaGraph(this);
      const formulaNodes = new Map(graph.nodes.map((node) => [node.key, node]));
      const cycleKeys = new Set(graph.cycles.flatMap((cycle) => cycle.keys));
      const evaluated = new Set();
      const evaluateNode = (node) => {
        if (!node) return null;
        if (cycleKeys.has(node.key)) {
          node.cell.value = "#CYCLE!";
          evaluated.add(node.key);
          return node.cell.value;
        }
        if (evaluated.has(node.key)) return node.cell.value;
        const value = evaluateFormula(node.sheetObject, node.formula, node.address, {
          getValue: (ref) => {
            const targetSheet = ref.sheetName ? this.worksheets.getItem(ref.sheetName) : node.sheetObject;
            if (!targetSheet) return "#REF!";
            const targetAddress = String(ref.address || "").replaceAll("$", "").toUpperCase();
            const targetKey = formulaCellKey(targetSheet.name, targetAddress);
            const targetNode = formulaNodes.get(targetKey);
            if (targetNode) return evaluateNode(targetNode);
            return targetSheet.store.get(targetAddress).value;
          },
        });
        node.cell.value = value;
        evaluated.add(node.key);
        return value;
      };
      for (const node of graph.nodes) evaluateNode(node);
      this._lastFormulaGraph = buildWorkbookFormulaGraph(this);
      return this._lastFormulaGraph;
    } finally {
      this._recalculating = false;
    }
  }

  formulaGraph(options = {}) {
    if (options.recalculate !== false) this.recalculate();
    const graph = this._lastFormulaGraph || buildWorkbookFormulaGraph(this);
    const records = formulaGraphRecords(graph, options);
    return {
      kind: "formulaGraph",
      nodes: graph.nodes.map(publicFormulaNode),
      edges: graph.edges.map((edge) => ({ ...edge })),
      cycles: graph.cycles.map((cycle) => ({ ...cycle, path: [...cycle.path], keys: [...cycle.keys] })),
      errors: graph.errors.map((error) => ({ ...error })),
      ...ndjson(records, options.maxChars ?? Infinity),
    };
  }

  inspect(options = {}) {
    this.recalculate();
    const kinds = normalizeKinds(options.kind, ["workbook", "sheet", "table", "formula"]);
    const records = [];
    const graph = (kinds.has("formula") || kinds.has("formulaGraph") || kinds.has("formulaNode") || kinds.has("formulaEdge") || kinds.has("formulaCycle")) ? this.formulaGraph({ ...options, recalculate: false, maxChars: Infinity }) : null;
    if (kinds.has("workbook")) records.push({ kind: "workbook", id: this.id, sheets: this.worksheets.items.length });
    for (const sheet of this.worksheets) {
      if (kinds.has("sheet")) records.push({ kind: "sheet", id: sheet.id, name: sheet.name, rows: sheet.usedBounds().rowCount, cols: sheet.usedBounds().colCount });
      if (kinds.has("table") || kinds.has("region")) records.push(sheet.tableRecord(options));
      if (kinds.has("table")) records.push(...sheet.tables.inspectRecords());
      if (kinds.has("drawing") || kinds.has("chart")) records.push(...sheet.charts.inspectRecords());
      if (kinds.has("drawing") || kinds.has("image")) records.push(...sheet.images.inspectRecords());
      if (kinds.has("sparkline") || kinds.has("drawing")) records.push(...sheet.sparklineGroups.inspectRecords());
      if (kinds.has("formula")) records.push(...sheet.formulaRecords({ ...options, graph }));
      if (kinds.has("style") || kinds.has("computedStyle")) records.push(...sheet.styleRecords(options));
      if (kinds.has("match")) records.push(...sheet.matchRecords(options));
      if (kinds.has("dataValidation")) records.push(...sheet.dataValidations.inspectRecords());
      if (kinds.has("conditionalFormat")) records.push(...sheet.conditionalFormattings.inspectRecords());
    }
    if (kinds.has("thread")) records.push(...this.comments.threads.map((thread) => thread.inspectRecord()));
    if (kinds.has("definedName") || kinds.has("name")) records.push(...this.definedNames.inspectRecords());
    if (kinds.has("formulaGraph") || kinds.has("formulaNode") || kinds.has("formulaEdge") || kinds.has("formulaCycle")) records.push(...formulaGraphRecords(graph, { ...options, kinds }));
    return ndjson(filterInspectRecords(records, options), options.maxChars ?? Infinity);
  }

  verify(options = {}) {
    this.recalculate();
    const graph = this.formulaGraph({ recalculate: false, maxChars: Infinity });
    const issues = [];
    if (this.worksheets.items.length === 0) issues.push(verificationIssue("workbook", "noSheets", "Workbook has no worksheets."));
    for (const definedName of this.definedNames.items) {
      const refersTo = String(definedName.refersTo || "").replace(/^=/, "");
      if (!formulaRefParts(refersTo)) issues.push(verificationIssue("workbook", "invalidDefinedName", `Defined name ${definedName.name} does not reference a valid A1 range.`, { id: definedName.id, name: definedName.name, refersTo: definedName.refersTo }));
    }
    for (const sheet of this.worksheets) {
      const entries = sheet.store.entries();
      if (entries.length === 0) issues.push(verificationIssue("workbook", "emptySheet", `Worksheet ${sheet.name} has no populated cells.`, { severity: "warning", sheet: sheet.name }));
      for (const [address, cell] of entries) {
        const value = String(cell.value ?? "");
        if (/^#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A|CYCLE!)/.test(value)) {
          issues.push(verificationIssue("workbook", "formulaError", `Formula error ${value} at ${sheet.name}!${address}.`, { sheet: sheet.name, address, value }));
        }
      }
      for (const table of sheet.tables.items) {
        const tableBounds = safeRangeBounds(table.range);
        if (!table.name) issues.push(verificationIssue("workbook", "unnamedTable", `A worksheet table on ${sheet.name} is missing a stable name.`, { sheet: sheet.name, id: table.id }));
        if (!tableBounds) issues.push(verificationIssue("workbook", "invalidTableRange", `Table ${table.name || table.id} on ${sheet.name} has invalid range ${table.range}.`, { sheet: sheet.name, id: table.id, range: table.range }));
        if (tableBounds && (tableBounds.rowCount !== Math.max(1, table.rowCount) || tableBounds.colCount !== Math.max(1, table.columnCount))) issues.push(verificationIssue("workbook", "tableRangeMismatch", `Table ${table.name || table.id} range does not match table values.`, { severity: "warning", sheet: sheet.name, id: table.id, range: table.range, rangeRows: tableBounds.rowCount, rangeCols: tableBounds.colCount, rows: table.rowCount, cols: table.columnCount }));
        if (table.values.some((row) => row.length !== table.columnCount)) issues.push(verificationIssue("workbook", "raggedWorksheetTable", `Table ${table.name || table.id} has ragged rows.`, { sheet: sheet.name, id: table.id }));
      }
      for (const chart of sheet.charts.items) {
        if (!chart.title) issues.push(verificationIssue("workbook", "untitledChart", `Chart ${chart.name} on ${sheet.name} has no title.`, { severity: "warning", sheet: sheet.name, id: chart.id }));
        if (chart.series.items.length === 0) issues.push(verificationIssue("workbook", "emptyChart", `Chart ${chart.name} on ${sheet.name} has no data series.`, { sheet: sheet.name, id: chart.id }));
        for (const series of chart.series.items) {
          if (chart.categories.length && series.values.length && chart.categories.length !== series.values.length) issues.push(verificationIssue("workbook", "chartDataMismatch", `Chart ${chart.name} series ${series.name || "Series"} has ${series.values.length} values for ${chart.categories.length} categories.`, { sheet: sheet.name, id: chart.id, series: series.name, values: series.values.length, categories: chart.categories.length }));
          if (series.formula && !workbookRangeValid(this, sheet, series.formula)) issues.push(verificationIssue("workbook", "chartFormulaInvalid", `Chart ${chart.name} series ${series.name || "Series"} references an invalid range.`, { sheet: sheet.name, id: chart.id, formula: series.formula }));
          if (series.categoryFormula && !workbookRangeValid(this, sheet, series.categoryFormula)) issues.push(verificationIssue("workbook", "chartCategoryFormulaInvalid", `Chart ${chart.name} categories reference an invalid range.`, { sheet: sheet.name, id: chart.id, formula: series.categoryFormula }));
        }
      }
      for (const image of sheet.images.items) {
        const frame = image.position;
        if (!image.dataUrl && !image.uri && !image.prompt) issues.push(verificationIssue("workbook", "emptyImage", `Worksheet image ${image.name || image.id} on ${sheet.name} has no dataUrl, uri, or prompt.`, { sheet: sheet.name, id: image.id }));
        if (image.dataUrl && !imageDataFromDataUrl(image.dataUrl)) issues.push(verificationIssue("workbook", "invalidImageDataUrl", `Worksheet image ${image.name || image.id} has an unsupported data URL.`, { sheet: sheet.name, id: image.id }));
        if (frame.left < 0 || frame.top < 0 || frame.width <= 0 || frame.height <= 0) issues.push(verificationIssue("workbook", "imageBoundsInvalid", `Worksheet image ${image.name || image.id} has invalid bounds.`, { sheet: sheet.name, id: image.id, bbox: [frame.left, frame.top, frame.width, frame.height] }));
      }
      for (const sparkline of sheet.sparklineGroups.items) {
        const targetValid = workbookRangeValid(this, sheet, sparkline.targetRange);
        const sourceValid = workbookRangeValid(this, sheet, sparkline.sourceData);
        if (!targetValid) issues.push(verificationIssue("workbook", "sparklineTargetInvalid", `Sparkline ${sparkline.id} on ${sheet.name} has invalid target range.`, { sheet: sheet.name, id: sparkline.id, targetRange: sparkline.targetRange.address }));
        if (!sourceValid) issues.push(verificationIssue("workbook", "sparklineSourceInvalid", `Sparkline ${sparkline.id} on ${sheet.name} has invalid source data range.`, { sheet: sheet.name, id: sparkline.id, sourceData: sparkline.sourceData.address }));
        if (targetValid && sourceValid && workbookRangeCount(this, sheet, sparkline.sourceData) === 0) issues.push(verificationIssue("workbook", "emptySparklineSource", `Sparkline ${sparkline.id} on ${sheet.name} has no source cells.`, { severity: "warning", sheet: sheet.name, id: sparkline.id }));
      }
      for (const validation of sheet.dataValidations.items) {
        if (!safeRangeBounds(validation.range || "")) issues.push(verificationIssue("workbook", "dataValidationRangeInvalid", `Data validation ${validation.id} on ${sheet.name} has invalid range ${validation.range}.`, { sheet: sheet.name, id: validation.id, range: validation.range }));
        const rule = validation.rule || validation;
        if ((rule.type || "").toLowerCase() === "list" && (!Array.isArray(rule.values) || rule.values.length === 0) && !rule.formula1) issues.push(verificationIssue("workbook", "dataValidationListEmpty", `List validation ${validation.id} on ${sheet.name} has no values or formula.`, { sheet: sheet.name, id: validation.id }));
      }
      for (const format of sheet.conditionalFormattings.items) {
        if (!safeRangeBounds(format.range || "")) issues.push(verificationIssue("workbook", "conditionalFormatRangeInvalid", `Conditional format ${format.id} on ${sheet.name} has invalid range ${format.range}.`, { sheet: sheet.name, id: format.id, range: format.range }));
        if ((format.ruleType === "expression" || format.kind === "conditionalFormat") && !format.formula && !format.expression && format.ruleType !== "cellIs") issues.push(verificationIssue("workbook", "conditionalFormatFormulaMissing", `Conditional format ${format.id} on ${sheet.name} has no formula/expression.`, { sheet: sheet.name, id: format.id }));
      }
    }
    for (const thread of this.comments.threads) {
      if (!thread.target?.address) issues.push(verificationIssue("workbook", "unanchoredComment", `Comment thread ${thread.id} is missing a target cell.`, { id: thread.id }));
      else if (!workbookRangeValid(this, thread.target.sheetName ? this.worksheets.getItem(thread.target.sheetName) : this.worksheets.items[0], thread.target)) issues.push(verificationIssue("workbook", "commentTargetInvalid", `Comment thread ${thread.id} points at an invalid target cell.`, { id: thread.id, target: thread.target }));
    }
    for (const cycle of graph.cycles) {
      issues.push(verificationIssue("workbook", "formulaCycle", `Formula cycle detected: ${cycle.path.join(" -> ")}.`, { cycle: cycle.path, keys: cycle.keys }));
    }
    for (const error of graph.errors) {
      if (error.type === "missingSheet") {
        const { type: _type, ...details } = error;
        issues.push(verificationIssue("workbook", "missingFormulaSheet", `Formula at ${error.from} references missing worksheet ${error.sheet}.`, details));
      }
    }
    return verificationResult("workbook", issues, options);
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
      for (const ref of formulaReferences(cell.formula, sheet)) {
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
    const thread = this.comments.getItem(id);
    if (thread) return thread;
    const definedName = this.definedNames.items.find((item) => item.id === id || item.name === id);
    if (definedName) return definedName;
    for (const sheet of this.worksheets) {
      if (sheet.id === id) return sheet;
      const table = sheet.tables.items.find((item) => item.id === id);
      if (table) return table;
      const chart = sheet.charts.items.find((item) => item.id === id);
      if (chart) return chart;
      const image = sheet.images.items.find((item) => item.id === id);
      if (image) return image;
      const sparkline = sheet.sparklineGroups.items.find((item) => item.id === id);
      if (sparkline) return sparkline;
      const rule = [...sheet.dataValidations.items, ...sheet.conditionalFormattings.items].find((item) => item.id === id);
      if (rule) return rule;
    }
    return undefined;
  }

  help(query = "*", options = {}) {
    return helpArtifact("workbook", query, options);
  }

  layoutJson(options = {}) {
    const selected = options.sheetName ? (this.worksheets.getItem(options.sheetName) || this.worksheets.getActiveWorksheet()) : undefined;
    const sheets = selected ? [selected] : this.worksheets.items;
    return {
      kind: "workbookLayout",
      id: this.id,
      activeSheet: this.worksheets.items[0]?.name,
      sheetCount: this.worksheets.items.length,
      sheets: sheets.map((sheet) => sheet.layoutJson(options)),
    };
  }

  async render(options = {}) {
    const format = String(options.format || "").trim().toLowerCase();
    if (format === "layout" || format === LAYOUT_MIME) {
      return new FileBlob(JSON.stringify(this.layoutJson(options), null, 2), {
        type: LAYOUT_MIME,
        metadata: { artifactKind: "workbook", format: "layout", sheetName: options.sheetName, range: options.range },
      });
    }
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
    this.charts = new WorksheetChartCollection(this);
    this.shapes = [];
    this.images = new WorksheetImageCollection(this);
    this.tables = new WorksheetTableCollection(this);
    this.sparklineGroups = new SparklineGroupCollection(this);
    this.sparklines = this.sparklineGroups;
    this.dataValidations = new WorksheetRuleCollection(this, "dataValidation");
    this.conditionalFormattings = new WorksheetRuleCollection(this, "conditionalFormat");
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
    this.charts.deleteAll();
    this.images.deleteAll();
    this.shapes = [];
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
    if (this.workbook && !this.workbook._recalculating) return this.workbook.recalculate();
    for (const [address, cell] of this.store.entries()) {
      if (cell.formula) cell.value = evaluateFormula(this, cell.formula, address);
    }
    return undefined;
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
      const graphNode = options.graph?.nodes?.find((node) => node.sheet === this.name && node.address === address);
      records.push({
        kind: "formula",
        sheet: this.name,
        address,
        formula: cell.formula,
        value: cell.value,
        precedents: graphNode?.precedents?.map((ref) => ref.key) || formulaReferences(cell.formula, this).map((ref) => formulaCellKey(ref.sheetName || this.name, ref.address)),
        dependents: graphNode?.dependents || [],
        error: formulaErrorCode(cell.value) || undefined,
        circular: graphNode?.circular || undefined,
      });
    }
    return records;
  }

  styleRecords(options = {}) {
    const bounds = options.range ? parseRangeAddress(options.range) : this.usedBounds();
    const records = [];
    for (const [address, cell] of this.store.entries()) {
      const coord = parseCellAddress(address);
      if (coord.row < bounds.top || coord.row > bounds.bottom || coord.col < bounds.left || coord.col > bounds.right) continue;
      if (!cell.style || Object.keys(cell.style).length === 0) continue;
      records.push({ kind: "style", sheet: this.name, address, style: { ...cell.style } });
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

  layoutJson(options = {}) {
    this.recalculate();
    const bounds = options.range ? parseRangeAddress(options.range) : this.usedBounds();
    const cellW = Number(options.cellWidthPx || options.cellW || 96);
    const cellH = Number(options.cellHeightPx || options.cellH || 28);
    const frameForBounds = (rangeBounds) => ({
      left: 40 + (rangeBounds.left - bounds.left) * cellW,
      top: 40 + (rangeBounds.top - bounds.top) * cellH,
      width: Math.max(cellW, rangeBounds.colCount * cellW),
      height: Math.max(cellH, rangeBounds.rowCount * cellH),
    });
    const frameForRange = (range) => {
      try { return frameForBounds(parseRangeAddress(range)); } catch { return undefined; }
    };
    const cells = [];
    for (let r = bounds.top; r <= bounds.bottom; r += 1) {
      for (let c = bounds.left; c <= bounds.right; c += 1) {
        const address = makeCellAddress(r, c);
        const cell = this.store.get(address);
        cells.push({
          kind: "cell",
          sheet: this.name,
          address,
          row: r,
          col: c,
          bbox: [40 + (c - bounds.left) * cellW, 40 + (r - bounds.top) * cellH, cellW, cellH],
          value: cell.value,
          formula: cell.formula || undefined,
          style: cell.style && Object.keys(cell.style).length ? { ...cell.style } : undefined,
        });
      }
    }
    const tables = this.tables.items.map((table) => ({ kind: "table", id: table.id, sheet: this.name, name: table.name, address: table.range, bbox: Object.values(frameForRange(table.range) || {}), rows: table.rowCount, cols: table.columnCount, hasHeaders: table.hasHeaders }));
    const charts = this.charts.items.map((chart) => ({ kind: "chart", id: chart.id, sheet: this.name, name: chart.name, chartType: chart.type, title: chart.title, bbox: [chart.position.left, chart.position.top, chart.position.width, chart.position.height], series: chart.series.items.length, categories: chart.categories }));
    const images = this.images.items.map((image) => { const p = image.position; return { kind: "image", id: image.id, sheet: this.name, name: image.name, alt: image.alt, bbox: [p.left, p.top, p.width, p.height], fit: image.fit, hasDataUrl: Boolean(image.dataUrl), uri: image.uri, prompt: image.prompt }; });
    const sparklines = this.sparklineGroups.items.map((group) => { const p = group.targetFrame(bounds); return { kind: "sparkline", id: group.id, sheet: this.name, type: group.type, targetRange: group.targetRange.address, sourceData: group.sourceData.address, bbox: [p.left, p.top, p.width, p.height], values: group.values() }; });
    const rules = [...this.dataValidations.items, ...this.conditionalFormattings.items].map((rule) => ({ kind: rule.kind, id: rule.id, sheet: this.name, range: rule.range, bbox: Object.values(frameForRange(rule.range) || {}), ruleType: rule.ruleType, rule: rule.rule }));
    const drawingFrames = [...charts.map((item) => item.bbox), ...images.map((item) => item.bbox), ...sparklines.map((item) => item.bbox), ...tables.map((item) => item.bbox)].filter((bbox) => bbox.length === 4);
    const width = Math.max(320, bounds.colCount * cellW + 80, ...drawingFrames.map((bbox) => bbox[0] + bbox[2] + 40));
    const height = Math.max(180, bounds.rowCount * cellH + 80, ...drawingFrames.map((bbox) => bbox[1] + bbox[3] + 40));
    return {
      kind: "worksheetLayout",
      id: this.id,
      name: this.name,
      bounds: { ...bounds, address: rangeToAddress(bounds) },
      unit: "px",
      origin: { left: 40, top: 40 },
      cell: { width: cellW, height: cellH },
      width,
      height,
      cells,
      tables,
      charts,
      images,
      sparklines,
      rules,
    };
  }

  toSvg() {
    this.recalculate();
    const bounds = this.usedBounds();
    const cellW = 96;
    const cellH = 28;
    const imageFrames = this.images.items.map((image) => image.position);
    const sparklineFrames = this.sparklineGroups.items.map((group) => group.targetFrame(bounds));
    const width = Math.max(320, bounds.colCount * cellW + 80, ...this.charts.items.map((chart) => chart.position.left + chart.position.width + 40), ...imageFrames.map((frame) => frame.left + frame.width + 40), ...sparklineFrames.map((frame) => frame.left + frame.width + 40));
    const height = Math.max(180, bounds.rowCount * cellH + 80, ...this.charts.items.map((chart) => chart.position.top + chart.position.height + 40), ...imageFrames.map((frame) => frame.top + frame.height + 40), ...sparklineFrames.map((frame) => frame.top + frame.height + 40));
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
    const tableOverlays = this.tables.items.map((table) => table.toSvg(bounds)).join("");
    const chartOverlays = this.charts.items.map((chart) => chart.toSvg()).join("");
    const imageOverlays = this.images.items.map((image) => image.toSvg()).join("");
    const sparklineOverlays = this.sparklineGroups.items.map((group) => group.toSvg(bounds)).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="#f6f8fa"/>${rows.join("")}${tableOverlays}${chartOverlays}${imageOverlays}${sparklineOverlays}</svg>`;
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

  get format() {
    const first = this.worksheet.store.get(makeCellAddress(this.bounds.top, this.bounds.left));
    return { ...(first.style || {}) };
  }

  set format(style) {
    this.writeStyle(style);
  }

  get style() { return this.format; }
  set style(value) { this.format = value; }
  setFormat(style = {}) { this.writeStyle(style); return this; }

  writeStyle(style = {}) {
    for (let r = this.bounds.top; r <= this.bounds.bottom; r++) {
      for (let c = this.bounds.left; c <= this.bounds.right; c++) {
        const cell = this.worksheet.store.get(makeCellAddress(r, c));
        cell.style = { ...(cell.style || {}), ...(style || {}) };
      }
    }
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

  get dataValidation() {
    const address = rangeToAddress(this.bounds);
    return this.worksheet.dataValidations.items.find((item) => item.range === address)?.rule ?? null;
  }

  set dataValidation(value) {
    const address = rangeToAddress(this.bounds);
    this.worksheet.dataValidations.items = this.worksheet.dataValidations.items.filter((item) => item.range !== address);
    if (value != null) this.worksheet.dataValidations.add({ range: address, ...value });
  }

  get conditionalFormats() { return new RangeConditionalFormatFacade(this); }

  get sparklines() { return new RangeSparklineFacade(this); }

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

function findWorkbookTable(sheet, tableName) {
  const workbook = sheet?.workbook;
  for (const candidateSheet of workbook?.worksheets || [sheet].filter(Boolean)) {
    const table = candidateSheet.tables.items.find((item) => item.name === tableName || item.id === tableName);
    if (table) return { sheet: candidateSheet, table };
  }
  return undefined;
}

function formulaDefinedNameRange(sheet, refText = "", seen = new Set()) {
  const raw = String(refText || "").trim();
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(raw)) return undefined;
  const workbook = sheet?.workbook;
  const item = workbook?.definedNames.getItem(raw, sheet?.name) || workbook?.definedNames.getItem(raw);
  if (!item) return undefined;
  if (seen.has(item.name)) return { missing: true, name: item.name, refersTo: item.refersTo };
  const target = String(item.refersTo || "").replace(/^=/, "").trim();
  const ref = formulaRefParts(target);
  if (ref) return { ...ref, name: item.name, id: item.id, refersTo: item.refersTo };
  return { missing: true, name: item.name, refersTo: item.refersTo };
}

function formulaStructuredRefRange(sheet, refText = "") {
  const match = /^([A-Za-z_][A-Za-z0-9_.]*)\[([^\]]+)\]$/.exec(String(refText || "").trim());
  if (!match) return undefined;
  const tableName = match[1];
  const columnName = match[2].trim();
  if (!columnName || columnName.startsWith("#")) return { missing: true, tableName, columnName };
  const found = findWorkbookTable(sheet, tableName);
  if (!found) return { missing: true, tableName, columnName };
  const { table, sheet: tableSheet } = found;
  const headers = table.hasHeaders ? (table.values[0] || []) : Array.from({ length: table.columnCount || 0 }, (_, index) => `Column${index + 1}`);
  const columnIndex = headers.findIndex((header) => String(header ?? "").trim() === columnName);
  if (columnIndex < 0) return { missing: true, tableName, columnName, sheetName: tableSheet.name };
  const bounds = parseRangeAddress(table.range);
  const top = bounds.top + (table.showHeaders ? 1 : 0);
  const bottom = bounds.bottom;
  if (top > bottom) return { sheetName: tableSheet.name, start: makeCellAddress(top, bounds.left + columnIndex), end: makeCellAddress(top - 1, bounds.left + columnIndex), empty: true, tableName, columnName, table, columnIndex };
  return { sheetName: tableSheet.name, start: makeCellAddress(top, bounds.left + columnIndex), end: makeCellAddress(bottom, bounds.left + columnIndex), tableName, columnName, table, columnIndex };
}

function formulaReferences(formula, sheet) {
  const raw = String(formula || "");
  const refs = [];
  const consumed = [];
  const structuredRegex = /\b([A-Za-z_][A-Za-z0-9_.]*)\[([^\]]+)\]/g;
  for (const match of raw.matchAll(structuredRegex)) {
    consumed.push([match.index, match.index + match[0].length]);
    const structured = formulaStructuredRefRange(sheet, match[0]);
    if (!structured || structured.missing || structured.empty) continue;
    const start = parseCellAddress(structured.start);
    const end = parseCellAddress(structured.end);
    for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
      for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
        refs.push({ sheetName: structured.sheetName, address: makeCellAddress(row, col), structuredRef: match[0], tableName: structured.tableName, columnName: structured.columnName });
      }
    }
  }
  const rangeRegex = /(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]+\$?\d+)\s*:\s*(\$?[A-Za-z]+\$?\d+)/g;
  for (const match of raw.matchAll(rangeRegex)) {
    consumed.push([match.index, match.index + match[0].length]);
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
    if (consumed.some(([start, end]) => match.index >= start && match.index < end)) continue;
    refs.push({ sheetName: match[1] || match[2] || undefined, address: match[3].replaceAll("$", "").toUpperCase() });
  }
  for (const definedName of sheet?.workbook?.definedNames.items || []) {
    const escaped = definedName.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(`\\b${escaped}\\b`, "g");
    for (const match of raw.matchAll(re)) {
      if (consumed.some(([start, end]) => match.index >= start && match.index < end)) continue;
      const resolved = formulaDefinedNameRange(sheet, definedName.name);
      if (!resolved || resolved.missing) continue;
      const start = parseCellAddress(resolved.start);
      const end = parseCellAddress(resolved.end);
      for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
        for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) refs.push({ sheetName: resolved.sheetName, address: makeCellAddress(row, col), definedName: definedName.name });
      }
    }
  }
  const seen = new Set();
  return refs.filter((ref) => {
    const key = `${ref.sheetName || ""}!${ref.address}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function formulaCellKey(sheetName, address) {
  return `${String(sheetName || "").replaceAll("'", "''")}!${String(address || "").replaceAll("$", "").toUpperCase()}`;
}

function formulaCellId(sheetName, address) {
  return `fx/${encodeURIComponent(formulaCellKey(sheetName, address))}`;
}

function displayFormulaRef(sheetName, address) {
  const safeSheet = String(sheetName || "").includes(" ") ? `'${String(sheetName).replaceAll("'", "''")}'` : String(sheetName || "");
  return `${safeSheet}!${String(address || "").replaceAll("$", "").toUpperCase()}`;
}

function formulaErrorCode(value) {
  const text = String(value ?? "");
  return /^#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A|CYCLE!)/.test(text) ? text.match(/^#[A-Z0-9\/?!]+/)?.[0] : undefined;
}

function publicFormulaNode(node) {
  return {
    kind: "formulaNode",
    id: node.id,
    key: node.key,
    sheet: node.sheet,
    address: node.address,
    formula: node.formula,
    value: node.cell?.value,
    precedents: node.precedents.map((ref) => ({ ...ref })),
    dependents: [...node.dependents],
    circular: node.circular || undefined,
    error: formulaErrorCode(node.cell?.value) || undefined,
  };
}

function buildWorkbookFormulaGraph(workbook) {
  const nodes = [];
  const nodeByKey = new Map();
  const errors = [];
  for (const sheet of workbook.worksheets) {
    for (const [address, cell] of sheet.store.entries()) {
      if (!cell.formula) continue;
      const key = formulaCellKey(sheet.name, address);
      const node = { kind: "formulaNode", id: formulaCellId(sheet.name, address), key, sheet: sheet.name, sheetObject: sheet, address, cell, formula: cell.formula, precedents: [], dependents: [], circular: false };
      nodes.push(node);
      nodeByKey.set(key, node);
    }
  }

  const edges = [];
  for (const node of nodes) {
    for (const ref of formulaReferences(node.formula, node.sheetObject)) {
      const sheetName = ref.sheetName || node.sheet;
      const targetSheet = workbook.worksheets.getItem(sheetName);
      const targetAddress = String(ref.address || "").replaceAll("$", "").toUpperCase();
      const targetKey = formulaCellKey(sheetName, targetAddress);
      const precedent = {
        sheet: sheetName,
        address: targetAddress,
        key: targetKey,
        missing: !targetSheet || undefined,
        hasFormula: nodeByKey.has(targetKey) || undefined,
      };
      node.precedents.push(precedent);
      edges.push({ kind: "formulaEdge", from: node.key, to: targetKey, fromSheet: node.sheet, fromAddress: node.address, toSheet: sheetName, toAddress: targetAddress, missing: !targetSheet || undefined });
      const targetNode = nodeByKey.get(targetKey);
      if (targetNode) targetNode.dependents.push(node.key);
      if (!targetSheet) errors.push({ kind: "formulaGraphError", type: "missingSheet", from: node.key, sheet: sheetName, address: targetAddress, ref: displayFormulaRef(sheetName, targetAddress) });
    }
  }

  const cycles = detectFormulaCycles(nodes, nodeByKey);
  const cycleKeys = new Set(cycles.flatMap((cycle) => cycle.keys));
  for (const node of nodes) if (cycleKeys.has(node.key)) node.circular = true;
  for (const node of nodes) {
    const error = formulaErrorCode(node.cell?.value);
    if (error) errors.push({ kind: "formulaGraphError", type: "formulaError", key: node.key, sheet: node.sheet, address: node.address, value: error });
  }
  return { kind: "formulaGraph", nodes, edges, cycles, errors };
}

function detectFormulaCycles(nodes, nodeByKey) {
  const cycles = [];
  const emitted = new Set();
  const visiting = new Set();
  const visited = new Set();
  const stack = [];

  const emitCycle = (keys) => {
    if (keys.length === 0) return;
    const canonical = [...keys].sort().join("|");
    if (emitted.has(canonical)) return;
    emitted.add(canonical);
    cycles.push({ kind: "formulaCycle", keys, path: [...keys, keys[0]] });
  };

  const visit = (node) => {
    if (visiting.has(node.key)) {
      const start = stack.indexOf(node.key);
      emitCycle(stack.slice(start));
      return;
    }
    if (visited.has(node.key)) return;
    visiting.add(node.key);
    stack.push(node.key);
    for (const ref of node.precedents) {
      const target = nodeByKey.get(ref.key);
      if (target) visit(target);
    }
    stack.pop();
    visiting.delete(node.key);
    visited.add(node.key);
  };

  for (const node of nodes) visit(node);
  return cycles;
}

function formulaGraphRecords(graph, options = {}) {
  const kinds = options.kinds || normalizeKinds(options.kind, ["formulaGraph", "formulaNode", "formulaEdge", "formulaCycle"]);
  const records = [];
  if (kinds.has("formulaGraph")) {
    records.push({ kind: "formulaGraph", formulas: graph.nodes.length, edges: graph.edges.length, cycles: graph.cycles.length, errors: graph.errors.length });
  }
  if (kinds.has("formulaNode")) records.push(...graph.nodes.map(publicFormulaNode));
  if (kinds.has("formulaEdge")) records.push(...graph.edges.map((edge) => ({ ...edge })));
  if (kinds.has("formulaCycle")) records.push(...graph.cycles.map((cycle) => ({ ...cycle })));
  return records;
}

function splitFormulaArgs(text = "") {
  const args = [];
  let current = "";
  let depth = 0;
  let inString = false;
  for (let i = 0; i < String(text).length; i++) {
    const ch = String(text)[i];
    if (ch === '"') {
      current += ch;
      if (String(text)[i + 1] === '"') { current += '"'; i += 1; continue; }
      inString = !inString;
      continue;
    }
    if (!inString && ch === "(") depth += 1;
    if (!inString && ch === ")") depth -= 1;
    if (!inString && ch === "," && depth === 0) { args.push(current.trim()); current = ""; continue; }
    current += ch;
  }
  if (current.trim() || text === "") args.push(current.trim());
  return args;
}

function formulaUnquote(value) {
  const text = String(value ?? "").trim();
  if (/^"(?:[^"]|"")*"$/.test(text)) return text.slice(1, -1).replaceAll('""', '"');
  return undefined;
}

function formulaRefParts(refText = "") {
  const raw = String(refText || "").trim();
  const match = /^(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]+\$?\d+)(?::(\$?[A-Za-z]+\$?\d+))?$/.exec(raw);
  if (!match) return undefined;
  return { sheetName: match[1] || match[2] || undefined, start: match[3].replaceAll("$", "").toUpperCase(), end: (match[4] || match[3]).replaceAll("$", "").toUpperCase() };
}

function formulaRangeMatrix(sheet, refText, context = {}) {
  const structured = formulaStructuredRefRange(sheet, refText);
  if (structured) {
    if (structured.missing) return [["#REF!"]];
    if (structured.empty) return [];
    const targetSheet = structured.sheetName ? sheet.workbook?.worksheets.getItem(structured.sheetName) : sheet;
    if (!targetSheet) return [["#REF!"]];
    const start = parseCellAddress(structured.start);
    const end = parseCellAddress(structured.end);
    const rows = [];
    for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
      const values = [];
      for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
        const address = makeCellAddress(row, col);
        let value = context.getValue ? context.getValue({ sheetName: structured.sheetName, address, structuredRef: refText, tableName: structured.tableName, columnName: structured.columnName }) : targetSheet.store.get(address).value;
        if ((value == null || value === "") && structured.table) {
          const tableBounds = parseRangeAddress(structured.table.range);
          const tableRow = row - tableBounds.top;
          value = structured.table.values[tableRow]?.[structured.columnIndex] ?? value;
        }
        values.push(value);
      }
      rows.push(values);
    }
    return rows;
  }
  const defined = formulaDefinedNameRange(sheet, refText);
  if (defined) {
    if (defined.missing) return [["#REF!"]];
    return formulaRangeMatrix(sheet, `${defined.sheetName ? `${defined.sheetName}!` : ""}${defined.start}${defined.end && defined.end !== defined.start ? `:${defined.end}` : ""}`, context);
  }
  const ref = formulaRefParts(refText);
  if (!ref) return undefined;
  const targetSheet = ref.sheetName ? sheet.workbook?.worksheets.getItem(ref.sheetName) : sheet;
  if (!targetSheet) return [["#REF!"]];
  const start = parseCellAddress(ref.start);
  const end = parseCellAddress(ref.end);
  const rows = [];
  for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
    const values = [];
    for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
      const address = makeCellAddress(row, col);
      values.push(context.getValue ? context.getValue({ sheetName: ref.sheetName, address }) : targetSheet.store.get(address).value);
    }
    rows.push(values);
  }
  return rows;
}

function formulaReferenceValues(sheet, refText, context = {}) {
  const matrix = formulaRangeMatrix(sheet, refText, context);
  if (matrix) return matrix.flat();
  const scalar = formulaScalar(sheet, refText, context);
  return scalar === undefined ? [] : [scalar];
}

function formulaScalar(sheet, expr, context = {}) {
  const text = String(expr ?? "").trim();
  const quoted = formulaUnquote(text);
  if (quoted !== undefined) return quoted;
  if (/^-?\d+(?:\.\d+)?$/.test(text)) return Number(text);
  if (/^TRUE$/i.test(text)) return true;
  if (/^FALSE$/i.test(text)) return false;
  const matrix = formulaRangeMatrix(sheet, text, context);
  if (matrix) return matrix[0]?.[0] ?? null;
  if (/^[A-Z][A-Z0-9.]*\(/i.test(text)) return evaluateFormula(sheet, `=${text}`, undefined, context);
  return undefined;
}

function formulaNumber(value) {
  if (formulaErrorCode(value)) return value;
  if (value === true) return 1;
  if (value === false || value == null || value === "") return 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function formulaText(value) {
  if (value == null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function formulaTruthy(value) {
  if (formulaErrorCode(value)) return false;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  const text = String(value ?? "").trim();
  if (/^TRUE$/i.test(text)) return true;
  if (/^FALSE$/i.test(text) || text === "") return false;
  return Boolean(Number(text));
}

function evaluateFormulaCondition(sheet, expr, context = {}) {
  const text = String(expr || "").trim();
  const comparison = /^(.*?)\s*(>=|<=|<>|=|>|<)\s*(.*?)$/.exec(text);
  if (comparison) {
    const left = formulaScalar(sheet, comparison[1], context);
    const right = formulaScalar(sheet, comparison[3], context);
    const leftNum = Number(left), rightNum = Number(right);
    const numeric = Number.isFinite(leftNum) && Number.isFinite(rightNum);
    const a = numeric ? leftNum : formulaText(left);
    const b = numeric ? rightNum : formulaText(right);
    switch (comparison[2]) {
      case ">=": return a >= b;
      case "<=": return a <= b;
      case "<>": return a !== b;
      case "=": return a === b;
      case ">": return a > b;
      case "<": return a < b;
    }
  }
  return formulaTruthy(formulaScalar(sheet, text, context));
}

function matchesCriteria(value, criteria) {
  const raw = formulaText(criteria).trim();
  const match = /^(>=|<=|<>|=|>|<)?\s*(.*)$/.exec(raw);
  const op = match?.[1] || "=";
  const expectedRaw = match?.[2] ?? raw;
  const actualNum = Number(value);
  const expectedNum = Number(expectedRaw);
  const numeric = Number.isFinite(actualNum) && Number.isFinite(expectedNum) && expectedRaw !== "";
  const actual = numeric ? actualNum : formulaText(value);
  const expected = numeric ? expectedNum : expectedRaw.replace(/^"|"$/g, "");
  switch (op) {
    case ">=": return actual >= expected;
    case "<=": return actual <= expected;
    case "<>": return actual !== expected;
    case ">": return actual > expected;
    case "<": return actual < expected;
    default: return actual === expected;
  }
}

function aggregateFormulaValues(values, fnName) {
  const errors = values.map(formulaErrorCode).filter(Boolean);
  if (errors.length) return errors[0];
  const nums = values.map(formulaNumber).filter((value) => Number.isFinite(value));
  if (fnName === "COUNT") return nums.length;
  if (nums.length === 0) return 0;
  if (fnName === "AVERAGE") return nums.reduce((acc, value) => acc + value, 0) / nums.length;
  if (fnName === "MIN") return Math.min(...nums);
  if (fnName === "MAX") return Math.max(...nums);
  return nums.reduce((acc, value) => acc + value, 0);
}

function evaluateFormulaFunction(sheet, fnName, args, context = {}) {
  const values = (parts = args) => parts.flatMap((part) => formulaReferenceValues(sheet, part, context));
  const scalar = (index, fallback = undefined) => {
    const value = formulaScalar(sheet, args[index], context);
    return value === undefined ? fallback : value;
  };
  switch (fnName) {
    case "SUM":
    case "AVERAGE":
    case "MIN":
    case "MAX":
    case "COUNT":
      return aggregateFormulaValues(values(), fnName);
    case "ABS": return Math.abs(formulaNumber(scalar(0, 0)));
    case "ROUND": return Number(formulaNumber(scalar(0, 0)).toFixed(Math.max(0, Number(scalar(1, 0)) || 0)));
    case "INT": return Math.floor(formulaNumber(scalar(0, 0)));
    case "CEILING": return Math.ceil(formulaNumber(scalar(0, 0)) / Math.max(1, formulaNumber(scalar(1, 1)))) * Math.max(1, formulaNumber(scalar(1, 1)));
    case "FLOOR": return Math.floor(formulaNumber(scalar(0, 0)) / Math.max(1, formulaNumber(scalar(1, 1)))) * Math.max(1, formulaNumber(scalar(1, 1)));
    case "IF": return evaluateFormulaCondition(sheet, args[0], context) ? scalar(1, true) : scalar(2, false);
    case "AND": return args.every((arg) => evaluateFormulaCondition(sheet, arg, context));
    case "OR": return args.some((arg) => evaluateFormulaCondition(sheet, arg, context));
    case "NOT": return !evaluateFormulaCondition(sheet, args[0], context);
    case "CONCAT":
    case "CONCATENATE": return values().map(formulaText).join("");
    case "TEXTJOIN": {
      const delimiter = formulaText(scalar(0, ""));
      const ignoreEmpty = formulaTruthy(scalar(1, false));
      const joined = values(args.slice(2)).map(formulaText).filter((value) => !ignoreEmpty || value !== "");
      return joined.join(delimiter);
    }
    case "LEFT": return formulaText(scalar(0, "")).slice(0, Number(scalar(1, 1)) || 1);
    case "RIGHT": { const text = formulaText(scalar(0, "")); const count = Number(scalar(1, 1)) || 1; return text.slice(Math.max(0, text.length - count)); }
    case "MID": { const text = formulaText(scalar(0, "")); const start = Math.max(1, Number(scalar(1, 1)) || 1); const count = Math.max(0, Number(scalar(2, 1)) || 1); return text.slice(start - 1, start - 1 + count); }
    case "LEN": return formulaText(scalar(0, "")).length;
    case "UPPER": return formulaText(scalar(0, "")).toUpperCase();
    case "LOWER": return formulaText(scalar(0, "")).toLowerCase();
    case "TRIM": return formulaText(scalar(0, "")).trim().replace(/\s+/g, " ");
    case "COUNTIF": { const range = values([args[0]]); const criteria = scalar(1, ""); return range.filter((value) => matchesCriteria(value, criteria)).length; }
    case "SUMIF": {
      const range = values([args[0]]);
      const criteria = scalar(1, "");
      const sumRange = args[2] ? values([args[2]]) : range;
      return range.reduce((sum, value, index) => sum + (matchesCriteria(value, criteria) ? formulaNumber(sumRange[index]) || 0 : 0), 0);
    }
    case "VLOOKUP": {
      const lookup = scalar(0, "");
      const matrix = formulaRangeMatrix(sheet, args[1], context) || [];
      const colIndex = Math.max(1, Number(scalar(2, 1)) || 1) - 1;
      const row = matrix.find((item) => formulaText(item[0]) === formulaText(lookup) || Number(item[0]) === Number(lookup));
      return row ? row[colIndex] ?? "#N/A" : "#N/A";
    }
    case "XLOOKUP": {
      const lookup = scalar(0, "");
      const lookupValues = values([args[1]]);
      const returnValues = values([args[2]]);
      const index = lookupValues.findIndex((value) => formulaText(value) === formulaText(lookup) || Number(value) === Number(lookup));
      return index >= 0 ? returnValues[index] : scalar(3, "#N/A");
    }
    default:
      return "#NAME?";
  }
}

function evaluateFormula(sheet, formula, _address, context = {}) {
  const raw = String(formula || "").trim();
  if (!raw.startsWith("=")) return raw;
  const expr = raw.slice(1).trim();
  const functionMatch = /^([A-Z][A-Z0-9.]*)\((.*)\)$/i.exec(expr);
  if (functionMatch) {
    return evaluateFormulaFunction(sheet, functionMatch[1].toUpperCase(), splitFormulaArgs(functionMatch[2]), context);
  }

  let replacementError;
  const safe = expr.replace(/(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]+\$?\d+)/g, (match, quotedSheet, bareSheet, address) => {
    const refSheetName = quotedSheet || bareSheet || undefined;
    const refAddress = address.replaceAll("$", "").toUpperCase();
    const targetSheet = refSheetName ? sheet.workbook?.worksheets.getItem(refSheetName) : sheet;
    const value = context.getValue ? context.getValue({ sheetName: refSheetName, address: refAddress }) : (targetSheet ? targetSheet.store.get(refAddress).value : "#REF!");
    const error = formulaErrorCode(value);
    if (error) replacementError = error;
    return Number(value) || 0;
  });
  if (replacementError) return replacementError;
  if (!/^[0-9+\-*/().\s]+$/.test(safe)) return "#NAME?";
  try {
    const value = Function(`"use strict"; return (${safe});`)();
    return Number.isFinite(value) ? value : "#DIV/0!";
  } catch {
    return "#VALUE!";
  }
}

function workbookMetadata(workbook) {
  return {
    version: 1,
    comments: workbook.comments.toJSON(),
    definedNames: workbook.definedNames.toJSON(),
    sheets: workbook.worksheets.items.map((sheet) => ({
      name: sheet.name,
      dataValidations: sheet.dataValidations.toJSON(),
      conditionalFormattings: sheet.conditionalFormattings.toJSON(),
      tables: sheet.tables.toJSON(),
      charts: sheet.charts.toJSON(),
      images: sheet.images.toJSON(),
      sparklineGroups: sheet.sparklineGroups.toJSON(),
    })),
  };
}

function applyWorkbookMetadata(workbook, metadata = {}) {
  workbook.definedNames.items = [];
  for (const item of metadata.definedNames || []) workbook.definedNames.add(item);
  if (metadata.comments?.self) workbook.comments.setSelf(metadata.comments.self);
  for (const threadData of metadata.comments?.threads || []) {
    const thread = new CommentThread(workbook, threadData.target, threadData.comments?.[0]?.text || "", threadData.author);
    thread.id = threadData.id || thread.id;
    thread.comments = threadData.comments || thread.comments;
    thread.resolved = Boolean(threadData.resolved);
    workbook.comments.threads.push(thread);
  }
  for (const sheetData of metadata.sheets || []) {
    const sheet = workbook.worksheets.getItem(sheetData.name);
    if (!sheet) continue;
    sheet.dataValidations.items = (sheetData.dataValidations || []).map((item) => ({ ...item }));
    sheet.conditionalFormattings.items = (sheetData.conditionalFormattings || []).map((item) => ({ ...item }));
    sheet.tables.items = [];
    for (const tableData of sheetData.tables || []) {
      const table = sheet.tables.add({ ...tableData });
      table.id = tableData.id || table.id;
      table.showHeaders = tableData.showHeaders ?? table.showHeaders;
      table.showTotals = Boolean(tableData.showTotals);
      table.showBandedColumns = Boolean(tableData.showBandedColumns);
      table.showFilterButton = tableData.showFilterButton ?? table.showFilterButton;
    }
    sheet.charts.items = [];
    for (const chartData of sheetData.charts || []) {
      const chart = sheet.charts.add(chartData.type || chartData.chartType || "bar", { ...chartData });
      chart.id = chartData.id || chart.id;
    }
    sheet.images.items = [];
    for (const imageData of sheetData.images || []) {
      const image = sheet.images.add({ ...imageData });
      image.id = imageData.id || image.id;
    }
    sheet.sparklineGroups.items = [];
    for (const sparklineData of sheetData.sparklineGroups || []) {
      const group = sheet.sparklineGroups.add({ ...sparklineData });
      group.id = sparklineData.id || group.id;
    }
  }
}

export class SpreadsheetFile {
  static async exportXlsx(workbook) {
    workbook.recalculate();
    const zip = new JSZip();
    const tableParts = collectWorkbookTableParts(workbook);
    const imageParts = collectWorkbookImageParts(workbook);
    const chartParts = collectWorkbookChartParts(workbook, imageParts);
    const threadParts = collectWorkbookThreadParts(workbook);
    const sharedStrings = collectWorkbookSharedStrings(workbook);
    const styleTable = collectWorkbookStyles(workbook);
    zip.file("[Content_Types].xml", xlsxContentTypes(workbook.worksheets.items.length, tableParts, imageParts, chartParts, threadParts, sharedStrings));
    zip.file("_rels/.rels", relsXml([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "xl/workbook.xml" }]));
    zip.file("xl/workbook.xml", workbookXml(workbook));
    zip.file("xl/_rels/workbook.xml.rels", workbookRelsXml(workbook.worksheets.items.length, threadParts.length > 0, sharedStrings.strings.length > 0));
    zip.file("xl/styles.xml", xlsxStylesXml(styleTable));
    if (sharedStrings.strings.length) zip.file("xl/sharedStrings.xml", sharedStringsXml(sharedStrings));
    zip.file("customXml/open-office-artifact.json", JSON.stringify(workbookMetadata(workbook), null, 2));
    workbook.worksheets.items.forEach((sheet, index) => {
      const sheetTableParts = tableParts.filter((part) => part.sheetIndex === index);
      const sheetImageParts = imageParts.filter((part) => part.sheetIndex === index);
      const sheetChartParts = chartParts.filter((part) => part.sheetIndex === index);
      const sheetThreadPart = threadParts.find((part) => part.sheetIndex === index);
      const drawingRelId = sheetImageParts.length || sheetChartParts.length ? `rId${sheetTableParts.length + 1}` : undefined;
      if (sheetThreadPart) sheetThreadPart.relId = `rId${sheetTableParts.length + (drawingRelId ? 1 : 0) + 1}`;
      zip.file(`xl/worksheets/sheet${index + 1}.xml`, worksheetXml(sheet, sheetTableParts, drawingRelId, sharedStrings, styleTable));
      if (sheetTableParts.length || sheetImageParts.length || sheetChartParts.length || sheetThreadPart) zip.file(`xl/worksheets/_rels/sheet${index + 1}.xml.rels`, worksheetRelsXml(sheetTableParts, drawingRelId ? { relId: drawingRelId, target: `../drawings/drawing${index + 1}.xml` } : undefined, sheetThreadPart));
      if (sheetImageParts.length || sheetChartParts.length) {
        zip.file(`xl/drawings/drawing${index + 1}.xml`, drawingXml(sheetImageParts, sheetChartParts));
        zip.file(`xl/drawings/_rels/drawing${index + 1}.xml.rels`, drawingRelsXml(sheetImageParts, sheetChartParts));
      }
    });
    tableParts.forEach((part) => zip.file(`xl/tables/table${part.tablePartId}.xml`, tableXml(part.table, part.tablePartId)));
    imageParts.forEach((part) => zip.file(`xl/media/image${part.imagePartId}.${part.extension}`, part.bytes));
    chartParts.forEach((part) => zip.file(`xl/charts/chart${part.chartPartId}.xml`, xlsxChartXml(part.chart)));
    threadParts.forEach((part) => zip.file(`xl/threadedComments/threadedComment${part.threadPartId}.xml`, threadedCommentsXml(part)));
    if (threadParts.length) zip.file("xl/persons/person.xml", personsXml(threadParts));
    const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return new FileBlob(bytes, { type: XLSX_MIME });
  }

  static async importXlsx(blobOrBuffer) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const workbook = Workbook.create();
    const sharedStrings = parseSharedStringsXml(await zip.file("xl/sharedStrings.xml")?.async("text"));
    const styles = parseXlsxStylesXml(await zip.file("xl/styles.xml")?.async("text"));
    const workbookText = await zip.file("xl/workbook.xml")?.async("text");
    const sheetNames = [...String(workbookText || "").matchAll(/<sheet[^>]*name="([^"]+)"[^>]*sheetId="(\d+)"/g)].map((m) => ({ name: decodeXml(m[1]), index: Number(m[2]) }));
    for (const { name, index } of sheetNames.length ? sheetNames : [{ name: "Sheet1", index: 1 }]) {
      const sheet = workbook.worksheets.add(name);
      const xml = await zip.file(`xl/worksheets/sheet${index}.xml`)?.async("text");
      if (xml) parseWorksheetXml(sheet, xml, { sharedStrings, styles });
    }
    parseWorkbookDefinedNames(workbook, workbookText);
    const metadataText = await zip.file("customXml/open-office-artifact.json")?.async("text");
    if (metadataText) applyWorkbookMetadata(workbook, JSON.parse(metadataText));
    else await importNativeThreadedComments(workbook, zip, sheetNames.length ? sheetNames : [{ name: "Sheet1", index: 1 }]);
    workbook.recalculate();
    return workbook;
  }
}

function collectWorkbookTableParts(workbook) {
  const parts = [];
  let tablePartId = 1;
  workbook.worksheets.items.forEach((sheet, sheetIndex) => {
    sheet.tables.items.forEach((table, tableIndex) => {
      parts.push({ sheet, sheetIndex, table, tableIndex, tablePartId: tablePartId++, relId: `rId${tableIndex + 1}` });
    });
  });
  return parts;
}

function imageDataFromDataUrl(dataUrl) {
  const match = /^data:([^;,]+);base64,(.+)$/i.exec(String(dataUrl || ""));
  if (!match) return undefined;
  const contentType = match[1].toLowerCase();
  const extension = contentType === "image/jpeg" ? "jpg" : contentType.split("/")[1] || "bin";
  return { contentType, extension: extension === "svg+xml" ? "svg" : extension, bytes: Buffer.from(match[2], "base64") };
}

function imageContentTypeDefaults(imageParts = []) {
  return [
    ["png", "image/png"],
    ["jpg", "image/jpeg"],
    ["jpeg", "image/jpeg"],
    ["gif", "image/gif"],
    ["svg", "image/svg+xml"],
  ].filter(([extension]) => imageParts.some((part) => part.extension === extension)).map(([extension, contentType]) => `<Default Extension="${extension}" ContentType="${contentType}"/>`).join("");
}

function imageContentTypeFromExtension(extension) {
  const normalized = String(extension || "").toLowerCase().replace(/^\./, "");
  if (normalized === "jpg" || normalized === "jpeg") return "image/jpeg";
  if (normalized === "png") return "image/png";
  if (normalized === "gif") return "image/gif";
  if (normalized === "svg" || normalized === "svg+xml") return "image/svg+xml";
  return "application/octet-stream";
}

function collectWorkbookImageParts(workbook) {
  const parts = [];
  let imagePartId = 1;
  workbook.worksheets.items.forEach((sheet, sheetIndex) => {
    let drawingRelIndex = 1;
    sheet.images.items.forEach((image, imageIndex) => {
      const data = imageDataFromDataUrl(image.dataUrl);
      if (!data) return;
      parts.push({
        sheet,
        sheetIndex,
        image,
        imageIndex,
        imagePartId: imagePartId++,
        drawingRelId: `rId${drawingRelIndex++}`,
        ...data,
      });
    });
  });
  return parts;
}

function collectWorkbookChartParts(workbook, imageParts = []) {
  const parts = [];
  let chartPartId = 1;
  workbook.worksheets.items.forEach((sheet, sheetIndex) => {
    let drawingRelIndex = imageParts.filter((part) => part.sheetIndex === sheetIndex).length + 1;
    sheet.charts.items.forEach((chart, chartIndex) => {
      parts.push({ sheet, sheetIndex, chart, chartIndex, chartPartId: chartPartId++, drawingRelId: `rId${drawingRelIndex++}` });
    });
  });
  return parts;
}

function collectWorkbookThreadParts(workbook) {
  const parts = [];
  let threadPartId = 1;
  workbook.worksheets.items.forEach((sheet, sheetIndex) => {
    const threads = workbook.comments.threads.filter((thread) => (thread.target.sheetName || sheet.name) === sheet.name);
    if (threads.length) parts.push({ sheet, sheetIndex, threads, threadPartId: threadPartId++, relId: undefined });
  });
  return parts;
}

function collectWorkbookSharedStrings(workbook) {
  const strings = [];
  const indexByText = new Map();
  for (const sheet of workbook.worksheets) {
    for (const [, cell] of sheet.store.entries()) {
      if (cell.formula || typeof cell.value !== "string") continue;
      if (!indexByText.has(cell.value)) {
        indexByText.set(cell.value, strings.length);
        strings.push(cell.value);
      }
    }
  }
  return { strings, indexByText };
}

function sharedStringsXml(sharedStrings) {
  const items = sharedStrings.strings.map((value) => `<si><t>${xmlEscape(value)}</t></si>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.strings.length}" uniqueCount="${sharedStrings.strings.length}">${items}</sst>`;
}

function parseSharedStringsXml(xml = "") {
  return [...String(xml || "").matchAll(/<si>([\s\S]*?)<\/si>/g)].map((match) => decodeXml([...match[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((text) => text[1]).join("")));
}

function normalizeXlsxColor(value, fallback = "000000") {
  const raw = String(value || fallback).trim();
  const token = resolveColorToken(raw, raw);
  return String(token || fallback).replace(/^#/, "").replace(/^FF/i, "").slice(0, 6).padEnd(6, "0").toUpperCase();
}

function normalizeXlsxStyle(style = {}) {
  const font = style.font || {};
  return {
    font: {
      bold: Boolean(style.bold ?? font.bold),
      italic: Boolean(style.italic ?? font.italic),
      color: style.fontColor || font.color || style.color || undefined,
      size: Number(style.fontSize || font.size || 11),
      name: style.fontFamily || font.name || "Aptos",
    },
    fill: style.fill || style.backgroundColor || style.fillColor || undefined,
    numberFormat: style.numberFormat || style.numFmt || undefined,
  };
}

function xlsxStyleKey(style = {}) {
  const normalized = normalizeXlsxStyle(style);
  if (!normalized.font.bold && !normalized.font.italic && !normalized.font.color && normalized.font.size === 11 && normalized.font.name === "Aptos" && !normalized.fill && !normalized.numberFormat) return "";
  return JSON.stringify(normalized);
}

function collectWorkbookStyles(workbook) {
  const styles = [{}];
  const indexByKey = new Map([["", 0]]);
  for (const sheet of workbook.worksheets) {
    for (const [, cell] of sheet.store.entries()) {
      const key = xlsxStyleKey(cell.style || {});
      if (!indexByKey.has(key)) {
        indexByKey.set(key, styles.length);
        styles.push(normalizeXlsxStyle(cell.style || {}));
      }
    }
  }
  return { styles, indexByKey };
}

function xlsxStyleIndex(cell, styleTable) {
  return styleTable.indexByKey.get(xlsxStyleKey(cell.style || {})) || 0;
}

function xlsxFontXml(style = {}) {
  const font = normalizeXlsxStyle(style).font;
  return `<font>${font.bold ? "<b/>" : ""}${font.italic ? "<i/>" : ""}<sz val="${Number(font.size) || 11}"/><color rgb="FF${normalizeXlsxColor(font.color, "000000")}"/><name val="${attrEscape(font.name || "Aptos")}"/></font>`;
}

function xlsxFillXml(style = {}) {
  const fill = normalizeXlsxStyle(style).fill;
  if (!fill) return `<fill><patternFill patternType="none"/></fill>`;
  const color = normalizeXlsxColor(fill, "FFFFFF");
  return `<fill><patternFill patternType="solid"><fgColor rgb="FF${color}"/><bgColor indexed="64"/></patternFill></fill>`;
}

function xlsxStylesXml(styleTable) {
  const styles = styleTable.styles || [{}];
  const customFormats = new Map();
  styles.forEach((style) => { if (style.numberFormat && !customFormats.has(style.numberFormat)) customFormats.set(style.numberFormat, 164 + customFormats.size); });
  const numFmts = customFormats.size ? `<numFmts count="${customFormats.size}">${[...customFormats.entries()].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${attrEscape(code)}"/>`).join("")}</numFmts>` : "";
  const fonts = styles.map((style, index) => index === 0 ? `<font><sz val="11"/><name val="Aptos"/></font>` : xlsxFontXml(style)).join("");
  const fills = [`<fill><patternFill patternType="none"/></fill>`, `<fill><patternFill patternType="gray125"/></fill>`, ...styles.slice(1).map(xlsxFillXml)].join("");
  const xfs = styles.map((style, index) => {
    if (index === 0) return `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`;
    const normalized = normalizeXlsxStyle(style);
    const numFmtId = normalized.numberFormat ? customFormats.get(normalized.numberFormat) : 0;
    const fillId = normalized.fill ? index + 1 : 0;
    return `<xf numFmtId="${numFmtId}" fontId="${index}" fillId="${fillId}" borderId="0" xfId="0"${numFmtId ? ` applyNumberFormat="1"` : ""} applyFont="1"${fillId ? ` applyFill="1"` : ""}/>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts}<fonts count="${styles.length}">${fonts}</fonts><fills count="${styles.length + 1}">${fills}</fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${styles.length}">${xfs}</cellXfs></styleSheet>`;
}

function parseAttrs(attrs = "") {
  return Object.fromEntries([...String(attrs || "").matchAll(/\b([A-Za-z_:][\w:.-]*)="([^"]*)"/g)].map((match) => [match[1], decodeXml(match[2])]));
}

function parseXlsxStylesXml(xml = "") {
  const text = String(xml || "");
  const numFmtById = new Map([...text.matchAll(/<numFmt\b([^>]*)\/>/g)].map((match) => { const attrs = parseAttrs(match[1]); return [Number(attrs.numFmtId), attrs.formatCode]; }));
  const fontsBody = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(text)?.[1] || "";
  const fonts = [...fontsBody.matchAll(/<font>([\s\S]*?)<\/font>/g)].map((match) => ({
    bold: /<b\b/.test(match[1]),
    italic: /<i\b/.test(match[1]),
    color: /<color[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(match[1])?.[1] ? `#${/<color[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(match[1])?.[1]}` : undefined,
    size: Number(/<sz[^>]*val="([^"]+)"/.exec(match[1])?.[1] || 11),
    name: /<name[^>]*val="([^"]+)"/.exec(match[1])?.[1] || "Aptos",
  }));
  const fillsBody = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(text)?.[1] || "";
  const fills = [...fillsBody.matchAll(/<fill>([\s\S]*?)<\/fill>/g)].map((match) => /<fgColor[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(match[1])?.[1]).map((color) => color ? `#${color}` : undefined);
  const xfsBody = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(text)?.[1] || "";
  return [...xfsBody.matchAll(/<xf\b([^>]*)\/?>(?:<\/xf>)?/g)].map((match) => {
    const attrs = parseAttrs(match[1]);
    const font = fonts[Number(attrs.fontId || 0)] || {};
    const fill = fills[Number(attrs.fillId || 0)];
    const numberFormat = numFmtById.get(Number(attrs.numFmtId || 0));
    const style = { font: { ...font } };
    if (fill) style.fill = fill;
    if (numberFormat) style.numberFormat = numberFormat;
    return style;
  });
}

function xlsxContentTypes(sheetCount, tableParts = [], imageParts = [], chartParts = [], threadParts = [], sharedStrings = { strings: [] }) {
  const sheets = Array.from({ length: sheetCount }, (_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("");
  const tables = tableParts.map((part) => `<Override PartName="/xl/tables/table${part.tablePartId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`).join("");
  const imageDefaults = [
    ["png", "image/png"],
    ["jpg", "image/jpeg"],
    ["jpeg", "image/jpeg"],
    ["gif", "image/gif"],
    ["svg", "image/svg+xml"],
  ].filter(([extension]) => imageParts.some((part) => part.extension === extension)).map(([extension, contentType]) => `<Default Extension="${extension}" ContentType="${contentType}"/>`).join("");
  const charts = chartParts.map((part) => `<Override PartName="/xl/charts/chart${part.chartPartId}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join("");
  const threadedComments = threadParts.map((part) => `<Override PartName="/xl/threadedComments/threadedComment${part.threadPartId}.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>`).join("");
  const persons = threadParts.length ? `<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>` : "";
  const shared = sharedStrings.strings?.length ? `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="json" ContentType="application/json"/>${imageDefaults}<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${shared}${sheets}${tables}${charts}${threadedComments}${persons}</Types>`;
}

function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map((rel) => `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${attrEscape(rel.target)}"${rel.targetMode ? ` TargetMode="${attrEscape(rel.targetMode)}"` : ""}/>`).join("")}</Relationships>`;
}

function parseRelsXml(xml) {
  return [...String(xml || "").matchAll(/<Relationship\b([^>]*)\/>/g)].map((match) => {
    const attrs = match[1] || "";
    return {
      id: decodeXml(/\bId="([^"]*)"/.exec(attrs)?.[1] || ""),
      type: decodeXml(/\bType="([^"]*)"/.exec(attrs)?.[1] || ""),
      target: decodeXml(/\bTarget="([^"]*)"/.exec(attrs)?.[1] || ""),
      targetMode: decodeXml(/\bTargetMode="([^"]*)"/.exec(attrs)?.[1] || ""),
    };
  }).filter((rel) => rel.id);
}

function workbookXml(workbook) {
  const sheets = workbook.worksheets.items.map((sheet, i) => `<sheet name="${attrEscape(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("");
  const definedNames = workbook.definedNames.items.length
    ? `<definedNames>${workbook.definedNames.items.map((item) => {
      const localSheetId = item.scope ? workbook.worksheets.items.findIndex((sheet) => sheet.name === item.scope) : -1;
      return `<definedName name="${attrEscape(item.name)}"${localSheetId >= 0 ? ` localSheetId="${localSheetId}"` : ""}>${xmlEscape(item.refersTo)}</definedName>`;
    }).join("")}</definedNames>`
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets}</sheets>${definedNames}</workbook>`;
}

function parseWorkbookDefinedNames(workbook, xml = "") {
  for (const match of String(xml || "").matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const attrs = match[1] || "";
    const name = decodeXml(/\bname="([^"]+)"/.exec(attrs)?.[1] || "");
    if (!name) continue;
    const localSheetIdRaw = /\blocalSheetId="(\d+)"/.exec(attrs)?.[1];
    const scope = localSheetIdRaw != null ? workbook.worksheets.items[Number(localSheetIdRaw)]?.name : undefined;
    workbook.definedNames.add({ name, refersTo: decodeXml(match[2] || ""), scope });
  }
}

function workbookRelsXml(sheetCount, hasThreadedComments = false, hasSharedStrings = false) {
  const rels = Array.from({ length: sheetCount }, (_, i) => ({ id: `rId${i + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", target: `worksheets/sheet${i + 1}.xml` }));
  rels.push({ id: `rId${sheetCount + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", target: "styles.xml" });
  let nextId = sheetCount + 2;
  if (hasSharedStrings) rels.push({ id: `rId${nextId++}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", target: "sharedStrings.xml" });
  if (hasThreadedComments) rels.push({ id: `rId${nextId++}`, type: "http://schemas.microsoft.com/office/2017/10/relationships/person", target: "persons/person.xml" });
  return relsXml(rels);
}

function worksheetRelsXml(tableParts, drawingRel, threadedPart) {
  const rels = tableParts.map((part) => ({
    id: part.relId,
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
    target: `../tables/table${part.tablePartId}.xml`,
  }));
  if (drawingRel) rels.push({ id: drawingRel.relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", target: drawingRel.target });
  if (threadedPart) rels.push({ id: threadedPart.relId, type: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment", target: `../threadedComments/threadedComment${threadedPart.threadPartId}.xml` });
  return relsXml(rels);
}

function stablePersonId(name) {
  const raw = String(name || "User");
  const hash = raw.split("").reduce((sum, ch) => ((sum * 33) + ch.charCodeAt(0)) >>> 0, 5381).toString(16).padStart(8, "0");
  return `{${hash.slice(0, 8)}-0000-4000-8000-000000000000}`;
}

function threadedCommentsXml(part) {
  const comments = part.threads.flatMap((thread) => thread.comments.map((comment, index) => {
    const id = `{${thread.id.replace(/[^A-Za-z0-9]/g, "")}-${index}}`;
    const parentId = index > 0 ? ` parentId="{${thread.id.replace(/[^A-Za-z0-9]/g, "")}-0}"` : "";
    return `<threadedComment ref="${attrEscape(thread.target.address)}" id="${attrEscape(id)}" personId="${attrEscape(stablePersonId(comment.author || thread.author))}" dT="${new Date(0).toISOString()}"${parentId} done="${thread.resolved ? 1 : 0}"><text>${xmlEscape(comment.text)}</text></threadedComment>`;
  })).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">${comments}</ThreadedComments>`;
}

function personsXml(threadParts) {
  const names = new Set(threadParts.flatMap((part) => part.threads.flatMap((thread) => thread.comments.map((comment) => comment.author || thread.author || "User"))));
  const persons = [...names].map((name) => `<person displayName="${attrEscape(name)}" id="${attrEscape(stablePersonId(name))}" userId="${attrEscape(name)}" providerId="None"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">${persons}</personList>`;
}

function parsePersonsXml(xml) {
  return new Map([...String(xml || "").matchAll(/<person\b([^>]*)\/>/g)].map((match) => {
    const attrs = match[1] || "";
    const id = decodeXml(/\bid="([^"]+)"/.exec(attrs)?.[1] || "");
    const displayName = decodeXml(/\bdisplayName="([^"]*)"/.exec(attrs)?.[1] || "User");
    return [id, displayName];
  }).filter(([id]) => id));
}

async function importNativeThreadedComments(workbook, zip, sheetRefs) {
  const persons = parsePersonsXml(await zip.file("xl/persons/person.xml")?.async("text"));
  for (const { index } of sheetRefs) {
    const sheet = workbook.worksheets.items[index - 1];
    if (!sheet) continue;
    const rels = parseRelsXml(await zip.file(`xl/worksheets/_rels/sheet${index}.xml.rels`)?.async("text"));
    for (const rel of rels.filter((item) => item.type.endsWith("/threadedComment"))) {
      const target = path.posix.normalize(`xl/worksheets/${rel.target}`).replace(/^\.\//, "");
      const xml = await zip.file(target)?.async("text");
      const byRef = new Map();
      for (const match of String(xml || "").matchAll(/<threadedComment\b([^>]*)>([\s\S]*?)<\/threadedComment>/g)) {
        const attrs = match[1] || "";
        const ref = decodeXml(/\bref="([^"]+)"/.exec(attrs)?.[1] || "A1");
        const personId = decodeXml(/\bpersonId="([^"]+)"/.exec(attrs)?.[1] || "");
        const author = persons.get(personId) || workbook.comments.self?.displayName || "User";
        const text = decodeXml(/<text[^>]*>([\s\S]*?)<\/text>/.exec(match[2])?.[1] || "");
        if (!byRef.has(ref)) {
          const thread = workbook.comments.addThread({ cell: sheet.getRange(ref) }, text);
          thread.author = author;
          thread.comments[0].author = author;
          thread.resolved = /\bdone="(?:1|true)"/.test(attrs);
          byRef.set(ref, thread);
        } else {
          byRef.get(ref).comments.push({ author, text });
        }
      }
    }
  }
}

function drawingRelsXml(imageParts, chartParts = []) {
  return relsXml([
    ...imageParts.map((part) => ({
      id: part.drawingRelId,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      target: `../media/image${part.imagePartId}.${part.extension}`,
    })),
    ...chartParts.map((part) => ({
      id: part.drawingRelId,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
      target: `../charts/chart${part.chartPartId}.xml`,
    })),
  ]);
}

function drawingXml(imageParts, chartParts = []) {
  const imageAnchors = imageParts.map((part, index) => {
    const from = part.image.anchor?.from || { row: 0, col: 0 };
    const extent = part.image.anchor?.extent || {};
    const widthPx = Number(extent.widthPx || part.image.anchor?.widthPx || 160);
    const heightPx = Number(extent.heightPx || part.image.anchor?.heightPx || 120);
    const cx = Math.round(widthPx * 9525);
    const cy = Math.round(heightPx * 9525);
    return `<xdr:oneCellAnchor><xdr:from><xdr:col>${Number(from.col || 0)}</xdr:col><xdr:colOff>${Math.round(Number(from.colOffsetPx || 0) * 9525)}</xdr:colOff><xdr:row>${Number(from.row || 0)}</xdr:row><xdr:rowOff>${Math.round(Number(from.rowOffsetPx || 0) * 9525)}</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${index + 2}" name="${attrEscape(part.image.name || `Image ${index + 1}`)}" descr="${attrEscape(part.image.alt || "")}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${part.drawingRelId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
  }).join("");
  const chartAnchors = chartParts.map((part, index) => {
    const p = part.chart.position;
    const col = Math.max(0, Math.floor((p.left - 40) / 96));
    const row = Math.max(0, Math.floor((p.top - 40) / 28));
    const colOff = Math.max(0, Math.round((p.left - 40 - col * 96) * 9525));
    const rowOff = Math.max(0, Math.round((p.top - 40 - row * 28) * 9525));
    const cx = Math.round(p.width * 9525);
    const cy = Math.round(p.height * 9525);
    const id = imageParts.length + index + 2;
    return `<xdr:oneCellAnchor><xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${attrEscape(part.chart.name || `Chart ${index + 1}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="${part.drawingRelId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${imageAnchors}${chartAnchors}</xdr:wsDr>`;
}

function sanitizeTableColumnName(value, index, seen) {
  const base = String(value ?? "").trim() || `Column${index + 1}`;
  let name = base;
  let suffix = 2;
  while (seen.has(name)) name = `${base}_${suffix++}`;
  seen.add(name);
  return name;
}

function tableXml(table, tablePartId) {
  const ref = table.range;
  const seen = new Set();
  const headers = table.showHeaders && table.values[0]
    ? table.values[0]
    : Array.from({ length: table.columnCount || 1 }, (_, index) => `Column${index + 1}`);
  const columns = Array.from({ length: table.columnCount || headers.length || 1 }, (_, index) => {
    const name = sanitizeTableColumnName(headers[index], index, seen);
    return `<tableColumn id="${index + 1}" name="${attrEscape(name)}"/>`;
  }).join("");
  const headerRowCount = table.showHeaders ? 1 : 0;
  const totalsRowShown = table.showTotals ? 1 : 0;
  const autoFilter = table.showFilterButton ? `<autoFilter ref="${attrEscape(ref)}"/>` : "";
  const styleName = table.style || "TableStyleMedium2";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${tablePartId}" name="${attrEscape(table.name)}" displayName="${attrEscape(table.name)}" ref="${attrEscape(ref)}" headerRowCount="${headerRowCount}" totalsRowShown="${totalsRowShown}">${autoFilter}<tableColumns count="${table.columnCount || headers.length || 1}">${columns}</tableColumns><tableStyleInfo name="${attrEscape(styleName)}" showFirstColumn="0" showLastColumn="0" showRowStripes="${table.showHeaders ? 1 : 0}" showColumnStripes="${table.showBandedColumns ? 1 : 0}"/></table>`;
}

function xlsxChartXml(chart) {
  const chartType = chart.type || chart.chartType || "bar";
  const chartElementName = chartType === "line" ? "lineChart" : "barChart";
  const grouping = chartType === "line" ? "<c:grouping val=\"standard\"/>" : "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>";
  const seriesItems = chart.series?.items || chart.series || [];
  const seriesXml = (seriesItems.length ? seriesItems : [{ name: chart.title || "Series", values: [] }]).map((series, index) => {
    const values = series.values || [];
    const categories = chart.categories || values.map((_, i) => String(i + 1));
    const catPts = categories.map((category, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${xmlEscape(category)}</c:v></c:pt>`).join("");
    const valPts = values.map((value, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${Number(value) || 0}</c:v></c:pt>`).join("");
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:v>${xmlEscape(series.name || `Series ${index + 1}`)}</c:v></c:tx><c:cat><c:strLit><c:ptCount val="${categories.length}"/>${catPts}</c:strLit></c:cat><c:val><c:numLit><c:ptCount val="${values.length}"/>${valPts}</c:numLit></c:val></c:ser>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(chart.title || chartType)}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:layout/><c:${chartElementName}>${grouping}${seriesXml}<c:axId val="1"/><c:axId val="2"/></c:${chartElementName}><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:layout/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

function dataValidationsXml(sheet) {
  if (sheet.dataValidations.items.length === 0) return "";
  const rules = sheet.dataValidations.items.map((item) => {
    const rule = item.rule || item;
    const type = rule.type || "custom";
    const operator = rule.operator ? ` operator="${attrEscape(rule.operator)}"` : "";
    const valuesFormula = Array.isArray(rule.values) ? `"${rule.values.map((value) => String(value).replaceAll('"', '""')).join(",")}"` : undefined;
    const formula1 = rule.formula1 ?? valuesFormula;
    const formula2 = rule.formula2;
    return `<dataValidation type="${attrEscape(type)}"${operator} allowBlank="1" sqref="${attrEscape(item.range || "A1")}">${formula1 != null ? `<formula1>${xmlEscape(formula1)}</formula1>` : ""}${formula2 != null ? `<formula2>${xmlEscape(formula2)}</formula2>` : ""}</dataValidation>`;
  }).join("");
  return `<dataValidations count="${sheet.dataValidations.items.length}">${rules}</dataValidations>`;
}

function conditionalFormattingXml(sheet) {
  return sheet.conditionalFormattings.items.map((item, index) => {
    const ruleType = item.ruleType || "expression";
    const type = ruleType === "cellIs" || ruleType === "CellValue" ? "cellIs" : ruleType === "containsText" ? "containsText" : "expression";
    const operator = item.operator ? ` operator="${attrEscape(item.operator)}"` : "";
    const formula = Array.isArray(item.formula) ? item.formula[0] : item.formula || item.expression || "TRUE";
    return `<conditionalFormatting sqref="${attrEscape(item.range || "A1")}"><cfRule type="${attrEscape(type)}" priority="${index + 1}"${operator}><formula>${xmlEscape(formula)}</formula></cfRule></conditionalFormatting>`;
  }).join("");
}

function worksheetXml(sheet, tableParts = [], drawingRelId, sharedStrings = { indexByText: new Map() }, styleTable = { styles: [{}], indexByKey: new Map([["", 0]]) }) {
  const rows = new Map();
  for (const [address, cell] of sheet.store.entries()) {
    const { row, col } = parseCellAddress(address);
    if (!rows.has(row)) rows.set(row, []);
    rows.get(row).push({ address, col, cell });
  }
  const rowXml = [...rows.entries()].sort((a, b) => a[0] - b[0]).map(([row, cells]) => `<row r="${row + 1}">${cells.sort((a, b) => a.col - b.col).map(({ address, cell }) => cellXml(address, cell, sharedStrings, styleTable)).join("")}</row>`).join("");
  const tablePartsXml = tableParts.length ? `<tableParts count="${tableParts.length}">${tableParts.map((part) => `<tablePart r:id="${part.relId}"/>`).join("")}</tableParts>` : "";
  const drawingXml = drawingRelId ? `<drawing r:id="${drawingRelId}"/>` : "";
  const sparklineXml = sparklineGroupExtXml(sheet);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>${rowXml}</sheetData>${conditionalFormattingXml(sheet)}${dataValidationsXml(sheet)}${drawingXml}${tablePartsXml}${sparklineXml}</worksheet>`;
}

function cellXml(address, cell, sharedStrings = { indexByText: new Map() }, styleTable = { styles: [{}], indexByKey: new Map([["", 0]]) }) {
  const f = cell.formula ? `<f>${xmlEscape(String(cell.formula).replace(/^=/, ""))}</f>` : "";
  const styleIndex = xlsxStyleIndex(cell, styleTable);
  const s = styleIndex ? ` s="${styleIndex}"` : "";
  if (typeof cell.value === "number") return `<c r="${address}"${s}>${f}<v>${cell.value}</v></c>`;
  if (typeof cell.value === "boolean") return `<c r="${address}" t="b"${s}>${f}<v>${cell.value ? 1 : 0}</v></c>`;
  if (cell.value == null && !f) return styleIndex ? `<c r="${address}"${s}/>` : "";
  if (cell.value == null) return `<c r="${address}"${s}>${f}</c>`;
  if (f) return `<c r="${address}" t="str"${s}>${f}<v>${xmlEscape(cell.value)}</v></c>`;
  const sharedIndex = sharedStrings.indexByText?.get(String(cell.value));
  if (sharedIndex !== undefined) return `<c r="${address}" t="s"${s}><v>${sharedIndex}</v></c>`;
  return `<c r="${address}" t="inlineStr"${s}><is><t>${xmlEscape(cell.value)}</t></is></c>`;
}

function parseWorksheetXml(sheet, xml, options = {}) {
  for (const match of xml.matchAll(/<c\s+([^>]*)>([\s\S]*?)<\/c>/g)) {
    const attrs = match[1];
    const body = match[2];
    const address = /r="([^"]+)"/.exec(attrs)?.[1];
    if (!address) continue;
    const cell = sheet.store.get(address);
    const styleIndex = Number(/\bs="([^"]+)"/.exec(attrs)?.[1] || 0);
    if (options.styles?.[styleIndex]) cell.style = { ...options.styles[styleIndex] };
    const formula = /<f[^>]*>([\s\S]*?)<\/f>/.exec(body)?.[1];
    if (formula) cell.formula = `=${decodeXml(formula)}`;
    const text = /<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/is>/.exec(body)?.[1];
    const value = /<v[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1];
    const type = /\bt="([^"]+)"/.exec(attrs)?.[1];
    if (type === "s" && value !== undefined) cell.value = options.sharedStrings?.[Number(value)] ?? "";
    else if (text !== undefined) cell.value = decodeXml(text);
    else if (value !== undefined) cell.value = Number.isFinite(Number(value)) && type !== "str" ? Number(value) : decodeXml(value);
  }
  parseSparklineGroupsXml(sheet, xml);
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

class PresentationTheme {
  constructor(presentation, config = {}) {
    this.presentation = presentation;
    this.id = config.id || "theme/default";
    this.name = config.name || "Open Office Clean Room";
    this.colors = {
      accent1: "#156082",
      accent2: "#0ea5e9",
      accent3: "#64748b",
      bg1: "#ffffff",
      tx1: "#0f172a",
      ...config.colors,
    };
    this.fonts = { major: "Aptos Display", minor: "Aptos", ...config.fonts };
  }

  setColors(colors = {}) { Object.assign(this.colors, colors); return this; }
  setFonts(fonts = {}) { Object.assign(this.fonts, fonts); return this; }
  inspectRecord() { return { kind: "theme", id: this.id, name: this.name, colors: this.colors, fonts: this.fonts }; }
  toJSON() { return { id: this.id, name: this.name, colors: this.colors, fonts: this.fonts }; }
}

class SlideLayoutTemplate {
  constructor(presentation, config = {}) {
    this.presentation = presentation;
    this.id = config.id || aid("lo");
    this.name = config.name || "Blank";
    this.type = config.type || "blank";
    this.masterId = config.masterId || "master/default";
    this.placeholders = (config.placeholders || []).map((placeholder, index) => ({
      id: placeholder.id || `${this.id}/ph/${index + 1}`,
      type: placeholder.type || "body",
      name: placeholder.name || `${placeholder.type || "body"} placeholder`,
      position: normalizeFrame(placeholder, { left: 80, top: 80 + index * 80, width: 640, height: 64 }),
      text: placeholder.text || "",
      required: Boolean(placeholder.required),
      style: placeholder.style || {},
    }));
  }

  apply(slide) {
    slide.layoutId = this.id;
    return this.placeholders.map((placeholder) => {
      const shape = slide.shapes.add({
        id: placeholder.id,
        name: placeholder.name,
        geometry: "rect",
        position: placeholder.position,
        fill: "transparent",
        line: { fill: "transparent", width: 0 },
        text: placeholder.text,
        placeholder: { layoutId: this.id, type: placeholder.type, name: placeholder.name, required: placeholder.required, idx: this.placeholders.indexOf(placeholder) + 1 },
      });
      shape.text.style = { ...placeholder.style };
      return shape;
    });
  }

  inspectRecord() { return { kind: "layoutTemplate", id: this.id, name: this.name, type: this.type, masterId: this.masterId, placeholders: this.placeholders.length, placeholderTypes: this.placeholders.map((placeholder) => placeholder.type) }; }
  toJSON() { return { id: this.id, name: this.name, type: this.type, masterId: this.masterId, placeholders: this.placeholders.map((placeholder) => ({ ...placeholder })) }; }
}

class SlideLayoutCollection {
  constructor(presentation) { this.presentation = presentation; this.items = []; }
  add(config = {}) { const layout = new SlideLayoutTemplate(this.presentation, config); this.items.push(layout); return layout; }
  getItem(idOrName) { return this.items.find((layout) => layout.id === idOrName || layout.name === idOrName || layout.type === idOrName); }
  inspectRecords() { return this.items.map((layout) => layout.inspectRecord()); }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

export class Presentation {
  constructor(options = {}) {
    this.id = aid("pr");
    this.slideSize = options.slideSize || { width: 1280, height: 720 };
    this.theme = new PresentationTheme(this, options.theme || {});
    this.layouts = new SlideLayoutCollection(this);
    for (const layout of options.layouts || []) this.layouts.add(layout);
    this.slides = new SlideCollection(this);
  }

  static create(options = {}) { return new Presentation(options); }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["deck", "slide", "textbox", "shape", "layout"]);
    const records = [];
    if (kinds.has("deck")) records.push({ kind: "deck", id: this.id, slides: this.slides.count });
    if (kinds.has("theme")) records.push(this.theme.inspectRecord());
    if (kinds.has("layout") || kinds.has("layoutTemplate")) records.push(...this.layouts.inspectRecords());
    for (const slide of this.slides) records.push(...slide.inspectRecords(kinds));
    return ndjson(filterInspectRecords(records, options), options.maxChars ?? Infinity);
  }

  validateLayout(options = {}) {
    const issues = this.slides.items.flatMap((slide) => slide.validateLayout(options).issues);
    return { ok: issues.length === 0, issues, ...ndjson(issues, options.maxChars ?? Infinity) };
  }

  verify(options = {}) {
    const issues = [];
    if (this.slides.items.length === 0) issues.push(verificationIssue("presentation", "noSlides", "Presentation has no slides."));
    issues.push(...this.validateLayout(options).issues.map((issue) => ({ ...issue, artifactKind: "presentation" })));
    for (const slide of this.slides) {
      if (slide.layoutId && !this.layouts.getItem(slide.layoutId)) issues.push(verificationIssue("presentation", "missingLayout", `Slide ${slide.index + 1} references missing layout ${slide.layoutId}.`, { slide: slide.index + 1, layoutId: slide.layoutId }));
      for (const shape of slide.shapes.items) {
        if (shape.placeholder?.required && !shape.text.value.trim()) issues.push(verificationIssue("presentation", "placeholderMissingContent", `Required ${shape.placeholder.type || "placeholder"} placeholder ${shape.name || shape.id} on slide ${slide.index + 1} is empty.`, { slide: slide.index + 1, id: shape.id, placeholder: shape.placeholder }));
      }
      for (const table of slide.tables.items) {
        if (!table.rows || !table.columns || table.values.length === 0 || table.values.every((row) => row.every((cell) => String(cell ?? "").trim() === ""))) issues.push(verificationIssue("presentation", "emptyTable", `Table ${table.name || table.id} on slide ${slide.index + 1} has no visible cell data.`, { slide: slide.index + 1, id: table.id }));
        if (table.values.length !== table.rows) issues.push(verificationIssue("presentation", "tableDataMismatch", `Table ${table.name || table.id} declares ${table.rows} rows but has ${table.values.length} value rows.`, { slide: slide.index + 1, id: table.id, rows: table.rows, valueRows: table.values.length }));
        if (table.values.some((row) => row.length !== table.columns)) issues.push(verificationIssue("presentation", "raggedTableRows", `Table ${table.name || table.id} has rows that do not match its declared column count.`, { slide: slide.index + 1, id: table.id, columns: table.columns, rowLengths: table.values.map((row) => row.length) }));
      }
      for (const chart of slide.charts.items) {
        if (!chart.series.length) issues.push(verificationIssue("presentation", "emptyChart", `Chart ${chart.name || chart.id} on slide ${slide.index + 1} has no data series.`, { slide: slide.index + 1, id: chart.id }));
        for (const series of chart.series) {
          const values = Array.isArray(series.values) ? series.values : [];
          if (chart.categories.length && values.length && chart.categories.length !== values.length) issues.push(verificationIssue("presentation", "chartDataMismatch", `Chart ${chart.name || chart.id} series ${series.name || "Series"} has ${values.length} values for ${chart.categories.length} categories.`, { slide: slide.index + 1, id: chart.id, series: series.name, values: values.length, categories: chart.categories.length }));
          if (values.some((value) => value !== "" && value != null && !Number.isFinite(Number(value)))) issues.push(verificationIssue("presentation", "chartDataNonNumeric", `Chart ${chart.name || chart.id} series ${series.name || "Series"} contains non-numeric values.`, { slide: slide.index + 1, id: chart.id, series: series.name }));
        }
      }
      for (const image of slide.images.items) {
        if (!image.dataUrl && !image.uri && !image.prompt) issues.push(verificationIssue("presentation", "emptyImage", `Image ${image.name || image.id} on slide ${slide.index + 1} has no dataUrl, uri, or prompt.`, { slide: slide.index + 1, id: image.id }));
        if (image.dataUrl && !imageDataFromDataUrl(image.dataUrl)) issues.push(verificationIssue("presentation", "invalidImageDataUrl", `Image ${image.name || image.id} on slide ${slide.index + 1} has an unsupported data URL.`, { slide: slide.index + 1, id: image.id }));
      }
      for (const comment of slide.comments) {
        if (comment.targetId && !slide.resolve(comment.targetId)) issues.push(verificationIssue("presentation", "danglingComment", `Slide ${slide.index + 1} comment ${comment.id} targets missing element ${comment.targetId}.`, { slide: slide.index + 1, id: comment.id, targetId: comment.targetId }));
      }
    }
    return verificationResult("presentation", issues, options);
  }

  resolve(id) {
    if (id === this.id) return this;
    if (id === this.theme.id) return this.theme;
    const layout = this.layouts.getItem(id);
    if (layout) return layout;
    for (const slide of this.slides) {
      if (slide.id === id) return slide;
      const found = slide.resolve(id);
      if (found) return found;
    }
    return undefined;
  }

  help(query = "*", options = {}) {
    return helpArtifact("presentation", query, options);
  }

  async export(options = {}) {
    const slide = options.slide || this.slides.getItem(0) || this.slides.add();
    if (options.format === "layout") return slide.export({ format: "layout" });
    return slide.export(options);
  }

  toProto() {
    return { id: this.id, slideSize: this.slideSize, theme: this.theme.toJSON(), layouts: this.layouts.items.map((layout) => layout.toJSON()), slides: this.slides.items.map((slide) => slide.toProto()) };
  }
}

class ShapeCollection {
  constructor(slide) { this.slide = slide; this.items = []; }
  add(config = {}) { const shape = new Shape(this.slide, config); this.items.push(shape); return shape; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class ElementCollection {
  constructor(slide, ElementClass) { this.slide = slide; this.ElementClass = ElementClass; this.items = []; }
  add(...args) { const element = new this.ElementClass(this.slide, ...args); this.items.push(element); return element; }
  getItemAt(index) { return this.items[index]; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

function normalizeFrame(config = {}, fallback = { left: 0, top: 0, width: 240, height: 160 }) {
  const source = config.position || config.frame || config;
  return {
    left: source.left ?? fallback.left,
    top: source.top ?? fallback.top,
    width: source.width ?? fallback.width,
    height: source.height ?? fallback.height,
  };
}

function resolveAutoLayoutFrame(slide, frame) {
  if (frame === "slide") return slide.frame;
  if (frame?.position) return frame.position;
  if (frame && typeof frame.left === "number" && typeof frame.top === "number" && typeof frame.width === "number" && typeof frame.height === "number") return frame;
  return slide.frame;
}

function elementFrame(element) {
  return element.position || element.frame || element.layoutJson?.().frame;
}

function elementLabel(element) {
  return element.name || element.id;
}

function overlapArea(a, b) {
  const left = Math.max(a.left, b.left);
  const top = Math.max(a.top, b.top);
  const right = Math.min(a.left + a.width, b.left + b.width);
  const bottom = Math.min(a.top + a.height, b.top + b.height);
  return Math.max(0, right - left) * Math.max(0, bottom - top);
}

function textOverflowIssue(slide, element, frame) {
  const text = element.text?.value || "";
  if (!text) return undefined;
  const fontSize = element.text.style.fontSize || 24;
  const lineSpacing = element.text.style.lineSpacing || 1.2;
  const availableWidth = Math.max(1, frame.width - 18);
  const charsPerLine = Math.max(1, Math.floor(availableWidth / (fontSize * 0.55)));
  const requiredLines = String(text).split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);
  const requiredHeight = requiredLines * fontSize * lineSpacing + 12;
  if (requiredHeight <= frame.height) return undefined;
  return {
    kind: "layoutIssue",
    type: "textOverflow",
    severity: "warning",
    slide: slide.index + 1,
    id: element.id,
    name: element.name || undefined,
    bbox: [frame.left, frame.top, frame.width, frame.height],
    requiredHeight: Math.round(requiredHeight),
    message: `Text may overflow ${elementLabel(element)}: estimated ${Math.round(requiredHeight)}px required for ${Math.round(frame.height)}px frame.`,
  };
}

function tableOverflowIssues(slide, tableElement) {
  const issues = [];
  const frame = tableElement.position;
  const cellW = frame.width / Math.max(1, tableElement.columns);
  const cellH = frame.height / Math.max(1, tableElement.rows);
  const fontSize = 13;
  for (let row = 0; row < tableElement.rows; row++) {
    for (let column = 0; column < tableElement.columns; column++) {
      const value = String(tableElement.values[row]?.[column] ?? "");
      const requiredWidth = value.length * fontSize * 0.55 + 12;
      if (requiredWidth > cellW || cellH < fontSize * 1.4) {
        issues.push({
          kind: "layoutIssue",
          type: "tableTextOverflow",
          severity: "warning",
          slide: slide.index + 1,
          id: tableElement.id,
          name: tableElement.name || undefined,
          row,
          column,
          bbox: [frame.left + column * cellW, frame.top + row * cellH, cellW, cellH],
          message: `Table cell ${elementLabel(tableElement)}[${row},${column}] may overflow its cell.`,
        });
      }
    }
  }
  return issues;
}

function pointFromElement(element, fallback = { x: 0, y: 0 }) {
  const frame = elementFrame(element);
  return frame ? { x: frame.left + frame.width / 2, y: frame.top + frame.height / 2 } : fallback;
}

function connectorPoint(slide, pointOrTarget, fallback = { x: 0, y: 0 }) {
  if (!pointOrTarget) return fallback;
  if (typeof pointOrTarget === "string") return pointFromElement(slide.resolve(pointOrTarget), fallback);
  if (pointOrTarget.id) return pointFromElement(slide.resolve(pointOrTarget.id) || pointOrTarget, fallback);
  if (pointOrTarget.element) return pointFromElement(pointOrTarget.element, fallback);
  if (pointOrTarget.targetId) return pointFromElement(slide.resolve(pointOrTarget.targetId), fallback);
  if (Number.isFinite(pointOrTarget.x) && Number.isFinite(pointOrTarget.y)) return { x: Number(pointOrTarget.x), y: Number(pointOrTarget.y) };
  return fallback;
}

class SlideCommentThread {
  constructor(slide, target, text, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("pc");
    this.targetId = typeof target === "string" ? target : target?.id || config.targetId;
    this.author = config.author || "User";
    this.resolved = Boolean(config.resolved);
    this.created = config.created || new Date(0).toISOString();
    this.comments = (config.comments || [{ author: this.author, text: String(text ?? ""), created: this.created }]).map((comment) => ({ author: comment.author || this.author, text: String(comment.text ?? ""), created: comment.created || this.created }));
  }

  addReply(text, config = {}) {
    this.comments.push({ author: config.author || this.author, text: String(text ?? ""), created: config.created || new Date(0).toISOString() });
    return this;
  }

  resolve() { this.resolved = true; return this; }
  reopen() { this.resolved = false; return this; }

  inspectRecord() {
    return { kind: "comment", id: this.id, slide: this.slide.index + 1, targetId: this.targetId, author: this.author, resolved: this.resolved, replies: Math.max(0, this.comments.length - 1), textPreview: this.comments.map((comment) => comment.text).join("\n").slice(0, 300) };
  }

  toJSON() { return { id: this.id, targetId: this.targetId, author: this.author, resolved: this.resolved, created: this.created, comments: this.comments.map((comment) => ({ ...comment })) }; }
}

class SlideCommentCollection {
  constructor(slide) { this.slide = slide; this.items = []; }
  addThread(target, text, config = {}) { const thread = new SlideCommentThread(this.slide, target, text, config); this.items.push(thread); return thread; }
  add(target, text, config = {}) { return this.addThread(target, text, config); }
  getItem(id) { return this.items.find((thread) => thread.id === id); }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class ConnectorElement {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("cx");
    this.name = config.name || "";
    this.connectorType = config.connectorType || config.type || "straight";
    this.startTargetId = typeof config.from === "string" ? config.from : config.from?.id || config.startTargetId;
    this.endTargetId = typeof config.to === "string" ? config.to : config.to?.id || config.endTargetId;
    this.start = config.start || connectorPoint(slide, config.from || config.startTargetId, { x: 0, y: 0 });
    this.end = config.end || connectorPoint(slide, config.to || config.endTargetId, { x: 160, y: 0 });
    this.line = config.line || { fill: "#334155", width: 2, endArrow: config.endArrow || "triangle" };
  }

  get position() {
    const left = Math.min(this.start.x, this.end.x);
    const top = Math.min(this.start.y, this.end.y);
    return { left, top, width: Math.abs(this.end.x - this.start.x), height: Math.abs(this.end.y - this.start.y) };
  }

  inspectRecord() {
    return { kind: "connector", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, connectorType: this.connectorType, start: this.start, end: this.end, startTargetId: this.startTargetId, endTargetId: this.endTargetId, line: this.line };
  }

  layoutJson() { return { kind: "connector", id: this.id, name: this.name, connectorType: this.connectorType, start: this.start, end: this.end, startTargetId: this.startTargetId, endTargetId: this.endTargetId, line: this.line, frame: this.position }; }

  toSvg() {
    const stroke = resolveColorToken(this.line?.fill || this.line?.color || "#334155", "#334155");
    const width = this.line?.width ?? 2;
    const markerId = `${this.id.replace(/[^A-Za-z0-9_-]/g, "")}-arrow`;
    const marker = this.line?.endArrow ? `<defs><marker id="${markerId}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" fill="${xmlEscape(stroke)}"/></marker></defs>` : "";
    return `${marker}<line x1="${this.start.x}" y1="${this.start.y}" x2="${this.end.x}" y2="${this.end.y}" stroke="${xmlEscape(stroke)}" stroke-width="${width}" marker-end="${this.line?.endArrow ? `url(#${markerId})` : ""}"/>`;
  }

  toPptxShape(index) { return pptxConnectorXml(index, this); }
}

export class Slide {
  constructor(presentation, options = {}) {
    this.presentation = presentation;
    this.id = aid("sl");
    this.name = options.name || "";
    this.shapes = new ShapeCollection(this);
    this.images = new ElementCollection(this, ImageElement);
    this.tables = new ElementCollection(this, TableElement);
    this.charts = new ElementCollection(this, ChartElement);
    this.connectors = new ElementCollection(this, ConnectorElement);
    this.comments = new SlideCommentCollection(this);
    this.layoutId = options.layoutId || options.layout?.id || (typeof options.layout === "string" ? options.layout : undefined);
    this.speakerNotes = { text: String(options.notes || options.speakerNotes?.text || "") };
    this.background = { fill: "white" };
  }

  get index() { return this.presentation.slides.items.indexOf(this); }
  get frame() { return { left: 0, top: 0, ...this.presentation.slideSize }; }

  addNotes(text) { this.speakerNotes.text = String(text ?? ""); return this.speakerNotes; }
  addComment(target, text, config = {}) { return this.comments.addThread(target, text, config); }
  addConnector(config = {}) { return this.connectors.add(config); }
  applyLayout(layoutOrName) { const layout = typeof layoutOrName === "string" ? this.presentation.layouts.getItem(layoutOrName) : layoutOrName; if (!layout) throw new Error(`Unknown slide layout: ${layoutOrName}`); return layout.apply(this); }

  inspectRecords(kinds) {
    const records = [];
    if (kinds.has("layout")) { const layout = this.presentation.layouts.getItem(this.layoutId); records.push({ kind: "layout", layoutId: this.layoutId || `${this.id}/layout`, name: layout?.name || "Blank", type: layout?.type || "blank", placeholders: layout?.placeholders.length || 0 }); }
    if (kinds.has("slide")) records.push({ kind: "slide", id: this.id, slide: this.index + 1, title: this.title(), textShapes: this.shapes.items.filter((s) => s.text.value).length, tables: this.tables.items.length, charts: this.charts.items.length, images: this.images.items.length, connectors: this.connectors.items.length, comments: this.comments.items.length, hasNotes: Boolean(this.speakerNotes.text) });
    for (const shape of this.shapes) {
      if (kinds.has("textbox") && shape.text.value) records.push(shape.inspectRecord("textbox"));
      else if (kinds.has("shape")) records.push(shape.inspectRecord("shape"));
    }
    if (kinds.has("table")) records.push(...this.tables.items.map((table) => table.inspectRecord()));
    if (kinds.has("chart")) records.push(...this.charts.items.map((chart) => chart.inspectRecord()));
    if (kinds.has("image")) records.push(...this.images.items.map((image) => image.inspectRecord()));
    if (kinds.has("connector")) records.push(...this.connectors.items.map((connector) => connector.inspectRecord()));
    if (kinds.has("comment") || kinds.has("thread")) records.push(...this.comments.items.map((comment) => comment.inspectRecord()));
    if (kinds.has("notes")) records.push({ kind: "notes", id: `${this.id}/notes`, slide: this.index + 1, text: this.speakerNotes.text, textPreview: this.speakerNotes.text.slice(0, 300), textChars: this.speakerNotes.text.length });
    return records;
  }

  title() { return this.shapes.items.find((shape) => shape.text.value)?.text.value || this.charts.items[0]?.title || ""; }
  resolve(id) { return [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.connectors.items, ...this.comments.items].find((element) => element.id === id); }

  validateLayout(options = {}) {
    const issues = [];
    const slideFrame = this.frame;
    const elements = [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items];
    const connectors = this.connectors.items;
    const minOverlapArea = options.minOverlapArea ?? 64;
    const padding = options.boundsPadding ?? 0;
    for (const element of elements) {
      const frame = elementFrame(element);
      if (!frame) continue;
      const offCanvas = frame.left < slideFrame.left - padding || frame.top < slideFrame.top - padding || frame.left + frame.width > slideFrame.left + slideFrame.width + padding || frame.top + frame.height > slideFrame.top + slideFrame.height + padding;
      if (offCanvas) {
        issues.push({
          kind: "layoutIssue",
          type: "offCanvas",
          severity: "error",
          slide: this.index + 1,
          id: element.id,
          name: element.name || undefined,
          bbox: [frame.left, frame.top, frame.width, frame.height],
          message: `${elementLabel(element)} extends outside the slide frame.`,
        });
      }
      const textIssue = textOverflowIssue(this, element, frame);
      if (textIssue) issues.push(textIssue);
      if (element instanceof TableElement) issues.push(...tableOverflowIssues(this, element));
    }
    for (const connector of connectors) {
      const points = [connector.start, connector.end];
      if (points.some((point) => point.x < slideFrame.left - padding || point.y < slideFrame.top - padding || point.x > slideFrame.left + slideFrame.width + padding || point.y > slideFrame.top + slideFrame.height + padding)) {
        issues.push({ kind: "layoutIssue", type: "connectorOffCanvas", severity: "error", slide: this.index + 1, id: connector.id, name: connector.name || undefined, start: connector.start, end: connector.end, message: `${elementLabel(connector)} connector endpoint extends outside the slide frame.` });
      }
    }
    for (let leftIndex = 0; leftIndex < elements.length; leftIndex++) {
      for (let rightIndex = leftIndex + 1; rightIndex < elements.length; rightIndex++) {
        const left = elements[leftIndex];
        const right = elements[rightIndex];
        const leftFrame = elementFrame(left);
        const rightFrame = elementFrame(right);
        if (!leftFrame || !rightFrame) continue;
        const area = overlapArea(leftFrame, rightFrame);
        if (area >= minOverlapArea) {
          issues.push({
            kind: "layoutIssue",
            type: "overlap",
            severity: "error",
            slide: this.index + 1,
            ids: [left.id, right.id],
            names: [elementLabel(left), elementLabel(right)],
            overlapArea: Math.round(area),
            message: `${elementLabel(left)} overlaps ${elementLabel(right)} by about ${Math.round(area)}px².`,
          });
        }
      }
    }
    return { ok: issues.length === 0, issues, ...ndjson(issues, options.maxChars ?? Infinity) };
  }

  async export(options = {}) {
    if (options.format === "layout") return new FileBlob(JSON.stringify(this.layoutJson(), null, 2), { type: LAYOUT_MIME });
    return new FileBlob(this.toSvg(), { type: "image/svg+xml" });
  }

  layoutJson() {
    return {
      schema: "open-office-artifact.layout/v1",
      unit: "px",
      slide: { id: this.id, slide: this.index + 1, frame: this.frame, notes: this.speakerNotes.text || undefined },
      elements: [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.connectors.items].map((element) => element.layoutJson()),
    };
  }

  toSvg() {
    const { width, height } = this.presentation.slideSize;
    const elements = [...this.connectors.items, ...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items].map((element) => element.toSvg()).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="${xmlEscape(this.background.fill || "white")}"/>${elements}</svg>`;
  }

  toProto() { return { id: this.id, notes: this.speakerNotes.text || undefined, comments: this.comments.items.map((comment) => comment.toJSON()), elements: [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.connectors.items].map((element) => element.layoutJson()) }; }

  compose(composeNode, options = {}) {
    const frame = options.frame || { left: 72, top: 64, width: this.presentation.slideSize.width - 144, height: this.presentation.slideSize.height - 128 };
    return materializeComposeNode(this, composeNode, frame);
  }

  autoLayout(shapes, options = {}) {
    const items = Array.from(shapes || []).filter(Boolean);
    if (items.length === 0) return items;
    const frame = resolveAutoLayoutFrame(this, options.frame || "slide");
    const inner = innerFrame(frame, {
      left: options.horizontalPadding ?? 0,
      right: options.horizontalPadding ?? 0,
      top: options.verticalPadding ?? 0,
      bottom: options.verticalPadding ?? 0,
    });
    const direction = options.direction || "horizontal";
    const horizontal = direction === "horizontal";
    const mainSize = horizontal ? "width" : "height";
    const crossSize = horizontal ? "height" : "width";
    const requestedGap = horizontal ? options.horizontalGap : options.verticalGap;
    const totalMain = items.reduce((sum, shape) => sum + (shape.position?.[mainSize] ?? 0), 0);
    const gap = requestedGap === "auto"
      ? items.length > 1 ? Math.max(0, (inner[mainSize] - totalMain) / (items.length - 1)) : 0
      : Number(requestedGap ?? 0);
    const usedMain = totalMain + gap * Math.max(0, items.length - 1);
    const align = options.align || "center";
    const mainStart = align.includes("Right") || align === "right" || align.includes("Bottom")
      ? inner[horizontal ? "left" : "top"] + inner[mainSize] - usedMain
      : align === "center" || align === "left" || align === "right"
        ? inner[horizontal ? "left" : "top"] + Math.max(0, (inner[mainSize] - usedMain) / 2)
        : inner[horizontal ? "left" : "top"];
    let cursor = mainStart;
    for (const shape of items) {
      const crossStart = align.includes("Bottom")
        ? inner[horizontal ? "top" : "left"] + inner[crossSize] - shape.position[crossSize]
        : align.includes("Center") || align === "center" || align === "left" || align === "right"
          ? inner[horizontal ? "top" : "left"] + Math.max(0, (inner[crossSize] - shape.position[crossSize]) / 2)
          : inner[horizontal ? "top" : "left"];
      shape.position = horizontal
        ? { ...shape.position, left: cursor, top: crossStart }
        : { ...shape.position, left: crossStart, top: cursor };
      cursor += shape.position[mainSize] + gap;
    }
    return items;
  }
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
    this.id = config.id || aid("sh");
    this.geometry = config.geometry || "rect";
    this.name = config.name || "";
    this.position = config.position || { left: 0, top: 0, width: 160, height: 80 };
    this.fill = config.fill || "transparent";
    this.line = config.line || { fill: "#334155", width: 1 };
    this.borderRadius = config.borderRadius;
    this.placeholder = config.placeholder;
    this._text = new TextFrame(config.text || "");
  }

  get text() { return this._text; }
  set text(value) { this._text.set(value); }

  inspectRecord(kind = "shape") {
    const p = this.position;
    return { kind, id: this.id, slide: this.slide.index + 1, name: this.name || undefined, text: this.text.value || undefined, textPreview: this.text.value || undefined, textChars: this.text.value.length || undefined, textLines: this.text.value ? this.text.value.split(/\r?\n/).length : undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", placeholder: this.placeholder || undefined };
  }

  layoutJson() { return { kind: this.text.value ? "textbox" : "shape", id: this.id, name: this.name, geometry: this.geometry, frame: this.position, text: this.text.value, placeholder: this.placeholder, style: { fill: this.fill, line: this.line, borderRadius: this.borderRadius, text: this.text.style } }; }

  toSvg() {
    const p = this.position;
    const fill = typeof this.fill === "string" ? resolveColorToken(this.fill, this.fill) : this.fill?.color || "transparent";
    const stroke = resolveColorToken(this.line?.fill || this.line?.color || "#334155", "#334155");
    const sw = this.line?.width ?? 1;
    const rect = this.geometry === "ellipse"
      ? `<ellipse cx="${p.left + p.width / 2}" cy="${p.top + p.height / 2}" rx="${p.width / 2}" ry="${p.height / 2}" fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}"/>`
      : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" rx="${this.borderRadius ? 12 : 0}" fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}"/>`;
    const text = this.text.value ? `<text x="${p.left + 12}" y="${p.top + 36}" font-family="Arial" font-size="${this.text.style.fontSize || 24}" font-weight="${this.text.style.bold ? "700" : "400"}" fill="${xmlEscape(this.text.style.color || "#0f172a")}">${xmlEscape(this.text.value)}</text>` : "";
    return rect + text;
  }

  toPptxShape(index) {
    return pptxTextShapeXml(index, this.name || this.id, this.geometry, this.position, this.text.value, this.placeholder);
  }
}

class TableCellFacade {
  constructor(table, row, column) { this.table = table; this.row = row; this.column = column; this.text = new TextFrame(); }
  get value() { return this.table.values[this.row]?.[this.column] ?? ""; }
  set value(value) { this.table.ensureCell(this.row, this.column); this.table.values[this.row][this.column] = value; }
}

export class TableElement {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("tb");
    this.name = config.name || "";
    this.rows = Number(config.rows || config.values?.length || 1);
    this.columns = Number(config.columns || config.values?.[0]?.length || 1);
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 320, height: 160 });
    this.values = Array.from({ length: this.rows }, (_, r) => Array.from({ length: this.columns }, (_, c) => config.values?.[r]?.[c] ?? ""));
    this.style = config.style;
    this.styleOptions = config.styleOptions || {};
    this.cells = { set: (row, column, value) => { this.getCell(row, column).value = value; }, block: (range) => ({ table: this, range }) };
    this.borders = { assign: (configValue) => { this.border = configValue; } };
  }

  ensureCell(row, column) {
    while (this.values.length <= row) this.values.push([]);
    while (this.values[row].length <= column) this.values[row].push("");
  }

  getCell(row, column) { return new TableCellFacade(this, row, column); }
  merge(range) { this.mergeRange = range; }

  inspectRecord() {
    const p = this.position;
    return { kind: "table", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, rows: this.rows, cols: this.columns, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", values: this.values };
  }

  layoutJson() { return { kind: "table", id: this.id, name: this.name, frame: this.position, rows: this.rows, columns: this.columns, values: this.values, style: this.style, styleOptions: this.styleOptions }; }

  toSvg() {
    const p = this.position;
    const cellW = p.width / Math.max(1, this.columns);
    const cellH = p.height / Math.max(1, this.rows);
    const parts = [`<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>`];
    for (let r = 0; r < this.rows; r++) {
      for (let c = 0; c < this.columns; c++) {
        const x = p.left + c * cellW;
        const y = p.top + r * cellH;
        const fill = this.styleOptions.headerRow && r === 0 ? "#0f172a" : r % 2 ? "#f8fafc" : "#ffffff";
        const color = this.styleOptions.headerRow && r === 0 ? "#ffffff" : "#0f172a";
        parts.push(`<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" fill="${fill}" stroke="#cbd5e1"/>`);
        parts.push(`<text x="${x + 6}" y="${y + Math.min(22, cellH - 6)}" font-family="Arial" font-size="13" fill="${color}">${xmlEscape(this.values[r]?.[c] ?? "")}</text>`);
      }
    }
    return parts.join("");
  }

  toPptxShape(index) { return pptxTableXml(index, this); }
}

export class ChartElement {
  constructor(slide, chartType = "bar", config = {}) {
    this.slide = slide;
    this.id = config.id || aid("ch");
    this.name = config.name || "";
    this.chartType = chartType || config.chartType || "bar";
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 360, height: 220 });
    this.title = config.title || "";
    this.categories = config.categories || [];
    this.series = config.series || [];
    this.hasLegend = config.hasLegend ?? this.series.length > 1;
  }

  inspectRecord() {
    const p = this.position;
    return { kind: "chart", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, chartType: this.chartType, title: this.title, series: this.series.length, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px" };
  }

  layoutJson() { return { kind: "chart", id: this.id, name: this.name, chartType: this.chartType, title: this.title, frame: this.position, categories: this.categories, series: this.series }; }

  toSvg() {
    const p = this.position;
    const values = this.series[0]?.values || [];
    const max = Math.max(1, ...values.map((value) => Number(value) || 0));
    const plot = { left: p.left + 36, top: p.top + 42, width: Math.max(0, p.width - 56), height: Math.max(0, p.height - 72) };
    const barW = values.length ? plot.width / values.length * 0.65 : 0;
    const gap = values.length ? plot.width / values.length * 0.35 : 0;
    const bars = values.map((value, index) => {
      const h = plot.height * (Number(value) || 0) / max;
      const x = plot.left + index * (barW + gap) + gap / 2;
      const y = plot.top + plot.height - h;
      return `<rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="#0ea5e9"/>`;
    }).join("");
    const labels = this.categories.map((category, index) => `<text x="${plot.left + index * (barW + gap)}" y="${p.top + p.height - 12}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(category)}</text>`).join("");
    return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/><text x="${p.left + 12}" y="${p.top + 24}" font-family="Arial" font-size="16" font-weight="700" fill="#0f172a">${xmlEscape(this.title || this.chartType)}</text>${bars}${labels}`;
  }

  toPptxShape(index, relId) {
    if (relId) return pptxChartFrameXml(index, this.name || this.id, this.position, relId);
    return pptxTextShapeXml(index, this.name || this.id, "rect", this.position, `${this.title || this.chartType}\n${this.series.map((series) => `${series.name || "Series"}: ${(series.values || []).join(", ")}`).join("\n")}`);
  }
}

export class ImageElement {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("im");
    this.name = config.name || "";
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 320, height: 180 });
    this.alt = config.alt || "";
    this.prompt = config.prompt;
    this.uri = config.uri;
    this.dataUrl = config.dataUrl;
    this.contentType = config.contentType;
    this.fit = config.fit || "contain";
    this.geometry = config.geometry || "rect";
    this.borderRadius = config.borderRadius;
  }

  get frame() { return this.position; }
  set frame(value) { this.position = normalizeFrame(value, this.position); }
  replace(config = {}) { Object.assign(this, config); }

  inspectRecord() {
    const p = this.position;
    return { kind: "image", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, alt: this.alt || undefined, prompt: this.prompt || undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", fit: this.fit };
  }

  layoutJson() { return { kind: "image", id: this.id, name: this.name, frame: this.position, alt: this.alt, prompt: this.prompt, uri: this.uri, dataUrl: this.dataUrl, fit: this.fit, geometry: this.geometry, borderRadius: this.borderRadius }; }

  toSvg() {
    const p = this.position;
    const label = this.alt || this.prompt || this.uri || "image";
    if (this.dataUrl) return `<image href="${attrEscape(this.dataUrl)}" x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" preserveAspectRatio="xMidYMid meet"/><text x="${p.left + 12}" y="${p.top + 28}" font-family="Arial" font-size="14" fill="#075985">${xmlEscape(label)}</text>`;
    const rect = this.geometry === "ellipse"
      ? `<ellipse cx="${p.left + p.width / 2}" cy="${p.top + p.height / 2}" rx="${p.width / 2}" ry="${p.height / 2}" fill="#e0f2fe" stroke="#0284c7"/>`
      : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" rx="${this.borderRadius ? 12 : 0}" fill="#e0f2fe" stroke="#0284c7"/>`;
    return `${rect}<text x="${p.left + 12}" y="${p.top + 28}" font-family="Arial" font-size="14" fill="#075985">${xmlEscape(label)}</text>`;
  }

  toPptxShape(index, relId) {
    if (!relId) return pptxTextShapeXml(index, this.name || this.id, this.geometry === "ellipse" ? "ellipse" : "rect", this.position, this.alt || this.prompt || "Image");
    return pptxPictureXml(index, this.name || this.id, this.alt || this.prompt || "", this.position, relId);
  }
}

export class PresentationFile {
  static async exportPptx(presentation) {
    const zip = new JSZip();
    const imageParts = collectPresentationImageParts(presentation);
    const chartParts = collectPresentationChartParts(presentation, imageParts);
    const layoutParts = collectPresentationLayoutParts(presentation);
    zip.file("[Content_Types].xml", pptxContentTypes(presentation.slides.count, imageParts, chartParts, presentation, layoutParts));
    zip.file("_rels/.rels", relsXml([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "ppt/presentation.xml" }]));
    zip.file("ppt/presentation.xml", presentationXml(presentation));
    zip.file("ppt/_rels/presentation.xml.rels", pptxPresentationRelsXml(presentation, layoutParts));
    zip.file("ppt/theme/theme1.xml", pptxThemeXml(presentation.theme));
    if (layoutParts.length) {
      zip.file("ppt/slideMasters/slideMaster1.xml", pptxSlideMasterXml(layoutParts));
      zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", pptxSlideMasterRelsXml(layoutParts));
      for (const part of layoutParts) zip.file(`ppt/slideLayouts/slideLayout${part.layoutPartId}.xml`, pptxSlideLayoutXml(part.layout));
    }
    presentation.slides.items.forEach((slide, i) => {
      const slideImageParts = imageParts.filter((part) => part.slideIndex === i);
      const slideChartParts = chartParts.filter((part) => part.slideIndex === i);
      const nextRelIndex = slideImageParts.length + slideChartParts.length + 1;
      const slideLayoutPart = layoutParts.find((part) => part.layout.id === slide.layoutId || part.layout.name === slide.layoutId) || layoutParts[0];
      const layoutRelId = slideLayoutPart ? `rId${nextRelIndex}` : undefined;
      const notesRelId = slide.speakerNotes.text ? `rId${nextRelIndex + (layoutRelId ? 1 : 0)}` : undefined;
      const commentsRelId = slide.comments.items.length ? `rId${nextRelIndex + (layoutRelId ? 1 : 0) + (notesRelId ? 1 : 0)}` : undefined;
      zip.file(`ppt/slides/slide${i + 1}.xml`, slideXml(slide, slideImageParts, slideChartParts));
      if (slideImageParts.length || slideChartParts.length || layoutRelId || notesRelId || commentsRelId) zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, pptxSlideRelsXml(slideImageParts, slideChartParts, { slideIndex: i, layoutRelId, layoutPartId: slideLayoutPart?.layoutPartId, notesRelId, commentsRelId }));
      if (notesRelId) zip.file(`ppt/notesSlides/notesSlide${i + 1}.xml`, pptxNotesSlideXml(slide));
      if (commentsRelId) zip.file(`ppt/comments/comment${i + 1}.xml`, pptxCommentsXml(slide));
    });
    imageParts.forEach((part) => zip.file(`ppt/media/image${part.imagePartId}.${part.extension}`, part.bytes));
    chartParts.forEach((part) => zip.file(`ppt/charts/chart${part.chartPartId}.xml`, pptxChartXml(part.chart)));
    const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return new FileBlob(bytes, { type: PPTX_MIME });
  }

  static async importPptx(blobOrBuffer) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const presentation = Presentation.create();
    const presentationRels = parseRelsXml(await zip.file("ppt/_rels/presentation.xml.rels")?.async("text"));
    const themeRel = presentationRels.find((rel) => rel.type.endsWith("/theme"));
    const themeTarget = themeRel?.target ? (themeRel.target.replace(/^\//, "").startsWith("ppt/") ? themeRel.target.replace(/^\//, "") : path.posix.normalize(`ppt/${themeRel.target}`).replace(/^\.\//, "")) : undefined;
    const themeXml = themeTarget ? await zip.file(themeTarget)?.async("text") : await zip.file("ppt/theme/theme1.xml")?.async("text");
    if (themeXml) presentation.theme = parsePptxTheme(presentation, themeXml);
    const layoutByTarget = new Map();
    const slideFiles = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort();
    for (const file of slideFiles) {
      const slide = presentation.slides.add();
      const relsFile = file.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
      const rels = parseRelsXml(await zip.file(relsFile)?.async("text"));
      const layoutRel = rels.find((rel) => rel.type.endsWith("/slideLayout"));
      const layoutTarget = layoutRel ? pptxRelationshipTarget(rels, layoutRel.id) : undefined;
      if (layoutTarget) {
        let layout = layoutByTarget.get(layoutTarget);
        if (!layout) {
          layout = parsePptxSlideLayout(presentation, await zip.file(layoutTarget)?.async("text"), `imported-layout-${layoutByTarget.size + 1}`);
          layoutByTarget.set(layoutTarget, layout);
        }
        slide.layoutId = layout?.id;
      }
      await parseSlideXml(slide, await zip.file(file).async("text"), { rels, zip });
      const notesRel = rels.find((rel) => rel.type.endsWith("/notesSlide"));
      const commentsRel = rels.find((rel) => rel.type.endsWith("/comments"));
      const notesTarget = notesRel ? pptxRelationshipTarget(rels, notesRel.id) : undefined;
      const commentsTarget = commentsRel ? pptxRelationshipTarget(rels, commentsRel.id) : undefined;
      if (notesTarget) parsePptxNotes(slide, await zip.file(notesTarget)?.async("text"));
      if (commentsTarget) parsePptxComments(slide, await zip.file(commentsTarget)?.async("text"));
    }
    return presentation;
  }
}

function collectPresentationImageParts(presentation) {
  const parts = [];
  let imagePartId = 1;
  presentation.slides.items.forEach((slide, slideIndex) => {
    let relIndex = 1;
    slide.images.items.forEach((image, imageIndex) => {
      const data = imageDataFromDataUrl(image.dataUrl);
      if (!data) return;
      parts.push({
        slide,
        slideIndex,
        image,
        imageIndex,
        imagePartId: imagePartId++,
        slideRelId: `rId${relIndex++}`,
        ...data,
      });
    });
  });
  return parts;
}

function collectPresentationChartParts(presentation, imageParts = []) {
  const parts = [];
  let chartPartId = 1;
  presentation.slides.items.forEach((slide, slideIndex) => {
    let relIndex = imageParts.filter((part) => part.slideIndex === slideIndex).length + 1;
    slide.charts.items.forEach((chart, chartIndex) => {
      parts.push({ slide, slideIndex, chart, chartIndex, chartPartId: chartPartId++, slideRelId: `rId${relIndex++}` });
    });
  });
  return parts;
}

function collectPresentationLayoutParts(presentation) {
  return presentation.layouts.items.map((layout, index) => ({ layout, layoutPartId: index + 1, masterRelId: `rId${index + 1}` }));
}

function pptxContentTypes(slideCount, imageParts = [], chartParts = [], presentation, layoutParts = []) {
  const slides = Array.from({ length: slideCount }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("");
  const charts = chartParts.map((part) => `<Override PartName="/ppt/charts/chart${part.chartPartId}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join("");
  const notes = presentation?.slides.items.map((slide, i) => slide.speakerNotes.text ? `<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>` : "").join("") || "";
  const comments = presentation?.slides.items.map((slide, i) => slide.comments.items.length ? `<Override PartName="/ppt/comments/comment${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>` : "").join("") || "";
  const theme = `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`;
  const master = layoutParts.length ? `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` : "";
  const layouts = layoutParts.map((part) => `<Override PartName="/ppt/slideLayouts/slideLayout${part.layoutPartId}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`).join("");
  const imageDefaults = [
    ["png", "image/png"],
    ["jpg", "image/jpeg"],
    ["jpeg", "image/jpeg"],
    ["gif", "image/gif"],
    ["svg", "image/svg+xml"],
  ].filter(([extension]) => imageParts.some((part) => part.extension === extension)).map(([extension, contentType]) => `<Default Extension="${extension}" ContentType="${contentType}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${imageDefaults}<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${slides}${charts}${theme}${master}${layouts}${notes}${comments}</Types>`;
}

function pptxSlideRelsXml(imageParts, chartParts = [], extras = {}) {
  return relsXml([
    ...imageParts.map((part) => ({ id: part.slideRelId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: `../media/image${part.imagePartId}.${part.extension}` })),
    ...chartParts.map((part) => ({ id: part.slideRelId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", target: `../charts/chart${part.chartPartId}.xml` })),
    ...(extras.layoutRelId ? [{ id: extras.layoutRelId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", target: `../slideLayouts/slideLayout${extras.layoutPartId}.xml` }] : []),
    ...(extras.notesRelId ? [{ id: extras.notesRelId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide", target: `../notesSlides/notesSlide${extras.slideIndex + 1}.xml` }] : []),
    ...(extras.commentsRelId ? [{ id: extras.commentsRelId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", target: `../comments/comment${extras.slideIndex + 1}.xml` }] : []),
  ]);
}

function pptxPresentationRelsXml(presentation, layoutParts = []) {
  const slideRels = presentation.slides.items.map((_, i) => ({ id: `rId${i + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", target: `slides/slide${i + 1}.xml` }));
  const themeRel = { id: `rId${slideRels.length + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", target: "theme/theme1.xml" };
  const masterRel = layoutParts.length ? [{ id: `rId${slideRels.length + 2}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", target: "slideMasters/slideMaster1.xml" }] : [];
  return relsXml([...slideRels, themeRel, ...masterRel]);
}

function pptxColorValue(value, fallback) {
  return String(value || fallback || "#000000").replace(/^#/, "").slice(0, 6).padEnd(6, "0");
}

function pptxThemeXml(theme) {
  const colors = theme.colors || {};
  const fonts = theme.fonts || {};
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${attrEscape(theme.name || "Open Office Clean Room")}"><a:themeElements><a:clrScheme name="${attrEscape(theme.name || "Clean Room")}"><a:dk1><a:srgbClr val="${pptxColorValue(colors.tx1, "#0f172a")}"/></a:dk1><a:lt1><a:srgbClr val="${pptxColorValue(colors.bg1, "#ffffff")}"/></a:lt1><a:accent1><a:srgbClr val="${pptxColorValue(colors.accent1, "#156082")}"/></a:accent1><a:accent2><a:srgbClr val="${pptxColorValue(colors.accent2, "#0ea5e9")}"/></a:accent2><a:accent3><a:srgbClr val="${pptxColorValue(colors.accent3, "#64748b")}"/></a:accent3><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Clean Room"><a:majorFont><a:latin typeface="${attrEscape(fonts.major || "Aptos Display")}"/></a:majorFont><a:minorFont><a:latin typeface="${attrEscape(fonts.minor || "Aptos")}"/></a:minorFont></a:fontScheme><a:fmtScheme name="Clean Room"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>`;
}

function pptxSlideMasterXml(layoutParts = []) {
  const ids = layoutParts.map((part, index) => `<p:sldLayoutId id="${2147483649 + index}" r:id="${part.masterRelId}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst><p:txStyles/></p:sldMaster>`;
}

function pptxSlideMasterRelsXml(layoutParts = []) {
  return relsXml(layoutParts.map((part) => ({ id: part.masterRelId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", target: `../slideLayouts/slideLayout${part.layoutPartId}.xml` })));
}

function pptxSlideLayoutXml(layout) {
  const placeholders = layout.placeholders.map((placeholder, index) => pptxTextShapeXml(index, placeholder.name, "rect", placeholder.position, placeholder.text || "", { type: placeholder.type, idx: index + 1, required: placeholder.required })).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="${attrEscape(layout.type)}" preserve="1"><p:cSld name="${attrEscape(layout.name)}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${placeholders}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function presentationXml(presentation) {
  const ids = presentation.slides.items.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="12192000" cy="6858000"/><p:sldIdLst>${ids}</p:sldIdLst></p:presentation>`;
}

function pptxTextShapeXml(index, name, geometry, position, text = "", placeholder) {
  const p = position;
  const x = Math.round(p.left * 9525), y = Math.round(p.top * 9525), cx = Math.round(p.width * 9525), cy = Math.round(p.height * 9525);
  const paragraphs = String(text || "").split(/\r?\n/).map((line) => `<a:p><a:r><a:t>${xmlEscape(line)}</a:t></a:r></a:p>`).join("");
  const ph = placeholder ? `<p:ph type="${attrEscape(placeholder.type || "body")}" idx="${Number(placeholder.idx || 1)}"/>` : "";
  return `<p:sp><p:nvSpPr><p:cNvPr id="${index + 2}" name="${attrEscape(name)}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${attrEscape(geometry === "textbox" ? "rect" : geometry)}"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs || "<a:p/>"}</p:txBody></p:sp>`;
}

function pptxPictureXml(index, name, alt, position, relId) {
  const p = position;
  const x = Math.round(p.left * 9525), y = Math.round(p.top * 9525), cx = Math.round(p.width * 9525), cy = Math.round(p.height * 9525);
  return `<p:pic><p:nvPicPr><p:cNvPr id="${index + 2}" name="${attrEscape(name)}" descr="${attrEscape(alt)}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

function pptxConnectorXml(index, connector) {
  const x1 = Math.round(connector.start.x * 9525);
  const y1 = Math.round(connector.start.y * 9525);
  const x2 = Math.round(connector.end.x * 9525);
  const y2 = Math.round(connector.end.y * 9525);
  const x = Math.min(x1, x2), y = Math.min(y1, y2), cx = Math.max(1, Math.abs(x2 - x1)), cy = Math.max(1, Math.abs(y2 - y1));
  const stroke = String(connector.line?.fill || connector.line?.color || "#334155").replace(/^#/, "");
  const width = Math.round((connector.line?.width ?? 2) * 12700);
  const flipH = x2 < x1 ? ` flipH="1"` : "";
  const flipV = y2 < y1 ? ` flipV="1"` : "";
  const arrow = connector.line?.endArrow ? `<a:tailEnd type="triangle"/>` : "";
  return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${index + 2}" name="${attrEscape(connector.name || connector.id)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm${flipH}${flipV}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="${width}"><a:solidFill><a:srgbClr val="${attrEscape(stroke)}"/></a:solidFill>${arrow}</a:ln></p:spPr><p:extLst><p:ext uri="urn:open-office-artifact:connector" startX="${connector.start.x}" startY="${connector.start.y}" endX="${connector.end.x}" endY="${connector.end.y}" startTargetId="${attrEscape(connector.startTargetId || "")}" endTargetId="${attrEscape(connector.endTargetId || "")}"/></p:extLst></p:cxnSp>`;
}

function pptxNotesSlideXml(slide) {
  const paragraphs = String(slide.speakerNotes.text || "").split(/\r?\n/).map((line) => `<a:p><a:r><a:t>${xmlEscape(line)}</a:t></a:r></a:p>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs || "<a:p/>"}</p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
}

function pptxCommentsXml(slide) {
  const comments = slide.comments.items.flatMap((thread, threadIndex) => thread.comments.map((comment, commentIndex) => {
    const idx = threadIndex * 100 + commentIndex;
    return `<p:cm authorId="0" dt="${attrEscape(comment.created || thread.created)}" idx="${idx}" ooa:threadId="${attrEscape(thread.id)}" ooa:targetId="${attrEscape(thread.targetId || "")}" ooa:resolved="${thread.resolved ? 1 : 0}" xmlns:ooa="urn:open-office-artifact"><p:pos x="${100 + commentIndex * 24}" y="${100 + threadIndex * 32}"/><p:text>${xmlEscape(comment.text)}</p:text></p:cm>`;
  })).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${comments}</p:cmLst>`;
}

function parsePptxNotes(slide, xml = "") {
  const text = [...String(xml || "").matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)].map((match) => decodeXml(match[1])).join("\n");
  if (text) slide.addNotes(text);
}

function parsePptxComments(slide, xml = "") {
  const byThread = new Map();
  for (const match of String(xml || "").matchAll(/<p:cm\b([^>]*)>([\s\S]*?)<\/p:cm>/g)) {
    const attrs = match[1];
    const body = match[2];
    const threadId = /\bthreadId="([^"]+)"/.exec(attrs)?.[1] || /\booa:threadId="([^"]+)"/.exec(attrs)?.[1] || aid("pc");
    const targetId = /\btargetId="([^"]*)"/.exec(attrs)?.[1] || /\booa:targetId="([^"]*)"/.exec(attrs)?.[1] || undefined;
    const resolved = /\bresolved="1"/.test(attrs) || /\booa:resolved="1"/.test(attrs);
    const created = /\bdt="([^"]+)"/.exec(attrs)?.[1] || new Date(0).toISOString();
    const text = decodeXml(/<p:text[^>]*>([\s\S]*?)<\/p:text>/.exec(body)?.[1] || "");
    const thread = byThread.get(threadId) || { id: threadId, targetId, resolved, created, comments: [] };
    thread.comments.push({ author: "User", text, created });
    byThread.set(threadId, thread);
  }
  for (const thread of byThread.values()) slide.comments.addThread(thread.targetId, thread.comments[0]?.text || "", thread);
}

function pptxTableXml(index, table) {
  const p = table.position;
  const x = Math.round(p.left * 9525), y = Math.round(p.top * 9525), cx = Math.round(p.width * 9525), cy = Math.round(p.height * 9525);
  const colWidth = Math.max(1, Math.floor(cx / Math.max(1, table.columns)));
  const rowHeight = Math.max(1, Math.floor(cy / Math.max(1, table.rows)));
  const grid = Array.from({ length: Math.max(1, table.columns) }, () => `<a:gridCol w="${colWidth}"/>`).join("");
  const rows = Array.from({ length: Math.max(1, table.rows) }, (_, rowIndex) => {
    const cells = Array.from({ length: Math.max(1, table.columns) }, (_, colIndex) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(table.values[rowIndex]?.[colIndex] ?? "")}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`).join("");
    return `<a:tr h="${rowHeight}">${cells}</a:tr>`;
  }).join("");
  const firstRow = table.styleOptions.headerRow ? 1 : 0;
  const bandRow = table.styleOptions.bandedRows ? 1 : 0;
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${index + 2}" name="${attrEscape(table.name || table.id)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="${firstRow}" bandRow="${bandRow}"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

function pptxChartFrameXml(index, name, position, relId) {
  const p = position;
  const x = Math.round(p.left * 9525), y = Math.round(p.top * 9525), cx = Math.round(p.width * 9525), cy = Math.round(p.height * 9525);
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${index + 2}" name="${attrEscape(name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${relId}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

function pptxChartXml(chart) {
  const chartElementName = chart.chartType === "line" ? "lineChart" : "barChart";
  const grouping = chart.chartType === "line" ? "<c:grouping val=\"standard\"/>" : "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>";
  const seriesXml = (chart.series.length ? chart.series : [{ name: chart.title || "Series", values: [] }]).map((series, index) => {
    const values = series.values || [];
    const categories = series.categories || chart.categories || values.map((_, i) => String(i + 1));
    const catPts = categories.map((category, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${xmlEscape(category)}</c:v></c:pt>`).join("");
    const valPts = values.map((value, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${Number(value) || 0}</c:v></c:pt>`).join("");
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:v>${xmlEscape(series.name || `Series ${index + 1}`)}</c:v></c:tx><c:cat><c:strLit><c:ptCount val="${categories.length}"/>${catPts}</c:strLit></c:cat><c:val><c:numLit><c:ptCount val="${values.length}"/>${valPts}</c:numLit></c:val></c:ser>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(chart.title || chart.chartType)}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:layout/><c:${chartElementName}>${grouping}${seriesXml}<c:axId val="1"/><c:axId val="2"/></c:${chartElementName}><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:layout/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

function slideXml(slide, imageParts = [], chartParts = []) {
  const imageRelById = new Map(imageParts.map((part) => [part.image.id, part.slideRelId]));
  const chartRelById = new Map(chartParts.map((part) => [part.chart.id, part.slideRelId]));
  const elements = [...slide.connectors.items, ...slide.shapes.items, ...slide.tables.items, ...slide.charts.items, ...slide.images.items];
  const shapes = elements.map((element, index) => element.toPptxShape(index, imageRelById.get(element.id) || chartRelById.get(element.id))).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld></p:sld>`;
}

function pptxFrameFromXml(part, fallback = { left: 0, top: 0, width: 160, height: 80 }) {
  const off = /<a:off[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/.exec(part);
  const ext = /<a:ext[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(part);
  return { left: off ? Number(off[1]) / 9525 : fallback.left, top: off ? Number(off[2]) / 9525 : fallback.top, width: ext ? Number(ext[1]) / 9525 : fallback.width, height: ext ? Number(ext[2]) / 9525 : fallback.height };
}

function pptxRelationshipTarget(rels, relId) {
  const rel = rels.find((item) => item.id === relId);
  if (!rel?.target) return undefined;
  const target = rel.target.replace(/^\//, "");
  return target.startsWith("ppt/") ? target : path.posix.normalize(`ppt/slides/${target}`).replace(/^\.\//, "");
}

function parsePptxTheme(presentation, xml = "") {
  const name = decodeXml(/<a:theme[^>]*name="([^"]*)"/.exec(xml)?.[1] || "Imported Theme");
  const colors = {};
  for (const key of ["accent1", "accent2", "accent3"]) {
    const value = new RegExp(`<a:${key}>[\\s\\S]*?<a:srgbClr[^>]*val="([^"]+)"`).exec(xml)?.[1];
    if (value) colors[key] = `#${value}`;
  }
  const tx1 = /<a:dk1>[\s\S]*?<a:srgbClr[^>]*val="([^"]+)"/.exec(xml)?.[1];
  const bg1 = /<a:lt1>[\s\S]*?<a:srgbClr[^>]*val="([^"]+)"/.exec(xml)?.[1];
  if (tx1) colors.tx1 = `#${tx1}`;
  if (bg1) colors.bg1 = `#${bg1}`;
  const fonts = {
    major: decodeXml(/<a:majorFont>[\s\S]*?<a:latin[^>]*typeface="([^"]*)"/.exec(xml)?.[1] || "Aptos Display"),
    minor: decodeXml(/<a:minorFont>[\s\S]*?<a:latin[^>]*typeface="([^"]*)"/.exec(xml)?.[1] || "Aptos"),
  };
  return new PresentationTheme(presentation, { name, colors, fonts });
}

function parsePptxSlideLayout(presentation, xml = "", fallbackId = "imported-layout") {
  const text = String(xml || "");
  const name = decodeXml(/<p:cSld[^>]*name="([^"]*)"/.exec(text)?.[1] || fallbackId);
  const type = /<p:sldLayout[^>]*type="([^"]*)"/.exec(text)?.[1] || "custom";
  const placeholders = [...text.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].flatMap((match, index) => {
    const part = match[0];
    const phAttrs = /<p:ph\b([^>]*)\/?>(?:<\/p:ph>)?/.exec(part)?.[1];
    if (!phAttrs) return [];
    const phType = /\btype="([^"]+)"/.exec(phAttrs)?.[1] || "body";
    const phIdx = Number(/\bidx="([^"]+)"/.exec(phAttrs)?.[1] || index + 1);
    const phName = decodeXml(/<p:cNvPr[^>]*name="([^"]*)"/.exec(part)?.[1] || `${phType} placeholder`);
    const phText = [...part.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decodeXml(m[1])).join("");
    return [{ type: phType, name: phName, text: phText, idx: phIdx, position: pptxFrameFromXml(part, { left: 80, top: 80 + index * 80, width: 640, height: 64 }) }];
  });
  const existing = presentation.layouts.getItem(name);
  if (existing) return existing;
  return presentation.layouts.add({ id: fallbackId, name, type, placeholders });
}

function parsePptxTableGraphic(slide, part) {
  const name = decodeXml(/<p:cNvPr[^>]*name="([^"]*)"/.exec(part)?.[1] || "");
  const values = [...part.matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map((rowMatch) => [...rowMatch[0].matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)].map((cellMatch) => decodeXml([...cellMatch[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((t) => t[1]).join(""))));
  if (!values.length) return;
  slide.tables.add({ name, position: pptxFrameFromXml(part, { left: 0, top: 0, width: 320, height: 160 }), values, rows: values.length, columns: Math.max(1, ...values.map((row) => row.length)) });
}

async function parsePptxChartGraphic(slide, part, context) {
  const relId = /<c:chart[^>]*r:id="([^"]+)"/.exec(part)?.[1];
  const target = pptxRelationshipTarget(context.rels, relId);
  const chartXml = target ? await context.zip.file(target)?.async("text") : "";
  const chartType = /<c:lineChart>/.test(chartXml) ? "line" : "bar";
  const title = decodeXml(/<c:title>[\s\S]*?<a:t>([\s\S]*?)<\/a:t>[\s\S]*?<\/c:title>/.exec(chartXml)?.[1] || "");
  const series = [...String(chartXml || "").matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((seriesMatch, index) => {
    const seriesXml = seriesMatch[0];
    const name = decodeXml(/<c:tx>[\s\S]*?<c:v>([\s\S]*?)<\/c:v>[\s\S]*?<\/c:tx>/.exec(seriesXml)?.[1] || `Series ${index + 1}`);
    const categories = [...( /<c:cat>([\s\S]*?)<\/c:cat>/.exec(seriesXml)?.[1] || "").matchAll(/<c:v>([\s\S]*?)<\/c:v>/g)].map((m) => decodeXml(m[1]));
    const values = [...( /<c:val>([\s\S]*?)<\/c:val>/.exec(seriesXml)?.[1] || "").matchAll(/<c:v>([\s\S]*?)<\/c:v>/g)].map((m) => Number(decodeXml(m[1])) || 0);
    return { name, values, categories };
  });
  const name = decodeXml(/<p:cNvPr[^>]*name="([^"]*)"/.exec(part)?.[1] || title || "chart");
  slide.charts.add(chartType, { name, title, position: pptxFrameFromXml(part, { left: 0, top: 0, width: 360, height: 220 }), categories: series[0]?.categories || [], series });
}

async function parsePptxPicture(slide, part, context) {
  const attrs = /<p:cNvPr\b([^>]*)\/>/.exec(part)?.[1] || "";
  const name = decodeXml(/\bname="([^"]*)"/.exec(attrs)?.[1] || "");
  const alt = decodeXml(/\bdescr="([^"]*)"/.exec(attrs)?.[1] || "");
  const relId = /<a:blip[^>]*r:embed="([^"]+)"/.exec(part)?.[1];
  const target = pptxRelationshipTarget(context.rels, relId);
  const bytes = target ? await context.zip.file(target)?.async("uint8array") : undefined;
  const extension = /\.([A-Za-z0-9+]+)$/.exec(target || "")?.[1] || "png";
  slide.images.add({ name, alt, position: pptxFrameFromXml(part, { left: 0, top: 0, width: 320, height: 180 }), dataUrl: bytes ? `data:${imageContentTypeFromExtension(extension)};base64,${Buffer.from(bytes).toString("base64")}` : undefined, uri: bytes ? undefined : target });
}

function parsePptxConnector(slide, part) {
  const name = decodeXml(/<p:cNvPr[^>]*name="([^"]*)"/.exec(part)?.[1] || "");
  const extAttrs = [...part.matchAll(/<p:ext\b([^>]*)>/g)].map((match) => match[1]).find((attrs) => attrs.includes("urn:open-office-artifact:connector"));
  const attrs = extAttrs || "";
  const startX = Number(/\bstartX="([^"]+)"/.exec(attrs)?.[1]);
  const startY = Number(/\bstartY="([^"]+)"/.exec(attrs)?.[1]);
  const endX = Number(/\bendX="([^"]+)"/.exec(attrs)?.[1]);
  const endY = Number(/\bendY="([^"]+)"/.exec(attrs)?.[1]);
  const frame = pptxFrameFromXml(part, { left: 0, top: 0, width: 160, height: 1 });
  const start = Number.isFinite(startX) && Number.isFinite(startY) ? { x: startX, y: startY } : { x: frame.left, y: frame.top };
  const end = Number.isFinite(endX) && Number.isFinite(endY) ? { x: endX, y: endY } : { x: frame.left + frame.width, y: frame.top + frame.height };
  const startTargetId = decodeXml(/\bstartTargetId="([^"]*)"/.exec(attrs)?.[1] || "") || undefined;
  const endTargetId = decodeXml(/\bendTargetId="([^"]*)"/.exec(attrs)?.[1] || "") || undefined;
  slide.connectors.add({ name, start, end, startTargetId, endTargetId, line: { fill: "#334155", width: 2, endArrow: /<a:tailEnd/.test(part) ? "triangle" : undefined } });
}

async function parseSlideXml(slide, xml, context = { rels: [], zip: undefined }) {
  for (const match of xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const part = match[0];
    const name = decodeXml(/<p:cNvPr[^>]*name="([^"]*)"/.exec(part)?.[1] || "");
    const text = [...part.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decodeXml(m[1])).join("");
    const phAttrs = /<p:ph\b([^>]*)\/?>(?:<\/p:ph>)?/.exec(part)?.[1];
    const placeholder = phAttrs ? { type: /\btype="([^"]+)"/.exec(phAttrs)?.[1] || "body", idx: Number(/\bidx="([^"]+)"/.exec(phAttrs)?.[1] || 1), name } : undefined;
    const shape = slide.shapes.add({ name, position: pptxFrameFromXml(part), placeholder });
    shape.text = text;
  }
  for (const match of xml.matchAll(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g)) {
    const part = match[0];
    if (part.includes("/drawingml/2006/table")) parsePptxTableGraphic(slide, part);
    else if (part.includes("/drawingml/2006/chart")) await parsePptxChartGraphic(slide, part, context);
  }
  for (const match of xml.matchAll(/<p:pic>[\s\S]*?<\/p:pic>/g)) await parsePptxPicture(slide, match[0], context);
  for (const match of xml.matchAll(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g)) parsePptxConnector(slide, match[0]);
}

class DocumentStyleCollection {
  constructor(styles = {}) {
    this.items = new Map(Object.entries({
      Normal: { id: "Normal", name: "Normal", type: "paragraph", fontSize: 22, fontFamily: "Aptos" },
      Title: { id: "Title", name: "Title", type: "paragraph", fontSize: 48, bold: true, fontFamily: "Aptos Display" },
      Heading1: { id: "Heading1", name: "Heading 1", type: "paragraph", fontSize: 32, bold: true, fontFamily: "Aptos Display" },
      Heading2: { id: "Heading2", name: "Heading 2", type: "paragraph", fontSize: 26, bold: true, fontFamily: "Aptos" },
      ...styles,
    }));
  }

  add(id, config = {}) { const style = { id, name: config.name || id, type: config.type || "paragraph", ...config }; this.items.set(id, style); return style; }
  get(id) { return this.items.get(id); }
  values() { return [...this.items.values()]; }
}

class DocumentTableCell {
  constructor(table, row, column) { this.table = table; this.row = row; this.column = column; }
  get value() { return this.table.values[this.row]?.[this.column] ?? ""; }
  set value(value) { this.table.ensureCell(this.row, this.column); this.table.values[this.row][this.column] = value; }
}

class DocumentTableBlock {
  constructor(document, config = {}) {
    this.document = document;
    this.kind = "table";
    this.id = config.id || aid("dtb");
    this.name = config.name || "";
    this.styleId = config.styleId || config.style || "TableGrid";
    this.values = (config.values || Array.from({ length: config.rows || 1 }, () => Array.from({ length: config.columns || 1 }, () => ""))).map((row) => [...row]);
    this.rows = this.values.length;
    this.columns = Math.max(0, ...this.values.map((row) => row.length));
  }

  ensureCell(row, column) { while (this.values.length <= row) this.values.push([]); while (this.values[row].length <= column) this.values[row].push(""); this.rows = this.values.length; this.columns = Math.max(this.columns, column + 1); }
  getCell(row, column) { return new DocumentTableCell(this, row, column); }
  inspectRecord(index) { return { kind: "table", id: this.id, index, name: this.name || undefined, rows: this.rows, cols: this.columns, styleId: this.styleId, values: this.values }; }
  toProto() { return { kind: "table", id: this.id, name: this.name, styleId: this.styleId, values: this.values }; }
}

class DocumentParagraphBlock {
  constructor(document, text, config = {}) {
    this.document = document;
    this.kind = "paragraph";
    this.id = config.id || aid("dp");
    this.text = String(text ?? "");
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "paragraph", id: this.id, index, name: this.name || undefined, styleId: this.styleId, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: "paragraph", id: this.id, name: this.name, styleId: this.styleId, text: this.text }; }
}

class DocumentChangeBlock {
  constructor(document, changeType, text, config = {}) {
    this.document = document;
    this.kind = "change";
    this.id = config.id || aid("dchg");
    const rawType = String(changeType ?? config.changeType ?? config.type ?? "insert").toLowerCase();
    this.changeType = rawType === "delete" || rawType === "deletion" || rawType === "del" ? "delete" : "insert";
    this.text = String(text ?? "");
    this.author = config.author || "User";
    this.date = config.date || new Date().toISOString();
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "change", id: this.id, index, name: this.name || undefined, changeType: this.changeType, author: this.author, date: this.date, styleId: this.styleId, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: "change", id: this.id, name: this.name, changeType: this.changeType, author: this.author, date: this.date, styleId: this.styleId, text: this.text }; }
}

class DocumentListItemBlock {
  constructor(document, text, config = {}) {
    this.document = document;
    this.kind = "listItem";
    this.id = config.id || aid("dli");
    this.text = String(text ?? "");
    this.level = Number(config.level ?? 0);
    this.listType = config.listType || config.type || "bullet";
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "listItem", id: this.id, index, name: this.name || undefined, styleId: this.styleId, listType: this.listType, level: this.level, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: "listItem", id: this.id, name: this.name, styleId: this.styleId, listType: this.listType, level: this.level, text: this.text }; }
}

class DocumentHyperlinkBlock {
  constructor(document, text, url, config = {}) {
    this.document = document;
    this.kind = "hyperlink";
    this.id = config.id || aid("dhl");
    this.text = String(text ?? "");
    this.url = String(url ?? config.url ?? "");
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "hyperlink", id: this.id, index, name: this.name || undefined, styleId: this.styleId, text: this.text, url: this.url, textChars: this.text.length }; }
  toProto() { return { kind: "hyperlink", id: this.id, name: this.name, styleId: this.styleId, text: this.text, url: this.url }; }
}

class DocumentFieldBlock {
  constructor(document, instruction, display, config = {}) {
    this.document = document;
    this.kind = "field";
    this.id = config.id || aid("dfld");
    this.instruction = String(instruction ?? config.instruction ?? "PAGE");
    this.display = String(display ?? config.display ?? "1");
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  get text() { return this.display; }
  inspectRecord(index) { return { kind: "field", id: this.id, index, name: this.name || undefined, styleId: this.styleId, instruction: this.instruction, display: this.display }; }
  toProto() { return { kind: "field", id: this.id, name: this.name, styleId: this.styleId, instruction: this.instruction, display: this.display }; }
}

class DocumentCitationBlock {
  constructor(document, text, metadata = {}, config = {}) {
    this.document = document;
    this.kind = "citation";
    this.id = config.id || aid("dct");
    this.text = String(text ?? "");
    this.metadata = { ...metadata };
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "citation", id: this.id, index, name: this.name || undefined, styleId: this.styleId, text: this.text, metadata: this.metadata, textChars: this.text.length }; }
  toProto() { return { kind: "citation", id: this.id, name: this.name, styleId: this.styleId, text: this.text, metadata: this.metadata }; }
}

class DocumentImageBlock {
  constructor(document, config = {}) {
    this.document = document;
    this.kind = "image";
    this.id = config.id || aid("dim");
    this.name = config.name || "";
    this.dataUrl = config.dataUrl;
    this.uri = config.uri;
    this.prompt = config.prompt;
    this.alt = config.alt || config.altText || config.name || "image";
    this.widthPx = Number(config.widthPx || config.width || 240);
    this.heightPx = Number(config.heightPx || config.height || 160);
    this.styleId = config.styleId || config.style || "Normal";
  }

  inspectRecord(index) { return { kind: "image", id: this.id, index, name: this.name || undefined, styleId: this.styleId, alt: this.alt, uri: this.uri, prompt: this.prompt, widthPx: this.widthPx, heightPx: this.heightPx, hasDataUrl: Boolean(this.dataUrl) }; }
  toProto() { return { kind: "image", id: this.id, name: this.name, styleId: this.styleId, dataUrl: this.dataUrl, uri: this.uri, prompt: this.prompt, alt: this.alt, widthPx: this.widthPx, heightPx: this.heightPx }; }
}

class DocumentSectionBlock {
  constructor(document, config = {}) {
    this.document = document;
    this.kind = "section";
    this.id = config.id || aid("dsec");
    this.name = config.name || "";
    this.breakType = config.breakType || config.type || "nextPage";
    this.orientation = config.orientation || "portrait";
    const pageSize = config.pageSize || {};
    this.pageSize = {
      widthTwips: Number(config.widthTwips || pageSize.widthTwips || (this.orientation === "landscape" ? 15840 : 12240)),
      heightTwips: Number(config.heightTwips || pageSize.heightTwips || (this.orientation === "landscape" ? 12240 : 15840)),
    };
    const margins = config.margins || {};
    this.margins = {
      top: Number(margins.top ?? config.marginTop ?? 1440),
      right: Number(margins.right ?? config.marginRight ?? 1440),
      bottom: Number(margins.bottom ?? config.marginBottom ?? 1440),
      left: Number(margins.left ?? config.marginLeft ?? 1440),
    };
  }

  inspectRecord(index) { return { kind: "section", id: this.id, index, name: this.name || undefined, breakType: this.breakType, orientation: this.orientation, pageSize: this.pageSize, margins: this.margins }; }
  toProto() { return { kind: "section", id: this.id, name: this.name, breakType: this.breakType, orientation: this.orientation, pageSize: this.pageSize, margins: this.margins }; }
}

class DocumentHeaderFooterBlock {
  constructor(document, kind, text, config = {}) {
    this.document = document;
    this.kind = kind;
    this.id = config.id || aid(kind === "header" ? "dh" : "df");
    this.text = String(text ?? "");
    this.name = config.name || kind;
    this.styleId = config.styleId || "Normal";
  }

  inspectRecord(index) { return { kind: this.kind, id: this.id, index, name: this.name || undefined, styleId: this.styleId, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: this.kind, id: this.id, name: this.name, styleId: this.styleId, text: this.text }; }
}

class DocumentComment {
  constructor(document, targetId, text, config = {}) {
    this.document = document;
    this.kind = "comment";
    this.id = config.id || aid("dc");
    this.targetId = targetId;
    this.author = config.author || "User";
    this.text = String(text ?? "");
    this.resolved = Boolean(config.resolved);
  }

  inspectRecord() { return { kind: "comment", id: this.id, targetId: this.targetId, author: this.author, resolved: this.resolved, textPreview: this.text.slice(0, 300) }; }
  toProto() { return { kind: "comment", id: this.id, targetId: this.targetId, author: this.author, text: this.text, resolved: this.resolved }; }
}

function documentBlockHeight(document, block, pageWidth = 612, margin = 72) {
  if (block.kind === "table") return Math.max(24, block.rows * 24 + 16);
  if (block.kind === "image") return Math.max(32, Math.min(360, Number(block.heightPx) || 160)) + 20;
  if (block.kind === "section") return 34;
  if (block.kind === "change") return 22;
  const style = document.styles.get(block.styleId) || document.styles.get("Normal") || {};
  const fontSize = Math.max(10, (style.fontSize || 22) / 2);
  const text = block.text || block.display || "";
  const charsPerLine = Math.max(8, Math.floor((pageWidth - margin * 2) / (fontSize * 0.55)));
  const lines = String(text).split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);
  return Math.max(20, lines * fontSize * 1.6);
}

function documentLayoutJson(document, options = {}) {
  const pageWidth = Number(options.pageWidth || 612);
  const pageHeight = Number(options.pageHeight || 792);
  const margin = Number(options.margin || 72);
  const pages = [];
  const elements = [];
  let page = 1;
  let y = margin;
  const ensurePage = () => {
    if (!pages.find((item) => item.page === page)) pages.push({ page, width: pageWidth, height: pageHeight, margin, headers: document.headers.map((header) => header.id), footers: document.footers.map((footer) => footer.id) });
  };
  ensurePage();
  for (const block of document.blocks) {
    const height = documentBlockHeight(document, block, pageWidth, margin);
    if (y + height > pageHeight - margin && y > margin) { page += 1; y = margin; ensurePage(); }
    elements.push({ kind: "layoutElement", id: block.id, blockKind: block.kind, name: block.name || undefined, page, bbox: [margin, y, pageWidth - margin * 2, height], styleId: block.styleId, textPreview: String(block.text || block.display || block.alt || "").slice(0, 120) });
    y += height;
    if (block.kind === "section" && block.breakType === "nextPage") { page += 1; y = margin; ensurePage(); }
  }
  return { schema: "open-office-artifact.document-layout/v1", unit: "px", document: { id: document.id, name: document.name, designPreset: document.designPreset }, pages, elements };
}

function documentLayoutRecords(document, options = {}) {
  const layout = documentLayoutJson(document, options);
  return [
    { kind: "layout", id: `${document.id}/layout`, pages: layout.pages.length, elements: layout.elements.length, designPreset: document.designPreset },
    ...layout.pages.map((page) => ({ kind: "page", id: `${document.id}/page/${page.page}`, ...page })),
    ...layout.elements,
  ];
}

export class DocumentModel {
  constructor(options = {}) {
    this.id = aid("doc");
    this.name = options.name || "New document";
    this.designPreset = options.designPreset || "default";
    this.styles = new DocumentStyleCollection(options.styles || {});
    this.blocks = [];
    this.comments = [];
    this.headers = [];
    this.footers = [];
    const sourceBlocks = options.blocks || (options.paragraphs ? options.paragraphs.map((text, index) => ({ kind: "paragraph", text, styleId: index === 0 ? "Title" : "Normal" })) : [{ kind: "paragraph", text: "Start writing here...", styleId: "Normal" }]);
    for (const block of sourceBlocks) {
      if (block.kind === "table") this.addTable(block);
      else if (block.kind === "listItem") this.addListItem(block.text ?? "", block);
      else if (block.kind === "hyperlink") this.addHyperlink(block.text ?? "", block.url, block);
      else if (block.kind === "field") this.addField(block.instruction, block.display, block);
      else if (block.kind === "citation") this.addCitation(block.text ?? "", block.metadata || {}, block);
      else if (block.kind === "image") this.addImage(block);
      else if (block.kind === "section") this.addSection(block);
      else if (block.kind === "change") this.addChange(block.changeType || block.type, block.text ?? "", block);
      else this.addParagraph(block.text ?? "", block);
    }
    for (const header of options.headers || []) this.addHeader(header.text, header);
    for (const footer of options.footers || []) this.addFooter(footer.text, footer);
    for (const comment of options.comments || []) this.addComment(comment.targetId, comment.text, comment);
  }

  static create(options = {}) { return new DocumentModel(options); }
  get paragraphs() { return this.blocks.filter((block) => block.kind === "paragraph").map((block) => block.text); }

  applyDesignPreset(name = "report", options = {}) {
    const presetName = String(name || "report");
    const presets = {
      report: {
        Title: { fontSize: 52, bold: true, fontFamily: "Aptos Display" },
        Heading1: { fontSize: 34, bold: true, fontFamily: "Aptos Display" },
        Heading2: { fontSize: 28, bold: true, fontFamily: "Aptos" },
        Normal: { fontSize: 22, fontFamily: "Aptos" },
        Caption: { name: "Caption", fontSize: 18, italic: true, fontFamily: "Aptos" },
        Callout: { name: "Callout", fontSize: 24, bold: true, fontFamily: "Aptos" },
        TableGrid: { name: "Table Grid", type: "table", fontSize: 20, fontFamily: "Aptos" },
      },
      memo: {
        Title: { fontSize: 44, bold: true, fontFamily: "Aptos Display" },
        Heading1: { fontSize: 30, bold: true, fontFamily: "Aptos" },
        Normal: { fontSize: 21, fontFamily: "Aptos" },
        Metadata: { name: "Metadata", fontSize: 18, fontFamily: "Aptos" },
        TableGrid: { name: "Table Grid", type: "table", fontSize: 20, fontFamily: "Aptos" },
      },
    };
    const preset = presets[presetName] || presets.report;
    for (const [id, style] of Object.entries(preset)) this.styles.add(id, { ...(this.styles.get(id) || {}), ...style, ...(options.styles?.[id] || {}) });
    this.designPreset = presetName;
    return this;
  }

  layoutJson(options = {}) {
    return documentLayoutJson(this, options);
  }

  addParagraph(text, config = {}) { const block = new DocumentParagraphBlock(this, text, config); this.blocks.push(block); return block; }
  addListItem(text, config = {}) { const block = new DocumentListItemBlock(this, text, config); this.blocks.push(block); return block; }
  addList(items = [], config = {}) { return items.map((item) => this.addListItem(typeof item === "string" ? item : item.text, { ...config, ...(typeof item === "string" ? {} : item) })); }
  addHyperlink(text, url, config = {}) { const block = new DocumentHyperlinkBlock(this, text, url, config); this.blocks.push(block); return block; }
  addField(instruction, display, config = {}) { const block = new DocumentFieldBlock(this, instruction, display, config); this.blocks.push(block); return block; }
  addCitation(text, metadata = {}, config = {}) { const block = new DocumentCitationBlock(this, text, metadata, config); this.blocks.push(block); return block; }
  addImage(config = {}) { const block = new DocumentImageBlock(this, config); this.blocks.push(block); return block; }
  addSection(config = {}) { const block = new DocumentSectionBlock(this, config); this.blocks.push(block); return block; }
  addPageBreakSection(config = {}) { return this.addSection({ ...config, breakType: "nextPage" }); }
  addChange(changeType, text, config = {}) { const block = new DocumentChangeBlock(this, changeType, text, config); this.blocks.push(block); return block; }
  addInsertion(text, config = {}) { return this.addChange("insert", text, config); }
  addDeletion(text, config = {}) { return this.addChange("delete", text, config); }
  addTable(config = {}) { const block = new DocumentTableBlock(this, config); this.blocks.push(block); return block; }
  addHeader(text, config = {}) { const block = new DocumentHeaderFooterBlock(this, "header", text, config); this.headers.push(block); return block; }
  addFooter(text, config = {}) { const block = new DocumentHeaderFooterBlock(this, "footer", text, config); this.footers.push(block); return block; }
  addComment(target, text, config = {}) { const targetId = typeof target === "string" ? target : target?.id; const comment = new DocumentComment(this, targetId, text, config); this.comments.push(comment); return comment; }
  resolve(id) { return this.id === id ? this : this.blocks.find((block) => block.id === id) || this.headers.find((block) => block.id === id) || this.footers.find((block) => block.id === id) || this.comments.find((comment) => comment.id === id) || this.styles.get(id); }

  toProto() { return { id: this.id, name: this.name, designPreset: this.designPreset, styles: Object.fromEntries(this.styles.values().map((style) => [style.id, style])), blocks: this.blocks.map((block) => block.toProto()), headers: this.headers.map((block) => block.toProto()), footers: this.footers.map((block) => block.toProto()), comments: this.comments.map((comment) => comment.toProto()) }; }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["paragraph", "table", "listItem", "hyperlink", "field", "citation", "image", "section", "change", "comment", "header", "footer"]);
    const records = [];
    if (kinds.has("document")) records.push({ kind: "document", id: this.id, name: this.name, blocks: this.blocks.length, designPreset: this.designPreset });
    if (kinds.has("layout")) records.push(...documentLayoutRecords(this, options));
    this.blocks.forEach((block, index) => { if (kinds.has(block.kind)) records.push(block.inspectRecord(index)); });
    if (kinds.has("header")) records.push(...this.headers.map((block, index) => block.inspectRecord(index)));
    if (kinds.has("footer")) records.push(...this.footers.map((block, index) => block.inspectRecord(index)));
    if (kinds.has("comment")) records.push(...this.comments.map((comment) => comment.inspectRecord()));
    if (kinds.has("style")) records.push(...this.styles.values().map((style) => ({ kind: "style", ...style })));
    return ndjson(filterInspectRecords(records, options), options.maxChars ?? Infinity);
  }

  verify(options = {}) {
    const issues = [];
    const knownStyleIds = new Set(this.styles.values().map((style) => style.id));
    const checkStyle = (block) => {
      if (block.styleId && !knownStyleIds.has(block.styleId)) {
        issues.push(verificationIssue("document", "unknownStyle", `${block.kind} ${block.id} references missing style ${block.styleId}.`, { severity: "warning", id: block.id, styleId: block.styleId }));
      }
    };
    if (this.blocks.length === 0) issues.push(verificationIssue("document", "emptyDocument", "Document has no body blocks."));
    for (const block of [...this.blocks, ...this.headers, ...this.footers]) checkStyle(block);
    for (const block of this.blocks) {
      if (block.kind === "paragraph" && /^\s*([-*•]|\d+[.)])\s+/.test(block.text)) {
        issues.push(verificationIssue("document", "fakeList", `Paragraph ${block.id} looks like a fake list item; use addListItem instead.`, { id: block.id }));
      }
      if (block.kind === "hyperlink" && !/^https?:\/\//.test(block.url)) {
        issues.push(verificationIssue("document", "invalidHyperlink", `Hyperlink ${block.id} is missing an absolute http(s) URL.`, { id: block.id, url: block.url }));
      }
      if (block.kind === "field" && !block.instruction.trim()) {
        issues.push(verificationIssue("document", "emptyField", `Field ${block.id} is missing an instruction.`, { id: block.id }));
      }
      if (block.kind === "citation" && block.metadata?.url && !/^https?:\/\//.test(String(block.metadata.url))) {
        issues.push(verificationIssue("document", "invalidCitationUrl", `Citation ${block.id} has a non-http(s) URL.`, { severity: "warning", id: block.id, url: block.metadata.url }));
      }
      if (block.kind === "image") {
        if (!block.dataUrl && !block.uri && !block.prompt) issues.push(verificationIssue("document", "emptyImage", `Image ${block.id} has no dataUrl, uri, or prompt.`, { id: block.id }));
        if (block.dataUrl && !imageDataFromDataUrl(block.dataUrl)) issues.push(verificationIssue("document", "invalidImageDataUrl", `Image ${block.id} has an unsupported data URL.`, { id: block.id }));
        if (!Number.isFinite(block.widthPx) || !Number.isFinite(block.heightPx) || block.widthPx <= 0 || block.heightPx <= 0) issues.push(verificationIssue("document", "invalidImageDimensions", `Image ${block.id} has invalid dimensions.`, { id: block.id, widthPx: block.widthPx, heightPx: block.heightPx }));
      }
      if (block.kind === "section") {
        if (!["portrait", "landscape"].includes(block.orientation)) issues.push(verificationIssue("document", "invalidSectionOrientation", `Section ${block.id} has invalid orientation ${block.orientation}.`, { id: block.id, orientation: block.orientation }));
        if (!["nextPage", "continuous", "evenPage", "oddPage", ""].includes(block.breakType)) issues.push(verificationIssue("document", "invalidSectionBreak", `Section ${block.id} has invalid break type ${block.breakType}.`, { id: block.id, breakType: block.breakType }));
        for (const [side, value] of Object.entries(block.margins || {})) if (!Number.isFinite(value) || value < 0) issues.push(verificationIssue("document", "invalidSectionMargin", `Section ${block.id} has an invalid ${side} margin.`, { id: block.id, side, value }));
        for (const [dimension, value] of Object.entries(block.pageSize || {})) if (!Number.isFinite(value) || value <= 0) issues.push(verificationIssue("document", "invalidSectionPageSize", `Section ${block.id} has invalid ${dimension}.`, { id: block.id, dimension, value }));
        const horizontalMargins = Number(block.margins?.left || 0) + Number(block.margins?.right || 0);
        if (Number.isFinite(block.pageSize?.widthTwips) && horizontalMargins >= block.pageSize.widthTwips) issues.push(verificationIssue("document", "sectionMarginsExceedPage", `Section ${block.id} horizontal margins exceed page width.`, { id: block.id, margins: block.margins, pageSize: block.pageSize }));
      }
      if (block.kind === "table") {
        if (!block.rows || !block.columns) issues.push(verificationIssue("document", "emptyTable", `Table ${block.id} has no rows or columns.`, { id: block.id, rows: block.rows, columns: block.columns }));
        if (block.columns > 12) issues.push(verificationIssue("document", "wideTable", `Table ${block.id} has ${block.columns} columns and may not fit the page.`, { severity: "warning", id: block.id, columns: block.columns }));
        block.values.forEach((row, rowIndex) => {
          if (row.length !== block.columns) issues.push(verificationIssue("document", "raggedTableRows", `Table ${block.id} row ${rowIndex} has ${row.length} cells; expected ${block.columns}.`, { id: block.id, row: rowIndex, cells: row.length, expected: block.columns }));
          for (const cell of row) {
            if (String(cell ?? "").length > 240) issues.push(verificationIssue("document", "tableCellTooLong", `Table ${block.id} contains paragraph-like cell content.`, { severity: "warning", id: block.id }));
          }
        });
      }
    }
    const blockIds = new Set(this.blocks.map((block) => block.id));
    for (const comment of this.comments) {
      if (!blockIds.has(comment.targetId)) issues.push(verificationIssue("document", "danglingComment", `Comment ${comment.id} points at a missing block.`, { id: comment.id, targetId: comment.targetId }));
    }
    if (options.visualQa || options.renderQa) {
      const layout = this.layoutJson(options);
      for (const element of layout.elements) {
        const pageInfo = layout.pages.find((page) => page.page === element.page);
        if (pageInfo && element.bbox[3] > pageInfo.height - pageInfo.margin * 2) issues.push(verificationIssue("document", "layoutElementTooTall", `Block ${element.id} is taller than the usable page area.`, { severity: "warning", id: element.id, page: element.page, bbox: element.bbox }));
        if (pageInfo && element.bbox[1] + element.bbox[3] > pageInfo.height - pageInfo.margin + 1) issues.push(verificationIssue("document", "layoutElementOverflow", `Block ${element.id} may overflow page ${element.page}.`, { severity: "warning", id: element.id, page: element.page, bbox: element.bbox }));
      }
    }
    return verificationResult("document", issues, options);
  }

  help(query = "*", options = {}) {
    return helpArtifact("document", query, options);
  }

  async render(options = {}) {
    if (options.format === "layout") return new FileBlob(JSON.stringify(this.layoutJson(options), null, 2), { type: LAYOUT_MIME });
    const width = 612;
    const margin = 72;
    let y = 72;
    const parts = [`<rect width="100%" height="100%" fill="white"/>`];
    for (const header of this.headers) {
      parts.push(`<text x="${margin}" y="${Math.max(28, y - 36)}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(header.text)}</text>`);
    }
    let listCounters = new Map();
    for (const block of this.blocks) {
      if (block.kind === "paragraph") {
        const style = this.styles.get(block.styleId) || this.styles.get("Normal");
        const fontSize = Math.max(10, (style.fontSize || 22) / 2);
        parts.push(`<text x="${margin}" y="${y}" font-family="${xmlEscape(style.fontFamily || "Arial")}" font-size="${fontSize}" font-weight="${style.bold ? "700" : "400"}" fill="#111827">${xmlEscape(block.text)}</text>`);
        y += fontSize * 1.6;
      } else if (block.kind === "hyperlink") {
        parts.push(`<text x="${margin}" y="${y}" font-family="Arial" font-size="11" fill="#2563eb" text-decoration="underline">${xmlEscape(block.text)}</text>`);
        y += 20;
      } else if (block.kind === "field") {
        parts.push(`<text x="${margin}" y="${y}" font-family="Arial" font-size="11" fill="#334155">${xmlEscape(block.display)} (${xmlEscape(block.instruction)})</text>`);
        y += 20;
      } else if (block.kind === "citation") {
        parts.push(`<text x="${margin}" y="${y}" font-family="Arial" font-size="11" fill="#475569">${xmlEscape(block.text)}</text>`);
        y += 20;
      } else if (block.kind === "image") {
        const imageWidth = Math.max(16, Math.min(width - margin * 2, Number(block.widthPx) || 240));
        const imageHeight = Math.max(16, Math.min(360, Number(block.heightPx) || 160));
        if (block.dataUrl) parts.push(`<image href="${attrEscape(block.dataUrl)}" x="${margin}" y="${y}" width="${imageWidth}" height="${imageHeight}" preserveAspectRatio="xMidYMid meet"/>`);
        else parts.push(`<rect x="${margin}" y="${y}" width="${imageWidth}" height="${imageHeight}" fill="#fef3c7" stroke="#f59e0b"/>`);
        parts.push(`<text x="${margin + 8}" y="${y + 18}" font-family="Arial" font-size="11" fill="#92400e">${xmlEscape(block.alt || block.prompt || block.uri || block.name || "image")}</text>`);
        y += imageHeight + 20;
      } else if (block.kind === "section") {
        parts.push(`<line x1="${margin}" x2="${width - margin}" y1="${y}" y2="${y}" stroke="#94a3b8" stroke-dasharray="4 4"/>`);
        parts.push(`<text x="${margin}" y="${y + 16}" font-family="Arial" font-size="10" fill="#64748b">section break: ${xmlEscape(block.breakType)} ${xmlEscape(block.orientation)} ${block.pageSize.widthTwips}x${block.pageSize.heightTwips}</text>`);
        y += 34;
      } else if (block.kind === "change") {
        const marker = block.changeType === "delete" ? "-" : "+";
        const fill = block.changeType === "delete" ? "#dc2626" : "#047857";
        const decoration = block.changeType === "delete" ? " text-decoration=\"line-through\"" : "";
        parts.push(`<text x="${margin}" y="${y}" font-family="Arial" font-size="10" fill="#64748b">tracked ${xmlEscape(block.changeType)} by ${xmlEscape(block.author)}</text>`);
        parts.push(`<text x="${margin + 92}" y="${y}" font-family="Arial" font-size="11" fill="${fill}"${decoration}>${xmlEscape(marker)} ${xmlEscape(block.text)}</text>`);
        y += 22;
      } else if (block.kind === "listItem") {
        const key = `${block.listType}:${block.level}`;
        const next = (listCounters.get(key) || 0) + 1;
        listCounters.set(key, next);
        const marker = block.listType === "number" ? `${next}.` : "•";
        const x = margin + block.level * 24;
        parts.push(`<text x="${x}" y="${y}" font-family="Arial" font-size="11" fill="#111827">${xmlEscape(marker)}</text>`);
        parts.push(`<text x="${x + 22}" y="${y}" font-family="Arial" font-size="11" fill="#111827">${xmlEscape(block.text)}</text>`);
        y += 20;
      } else if (block.kind === "table") {
        const cellW = (width - margin * 2) / Math.max(1, block.columns);
        const cellH = 24;
        for (let r = 0; r < block.rows; r++) {
          for (let c = 0; c < block.columns; c++) {
            const x = margin + c * cellW;
            const fill = r === 0 ? "#f1f5f9" : "#ffffff";
            parts.push(`<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" fill="${fill}" stroke="#cbd5e1"/>`);
            parts.push(`<text x="${x + 5}" y="${y + 16}" font-family="Arial" font-size="11" fill="#111827">${xmlEscape(block.values[r]?.[c] ?? "")}</text>`);
          }
          y += cellH;
        }
        y += 16;
      }
    }
    const height = Math.max(792, y + 72);
    for (const footer of this.footers) {
      parts.push(`<text x="${margin}" y="${height - 36}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(footer.text)}</text>`);
    }
    return new FileBlob(`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${parts.join("")}</svg>`, { type: "image/svg+xml" });
  }
}

function docxContentTypes({ hasComments, hasHeader, hasFooter, hasNumbering, imageParts = [] }) {
  const imageDefaults = imageContentTypeDefaults(imageParts);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="json" ContentType="application/json"/>${imageDefaults}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${hasNumbering ? `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` : ""}${hasComments ? `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` : ""}${hasHeader ? `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` : ""}${hasFooter ? `<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>` : ""}</Types>`;
}

function docxStylesXml(document) {
  const styles = document.styles.values().map((style) => `<w:style w:type="paragraph" w:styleId="${attrEscape(style.id)}"><w:name w:val="${attrEscape(style.name || style.id)}"/><w:rPr>${style.bold ? "<w:b/>" : ""}<w:sz w:val="${Math.round(style.fontSize || 22)}"/><w:rFonts w:ascii="${attrEscape(style.fontFamily || "Aptos")}" w:hAnsi="${attrEscape(style.fontFamily || "Aptos")}"/></w:rPr></w:style>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${styles}</w:styles>`;
}

function docxNumberingXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>`;
}

function docxHeaderFooterXml(kind, blocks) {
  const tag = kind === "header" ? "hdr" : "ftr";
  const body = blocks.map((block) => `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr><w:r><w:t>${xmlEscape(block.text)}</w:t></w:r></w:p>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${body}</w:${tag}>`;
}

function docxParagraphXml(block, commentIndexes) {
  const commentStart = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeStart w:id="${id}"/>`).join("") : "";
  const commentEnd = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeEnd w:id="${id}"/>`).join("") : "";
  const refs = commentIndexes.length ? commentIndexes.map((id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`).join("") : "";
  const numPr = block.kind === "listItem" ? `<w:numPr><w:ilvl w:val="${Math.max(0, block.level || 0)}"/><w:numId w:val="${block.listType === "number" ? 2 : 1}"/></w:numPr>` : "";
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/>${numPr}</w:pPr>${commentStart}<w:r><w:t>${xmlEscape(block.text)}</w:t></w:r>${commentEnd}${refs}</w:p>`;
}

function docxRevisionId(block) {
  const raw = String(block?.id || block?.name || block?.text || "1");
  const sum = raw.split("").reduce((acc, ch) => ((acc * 31) + ch.charCodeAt(0)) >>> 0, 17);
  return Math.max(1, sum % 2147483647);
}

function docxChangeXml(block, commentIndexes = []) {
  const commentStart = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeStart w:id="${id}"/>`).join("") : "";
  const commentEnd = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeEnd w:id="${id}"/>`).join("") : "";
  const refs = commentIndexes.length ? commentIndexes.map((id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`).join("") : "";
  const tag = block.changeType === "delete" ? "del" : "ins";
  const textTag = block.changeType === "delete" ? "delText" : "t";
  const date = block.date ? ` w:date="${attrEscape(block.date)}"` : "";
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr>${commentStart}<w:${tag} w:id="${docxRevisionId(block)}" w:author="${attrEscape(block.author || "User")}"${date}><w:r><w:${textTag}>${xmlEscape(block.text)}</w:${textTag}></w:r></w:${tag}>${commentEnd}${refs}</w:p>`;
}

function collectDocxImageParts(document) {
  const parts = [];
  let imagePartId = 1;
  let relIndex = 1;
  for (const block of document.blocks.filter((item) => item.kind === "image")) {
    const data = imageDataFromDataUrl(block.dataUrl);
    if (!data) continue;
    parts.push({ image: block, imagePartId: imagePartId++, relId: `rIdImage${relIndex++}`, ...data });
  }
  return parts;
}

function docxImageXml(block, relId, commentIndexes = []) {
  const commentStart = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeStart w:id="${id}"/>`).join("") : "";
  const commentEnd = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeEnd w:id="${id}"/>`).join("") : "";
  const refs = commentIndexes.length ? commentIndexes.map((id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`).join("") : "";
  if (!relId) return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr>${commentStart}<w:r><w:t>${xmlEscape(block.alt || block.prompt || block.uri || block.name || "image")}</w:t></w:r>${commentEnd}${refs}</w:p>`;
  const cx = Math.round(Math.max(16, Number(block.widthPx) || 240) * 9525);
  const cy = Math.round(Math.max(16, Number(block.heightPx) || 160) * 9525);
  const docPrId = docxRevisionId(block);
  const name = block.name || `Image ${docPrId}`;
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr>${commentStart}<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${docPrId}" name="${attrEscape(name)}" descr="${attrEscape(block.alt || "")}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${docPrId}" name="${attrEscape(name)}" descr="${attrEscape(block.alt || "")}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>${commentEnd}${refs}</w:p>`;
}

function docxSectionPrXml(section, refs = "") {
  const size = section.pageSize || {};
  const margins = section.margins || {};
  const orient = section.orientation === "landscape" ? ` w:orient="landscape"` : "";
  const type = section.breakType ? `<w:type w:val="${attrEscape(section.breakType)}"/>` : "";
  return `<w:sectPr>${refs}${type}<w:pgSz w:w="${Math.round(size.widthTwips || 12240)}" w:h="${Math.round(size.heightTwips || 15840)}"${orient}/><w:pgMar w:top="${Math.round(margins.top ?? 1440)}" w:right="${Math.round(margins.right ?? 1440)}" w:bottom="${Math.round(margins.bottom ?? 1440)}" w:left="${Math.round(margins.left ?? 1440)}" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
}

function docxSectionXml(block, commentIndexes = []) {
  const commentStart = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeStart w:id="${id}"/>`).join("") : "";
  const commentEnd = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeEnd w:id="${id}"/>`).join("") : "";
  const refs = commentIndexes.length ? commentIndexes.map((id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`).join("") : "";
  return `<w:p><w:pPr>${docxSectionPrXml(block)}</w:pPr>${commentStart}${commentEnd}${refs}</w:p>`;
}

function docxHyperlinkXml(block, relId) {
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr><w:hyperlink r:id="${relId}"><w:r><w:rPr><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t>${xmlEscape(block.text)}</w:t></w:r></w:hyperlink></w:p>`;
}

function docxFieldXml(block) {
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr><w:fldSimple w:instr="${attrEscape(block.instruction)}"><w:r><w:t>${xmlEscape(block.display)}</w:t></w:r></w:fldSimple></w:p>`;
}

function docxCitationXml(block) {
  const label = block.metadata?.source ? `${block.text} (${block.metadata.source})` : block.text;
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr><w:bookmarkStart w:id="${Math.abs(block.id.split("").reduce((sum, ch) => sum + ch.charCodeAt(0), 0))}" w:name="${attrEscape(`OpenOfficeCitation_${block.id.replace(/[^A-Za-z0-9_]/g, "_")}`)}"/><w:r><w:t>${xmlEscape(label)}</w:t></w:r><w:bookmarkEnd w:id="${Math.abs(block.id.split("").reduce((sum, ch) => sum + ch.charCodeAt(0), 0))}"/></w:p>`;
}

function docxTableXml(block) {
  const grid = Array.from({ length: block.columns || 1 }, () => `<w:gridCol w:w="${Math.floor(9360 / Math.max(1, block.columns || 1))}"/>`).join("");
  const rows = block.values.map((row) => `<w:tr>${Array.from({ length: block.columns || row.length || 1 }, (_, column) => `<w:tc><w:tcPr><w:tcW w:w="${Math.floor(9360 / Math.max(1, block.columns || 1))}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${xmlEscape(row[column] ?? "")}</w:t></w:r></w:p></w:tc>`).join("")}</w:tr>`).join("");
  return `<w:tbl><w:tblPr><w:tblStyle w:val="${attrEscape(block.styleId || "TableGrid")}"/><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
}

function docxDocumentXml(document, relIds = {}) {
  const commentIndex = new Map(document.comments.map((comment, index) => [comment, index]));
  const body = document.blocks.map((block) => {
    if (block.kind === "table") return docxTableXml(block);
    if (block.kind === "hyperlink") return docxHyperlinkXml(block, relIds.hyperlinks?.get(block.id));
    if (block.kind === "field") return docxFieldXml(block);
    if (block.kind === "citation") return docxCitationXml(block);
    const indexes = document.comments.filter((comment) => comment.targetId === block.id).map((comment) => commentIndex.get(comment));
    if (block.kind === "image") return docxImageXml(block, relIds.images?.get(block.id), indexes);
    if (block.kind === "section") return docxSectionXml(block, indexes);
    if (block.kind === "change") return docxChangeXml(block, indexes);
    return docxParagraphXml(block, indexes);
  }).join("");
  const refs = `${relIds.header ? `<w:headerReference w:type="default" r:id="${relIds.header}"/>` : ""}${relIds.footer ? `<w:footerReference w:type="default" r:id="${relIds.footer}"/>` : ""}`;
  const finalSection = docxSectionPrXml({ pageSize: {}, margins: {}, breakType: "" }, refs);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body}${finalSection}</w:body></w:document>`;
}

function docxCommentsXml(document) {
  const comments = document.comments.map((comment, index) => `<w:comment w:id="${index}" w:author="${attrEscape(comment.author)}"><w:p><w:r><w:t>${xmlEscape(comment.text)}</w:t></w:r></w:p></w:comment>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${comments}</w:comments>`;
}

function parseDocxParagraph(part, commentTextById, imageByRelId = new Map()) {
  const styleId = /<w:pStyle[^>]*w:val="([^"]+)"/.exec(part)?.[1] || "Normal";
  const commentIds = [...part.matchAll(/<w:commentRangeStart[^>]*w:id="(\d+)"/g)].map((match) => match[1]);
  const sectionMatch = /<w:sectPr\b[^>]*>([\s\S]*?)<\/w:sectPr>/.exec(part);
  if (sectionMatch) {
    const sectionXml = sectionMatch[1] || "";
    const type = /<w:type[^>]*w:val="([^"]+)"/.exec(sectionXml)?.[1] || "nextPage";
    const sizeAttrs = /<w:pgSz\b([^>]*)\/>/.exec(sectionXml)?.[1] || "";
    const marginAttrs = /<w:pgMar\b([^>]*)\/>/.exec(sectionXml)?.[1] || "";
    const orientation = /w:orient="landscape"/.test(sizeAttrs) ? "landscape" : "portrait";
    const readAttr = (attrs, name, fallback) => Number(new RegExp(`\\bw:${name}="(\\d+)"`).exec(attrs)?.[1] || fallback);
    return { block: { kind: "section", breakType: type, orientation, pageSize: { widthTwips: readAttr(sizeAttrs, "w", orientation === "landscape" ? 15840 : 12240), heightTwips: readAttr(sizeAttrs, "h", orientation === "landscape" ? 12240 : 15840) }, margins: { top: readAttr(marginAttrs, "top", 1440), right: readAttr(marginAttrs, "right", 1440), bottom: readAttr(marginAttrs, "bottom", 1440), left: readAttr(marginAttrs, "left", 1440) }, styleId }, commentIds };
  }
  const imageRelId = /<a:blip[^>]*r:embed="([^"]+)"/.exec(part)?.[1];
  if (imageRelId) {
    const image = imageByRelId.get(imageRelId) || {};
    const docPr = /<wp:docPr\b([^>]*)\/>/.exec(part)?.[1] || /<pic:cNvPr\b([^>]*)\/>/.exec(part)?.[1] || "";
    const extent = /<wp:extent[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(part);
    const name = decodeXml(/\bname="([^"]*)"/.exec(docPr)?.[1] || image.name || "");
    const alt = decodeXml(/\bdescr="([^"]*)"/.exec(docPr)?.[1] || image.alt || name || "image");
    return { block: { kind: "image", name, alt, dataUrl: image.dataUrl, uri: image.uri, widthPx: extent ? Math.round(Number(extent[1]) / 9525) : image.widthPx, heightPx: extent ? Math.round(Number(extent[2]) / 9525) : image.heightPx, styleId }, commentIds };
  }
  const changeMatch = /<w:(ins|del)\b([^>]*)>([\s\S]*?)<\/w:\1>/.exec(part);
  if (changeMatch) {
    const changeType = changeMatch[1] === "del" ? "delete" : "insert";
    const attrs = changeMatch[2] || "";
    const inner = changeMatch[3] || "";
    const text = decodeXml([...inner.matchAll(/<w:(?:t|delText)[^>]*>([\s\S]*?)<\/w:(?:t|delText)>/g)].map((t) => t[1]).join(""));
    const author = decodeXml(/w:author="([^"]*)"/.exec(attrs)?.[1] || "User");
    const date = decodeXml(/w:date="([^"]*)"/.exec(attrs)?.[1] || "");
    return { block: { kind: "change", changeType, text, styleId, author, date }, commentIds };
  }
  const text = decodeXml([...part.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join(""));
  const numId = /<w:numId[^>]*w:val="(\d+)"/.exec(part)?.[1];
  const level = Number(/<w:ilvl[^>]*w:val="(\d+)"/.exec(part)?.[1] || 0);
  if (numId) return { block: { kind: "listItem", text, styleId, level, listType: numId === "2" ? "number" : "bullet" }, commentIds };
  return { block: { kind: "paragraph", text, styleId }, commentIds };
}

function parseDocxTable(part) {
  const values = [...part.matchAll(/<w:tr[\s\S]*?<\/w:tr>/g)].map((rowMatch) => [...rowMatch[0].matchAll(/<w:tc[\s\S]*?<\/w:tc>/g)].map((cellMatch) => decodeXml([...cellMatch[0].matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join(""))));
  return { kind: "table", values, rows: values.length, columns: Math.max(0, ...values.map((row) => row.length)) };
}

function parseHeaderFooterXml(xml) {
  return [...String(xml || "").matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((match) => ({ text: decodeXml([...match[0].matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join("")) })).filter((item) => item.text.length > 0);
}

export class DocumentFile {
  static async exportDocx(document) {
    const zip = new JSZip();
    const imageParts = collectDocxImageParts(document);
    const hasNumbering = document.blocks.some((block) => block.kind === "listItem");
    const hasHeader = document.headers.length > 0;
    const hasFooter = document.footers.length > 0;
    zip.file("[Content_Types].xml", docxContentTypes({ hasComments: document.comments.length > 0, hasHeader, hasFooter, hasNumbering, imageParts }));
    zip.file("_rels/.rels", relsXml([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "word/document.xml" }]));
    const docRels = [{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", target: "styles.xml" }];
    const relIds = {};
    if (hasNumbering) docRels.push({ id: `rId${docRels.length + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", target: "numbering.xml" });
    if (document.comments.length) docRels.push({ id: `rId${docRels.length + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", target: "comments.xml" });
    if (hasHeader) { relIds.header = `rId${docRels.length + 1}`; docRels.push({ id: relIds.header, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", target: "header1.xml" }); }
    if (hasFooter) { relIds.footer = `rId${docRels.length + 1}`; docRels.push({ id: relIds.footer, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", target: "footer1.xml" }); }
    relIds.hyperlinks = new Map();
    for (const block of document.blocks.filter((item) => item.kind === "hyperlink")) {
      const relId = `rId${docRels.length + 1}`;
      relIds.hyperlinks.set(block.id, relId);
      docRels.push({ id: relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", target: block.url, targetMode: "External" });
    }
    relIds.images = new Map();
    for (const part of imageParts) {
      relIds.images.set(part.image.id, part.relId);
      docRels.push({ id: part.relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: `media/image${part.imagePartId}.${part.extension}` });
    }
    zip.file("word/_rels/document.xml.rels", relsXml(docRels));
    zip.file("word/styles.xml", docxStylesXml(document));
    if (hasNumbering) zip.file("word/numbering.xml", docxNumberingXml());
    if (document.comments.length) zip.file("word/comments.xml", docxCommentsXml(document));
    if (hasHeader) zip.file("word/header1.xml", docxHeaderFooterXml("header", document.headers));
    if (hasFooter) zip.file("word/footer1.xml", docxHeaderFooterXml("footer", document.footers));
    imageParts.forEach((part) => zip.file(`word/media/image${part.imagePartId}.${part.extension}`, part.bytes));
    zip.file("word/open-office-artifact.json", JSON.stringify(document.toProto(), null, 2));
    zip.file("word/document.xml", docxDocumentXml(document, relIds));
    return new FileBlob(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: DOCX_MIME });
  }

  static async importDocx(blobOrBuffer) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const metadataText = await zip.file("word/open-office-artifact.json")?.async("text");
    if (metadataText) return DocumentModel.create(JSON.parse(metadataText));
    const xml = await zip.file("word/document.xml")?.async("text");
    const commentsXml = await zip.file("word/comments.xml")?.async("text");
    const commentTextById = new Map([...String(commentsXml || "").matchAll(/<w:comment[^>]*w:id="(\d+)"[^>]*>([\s\S]*?)<\/w:comment>/g)].map((match) => [match[1], decodeXml([...match[2].matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join(""))]));
    const relsText = await zip.file("word/_rels/document.xml.rels")?.async("text");
    const imageByRelId = new Map();
    for (const rel of parseRelsXml(relsText).filter((item) => item.type.endsWith("/image"))) {
      const target = rel.target.replace(/^\//, "");
      const packagePath = target.startsWith("word/") ? target : path.posix.normalize(`word/${target}`).replace(/^\.\//, "");
      const bytes = await zip.file(packagePath)?.async("uint8array");
      const extension = /\.([A-Za-z0-9+]+)$/.exec(target)?.[1] || "bin";
      imageByRelId.set(rel.id, bytes ? { dataUrl: `data:${imageContentTypeFromExtension(extension)};base64,${Buffer.from(bytes).toString("base64")}` } : { uri: rel.target });
    }
    const blocks = [];
    const pendingComments = [];
    for (const match of String(xml || "").matchAll(/<w:tbl[\s\S]*?<\/w:tbl>|<w:p[\s\S]*?<\/w:p>/g)) {
      const part = match[0];
      if (part.startsWith("<w:tbl")) blocks.push(parseDocxTable(part));
      else {
        const parsed = parseDocxParagraph(part, commentTextById, imageByRelId);
        blocks.push(parsed.block);
        for (const commentId of parsed.commentIds) pendingComments.push({ blockIndex: blocks.length - 1, text: commentTextById.get(commentId) || "" });
      }
    }
    const document = DocumentModel.create({ blocks: blocks.length ? blocks : [{ kind: "paragraph", text: "" }] });
    for (const header of parseHeaderFooterXml(await zip.file("word/header1.xml")?.async("text"))) document.addHeader(header.text, header);
    for (const footer of parseHeaderFooterXml(await zip.file("word/footer1.xml")?.async("text"))) document.addFooter(footer.text, footer);
    pendingComments.forEach((comment) => document.addComment(document.blocks[comment.blockIndex]?.id, comment.text));
    return document;
  }
}

class PdfTable {
  constructor(page, config = {}) {
    this.page = page;
    this.id = config.id || aid("ptb");
    this.name = config.name || "";
    this.values = (config.values || [[]]).map((row) => [...row]);
    this.bbox = config.bbox || [72, 140, 468, Math.max(24, this.values.length * 24)];
  }

  inspectRecord(pageIndex) { return { kind: "table", id: this.id, page: pageIndex + 1, name: this.name || undefined, rows: this.values.length, cols: Math.max(0, ...this.values.map((row) => row.length)), bbox: this.bbox, values: this.values }; }
  toJSON() { return { id: this.id, name: this.name, values: this.values, bbox: this.bbox }; }
}

class PdfImage {
  constructor(page, config = {}) {
    this.page = page;
    this.id = config.id || aid("pim");
    this.name = config.name || "";
    this.dataUrl = config.dataUrl;
    this.uri = config.uri;
    this.prompt = config.prompt;
    this.alt = config.alt || config.altText || config.name || "image";
    this.bbox = config.bbox || [72, 280, Number(config.width || 180), Number(config.height || 120)];
    this.fit = config.fit || "contain";
  }

  inspectRecord(pageIndex) { return { kind: "image", id: this.id, page: pageIndex + 1, name: this.name || undefined, alt: this.alt, uri: this.uri, prompt: this.prompt, bbox: this.bbox, fit: this.fit, hasDataUrl: Boolean(this.dataUrl) }; }
  toJSON() { return { id: this.id, name: this.name, dataUrl: this.dataUrl, uri: this.uri, prompt: this.prompt, alt: this.alt, bbox: this.bbox, fit: this.fit }; }
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
    this.textItems = (config.textItems || []).map((item, index) => ({ id: item.id || `${this.id}/txt/${index + 1}`, text: String(item.text ?? item.str ?? ""), bbox: item.bbox || [Number(item.x || 0), Number(item.y || item.top || 0), Number(item.width || 0), Number(item.height || 0)], fontName: item.fontName, dir: item.dir }));
    this.regions = (config.regions || []).map((region, index) => ({ id: region.id || `${this.id}/rg/${index + 1}`, kind: region.kind || "region", bbox: region.bbox || [0, 0, this.width, this.height], label: region.label }));
  }

  addTable(config = {}) { const table = new PdfTable(this, config); this.tables.push(table); return table; }
  addImage(config = {}) { const image = new PdfImage(this, config); this.images.push(image); return image; }
  inspectRecord(index) { return { kind: "page", id: this.id, page: index + 1, width: this.width, height: this.height, textPreview: this.text.slice(0, 300), textChars: this.text.length, textItems: this.textItems.length, tables: this.tables.length, images: this.images.length, regions: this.regions.length }; }
  textRecord(index) { return { kind: "text", id: `${this.id}/text`, page: index + 1, text: this.text, textChars: this.text.length, textItems: this.textItems.length }; }
  textItemRecords(index) { return this.textItems.map((item) => ({ kind: "textItem", page: index + 1, ...item })); }
  regionRecords(index) { return this.regions.map((region) => ({ ...region, kind: "region", regionKind: region.kind || "region", page: index + 1 })); }
  toJSON() { return { id: this.id, text: this.text, width: this.width, height: this.height, textItems: this.textItems, regions: this.regions, tables: this.tables.map((table) => table.toJSON()), images: this.images.map((image) => image.toJSON()) }; }
}

export class PdfArtifact {
  constructor(options = {}) {
    this.id = options.id || aid("pdf");
    this.metadata = options.metadata || {};
    const pages = options.pages || [{ text: options.text || "", tables: options.tables || [], images: options.images || [] }];
    this.pages = pages.map((page) => new PdfPage(this, page));
  }

  static create(options = {}) { return new PdfArtifact(options); }
  addPage(config = {}) { const page = new PdfPage(this, config); this.pages.push(page); return page; }
  addTable(config = {}) { return (this.pages[0] || this.addPage()).addTable(config); }
  addImage(config = {}) { const pageIndex = Number(config.pageIndex ?? config.page ?? 0); return (this.pages[pageIndex] || this.pages[0] || this.addPage()).addImage(config); }
  extractText(options = {}) { const pages = options.page == null ? this.pages : [this.pages[Number(options.page) - 1]].filter(Boolean); return pages.map((page) => page.text).join("\n\n"); }
  extractTables(options = {}) { const pages = options.page == null ? this.pages : [this.pages[Number(options.page) - 1]].filter(Boolean); return pages.flatMap((page, index) => page.tables.map((table) => ({ page: options.page || index + 1, id: table.id, name: table.name, values: table.values, bbox: table.bbox }))); }
  resolve(id) {
    if (id === this.id) return this;
    for (const [pageIndex, page] of this.pages.entries()) {
      if (id === page.id) return page;
      if (id === `${page.id}/text`) return { kind: "text", id, page: pageIndex + 1, text: page.text, textChars: page.text.length, pageObject: page };
      const textItem = page.textItems.find((item) => item.id === id);
      if (textItem) return textItem;
      const region = page.regions.find((item) => item.id === id);
      if (region) return region;
      const table = page.tables.find((item) => item.id === id);
      if (table) return table;
      const image = page.images.find((item) => item.id === id);
      if (image) return image;
    }
    return undefined;
  }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["page", "text", "table", "image"]);
    const records = [];
    this.pages.forEach((page, index) => {
      if (kinds.has("page")) records.push(page.inspectRecord(index));
      if (kinds.has("text")) records.push(page.textRecord(index));
      if (kinds.has("textItem")) records.push(...page.textItemRecords(index));
      if (kinds.has("region")) records.push(...page.regionRecords(index));
      if (kinds.has("table")) records.push(...page.tables.map((table) => table.inspectRecord(index)));
      if (kinds.has("image")) records.push(...page.images.map((image) => image.inspectRecord(index)));
    });
    return ndjson(filterInspectRecords(records, options), options.maxChars ?? Infinity);
  }

  verify(options = {}) {
    const issues = [];
    if (this.pages.length === 0) issues.push(verificationIssue("pdf", "noPages", "PDF artifact has no pages."));
    this.pages.forEach((page, pageIndex) => {
      if (!page.text.trim() && page.tables.length === 0 && page.images.length === 0) issues.push(verificationIssue("pdf", "emptyPage", `PDF page ${pageIndex + 1} has no modeled text, tables, or images.`, { page: pageIndex + 1 }));
      if (page.textItems.length && !page.text.trim()) issues.push(verificationIssue("pdf", "textExtractionMismatch", `PDF page ${pageIndex + 1} has positioned text items but no extracted page text.`, { severity: "warning", page: pageIndex + 1 }));
      if (page.textItems.some((item) => !item.text)) issues.push(verificationIssue("pdf", "emptyTextItem", `PDF page ${pageIndex + 1} contains an empty positioned text item.`, { severity: "warning", page: pageIndex + 1 }));
      if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(page.text)) issues.push(verificationIssue("pdf", "textExtractionControlChars", `PDF page ${pageIndex + 1} extracted text contains control characters.`, { page: pageIndex + 1 }));
      if (page.width <= 0 || page.height <= 0) issues.push(verificationIssue("pdf", "invalidPageGeometry", `PDF page ${pageIndex + 1} has invalid page geometry.`, { page: pageIndex + 1, width: page.width, height: page.height }));
      if (/[\u2010-\u2015]/.test(page.text)) issues.push(verificationIssue("pdf", "unicodeDash", `PDF page ${pageIndex + 1} contains a Unicode dash; use ASCII hyphen for compatibility.`, { page: pageIndex + 1 }));
      for (const table of page.tables) {
        if (!table.values.length || !table.values[0]?.length) issues.push(verificationIssue("pdf", "emptyTable", `PDF table ${table.id} on page ${pageIndex + 1} has no cells.`, { page: pageIndex + 1, id: table.id }));
        const [left, top, width, height] = table.bbox || [];
        if (left < 0 || top < 0 || width <= 0 || height <= 0 || left + width > page.width || top + height > page.height) {
          issues.push(verificationIssue("pdf", "tableOutOfBounds", `PDF table ${table.id} extends outside page ${pageIndex + 1}.`, { page: pageIndex + 1, id: table.id, bbox: table.bbox }));
        }
      }
      for (const image of page.images) {
        if (!image.dataUrl && !image.uri && !image.prompt) issues.push(verificationIssue("pdf", "emptyImage", `PDF image ${image.id} on page ${pageIndex + 1} has no dataUrl, uri, or prompt.`, { page: pageIndex + 1, id: image.id }));
        const [left, top, width, height] = image.bbox || [];
        if (left < 0 || top < 0 || width <= 0 || height <= 0 || left + width > page.width || top + height > page.height) {
          issues.push(verificationIssue("pdf", "imageOutOfBounds", `PDF image ${image.id} extends outside page ${pageIndex + 1}.`, { page: pageIndex + 1, id: image.id, bbox: image.bbox }));
        }
      }
    });
    return verificationResult("pdf", issues, options);
  }

  help(query = "*", options = {}) { return helpArtifact("pdf", query, options); }

  layoutJson(options = {}) {
    const pageNumber = options.page != null ? Number(options.page) : options.pageIndex != null ? Number(options.pageIndex) + 1 : undefined;
    const selectedPages = pageNumber ? [this.pages[pageNumber - 1]].filter(Boolean) : this.pages;
    return {
      kind: "pdfLayout",
      id: this.id,
      pageCount: this.pages.length,
      metadata: this.metadata,
      pages: selectedPages.map((page) => {
        const pageIndex = this.pages.indexOf(page);
        return {
          kind: "pdfPageLayout",
          id: page.id,
          page: pageIndex + 1,
          width: page.width,
          height: page.height,
          unit: "pt",
          text: { id: `${page.id}/text`, text: page.text, textChars: page.text.length, bbox: [0, 0, page.width, page.height] },
          textItems: page.textItems.map((item) => ({ kind: "textItem", ...item })),
          regions: page.regions.map((region) => ({ kind: "region", ...region })),
          tables: page.tables.map((table) => ({ kind: "table", id: table.id, name: table.name || undefined, values: table.values, bbox: table.bbox })),
          images: page.images.map((image) => ({ kind: "image", id: image.id, name: image.name || undefined, alt: image.alt, bbox: image.bbox, fit: image.fit, hasDataUrl: Boolean(image.dataUrl), uri: image.uri, prompt: image.prompt })),
        };
      }),
    };
  }

  async render(options = {}) {
    const format = String(options.format || "").trim().toLowerCase();
    if (format === "layout" || format === LAYOUT_MIME) return new FileBlob(JSON.stringify(this.layoutJson(options), null, 2), { type: LAYOUT_MIME, metadata: { artifactKind: "pdf", format: "layout", page: options.page, pageIndex: options.pageIndex } });
    return new FileBlob(pdfPageSvg(this.pages[options.pageIndex || 0] || new PdfPage(this)), { type: "image/svg+xml" });
  }
  toJSON() { return { id: this.id, metadata: this.metadata, pages: this.pages.map((page) => page.toJSON()) }; }
}

export class PdfFile {
  static async exportPdf(artifact) {
    return new FileBlob(buildMinimalPdf(artifact), { type: PDF_MIME });
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
    tables: (page.tables || []).map((table, tableIndex) => ({ name: table.name || `parsed-table-${index + 1}-${tableIndex + 1}`, values: table.values || table.rows || [], bbox: table.bbox || table.bounds || [72, 140 + tableIndex * 120, 468, 96] })),
    images: (page.images || []).map((image, imageIndex) => ({ name: image.name || `parsed-image-${index + 1}-${imageIndex + 1}`, alt: image.alt || image.altText || image.name || "parsed image", dataUrl: image.dataUrl, uri: image.uri, prompt: image.prompt, bbox: image.bbox || image.bounds || [72, 280 + imageIndex * 140, 180, 120] })),
  }));
  return PdfArtifact.create({ metadata: { ...metadata, ...(source.metadata || parsed?.metadata || {}) }, pages: pages.length ? pages : [{ text: "" }] });
}

function pdfPageSvg(page) {
  const width = page.width || 612;
  const height = page.height || 792;
  const lines = String(page.text || "").split(/\r?\n/).filter(Boolean);
  const text = lines.map((line, index) => `<text x="72" y="${96 + index * 20}" font-family="Helvetica" font-size="14" fill="#111827">${xmlEscape(line)}</text>`).join("");
  const tables = (page.tables || []).map((table) => {
    const [left, top, tableWidth, tableHeight] = table.bbox;
    const rows = table.values.length;
    const cols = Math.max(1, ...table.values.map((row) => row.length));
    const cellW = tableWidth / cols;
    const cellH = tableHeight / Math.max(1, rows);
    const cells = [];
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const x = left + c * cellW;
        const y = top + r * cellH;
        cells.push(`<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" fill="${r === 0 ? "#f1f5f9" : "#ffffff"}" stroke="#cbd5e1"/>`);
        cells.push(`<text x="${x + 5}" y="${y + Math.min(16, cellH - 4)}" font-family="Helvetica" font-size="11" fill="#111827">${xmlEscape(table.values[r]?.[c] ?? "")}</text>`);
      }
    }
    return cells.join("");
  }).join("");
  const images = (page.images || []).map((image) => {
    const [left, top, imageWidth, imageHeight] = image.bbox || [72, 280, 180, 120];
    const visual = image.dataUrl ? `<image href="${attrEscape(image.dataUrl)}" x="${left}" y="${top}" width="${imageWidth}" height="${imageHeight}" preserveAspectRatio="xMidYMid meet"/>` : `<rect x="${left}" y="${top}" width="${imageWidth}" height="${imageHeight}" fill="#fef3c7" stroke="#f59e0b"/>`;
    return `${visual}<text x="${left + 8}" y="${top + 18}" font-family="Helvetica" font-size="11" fill="#92400e">${xmlEscape(image.alt || image.prompt || image.uri || image.name || "image")}</text>`;
  }).join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="white"/>${text}${tables}${images}</svg>`;
}

function escapePdfString(text) {
  return String(text).replaceAll("\\", "\\\\").replaceAll("(", "\\(").replaceAll(")", "\\)");
}

function buildMinimalPdf(artifact) {
  const page = artifact.pages[0] || new PdfPage(artifact);
  const tableLines = page.tables.flatMap((table) => table.values.map((row) => row.join(" | ")));
  const imageLines = page.images.map((image) => `[Image: ${image.alt || image.prompt || image.uri || image.name || "image"}]`);
  const lines = [...String(page.text || "").split(/\r?\n/), ...tableLines, ...imageLines].filter(Boolean);
  const content = `BT\n/F1 18 Tf\n72 720 Td\n${lines.map((line, index) => `${index === 0 ? "" : "0 -24 Td\n"}(${escapePdfString(line)}) Tj`).join("\n")}\nET`;
  const metadata = Buffer.from(JSON.stringify(artifact.toJSON()), "utf8").toString("base64");
  const objects = [
    `1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n`,
    `2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n`,
    `3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${page.width || 612} ${page.height || 792}] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>\nendobj\n`,
    `4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n`,
    `5 0 obj\n<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}\nendstream\nendobj\n`,
  ];
  let pdf = `%PDF-1.4\n%OPEN_OFFICE_ARTIFACT ${metadata}\n`;
  const offsets = [0];
  for (const obj of objects) { offsets.push(Buffer.byteLength(pdf)); pdf += obj; }
  const xref = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (let i = 1; i < offsets.length; i++) pdf += `${String(offsets[i]).padStart(10, "0")} 00000 n \n`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return encoder.encode(pdf);
}
