import fs from "node:fs/promises";
import path from "node:path";
import { deflateSync, inflateSync } from "node:zlib";
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
  const result = compareRgbaPixels(actual, expected, { ...options, threshold, format: "png" });
  return result;
}

function isPpmBytes(bytes) {
  return bytes?.[0] === 0x50 && (bytes?.[1] === 0x36 || bytes?.[1] === 0x33);
}

function ppmTokens(bytes) {
  const tokens = [];
  let i = 0;
  while (i < bytes.length) {
    while (i < bytes.length && /\s/.test(String.fromCharCode(bytes[i]))) i += 1;
    if (bytes[i] === 0x23) { while (i < bytes.length && bytes[i] !== 0x0a) i += 1; continue; }
    if (i >= bytes.length) break;
    const start = i;
    while (i < bytes.length && !/\s/.test(String.fromCharCode(bytes[i])) && bytes[i] !== 0x23) i += 1;
    tokens.push({ text: decoder.decode(bytes.slice(start, i)), end: i });
    if (tokens.length >= 4) break;
  }
  return tokens;
}

function decodePpmRgba(bytes) {
  if (!isPpmBytes(bytes)) throw new Error("not a PPM file");
  const tokens = ppmTokens(bytes);
  const magic = tokens[0]?.text;
  const width = Number(tokens[1]?.text);
  const height = Number(tokens[2]?.text);
  const max = Number(tokens[3]?.text);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0 || max <= 0) throw new Error("PPM is missing valid geometry or max value");
  const rgba = new Uint8Array(width * height * 4);
  if (magic === "P6") {
    let offset = tokens[3].end;
    while (offset < bytes.length && /\s/.test(String.fromCharCode(bytes[offset]))) offset += 1;
    const needed = width * height * 3;
    if (bytes.length - offset < needed) throw new Error("PPM image data is truncated");
    for (let p = 0; p < width * height; p += 1) {
      const input = offset + p * 3;
      const output = p * 4;
      rgba[output] = Math.round(bytes[input] * 255 / max);
      rgba[output + 1] = Math.round(bytes[input + 1] * 255 / max);
      rgba[output + 2] = Math.round(bytes[input + 2] * 255 / max);
      rgba[output + 3] = 255;
    }
  } else if (magic === "P3") {
    const text = decoder.decode(bytes.slice(tokens[3].end));
    const values = text.replace(/#[^\n\r]*/g, " ").trim().split(/\s+/).filter(Boolean).map(Number);
    if (values.length < width * height * 3) throw new Error("PPM image data is truncated");
    for (let p = 0; p < width * height; p += 1) {
      const input = p * 3;
      const output = p * 4;
      rgba[output] = Math.round(values[input] * 255 / max);
      rgba[output + 1] = Math.round(values[input + 1] * 255 / max);
      rgba[output + 2] = Math.round(values[input + 2] * 255 / max);
      rgba[output + 3] = 255;
    }
  } else {
    throw new Error(`unsupported PPM magic ${magic}`);
  }
  return { width, height, pixels: rgba };
}

function comparePpmPixels(bytes, baselineBytes, options = {}) {
  const threshold = Math.max(0, Number(options.pixelThreshold ?? options.threshold ?? 0));
  return compareRgbaPixels(decodePpmRgba(bytes), decodePpmRgba(baselineBytes), { ...options, threshold, format: "ppm" });
}

function compareRgbaPixels(actual, expected, options = {}) {
  const threshold = Math.max(0, Number(options.threshold ?? options.pixelThreshold ?? 0));
  const result = {
    format: options.format || "rgba",
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
      if ((isPngBytes(bytes) && isPngBytes(baselineBytes)) || (isPpmBytes(bytes) && isPpmBytes(baselineBytes))) {
        try {
          const pixelDiffOptions = typeof options.pixelDiff === "object" ? options.pixelDiff : options;
          const pixelDiff = isPngBytes(bytes) ? comparePngPixels(bytes, baselineBytes, pixelDiffOptions) : comparePpmPixels(bytes, baselineBytes, pixelDiffOptions);
          summary.pixelDiff = pixelDiff;
          if (pixelDiff.changed && options.allowChange !== true && options.allowPixelChange !== true) {
            issues.push(verificationIssue(artifactKind, "visualPixelDiff", `Rendered ${pixelDiff.format.toUpperCase()} differs from the baseline in ${pixelDiff.differentPixels} pixels.`, { severity: options.diffSeverity || "warning", ...pixelDiff }));
          }
        } catch (error) {
          summary.pixelDiff = { skipped: true, reason: error.message };
        }
      } else {
        summary.pixelDiff = { skipped: true, reason: "pixelDiff currently supports PNG and PPM baselines only" };
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
  { artifactKind: "workbook", kind: "api", name: "workbook.layoutJson", summary: "Return workbook/worksheet layout JSON with cell, table, chart, image, sparkline, rule bounding boxes, and target/search context slicing." },
  { artifactKind: "workbook", kind: "api", name: "workbook.verify", summary: "Return bounded QA issues for sheets, formulas, tables, charts, and comments." },
  { artifactKind: "workbook", kind: "api", name: "workbook.trace", summary: "Return a formula precedent tree and bounded NDJSON trace for a target cell, with circular references flagged." },
  { artifactKind: "workbook", kind: "api", name: "workbook.formulaGraph", summary: "Return a dependency graph of formula nodes, edges, dependents, cycles, and formula errors for workbook QA." },
  { artifactKind: "workbook", kind: "formula", name: "workbook.structuredReferences", summary: "Evaluate Excel-style table structured references such as TableName[Column], TableName[#Headers], TableName[[#Data],[Column]], and TableName[[#Data],[First]:[Last]] in formulas, expanding them to stable table cell precedents." },
  { artifactKind: "workbook", kind: "formula", name: "workbook.sharedArrayFormulas", summary: "Import and export native XLSX shared formulas (t=shared) by translating relative A1 references and surface native array formulas (t=array) with formulaType/sharedRef/arrayRef inspect metadata." },
  { artifactKind: "workbook", kind: "api", name: "workbook.definedNames.add", summary: "Create a workbook or sheet-scoped defined name over an A1 range; exported as native workbook.xml definedName and usable in formulas such as SUM(RevenueData)." },
  { artifactKind: "workbook", kind: "api", name: "range.dataValidation", summary: "Assign a validation rule to a range or use sheet.dataValidations.add({ range, rule })." },
  { artifactKind: "workbook", kind: "api", name: "range.format", summary: "Assign basic cell style metadata such as fill, font, numberFormat, alignment, and borders; XLSX export writes native styles.xml and cell style indexes." },
  { artifactKind: "workbook", kind: "api", name: "range.conditionalFormats.add", summary: "Add a conditional formatting rule; cellIs/expression/containsText/colorScale rules are evaluated into computedStyle inspect records, layout JSON hints, and SVG preview fills." },
  { artifactKind: "workbook", kind: "api", name: "workbook.comments.addThread", summary: "Create threaded comments after comments.setSelf({ displayName }); resolve with wb.resolve('th/...')." },
  { artifactKind: "workbook", kind: "api", name: "sheet.tables.add", summary: "Create an inspectable worksheet table over an A1 range with rows.add, getDataRows, getHeaderRowRange, style, and visibility toggles." },
  { artifactKind: "workbook", kind: "api", name: "sheet.pivotTables.add", summary: "Create a clean-room pivot table facade over a source range with row/value fields, computed summary values, inspect/resolve/layout records, verification, and metadata roundtrip." },
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
  { artifactKind: "workbook", kind: "formula", name: "fx.COUNTIFS", category: "statistical", summary: "Count rows where multiple criteria ranges all match their criteria.", examples: ["=COUNTIFS(A1:A10,\"East\",B1:B10,\">=10\")"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.AVERAGEIF", category: "statistical", summary: "Average values whose corresponding criteria range entries match a criterion.", examples: ["=AVERAGEIF(A1:A10,\"East\",B1:B10)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.AVERAGEIFS", category: "statistical", summary: "Average values where all supplied criteria ranges match their criteria.", examples: ["=AVERAGEIFS(C1:C10,A1:A10,\"East\",B1:B10,\">=10\")"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.SUMIF", category: "math-trig", summary: "Sum values whose corresponding criteria range entries match a criterion.", examples: ["=SUMIF(A1:A10,\"East\",B1:B10)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.SUMIFS", category: "math-trig", summary: "Sum values where all supplied criteria ranges match their criteria.", examples: ["=SUMIFS(C1:C10,A1:A10,\"East\",B1:B10,\">=10\")"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.SUMPRODUCT", category: "math-trig", summary: "Multiply corresponding numeric values in equally sized arrays and return the sum of those products.", examples: ["=SUMPRODUCT(A1:A10,B1:B10)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.VLOOKUP", category: "lookup-reference", summary: "Look up a value in the first column of a table range and return a value from another column.", examples: ["=VLOOKUP(\"Beta\",A2:B4,2,FALSE)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.HLOOKUP", category: "lookup-reference", summary: "Look up a value in the first row of a table range and return a value from another row.", examples: ["=HLOOKUP(\"Revenue\",A1:D4,3,FALSE)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.XLOOKUP", category: "lookup-reference", summary: "Look up a value in one range and return the corresponding value from another range.", examples: ["=XLOOKUP(\"Gamma\",A2:A4,B2:B4,\"missing\")"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.INDEX", category: "lookup-reference", summary: "Return a value from a range by 1-based row and optional column index.", examples: ["=INDEX(A2:C4,2,3)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.MATCH", category: "lookup-reference", summary: "Return the 1-based position of a lookup value in a range, with exact match and basic ascending/descending approximate modes.", examples: ["=MATCH(\"Beta\",A2:A4,0)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.SEQUENCE", category: "dynamic-array", summary: "Return a dynamic array sequence that spills into neighboring cells in the clean-room formula engine.", examples: ["=SEQUENCE(2,3,10,2)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.TRANSPOSE", category: "dynamic-array", summary: "Transpose a source range into a spilled dynamic array with spillRange/spillValues inspect metadata.", examples: ["=TRANSPOSE(A1:C2)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.FILTER", category: "dynamic-array", summary: "Filter rows from a source range with a boolean or comparison include array and spill the matching rows.", examples: ["=FILTER(A2:C10,B2:B10=\"East\")"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.UNIQUE", category: "dynamic-array", summary: "Return unique rows from a range as a spilled dynamic array.", examples: ["=UNIQUE(A2:A10)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.SORT", category: "dynamic-array", summary: "Sort a range by a 1-based column index and spill the sorted rows.", examples: ["=SORT(A2:C10,3,-1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.TEXTJOIN", category: "text", summary: "Join text values with a delimiter and optional empty-value skipping.", examples: ["=TEXTJOIN(\"/\",TRUE,A1:A3)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.CONCAT", category: "text", summary: "Concatenate text values and ranges.", examples: ["=CONCAT(A1,\"-\",B1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.LEFT", category: "text", summary: "Return characters from the start of a text value.", examples: ["=LEFT(A1,3)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.RIGHT", category: "text", summary: "Return characters from the end of a text value.", examples: ["=RIGHT(A1,3)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.LEN", category: "text", summary: "Return the length of a text value.", examples: ["=LEN(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.MID", category: "text", summary: "Return characters from the middle of a text value.", examples: ["=MID(A1,2,3)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.UPPER", category: "text", summary: "Convert text to uppercase.", examples: ["=UPPER(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.LOWER", category: "text", summary: "Convert text to lowercase.", examples: ["=LOWER(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.TRIM", category: "text", summary: "Trim leading/trailing whitespace and collapse internal whitespace.", examples: ["=TRIM(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.IFERROR", category: "logical", summary: "Return a fallback value when an expression evaluates to a formula error.", examples: ["=IFERROR(XLOOKUP(\"missing\",A1:A10,B1:B10),\"not found\")"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.ISNUMBER", category: "information", summary: "Return TRUE when a value is numeric.", examples: ["=ISNUMBER(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.ISTEXT", category: "information", summary: "Return TRUE when a value is text and not a formula error.", examples: ["=ISTEXT(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.ISBLANK", category: "information", summary: "Return TRUE when a referenced value is empty.", examples: ["=ISBLANK(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.ISERROR", category: "information", summary: "Return TRUE when a value is any recognized formula error.", examples: ["=ISERROR(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.ABS", category: "math-trig", summary: "Return the absolute value of a number.", examples: ["=ABS(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.INT", category: "math-trig", summary: "Round a number down to the nearest integer.", examples: ["=INT(A1)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.CEILING", category: "math-trig", summary: "Round a number up to the nearest significance.", examples: ["=CEILING(A1,5)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.FLOOR", category: "math-trig", summary: "Round a number down to the nearest significance.", examples: ["=FLOOR(A1,5)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.PMT", category: "financial", summary: "Calculate a loan payment for constant payments and constant interest rate.", examples: ["=PMT(rate,nper,pv)"], notes: ["Catalog entry only in MVP; full financial formula evaluation is roadmap."] },

  { artifactKind: "presentation", kind: "api", name: "Presentation.create", summary: "Create a deck with a default or explicit slide size." },
  { artifactKind: "presentation", kind: "api", name: "presentation.inspect", summary: "Emit NDJSON for deck, slides, textboxes, shapes, tables, charts, images, notes, comments, and layout; narrow with search/target anchors and shape fields with include/exclude." },
  { artifactKind: "presentation", kind: "api", name: "presentation.textRange", summary: "Inspect or resolve stable textRange anchors such as shapeId/text for editable slide text frames." },
  { artifactKind: "presentation", kind: "api", name: "presentation.resolve", summary: "Map stable inspect anchor IDs back to editable facade objects." },
  { artifactKind: "presentation", kind: "api", name: "presentation.export", summary: "Export a slide SVG preview, deck SVG montage via { format: 'montage' }, or target/search-sliced layout JSON." },
  { artifactKind: "presentation", kind: "api", name: "presentation.validateLayout", summary: "Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow." },
  { artifactKind: "presentation", kind: "api", name: "presentation.verify", summary: "Return presentation QA issues for layout validation, placeholder/template fidelity, chart/data consistency, table shape, image data, and dangling comments." },
  { artifactKind: "presentation", kind: "api", name: "slide.shapes.add", summary: "Add a shape/textbox with geometry, position, fill, line, and text." },
  { artifactKind: "presentation", kind: "api", name: "slide.compose", summary: "Materialize a clean-room compose tree with row, column, grid, layers, box, paragraph, shape, table, chart, image, and rule nodes into editable slide objects." },
  { artifactKind: "presentation", kind: "api", name: "slide.autoLayout", summary: "Place existing shapes inside a frame using horizontal or vertical flow, gap, padding, and alignment options." },
  { artifactKind: "presentation", kind: "api", name: "slide.tables.add", summary: "Add an inspectable native-style table facade with rows, columns, values, cells, layout JSON, and SVG/PPTX placeholder output." },
  { artifactKind: "presentation", kind: "api", name: "slide.charts.add", summary: "Add an inspectable bar/line/pie chart facade with chartType, title, categories, series colors, axes, legend, data labels, layout JSON, SVG preview, and PPTX chart output." },
  { artifactKind: "presentation", kind: "api", name: "slide.images.add", summary: "Add an inspectable image facade with alt text, prompt/URI/data URL metadata, fit, frame, layout JSON, SVG preview, and PPTX placeholder output." },
  { artifactKind: "presentation", kind: "api", name: "presentation.theme", summary: "Configure inspectable theme colors and major/minor fonts; export writes a real ppt/theme/theme1.xml part." },
  { artifactKind: "presentation", kind: "api", name: "presentation.layouts.add", summary: "Create a reusable slide layout with placeholders; export writes slideLayout and slideMaster parts for clean-room PPTX roundtrip." },
  { artifactKind: "presentation", kind: "api", name: "slide.applyLayout", summary: "Apply a slide layout to materialize editable placeholder shapes and preserve layout identity for inspect, verify, and PPTX export." },
  { artifactKind: "presentation", kind: "api", name: "slide.addNotes", summary: "Set speaker notes for a slide; exported as a PPTX notesSlide part and surfaced through inspect({ kind: 'notes' })." },
  { artifactKind: "presentation", kind: "api", name: "slide.comments.addThread", summary: "Attach threaded comments to slide elements; exported as PPTX comments parts and verified for dangling targets." },
  { artifactKind: "presentation", kind: "api", name: "slide.connectors.add", summary: "Add an inspectable connector line between points or element IDs with SVG preview, layout JSON, PPTX p:cxnSp export, and off-canvas QA." },
  { artifactKind: "presentation", kind: "api", name: "PresentationFile.inspectPptx", summary: "Inspect a PPTX zip package as bounded NDJSON part records with paths, sizes, content types, and optional XML/relationship previews." },
  { artifactKind: "presentation", kind: "api", name: "compose.column", summary: "Create a vertical compose container. Use width/height fill, hug, or fixed pixels; gap and padding are in pixels." },
  { artifactKind: "presentation", kind: "api", name: "compose.paragraph", summary: "Create an editable text block with name, className/style text tokens, and stable inspect output." },

  { artifactKind: "document", kind: "api", name: "DocumentModel.create", summary: "Create a document with paragraph, list, table, header/footer, style, and comment blocks." },
  { artifactKind: "document", kind: "api", name: "document.addParagraph", summary: "Append a styled paragraph block with optional run-level styles and return an inspectable/resolveable paragraph object." },
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
  { artifactKind: "document", kind: "api", name: "document.styles.effective", summary: "Resolve a named document style through basedOn inheritance so inspect/layout/render/DOCX export share the same effective style metadata." },
  { artifactKind: "document", kind: "api", name: "document.inspect", summary: "Emit bounded NDJSON for document blocks, comments, styles, headers/footers, and layout; narrow with search/target anchors and shape fields with include/exclude." },
  { artifactKind: "document", kind: "api", name: "document.textRange", summary: "Inspect or resolve stable textRange anchors such as blockId/text for editable document block, header/footer, and comment text." },
  { artifactKind: "document", kind: "api", name: "document.layoutJson", summary: "Return page-aware layout JSON with block bounding boxes, page records, style IDs, design preset metadata, and target/search context slicing." },
  { artifactKind: "document", kind: "api", name: "document.render", summary: "Render an SVG preview by default, return layout JSON with { format: 'layout' }, or use { source: 'docx', renderer } to feed native DOCX into LibreOffice/native Office render adapters for PDF/PNG outputs." },
  { artifactKind: "document", kind: "api", name: "document.verify", summary: "Return QA issues for fake lists, invalid links/citations, unknown styles, malformed tables, bad image dimensions/data URLs, section setup, dangling comments, visual layout overflow, and prose-like table cells." },
  { artifactKind: "document", kind: "api", name: "DocumentFile.exportDocx", summary: "Export DocumentModel to a DOCX package with document.xml, styles.xml, comments.xml, numbering.xml, header/footer parts, hyperlinks, fields, citations, and metadata." },
  { artifactKind: "document", kind: "api", name: "DocumentFile.inspectDocx", summary: "Inspect a DOCX zip package as bounded NDJSON part records with safe part paths, sizes, content types, and optional XML/JSON previews." },
  { artifactKind: "document", kind: "api", name: "DocumentFile.patchDocx", summary: "Apply safe in-package DOCX XML/JSON/binary patches with path traversal validation and return a patched DOCX FileBlob." },

  { artifactKind: "pdf", kind: "api", name: "PdfArtifact.create", summary: "Create a modeled PDF artifact with pages, text, table regions, and image regions." },
  { artifactKind: "pdf", kind: "api", name: "pdf.addText", summary: "Add positioned PDF text with page-space bbox, font metadata, inspect/resolve/layout records, and SVG preview rendering." },
  { artifactKind: "pdf", kind: "api", name: "pdf.addImage", summary: "Add a modeled PDF image region with dataUrl/URI/prompt metadata, alt text, and page-space bounding box." },
  { artifactKind: "pdf", kind: "api", name: "pdf.addChart", summary: "Add a modeled bar/line chart region with categories, series, title, bbox, inspect/resolve/layout records, SVG preview, and PDF metadata roundtrip." },
  { artifactKind: "pdf", kind: "api", name: "pdf.extractText", summary: "Extract modeled text across all pages or a selected page." },
  { artifactKind: "pdf", kind: "api", name: "pdf.extractTables", summary: "Extract modeled table values and bounding boxes across all pages or a selected page." },
  { artifactKind: "pdf", kind: "api", name: "pdf.inspect", summary: "Emit bounded NDJSON for pages, text, positioned text items, layout regions, tables, images, and charts; narrow with search/target anchors and shape fields with include/exclude." },
  { artifactKind: "pdf", kind: "api", name: "pdf.resolve", summary: "Resolve stable PDF artifact IDs for pages, page text blocks, positioned text items, layout regions, tables, images, and charts." },
  { artifactKind: "pdf", kind: "api", name: "pdf.render", summary: "Render a modeled PDF page to SVG by default, return page layout JSON with { format: 'layout' }, or use { source: 'pdf', renderer } to feed the exported PDF into Poppler/PDF-capable raster adapters." },
  { artifactKind: "pdf", kind: "api", name: "pdf.layoutJson", summary: "Return modeled PDF page layout JSON with page text, positioned text items, layout regions, tables, images, charts, and target/search context slicing." },
  { artifactKind: "pdf", kind: "api", name: "pdf.verify", summary: "Return QA issues for empty pages, Unicode dashes, text extraction sanity, page geometry, text/region/table/image/chart bounds, invalid image data URLs, malformed tables, and chart data." },
  { artifactKind: "pdf", kind: "api", name: "PdfFile.exportPdf", summary: "Export a modeled artifact as a real multi-page PDF with positioned text, vector tables/charts, embedded PNG images, and clean-room metadata." },
  { artifactKind: "pdf", kind: "api", name: "PdfFile.inspectPdf", summary: "Inspect PDF bytes as bounded file/object records including version, byte size, page/object counts, embedded clean-room model presence, and EOF integrity." },
  { artifactKind: "pdf", kind: "api", name: "PdfFile.importPdf", summary: "Import clean-room generated PDFs from metadata, use an injected parser adapter for arbitrary PDFs, normalize parser image bytes/base64 into data URLs, reconstruct tables from positioned text geometry when explicit tables are absent, or fall back to heuristic visible-text/table extraction." },
  { artifactKind: "pdf", kind: "api", name: "createPdfjsParser", summary: "Create an optional PDF.js parser adapter from open-office-artifact-tool/pdf/pdfjs to extract page geometry, positioned text, heuristic tables, and image placeholders." },

  { artifactKind: "shared", kind: "api", name: "verifyArtifact", summary: "Run an artifact's verify() method and return a bounded NDJSON QA report." },
  { artifactKind: "shared", kind: "api", name: "visualQaArtifact", summary: "Render an artifact, record deterministic render metadata/hash, validate empty or malformed render output, optionally compare against a baseline render, and compute PNG/PPM pixel-diff metrics when requested." },
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
    schema: {
      parameters: {
        kind: { type: "string", description: "Comma-separated record kinds such as formula, table, style, computedStyle, chart, image." },
        target: { type: "string", description: "Stable ID, anchor, or A1 cell/range to slice results around." },
        search: { type: "string", description: "Case-insensitive text filter over inspect records." },
        include: { type: "string", description: "Comma-separated top-level fields to keep." },
        exclude: { type: "string", description: "Comma-separated top-level fields to omit." },
        maxChars: { type: "number", description: "Maximum NDJSON output size before truncation notice." },
      },
      returns: {
        ndjson: { type: "string", description: "Bounded newline-delimited JSON records." },
        truncated: { type: "boolean", description: "True when maxChars truncated the output." },
      },
    },
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
    schema: {
      parameters: {
        artifact: { type: "Workbook|Presentation|DocumentModel|PdfArtifact", required: true, description: "Artifact facade to render through its native preview/export path." },
        format: { type: "string", description: "svg, png, webp, jpeg, pdf, layout, or an output MIME type." },
        renderer: { type: "function", description: "Optional pluggable renderer adapter for raster/PDF conversion." },
        source: { type: "string", description: "Optional native source such as docx or pdf for renderer gates." },
      },
      returns: {
        blob: { type: "FileBlob", description: "Rendered output with normalized metadata." },
      },
    },
  },
  visualQaArtifact: {
    examples: ["await visualQaArtifact(document, { baseline, pixelDiff: true, minBytes: 100 })"],
    options: ["baseline/expected/baselineBlob", "pixelDiff", "PNG/PPM raster pixel comparison", "allowChange", "minBytes", "maxBytes", "maxChars"],
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
  "range.format": {
    examples: ["sheet.getRange('A1:D1').format = { fill: '#0f172a', font: { bold: true }, alignment: { horizontal: 'center' }, border: { style: 'thin' } }"],
    schema: {
      parameters: {
        fill: { type: "string", description: "Cell background color token or hex color." },
        font: { type: "object", description: "Font properties: bold, italic, color, size, name." },
        numberFormat: { type: "string", description: "Excel number format code." },
        alignment: { type: "object", description: "horizontal, vertical, and wrapText alignment options." },
        border: { type: "object", description: "Basic border style and color." },
      },
      returns: { range: { type: "Range", description: "The formatted range facade." } },
    },
  },
  "range.conditionalFormats.add": {
    examples: ["range.conditionalFormats.add('cellIs', { operator: 'greaterThan', formula: 10, format: { fill: 'green' } })", "range.conditionalFormats.addColorScale({ colors: ['#fee2e2', '#fef3c7', '#22c55e'] })"],
    schema: {
      parameters: {
        ruleType: { type: "string", required: true, description: "cellIs, expression, containsText, or colorScale." },
        formula: { type: "string|number", description: "Rule formula or scalar threshold." },
        operator: { type: "string", description: "Comparison operator for cellIs rules." },
        format: { type: "object", description: "Style patch applied when the rule matches." },
        colors: { type: "string[]", description: "Two or three colors for colorScale rules." },
      },
      returns: { conditionalFormat: { type: "object", description: "Inspectable conditional-format rule with stable id." } },
    },
  },
  "PdfFile.importPdf": {
    examples: ["await PdfFile.importPdf(blob, { parser: createPdfjsParser() })"],
    schema: {
      parameters: {
        blob: { type: "FileBlob|Uint8Array", required: true, description: "PDF input bytes." },
        parser: { type: "function", description: "Optional parser adapter returning pages/textItems/tables/images." },
        preferParser: { type: "boolean", description: "Use parser even if clean-room metadata is embedded." },
        parserName: { type: "string", description: "Name recorded in artifact metadata." },
      },
      returns: { pdf: { type: "PdfArtifact", description: "Modeled PDF artifact with inspect/resolve/render/verify APIs." } },
    },
  },
  "DocumentFile.patchDocx": {
    examples: ["await DocumentFile.patchDocx(docx, [{ path: 'customXml/review-note.xml', text: '<review>ok</review>' }])"],
    schema: {
      parameters: {
        docx: { type: "FileBlob|Uint8Array", required: true, description: "DOCX package bytes." },
        patches: { type: "array|object", required: true, description: "Path-validated package part edits with text/xml/json/bytes/remove." },
        maxPatchBytes: { type: "number", description: "Per-part patch size limit." },
      },
      returns: { docx: { type: "FileBlob", description: "Patched DOCX FileBlob with metadata.patchedParts." } },
    },
  },
  "PresentationFile.inspectPptx": {
    examples: ["await PresentationFile.inspectPptx(pptx, { includeText: true, maxChars: 12000 })"],
    schema: {
      parameters: {
        pptx: { type: "FileBlob|Uint8Array", required: true, description: "PPTX package bytes." },
        includeText: { type: "boolean", description: "Include bounded XML, relationship, and JSON text previews." },
        maxPreviewChars: { type: "number", description: "Maximum preview characters per textual package part." },
        maxChars: { type: "number", description: "Maximum bounded NDJSON output size." },
      },
      returns: { package: { type: "object", description: "PPTX package and part records with paths, sizes, content types, and optional previews." } },
    },
  },
  "PdfFile.inspectPdf": {
    examples: ["await PdfFile.inspectPdf(pdf, { maxObjects: 200, maxChars: 12000 })"],
    schema: {
      parameters: {
        pdf: { type: "FileBlob|Uint8Array", required: true, description: "PDF file bytes." },
        maxObjects: { type: "number", description: "Maximum indirect object records to inspect." },
        maxChars: { type: "number", description: "Maximum bounded NDJSON output size." },
      },
      returns: { inspection: { type: "object", description: "PDF file summary plus bounded indirect object records." } },
    },
  },
  "workbook.structuredReferences": {
    examples: ["=SUM(TasksTable[Revenue])", "=TEXTJOIN(\"|\",TRUE,TasksTable[#Headers])", "=SUM(TasksTable[[#Data],[Revenue]])", "=SUM(TasksTable[[#Data],[Revenue]:[Cost]])", "=TEXTJOIN(\"|\",TRUE,TasksTable[[#Data],[Region],[Code]])"],
    notes: ["Current clean-room subset supports #Headers/#Data/#All/#Totals sections, single-column selectors, contiguous column ranges, and comma-separated column unions; special escaping for headers containing brackets remains roadmap."],
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
    green: "#22c55e",
    red: "#ef4444",
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

function pivotValueLabel(valueField = {}) {
  const summarizeBy = valueField.summarizeBy || valueField.aggregation || "sum";
  return valueField.name || `${summarizeBy} of ${valueField.field || valueField.name || "Value"}`;
}

function summarizePivotValues(values = [], summarizeBy = "sum") {
  const nums = values.map((value) => Number(value)).filter((value) => Number.isFinite(value));
  const mode = String(summarizeBy || "sum").toLowerCase();
  if (mode === "count") return values.filter((value) => value != null && value !== "").length;
  if (!nums.length) return 0;
  if (mode === "average" || mode === "avg") return nums.reduce((sum, value) => sum + value, 0) / nums.length;
  if (mode === "min") return Math.min(...nums);
  if (mode === "max") return Math.max(...nums);
  return nums.reduce((sum, value) => sum + value, 0);
}

class WorksheetPivotTable {
  constructor(worksheet, config = {}) {
    this.worksheet = worksheet;
    this.id = config.id || aid("pvt");
    this.name = config.name || `PivotTable${worksheet.pivotTables.items.length + 1}`;
    this.sourceRange = workbookRangeRef(config.sourceRange || config.source || config.range || "A1");
    this.targetRange = workbookRangeRef(config.targetRange || config.destination || config.target || "A6");
    this.rowFields = [...(config.rowFields || config.rows || [])];
    this.columnFields = [...(config.columnFields || config.columns || [])];
    this.valueFields = (config.valueFields || config.values || []).map((field) => typeof field === "string" ? { field, summarizeBy: "sum" } : { ...field });
    this.filters = config.filters || {};
  }

  sourceValues() {
    const target = workbookRangeTarget(this.worksheet.workbook, this.worksheet, this.sourceRange);
    if (!target.sheet || !target.bounds) return [];
    return target.sheet.getRange(target.address).values;
  }

  computedValues() {
    const matrix = this.sourceValues();
    if (!matrix.length) return [];
    const headers = matrix[0].map((value) => String(value ?? ""));
    let rowIndexes = this.rowFields.map((field) => headers.indexOf(String(field))).filter((index) => index >= 0);
    if (this.rowFields.length && rowIndexes.length === 0 && headers.length) rowIndexes = [0];
    const valueConfigs = this.valueFields.length ? this.valueFields : headers.slice(1, 2).map((field) => ({ field, summarizeBy: "sum" }));
    const valueIndexes = valueConfigs.map((field) => headers.indexOf(String(field.field || field.name))).map((index, i) => ({ index, config: valueConfigs[i] })).filter((item) => item.index >= 0);
    const groups = new Map();
    for (const row of matrix.slice(1)) {
      const keyValues = rowIndexes.length ? rowIndexes.map((index) => row[index]) : ["(all)"];
      const key = JSON.stringify(keyValues);
      if (!groups.has(key)) groups.set(key, { keyValues, rows: [] });
      groups.get(key).rows.push(row);
    }
    const header = [...(this.rowFields.length ? this.rowFields : ["Group"]), ...valueIndexes.map((item) => pivotValueLabel(item.config))];
    const rows = [...groups.values()].map((group) => [
      ...group.keyValues,
      ...valueIndexes.map((item) => summarizePivotValues(group.rows.map((row) => row[item.index]), item.config.summarizeBy)),
    ]);
    return [header, ...rows];
  }

  inspectRecord() {
    const values = this.computedValues();
    return { kind: "pivotTable", id: this.id, sheet: this.worksheet.name, name: this.name, sourceRange: this.sourceRange.address, sourceSheet: this.sourceRange.sheetName || this.worksheet.name, targetRange: this.targetRange.address, rowFields: this.rowFields, columnFields: this.columnFields, valueFields: this.valueFields, values, rows: Math.max(0, values.length - 1), cols: values[0]?.length || 0 };
  }

  layoutJson(bounds) {
    const values = this.computedValues();
    const target = safeRangeBounds(this.targetRange.address) || { top: 0, left: 0, rowCount: Math.max(1, values.length), colCount: Math.max(1, values[0]?.length || 1) };
    const frame = { left: 40 + (target.left - bounds.left) * 96, top: 40 + (target.top - bounds.top) * 28, width: Math.max(96, (values[0]?.length || target.colCount || 1) * 96), height: Math.max(28, Math.max(values.length, target.rowCount || 1) * 28) };
    return { kind: "pivotTable", id: this.id, sheet: this.worksheet.name, name: this.name, sourceRange: this.sourceRange.address, targetRange: this.targetRange.address, rowFields: this.rowFields, valueFields: this.valueFields, values, bbox: [frame.left, frame.top, frame.width, frame.height] };
  }

  toSvg(bounds) {
    const layout = this.layoutJson(bounds);
    const [left, top] = layout.bbox;
    const values = layout.values || [];
    const cellW = values[0]?.length ? Math.max(72, layout.bbox[2] / values[0].length) : 96;
    const cellH = 24;
    const cells = [];
    values.forEach((row, r) => row.forEach((value, c) => {
      const x = left + c * cellW;
      const y = top + r * cellH;
      cells.push(`<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" fill="${r === 0 ? "#ecfeff" : "#ffffff"}" stroke="#67e8f9"/>`);
      cells.push(`<text x="${x + 5}" y="${y + 16}" font-family="Arial" font-size="11" fill="#155e75">${xmlEscape(value ?? "")}</text>`);
    }));
    return `<rect x="${left}" y="${Math.max(0, top - 18)}" width="${Math.max(120, layout.bbox[2])}" height="18" fill="#cffafe" stroke="#06b6d4"/><text x="${left + 5}" y="${Math.max(12, top - 5)}" font-family="Arial" font-size="11" font-weight="700" fill="#155e75">${xmlEscape(this.name)}</text>${cells.join("")}`;
  }

  toJSON() { return { id: this.id, name: this.name, sourceRange: this.sourceRange, targetRange: this.targetRange, rowFields: this.rowFields, columnFields: this.columnFields, valueFields: this.valueFields, filters: this.filters }; }
}

class WorksheetPivotTableCollection {
  constructor(worksheet) { this.worksheet = worksheet; this.items = []; }
  add(config = {}) { const pivot = new WorksheetPivotTable(this.worksheet, config); this.items.push(pivot); return pivot; }
  getItemOrNullObject(name) { return this.items.find((pivot) => pivot.name === name || pivot.id === name) || { isNullObject: true }; }
  deleteAll() { this.items = []; }
  inspectRecords() { return this.items.map((pivot) => pivot.inspectRecord()); }
  toJSON() { return this.items.map((pivot) => pivot.toJSON()); }
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
  addColorScale(config = {}) { return this.add("colorScale", config); }
  deleteAll() { this.range.worksheet.conditionalFormattings.items = this.range.worksheet.conditionalFormattings.items.filter((item) => item.range !== rangeToAddress(this.range.bounds)); }
  clear() { this.deleteAll(); }
}

function cellAddressWithinBounds(address, bounds) {
  if (!bounds) return false;
  try {
    const coord = parseCellAddress(String(address || "").replaceAll("$", ""));
    return coord.row >= bounds.top && coord.row <= bounds.bottom && coord.col >= bounds.left && coord.col <= bounds.right;
  } catch {
    return false;
  }
}

function mergeCellStyle(base = {}, patch = {}) {
  const out = { ...(base || {}) };
  for (const [key, value] of Object.entries(patch || {})) {
    if (value && typeof value === "object" && !Array.isArray(value) && out[key] && typeof out[key] === "object" && !Array.isArray(out[key])) out[key] = { ...out[key], ...value };
    else out[key] = value;
  }
  return out;
}

function conditionalFormulaForCell(formula, ruleBounds, address) {
  const raw = String(formula || "").trim();
  if (!ruleBounds || !address) return raw;
  let target;
  try { target = parseCellAddress(String(address).replaceAll("$", "")); } catch { return raw; }
  const rowOffset = target.row - ruleBounds.top;
  const colOffset = target.col - ruleBounds.left;
  return raw.replace(/(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?)([A-Za-z]+)(\$?)(\d+)/g, (match, quotedSheet, bareSheet, absCol, colText, absRow, rowText) => {
    if (quotedSheet || bareSheet) return match;
    const col = columnLabelToNumber(colText);
    const row = Number(rowText) - 1;
    const shiftedCol = absCol ? col : Math.max(0, col + colOffset);
    const shiftedRow = absRow ? row : Math.max(0, row + rowOffset);
    return `${absCol || ""}${columnNumberToLabel(shiftedCol)}${absRow || ""}${shiftedRow + 1}`;
  });
}

function conditionalScalar(sheet, value, address, ruleBounds) {
  if (value == null || value === "") return undefined;
  const shifted = conditionalFormulaForCell(String(value).replace(/^=/, ""), ruleBounds, address);
  const scalar = formulaScalar(sheet, shifted, {});
  return scalar === undefined ? value : scalar;
}

function compareConditionalValues(actualValue, expectedValue, operator = "equalTo") {
  const actualNum = Number(actualValue);
  const expectedNum = Number(expectedValue);
  const numeric = Number.isFinite(actualNum) && Number.isFinite(expectedNum) && String(expectedValue ?? "") !== "";
  const actual = numeric ? actualNum : formulaText(actualValue);
  const expected = numeric ? expectedNum : formulaText(expectedValue);
  switch (String(operator || "equalTo").toLowerCase()) {
    case "greaterthan": return actual > expected;
    case "greaterthanorequal":
    case "greaterthanorequalto": return actual >= expected;
    case "lessthan": return actual < expected;
    case "lessthanorequal":
    case "lessthanorequalto": return actual <= expected;
    case "notequal":
    case "not equal": return actual !== expected;
    case "equal":
    case "equalto": return actual === expected;
    default: return matchesCriteria(actualValue, `${operator}${expectedValue}`);
  }
}

function xlsxColorToRgb(value, fallback = "FFFFFF") {
  const hex = normalizeXlsxColor(value, fallback);
  return [Number.parseInt(hex.slice(0, 2), 16), Number.parseInt(hex.slice(2, 4), 16), Number.parseInt(hex.slice(4, 6), 16)];
}

function rgbToHex(rgb) {
  return `#${rgb.map((value) => Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, "0")).join("")}`;
}

function interpolateColor(left, right, amount) {
  const a = xlsxColorToRgb(left);
  const b = xlsxColorToRgb(right);
  return rgbToHex(a.map((value, index) => value + (b[index] - value) * Math.max(0, Math.min(1, amount))));
}

function colorScaleColors(rule = {}) {
  const raw = rule.colors || rule.colorScale?.colors || rule.rule?.colors || rule.rule?.colorScale?.colors;
  const colors = Array.isArray(raw) && raw.length >= 2 ? raw : [rule.minColor || "#ef4444", rule.midColor || "#facc15", rule.maxColor || "#22c55e"];
  return colors.slice(0, 3);
}

function colorScaleFormatForCell(sheet, rule, address) {
  const bounds = safeRangeBounds(rule.range || "");
  if (!cellAddressWithinBounds(address, bounds)) return undefined;
  const cell = sheet.store.get(String(address).toUpperCase());
  const value = Number(cell.value);
  if (!Number.isFinite(value)) return undefined;
  const values = sheet.getRange(rule.range || address).values.flat().map(Number).filter(Number.isFinite);
  if (!values.length) return undefined;
  const min = Math.min(...values);
  const max = Math.max(...values);
  const colors = colorScaleColors(rule);
  const ratio = max === min ? 1 : (value - min) / (max - min);
  const fill = colors.length >= 3
    ? ratio <= 0.5 ? interpolateColor(colors[0], colors[1], ratio * 2) : interpolateColor(colors[1], colors[2], (ratio - 0.5) * 2)
    : interpolateColor(colors[0], colors[1], ratio);
  return { fill };
}

function evaluateConditionalFormatRule(sheet, rule, address) {
  const bounds = safeRangeBounds(rule.range || "");
  if (!cellAddressWithinBounds(address, bounds)) return false;
  const cell = sheet.store.get(String(address).toUpperCase());
  const ruleType = String(rule.ruleType || rule.type || "expression");
  const normalizedType = ruleType.toLowerCase();
  if (normalizedType === "colorscale") return Boolean(colorScaleFormatForCell(sheet, rule, address));
  if (normalizedType === "cellis" || normalizedType === "cellvalue") {
    const first = conditionalScalar(sheet, rule.formula ?? rule.formula1 ?? rule.rule?.formula ?? rule.rule?.formula1, address, bounds);
    const second = conditionalScalar(sheet, rule.formula2 ?? rule.rule?.formula2, address, bounds);
    const operator = rule.operator || rule.rule?.operator || "equalTo";
    if (String(operator).toLowerCase() === "between") return compareConditionalValues(cell.value, first, "greaterThanOrEqual") && compareConditionalValues(cell.value, second, "lessThanOrEqual");
    if (String(operator).toLowerCase() === "notbetween") return !(compareConditionalValues(cell.value, first, "greaterThanOrEqual") && compareConditionalValues(cell.value, second, "lessThanOrEqual"));
    return compareConditionalValues(cell.value, first, operator);
  }
  if (normalizedType === "containstext") {
    const expected = rule.text ?? rule.formula ?? rule.rule?.text ?? rule.rule?.formula;
    return formulaText(cell.value).includes(formulaText(expected));
  }
  const formula = rule.formula || rule.expression || rule.rule?.formula || rule.rule?.expression;
  if (!formula) return false;
  const shifted = conditionalFormulaForCell(String(formula).replace(/^=/, ""), bounds, address);
  return Boolean(evaluateFormulaCondition(sheet, shifted));
}

function worksheetConditionalFormatMatches(sheet, address) {
  const matches = [];
  for (const rule of sheet.conditionalFormattings.items) {
    const ruleType = rule.ruleType || rule.type || "expression";
    if (String(ruleType).toLowerCase() === "colorscale") {
      const format = colorScaleFormatForCell(sheet, rule, address);
      if (format) matches.push({ id: rule.id, range: rule.range, ruleType, operator: rule.operator || rule.rule?.operator, formula: rule.formula || rule.expression || rule.rule?.formula, format, colors: colorScaleColors(rule) });
    } else if (evaluateConditionalFormatRule(sheet, rule, address)) {
      matches.push({ id: rule.id, range: rule.range, ruleType, operator: rule.operator || rule.rule?.operator, formula: rule.formula || rule.expression || rule.rule?.formula, format: rule.format || rule.rule?.format || {} });
    }
  }
  return matches;
}

function worksheetComputedCellStyle(sheet, address, baseStyle = {}) {
  const matches = worksheetConditionalFormatMatches(sheet, address);
  const style = matches.reduce((acc, match) => mergeCellStyle(acc, match.format), { ...(baseStyle || {}) });
  return { style, matches };
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
      clearFormulaSpills(this);
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
        if (isFormulaMatrix(value)) {
          const spill = writeFormulaSpill(node.sheetObject, node.address, value, node.key);
          node.cell.value = spill.value;
          node.cell.spillRange = spill.range;
          node.cell.spillValues = spill.values;
          if (spill.blocked) node.cell.spillError = { type: "blocked", addresses: spill.blocked };
        } else {
          node.cell.value = value;
        }
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
      if (kinds.has("pivotTable") || kinds.has("pivot")) records.push(...sheet.pivotTables.inspectRecords());
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
        if (/^#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A|CYCLE!|SPILL!)/.test(value)) {
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
      for (const pivot of sheet.pivotTables.items) {
        const source = workbookRangeTarget(this, sheet, pivot.sourceRange);
        if (!source.sheet || !source.bounds) issues.push(verificationIssue("workbook", "pivotSourceInvalid", `Pivot table ${pivot.name || pivot.id} on ${sheet.name} has invalid source range.`, { sheet: sheet.name, id: pivot.id, sourceRange: pivot.sourceRange.address }));
        const headers = pivot.sourceValues()[0]?.map((value) => String(value ?? "")) || [];
        for (const field of [...pivot.rowFields, ...pivot.valueFields.map((item) => item.field || item.name)]) {
          if (field && !headers.includes(String(field))) issues.push(verificationIssue("workbook", "pivotFieldMissing", `Pivot table ${pivot.name || pivot.id} references missing field ${field}.`, { sheet: sheet.name, id: pivot.id, field }));
        }
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
        const cfType = String(format.ruleType || format.type || "").toLowerCase();
        if ((cfType === "expression" || (!format.ruleType && format.kind === "conditionalFormat")) && !format.formula && !format.expression) issues.push(verificationIssue("workbook", "conditionalFormatFormulaMissing", `Conditional format ${format.id} on ${sheet.name} has no formula/expression.`, { sheet: sheet.name, id: format.id }));
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
      const pivot = sheet.pivotTables.items.find((item) => item.id === id || item.name === id);
      if (pivot) return pivot;
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
    const targets = inspectTargetTokens(options);
    const search = String(options.search || options.searchTerm || "").trim();
    const targetsWorkbook = targets.some((target) => target === this.id || target === "workbookLayout" || target === "workbook");
    const sheetOptions = targetsWorkbook ? { ...options, target: undefined, targetId: undefined, id: undefined, anchor: undefined } : options;
    let sheetLayouts = sheets.map((sheet) => sheet.layoutJson(sheetOptions));
    if ((targets.length || search) && !targetsWorkbook && !selected) sheetLayouts = sheetLayouts.filter((sheetLayout) => !sheetLayout.slice || sheetLayout.slice.returnedRecords > 0);
    return {
      kind: "workbookLayout",
      id: this.id,
      activeSheet: this.worksheets.items[0]?.name,
      sheetCount: this.worksheets.items.length,
      sheets: sheetLayouts,
      slice: (targets.length || search) ? { targets, search: search || undefined, returnedSheets: sheetLayouts.length } : undefined,
    };
  }

  async render(options = {}) {
    const format = String(options.format || "").trim().toLowerCase();
    if (format === "layout" || format === LAYOUT_MIME) {
      return new FileBlob(JSON.stringify(this.layoutJson(options), null, 2), {
        type: LAYOUT_MIME,
        metadata: { artifactKind: "workbook", format: "layout", sheetName: options.sheetName, range: options.range, target: options.target ?? options.targetId ?? options.id ?? options.anchor, search: options.search ?? options.searchTerm },
      });
    }
    const sheet = this.worksheets.getItem(options.sheetName) || this.worksheets.getActiveWorksheet();
    return new FileBlob(sheet.toSvg(options), { type: "image/svg+xml" });
  }
}

function worksheetLayoutEntries(layout) {
  const entries = [];
  for (const collection of ["cells", "tables", "pivots", "charts", "images", "sparklines", "rules"]) {
    layout[collection].forEach((record, itemIndex) => entries.push({ collection, itemIndex, record }));
  }
  return entries;
}

function worksheetLayoutSlice(layout, options = {}) {
  const targets = inspectTargetTokens(options);
  const search = String(options.search || options.searchTerm || "").trim().toLowerCase();
  if (!targets.length && !search) return layout;
  const before = Math.max(0, Number(options.before ?? options.contextBefore ?? options.context ?? 0) || 0);
  const after = Math.max(0, Number(options.after ?? options.contextAfter ?? options.context ?? 0) || 0);
  const targetsSheet = targets.some((target) => target === layout.id || target === layout.name || target === layout.sheet);
  if (targetsSheet && !search) {
    const returnedRecords = worksheetLayoutEntries(layout).length;
    return { ...layout, slice: { targets, before, after, matchedRecords: returnedRecords, returnedRecords } };
  }
  const entries = worksheetLayoutEntries(layout);
  const matches = [];
  entries.forEach((entry, index) => {
    const matchesSearch = !search || JSON.stringify(entry.record).toLowerCase().includes(search);
    const matchesTarget = !targets.length || targetsSheet || inspectRecordMatchesTarget(entry.record, targets);
    if (matchesSearch && matchesTarget) matches.push(index);
  });
  const keep = new Set();
  for (const index of matches) {
    for (let i = Math.max(0, index - before); i <= Math.min(entries.length - 1, index + after); i += 1) keep.add(i);
  }
  const keepByCollection = new Map();
  for (const entryIndex of keep) {
    const entry = entries[entryIndex];
    if (!keepByCollection.has(entry.collection)) keepByCollection.set(entry.collection, new Set());
    keepByCollection.get(entry.collection).add(entry.itemIndex);
  }
  const sliced = { ...layout };
  for (const collection of ["cells", "tables", "pivots", "charts", "images", "sparklines", "rules"]) {
    const kept = keepByCollection.get(collection) || new Set();
    sliced[collection] = layout[collection].filter((_, index) => kept.has(index));
  }
  sliced.slice = { targets, search: search || undefined, before, after, matchedRecords: matches.length, returnedRecords: keep.size };
  return sliced;
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
    this.pivotTables = new WorksheetPivotTableCollection(this);
    this.pivots = this.pivotTables;
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
        formulaType: cell.formulaType || undefined,
        sharedIndex: cell.sharedIndex,
        sharedRef: cell.sharedRef,
        arrayRef: cell.arrayRef,
        spillRange: cell.spillRange,
        spillValues: cell.spillValues,
        spillError: cell.spillError,
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
    const kinds = normalizeKinds(options.kind, ["style"]);
    const includeBaseStyle = kinds.has("style");
    const includeComputedStyle = kinds.has("computedStyle");
    const records = [];
    for (const [address, cell] of this.store.entries()) {
      const coord = parseCellAddress(address);
      if (coord.row < bounds.top || coord.row > bounds.bottom || coord.col < bounds.left || coord.col > bounds.right) continue;
      if (includeBaseStyle && cell.style && Object.keys(cell.style).length > 0) records.push({ kind: "style", sheet: this.name, address, style: { ...cell.style } });
      if (includeComputedStyle) {
        const computed = worksheetComputedCellStyle(this, address, cell.style);
        if (computed.matches.length || Object.keys(computed.style || {}).length) records.push({ kind: "computedStyle", sheet: this.name, address, style: computed.style, conditionalFormats: computed.matches.map((match) => ({ id: match.id, ruleType: match.ruleType, operator: match.operator, formula: match.formula, format: match.format })) });
      }
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
        const computed = worksheetComputedCellStyle(this, address, cell.style);
        cells.push({
          kind: "cell",
          sheet: this.name,
          address,
          row: r,
          col: c,
          bbox: [40 + (c - bounds.left) * cellW, 40 + (r - bounds.top) * cellH, cellW, cellH],
          value: cell.value,
          formula: cell.formula || undefined,
          spillParent: cell.spillParent,
          spillRange: cell.spillRange,
          style: cell.style && Object.keys(cell.style).length ? { ...cell.style } : undefined,
          computedStyle: computed.matches.length || Object.keys(computed.style || {}).length ? computed.style : undefined,
          conditionalFormats: computed.matches.length ? computed.matches.map((match) => ({ id: match.id, ruleType: match.ruleType, operator: match.operator, formula: match.formula, format: match.format })) : undefined,
        });
      }
    }
    const tables = this.tables.items.map((table) => ({ kind: "table", id: table.id, sheet: this.name, name: table.name, address: table.range, bbox: Object.values(frameForRange(table.range) || {}), rows: table.rowCount, cols: table.columnCount, hasHeaders: table.hasHeaders }));
    const pivots = this.pivotTables.items.map((pivot) => pivot.layoutJson(bounds));
    const charts = this.charts.items.map((chart) => ({ kind: "chart", id: chart.id, sheet: this.name, name: chart.name, chartType: chart.type, title: chart.title, bbox: [chart.position.left, chart.position.top, chart.position.width, chart.position.height], series: chart.series.items.length, categories: chart.categories }));
    const images = this.images.items.map((image) => { const p = image.position; return { kind: "image", id: image.id, sheet: this.name, name: image.name, alt: image.alt, bbox: [p.left, p.top, p.width, p.height], fit: image.fit, hasDataUrl: Boolean(image.dataUrl), uri: image.uri, prompt: image.prompt }; });
    const sparklines = this.sparklineGroups.items.map((group) => { const p = group.targetFrame(bounds); return { kind: "sparkline", id: group.id, sheet: this.name, type: group.type, targetRange: group.targetRange.address, sourceData: group.sourceData.address, bbox: [p.left, p.top, p.width, p.height], values: group.values() }; });
    const rules = [...this.dataValidations.items, ...this.conditionalFormattings.items].map((rule) => ({ kind: rule.kind, id: rule.id, sheet: this.name, range: rule.range, bbox: Object.values(frameForRange(rule.range) || {}), ruleType: rule.ruleType, rule: rule.rule }));
    const drawingFrames = [...charts.map((item) => item.bbox), ...images.map((item) => item.bbox), ...sparklines.map((item) => item.bbox), ...tables.map((item) => item.bbox), ...pivots.map((item) => item.bbox)].filter((bbox) => bbox.length === 4);
    const width = Math.max(320, bounds.colCount * cellW + 80, ...drawingFrames.map((bbox) => bbox[0] + bbox[2] + 40));
    const height = Math.max(180, bounds.rowCount * cellH + 80, ...drawingFrames.map((bbox) => bbox[1] + bbox[3] + 40));
    const layout = {
      kind: "worksheetLayout",
      id: this.id,
      name: this.name,
      sheet: this.name,
      bounds: { ...bounds, address: rangeToAddress(bounds) },
      unit: "px",
      origin: { left: 40, top: 40 },
      cell: { width: cellW, height: cellH },
      width,
      height,
      cells,
      tables,
      pivots,
      charts,
      images,
      sparklines,
      rules,
    };
    return worksheetLayoutSlice(layout, options);
  }

  toSvg() {
    this.recalculate();
    const bounds = this.usedBounds();
    const cellW = 96;
    const cellH = 28;
    const imageFrames = this.images.items.map((image) => image.position);
    const sparklineFrames = this.sparklineGroups.items.map((group) => group.targetFrame(bounds));
    const pivotFrames = this.pivotTables.items.map((pivot) => { const bbox = pivot.layoutJson(bounds).bbox; return { left: bbox[0], top: bbox[1], width: bbox[2], height: bbox[3] }; });
    const width = Math.max(320, bounds.colCount * cellW + 80, ...this.charts.items.map((chart) => chart.position.left + chart.position.width + 40), ...imageFrames.map((frame) => frame.left + frame.width + 40), ...sparklineFrames.map((frame) => frame.left + frame.width + 40), ...pivotFrames.map((frame) => frame.left + frame.width + 40));
    const height = Math.max(180, bounds.rowCount * cellH + 80, ...this.charts.items.map((chart) => chart.position.top + chart.position.height + 40), ...imageFrames.map((frame) => frame.top + frame.height + 40), ...sparklineFrames.map((frame) => frame.top + frame.height + 40), ...pivotFrames.map((frame) => frame.top + frame.height + 40));
    const rows = [];
    for (let r = bounds.top; r <= bounds.bottom; r++) {
      for (let c = bounds.left; c <= bounds.right; c++) {
        const address = makeCellAddress(r, c);
        const cell = this.store.get(address);
        const value = cell.value ?? "";
        const computed = worksheetComputedCellStyle(this, address, cell.style);
        const x = 40 + (c - bounds.left) * cellW;
        const y = 40 + (r - bounds.top) * cellH;
        const fill = resolveColorToken(computed.style?.fill || "white", "white");
        const font = computed.style?.font || {};
        const fontWeight = font.bold || computed.style?.bold ? "700" : "400";
        const fontFill = resolveColorToken(font.color || computed.style?.color || "#24292f", "#24292f");
        rows.push(`<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" fill="${xmlEscape(fill)}" stroke="#d0d7de"/>`);
        rows.push(`<text x="${x + 6}" y="${y + 18}" font-family="${xmlEscape(font.name || "Arial")}" font-size="13" font-weight="${fontWeight}" fill="${xmlEscape(fontFill)}">${xmlEscape(value)}</text>`);
      }
    }
    const tableOverlays = this.tables.items.map((table) => table.toSvg(bounds)).join("");
    const pivotOverlays = this.pivotTables.items.map((pivot) => pivot.toSvg(bounds)).join("");
    const chartOverlays = this.charts.items.map((chart) => chart.toSvg()).join("");
    const imageOverlays = this.images.items.map((image) => image.toSvg()).join("");
    const sparklineOverlays = this.sparklineGroups.items.map((group) => group.toSvg(bounds)).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="#f6f8fa"/>${rows.join("")}${tableOverlays}${pivotOverlays}${chartOverlays}${imageOverlays}${sparklineOverlays}</svg>`;
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
        delete cell.spillParent;
        delete cell.spillAnchor;
        delete cell.spillRange;
        delete cell.spillValues;
        delete cell.spillError;
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

function structuredRefSegments(refBody = "") {
  const body = String(refBody || "").trim();
  const segments = [];
  let last = 0;
  for (let i = 0; i < body.length; i++) {
    if (body[i] !== "[") continue;
    let end = i + 1;
    while (end < body.length && body[end] !== "]") end += 1;
    if (end >= body.length) break;
    if (i > last) segments.push({ type: "sep", text: body.slice(last, i) });
    segments.push({ type: "token", text: body.slice(i + 1, end).trim() });
    i = end;
    last = end + 1;
  }
  if (segments.length && last < body.length) segments.push({ type: "sep", text: body.slice(last) });
  return segments;
}

function structuredRefParts(refBody = "") {
  const body = String(refBody || "").trim();
  const segments = structuredRefSegments(body);
  if (!segments.length) {
    const tokens = body.split(",").map((item) => item.trim()).filter(Boolean);
    return {
      tokens,
      sectionTokens: tokens.filter((token) => token.startsWith("#")),
      columnSelectors: tokens.filter((token) => !token.startsWith("#")).map((token) => {
        const range = token.split(":").map((item) => item.trim()).filter(Boolean);
        return range.length === 2 ? { start: range[0], end: range[1] } : { name: token };
      }),
    };
  }
  const tokens = segments.filter((segment) => segment.type === "token").map((segment) => segment.text).filter(Boolean);
  const columnSelectors = [];
  for (let i = 0; i < segments.length; i++) {
    const segment = segments[i];
    if (segment.type !== "token" || !segment.text || segment.text.startsWith("#")) continue;
    const sep = segments[i + 1]?.type === "sep" ? segments[i + 1].text : "";
    const next = segments[i + 2];
    if (sep.includes(":") && next?.type === "token" && next.text && !next.text.startsWith("#")) {
      columnSelectors.push({ start: segment.text, end: next.text });
      i += 2;
    } else {
      columnSelectors.push({ name: segment.text });
    }
  }
  return { tokens, sectionTokens: tokens.filter((token) => token.startsWith("#")), columnSelectors };
}

function structuredColumnIndex(headers, columnName) {
  return headers.findIndex((header) => String(header ?? "").trim() === String(columnName ?? "").trim());
}

function structuredColumnIndexes(selectors, headers) {
  const columnCount = headers.length;
  if (!selectors.length) return { indexes: Array.from({ length: columnCount }, (_, index) => index) };
  const indexes = [];
  const push = (index) => { if (!indexes.includes(index)) indexes.push(index); };
  for (const selector of selectors) {
    if (selector.start && selector.end) {
      const start = structuredColumnIndex(headers, selector.start);
      const end = structuredColumnIndex(headers, selector.end);
      if (start < 0) return { missing: selector.start };
      if (end < 0) return { missing: selector.end };
      const step = start <= end ? 1 : -1;
      for (let index = start; step > 0 ? index <= end : index >= end; index += step) push(index);
    } else {
      const index = structuredColumnIndex(headers, selector.name);
      if (index < 0) return { missing: selector.name };
      push(index);
    }
  }
  return { indexes };
}

function formulaStructuredRefRange(sheet, refText = "") {
  const match = /^([A-Za-z_][A-Za-z0-9_.]*)\[((?:[^\[\]]+|\[[^\]]+\])+)\]$/.exec(String(refText || "").trim());
  if (!match) return undefined;
  const tableName = match[1];
  const parts = structuredRefParts(match[2]);
  const tokens = parts.tokens;
  const found = findWorkbookTable(sheet, tableName);
  if (!found) return { missing: true, tableName, columnName: tokens.join(",") };
  const { table, sheet: tableSheet } = found;
  const bounds = parseRangeAddress(table.range);
  const columnCount = bounds.right - bounds.left + 1;
  const headers = Array.from({ length: columnCount }, (_, index) => table.hasHeaders ? (table.values[0]?.[index] ?? `Column${index + 1}`) : `Column${index + 1}`);
  const section = parts.sectionTokens.at(-1) || "#Data";
  const firstDataRow = bounds.top + (table.showHeaders ? 1 : 0);
  const totalsRow = table.showTotals ? bounds.bottom : undefined;
  const lastDataRow = bounds.bottom - (table.showTotals ? 1 : 0);
  let top = firstDataRow;
  let bottom = lastDataRow;
  if (/^#Headers$/i.test(section)) {
    if (!table.showHeaders) return { sheetName: tableSheet.name, start: makeCellAddress(bounds.top, bounds.left), end: makeCellAddress(bounds.top - 1, bounds.left), empty: true, tableName, columnName: parts.columnSelectors.map((item) => item.name || `${item.start}:${item.end}`).join(","), table };
    top = bottom = bounds.top;
  } else if (/^#Totals$/i.test(section)) {
    if (totalsRow == null) return { sheetName: tableSheet.name, start: makeCellAddress(bounds.bottom + 1, bounds.left), end: makeCellAddress(bounds.bottom, bounds.left), empty: true, tableName, columnName: parts.columnSelectors.map((item) => item.name || `${item.start}:${item.end}`).join(","), table };
    top = bottom = totalsRow;
  } else if (/^#All$/i.test(section)) {
    top = bounds.top;
    bottom = bounds.bottom;
  } else if (/^#Data$/i.test(section)) {
    top = firstDataRow;
    bottom = lastDataRow;
  } else if (parts.sectionTokens.length) {
    return { missing: true, tableName, columnName: section, sheetName: tableSheet.name };
  }
  const selected = structuredColumnIndexes(parts.columnSelectors, headers);
  if (selected.missing) return { missing: true, tableName, columnName: selected.missing, sheetName: tableSheet.name };
  const columns = selected.indexes;
  if (!columns.length) return { missing: true, tableName, columnName: tokens.join(","), sheetName: tableSheet.name };
  const left = bounds.left + Math.min(...columns);
  const right = bounds.left + Math.max(...columns);
  const columnNames = columns.map((index) => String(headers[index] ?? `Column${index + 1}`));
  const columnName = columnNames.join(",");
  const columnIndex = columns.length === 1 ? columns[0] : undefined;
  if (top > bottom) return { sheetName: tableSheet.name, start: makeCellAddress(top, left), end: makeCellAddress(bottom, right), empty: true, tableName, columnName, table, columnIndex, columns, columnNames, section };
  return { sheetName: tableSheet.name, start: makeCellAddress(top, left), end: makeCellAddress(bottom, right), tableName, columnName, table, columnIndex, columns, columnNames, section };
}

function formulaReferences(formula, sheet) {
  const raw = String(formula || "");
  const refs = [];
  const consumed = [];
  const structuredRegex = /\b([A-Za-z_][A-Za-z0-9_.]*)\[((?:[^\[\]]+|\[[^\]]+\])+)\]/g;
  for (const match of raw.matchAll(structuredRegex)) {
    consumed.push([match.index, match.index + match[0].length]);
    const structured = formulaStructuredRefRange(sheet, match[0]);
    if (!structured || structured.missing || structured.empty) continue;
    const start = parseCellAddress(structured.start);
    const end = parseCellAddress(structured.end);
    const tableBounds = structured.table ? parseRangeAddress(structured.table.range) : undefined;
    const cols = tableBounds && structured.columns?.length ? structured.columns.map((index) => tableBounds.left + index) : Array.from({ length: Math.abs(end.col - start.col) + 1 }, (_, index) => Math.min(start.col, end.col) + index);
    for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
      for (const col of cols) {
        refs.push({ sheetName: structured.sheetName, address: makeCellAddress(row, col), structuredRef: match[0], tableName: structured.tableName, columnName: structured.columnName, columnNames: structured.columnNames });
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
  return /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|CYCLE!|FIELD!|DATA!|CONNECT!|BLOCKED!|UNKNOWN!|BUSY!)/.test(text)
    ? text.match(/^#[A-Z0-9\/?!_]+/)?.[0]
    : undefined;
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
  let bracketDepth = 0;
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
    if (!inString && ch === "[") bracketDepth += 1;
    if (!inString && ch === "]") bracketDepth -= 1;
    if (!inString && ch === "," && depth === 0 && bracketDepth === 0) { args.push(current.trim()); current = ""; continue; }
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
    const tableBounds = structured.table ? parseRangeAddress(structured.table.range) : undefined;
    const cols = tableBounds && structured.columns?.length ? structured.columns.map((index) => tableBounds.left + index) : Array.from({ length: Math.abs(end.col - start.col) + 1 }, (_, index) => Math.min(start.col, end.col) + index);
    const rows = [];
    for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
      const values = [];
      for (const col of cols) {
        const address = makeCellAddress(row, col);
        let value = context.getValue ? context.getValue({ sheetName: structured.sheetName, address, structuredRef: refText, tableName: structured.tableName, columnName: structured.columnName, columnNames: structured.columnNames }) : targetSheet.store.get(address).value;
        if ((value == null || value === "") && structured.table) {
          const tableBounds = parseRangeAddress(structured.table.range);
          const tableRow = row - tableBounds.top;
          const tableCol = col - tableBounds.left;
          value = structured.table.values[tableRow]?.[tableCol] ?? value;
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
  const error = formulaErrorCode(text);
  if (error) return error;
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

function isFormulaMatrix(value) {
  return Array.isArray(value) && (Array.isArray(value[0]) || value.length === 0);
}

function normalizeFormulaMatrix(value) {
  if (!Array.isArray(value)) return [[value]];
  if (value.length === 0) return [];
  return Array.isArray(value[0]) ? value.map((row) => Array.isArray(row) ? row : [row]) : value.map((item) => [item]);
}

function clearFormulaSpills(workbook) {
  for (const sheet of workbook.worksheets) {
    for (const [address, cell] of sheet.store.entries()) {
      if (cell.spillParent) sheet.store.cells.delete(address);
      else {
        delete cell.spillRange;
        delete cell.spillValues;
        delete cell.spillError;
      }
    }
  }
}

function writeFormulaSpill(sheet, anchorAddress, matrixValue, parentKey) {
  const matrix = normalizeFormulaMatrix(matrixValue);
  const rows = matrix.length;
  const cols = Math.max(0, ...matrix.map((row) => row.length));
  if (!rows || !cols) return { value: null, range: anchorAddress, values: matrix };
  const anchor = parseCellAddress(anchorAddress);
  const spillRange = rangeToAddress({ top: anchor.row, left: anchor.col, bottom: anchor.row + rows - 1, right: anchor.col + cols - 1 });
  const collisions = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (r === 0 && c === 0) continue;
      const address = makeCellAddress(anchor.row + r, anchor.col + c);
      const cell = sheet.store.get(address);
      if (!cell.spillParent && (cell.formula || cell.value != null)) collisions.push(address);
    }
  }
  const anchorCell = sheet.store.get(anchorAddress);
  if (collisions.length) {
    anchorCell.value = "#SPILL!";
    anchorCell.spillRange = spillRange;
    anchorCell.spillValues = matrix;
    anchorCell.spillError = { type: "blocked", addresses: collisions };
    return { value: "#SPILL!", range: spillRange, values: matrix, blocked: collisions };
  }
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const address = makeCellAddress(anchor.row + r, anchor.col + c);
      const cell = sheet.store.get(address);
      cell.value = matrix[r]?.[c] ?? null;
      cell.formula = r === 0 && c === 0 ? anchorCell.formula : null;
      cell.spillParent = r === 0 && c === 0 ? undefined : parentKey;
      cell.spillAnchor = anchorAddress;
      cell.spillRange = spillRange;
      cell.spillValues = matrix;
      delete cell.spillError;
    }
  }
  return { value: matrix[0]?.[0] ?? null, range: spillRange, values: matrix };
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

function compareFormulaValues(left, op, right) {
  const leftNum = Number(left), rightNum = Number(right);
  const numeric = Number.isFinite(leftNum) && Number.isFinite(rightNum) && String(left ?? "").trim() !== "" && String(right ?? "").trim() !== "";
  const a = numeric ? leftNum : formulaText(left);
  const b = numeric ? rightNum : formulaText(right);
  switch (op) {
    case ">=": return a >= b;
    case "<=": return a <= b;
    case "<>": return a !== b;
    case ">": return a > b;
    case "<": return a < b;
    default: return a === b;
  }
}

function formulaCriteriaArray(sheet, expr, context = {}) {
  const text = String(expr || "").trim();
  const comparison = /^(.*?)\s*(>=|<=|<>|=|>|<)\s*(.*?)$/.exec(text);
  if (comparison) {
    const leftMatrix = formulaRangeMatrix(sheet, comparison[1], context);
    const rightMatrix = formulaRangeMatrix(sheet, comparison[3], context);
    const leftValues = leftMatrix ? leftMatrix.flat() : [formulaScalar(sheet, comparison[1], context)];
    const rightValues = rightMatrix ? rightMatrix.flat() : [formulaScalar(sheet, comparison[3], context)];
    const length = Math.max(leftValues.length, rightValues.length);
    return Array.from({ length }, (_, index) => compareFormulaValues(leftValues[leftValues.length === 1 ? 0 : index], comparison[2], rightValues[rightValues.length === 1 ? 0 : index]));
  }
  const matrix = formulaRangeMatrix(sheet, text, context);
  if (matrix) return matrix.flat().map(formulaTruthy);
  return [formulaTruthy(formulaScalar(sheet, text, context))];
}

function formulaSortCompare(left, right) {
  const leftNum = Number(left), rightNum = Number(right);
  if (Number.isFinite(leftNum) && Number.isFinite(rightNum)) return leftNum - rightNum;
  return formulaText(left).localeCompare(formulaText(right));
}

function formulaMatchIndex(lookup, lookupValues = [], matchType = 0) {
  const values = lookupValues.flat ? lookupValues.flat() : lookupValues;
  const same = (value) => formulaText(value) === formulaText(lookup) || Number(value) === Number(lookup);
  const exact = values.findIndex(same);
  if (matchType === 0 || exact >= 0) return exact >= 0 ? exact + 1 : "#N/A";
  const lookupNum = Number(lookup);
  if (Number.isFinite(lookupNum)) {
    if (matchType < 0) {
      const index = values.findIndex((value) => Number(value) <= lookupNum);
      return index >= 0 ? index + 1 : "#N/A";
    }
    let best = -1;
    for (let i = 0; i < values.length; i++) if (Number(values[i]) <= lookupNum) best = i;
    return best >= 0 ? best + 1 : "#N/A";
  }
  return "#N/A";
}

function uniqueFormulaRows(matrix) {
  const seen = new Set();
  const rows = [];
  for (const row of normalizeFormulaMatrix(matrix)) {
    const key = JSON.stringify(row.map((value) => [typeof value, value]));
    if (seen.has(key)) continue;
    seen.add(key);
    rows.push(row);
  }
  return rows;
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
    case "IFERROR": { const value = scalar(0); return formulaErrorCode(value) ? scalar(1, "") : value; }
    case "AND": return args.every((arg) => evaluateFormulaCondition(sheet, arg, context));
    case "OR": return args.some((arg) => evaluateFormulaCondition(sheet, arg, context));
    case "NOT": return !evaluateFormulaCondition(sheet, args[0], context);
    case "ISNUMBER": { const value = scalar(0); return typeof value === "number" && Number.isFinite(value); }
    case "ISTEXT": { const value = scalar(0); return typeof value === "string" && !formulaErrorCode(value); }
    case "ISBLANK": { const value = scalar(0); return value == null; }
    case "ISERROR": return Boolean(formulaErrorCode(scalar(0)));
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
    case "COUNTIFS": {
      const pairs = [];
      for (let i = 0; i < args.length; i += 2) pairs.push({ range: values([args[i]]), criteria: scalar(i + 1, "") });
      const length = Math.max(0, ...pairs.map((pair) => pair.range.length));
      let count = 0;
      for (let index = 0; index < length; index++) if (pairs.every((pair) => matchesCriteria(pair.range[index], pair.criteria))) count += 1;
      return count;
    }
    case "SEQUENCE": {
      const rows = Math.max(1, Math.floor(formulaNumber(scalar(0, 1))) || 1);
      const cols = Math.max(1, Math.floor(formulaNumber(scalar(1, 1))) || 1);
      const start = formulaNumber(scalar(2, 1));
      const step = formulaNumber(scalar(3, 1));
      return Array.from({ length: rows }, (_, row) => Array.from({ length: cols }, (_, col) => start + (row * cols + col) * step));
    }
    case "TRANSPOSE": {
      const matrix = formulaRangeMatrix(sheet, args[0], context) || [];
      const rows = matrix.length;
      const cols = Math.max(0, ...matrix.map((row) => row.length));
      return Array.from({ length: cols }, (_, col) => Array.from({ length: rows }, (_, row) => matrix[row]?.[col] ?? null));
    }
    case "FILTER": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const include = formulaCriteriaArray(sheet, args[1], context);
      const rows = matrix.filter((_, index) => formulaTruthy(include[index]));
      if (rows.length) return rows;
      return [[args[2] ? scalar(2, "") : "#CALC!"]];
    }
    case "UNIQUE": {
      return uniqueFormulaRows(formulaRangeMatrix(sheet, args[0], context) || []);
    }
    case "SORT": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const sortIndex = Math.max(0, Math.floor(formulaNumber(scalar(1, 1))) - 1);
      const sortOrder = formulaNumber(scalar(2, 1)) < 0 ? -1 : 1;
      return [...matrix].sort((a, b) => formulaSortCompare(a[sortIndex], b[sortIndex]) * sortOrder);
    }
    case "SUMIF": {
      const range = values([args[0]]);
      const criteria = scalar(1, "");
      const sumRange = args[2] ? values([args[2]]) : range;
      return range.reduce((sum, value, index) => sum + (matchesCriteria(value, criteria) ? formulaNumber(sumRange[index]) || 0 : 0), 0);
    }
    case "SUMIFS": {
      const sumRange = values([args[0]]);
      const pairs = [];
      for (let i = 1; i < args.length; i += 2) pairs.push({ range: values([args[i]]), criteria: scalar(i + 1, "") });
      return sumRange.reduce((sum, value, index) => sum + (pairs.every((pair) => matchesCriteria(pair.range[index], pair.criteria)) ? formulaNumber(value) || 0 : 0), 0);
    }
    case "AVERAGEIF": {
      const range = values([args[0]]);
      const criteria = scalar(1, "");
      const averageRange = args[2] ? values([args[2]]) : range;
      if (averageRange.length !== range.length) return "#VALUE!";
      const matched = averageRange.filter((_, index) => matchesCriteria(range[index], criteria));
      const error = matched.map(formulaErrorCode).find(Boolean);
      if (error) return error;
      const numbers = matched.filter((value) => value !== "" && value != null && Number.isFinite(Number(value))).map(Number);
      return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : "#DIV/0!";
    }
    case "AVERAGEIFS": {
      const averageRange = values([args[0]]);
      const pairs = [];
      for (let i = 1; i < args.length; i += 2) pairs.push({ range: values([args[i]]), criteria: scalar(i + 1, "") });
      if (pairs.some((pair) => pair.range.length !== averageRange.length)) return "#VALUE!";
      const matched = averageRange.filter((_, index) => pairs.every((pair) => matchesCriteria(pair.range[index], pair.criteria)));
      const error = matched.map(formulaErrorCode).find(Boolean);
      if (error) return error;
      const numbers = matched.filter((value) => value !== "" && value != null && Number.isFinite(Number(value))).map(Number);
      return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : "#DIV/0!";
    }
    case "SUMPRODUCT": {
      const matrices = args.map((arg) => formulaRangeMatrix(sheet, arg, context) || [[formulaScalar(sheet, arg, context)]]);
      if (!matrices.length) return 0;
      const shape = (matrix) => [matrix.length, matrix[0]?.length || 0];
      const [rows, cols] = shape(matrices[0]);
      if (matrices.some((matrix) => {
        const [matrixRows, matrixCols] = shape(matrix);
        return matrixRows !== rows || matrixCols !== cols || matrix.some((row) => row.length !== cols);
      })) return "#VALUE!";
      const arrays = matrices.map((matrix) => matrix.flat());
      const length = rows * cols;
      const error = arrays.flat().map(formulaErrorCode).find(Boolean);
      if (error) return error;
      return Array.from({ length }, (_, index) => arrays.reduce((product, array) => product * formulaNumber(array[index]), 1)).reduce((sum, value) => sum + value, 0);
    }
    case "INDEX": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      if (!matrix.length) return "#REF!";
      const rowIndex = Math.max(1, Math.floor(formulaNumber(scalar(1, 1))) || 1) - 1;
      const colIndex = Math.max(1, Math.floor(formulaNumber(scalar(2, 1))) || 1) - 1;
      return matrix[rowIndex]?.[colIndex] ?? "#REF!";
    }
    case "MATCH": {
      const lookup = scalar(0, "");
      const matchValues = formulaRangeMatrix(sheet, args[1], context) || [];
      const matchType = Math.sign(formulaNumber(scalar(2, 0)) || 0);
      return formulaMatchIndex(lookup, matchValues.flat(), matchType);
    }
    case "VLOOKUP": {
      const lookup = scalar(0, "");
      const matrix = formulaRangeMatrix(sheet, args[1], context) || [];
      const colIndex = Math.max(1, Number(scalar(2, 1)) || 1) - 1;
      const row = matrix.find((item) => formulaText(item[0]) === formulaText(lookup) || Number(item[0]) === Number(lookup));
      return row ? row[colIndex] ?? "#N/A" : "#N/A";
    }
    case "HLOOKUP": {
      const lookup = scalar(0, "");
      const matrix = formulaRangeMatrix(sheet, args[1], context) || [];
      const rowIndex = Math.floor(Number(scalar(2, 1)) || 1) - 1;
      if (rowIndex < 0 || rowIndex >= matrix.length) return "#REF!";
      const header = matrix[0] || [];
      const equals = (value) => formulaText(value).toLowerCase() === formulaText(lookup).toLowerCase()
        || (formulaText(value).trim() !== "" && formulaText(lookup).trim() !== "" && Number.isFinite(Number(value)) && Number(value) === Number(lookup));
      let columnIndex = header.findIndex(equals);
      if (columnIndex < 0 && formulaTruthy(scalar(3, true))) {
        for (let index = 0; index < header.length; index++) if (compareFormulaValues(header[index], "<=", lookup)) columnIndex = index;
      }
      return columnIndex >= 0 ? matrix[rowIndex]?.[columnIndex] ?? "#N/A" : "#N/A";
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
      pivots: sheet.pivotTables.toJSON(),
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
    sheet.pivotTables.items = [];
    for (const pivotData of sheetData.pivots || sheetData.pivotTables || []) {
      const pivot = sheet.pivotTables.add({ ...pivotData });
      pivot.id = pivotData.id || pivot.id;
    }
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
    const pivotParts = collectWorkbookPivotParts(workbook);
    const imageParts = collectWorkbookImageParts(workbook);
    const chartParts = collectWorkbookChartParts(workbook, imageParts);
    const threadParts = collectWorkbookThreadParts(workbook);
    const sharedStrings = collectWorkbookSharedStrings(workbook);
    const styleTable = collectWorkbookStyles(workbook);
    const workbookRels = workbookRelsXml(workbook.worksheets.items.length, threadParts.length > 0, sharedStrings.strings.length > 0, pivotParts);
    zip.file("[Content_Types].xml", xlsxContentTypes(workbook.worksheets.items.length, tableParts, imageParts, chartParts, threadParts, sharedStrings, pivotParts));
    zip.file("_rels/.rels", relsXml([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "xl/workbook.xml" }]));
    zip.file("xl/workbook.xml", workbookXml(workbook, pivotParts));
    zip.file("xl/_rels/workbook.xml.rels", workbookRels);
    zip.file("xl/styles.xml", xlsxStylesXml(styleTable));
    if (sharedStrings.strings.length) zip.file("xl/sharedStrings.xml", sharedStringsXml(sharedStrings));
    zip.file("customXml/open-office-artifact.json", JSON.stringify(workbookMetadata(workbook), null, 2));
    workbook.worksheets.items.forEach((sheet, index) => {
      const sheetTableParts = tableParts.filter((part) => part.sheetIndex === index);
      const sheetImageParts = imageParts.filter((part) => part.sheetIndex === index);
      const sheetChartParts = chartParts.filter((part) => part.sheetIndex === index);
      const sheetPivotParts = pivotParts.filter((part) => part.sheetIndex === index);
      const sheetThreadPart = threadParts.find((part) => part.sheetIndex === index);
      let nextRelIndex = sheetTableParts.length + 1;
      const drawingRelId = sheetImageParts.length || sheetChartParts.length ? `rId${nextRelIndex++}` : undefined;
      if (sheetThreadPart) sheetThreadPart.relId = `rId${nextRelIndex++}`;
      for (const part of sheetPivotParts) part.relId = `rId${nextRelIndex++}`;
      zip.file(`xl/worksheets/sheet${index + 1}.xml`, worksheetXml(sheet, sheetTableParts, drawingRelId, sharedStrings, styleTable));
      if (sheetTableParts.length || sheetImageParts.length || sheetChartParts.length || sheetThreadPart || sheetPivotParts.length) zip.file(`xl/worksheets/_rels/sheet${index + 1}.xml.rels`, worksheetRelsXml(sheetTableParts, drawingRelId ? { relId: drawingRelId, target: `../drawings/drawing${index + 1}.xml` } : undefined, sheetThreadPart, sheetPivotParts));
      if (sheetImageParts.length || sheetChartParts.length) {
        zip.file(`xl/drawings/drawing${index + 1}.xml`, drawingXml(sheetImageParts, sheetChartParts));
        zip.file(`xl/drawings/_rels/drawing${index + 1}.xml.rels`, drawingRelsXml(sheetImageParts, sheetChartParts));
      }
    });
    tableParts.forEach((part) => zip.file(`xl/tables/table${part.tablePartId}.xml`, tableXml(part.table, part.tablePartId)));
    pivotParts.forEach((part) => {
      zip.file(`xl/pivotTables/pivotTable${part.pivotPartId}.xml`, pivotTableXml(part));
      zip.file(`xl/pivotCache/pivotCacheDefinition${part.cachePartId}.xml`, pivotCacheDefinitionXml(part));
      zip.file(`xl/pivotCache/pivotCacheRecords${part.recordsPartId}.xml`, pivotCacheRecordsXml(part));
      zip.file(`xl/pivotCache/_rels/pivotCacheDefinition${part.cachePartId}.xml.rels`, pivotCacheDefinitionRelsXml(part));
    });
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

function collectWorkbookPivotParts(workbook) {
  const parts = [];
  let pivotPartId = 1;
  workbook.worksheets.items.forEach((sheet, sheetIndex) => {
    sheet.pivotTables.items.forEach((pivot, pivotIndex) => {
      parts.push({ sheet, sheetIndex, pivot, pivotIndex, pivotPartId: pivotPartId, cacheId: pivotPartId, cachePartId: pivotPartId, recordsPartId: pivotPartId, recordsRelId: "rId1", relId: undefined, cacheRelId: undefined });
      pivotPartId += 1;
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

function normalizeXlsxBorder(style = {}) {
  const raw = style.border || style.borders;
  if (!raw) return undefined;
  const base = raw.outside || raw.all || raw;
  const borderStyle = base.style || base.lineStyle || base.weight || "thin";
  const color = base.color || base.fill || base.borderColor || "#CBD5E1";
  return { style: borderStyle, color };
}

function normalizeXlsxAlignment(style = {}) {
  const raw = style.alignment || style.align || {};
  const horizontal = raw.horizontal || style.horizontalAlignment || style.textAlign;
  const vertical = raw.vertical || style.verticalAlignment;
  const wrapText = raw.wrapText ?? style.wrapText;
  const result = {};
  if (horizontal) result.horizontal = horizontal;
  if (vertical) result.vertical = vertical;
  if (wrapText != null) result.wrapText = Boolean(wrapText);
  return Object.keys(result).length ? result : undefined;
}

function normalizeXlsxStyle(style = {}) {
  const font = style.font || {};
  const alignment = normalizeXlsxAlignment(style);
  const border = normalizeXlsxBorder(style);
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
    alignment,
    border,
  };
}

function xlsxStyleKey(style = {}) {
  const normalized = normalizeXlsxStyle(style);
  if (!normalized.font.bold && !normalized.font.italic && !normalized.font.color && normalized.font.size === 11 && normalized.font.name === "Aptos" && !normalized.fill && !normalized.numberFormat && !normalized.alignment && !normalized.border) return "";
  return JSON.stringify(normalized);
}

function collectWorkbookStyles(workbook) {
  const styles = [{}];
  const indexByKey = new Map([["", 0]]);
  const dxfs = [];
  const dxfIndexByKey = new Map();
  const addStyle = (style = {}) => {
    const key = xlsxStyleKey(style || {});
    if (!indexByKey.has(key)) {
      indexByKey.set(key, styles.length);
      styles.push(normalizeXlsxStyle(style || {}));
    }
  };
  const addDxf = (style = {}) => {
    const key = xlsxStyleKey(style || {});
    if (!key || dxfIndexByKey.has(key)) return;
    dxfIndexByKey.set(key, dxfs.length);
    dxfs.push(normalizeXlsxStyle(style || {}));
  };
  for (const sheet of workbook.worksheets) {
    for (const [, cell] of sheet.store.entries()) addStyle(cell.style || {});
    for (const rule of sheet.conditionalFormattings.items) addDxf(rule.format || rule.rule?.format || {});
  }
  return { styles, indexByKey, dxfs, dxfIndexByKey };
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

function xlsxBorderXml(style = {}) {
  const border = normalizeXlsxStyle(style).border;
  if (!border) return `<border/>`;
  const color = normalizeXlsxColor(border.color, "CBD5E1");
  const lineStyle = attrEscape(border.style || "thin");
  const edge = (name) => `<${name} style="${lineStyle}"><color rgb="FF${color}"/></${name}>`;
  return `<border>${edge("left")}${edge("right")}${edge("top")}${edge("bottom")}<diagonal/></border>`;
}

function xlsxDxfXml(style = {}) {
  const normalized = normalizeXlsxStyle(style || {});
  const font = normalized.font || {};
  const fontXml = (font.bold || font.italic || font.color || font.size || font.name)
    ? `<font>${font.bold ? "<b/>" : ""}${font.italic ? "<i/>" : ""}${font.size ? `<sz val="${Number(font.size) || 11}"/>` : ""}${font.color ? `<color rgb="FF${normalizeXlsxColor(font.color, "000000")}"/>` : ""}${font.name ? `<name val="${attrEscape(font.name)}"/>` : ""}</font>`
    : "";
  const fillXml = normalized.fill ? xlsxFillXml(normalized) : "";
  const numFmtXml = normalized.numberFormat ? `<numFmt numFmtId="0" formatCode="${attrEscape(normalized.numberFormat)}"/>` : "";
  return `<dxf>${fontXml}${fillXml}${numFmtXml}</dxf>`;
}

function xlsxStylesXml(styleTable) {
  const styles = styleTable.styles || [{}];
  const dxfs = styleTable.dxfs || [];
  const customFormats = new Map();
  styles.forEach((style) => { if (style.numberFormat && !customFormats.has(style.numberFormat)) customFormats.set(style.numberFormat, 164 + customFormats.size); });
  const numFmts = customFormats.size ? `<numFmts count="${customFormats.size}">${[...customFormats.entries()].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${attrEscape(code)}"/>`).join("")}</numFmts>` : "";
  const fonts = styles.map((style, index) => index === 0 ? `<font><sz val="11"/><name val="Aptos"/></font>` : xlsxFontXml(style)).join("");
  const fills = [`<fill><patternFill patternType="none"/></fill>`, `<fill><patternFill patternType="gray125"/></fill>`, ...styles.slice(1).map(xlsxFillXml)].join("");
  const borders = [`<border/>`, ...styles.slice(1).map(xlsxBorderXml)].join("");
  const xfs = styles.map((style, index) => {
    if (index === 0) return `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`;
    const normalized = normalizeXlsxStyle(style);
    const numFmtId = normalized.numberFormat ? customFormats.get(normalized.numberFormat) : 0;
    const fillId = normalized.fill ? index + 1 : 0;
    const borderId = normalized.border ? index : 0;
    const alignmentXml = normalized.alignment ? `<alignment${normalized.alignment.horizontal ? ` horizontal="${attrEscape(normalized.alignment.horizontal)}"` : ""}${normalized.alignment.vertical ? ` vertical="${attrEscape(normalized.alignment.vertical)}"` : ""}${normalized.alignment.wrapText ? ` wrapText="1"` : ""}/>` : "";
    const attrs = `numFmtId="${numFmtId}" fontId="${index}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ` applyNumberFormat="1"` : ""} applyFont="1"${fillId ? ` applyFill="1"` : ""}${borderId ? ` applyBorder="1"` : ""}${normalized.alignment ? ` applyAlignment="1"` : ""}`;
    return alignmentXml ? `<xf ${attrs}>${alignmentXml}</xf>` : `<xf ${attrs}/>`;
  }).join("");
  const dxfXml = `<dxfs count="${dxfs.length}">${dxfs.map(xlsxDxfXml).join("")}</dxfs>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts}<fonts count="${styles.length}">${fonts}</fonts><fills count="${styles.length + 1}">${fills}</fills><borders count="${styles.length}">${borders}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${styles.length}">${xfs}</cellXfs>${dxfXml}</styleSheet>`;
}

function parseAttrs(attrs = "") {
  return Object.fromEntries([...String(attrs || "").matchAll(/\b([A-Za-z_:][\w:.-]*)="([^"]*)"/g)].map((match) => [match[1], decodeXml(match[2])]));
}

function parseXlsxFontStyle(body = "") {
  return {
    bold: /<b\b/.test(body),
    italic: /<i\b/.test(body),
    color: /<color[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(body)?.[1] ? `#${/<color[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(body)?.[1]}` : undefined,
    size: Number(/<sz[^>]*val="([^"]+)"/.exec(body)?.[1] || 11),
    name: /<name[^>]*val="([^"]+)"/.exec(body)?.[1] || "Aptos",
  };
}

function parseXlsxFillStyle(body = "") {
  const color = /<fgColor[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(body)?.[1];
  return color ? `#${color}` : undefined;
}

function parseXlsxBorderStyle(body = "") {
  const edge = /<(left|right|top|bottom)\b([^>]*)>([\s\S]*?)<\/\1>/.exec(body);
  if (!edge) return undefined;
  const attrs = parseAttrs(edge[2]);
  const color = /<color[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(edge[3])?.[1];
  return { style: attrs.style || "thin", ...(color ? { color: `#${color}` } : {}) };
}

function parseXlsxDxfXml(body = "") {
  const style = {};
  const fontBody = /<font>([\s\S]*?)<\/font>/.exec(body)?.[1];
  const fillBody = /<fill>([\s\S]*?)<\/fill>/.exec(body)?.[1];
  const numFmt = /<numFmt\b([^>]*)\/?>(?:<\/numFmt>)?/.exec(body)?.[1];
  if (fontBody != null) style.font = parseXlsxFontStyle(fontBody);
  const fill = fillBody != null ? parseXlsxFillStyle(fillBody) : undefined;
  if (fill) style.fill = fill;
  if (numFmt) style.numberFormat = parseAttrs(numFmt).formatCode;
  return style;
}

function parseXlsxStylesXml(xml = "") {
  const text = String(xml || "");
  const numFmtById = new Map([...text.matchAll(/<numFmt\b([^>]*)\/>/g)].map((match) => { const attrs = parseAttrs(match[1]); return [Number(attrs.numFmtId), attrs.formatCode]; }));
  const fontsBody = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(text)?.[1] || "";
  const fonts = [...fontsBody.matchAll(/<font>([\s\S]*?)<\/font>/g)].map((match) => parseXlsxFontStyle(match[1]));
  const fillsBody = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(text)?.[1] || "";
  const fills = [...fillsBody.matchAll(/<fill>([\s\S]*?)<\/fill>/g)].map((match) => parseXlsxFillStyle(match[1]));
  const bordersBody = /<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(text)?.[1] || "";
  const borders = [...bordersBody.matchAll(/<border>([\s\S]*?)<\/border>/g)].map((match) => parseXlsxBorderStyle(match[1]));
  const xfsBody = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(text)?.[1] || "";
  const styles = [...xfsBody.matchAll(/<xf\b([^>]*)(?:\/>|>([\s\S]*?)<\/xf>)/g)].map((match) => {
    const attrs = parseAttrs(match[1]);
    const body = match[2] || "";
    const font = fonts[Number(attrs.fontId || 0)] || {};
    const fill = fills[Number(attrs.fillId || 0)];
    const border = borders[Number(attrs.borderId || 0)];
    const numberFormat = numFmtById.get(Number(attrs.numFmtId || 0));
    const alignmentAttrs = parseAttrs(/<alignment\b([^>]*)\/>/.exec(body)?.[1] || "");
    const style = { font: { ...font } };
    if (fill) style.fill = fill;
    if (border) style.border = border;
    if (Object.keys(alignmentAttrs).length) style.alignment = { horizontal: alignmentAttrs.horizontal, vertical: alignmentAttrs.vertical, wrapText: alignmentAttrs.wrapText === "1" || alignmentAttrs.wrapText === "true" };
    if (numberFormat) style.numberFormat = numberFormat;
    return style;
  });
  const dxfsBody = /<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/.exec(text)?.[1] || "";
  styles.dxfs = [...dxfsBody.matchAll(/<dxf>([\s\S]*?)<\/dxf>/g)].map((match) => parseXlsxDxfXml(match[1]));
  return styles;
}

function xlsxContentTypes(sheetCount, tableParts = [], imageParts = [], chartParts = [], threadParts = [], sharedStrings = { strings: [] }, pivotParts = []) {
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
  const pivots = pivotParts.map((part) => `<Override PartName="/xl/pivotTables/pivotTable${part.pivotPartId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/><Override PartName="/xl/pivotCache/pivotCacheDefinition${part.cachePartId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/><Override PartName="/xl/pivotCache/pivotCacheRecords${part.recordsPartId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>`).join("");
  const threadedComments = threadParts.map((part) => `<Override PartName="/xl/threadedComments/threadedComment${part.threadPartId}.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>`).join("");
  const persons = threadParts.length ? `<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>` : "";
  const shared = sharedStrings.strings?.length ? `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="json" ContentType="application/json"/>${imageDefaults}<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${shared}${sheets}${tables}${charts}${pivots}${threadedComments}${persons}</Types>`;
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

function workbookXml(workbook, pivotParts = []) {
  const sheets = workbook.worksheets.items.map((sheet, i) => `<sheet name="${attrEscape(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("");
  const definedNames = workbook.definedNames.items.length
    ? `<definedNames>${workbook.definedNames.items.map((item) => {
      const localSheetId = item.scope ? workbook.worksheets.items.findIndex((sheet) => sheet.name === item.scope) : -1;
      return `<definedName name="${attrEscape(item.name)}"${localSheetId >= 0 ? ` localSheetId="${localSheetId}"` : ""}>${xmlEscape(item.refersTo)}</definedName>`;
    }).join("")}</definedNames>`
    : "";
  const pivotCaches = pivotParts.length ? `<pivotCaches>${pivotParts.map((part) => `<pivotCache cacheId="${part.cacheId}" r:id="${part.cacheRelId}"/>`).join("")}</pivotCaches>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets}</sheets>${definedNames}${pivotCaches}</workbook>`;
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

function workbookRelsXml(sheetCount, hasThreadedComments = false, hasSharedStrings = false, pivotParts = []) {
  const rels = Array.from({ length: sheetCount }, (_, i) => ({ id: `rId${i + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", target: `worksheets/sheet${i + 1}.xml` }));
  rels.push({ id: `rId${sheetCount + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", target: "styles.xml" });
  let nextId = sheetCount + 2;
  if (hasSharedStrings) rels.push({ id: `rId${nextId++}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", target: "sharedStrings.xml" });
  if (hasThreadedComments) rels.push({ id: `rId${nextId++}`, type: "http://schemas.microsoft.com/office/2017/10/relationships/person", target: "persons/person.xml" });
  for (const part of pivotParts) {
    part.cacheRelId = `rId${nextId++}`;
    rels.push({ id: part.cacheRelId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition", target: `pivotCache/pivotCacheDefinition${part.cachePartId}.xml` });
  }
  return relsXml(rels);
}

function worksheetRelsXml(tableParts, drawingRel, threadedPart, pivotParts = []) {
  const rels = tableParts.map((part) => ({
    id: part.relId,
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
    target: `../tables/table${part.tablePartId}.xml`,
  }));
  if (drawingRel) rels.push({ id: drawingRel.relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", target: drawingRel.target });
  if (threadedPart) rels.push({ id: threadedPart.relId, type: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment", target: `../threadedComments/threadedComment${threadedPart.threadPartId}.xml` });
  for (const part of pivotParts) rels.push({ id: part.relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable", target: `../pivotTables/pivotTable${part.pivotPartId}.xml` });
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

function pivotSourceHeaders(pivot) {
  return (pivot.sourceValues()[0] || []).map((value) => String(value ?? ""));
}

function pivotCacheDefinitionXml(part) {
  const { pivot } = part;
  const sourceSheet = pivot.sourceRange.sheetName || pivot.worksheet.name;
  const headers = pivotSourceHeaders(pivot);
  const cacheFields = headers.map((header, index) => {
    const values = [...new Set(pivot.sourceValues().slice(1).map((row) => row[index]).filter((value) => value != null && value !== ""))];
    const numeric = values.every((value) => Number.isFinite(Number(value)));
    const shared = values.length ? `<sharedItems${numeric ? ` containsNumber="1"` : ` containsString="1"`} count="${values.length}">${values.map((value) => numeric ? `<n v="${Number(value)}"/>` : `<s v="${attrEscape(value)}"/>`).join("")}</sharedItems>` : `<sharedItems count="0"/>`;
    return `<cacheField name="${attrEscape(header || `Field${index + 1}`)}" numFmtId="0">${shared}</cacheField>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="1" recordCount="${Math.max(0, pivot.sourceValues().length - 1)}"><cacheSource type="worksheet"><worksheetSource ref="${attrEscape(pivot.sourceRange.address)}" sheet="${attrEscape(sourceSheet)}"/></cacheSource><cacheFields count="${headers.length}">${cacheFields}</cacheFields></pivotCacheDefinition>`;
}

function pivotCacheDefinitionRelsXml(part) {
  return relsXml([{ id: part.recordsRelId || "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords", target: `pivotCacheRecords${part.recordsPartId}.xml` }]);
}

function pivotCacheRecordValueXml(value) {
  if (value == null || value === "") return "<m/>";
  const numeric = Number.isFinite(Number(value)) && String(value).trim() !== "";
  return numeric ? `<n v="${Number(value)}"/>` : `<s v="${attrEscape(value)}"/>`;
}

function pivotCacheRecordsXml(part) {
  const rows = part.pivot.sourceValues().slice(1);
  const records = rows.map((row) => `<r>${row.map(pivotCacheRecordValueXml).join("")}</r>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${rows.length}">${records}</pivotCacheRecords>`;
}

function pivotTableXml(part) {
  const { pivot } = part;
  const headers = pivotSourceHeaders(pivot);
  const target = safeRangeBounds(pivot.targetRange.address) || { top: 0, left: 0, bottom: Math.max(0, pivot.computedValues().length - 1), right: Math.max(0, pivot.computedValues()[0]?.length - 1) };
  const values = pivot.computedValues();
  const targetEnd = makeCellAddress(target.top + Math.max(0, values.length - 1), target.left + Math.max(0, (values[0]?.length || 1) - 1));
  const ref = `${makeCellAddress(target.top, target.left)}:${targetEnd}`;
  const rowIndexes = pivot.rowFields.map((field) => headers.indexOf(String(field))).filter((index) => index >= 0);
  const valueIndexes = pivot.valueFields.map((field) => headers.indexOf(String(field.field || field.name))).filter((index) => index >= 0);
  const pivotFields = headers.map((header, index) => `<pivotField${rowIndexes.includes(index) ? ` axis="axisRow"` : ""}${valueIndexes.includes(index) ? ` dataField="1"` : ""} showAll="0"><items count="1"><item t="default"/></items></pivotField>`).join("");
  const rowFields = rowIndexes.length ? `<rowFields count="${rowIndexes.length}">${rowIndexes.map((index) => `<field x="${index}"/>`).join("")}</rowFields>` : "";
  const dataFields = pivot.valueFields.length ? `<dataFields count="${pivot.valueFields.length}">${pivot.valueFields.map((field) => `<dataField name="${attrEscape(pivotValueLabel(field))}" fld="${Math.max(0, headers.indexOf(String(field.field || field.name)))}" subtotal="${attrEscape(field.summarizeBy || "sum")}"/>`).join("")}</dataFields>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${attrEscape(pivot.name)}" cacheId="${part.cacheId}" dataCaption="Values" updatedVersion="7" minRefreshableVersion="3"><location ref="${attrEscape(ref)}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="${headers.length}">${pivotFields}</pivotFields>${rowFields}${dataFields}</pivotTableDefinition>`;
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

function conditionalFormattingXml(sheet, styleTable = {}) {
  return sheet.conditionalFormattings.items.map((item, index) => {
    const ruleType = item.ruleType || "expression";
    if (String(ruleType).toLowerCase() === "colorscale") {
      const colors = colorScaleColors(item);
      const cfvos = colors.length >= 3 ? `<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>` : `<cfvo type="min"/><cfvo type="max"/>`;
      const colorXml = colors.map((color) => `<color rgb="FF${normalizeXlsxColor(color, "FFFFFF")}"/>`).join("");
      return `<conditionalFormatting sqref="${attrEscape(item.range || "A1")}"><cfRule type="colorScale" priority="${index + 1}"><colorScale>${cfvos}${colorXml}</colorScale></cfRule></conditionalFormatting>`;
    }
    const type = ruleType === "cellIs" || ruleType === "CellValue" ? "cellIs" : ruleType === "containsText" ? "containsText" : "expression";
    const operator = item.operator ? ` operator="${attrEscape(item.operator)}"` : "";
    const formula = Array.isArray(item.formula) ? item.formula[0] : item.formula || item.expression || "TRUE";
    const text = item.text ? ` text="${attrEscape(item.text)}"` : "";
    const dxfId = styleTable.dxfIndexByKey?.get(xlsxStyleKey(item.format || item.rule?.format || {}));
    const dxfAttr = dxfId != null ? ` dxfId="${dxfId}"` : "";
    return `<conditionalFormatting sqref="${attrEscape(item.range || "A1")}"><cfRule type="${attrEscape(type)}" priority="${index + 1}"${operator}${text}${dxfAttr}><formula>${xmlEscape(formula)}</formula></cfRule></conditionalFormatting>`;
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
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>${rowXml}</sheetData>${conditionalFormattingXml(sheet, styleTable)}${dataValidationsXml(sheet)}${drawingXml}${tablePartsXml}${sparklineXml}</worksheet>`;
}

function cellFormulaXml(address, cell) {
  if (!cell.formula) return "";
  const text = String(cell.formula).replace(/^=/, "");
  const formulaType = cell.formulaType || "";
  if (formulaType === "shared") {
    const attrs = [`t="shared"`];
    if (cell.sharedIndex != null) attrs.push(`si="${Number(cell.sharedIndex) || 0}"`);
    let includeBody = true;
    if (cell.sharedRef) {
      try {
        const bounds = parseRangeAddress(cell.sharedRef);
        includeBody = String(address).toUpperCase() === makeCellAddress(bounds.top, bounds.left);
        if (includeBody) attrs.push(`ref="${attrEscape(cell.sharedRef)}"`);
      } catch {
        includeBody = true;
        attrs.push(`ref="${attrEscape(cell.sharedRef)}"`);
      }
    }
    return `<f ${attrs.join(" ")}>${includeBody ? xmlEscape(text) : ""}</f>`;
  }
  if (formulaType === "array") {
    const ref = cell.arrayRef ? ` ref="${attrEscape(cell.arrayRef)}"` : "";
    return `<f t="array"${ref}>${xmlEscape(text)}</f>`;
  }
  if (cell.spillRange && !cell.spillError) {
    return `<f t="array" ref="${attrEscape(cell.spillRange)}">${xmlEscape(text)}</f>`;
  }
  const type = formulaType ? ` t="${attrEscape(formulaType)}"` : "";
  return `<f${type}>${xmlEscape(text)}</f>`;
}

function cellXml(address, cell, sharedStrings = { indexByText: new Map() }, styleTable = { styles: [{}], indexByKey: new Map([["", 0]]) }) {
  const f = cellFormulaXml(address, cell);
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

function parseXlsxSqref(value = "") {
  return String(value || "").trim().split(/\s+/).map((item) => item.trim()).filter(Boolean);
}

function parseDataValidationsXml(sheet, xml = "") {
  const text = String(xml || "");
  for (const match of text.matchAll(/<dataValidation\b([^>]*)>([\s\S]*?)<\/dataValidation>/g)) {
    const attrs = parseAttrs(match[1]);
    const body = match[2] || "";
    const formula1 = decodeXml(/<formula1[^>]*>([\s\S]*?)<\/formula1>/.exec(body)?.[1] || "");
    const formula2 = decodeXml(/<formula2[^>]*>([\s\S]*?)<\/formula2>/.exec(body)?.[1] || "");
    const listMatch = /^"([\s\S]*)"$/.exec(formula1);
    const rule = {
      type: attrs.type || "custom",
      operator: attrs.operator || undefined,
      formula1: formula1 || undefined,
      formula2: formula2 || undefined,
    };
    if (listMatch) rule.values = listMatch[1].split(",").map((item) => item.replaceAll('""', '"'));
    for (const range of parseXlsxSqref(attrs.sqref || attrs.ref || "A1")) sheet.dataValidations.add({ range, rule });
  }
}

function parseConditionalFormattingXml(sheet, xml = "", styles = []) {
  const text = String(xml || "");
  for (const block of text.matchAll(/<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g)) {
    const attrs = parseAttrs(block[1]);
    const ranges = parseXlsxSqref(attrs.sqref || "A1");
    for (const ruleMatch of String(block[2] || "").matchAll(/<cfRule\b([^>]*)>([\s\S]*?)<\/cfRule>/g)) {
      const ruleAttrs = parseAttrs(ruleMatch[1]);
      const type = ruleAttrs.type || "expression";
      if (type === "colorScale") {
        const colorScaleBody = /<colorScale>([\s\S]*?)<\/colorScale>/.exec(ruleMatch[2])?.[1] || "";
        const colors = [...colorScaleBody.matchAll(/<color[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"\/>/g)].map((m) => `#${m[1]}`);
        for (const range of ranges) sheet.conditionalFormattings.add({ range, ruleType: "colorScale", colors });
        continue;
      }
      const formula = decodeXml(/<formula[^>]*>([\s\S]*?)<\/formula>/.exec(ruleMatch[2])?.[1] || "");
      const ruleType = type === "cellIs" ? "cellIs" : type === "containsText" ? "containsText" : "expression";
      const format = styles?.dxfs?.[Number(ruleAttrs.dxfId)] || undefined;
      for (const range of ranges) {
        sheet.conditionalFormattings.add({
          range,
          ruleType,
          operator: ruleAttrs.operator || undefined,
          text: ruleAttrs.text || undefined,
          formula: formula || (type === "containsText" ? ruleAttrs.text : undefined),
          format,
          priority: ruleAttrs.priority ? Number(ruleAttrs.priority) : undefined,
        });
      }
    }
  }
}

function parseWorksheetXml(sheet, xml, options = {}) {
  const cellMatches = [...String(xml || "").matchAll(/<c\s+([^>]*)>([\s\S]*?)<\/c>/g)].map((match) => {
    const attrs = match[1];
    const body = match[2];
    const address = /r="([^"]+)"/.exec(attrs)?.[1];
    const formulaMatch = /<f\b([^>]*)>([\s\S]*?)<\/f>/.exec(body);
    return { attrs, body, address, formulaAttrs: parseAttrs(formulaMatch?.[1] || ""), formulaBody: formulaMatch ? decodeXml(formulaMatch[2] || "") : undefined };
  }).filter((item) => item.address);
  const sharedFormulas = new Map();
  for (const item of cellMatches) {
    if (item.formulaAttrs.t === "shared" && item.formulaAttrs.si != null && item.formulaBody) {
      sharedFormulas.set(String(item.formulaAttrs.si), { address: item.address, formula: item.formulaBody, ref: item.formulaAttrs.ref });
    }
  }
  for (const item of cellMatches) {
    const { attrs, body, address, formulaAttrs, formulaBody } = item;
    const cell = sheet.store.get(address);
    const styleIndex = Number(/\bs="([^"]+)"/.exec(attrs)?.[1] || 0);
    if (options.styles?.[styleIndex]) cell.style = { ...options.styles[styleIndex] };
    if (formulaBody !== undefined || formulaAttrs.t === "shared") {
      if (formulaAttrs.t === "shared") {
        const shared = sharedFormulas.get(String(formulaAttrs.si));
        const formulaText = formulaBody || (shared ? conditionalFormulaForCell(shared.formula, parseRangeAddress(shared.address), address) : "");
        if (formulaText) cell.formula = `=${formulaText}`;
        cell.formulaType = "shared";
        cell.sharedIndex = formulaAttrs.si != null ? Number(formulaAttrs.si) : undefined;
        cell.sharedRef = formulaAttrs.ref || shared?.ref;
      } else if (formulaAttrs.t === "array") {
        if (formulaBody) cell.formula = `=${formulaBody}`;
        cell.formulaType = "array";
        cell.arrayRef = formulaAttrs.ref;
      } else if (formulaBody) {
        cell.formula = `=${formulaBody}`;
        cell.formulaType = formulaAttrs.t || undefined;
      }
    }
    const text = /<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/is>/.exec(body)?.[1];
    const value = /<v[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1];
    const type = /\bt="([^"]+)"/.exec(attrs)?.[1];
    if (type === "s" && value !== undefined) cell.value = options.sharedStrings?.[Number(value)] ?? "";
    else if (text !== undefined) cell.value = decodeXml(text);
    else if (value !== undefined) cell.value = Number.isFinite(Number(value)) && type !== "str" ? Number(value) : decodeXml(value);
  }
  parseDataValidationsXml(sheet, xml);
  parseConditionalFormattingXml(sheet, xml, options.styles || []);
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

function svgInner(svg = "") {
  return String(svg || "").replace(/^<svg\b[^>]*>/i, "").replace(/<\/svg>\s*$/i, "");
}

function presentationMontageSvg(presentation, options = {}) {
  const slides = presentation.slides.items.length ? presentation.slides.items : [presentation.slides.add()];
  const gap = Number(options.gap ?? 24);
  const scale = Number(options.scale ?? 0.25);
  const columns = Math.max(1, Number(options.columns ?? 1) || 1);
  const slideW = Number(presentation.slideSize.width || 1280);
  const slideH = Number(presentation.slideSize.height || 720);
  const thumbW = slideW * scale;
  const thumbH = slideH * scale;
  const labelH = 20;
  const rows = Math.ceil(slides.length / columns);
  const width = Math.max(1, columns * thumbW + (columns + 1) * gap);
  const height = Math.max(1, rows * (thumbH + labelH) + (rows + 1) * gap);
  const thumbs = slides.map((slide, index) => {
    const col = index % columns;
    const row = Math.floor(index / columns);
    const x = gap + col * (thumbW + gap);
    const y = gap + row * (thumbH + labelH + gap);
    return `<g data-slide="${index + 1}"><rect x="${x - 1}" y="${y - 1}" width="${thumbW + 2}" height="${thumbH + 2}" fill="#ffffff" stroke="#94a3b8"/><g transform="translate(${x},${y}) scale(${scale})">${svgInner(slide.toSvg())}</g><text x="${x}" y="${y + thumbH + 15}" font-family="Arial" font-size="12" fill="#475569">Slide ${index + 1}${slide.title() ? ` — ${xmlEscape(slide.title()).slice(0, 80)}` : ""}</text></g>`;
  }).join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="#f8fafc"/>${thumbs}</svg>`;
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
        if (!/^(bar|line|pie)$/i.test(chart.chartType)) issues.push(verificationIssue("presentation", "unsupportedChartType", `Chart ${chart.name || chart.id} uses unsupported chart type ${chart.chartType}.`, { severity: "warning", slide: slide.index + 1, id: chart.id, chartType: chart.chartType }));
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
    if (options.format === "montage" || options.montage === true) return new FileBlob(presentationMontageSvg(this, options), { type: "image/svg+xml", metadata: { format: "montage", slides: this.slides.count, artifactKind: "presentation" } });
    const slide = options.slide || this.slides.getItem(0) || this.slides.add();
    if (options.format === "layout") return slide.export({ ...options, format: "layout" });
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

function slideLayoutSlice(slide, layout, options = {}) {
  const targets = inspectTargetTokens(options);
  const search = String(options.search || options.searchTerm || "").trim().toLowerCase();
  if (!targets.length && !search) return layout;
  const before = Math.max(0, Number(options.before ?? options.contextBefore ?? options.context ?? 0) || 0);
  const after = Math.max(0, Number(options.after ?? options.contextAfter ?? options.context ?? 0) || 0);
  const targetsSlide = targets.some((target) => target === slide.id || target === slide.name || target === String(slide.index + 1) || target === "slide");
  if (targetsSlide && !search) return { ...layout, slice: { targets, before, after, matchedElements: layout.elements.length, returnedElements: layout.elements.length } };
  const matches = [];
  layout.elements.forEach((element, index) => {
    const matchesSearch = !search || JSON.stringify(element).toLowerCase().includes(search);
    const matchesTarget = !targets.length || targetsSlide || inspectRecordMatchesTarget(element, targets);
    if (matchesSearch && matchesTarget) matches.push(index);
  });
  const keep = new Set();
  for (const index of matches) {
    for (let i = Math.max(0, index - before); i <= Math.min(layout.elements.length - 1, index + after); i += 1) keep.add(i);
  }
  const elements = layout.elements.filter((_, index) => keep.has(index));
  return { ...layout, elements, slice: { targets, search: search || undefined, before, after, matchedElements: matches.length, returnedElements: elements.length } };
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
      if (kinds.has("textRange") && shape.text.value) records.push(textRangeRecord(shape, { parentKind: "shape", record: { slide: this.index + 1, bbox: [shape.position.left, shape.position.top, shape.position.width, shape.position.height], bboxUnit: "px" } }));
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
  resolve(id) {
    if (String(id || "").endsWith("/text")) {
      const parentId = String(id).slice(0, -5);
      const shape = this.shapes.items.find((item) => item.id === parentId);
      if (shape) return createTextRange(shape, id, { parentKind: "shape" });
    }
    return [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.connectors.items, ...this.comments.items].find((element) => element.id === id);
  }

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
    if (options.format === "layout" || options.format === LAYOUT_MIME) return new FileBlob(JSON.stringify(this.layoutJson(options), null, 2), { type: LAYOUT_MIME, metadata: { artifactKind: "presentation", format: "layout", slide: this.index + 1, target: options.target ?? options.targetId ?? options.id ?? options.anchor, search: options.search ?? options.searchTerm } });
    return new FileBlob(this.toSvg(), { type: "image/svg+xml" });
  }

  layoutJson(options = {}) {
    const elements = [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.connectors.items].map((element) => {
      const record = element.layoutJson();
      const comments = this.comments.items.filter((comment) => comment.targetId === element.id);
      return {
        ...record,
        slide: this.index + 1,
        textRangeId: element.text?.value ? `${element.id}/text` : undefined,
        commentIds: comments.length ? comments.map((comment) => comment.id) : undefined,
        commentTextPreview: comments.length ? comments.flatMap((comment) => comment.comments.map((item) => item.text)).join("\n").slice(0, 300) : undefined,
      };
    });
    return slideLayoutSlice(this, {
      schema: "open-office-artifact.layout/v1",
      unit: "px",
      slide: { id: this.id, slide: this.index + 1, frame: this.frame, notes: this.speakerNotes.text || undefined },
      elements,
    }, options);
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

function createTextRange(parent, id, options = {}) {
  const getText = options.getText || (() => parent.text?.value ?? parent.text ?? parent.display ?? "");
  const setText = options.setText || ((value) => {
    if (parent.text && typeof parent.text.set === "function") parent.text.set(value);
    else if (parent.text && typeof parent.text === "object" && "value" in parent.text) parent.text.value = String(value ?? "");
    else if ("text" in parent) parent.text = String(value ?? "");
    else if ("display" in parent) parent.display = String(value ?? "");
  });
  return {
    kind: "textRange",
    id,
    parentId: parent.id,
    parentKind: options.parentKind || parent.kind || parent.constructor?.name,
    get text() { return getText(); },
    set text(value) { setText(value); },
    replace(search, replacement) { const next = String(getText()).replace(search, replacement); setText(next); return this; },
  };
}

function textRangeRecord(parent, options = {}) {
  const range = createTextRange(parent, `${parent.id}/text`, options);
  const text = String(range.text || "");
  return { kind: "textRange", id: range.id, parentId: parent.id, parentKind: range.parentKind, text, textPreview: text.slice(0, 300), textChars: text.length, ...(options.record || {}) };
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
    return pptxTextShapeXml(index, this.name || this.id, this.geometry, this.position, this.text.value, this.placeholder, { fill: this.fill, line: this.line, textStyle: this.text.style });
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

function normalizeChartSeries(seriesItems = []) {
  return (seriesItems || []).map((series, index) => ({
    name: series.name || `Series ${index + 1}`,
    values: (series.values || series.data || []).map((value) => value),
    categories: series.categories,
    color: series.color || series.fill || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4],
  }));
}

function normalizeChartAxes(config = {}) {
  const axes = config.axes || {};
  const axisTitles = config.axisTitles || {};
  return {
    category: { ...(axes.category || axes.x || {}), title: axes.category?.title || axes.x?.title || axisTitles.category || axisTitles.x || config.categoryAxisTitle || config.xAxisTitle || "" },
    value: { ...(axes.value || axes.y || {}), title: axes.value?.title || axes.y?.title || axisTitles.value || axisTitles.y || config.valueAxisTitle || config.yAxisTitle || "" },
  };
}

function normalizeChartLegend(config = {}, seriesLength = 0) {
  const raw = config.legend;
  if (raw === false || config.hasLegend === false) return { visible: false, position: "r" };
  if (typeof raw === "string") return { visible: true, position: raw };
  return { visible: raw?.visible ?? config.hasLegend ?? seriesLength > 1, position: raw?.position || config.legendPosition || "r" };
}

function normalizeChartDataLabels(config = {}) {
  const raw = config.dataLabels || config.labels || {};
  if (raw === true) return { showValue: true, showCategoryName: false, position: "bestFit" };
  if (raw === false) return { showValue: false, showCategoryName: false, position: "bestFit" };
  return { showValue: Boolean(raw.showValue ?? config.showValues), showCategoryName: Boolean(raw.showCategoryName ?? raw.showCategory ?? config.showCategoryLabels), position: raw.position || "bestFit" };
}

function pieSlicePath(cx, cy, radius, startAngle, endAngle) {
  const startX = cx + radius * Math.cos(startAngle);
  const startY = cy + radius * Math.sin(startAngle);
  const endX = cx + radius * Math.cos(endAngle);
  const endY = cy + radius * Math.sin(endAngle);
  const largeArc = endAngle - startAngle > Math.PI ? 1 : 0;
  return `M ${cx} ${cy} L ${startX} ${startY} A ${radius} ${radius} 0 ${largeArc} 1 ${endX} ${endY} Z`;
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
    this.series = normalizeChartSeries(config.series || []);
    this.axes = normalizeChartAxes(config);
    this.legend = normalizeChartLegend(config, this.series.length);
    this.hasLegend = this.legend.visible;
    this.dataLabels = normalizeChartDataLabels(config);
  }

  inspectRecord() {
    const p = this.position;
    return { kind: "chart", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, chartType: this.chartType, title: this.title, categories: this.categories, series: this.series.length, seriesDetails: this.series, axes: this.axes, legend: this.legend, dataLabels: this.dataLabels, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px" };
  }

  layoutJson() { return { kind: "chart", id: this.id, name: this.name, chartType: this.chartType, title: this.title, frame: this.position, categories: this.categories, series: this.series, axes: this.axes, legend: this.legend, dataLabels: this.dataLabels }; }

  toSvg() {
    const p = this.position;
    const categories = this.categories.length ? this.categories : Array.from({ length: Math.max(0, ...this.series.map((series) => series.values?.length || 0)) }, (_, index) => String(index + 1));
    const allValues = this.series.flatMap((series) => series.values || []).map((value) => Number(value) || 0);
    const max = Math.max(1, ...allValues);
    const plot = { left: p.left + 42, top: p.top + 42, width: Math.max(0, p.width - 72), height: Math.max(0, p.height - 82) };
    const title = `<text x="${p.left + 12}" y="${p.top + 24}" font-family="Arial" font-size="16" font-weight="700" fill="#0f172a">${xmlEscape(this.title || this.chartType)}</text>`;
    const axes = `<line x1="${plot.left}" y1="${plot.top + plot.height}" x2="${plot.left + plot.width}" y2="${plot.top + plot.height}" stroke="#94a3b8"/><line x1="${plot.left}" y1="${plot.top}" x2="${plot.left}" y2="${plot.top + plot.height}" stroke="#94a3b8"/>${this.axes.category.title ? `<text x="${plot.left + plot.width / 2 - 24}" y="${p.top + p.height - 4}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.category.title)}</text>` : ""}${this.axes.value.title ? `<text x="${p.left + 8}" y="${plot.top + 10}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.value.title)}</text>` : ""}`;
    const legend = this.legend.visible ? this.series.map((series, index) => `<rect x="${p.left + p.width - 82}" y="${p.top + 18 + index * 16}" width="10" height="10" fill="${xmlEscape(resolveColorToken(series.color, series.color))}"/><text x="${p.left + p.width - 68}" y="${p.top + 27 + index * 16}" font-family="Arial" font-size="10" fill="#334155">${xmlEscape(series.name)}</text>`).join("") : "";
    if (/^pie$/i.test(this.chartType)) {
      const series = this.series[0] || { values: [] };
      const values = (series.values || []).map((value) => Math.max(0, Number(value) || 0));
      const total = values.reduce((sum, value) => sum + value, 0) || 1;
      const radius = Math.max(8, Math.min(plot.width, plot.height) / 2);
      const cx = plot.left + plot.width / 2;
      const cy = plot.top + plot.height / 2;
      let angle = -Math.PI / 2;
      const slices = values.map((value, index) => {
        const next = angle + (value / total) * Math.PI * 2;
        const color = resolveColorToken(["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4], "#0ea5e9");
        const label = this.dataLabels.showValue ? `<text x="${cx + (radius + 8) * Math.cos((angle + next) / 2)}" y="${cy + (radius + 8) * Math.sin((angle + next) / 2)}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(categories[index] ?? value)}</text>` : "";
        const path = `<path d="${pieSlicePath(cx, cy, radius, angle, next)}" fill="${xmlEscape(color)}" stroke="#ffffff"/>${label}`;
        angle = next;
        return path;
      }).join("");
      const categoryLegend = categories.map((category, index) => `<rect x="${p.left + p.width - 82}" y="${p.top + 18 + index * 16}" width="10" height="10" fill="${xmlEscape(["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4])}"/><text x="${p.left + p.width - 68}" y="${p.top + 27 + index * 16}" font-family="Arial" font-size="10" fill="#334155">${xmlEscape(category)}</text>`).join("");
      return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>${title}${slices}${this.legend.visible ? categoryLegend : ""}`;
    }
    let body = "";
    if (/^line$/i.test(this.chartType)) {
      body = this.series.map((series) => {
        const points = (series.values || []).map((value, index) => {
          const x = plot.left + (categories.length <= 1 ? plot.width / 2 : (index / Math.max(1, categories.length - 1)) * plot.width);
          const y = plot.top + plot.height - ((Number(value) || 0) / max) * plot.height;
          return `${x},${y}`;
        }).join(" ");
        return `<polyline points="${points}" fill="none" stroke="${xmlEscape(resolveColorToken(series.color, series.color))}" stroke-width="2"/>`;
      }).join("");
    } else {
      const groupW = categories.length ? plot.width / categories.length : 0;
      const barW = Math.max(1, groupW / Math.max(1, this.series.length) * 0.72);
      body = this.series.flatMap((series, seriesIndex) => (series.values || []).map((value, index) => {
        const h = plot.height * (Number(value) || 0) / max;
        const x = plot.left + index * groupW + seriesIndex * barW + groupW * 0.12;
        const y = plot.top + plot.height - h;
        const label = this.dataLabels.showValue ? `<text x="${x}" y="${y - 4}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(value)}</text>` : "";
        return `<rect x="${x}" y="${y}" width="${Math.max(1, barW - 2)}" height="${h}" fill="${xmlEscape(resolveColorToken(series.color, series.color))}"/>${label}`;
      })).join("");
    }
    const labels = categories.map((category, index) => `<text x="${plot.left + index * (plot.width / Math.max(1, categories.length))}" y="${p.top + p.height - 18}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(category)}</text>`).join("");
    return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>${title}${axes}${body}${labels}${legend}`;
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
    if (this.dataUrl) return `<image href="${attrEscape(this.dataUrl)}" x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" preserveAspectRatio="xMidYMid meet"/>`;
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
  static async inspectPptx(blobOrBuffer, options = {}) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const files = Object.values(zip.files).filter((file) => !file.dir).sort((a, b) => a.name.localeCompare(b.name));
    const slideParts = files.filter((file) => /^ppt\/slides\/slide\d+\.xml$/.test(file.name)).length;
    const records = [{ kind: "pptxPackage", parts: files.length, slides: slideParts }];
    const includeText = Boolean(options.includeText || options.preview || options.includeXml);
    const maxPreviewChars = Math.max(0, Number(options.maxPreviewChars ?? 400) || 0);
    for (const file of files) {
      const partBytes = await file.async("uint8array");
      const contentType = file.name.endsWith(".rels") ? "application/vnd.openxmlformats-package.relationships+xml" : file.name.endsWith(".xml") ? "application/xml" : file.name.endsWith(".json") ? "application/json" : file.name.endsWith(".png") ? "image/png" : file.name.endsWith(".jpg") || file.name.endsWith(".jpeg") ? "image/jpeg" : "application/octet-stream";
      const record = { kind: "pptxPart", path: file.name, size: partBytes.byteLength, contentType };
      if (includeText && /\.(xml|json|rels)$/i.test(file.name)) record.textPreview = decoder.decode(partBytes).slice(0, maxPreviewChars);
      records.push(record);
    }
    return { parts: records.filter((record) => record.kind === "pptxPart"), records, ...ndjson(records, options.maxChars ?? Infinity) };
  }

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
  const placeholders = layout.placeholders.map((placeholder, index) => pptxTextShapeXml(index, placeholder.name, "rect", placeholder.position, placeholder.text || "", { type: placeholder.type, idx: index + 1, required: placeholder.required }, { fill: "transparent", line: { fill: "transparent", width: 0 }, textStyle: placeholder.style })).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="${attrEscape(layout.type)}" preserve="1"><p:cSld name="${attrEscape(layout.name)}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${placeholders}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function presentationXml(presentation) {
  const ids = presentation.slides.items.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="12192000" cy="6858000"/><p:sldIdLst>${ids}</p:sldIdLst></p:presentation>`;
}

function pptxDrawingFillXml(value, fallback = "transparent") {
  const raw = typeof value === "object" ? value?.color || value?.fill : value;
  const resolved = resolveColorToken(raw || fallback, raw || fallback);
  if (!resolved || resolved === "transparent" || resolved === "none") return "<a:noFill/>";
  return `<a:solidFill><a:srgbClr val="${attrEscape(pptxColorValue(resolved, fallback))}"/></a:solidFill>`;
}

function pptxDrawingLineXml(line = {}) {
  const width = Math.max(0, Number(line.width ?? 1));
  const raw = line.fill || line.color || (width > 0 ? "#334155" : "transparent");
  return `<a:ln w="${Math.round(width * 12700)}">${pptxDrawingFillXml(raw)}</a:ln>`;
}

function pptxTextRunPropertiesXml(style = {}, options = {}) {
  const fontSize = Math.max(1, Number(style.fontSize || options.fontSize || 24));
  const color = resolveColorToken(style.color || options.color || "#0f172a", style.color || options.color || "#0f172a");
  const attrs = ` lang="en-US" sz="${Math.round(fontSize * 75)}"${style.bold ? ' b="1"' : ""}${style.italic ? ' i="1"' : ""}`;
  const typeface = style.fontFamily || style.typeface || options.fontFamily;
  return `<a:rPr${attrs}>${pptxDrawingFillXml(color, "#0f172a")}${typeface ? `<a:latin typeface="${attrEscape(typeface)}"/>` : ""}</a:rPr>`;
}

function pptxTextShapeXml(index, name, geometry, position, text = "", placeholder, options = {}) {
  const p = position;
  const x = Math.round(p.left * 9525), y = Math.round(p.top * 9525), cx = Math.round(p.width * 9525), cy = Math.round(p.height * 9525);
  const textStyle = options.textStyle || {};
  const paragraphs = String(text || "").split(/\r?\n/).map((line) => `<a:p><a:pPr algn="${attrEscape(textStyle.alignment === "center" ? "ctr" : textStyle.alignment === "right" ? "r" : "l")}"/><a:r>${pptxTextRunPropertiesXml(textStyle)}<a:t>${xmlEscape(line)}</a:t></a:r><a:endParaRPr lang="en-US" sz="${Math.round(Math.max(1, Number(textStyle.fontSize || 24)) * 75)}"/></a:p>`).join("");
  const ph = placeholder ? `<p:ph type="${attrEscape(placeholder.type || "body")}" idx="${Number(placeholder.idx || 1)}"/>` : "";
  return `<p:sp><p:nvSpPr><p:cNvPr id="${index + 2}" name="${attrEscape(name)}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${attrEscape(geometry === "textbox" ? "rect" : geometry)}"><a:avLst/></a:prstGeom>${pptxDrawingFillXml(options.fill)}${pptxDrawingLineXml(options.line)}</p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/><a:lstStyle/>${paragraphs || "<a:p/>"}</p:txBody></p:sp>`;
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
    const targetName = slide.resolve(thread.targetId)?.name || "";
    return `<p:cm authorId="0" dt="${attrEscape(comment.created || thread.created)}" idx="${idx}" ooa:threadId="${attrEscape(thread.id)}" ooa:targetId="${attrEscape(thread.targetId || "")}" ooa:targetName="${attrEscape(targetName)}" ooa:resolved="${thread.resolved ? 1 : 0}" xmlns:ooa="urn:open-office-artifact"><p:pos x="${100 + commentIndex * 24}" y="${100 + threadIndex * 32}"/><p:text>${xmlEscape(comment.text)}</p:text></p:cm>`;
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
    const originalTargetId = /\btargetId="([^"]*)"/.exec(attrs)?.[1] || /\booa:targetId="([^"]*)"/.exec(attrs)?.[1] || undefined;
    const targetName = decodeXml(/\btargetName="([^"]*)"/.exec(attrs)?.[1] || /\booa:targetName="([^"]*)"/.exec(attrs)?.[1] || "");
    const namedTarget = targetName ? [...slide.shapes.items, ...slide.tables.items, ...slide.charts.items, ...slide.images.items, ...slide.connectors.items].find((item) => item.name === targetName) : undefined;
    const targetId = slide.resolve(originalTargetId)?.id || namedTarget?.id || originalTargetId;
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
  const borderXml = `<a:lnL w="12700"><a:solidFill><a:srgbClr val="B8BCC4"/></a:solidFill></a:lnL><a:lnR w="12700"><a:solidFill><a:srgbClr val="B8BCC4"/></a:solidFill></a:lnR><a:lnT w="12700"><a:solidFill><a:srgbClr val="B8BCC4"/></a:solidFill></a:lnT><a:lnB w="12700"><a:solidFill><a:srgbClr val="B8BCC4"/></a:solidFill></a:lnB>`;
  const rows = Array.from({ length: Math.max(1, table.rows) }, (_, rowIndex) => {
    const header = table.styleOptions.headerRow && rowIndex === 0;
    const cells = Array.from({ length: Math.max(1, table.columns) }, (_, colIndex) => {
      const cellText = table.values[rowIndex]?.[colIndex] ?? "";
      const textStyle = { fontSize: table.styleOptions.fontSize || 18, bold: header, color: header ? "#000000" : "#0f172a", fontFamily: table.styleOptions.fontFamily };
      const fill = header ? "EDEDED" : rowIndex % 2 && table.styleOptions.bandedRows ? "F8FAFC" : "FFFFFF";
      return `<a:tc><a:txBody><a:bodyPr wrap="square" lIns="76200" tIns="45720" rIns="76200" bIns="45720" anchor="ctr"/><a:lstStyle/><a:p><a:r>${pptxTextRunPropertiesXml(textStyle)}<a:t>${xmlEscape(cellText)}</a:t></a:r><a:endParaRPr lang="en-US" sz="${Math.round(textStyle.fontSize * 75)}"/></a:p></a:txBody><a:tcPr marL="76200" marR="76200" marT="45720" marB="45720"><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>${borderXml}</a:tcPr></a:tc>`;
    }).join("");
    return `<a:tr h="${rowHeight}">${cells}</a:tr>`;
  }).join("");
  const firstRow = table.styleOptions.headerRow ? 1 : 0;
  const bandRow = table.styleOptions.bandedRows ? 1 : 0;
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${index + 2}" name="${attrEscape(table.name || table.id)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="${firstRow}" bandRow="${bandRow}"/><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

function pptxChartFrameXml(index, name, position, relId) {
  const p = position;
  const x = Math.round(p.left * 9525), y = Math.round(p.top * 9525), cx = Math.round(p.width * 9525), cy = Math.round(p.height * 9525);
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${index + 2}" name="${attrEscape(name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${relId}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

function pptxChartTextTitleXml(text = "") {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(text)}</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
}

function pptxChartXml(chart) {
  const chartElementName = chart.chartType === "line" ? "lineChart" : chart.chartType === "pie" ? "pieChart" : "barChart";
  const grouping = chart.chartType === "line" ? "<c:grouping val=\"standard\"/>" : chart.chartType === "pie" ? "" : "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>";
  const dataLabels = chart.dataLabels || {};
  const dataLabelsXml = dataLabels.showValue || dataLabels.showCategoryName ? `<c:dLbls><c:showVal val=\"${dataLabels.showValue ? 1 : 0}\"/><c:showCatName val=\"${dataLabels.showCategoryName ? 1 : 0}\"/><c:showLegendKey val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls>` : "";
  const seriesXml = (chart.series.length ? chart.series : [{ name: chart.title || "Series", values: [] }]).map((series, index) => {
    const values = series.values || [];
    const categories = series.categories || chart.categories || values.map((_, i) => String(i + 1));
    const catPts = categories.map((category, pointIndex) => `<c:pt idx=\"${pointIndex}\"><c:v>${xmlEscape(category)}</c:v></c:pt>`).join("");
    const valPts = values.map((value, pointIndex) => `<c:pt idx=\"${pointIndex}\"><c:v>${Number(value) || 0}</c:v></c:pt>`).join("");
    const color = String(series.color || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4]).replace(/^#/, "").slice(0, 6).padEnd(6, "0");
    return `<c:ser><c:idx val=\"${index}\"/><c:order val=\"${index}\"/><c:tx><c:v>${xmlEscape(series.name || `Series ${index + 1}`)}</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val=\"${attrEscape(color)}\"/></a:solidFill></c:spPr><c:cat><c:strLit><c:ptCount val=\"${categories.length}\"/>${catPts}</c:strLit></c:cat><c:val><c:numLit><c:ptCount val=\"${values.length}\"/>${valPts}</c:numLit></c:val></c:ser>`;
  }).join("");
  const categoryAxisTitle = chart.axes?.category?.title ? pptxChartTextTitleXml(chart.axes.category.title) : "";
  const valueAxisTitle = chart.axes?.value?.title ? pptxChartTextTitleXml(chart.axes.value.title) : "";
  const legendXml = chart.legend?.visible || chart.hasLegend ? `<c:legend><c:legendPos val=\"${attrEscape(chart.legend?.position || "r")}\"/><c:layout/></c:legend>` : "";
  const chartBody = chart.chartType === "pie"
    ? `<c:${chartElementName}>${seriesXml}${dataLabelsXml}</c:${chartElementName}>`
    : `<c:${chartElementName}>${grouping}${seriesXml}${dataLabelsXml}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:${chartElementName}><c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"b\"/>${categoryAxisTitle}<c:crossAx val=\"2\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"l\"/>${valueAxisTitle}<c:crossAx val=\"1\"/></c:valAx>`;
  return `<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><c:chart>${pptxChartTextTitleXml(chart.title || chart.chartType)}<c:plotArea><c:layout/>${chartBody}</c:plotArea>${legendXml}<c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>`;
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
  slide.tables.add({ name, position: pptxFrameFromXml(part, { left: 0, top: 0, width: 320, height: 160 }), values, rows: values.length, columns: Math.max(1, ...values.map((row) => row.length)), styleOptions: { headerRow: /<a:tblPr\b[^>]*firstRow="1"/.test(part), bandedRows: /<a:tblPr\b[^>]*bandRow="1"/.test(part) } });
}

async function parsePptxChartGraphic(slide, part, context) {
  const relId = /<c:chart[^>]*r:id="([^"]+)"/.exec(part)?.[1];
  const target = pptxRelationshipTarget(context.rels, relId);
  const chartXml = target ? await context.zip.file(target)?.async("text") : "";
  const chartType = /<c:pieChart>/.test(chartXml) ? "pie" : /<c:lineChart>/.test(chartXml) ? "line" : "bar";
  const title = decodeXml(/<c:chart>[\s\S]*?<c:title>[\s\S]*?<a:t>([\s\S]*?)<\/a:t>[\s\S]*?<\/c:title>/.exec(chartXml)?.[1] || "");
  const legendMatch = /<c:legend>[\s\S]*?<c:legendPos[^>]*val="([^"]+)"/.exec(chartXml);
  const dataLabelsBlock = /<c:dLbls>([\s\S]*?)<\/c:dLbls>/.exec(chartXml)?.[1] || "";
  const dataLabels = { showValue: /<c:showVal[^>]*val="1"/.test(dataLabelsBlock), showCategoryName: /<c:showCatName[^>]*val="1"/.test(dataLabelsBlock), position: "bestFit" };
  const catAxisTitle = decodeXml(/<c:catAx>[\s\S]*?<c:title>[\s\S]*?<a:t>([\s\S]*?)<\/a:t>[\s\S]*?<\/c:title>/.exec(chartXml)?.[1] || "");
  const valAxisTitle = decodeXml(/<c:valAx>[\s\S]*?<c:title>[\s\S]*?<a:t>([\s\S]*?)<\/a:t>[\s\S]*?<\/c:title>/.exec(chartXml)?.[1] || "");
  const series = [...String(chartXml || "").matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((seriesMatch, index) => {
    const seriesXml = seriesMatch[0];
    const name = decodeXml(/<c:tx>[\s\S]*?<c:v>([\s\S]*?)<\/c:v>[\s\S]*?<\/c:tx>/.exec(seriesXml)?.[1] || `Series ${index + 1}`);
    const color = /<a:srgbClr[^>]*val="([A-Fa-f0-9]{6})"/.exec(seriesXml)?.[1];
    const categories = [...( /<c:cat>([\s\S]*?)<\/c:cat>/.exec(seriesXml)?.[1] || "").matchAll(/<c:v>([\s\S]*?)<\/c:v>/g)].map((m) => decodeXml(m[1]));
    const values = [...( /<c:val>([\s\S]*?)<\/c:val>/.exec(seriesXml)?.[1] || "").matchAll(/<c:v>([\s\S]*?)<\/c:v>/g)].map((m) => Number(decodeXml(m[1])) || 0);
    return { name, values, categories, color: color ? `#${color}` : undefined };
  });
  const name = decodeXml(/<p:cNvPr[^>]*name="([^"]*)"/.exec(part)?.[1] || title || "chart");
  slide.charts.add(chartType, { name, title, position: pptxFrameFromXml(part, { left: 0, top: 0, width: 360, height: 220 }), categories: series[0]?.categories || [], series, axes: { category: { title: catAxisTitle }, value: { title: valAxisTitle } }, legend: { visible: Boolean(legendMatch), position: legendMatch?.[1] || "r" }, dataLabels });
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
    const spPr = /<p:spPr>([\s\S]*?)<\/p:spPr>/.exec(part)?.[1] || "";
    const fill = /<a:solidFill>[\s\S]*?<a:srgbClr[^>]*val="([A-Fa-f0-9]{6})"/.exec(spPr)?.[1];
    const lineBlock = /<a:ln\b([^>]*)>([\s\S]*?)<\/a:ln>/.exec(spPr);
    const lineColor = /<a:solidFill>[\s\S]*?<a:srgbClr[^>]*val="([A-Fa-f0-9]{6})"/.exec(lineBlock?.[2] || "")?.[1];
    const lineWidth = Number(/\bw="(\d+)"/.exec(lineBlock?.[1] || "")?.[1] || 0) / 12700;
    const geometry = /<a:prstGeom[^>]*prst="([^"]+)"/.exec(spPr)?.[1] || "rect";
    const rPr = /<a:rPr\b([^>]*)>([\s\S]*?)<\/a:rPr>/.exec(part);
    const fontSize = Number(/\bsz="(\d+)"/.exec(rPr?.[1] || "")?.[1] || 0) / 75;
    const textColor = /<a:solidFill>[\s\S]*?<a:srgbClr[^>]*val="([A-Fa-f0-9]{6})"/.exec(rPr?.[2] || "")?.[1];
    const fontFamily = decodeXml(/<a:latin[^>]*typeface="([^"]*)"/.exec(rPr?.[2] || "")?.[1] || "");
    const shape = slide.shapes.add({ name, geometry, position: pptxFrameFromXml(part), placeholder, fill: fill ? `#${fill}` : "transparent", line: lineColor && lineWidth > 0 ? { fill: `#${lineColor}`, width: lineWidth } : { fill: "transparent", width: 0 } });
    shape.text = text;
    shape.text.style = { ...(fontSize ? { fontSize } : {}), ...(/\bb="1"/.test(rPr?.[1] || "") ? { bold: true } : {}), ...(/\bi="1"/.test(rPr?.[1] || "") ? { italic: true } : {}), ...(textColor ? { color: `#${textColor}` } : {}), ...(fontFamily ? { fontFamily } : {}) };
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
  effective(id, seen = new Set()) {
    const style = this.get(id) || this.get("Normal");
    if (!style) return undefined;
    const parentId = style.basedOn || style.parent || style.extends;
    if (!parentId || seen.has(style.id)) return { ...style };
    seen.add(style.id);
    const parent = this.effective(parentId, seen) || {};
    return { ...parent, ...style, basedOn: parentId };
  }
  values() { return [...this.items.values()]; }
}

class DocumentTableCell {
  constructor(table, row, column) { this.table = table; this.row = row; this.column = column; }
  get value() { return this.table.values[this.row]?.[this.column] ?? ""; }
  set value(value) { this.table.ensureCell(this.row, this.column); this.table.values[this.row][this.column] = value; }
}

function documentTableDefaultColumnWidths(columns, widthDxa) {
  const count = Math.max(1, Number(columns) || 1);
  const total = Math.max(count, Math.round(Number(widthDxa) || 9360));
  const base = Math.floor(total / count);
  return Array.from({ length: count }, (_, index) => base + (index < total - base * count ? 1 : 0));
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
    this.widthDxa = Math.round(Number(config.widthDxa ?? 9360));
    this.indentDxa = Math.round(Number(config.indentDxa ?? 120));
    this.columnWidthsDxa = Array.isArray(config.columnWidthsDxa)
      ? config.columnWidthsDxa.map((value) => Math.round(Number(value)))
      : documentTableDefaultColumnWidths(this.columns, this.widthDxa);
    this.cellMarginsDxa = {
      top: Math.round(Number(config.cellMarginsDxa?.top ?? 80)),
      bottom: Math.round(Number(config.cellMarginsDxa?.bottom ?? 80)),
      start: Math.round(Number(config.cellMarginsDxa?.start ?? config.cellMarginsDxa?.left ?? 120)),
      end: Math.round(Number(config.cellMarginsDxa?.end ?? config.cellMarginsDxa?.right ?? 120)),
    };
    this.borderColor = String(config.borderColor || "D9D9D9").replace(/^#/, "").toUpperCase();
    this.borderSize = Math.round(Number(config.borderSize ?? 4));
    this.headerFill = String(config.headerFill || "F2F4F7").replace(/^#/, "").toUpperCase();
  }

  ensureCell(row, column) { while (this.values.length <= row) this.values.push([]); while (this.values[row].length <= column) this.values[row].push(""); this.rows = this.values.length; this.columns = Math.max(this.columns, column + 1); }
  getCell(row, column) { return new DocumentTableCell(this, row, column); }
  inspectRecord(index) { return { kind: "table", id: this.id, index, name: this.name || undefined, rows: this.rows, cols: this.columns, styleId: this.styleId, widthDxa: this.widthDxa, indentDxa: this.indentDxa, columnWidthsDxa: this.columnWidthsDxa, cellMarginsDxa: this.cellMarginsDxa, borderColor: this.borderColor, borderSize: this.borderSize, headerFill: this.headerFill, values: this.values }; }
  toProto() { return { kind: "table", id: this.id, name: this.name, styleId: this.styleId, widthDxa: this.widthDxa, indentDxa: this.indentDxa, columnWidthsDxa: this.columnWidthsDxa, cellMarginsDxa: this.cellMarginsDxa, borderColor: this.borderColor, borderSize: this.borderSize, headerFill: this.headerFill, values: this.values }; }
}

function normalizeDocumentRuns(text, config = {}) {
  const runs = (config.runs || config.textRuns || []).map((run) => ({ text: String(run.text ?? run.value ?? ""), style: { ...(run.style || run.textStyle || {}) } })).filter((run) => run.text.length > 0);
  if (runs.length) return runs;
  const rawText = String(text ?? "");
  return rawText ? [{ text: rawText, style: {} }] : [];
}

class DocumentParagraphBlock {
  constructor(document, text, config = {}) {
    this.document = document;
    this.kind = "paragraph";
    this.id = config.id || aid("dp");
    this.runs = normalizeDocumentRuns(text, config);
    this.text = this.runs.map((run) => run.text).join("") || String(text ?? "");
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "paragraph", id: this.id, index, name: this.name || undefined, styleId: this.styleId, text: this.text, textChars: this.text.length, runs: this.runs.length > 1 ? this.runs : undefined }; }
  toProto() { return { kind: "paragraph", id: this.id, name: this.name, styleId: this.styleId, text: this.text, runs: this.runs.length > 1 ? this.runs : undefined }; }
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
  const style = document.styles.effective(block.styleId) || document.styles.get("Normal") || {};
  const fontSize = Math.max(10, (style.fontSize || 22) / 2);
  const text = block.text || block.display || "";
  const charsPerLine = Math.max(8, Math.floor((pageWidth - margin * 2) / (fontSize * 0.55)));
  const lines = String(text).split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);
  return Math.max(20, lines * fontSize * 1.6);
}

function documentBlockLayoutText(block) {
  if (block.kind === "table") return block.values.map((row) => row.map((cell) => String(cell ?? "")).join(" ")).join(" ");
  if (block.kind === "section") return `section break: ${block.breakType || ""} ${block.orientation || ""}`.trim();
  return String(block.text ?? block.display ?? block.alt ?? block.prompt ?? block.uri ?? block.name ?? "");
}

function documentLayoutSlice(document, layout, options = {}) {
  const targets = inspectTargetTokens(options);
  const search = String(options.search || options.searchTerm || "").trim().toLowerCase();
  if (!targets.length && !search) return layout;
  const before = Math.max(0, Number(options.before ?? options.contextBefore ?? options.context ?? 0) || 0);
  const after = Math.max(0, Number(options.after ?? options.contextAfter ?? options.context ?? 0) || 0);
  const targetsDocument = targets.some((target) => target === document.id || target === document.name);
  const matchedPages = new Set(layout.pages.filter((pageRecord) => inspectRecordMatchesTarget(pageRecord, targets)).map((pageRecord) => pageRecord.page));
  const matches = [];
  layout.elements.forEach((element, index) => {
    const matchesSearch = !search || JSON.stringify(element).toLowerCase().includes(search);
    const matchesTarget = !targets.length || targetsDocument || matchedPages.has(element.page) || inspectRecordMatchesTarget(element, targets);
    if (matchesSearch && matchesTarget) matches.push(index);
  });
  const keep = new Set();
  for (const index of matches) {
    for (let i = Math.max(0, index - before); i <= Math.min(layout.elements.length - 1, index + after); i += 1) keep.add(i);
  }
  const elements = layout.elements.filter((_, index) => keep.has(index));
  const referencedPages = new Set(elements.map((element) => element.page));
  const pages = layout.pages.filter((pageRecord) => referencedPages.has(pageRecord.page));
  return { ...layout, pages, elements, slice: { targets, search: search || undefined, before, after, matchedElements: matches.length, returnedElements: elements.length } };
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
    if (!pages.find((item) => item.page === page)) pages.push({ id: `${document.id}/page/${page}`, page, width: pageWidth, height: pageHeight, margin, headers: document.headers.map((header) => header.id), footers: document.footers.map((footer) => footer.id) });
  };
  ensurePage();
  for (const block of document.blocks) {
    const height = documentBlockHeight(document, block, pageWidth, margin);
    if (y + height > pageHeight - margin && y > margin) { page += 1; y = margin; ensurePage(); }
    const textPreview = documentBlockLayoutText(block).slice(0, 120);
    const comments = document.comments.filter((comment) => comment.targetId === block.id).map((comment) => comment.id);
    const effectiveStyle = block.styleId ? document.styles.effective(block.styleId) : undefined;
    elements.push({ kind: "layoutElement", id: block.id, layoutId: `${block.id}/layout`, blockKind: block.kind, name: block.name || undefined, textRangeId: ("text" in block || "display" in block) ? `${block.id}/text` : undefined, commentIds: comments.length ? comments : undefined, page, bbox: [margin, y, pageWidth - margin * 2, height], styleId: block.styleId, effectiveStyle, textPreview });
    y += height;
    if (block.kind === "section" && block.breakType === "nextPage") { page += 1; y = margin; ensurePage(); }
  }
  return documentLayoutSlice(document, { schema: "open-office-artifact.document-layout/v1", unit: "px", document: { id: document.id, name: document.name, designPreset: document.designPreset }, pages, elements }, options);
}

function documentLayoutRecords(document, options = {}) {
  const layout = documentLayoutJson(document, options);
  return [
    { kind: "layout", id: `${document.id}/layout`, pages: layout.pages.length, elements: layout.elements.length, designPreset: document.designPreset },
    ...layout.pages.map((page) => ({ kind: "page", id: `${document.id}/page/${page.page}`, ...page })),
    ...layout.elements,
  ];
}

function documentTextParent(document, parentId) {
  return [...document.blocks, ...document.headers, ...document.footers, ...document.comments].find((item) => item.id === parentId);
}

function documentTextRange(document, id) {
  const parentId = String(id || "").endsWith("/text") ? String(id).slice(0, -5) : undefined;
  const parent = parentId ? documentTextParent(document, parentId) : undefined;
  if (!parent) return undefined;
  return createTextRange(parent, id, {
    parentKind: parent.kind,
    getText: () => parent.text ?? parent.display ?? "",
    setText: (value) => {
      if (parent.kind === "field" || ("display" in parent && !("text" in parent))) parent.display = String(value ?? "");
      else {
        parent.text = String(value ?? "");
        if (parent.kind === "paragraph") parent.runs = parent.text ? [{ text: parent.text, style: {} }] : [];
      }
    },
  });
}

function documentInspectRecord(document, block, index) {
  const record = block.inspectRecord(index);
  if (record.styleId) record.effectiveStyle = document.styles.effective(record.styleId);
  return record;
}

function documentTextRangeRecords(document) {
  const parents = [...document.blocks, ...document.headers, ...document.footers, ...document.comments].filter((item) => item && ("text" in item || "display" in item));
  return parents.map((parent, index) => textRangeRecord(parent, {
    parentKind: parent.kind,
    getText: () => parent.text ?? parent.display ?? "",
    record: { index, styleId: parent.styleId, targetId: parent.targetId },
  }));
}

function documentRenderUsesDocxSource(options = {}) {
  const source = String(options.source || options.inputFormat || options.renderSource || "").trim().toLowerCase();
  return source === "docx" || options.docx === true || options.useDocx === true || options.native === true || options.nativeOffice === true || options.office === true;
}

async function renderDocumentFromDocx(document, options = {}) {
  const docx = await DocumentFile.exportDocx(document);
  docx.metadata = { ...(docx.metadata || {}), artifactKind: "document", format: "docx", renderSource: "docx" };
  const desiredType = renderTypeForOptions(options, docx.type);
  const format = String(options.format || "docx").trim().toLowerCase();
  if (!format || format === "docx" || desiredType === docx.type) return docx;
  const adapter = options.renderer || options.rasterRenderer || options.renderAdapter;
  if (typeof adapter !== "function") return docx;
  const converted = await adapter({ input: docx, source: docx, inputType: docx.type, outputType: desiredType, format: options.format, artifactKind: "document", options: { ...options, source: "docx" } });
  const blob = await fileBlobFromRenderOutput(converted, desiredType, { artifactKind: "document", format: options.format, renderedFrom: docx.type, renderSource: "docx" });
  if (!blob.type || blob.type === "application/octet-stream") blob.type = desiredType;
  return blob;
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
  resolve(id) { return String(id || "").endsWith("/text") ? documentTextRange(this, id) : this.id === id ? this : this.blocks.find((block) => block.id === id) || this.headers.find((block) => block.id === id) || this.footers.find((block) => block.id === id) || this.comments.find((comment) => comment.id === id) || this.styles.get(id); }

  toProto() { return { id: this.id, name: this.name, designPreset: this.designPreset, styles: Object.fromEntries(this.styles.values().map((style) => [style.id, style])), blocks: this.blocks.map((block) => block.toProto()), headers: this.headers.map((block) => block.toProto()), footers: this.footers.map((block) => block.toProto()), comments: this.comments.map((comment) => comment.toProto()) }; }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["paragraph", "table", "listItem", "hyperlink", "field", "citation", "image", "section", "change", "comment", "header", "footer"]);
    const records = [];
    if (kinds.has("document")) records.push({ kind: "document", id: this.id, name: this.name, blocks: this.blocks.length, designPreset: this.designPreset });
    if (kinds.has("layout")) records.push(...documentLayoutRecords(this, options));
    this.blocks.forEach((block, index) => { if (kinds.has(block.kind)) records.push(documentInspectRecord(this, block, index)); });
    if (kinds.has("header")) records.push(...this.headers.map((block, index) => documentInspectRecord(this, block, index)));
    if (kinds.has("footer")) records.push(...this.footers.map((block, index) => documentInspectRecord(this, block, index)));
    if (kinds.has("comment")) records.push(...this.comments.map((comment) => comment.inspectRecord()));
    if (kinds.has("textRange")) records.push(...documentTextRangeRecords(this));
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
        if (!Number.isFinite(block.widthDxa) || block.widthDxa <= 0) issues.push(verificationIssue("document", "invalidTableWidth", `Table ${block.id} has an invalid width.`, { id: block.id, widthDxa: block.widthDxa }));
        if (!Number.isFinite(block.indentDxa) || block.indentDxa < 0) issues.push(verificationIssue("document", "invalidTableIndent", `Table ${block.id} has an invalid indent.`, { id: block.id, indentDxa: block.indentDxa }));
        if (!Array.isArray(block.columnWidthsDxa) || block.columnWidthsDxa.length !== block.columns) issues.push(verificationIssue("document", "invalidTableColumnWidths", `Table ${block.id} needs one column width per column.`, { id: block.id, columns: block.columns, columnWidthsDxa: block.columnWidthsDxa }));
        else {
          if (block.columnWidthsDxa.some((value) => !Number.isFinite(value) || value <= 0)) issues.push(verificationIssue("document", "invalidTableColumnWidth", `Table ${block.id} contains an invalid column width.`, { id: block.id, columnWidthsDxa: block.columnWidthsDxa }));
          const widthSum = block.columnWidthsDxa.reduce((sum, value) => sum + value, 0);
          if (Number.isFinite(block.widthDxa) && widthSum !== block.widthDxa) issues.push(verificationIssue("document", "tableColumnWidthMismatch", `Table ${block.id} column widths do not equal the table width.`, { id: block.id, widthDxa: block.widthDxa, columnWidthsDxa: block.columnWidthsDxa, widthSum }));
        }
        for (const [side, value] of Object.entries(block.cellMarginsDxa || {})) if (!Number.isFinite(value) || value < 0) issues.push(verificationIssue("document", "invalidTableCellMargin", `Table ${block.id} has an invalid ${side} cell margin.`, { id: block.id, side, value }));
        if (!Number.isFinite(block.borderSize) || block.borderSize < 0) issues.push(verificationIssue("document", "invalidTableBorderSize", `Table ${block.id} has an invalid border size.`, { id: block.id, borderSize: block.borderSize }));
        if (!/^[A-F0-9]{6}$/.test(block.borderColor)) issues.push(verificationIssue("document", "invalidTableBorderColor", `Table ${block.id} has an invalid border color.`, { id: block.id, borderColor: block.borderColor }));
        if (!/^[A-F0-9]{6}$/.test(block.headerFill)) issues.push(verificationIssue("document", "invalidTableHeaderFill", `Table ${block.id} has an invalid header fill.`, { id: block.id, headerFill: block.headerFill }));
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
    if (options.format === "layout") return new FileBlob(JSON.stringify(this.layoutJson(options), null, 2), { type: LAYOUT_MIME, metadata: { artifactKind: "document", format: "layout", target: options.target ?? options.targetId ?? options.id ?? options.anchor, search: options.search ?? options.searchTerm } });
    if (documentRenderUsesDocxSource(options)) return renderDocumentFromDocx(this, options);
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
        const style = this.styles.effective(block.styleId) || this.styles.get("Normal");
        const fontSize = Math.max(10, (style.fontSize || 22) / 2);
        const runs = block.runs?.length ? block.runs : [{ text: block.text, style: {} }];
        const tspans = runs.map((run, index) => {
          const runStyle = { ...style, ...(run.style || {}) };
          return `<tspan${index ? "" : ` x=\"${margin}\"`} font-family="${xmlEscape(runStyle.fontFamily || "Arial")}" font-size="${fontSize}" font-style="${runStyle.italic ? "italic" : "normal"}" font-weight="${runStyle.bold ? "700" : "400"}" fill="${xmlEscape(runStyle.color || "#111827")}">${xmlEscape(run.text)}</tspan>`;
        }).join("");
        parts.push(`<text x="${margin}" y="${y}">${tspans}</text>`);
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
        if (block.dataUrl) {
          parts.push(`<image href="${attrEscape(block.dataUrl)}" x="${margin}" y="${y}" width="${imageWidth}" height="${imageHeight}" preserveAspectRatio="xMidYMid meet"/>`);
          parts.push(`<text x="${margin}" y="${y + imageHeight + 14}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(block.alt || block.name || "image")}</text>`);
        } else {
          parts.push(`<rect x="${margin}" y="${y}" width="${imageWidth}" height="${imageHeight}" fill="#fef3c7" stroke="#f59e0b"/>`);
          parts.push(`<text x="${margin + 8}" y="${y + 18}" font-family="Arial" font-size="11" fill="#92400e">${xmlEscape(block.alt || block.prompt || block.uri || block.name || "image")}</text>`);
        }
        y += imageHeight + (block.dataUrl ? 36 : 20);
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
        const style = this.styles.effective(block.styleId) || this.styles.get("Normal") || {};
        const fontSize = Math.max(10, (style.fontSize || 22) / 2);
        const key = `${block.listType}:${block.level}`;
        const next = (listCounters.get(key) || 0) + 1;
        listCounters.set(key, next);
        const marker = block.listType === "number" ? `${next}.` : "•";
        const x = margin + block.level * 24;
        parts.push(`<text x="${x}" y="${y}" font-family="${xmlEscape(style.fontFamily || "Arial")}" font-size="${fontSize}" font-style="${style.italic ? "italic" : "normal"}" font-weight="${style.bold ? "700" : "400"}" fill="${xmlEscape(style.color || "#111827")}">${xmlEscape(marker)}</text>`);
        parts.push(`<text x="${x + 22}" y="${y}" font-family="${xmlEscape(style.fontFamily || "Arial")}" font-size="${fontSize}" font-style="${style.italic ? "italic" : "normal"}" font-weight="${style.bold ? "700" : "400"}" fill="${xmlEscape(style.color || "#111827")}">${xmlEscape(block.text)}</text>`);
        y += Math.max(20, fontSize * 1.5);
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
  const styles = document.styles.values().map((style) => {
    const type = style.type || "paragraph";
    const basedOn = style.basedOn || style.parent || style.extends;
    const color = style.color ? `<w:color w:val="${attrEscape(String(style.color).replace(/^#/, ""))}"/>` : "";
    return `<w:style w:type="${attrEscape(type)}" w:styleId="${attrEscape(style.id)}"><w:name w:val="${attrEscape(style.name || style.id)}"/>${basedOn ? `<w:basedOn w:val="${attrEscape(basedOn)}"/>` : ""}<w:rPr>${style.bold ? "<w:b/>" : ""}${style.italic ? "<w:i/>" : ""}${color}<w:sz w:val="${Math.round(style.fontSize || 22)}"/><w:rFonts w:ascii="${attrEscape(style.fontFamily || "Aptos")}" w:hAnsi="${attrEscape(style.fontFamily || "Aptos")}"/></w:rPr></w:style>`;
  }).join("");
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

function docxRunPrXml(style = {}) {
  const color = style.color ? `<w:color w:val="${attrEscape(String(style.color).replace(/^#/, ""))}"/>` : "";
  const size = style.fontSize ? `<w:sz w:val="${Math.round(style.fontSize)}"/>` : "";
  const fonts = style.fontFamily ? `<w:rFonts w:ascii="${attrEscape(style.fontFamily)}" w:hAnsi="${attrEscape(style.fontFamily)}"/>` : "";
  const body = `${style.bold ? "<w:b/>" : ""}${style.italic ? "<w:i/>" : ""}${color}${size}${fonts}`;
  return body ? `<w:rPr>${body}</w:rPr>` : "";
}

function docxRunXml(run = {}) {
  return `<w:r>${docxRunPrXml(run.style || run.textStyle || {})}<w:t>${xmlEscape(run.text ?? "")}</w:t></w:r>`;
}

function docxParagraphXml(block, commentIndexes) {
  const commentStart = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeStart w:id="${id}"/>`).join("") : "";
  const commentEnd = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeEnd w:id="${id}"/>`).join("") : "";
  const refs = commentIndexes.length ? commentIndexes.map((id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`).join("") : "";
  const numPr = block.kind === "listItem" ? `<w:numPr><w:ilvl w:val="${Math.max(0, block.level || 0)}"/><w:numId w:val="${block.listType === "number" ? 2 : 1}"/></w:numPr>` : "";
  const runs = block.runs?.length ? block.runs.map(docxRunXml).join("") : `<w:r><w:t>${xmlEscape(block.text)}</w:t></w:r>`;
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/>${numPr}</w:pPr>${commentStart}${runs}${commentEnd}${refs}</w:p>`;
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
  const columns = Math.max(1, block.columns || 1);
  const widthDxa = Number.isFinite(block.widthDxa) && block.widthDxa > 0 ? Math.round(block.widthDxa) : 9360;
  const configuredWidths = Array.isArray(block.columnWidthsDxa) && block.columnWidthsDxa.length === columns && block.columnWidthsDxa.every((value) => Number.isFinite(value) && value > 0)
    ? block.columnWidthsDxa.map(Math.round)
    : documentTableDefaultColumnWidths(columns, widthDxa);
  const columnWidths = configuredWidths.reduce((sum, value) => sum + value, 0) === widthDxa ? configuredWidths : documentTableDefaultColumnWidths(columns, widthDxa);
  const margins = block.cellMarginsDxa || {};
  const border = `<w:top w:val="single" w:sz="${Math.max(0, Math.round(block.borderSize ?? 4))}" w:space="0" w:color="${attrEscape(block.borderColor || "D9D9D9")}"/><w:left w:val="single" w:sz="${Math.max(0, Math.round(block.borderSize ?? 4))}" w:space="0" w:color="${attrEscape(block.borderColor || "D9D9D9")}"/><w:bottom w:val="single" w:sz="${Math.max(0, Math.round(block.borderSize ?? 4))}" w:space="0" w:color="${attrEscape(block.borderColor || "D9D9D9")}"/><w:right w:val="single" w:sz="${Math.max(0, Math.round(block.borderSize ?? 4))}" w:space="0" w:color="${attrEscape(block.borderColor || "D9D9D9")}"/><w:insideH w:val="single" w:sz="${Math.max(0, Math.round(block.borderSize ?? 4))}" w:space="0" w:color="${attrEscape(block.borderColor || "D9D9D9")}"/><w:insideV w:val="single" w:sz="${Math.max(0, Math.round(block.borderSize ?? 4))}" w:space="0" w:color="${attrEscape(block.borderColor || "D9D9D9")}"/>`;
  const grid = columnWidths.map((width) => `<w:gridCol w:w="${width}"/>`).join("");
  const rows = block.values.map((row, rowIndex) => `<w:tr>${Array.from({ length: columns }, (_, column) => {
    const headerProperties = rowIndex === 0 ? `<w:shd w:val="clear" w:color="auto" w:fill="${attrEscape(block.headerFill || "F2F4F7")}"/>` : "";
    const runProperties = rowIndex === 0 ? "<w:rPr><w:b/></w:rPr>" : "";
    return `<w:tc><w:tcPr><w:tcW w:w="${columnWidths[column]}" w:type="dxa"/>${headerProperties}</w:tcPr><w:p><w:r>${runProperties}<w:t>${xmlEscape(row[column] ?? "")}</w:t></w:r></w:p></w:tc>`;
  }).join("")}</w:tr>`).join("");
  return `<w:tbl><w:tblPr><w:tblStyle w:val="${attrEscape(block.styleId || "TableGrid")}"/><w:tblW w:w="${widthDxa}" w:type="dxa"/><w:tblInd w:w="${Math.max(0, Math.round(block.indentDxa ?? 120))}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="${Math.max(0, Math.round(margins.top ?? 80))}" w:type="dxa"/><w:start w:w="${Math.max(0, Math.round(margins.start ?? 120))}" w:type="dxa"/><w:bottom w:w="${Math.max(0, Math.round(margins.bottom ?? 80))}" w:type="dxa"/><w:end w:w="${Math.max(0, Math.round(margins.end ?? 120))}" w:type="dxa"/></w:tblCellMar><w:tblBorders>${border}</w:tblBorders></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
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

function parseDocxStylesXml(xml = "") {
  const styles = {};
  for (const match of String(xml || "").matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    const attrs = match[1] || "";
    const body = match[2] || "";
    const id = decodeXml(/\bw:styleId="([^"]+)"/.exec(attrs)?.[1] || "");
    if (!id) continue;
    const type = /\bw:type="([^"]+)"/.exec(attrs)?.[1] || "paragraph";
    const name = decodeXml(/<w:name[^>]*w:val="([^"]*)"/.exec(body)?.[1] || id);
    const basedOn = decodeXml(/<w:basedOn[^>]*w:val="([^"]*)"/.exec(body)?.[1] || "") || undefined;
    const fontSize = Number(/<w:sz[^>]*w:val="([^"]+)"/.exec(body)?.[1]);
    const fontFamily = decodeXml(/<w:rFonts[^>]*w:ascii="([^"]*)"/.exec(body)?.[1] || /<w:rFonts[^>]*w:hAnsi="([^"]*)"/.exec(body)?.[1] || "");
    const color = /<w:color[^>]*w:val="([A-Fa-f0-9]{3,6})"/.exec(body)?.[1];
    const bold = /<w:b\b/.test(body);
    const italic = /<w:i\b/.test(body);
    styles[id] = { id, name, type, ...(basedOn ? { basedOn } : {}), ...(Number.isFinite(fontSize) ? { fontSize } : {}), ...(fontFamily ? { fontFamily } : {}), ...(bold ? { bold: true } : {}), ...(italic ? { italic: true } : {}), ...(color ? { color: `#${color}` } : {}) };
  }
  return styles;
}

function parseDocxRuns(part = "") {
  return [...String(part || "").matchAll(/<w:r\b[\s\S]*?<\/w:r>/g)].map((match) => {
    const runXml = match[0];
    const text = decodeXml([...runXml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join(""));
    if (!text) return undefined;
    const rPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(runXml)?.[1] || "";
    const color = /<w:color[^>]*w:val="([A-Fa-f0-9]{3,6})"/.exec(rPr)?.[1];
    const fontSize = Number(/<w:sz[^>]*w:val="([^"]+)"/.exec(rPr)?.[1]);
    const fontFamily = decodeXml(/<w:rFonts[^>]*w:ascii="([^"]*)"/.exec(rPr)?.[1] || /<w:rFonts[^>]*w:hAnsi="([^"]*)"/.exec(rPr)?.[1] || "");
    const style = { ...(/<w:b\b/.test(rPr) ? { bold: true } : {}), ...(/<w:i\b/.test(rPr) ? { italic: true } : {}), ...(color ? { color: `#${color}` } : {}), ...(Number.isFinite(fontSize) ? { fontSize } : {}), ...(fontFamily ? { fontFamily } : {}) };
    return { text, style };
  }).filter(Boolean);
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
  const runs = parseDocxRuns(part);
  const text = runs.length ? runs.map((run) => run.text).join("") : decodeXml([...part.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join(""));
  const numId = /<w:numId[^>]*w:val="(\d+)"/.exec(part)?.[1];
  const level = Number(/<w:ilvl[^>]*w:val="(\d+)"/.exec(part)?.[1] || 0);
  if (numId) return { block: { kind: "listItem", text, styleId, level, listType: numId === "2" ? "number" : "bullet" }, commentIds };
  return { block: { kind: "paragraph", text, styleId, runs: runs.length ? runs : undefined }, commentIds };
}

function parseDocxTable(part) {
  const values = [...part.matchAll(/<w:tr[\s\S]*?<\/w:tr>/g)].map((rowMatch) => [...rowMatch[0].matchAll(/<w:tc[\s\S]*?<\/w:tc>/g)].map((cellMatch) => decodeXml([...cellMatch[0].matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join(""))));
  const readDxa = (tag, fallback) => Number(new RegExp(`<w:${tag}\\b[^>]*w:w="(\\d+)"`).exec(part)?.[1] ?? fallback);
  const columns = Math.max(0, ...values.map((row) => row.length));
  const columnWidthsDxa = [...part.matchAll(/<w:gridCol\b[^>]*w:w="(\d+)"[^>]*\/?\s*>/g)].map((match) => Number(match[1]));
  const firstRow = /<w:tr[\s\S]*?<\/w:tr>/.exec(part)?.[0] || "";
  return {
    kind: "table",
    styleId: decodeXml(/<w:tblStyle\b[^>]*w:val="([^"]+)"/.exec(part)?.[1] || "TableGrid"),
    values,
    rows: values.length,
    columns,
    widthDxa: readDxa("tblW", columnWidthsDxa.reduce((sum, value) => sum + value, 0) || 9360),
    indentDxa: readDxa("tblInd", 0),
    columnWidthsDxa: columnWidthsDxa.length === columns ? columnWidthsDxa : undefined,
    cellMarginsDxa: {
      top: readDxa("top", 80),
      bottom: readDxa("bottom", 80),
      start: readDxa("start", readDxa("left", 120)),
      end: readDxa("end", readDxa("right", 120)),
    },
    borderColor: String(/<w:tblBorders>[\s\S]*?<w:(?:top|left|start)\b[^>]*w:color="([^"]+)"/.exec(part)?.[1] || "D9D9D9").toUpperCase(),
    borderSize: Number(/<w:tblBorders>[\s\S]*?<w:(?:top|left|start)\b[^>]*w:sz="(\d+)"/.exec(part)?.[1] || 4),
    headerFill: String(/<w:shd\b[^>]*w:fill="([^"]+)"/.exec(firstRow)?.[1] || "F2F4F7").toUpperCase(),
  };
}

function parseHeaderFooterXml(xml) {
  return [...String(xml || "").matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((match) => ({ text: decodeXml([...match[0].matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join("")) })).filter((item) => item.text.length > 0);
}

function docxSafePartPath(partPath) {
  const raw = String(partPath || "").replaceAll("\\", "/").trim();
  if (!raw || raw.startsWith("/") || raw.includes("\0")) throw new Error(`Unsafe DOCX part path: ${partPath}`);
  const normalized = path.posix.normalize(raw).replace(/^\.\//, "");
  if (!normalized || normalized === "." || normalized.startsWith("../") || normalized.includes("/../") || normalized === "..") throw new Error(`Unsafe DOCX part path: ${partPath}`);
  return normalized;
}

async function docxPackageRecords(zip, options = {}) {
  const includeText = Boolean(options.includeText || options.preview || options.includeXml);
  const maxPreviewChars = Math.max(0, Number(options.maxPreviewChars ?? 400) || 0);
  const files = Object.values(zip.files).filter((file) => !file.dir).sort((a, b) => a.name.localeCompare(b.name));
  const records = [{ kind: "docxPackage", parts: files.length }];
  for (const file of files) {
    const bytes = await file.async("uint8array");
    const record = { kind: "docxPart", path: file.name, size: bytes.byteLength, contentType: file.name.endsWith(".xml") ? "application/xml" : file.name.endsWith(".json") ? "application/json" : "application/octet-stream" };
    if (includeText && /\.(xml|json|rels)$/i.test(file.name)) record.textPreview = decoder.decode(bytes).slice(0, maxPreviewChars);
    records.push(record);
  }
  return records;
}

function docxPatchData(patch, options = {}) {
  if (patch.json !== undefined) return encoder.encode(JSON.stringify(patch.json, null, 2));
  if (patch.text !== undefined || patch.xml !== undefined) return encoder.encode(String(patch.text ?? patch.xml));
  if (patch.bytes !== undefined || patch.data !== undefined || patch.buffer !== undefined) return toUint8Array(patch.bytes ?? patch.data ?? patch.buffer);
  if (patch.content !== undefined) return typeof patch.content === "string" ? encoder.encode(patch.content) : toUint8Array(patch.content);
  if (patch.remove || patch.delete) return undefined;
  throw new Error(`DOCX patch for ${patch.path || patch.part || "unknown part"} has no content or remove flag.`);
}

export class DocumentFile {
  static async inspectDocx(blobOrBuffer, options = {}) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const records = await docxPackageRecords(zip, options);
    return { parts: records.filter((record) => record.kind === "docxPart"), records, ...ndjson(records, options.maxChars ?? Infinity) };
  }

  static async patchDocx(blobOrBuffer, patches = [], options = {}) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const list = Array.isArray(patches) ? patches : Object.entries(patches || {}).map(([partPath, content]) => ({ path: partPath, content }));
    const maxPatchBytes = Number(options.maxPatchBytes ?? 5 * 1024 * 1024);
    for (const patch of list) {
      const partPath = docxSafePartPath(patch.path || patch.part || patch.name);
      if (patch.remove || patch.delete) { zip.remove(partPath); continue; }
      const data = docxPatchData(patch, options);
      if (maxPatchBytes && data?.byteLength > maxPatchBytes) throw new Error(`DOCX patch for ${partPath} exceeds maxPatchBytes (${maxPatchBytes}).`);
      zip.file(partPath, data);
    }
    return new FileBlob(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: DOCX_MIME, metadata: { artifactKind: "document", patchedParts: list.length } });
  }

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
    const stylesText = await zip.file("word/styles.xml")?.async("text");
    const importedStyles = parseDocxStylesXml(stylesText);
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
    const document = DocumentModel.create({ styles: importedStyles, blocks: blocks.length ? blocks : [{ kind: "paragraph", text: "" }] });
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
    this.source = config.source;
  }

  inspectRecord(pageIndex) { return { kind: "table", id: this.id, page: pageIndex + 1, name: this.name || undefined, rows: this.values.length, cols: Math.max(0, ...this.values.map((row) => row.length)), bbox: this.bbox, values: this.values, source: this.source || undefined }; }
  toJSON() { return { id: this.id, name: this.name, values: this.values, bbox: this.bbox, source: this.source }; }
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

class PdfChart {
  constructor(page, config = {}) {
    this.page = page;
    this.id = config.id || aid("pch");
    this.name = config.name || "";
    this.title = config.title || config.name || "Chart";
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

  inspectRecord(pageIndex) { return { kind: "chart", id: this.id, page: pageIndex + 1, name: this.name || undefined, title: this.title, chartType: this.chartType, categories: this.categories, series: this.series, bbox: this.bbox }; }
  toJSON() { return { id: this.id, name: this.name, title: this.title, chartType: this.chartType, categories: this.categories, series: this.series, bbox: this.bbox }; }
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
  }

  normalizeTextItem(item = {}, index = this.textItems?.length || 0) {
    const bbox = pdfTextItemBBox(item);
    return { id: item.id || `${this.id}/txt/${index + 1}`, text: String(item.text ?? item.str ?? ""), bbox, fontName: item.fontName || item.fontFamily, fontSize: item.fontSize || item.size, color: item.color, bold: Boolean(item.bold), italic: Boolean(item.italic), dir: item.dir };
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
  inspectRecord(index) { return { kind: "page", id: this.id, page: index + 1, width: this.width, height: this.height, textPreview: this.text.slice(0, 300), textChars: this.text.length, textItems: this.textItems.length, tables: this.tables.length, images: this.images.length, charts: this.charts.length, regions: this.regions.length }; }
  textRecord(index) { return { kind: "text", id: `${this.id}/text`, page: index + 1, text: this.text, textChars: this.text.length, textItems: this.textItems.length }; }
  textItemRecords(index) { return this.textItems.map((item) => ({ kind: "textItem", page: index + 1, ...item })); }
  regionRecords(index) { return this.regions.map((region) => ({ ...region, kind: "region", regionKind: region.kind || "region", page: index + 1 })); }
  toJSON() { return { id: this.id, text: this.text, width: this.width, height: this.height, textItems: this.textItems, regions: this.regions, tables: this.tables.map((table) => table.toJSON()), images: this.images.map((image) => image.toJSON()), charts: this.charts.map((chart) => chart.toJSON()) }; }
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
    if (!keepByPage.has(pageArrayIndex)) keepByPage.set(pageArrayIndex, { full: false, text: false, textItems: new Set(), regions: new Set(), tables: new Set(), images: new Set(), charts: new Set() });
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
  addTable(config = {}) { return (this.pages[0] || this.addPage()).addTable(config); }
  addImage(config = {}) { const pageIndex = Number(config.pageIndex ?? config.page ?? 0); return (this.pages[pageIndex] || this.pages[0] || this.addPage()).addImage(config); }
  addChart(config = {}) { const pageIndex = Number(config.pageIndex ?? config.page ?? 0); return (this.pages[pageIndex] || this.pages[0] || this.addPage()).addChart(config); }
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
      const chart = page.charts.find((item) => item.id === id);
      if (chart) return chart;
    }
    return undefined;
  }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["page", "text", "table", "image", "chart"]);
    const records = [];
    this.pages.forEach((page, index) => {
      if (kinds.has("page")) records.push(page.inspectRecord(index));
      if (kinds.has("text")) records.push(page.textRecord(index));
      if (kinds.has("textItem")) records.push(...page.textItemRecords(index));
      if (kinds.has("region")) records.push(...page.regionRecords(index));
      if (kinds.has("table")) records.push(...page.tables.map((table) => table.inspectRecord(index)));
      if (kinds.has("image")) records.push(...page.images.map((image) => image.inspectRecord(index)));
      if (kinds.has("chart")) records.push(...page.charts.map((chart) => chart.inspectRecord(index)));
    });
    return ndjson(filterInspectRecords(records, options), options.maxChars ?? Infinity);
  }

  verify(options = {}) {
    const issues = [];
    if (this.pages.length === 0) issues.push(verificationIssue("pdf", "noPages", "PDF artifact has no pages."));
    this.pages.forEach((page, pageIndex) => {
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
        const [left, top, width, height] = table.bbox || [];
        if (left < 0 || top < 0 || width <= 0 || height <= 0 || left + width > page.width || top + height > page.height) {
          issues.push(verificationIssue("pdf", "tableOutOfBounds", `PDF table ${table.id} extends outside page ${pageIndex + 1}.`, { page: pageIndex + 1, id: table.id, bbox: table.bbox }));
        }
      }
      for (const image of page.images) {
        if (!image.dataUrl && !image.uri && !image.prompt) issues.push(verificationIssue("pdf", "emptyImage", `PDF image ${image.id} on page ${pageIndex + 1} has no dataUrl, uri, or prompt.`, { page: pageIndex + 1, id: image.id }));
        if (image.dataUrl && !imageDataFromDataUrl(image.dataUrl)) issues.push(verificationIssue("pdf", "invalidImageDataUrl", `PDF image ${image.id} on page ${pageIndex + 1} has an unsupported data URL.`, { page: pageIndex + 1, id: image.id }));
        const [left, top, width, height] = image.bbox || [];
        if (left < 0 || top < 0 || width <= 0 || height <= 0 || left + width > page.width || top + height > page.height) {
          issues.push(verificationIssue("pdf", "imageOutOfBounds", `PDF image ${image.id} extends outside page ${pageIndex + 1}.`, { page: pageIndex + 1, id: image.id, bbox: image.bbox }));
        }
      }
      for (const chart of page.charts) {
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
    return verificationResult("pdf", issues, options);
  }

  help(query = "*", options = {}) { return helpArtifact("pdf", query, options); }

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
          tables: page.tables.map((table) => ({ kind: "table", id: table.id, page: pageNumber, name: table.name || undefined, values: table.values, bbox: table.bbox, source: table.source || undefined })),
          images: page.images.map((image) => ({ kind: "image", id: image.id, page: pageNumber, name: image.name || undefined, alt: image.alt, bbox: image.bbox, fit: image.fit, hasDataUrl: Boolean(image.dataUrl), uri: image.uri, prompt: image.prompt })),
          charts: page.charts.map((chart) => ({ kind: "chart", id: chart.id, page: pageNumber, name: chart.name || undefined, title: chart.title, chartType: chart.chartType, categories: chart.categories, series: chart.series, bbox: chart.bbox })),
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
    const records = [
      { kind: "pdfFile", bytes: bytes.byteLength, version, pages, objects, hasEmbeddedModel: /%OPEN_OFFICE_ARTIFACT [A-Za-z0-9+/=]+/.test(text), hasEof: /%%EOF\s*$/.test(text) },
      ...[...text.matchAll(/(\d+)\s+0\s+obj\s*<<([\s\S]*?)>>/g)].slice(0, Math.max(0, Number(options.maxObjects ?? 200) || 0)).map((match) => ({ kind: "pdfObject", object: Number(match[1]), type: /\/Type\s*\/([A-Za-z0-9]+)/.exec(match[2])?.[1], subtype: /\/Subtype\s*\/([A-Za-z0-9]+)/.exec(match[2])?.[1], stream: /\bstream\b/.test(match[0]) })),
    ];
    return { records, summary: records[0], ...ndjson(records, options.maxChars ?? Infinity) };
  }

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
    tables: ((page.tables || []).length ? (page.tables || []) : reconstructPdfTablesFromTextGeometry(page, index)).map((table, tableIndex) => ({ name: table.name || `parsed-table-${index + 1}-${tableIndex + 1}`, values: table.values || table.rows || [], bbox: table.bbox || table.bounds || [72, 140 + tableIndex * 120, 468, 96], source: table.source })),
    images: (page.images || []).map((image, imageIndex) => ({ name: image.name || `parsed-image-${index + 1}-${imageIndex + 1}`, alt: image.alt || image.altText || image.name || "parsed image", dataUrl: pdfImageDataUrl(image), uri: image.uri, prompt: image.prompt, bbox: image.bbox || image.bounds || [72, 280 + imageIndex * 140, 180, 120] })),
    charts: (page.charts || []).map((chart, chartIndex) => ({ name: chart.name || `parsed-chart-${index + 1}-${chartIndex + 1}`, title: chart.title || chart.name || `Parsed chart ${chartIndex + 1}`, chartType: chart.chartType || chart.type || "bar", categories: chart.categories || chart.labels || [], series: chart.series || [{ name: chart.seriesName || "Series 1", values: chart.values || chart.data || [] }], bbox: chart.bbox || chart.bounds || [72, 430 + chartIndex * 180, 468, 160] })),
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
  const positionedValues = new Set((page.textItems || []).map((item) => String(item.text || "").trim()).filter(Boolean));
  const lines = String(page.text || "").split(/\r?\n/).map((line) => line.trim()).filter((line) => line && !positionedValues.has(line));
  const positionedText = (page.textItems || []).map((item) => {
    const [left, top, itemWidth, itemHeight] = item.bbox || [72, 72, 0, 14];
    const fontSize = Math.max(6, Number(item.fontSize || itemHeight || 12));
    return `<text x="${left}" y="${top + fontSize}" font-family="${xmlEscape(item.fontName || "Helvetica")}" font-size="${fontSize}" font-weight="${item.bold ? "700" : "400"}" font-style="${item.italic ? "italic" : "normal"}" fill="${xmlEscape(item.color || "#111827")}" data-text-item-id="${attrEscape(item.id || "")}">${xmlEscape(item.text || "")}</text>`;
  }).join("");
  const lineText = lines.map((line, index) => `<text x="72" y="${96 + index * (index ? 22 : 30)}" font-family="Helvetica" font-size="${index === 0 ? 24 : 14}" font-weight="${index === 0 ? "700" : "400"}" fill="${index === 0 ? "#0f172a" : "#334155"}">${xmlEscape(line)}</text>`).join("");
  const text = `${lineText}${positionedText}`;
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
    if (image.dataUrl) return `<image href="${attrEscape(image.dataUrl)}" x="${left}" y="${top}" width="${imageWidth}" height="${imageHeight}" preserveAspectRatio="xMidYMid meet"/><text x="${left}" y="${top + imageHeight + 14}" font-family="Helvetica" font-size="10" fill="#475569">${xmlEscape(image.alt || image.name || "image")}</text>`;
    return `<rect x="${left}" y="${top}" width="${imageWidth}" height="${imageHeight}" fill="#fef3c7" stroke="#f59e0b"/><text x="${left + 8}" y="${top + 18}" font-family="Helvetica" font-size="11" fill="#92400e">${xmlEscape(image.alt || image.prompt || image.uri || image.name || "image")}</text>`;
  }).join("");
  const charts = (page.charts || []).map(pdfChartSvg).join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="white"/>${text}${tables}${images}${charts}</svg>`;
}

function escapePdfString(text) {
  return String(text).replace(/[^\x20-\x7e]/g, "?").replaceAll("\\", "\\\\").replaceAll("(", "\\(").replaceAll(")", "\\)");
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

function pdfTextCommand(page, text, left, top, options = {}) {
  const fontSize = Math.max(1, Number(options.fontSize || 12));
  const baseline = page.height - Number(top) - fontSize;
  return `BT /${options.bold ? "F2" : "F1"} ${pdfNumber(fontSize)} Tf ${pdfColorCommand(options.color || "#111827")} ${pdfNumber(left)} ${pdfNumber(baseline)} Td (${escapePdfString(text)}) Tj ET`;
}

function pdfLineCommand(page, x1, top1, x2, top2, options = {}) {
  return `q ${pdfColorCommand(options.color || "#94a3b8", true)} ${pdfNumber(options.width || 1)} w ${pdfNumber(x1)} ${pdfNumber(page.height - top1)} m ${pdfNumber(x2)} ${pdfNumber(page.height - top2)} l S Q`;
}

function pdfRectCommand(page, bbox, options = {}) {
  const [left, top, width, height] = bbox.map(Number);
  const operator = options.fill === false ? "S" : options.stroke === false ? "f" : "B";
  return `q ${pdfColorCommand(options.fillColor || "#ffffff")} ${pdfColorCommand(options.strokeColor || "#cbd5e1", true)} ${pdfNumber(options.lineWidth || 1)} w ${pdfNumber(left)} ${pdfNumber(page.height - top - height)} ${pdfNumber(width)} ${pdfNumber(height)} re ${operator} Q`;
}

function pdfFitText(text, width, fontSize) {
  const maxChars = Math.max(1, Math.floor(Number(width) / Math.max(1, Number(fontSize) * 0.55)));
  const value = String(text ?? "");
  return value.length <= maxChars ? value : `${value.slice(0, Math.max(1, maxChars - 1))}...`;
}

function pdfPageTextCommands(page) {
  const positioned = new Set((page.textItems || []).map((item) => String(item.text || "").trim()).filter(Boolean));
  const lines = String(page.text || "").split(/\r?\n/).map((line) => line.trim()).filter((line) => line && !positioned.has(line));
  return lines.map((line, index) => pdfTextCommand(page, line, 72, 72 + index * (index ? 22 : 30), { fontSize: index === 0 ? 24 : 14, bold: index === 0, color: index === 0 ? "#0f172a" : "#334155" }));
}

function pdfPositionedTextCommands(page) {
  return (page.textItems || []).filter((item) => item.text).map((item) => {
    const [left, top, width, height] = item.bbox || [72, 72, 120, 14];
    const fontSize = Math.max(6, Number(item.fontSize || height || 12));
    return pdfTextCommand(page, pdfFitText(item.text, width || page.width - left, fontSize), left, top, { fontSize, color: item.color || "#111827", bold: Boolean(item.bold) });
  });
}

function pdfTableCommands(page, table) {
  const [left, top, tableWidth, tableHeight] = table.bbox.map(Number);
  const rows = Math.max(1, table.values.length);
  const columns = Math.max(1, ...table.values.map((row) => row.length));
  const cellWidth = tableWidth / columns;
  const cellHeight = tableHeight / rows;
  const commands = [];
  for (let row = 0; row < rows; row += 1) {
    for (let column = 0; column < columns; column += 1) {
      const bbox = [left + column * cellWidth, top + row * cellHeight, cellWidth, cellHeight];
      commands.push(pdfRectCommand(page, bbox, { fillColor: row === 0 ? "#e2e8f0" : row % 2 ? "#f8fafc" : "#ffffff", strokeColor: "#94a3b8", lineWidth: 0.75 }));
      const fontSize = Math.max(8, Math.min(12, cellHeight * 0.32));
      commands.push(pdfTextCommand(page, pdfFitText(table.values[row]?.[column] ?? "", cellWidth - 12, fontSize), bbox[0] + 6, bbox[1] + Math.max(5, (cellHeight - fontSize) / 2), { fontSize, bold: row === 0, color: "#0f172a" }));
    }
  }
  return commands;
}

function pdfChartCommands(page, chart) {
  const [left, top, width, height] = chart.bbox.map(Number);
  const commands = [pdfRectCommand(page, chart.bbox, { fillColor: "#ffffff", strokeColor: "#cbd5e1", lineWidth: 1 }), pdfTextCommand(page, chart.title, left + 10, top + 10, { fontSize: 13, bold: true })];
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
  categories.slice(0, 8).forEach((category, index) => commands.push(pdfTextCommand(page, pdfFitText(category, plot.width / Math.max(1, categories.length), 8), plot.left + index * (plot.width / Math.max(1, categories.length)), top + height - 18, { fontSize: 8, color: "#475569" })));
  chart.series.slice(0, 4).forEach((series, index) => {
    commands.push(pdfRectCommand(page, [left + width - 94, top + 12 + index * 14, 8, 8], { fillColor: series.color, stroke: false }));
    commands.push(pdfTextCommand(page, pdfFitText(series.name, 74, 8), left + width - 82, top + 10 + index * 14, { fontSize: 8, color: "#334155" }));
  });
  return commands;
}

function pdfPngAsset(image, objectId, resourceName) {
  const data = imageDataFromDataUrl(image.dataUrl);
  if (!data || data.contentType !== "image/png") return undefined;
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
    const compressed = deflateSync(rgb);
    return { image, objectId, resourceName, width: decoded.width, height: decoded.height, compressed };
  } catch (error) {
    throw new Error(`Unable to embed PNG image ${image.name || image.id || "image"}: ${error.message}`);
  }
}

function pdfImageCommands(page, image, asset) {
  const [left, top, width, height] = image.bbox.map(Number);
  if (!asset) return [pdfRectCommand(page, image.bbox, { fillColor: "#fef3c7", strokeColor: "#f59e0b" }), pdfTextCommand(page, pdfFitText(image.alt || image.prompt || image.uri || image.name || "image", width - 16, 10), left + 8, top + 8, { fontSize: 10, color: "#92400e" })];
  const sourceRatio = asset.width / asset.height;
  const frameRatio = width / height;
  const drawWidth = image.fit === "cover" ? (sourceRatio > frameRatio ? width : height * sourceRatio) : (sourceRatio > frameRatio ? width : height * sourceRatio);
  const drawHeight = image.fit === "cover" ? (sourceRatio > frameRatio ? width / sourceRatio : height) : (sourceRatio > frameRatio ? width / sourceRatio : height);
  const x = left + (width - drawWidth) / 2;
  const drawTop = top + (height - drawHeight) / 2;
  return [`q ${pdfNumber(drawWidth)} 0 0 ${pdfNumber(drawHeight)} ${pdfNumber(x)} ${pdfNumber(page.height - drawTop - drawHeight)} cm /${asset.resourceName} Do Q`];
}

function buildMinimalPdf(artifact) {
  const pages = artifact.pages.length ? artifact.pages : [new PdfPage(artifact)];
  const metadata = Buffer.from(JSON.stringify(artifact.toJSON()), "utf8").toString("base64");
  let nextObjectId = 5;
  const plans = pages.map((page) => {
    const imageAssets = [];
    for (const image of page.images) {
      const asset = pdfPngAsset(image, nextObjectId, `Im${imageAssets.length + 1}`);
      if (asset) { imageAssets.push(asset); nextObjectId += 1; }
    }
    const pageObjectId = nextObjectId++;
    const contentObjectId = nextObjectId++;
    return { page, imageAssets, pageObjectId, contentObjectId };
  });
  const objects = new Map();
  objects.set(1, Buffer.from("<< /Type /Catalog /Pages 2 0 R >>", "ascii"));
  objects.set(2, Buffer.from(`<< /Type /Pages /Kids [${plans.map((plan) => `${plan.pageObjectId} 0 R`).join(" ")}] /Count ${plans.length} >>`, "ascii"));
  objects.set(3, Buffer.from("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>", "ascii"));
  objects.set(4, Buffer.from("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>", "ascii"));
  for (const plan of plans) {
    const { page, imageAssets } = plan;
    const assetByImage = new Map(imageAssets.map((asset) => [asset.image, asset]));
    const commands = [
      ...pdfPageTextCommands(page),
      ...pdfPositionedTextCommands(page),
      ...page.tables.flatMap((table) => pdfTableCommands(page, table)),
      ...page.images.flatMap((image) => pdfImageCommands(page, image, assetByImage.get(image))),
      ...page.charts.flatMap((chart) => pdfChartCommands(page, chart)),
    ];
    const content = Buffer.from(`${commands.join("\n")}\n`, "ascii");
    const xobjects = imageAssets.length ? `/XObject << ${imageAssets.map((asset) => `/${asset.resourceName} ${asset.objectId} 0 R`).join(" ")} >>` : "";
    objects.set(plan.pageObjectId, Buffer.from(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pdfNumber(page.width || 612)} ${pdfNumber(page.height || 792)}] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> ${xobjects} >> /Contents ${plan.contentObjectId} 0 R >>`, "ascii"));
    objects.set(plan.contentObjectId, Buffer.concat([Buffer.from(`<< /Length ${content.byteLength} >>\nstream\n`, "ascii"), content, Buffer.from("endstream", "ascii")]));
    for (const asset of imageAssets) objects.set(asset.objectId, Buffer.concat([Buffer.from(`<< /Type /XObject /Subtype /Image /Width ${asset.width} /Height ${asset.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length ${asset.compressed.byteLength} >>\nstream\n`, "ascii"), asset.compressed, Buffer.from("\nendstream", "ascii")]));
  }
  const header = Buffer.from(`%PDF-1.4\n%OPEN_OFFICE_ARTIFACT ${metadata}\n`, "ascii");
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
  xref += `trailer\n<< /Size ${nextObjectId} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`;
  chunks.push(Buffer.from(xref, "ascii"));
  return new Uint8Array(Buffer.concat(chunks));
}
