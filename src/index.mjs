import path from "node:path";
import { deflateSync } from "node:zlib";
import JSZip from "jszip";
import { validatePptxPackageSemantics } from "./ooxml/pptx-package-semantics.mjs";
import { mutateOoxmlSourceReference, mutateOoxmlSourceReferenceTarget, supportedOoxmlSourceReferenceSummary, supportsOoxmlSourceReference, validateOoxmlSourceReferenceTarget } from "./ooxml/source-references.mjs";
import { normalizeDocxSettings } from "./ooxml/docx-settings.mjs";
import { effectiveDocxRunStyle, mergeDocxRunStyleCascade, normalizeDocxRunStyle, normalizeDocxThemeConfig } from "./ooxml/docx-run-styles.mjs";
import { DOCX_COMMENTS_EXTENDED_CONTENT_TYPE, DOCX_COMMENTS_EXTENDED_RELATIONSHIP_TYPE, DOCX_COMMENTS_EXTENSIBLE_CONTENT_TYPE, DOCX_COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE, DOCX_COMMENTS_IDS_CONTENT_TYPE, DOCX_COMMENTS_IDS_RELATIONSHIP_TYPE, DOCX_PEOPLE_CONTENT_TYPE, DOCX_PEOPLE_RELATIONSHIP_TYPE, validateDocxCommentPackageSemantics } from "./ooxml/docx-comments.mjs";
import { validateDocxLinkPackageSemantics } from "./ooxml/docx-links.mjs";
import { normalizeDocxBibliographySource, validateDocxBibliographyPackageSemantics } from "./ooxml/docx-bibliography.mjs";
import { normalizeDocxSectionSettings, planDocxHeaderFooterSections, resolveDocxPageHeaderFooter } from "./ooxml/docx-sections.mjs";
import { normalizeDocumentPictureBullet } from "./ooxml/docx-numbering.mjs";
import { resolveColorToken } from "./shared/colors.mjs";
import { matchesFormulaCriteria } from "./spreadsheet/formula-criteria.mjs";
import { normalizeSpreadsheetChartSeriesLine, spreadsheetChartLineDashArray } from "./spreadsheet/chart-line-style.mjs";
import { normalizeSpreadsheetChartLineOptions, spreadsheetChartSmoothLinePath } from "./spreadsheet/chart-line-options.mjs";
import { normalizeSpreadsheetChartSeriesMarker, spreadsheetChartMarkerSvg } from "./spreadsheet/chart-marker-style.mjs";
import { normalizeSpreadsheetChartDataLabels, spreadsheetChartDataLabelSvgPlacement, spreadsheetChartDataLabelText } from "./spreadsheet/chart-data-labels.mjs";
import { resolvedWorksheetChartCategories, resolvedWorksheetChartSeriesValues } from "./spreadsheet/chart-source-data.mjs";
import { computePivotValues, normalizePivotConfig } from "./spreadsheet/pivots.mjs";
import { formatSpreadsheetDisplayValue, normalizeXlsxColor, normalizeXlsxThemeConfig, xlsxColorCss, xlsxFillSvgPaint } from "./spreadsheet/ooxml-styles.mjs";
import { planSpreadsheetThreadedComments, validateSpreadsheetThreadedCommentPackageSemantics } from "./spreadsheet/ooxml-threaded-comments.mjs";
import { parseStructuredReference, scanStructuredReferenceIntersections, scanStructuredReferences, splitReferenceIntersectionOperands } from "./spreadsheet/structured-references.mjs";
import {
  assertWorksheetBounds,
  columnLabelToNumber,
  columnNumberToLabel,
  formulaA1ToR1C1,
  formulaR1C1ToA1,
  makeCellAddress,
  parseCellAddress,
  parseRangeAddress,
  rangeBounds,
  rangeToAddress,
  translateA1Formula,
  XLSX_MAX_COLUMN_INDEX,
  XLSX_MAX_ROW_INDEX,
} from "./spreadsheet/range-addressing.mjs";
import {
  assertRangeCopyShape,
  currentRegionBounds,
  normalizeRangeCopyMode,
  normalizeRangeWrite,
  writtenRangeBounds,
} from "./spreadsheet/range-operations.mjs";
import { normalizePresentationThemeConfig } from "./presentation/ooxml-theme.mjs";
import { mergePresentationPlaceholders, normalizePresentationBackground, resolvePresentationBackgroundColor } from "./presentation/ooxml-masters.mjs";
import { createPresentationGroupShapeClass } from "./presentation/group-shapes.mjs";
import { normalizePresentationChartAxisGroup, normalizePresentationChartDataLabels, normalizePresentationChartErrorBars, normalizePresentationChartSeriesStyle, normalizePresentationChartStyle, normalizePresentationChartTrendlines } from "./presentation/ooxml-charts.mjs";
import { normalizePresentationChartExternalData, presentationChartUsesFormulaReferences } from "./presentation/ooxml-chart-data.mjs";
import { presentationChartLineSvgAttributes, presentationChartTrendlinesSvg } from "./presentation/chart-trendline-svg.mjs";
import { planPresentationCustomShows, PresentationCustomShowCollection } from "./presentation/ooxml-custom-shows.mjs";
import { inheritPresentationParagraphs, normalizePresentationParagraphs, normalizePresentationParagraphStyles, presentationParagraphsNeedSerialization, presentationParagraphsSvg, presentationParagraphsText, replacePresentationParagraphText } from "./presentation/text-paragraphs.mjs";
import { normalizePresentationTextBodyProperties } from "./presentation/text-body-properties.mjs";
import { normalizePresentationCustomPaths, presentationCustomPathsSvg } from "./presentation/custom-geometry.mjs";
import { PPTX_MODERN_AUTHOR_CONTENT_TYPE, PPTX_MODERN_AUTHOR_RELATIONSHIP_TYPE, PPTX_MODERN_COMMENT_CONTENT_TYPE, PPTX_MODERN_COMMENT_RELATIONSHIP_TYPE, planPresentationModernComments } from "./presentation/ooxml-modern-comments.mjs";
import { PdfArtifact, PdfFile } from "./pdf/index.mjs";
import { formulaTimeParts, formulaTimeSerial, parseFormulaDateText, parseFormulaNumberText, parseFormulaTimeText } from "./spreadsheet/formula-coercion.mjs";
import { createWorkbookWindowCollection, worksheetWindowMemberships } from "./spreadsheet/workbook-windows.mjs";
import { decoder, encoder, toUint8Array } from "./shared/binary.mjs";
import { FileBlob } from "./shared/file-blob.mjs";
import { aid } from "./shared/ids.mjs";
import { imageContentTypeFromExtension, imageDataFromDataUrl } from "./shared/images.mjs";
import { filterInspectRecords, inspectRecordMatchesTarget, inspectTargetTokens, ndjson, normalizeKinds, verificationIssue, verificationResult } from "./shared/inspection.mjs";
import { decodePngRgba, isPngBytes } from "./shared/png.mjs";
import { fileBlobFromRenderOutput, LAYOUT_MIME, renderTypeForOptions } from "./shared/render-output.mjs";
import { attrEscape, xmlEscape } from "./shared/xml.mjs";
import { queryHelpRecords } from "./help/index.mjs";
import { materializeComposeNode } from "./presentation/compose.mjs";

export { FileBlob } from "./shared/file-blob.mjs";
export { HELP_CATALOG } from "./help/index.mjs";
export { box, chart, column, grid, image, layers, node, paragraph, row, rule, run, shape, table, text } from "./presentation/compose.mjs";
export { PdfArtifact, PdfFile };

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const XLSX_DYNAMIC_ARRAY_METADATA_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
const XLSX_DYNAMIC_ARRAY_METADATA_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
const XLSX_DYNAMIC_ARRAY_METADATA_EXTENSION_URI = "{BDBB8CDC-FA1E-496E-A857-3C3F30C029C3}";
const CSV_MIME = "text/csv";
const TSV_MIME = "text/tab-separated-values";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const PPTX_PACKAGE_CONFIG = {
  family: "PPTX",
  packageKind: "pptxPackage",
  partKind: "pptxPart",
  counts: { slides: /^ppt\/slides\/slide\d+\.xml$/ },
  semanticIssues: validatePptxPackageSemantics,
};

const DOCX_PACKAGE_CONFIG = {
  family: "DOCX",
  packageKind: "docxPackage",
  partKind: "docxPart",
  semanticIssues: (context) => [...validateDocxCommentPackageSemantics(context), ...validateDocxLinkPackageSemantics(context), ...validateDocxBibliographyPackageSemantics(context)],
};

const XLSX_PACKAGE_CONFIG = {
  family: "XLSX",
  packageKind: "xlsxPackage",
  partKind: "xlsxPart",
  counts: { sheets: /^xl\/worksheets\/sheet\d+\.xml$/ },
  semanticIssues: validateSpreadsheetThreadedCommentPackageSemantics,
};

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

function isJpegBytes(bytes) {
  return bytes?.[0] === 0xff && bytes?.[1] === 0xd8 && bytes?.[2] === 0xff;
}

function isWebpBytes(bytes) {
  return bytes?.[0] === 0x52 && bytes?.[1] === 0x49 && bytes?.[2] === 0x46 && bytes?.[3] === 0x46
    && bytes?.[8] === 0x57 && bytes?.[9] === 0x45 && bytes?.[10] === 0x42 && bytes?.[11] === 0x50;
}

function rasterByteFormat(bytes) {
  if (isPngBytes(bytes)) return "png";
  if (isPpmBytes(bytes)) return "ppm";
  if (isJpegBytes(bytes)) return "jpeg";
  if (isWebpBytes(bytes)) return "webp";
  return undefined;
}

async function decodeRasterRgba(bytes, options = {}) {
  const format = rasterByteFormat(bytes);
  if (format === "png") return decodePngRgba(bytes);
  if (format === "ppm") return decodePpmRgba(bytes);
  if (format !== "jpeg" && format !== "webp") throw new Error("unsupported raster format");
  let sharp;
  try {
    const module = await import("sharp");
    sharp = module.default || module;
  } catch (error) {
    throw new Error(`JPEG/WebP pixel diff requires the optional peer dependency \"sharp\": ${error.message}`);
  }
  const maxPixels = Math.max(1, Number(options.maxDecodedPixels ?? options.maxPixels ?? 40_000_000));
  const decoded = await sharp(Buffer.from(bytes), { failOn: "error", limitInputPixels: maxPixels })
    .ensureAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });
  if (!decoded.info?.width || !decoded.info?.height || decoded.info.channels !== 4) throw new Error(`Sharp did not return RGBA pixels for ${format}`);
  return { width: decoded.info.width, height: decoded.info.height, pixels: new Uint8Array(decoded.data) };
}

async function compareRasterPixels(bytes, baselineBytes, options = {}) {
  const actualFormat = rasterByteFormat(bytes);
  const baselineFormat = rasterByteFormat(baselineBytes);
  if (!actualFormat || !baselineFormat) throw new Error("pixelDiff supports PNG, JPEG, WebP, and PPM raster baselines");
  const [actual, expected] = await Promise.all([decodeRasterRgba(bytes, options), decodeRasterRgba(baselineBytes, options)]);
  const threshold = Math.max(0, Number(options.pixelThreshold ?? options.threshold ?? 0));
  const format = actualFormat === baselineFormat ? actualFormat : `${actualFormat}/${baselineFormat}`;
  const metrics = compareRgbaPixels(actual, expected, { ...options, threshold, format, actualFormat, baselineFormat });
  let diffBytes;
  if (metrics.changed && metrics.diffPixels) diffBytes = encodePngRgba(metrics.diffWidth || metrics.width, metrics.diffHeight || metrics.height, metrics.diffPixels);
  delete metrics.diffPixels;
  return { metrics, diffBytes };
}

function pngCrc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data = new Uint8Array()) {
  const typeBytes = encoder.encode(type);
  const payload = Buffer.from(data);
  const chunk = Buffer.alloc(12 + payload.length);
  chunk.writeUInt32BE(payload.length, 0);
  Buffer.from(typeBytes).copy(chunk, 4);
  payload.copy(chunk, 8);
  chunk.writeUInt32BE(pngCrc32(new Uint8Array(chunk.subarray(4, 8 + payload.length))), 8 + payload.length);
  return chunk;
}

function encodePngRgba(width, height, pixels) {
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const rowBytes = width * 4;
  const raw = Buffer.alloc((rowBytes + 1) * height);
  for (let row = 0; row < height; row += 1) Buffer.from(pixels.subarray(row * rowBytes, (row + 1) * rowBytes)).copy(raw, row * (rowBytes + 1) + 1);
  return new Uint8Array(Buffer.concat([signature, pngChunk("IHDR", ihdr), pngChunk("IDAT", deflateSync(raw)), pngChunk("IEND")]));
}

function qaRgbColor(value, fallback) {
  if (typeof value === "string") {
    const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(value);
    if (short) return short.slice(1).map((part) => Number.parseInt(part + part, 16));
    const hex = /^#([0-9a-f]{6})$/i.exec(value)?.[1];
    if (hex) return [0, 2, 4].map((offset) => Number.parseInt(hex.slice(offset, offset + 2), 16));
  }
  if (Array.isArray(value) && value.length >= 3) return value.slice(0, 3).map((part) => Math.max(0, Math.min(255, Math.round(Number(part) || 0))));
  return fallback;
}

function rgbaAt(image, canvasX, canvasY, offsetX, offsetY) {
  const x = canvasX - offsetX;
  const y = canvasY - offsetY;
  if (x < 0 || y < 0 || x >= image.width || y >= image.height) return [0, 0, 0, 0];
  const index = (y * image.width + x) * 4;
  return image.pixels.subarray(index, index + 4);
}

function rgbaPixelChanged(actual, expected, canvasX, canvasY, actualOffsetX, actualOffsetY, expectedOffsetX, expectedOffsetY, threshold) {
  const actualX = canvasX - actualOffsetX;
  const actualY = canvasY - actualOffsetY;
  const expectedX = canvasX - expectedOffsetX;
  const expectedY = canvasY - expectedOffsetY;
  const actualIndex = actualX >= 0 && actualY >= 0 && actualX < actual.width && actualY < actual.height ? (actualY * actual.width + actualX) * 4 : -1;
  const expectedIndex = expectedX >= 0 && expectedY >= 0 && expectedX < expected.width && expectedY < expected.height ? (expectedY * expected.width + expectedX) * 4 : -1;
  for (let channel = 0; channel < 4; channel += 1) {
    const actualValue = actualIndex < 0 ? 0 : actual.pixels[actualIndex + channel];
    const expectedValue = expectedIndex < 0 ? 0 : expected.pixels[expectedIndex + channel];
    if (Math.abs(actualValue - expectedValue) > threshold) return true;
  }
  return false;
}

function countRgbaMismatches(actual, expected, geometry, threshold, expectedShiftX = 0, expectedShiftY = 0, stride = 1) {
  let differentPixels = 0;
  let sampledPixels = 0;
  for (let y = 0; y < geometry.canvasHeight; y += stride) {
    for (let x = 0; x < geometry.canvasWidth; x += stride) {
      const expectedX = x - geometry.expectedOffsetX - expectedShiftX;
      const expectedY = y - geometry.expectedOffsetY - expectedShiftY;
      if ((expectedShiftX || expectedShiftY) && (expectedX < 0 || expectedY < 0 || expectedX >= expected.width || expectedY >= expected.height)) continue;
      sampledPixels += 1;
      if (rgbaPixelChanged(actual, expected, x, y, geometry.actualOffsetX, geometry.actualOffsetY, geometry.expectedOffsetX + expectedShiftX, geometry.expectedOffsetY + expectedShiftY, threshold)) differentPixels += 1;
    }
  }
  return { differentPixels, sampledPixels };
}

function pixelRegistrationConfig(options = {}) {
  const request = options.pixelRegistration ?? options.registration;
  if (!request) return undefined;
  const config = typeof request === "object" ? request : {};
  const requestedOffset = request === true ? 2 : typeof request === "number" ? request : config.maxOffset ?? config.maxPixels ?? 2;
  const maxOffset = Math.max(0, Math.min(8, Math.floor(Number(requestedOffset) || 0)));
  if (!maxOffset) return undefined;
  return {
    maxOffset,
    minImprovementRatio: Math.max(0, Math.min(1, Number(config.minImprovementRatio ?? config.minImprovement ?? 0.05))),
    maxSamples: Math.max(1_000, Math.min(1_000_000, Math.floor(Number(config.maxSamples ?? config.samples ?? 100_000) || 100_000))),
  };
}

function findPixelRegistration(actual, expected, geometry, threshold, config) {
  const candidateCount = (config.maxOffset * 2 + 1) ** 2;
  const samplesPerCandidate = Math.max(16, Math.floor(config.maxSamples / candidateCount));
  const stride = Math.max(1, Math.ceil(Math.sqrt((geometry.canvasWidth * geometry.canvasHeight) / samplesPerCandidate)));
  const baseline = countRgbaMismatches(actual, expected, geometry, threshold, 0, 0, stride);
  let best = { x: 0, y: 0, ...baseline };
  for (let y = -config.maxOffset; y <= config.maxOffset; y += 1) {
    for (let x = -config.maxOffset; x <= config.maxOffset; x += 1) {
      if (x === 0 && y === 0) continue;
      const candidate = countRgbaMismatches(actual, expected, geometry, threshold, x, y, stride);
      const candidateRatio = candidate.sampledPixels ? candidate.differentPixels / candidate.sampledPixels : 1;
      const bestRatio = best.sampledPixels ? best.differentPixels / best.sampledPixels : 1;
      const candidateDistance = Math.abs(x) + Math.abs(y);
      const bestDistance = Math.abs(best.x) + Math.abs(best.y);
      if (candidateRatio < bestRatio || (candidateRatio === bestRatio && candidateDistance < bestDistance)) best = { x, y, ...candidate };
    }
  }
  const baselineRatio = baseline.sampledPixels ? baseline.differentPixels / baseline.sampledPixels : 0;
  const bestRatio = best.sampledPixels ? best.differentPixels / best.sampledPixels : 1;
  const improvementRatio = baselineRatio ? (baselineRatio - bestRatio) / baselineRatio : 0;
  const applied = (best.x !== 0 || best.y !== 0) && improvementRatio >= config.minImprovementRatio;
  return {
    requested: true,
    applied,
    maxOffset: config.maxOffset,
    minImprovementRatio: config.minImprovementRatio,
    maxSamples: config.maxSamples,
    candidateCount,
    samplesPerCandidate,
    estimatedComparisons: baseline.sampledPixels * candidateCount,
    sampleStride: stride,
    sampledPixels: baseline.sampledPixels,
    sampledPixelsAfter: applied ? best.sampledPixels : baseline.sampledPixels,
    sampledDifferentPixelsBefore: baseline.differentPixels,
    sampledDifferentPixelsAfter: applied ? best.differentPixels : baseline.differentPixels,
    sampledImprovementRatio: applied ? improvementRatio : 0,
    offset: applied ? { x: best.x, y: best.y } : { x: 0, y: 0 },
  };
}

function compareRgbaPixels(actual, expected, options = {}) {
  const threshold = Math.max(0, Number(options.threshold ?? options.pixelThreshold ?? 0));
  const requestedAlignment = String(options.diffAlignment ?? options.alignment ?? "strict").trim().toLowerCase();
  const alignment = ["strict", "top-left", "center"].includes(requestedAlignment) ? requestedAlignment : "strict";
  const dimensionMismatch = actual.width !== expected.width || actual.height !== expected.height;
  const requestedRegistrationConfig = pixelRegistrationConfig(options);
  const canvasWidth = dimensionMismatch && alignment !== "strict" ? Math.max(actual.width, expected.width) : actual.width;
  const canvasHeight = dimensionMismatch && alignment !== "strict" ? Math.max(actual.height, expected.height) : actual.height;
  const result = {
    format: options.format || "rgba",
    actualFormat: options.actualFormat,
    baselineFormat: options.baselineFormat,
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
    alignment,
  };
  if (dimensionMismatch && alignment === "strict") {
    result.dimensionMismatch = true;
    result.differentPixels = Math.max(result.pixels, expected.width * expected.height);
    result.mismatchRatio = 1;
    result.changed = true;
    if (requestedRegistrationConfig) result.registration = { requested: true, applied: false, skipped: "dimensionMismatch", ...requestedRegistrationConfig };
    return result;
  }
  if (dimensionMismatch) {
    result.dimensionMismatch = true;
    result.diffWidth = canvasWidth;
    result.diffHeight = canvasHeight;
    result.comparisonPixels = canvasWidth * canvasHeight;
    result.pixels = result.comparisonPixels;
  }
  const actualOffsetX = alignment === "center" ? Math.floor((canvasWidth - actual.width) / 2) : 0;
  const actualOffsetY = alignment === "center" ? Math.floor((canvasHeight - actual.height) / 2) : 0;
  const baseExpectedOffsetX = alignment === "center" ? Math.floor((canvasWidth - expected.width) / 2) : 0;
  const baseExpectedOffsetY = alignment === "center" ? Math.floor((canvasHeight - expected.height) / 2) : 0;
  const geometry = { canvasWidth, canvasHeight, actualOffsetX, actualOffsetY, expectedOffsetX: baseExpectedOffsetX, expectedOffsetY: baseExpectedOffsetY };
  const registrationConfig = dimensionMismatch ? undefined : requestedRegistrationConfig;
  const registration = registrationConfig ? findPixelRegistration(actual, expected, geometry, threshold, registrationConfig) : undefined;
  const expectedOffsetX = baseExpectedOffsetX + (registration?.offset.x || 0);
  const expectedOffsetY = baseExpectedOffsetY + (registration?.offset.y || 0);
  const diffPixels = options.diffImage === false ? undefined : new Uint8Array(canvasWidth * canvasHeight * 4);
  const palette = options.diffPalette || options.palette || {};
  const changedColor = qaRgbColor(palette.changed, [255, 24, 72]);
  const unchangedColor = palette.unchanged == null ? undefined : qaRgbColor(palette.unchanged, [64, 64, 64]);
  const changedAlpha = Math.max(0, Math.min(255, Math.round(Number(palette.changedAlpha ?? 255))));
  const unchangedAlpha = Math.max(0, Math.min(255, Math.round(Number(palette.unchangedAlpha ?? 255))));
  let changedPixels = 0;
  let channelDeltaSum = 0;
  let registrationIgnoredPixels = 0;
  for (let y = 0; y < canvasHeight; y += 1) for (let x = 0; x < canvasWidth; x += 1) {
    const actualPixel = rgbaAt(actual, x, y, actualOffsetX, actualOffsetY);
    const expectedX = x - expectedOffsetX;
    const expectedY = y - expectedOffsetY;
    const ignoredRegistrationEdge = registration?.applied && (expectedX < 0 || expectedY < 0 || expectedX >= expected.width || expectedY >= expected.height);
    const expectedPixel = ignoredRegistrationEdge ? actualPixel : rgbaAt(expected, x, y, expectedOffsetX, expectedOffsetY);
    const i = (y * canvasWidth + x) * 4;
    let pixelChanged = false;
    if (ignoredRegistrationEdge) registrationIgnoredPixels += 1;
    for (let c = 0; c < 4; c += 1) {
      const delta = Math.abs(actualPixel[c] - expectedPixel[c]);
      channelDeltaSum += delta;
      if (delta > result.maxChannelDelta) result.maxChannelDelta = delta;
      if (delta > threshold) pixelChanged = true;
    }
    if (diffPixels) {
      const brightness = Math.round((actualPixel[0] + actualPixel[1] + actualPixel[2]) / 3 * 0.35 + 32);
      const color = pixelChanged ? changedColor : unchangedColor || [brightness, brightness, brightness];
      diffPixels[i] = color[0];
      diffPixels[i + 1] = color[1];
      diffPixels[i + 2] = color[2];
      diffPixels[i + 3] = pixelChanged ? changedAlpha : unchangedAlpha;
    }
    if (pixelChanged) changedPixels += 1;
  }
  result.differentPixels = changedPixels;
  const comparedPixels = Math.max(0, result.pixels - registrationIgnoredPixels);
  result.mismatchRatio = comparedPixels ? changedPixels / comparedPixels : 0;
  result.meanChannelDelta = comparedPixels ? channelDeltaSum / (comparedPixels * 4) : 0;
  result.diffPalette = { changed: changedColor, changedAlpha, unchanged: unchangedColor || "actual-grayscale", unchangedAlpha };
  if (registration) {
    const before = registration.applied ? countRgbaMismatches(actual, expected, geometry, threshold) : { differentPixels: changedPixels, sampledPixels: result.pixels };
    result.registration = {
      ...registration,
      differentPixelsBefore: before.differentPixels,
      mismatchRatioBefore: result.pixels ? before.differentPixels / result.pixels : 0,
      differentPixelsAfter: changedPixels,
      mismatchRatioAfter: result.mismatchRatio,
      comparedPixelsAfter: comparedPixels,
      ignoredEdgePixels: registrationIgnoredPixels,
    };
  } else if (requestedRegistrationConfig && dimensionMismatch) result.registration = { requested: true, applied: false, skipped: "dimensionMismatch", ...requestedRegistrationConfig };
  result.changed = changedPixels > 0;
  if (result.changed && diffPixels) result.diffPixels = diffPixels;
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
  let diffBlob;
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
    summary.byteChanged = baselineHash !== hash;
    summary.changed = summary.byteChanged;
    const pixelDiffEnabled = options.pixelDiff === true || typeof options.pixelDiff === "object";
    if (pixelDiffEnabled) {
      if (rasterByteFormat(bytes) && rasterByteFormat(baselineBytes)) {
        try {
          const pixelDiffOptions = typeof options.pixelDiff === "object" ? { ...options, ...options.pixelDiff } : options;
          const compared = await compareRasterPixels(bytes, baselineBytes, pixelDiffOptions);
          const pixelDiff = compared.metrics;
          summary.pixelDiff = pixelDiff;
          summary.changed = pixelDiff.changed;
          if (compared.diffBytes) {
            diffBlob = new FileBlob(compared.diffBytes, { type: "image/png", metadata: { artifactKind, format: "pixel-diff", actualFormat: pixelDiff.actualFormat, baselineFormat: pixelDiff.baselineFormat, alignment: pixelDiff.alignment, registration: pixelDiff.registration, width: pixelDiff.diffWidth || pixelDiff.width, height: pixelDiff.diffHeight || pixelDiff.height, palette: pixelDiff.diffPalette } });
            summary.diff = { type: diffBlob.type, bytes: diffBlob.bytes.length, hash: stableByteHash(diffBlob.bytes) };
          }
          if (pixelDiff.changed && options.allowChange !== true && options.allowPixelChange !== true) {
            issues.push(verificationIssue(artifactKind, "visualPixelDiff", `Rendered ${pixelDiff.format.toUpperCase()} differs from the baseline in ${pixelDiff.differentPixels} pixels.`, { severity: options.diffSeverity || "warning", ...pixelDiff }));
          }
        } catch (error) {
          summary.pixelDiff = { skipped: true, reason: error.message };
        }
      } else {
        summary.pixelDiff = { skipped: true, reason: "pixelDiff supports PNG, JPEG, WebP, and PPM raster baselines" };
      }
    }
    const pixelsEquivalent = summary.pixelDiff && !summary.pixelDiff.skipped && summary.pixelDiff.changed === false;
    if (baselineHash !== hash && !pixelsEquivalent && options.allowChange !== true) issues.push(verificationIssue(artifactKind, "visualDiff", "Rendered output differs from the supplied baseline.", { severity: options.diffSeverity || "warning", hash, baselineHash }));
  }
  const records = [summary, ...issues];
  return { artifactKind, ok: issues.length === 0, blob, diffBlob, summary, issues, ...ndjson(records, options.maxChars ?? Infinity) };
}

export function helpArtifact(artifactOrKind = "*", query = "*", options = {}) {
  const artifactKind = typeof artifactOrKind === "string" ? artifactOrKind : inferArtifactKind(artifactOrKind);
  return ndjson(queryHelpRecords(artifactKind, query, options), options.maxChars ?? Infinity);
}


class WorksheetCollection {
  constructor(workbook) {
    this.workbook = workbook;
    this.items = [];
  }

  add(name = `Sheet${this.items.length + 1}`, options = {}) {
    if (name && typeof name === "object") {
      options = name;
      name = options.name || `Sheet${this.items.length + 1}`;
    }
    const worksheet = new Worksheet(this.workbook, name, options);
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

  _resolveWorksheet(nameOrIndexOrWorksheet, label = "worksheet") {
    const worksheet = typeof nameOrIndexOrWorksheet === "number"
      ? this.getItem(nameOrIndexOrWorksheet)
      : typeof nameOrIndexOrWorksheet === "string"
        ? this.getItem(nameOrIndexOrWorksheet) || this.items.find((sheet) => sheet.id === nameOrIndexOrWorksheet)
        : nameOrIndexOrWorksheet;
    if (!worksheet || worksheet.workbook !== this.workbook || !this.items.includes(worksheet))
      throw new Error(`Workbook ${label} ${String(nameOrIndexOrWorksheet)} was not found.`);
    return worksheet;
  }

  getActiveWorksheet() {
    return this.workbook.windows.getItemAt(0).getActiveWorksheet();
  }

  setActiveWorksheet(nameOrIndexOrWorksheet) {
    return this.workbook.windows.getItemAt(0).setActiveWorksheet(nameOrIndexOrWorksheet);
  }

  getSelectedWorksheets() {
    return this.workbook.windows.getItemAt(0).getSelectedWorksheets();
  }

  setSelectedWorksheets(worksheets) {
    return this.workbook.windows.getItemAt(0).setSelectedWorksheets(worksheets);
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
  constructor(workbook, target, text, author, config = {}) {
    this.workbook = workbook;
    this.id = config.id || aid("th");
    this.target = workbookRangeRef(target);
    this.author = author || workbook.comments.self?.displayName || "User";
    this.comments = [{ author: this.author, text: String(text ?? ""), ...(config.comment || {}) }];
    this.resolved = Boolean(config.resolved);
  }

  addReply(text, config = {}) {
    this.comments.push({ author: config.author || this.workbook.comments.self?.displayName || this.author, text: String(text ?? ""), ...config });
    return this;
  }

  resolve() { this.resolved = true; this.comments.forEach((comment) => { comment.done = true; }); return this; }
  reopen() { this.resolved = false; this.comments.forEach((comment) => { comment.done = false; }); return this; }

  inspectRecord() {
    return { kind: "thread", id: this.id, sheet: this.target.sheetName, address: this.target.address, author: this.author, resolved: this.resolved, replies: Math.max(0, this.comments.length - 1), commentIds: this.comments.map((comment) => comment.id).filter(Boolean), personIds: this.comments.map((comment) => comment.personId).filter(Boolean), dates: this.comments.map((comment) => comment.date).filter(Boolean), textPreview: this.comments.map((comment) => comment.text).join("\n").slice(0, 300) };
  }

  toJSON() { return { id: this.id, target: this.target, author: this.author, comments: this.comments, resolved: this.resolved }; }
}

class CommentsCollection {
  constructor(workbook) { this.workbook = workbook; this.self = undefined; this.threads = []; }
  setSelf(self) { this.self = { displayName: self?.displayName || "User" }; return this.self; }
  addThread(target, text, config = {}) { const thread = new CommentThread(this.workbook, target, text, config.author || this.self?.displayName, config); this.threads.push(thread); return thread; }
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
    if (config.hidden !== undefined) this.hidden = Boolean(config.hidden);
  }
  inspectRecord() { return { kind: "definedName", id: this.id, name: this.name, refersTo: this.refersTo, scope: this.scope, comment: this.comment, hidden: this.hidden }; }
  toJSON() { return { id: this.id, name: this.name, refersTo: this.refersTo, scope: this.scope, comment: this.comment, hidden: this.hidden }; }
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

const XLSX_MAX_COLUMN_WIDTH = 255;
const XLSX_MAX_ROW_HEIGHT = 409.5;
const XLSX_DEFAULT_COLUMN_WIDTH = 8.43;
const XLSX_DEFAULT_ROW_HEIGHT = 15;
const XLSX_MAX_DIGIT_WIDTH_PX = 7;

function xlsxBoolean(value) {
  return ["1", "true", "on"].includes(String(value ?? "").trim().toLowerCase());
}

function xlsxColumnCharactersToWidth(value) {
  const characters = Number(value);
  if (!Number.isFinite(characters) || characters <= 0 || characters > XLSX_MAX_COLUMN_WIDTH) throw new Error(`Worksheet column width must be greater than 0 and at most ${XLSX_MAX_COLUMN_WIDTH}.`);
  return Math.min(XLSX_MAX_COLUMN_WIDTH, Math.trunc(((characters * XLSX_MAX_DIGIT_WIDTH_PX + 5) / XLSX_MAX_DIGIT_WIDTH_PX) * 256) / 256);
}

function xlsxColumnWidthToPixels(width) {
  return Math.trunc(((256 * Number(width) + Math.trunc(128 / XLSX_MAX_DIGIT_WIDTH_PX)) / 256) * XLSX_MAX_DIGIT_WIDTH_PX);
}

function xlsxColumnPixelsToWidth(pixelsValue) {
  const pixels = Number(pixelsValue);
  if (!Number.isFinite(pixels) || pixels <= 0 || pixels > 1_790) throw new Error("Worksheet column pixel width must be greater than 0 and at most 1790.");
  const characters = Math.trunc((((pixels - 5) / XLSX_MAX_DIGIT_WIDTH_PX) * 100 + 0.5)) / 100;
  return xlsxColumnCharactersToWidth(Math.max(0.01, characters));
}

function xlsxColumnWidthToCharacters(width) {
  const pixels = xlsxColumnWidthToPixels(width);
  return Math.max(0, Math.trunc((((pixels - 5) / XLSX_MAX_DIGIT_WIDTH_PX) * 100 + 0.5)) / 100);
}

function xlsxRowHeight(value) {
  const height = Number(value);
  if (!Number.isFinite(height) || height <= 0 || height > XLSX_MAX_ROW_HEIGHT) throw new Error(`Worksheet row height must be greater than 0 and at most ${XLSX_MAX_ROW_HEIGHT}.`);
  return Math.round(height * 100) / 100;
}

function worksheetColumnDimension(sheet, column) {
  return sheet.columnDimensions?.get(column) || {};
}

function worksheetRowDimension(sheet, row) {
  return sheet.rowDimensions?.get(row) || {};
}

function worksheetColumnWidthPx(sheet, column, fallback = undefined, collapseHidden = true) {
  const dimension = worksheetColumnDimension(sheet, column);
  if (collapseHidden && dimension.hidden) return 0;
  if (dimension.width != null) return xlsxColumnWidthToPixels(dimension.width);
  return Number(fallback ?? xlsxColumnWidthToPixels(xlsxColumnCharactersToWidth(XLSX_DEFAULT_COLUMN_WIDTH)));
}

function worksheetRowHeightPx(sheet, row, fallback = undefined, collapseHidden = true) {
  const dimension = worksheetRowDimension(sheet, row);
  if (collapseHidden && dimension.hidden) return 0;
  if (dimension.height != null) return dimension.height * 96 / 72;
  return Number(fallback ?? XLSX_DEFAULT_ROW_HEIGHT * 96 / 72);
}

function worksheetAxisOffset(sheet, axis, index, start = 0, fallback = undefined) {
  let offset = 0;
  for (let current = start; current < index; current += 1) offset += axis === "column" ? worksheetColumnWidthPx(sheet, current, fallback) : worksheetRowHeightPx(sheet, current, fallback);
  return offset;
}

function worksheetCellAtPixel(sheet, axis, pixelValue) {
  let remaining = Math.max(0, Number(pixelValue) - 40);
  const maximum = axis === "column" ? XLSX_MAX_FREEZE_COLUMNS : XLSX_MAX_FREEZE_ROWS;
  for (let index = 0; index <= maximum; index += 1) {
    const size = axis === "column" ? worksheetColumnWidthPx(sheet, index) : worksheetRowHeightPx(sheet, index);
    if (size > 0 && remaining < size) return { index, offset: remaining };
    if (size > 0) remaining -= size;
  }
  return { index: maximum, offset: 0 };
}

function worksheetRangeFrame(sheet, rangeBounds, viewportBounds = { top: 0, left: 0 }, options = {}) {
  const defaultColumnWidth = options.cellWidthPx ?? options.cellW;
  const defaultRowHeight = options.cellHeightPx ?? options.cellH;
  const left = 40 + worksheetAxisOffset(sheet, "column", rangeBounds.left, viewportBounds.left, defaultColumnWidth);
  const top = 40 + worksheetAxisOffset(sheet, "row", rangeBounds.top, viewportBounds.top, defaultRowHeight);
  let width = 0;
  let height = 0;
  for (let column = rangeBounds.left; column <= rangeBounds.right; column += 1) width += worksheetColumnWidthPx(sheet, column, defaultColumnWidth);
  for (let row = rangeBounds.top; row <= rangeBounds.bottom; row += 1) height += worksheetRowHeightPx(sheet, row, defaultRowHeight);
  return { left, top, width, height };
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

const WORKSHEET_TABLE_QUERY_UNSUPPORTED = Symbol("open-office-artifact-tool.worksheet-table-query-unsupported");
const WORKSHEET_TABLE_QUERY_BOOLEAN_FIELDS = [
  "headers", "rowNumbers", "disableRefresh", "backgroundRefresh", "firstBackgroundRefresh", "refreshOnLoad",
  "fillFormulas", "removeDataOnSave", "disableEdit", "preserveFormatting", "adjustColumnWidth", "intermediate",
  "applyNumberFormats", "applyBorderFormats", "applyFontFormats", "applyPatternFormats", "applyAlignmentFormats",
  "applyWidthHeightFormats",
];

const WORKSHEET_TABLE_QUERY_REFRESH_BOOLEAN_FIELDS = ["preserveSortFilterLayout", "fieldIdWrapped", "headersInLastRefresh"];
const WORKSHEET_TABLE_QUERY_REFRESH_UINT_FIELDS = ["minimumVersion", "nextId", "unboundColumnsLeft", "unboundColumnsRight"];
const WORKSHEET_TABLE_QUERY_FIELD_BOOLEAN_FIELDS = ["dataBound", "rowNumbers", "fillFormulas", "clipped"];

function normalizeWorksheetTableSortState(sortState) {
  if (!sortState) return undefined;
  return {
    reference: String(sortState.reference ?? ""),
    caseSensitive: Boolean(sortState.caseSensitive),
    ...(sortState.sortMethod == null ? {} : { sortMethod: String(sortState.sortMethod) }),
    ...(sortState.columnSort == null ? {} : { columnSort: Boolean(sortState.columnSort) }),
    conditions: Array.isArray(sortState.conditions)
      ? sortState.conditions.map((condition) => ({
          reference: String(condition?.reference ?? ""),
          descending: Boolean(condition?.descending),
          ...((condition?.kind === "icon" || condition?.iconSet) ? {
            kind: "icon",
            iconSet: String(condition.iconSet ?? ""),
            ...(condition.iconId == null ? {} : { iconId: Number(condition.iconId) }),
          } : condition?.kind === "color" ? {
            kind: "color",
            target: String(condition.target ?? ""),
            color: condition.color,
          } : condition?.customList == null ? {} : { customList: String(condition.customList) }),
        }))
      : [],
  };
}

function normalizeWorksheetTableQueryRefresh(refresh) {
  if (refresh == null) return undefined;
  const normalized = {
    fields: Array.isArray(refresh.fields) ? refresh.fields.map((field) => {
      const result = { id: Number(field?.id ?? 0) };
      if (field?.name !== undefined) result.name = String(field.name);
      for (const name of WORKSHEET_TABLE_QUERY_FIELD_BOOLEAN_FIELDS) if (field?.[name] !== undefined) result[name] = Boolean(field[name]);
      if (field?.tableColumnId !== undefined) result.tableColumnId = Number(field.tableColumnId);
      return result;
    }) : [],
  };
  for (const field of WORKSHEET_TABLE_QUERY_REFRESH_BOOLEAN_FIELDS) if (refresh[field] !== undefined) normalized[field] = Boolean(refresh[field]);
  for (const field of WORKSHEET_TABLE_QUERY_REFRESH_UINT_FIELDS) if (refresh[field] !== undefined) normalized[field] = Number(refresh[field]);
  if (Array.isArray(refresh.deletedFieldNames)) normalized.deletedFieldNames = refresh.deletedFieldNames.map((name) => String(name));
  if (refresh.sortState) normalized.sortState = normalizeWorksheetTableSortState(refresh.sortState);
  return normalized;
}

function normalizeWorksheetTableQuery(query) {
  if (query == null) return undefined;
  const normalized = { name: String(query.name ?? ""), connectionId: Number(query.connectionId ?? 0) };
  for (const field of WORKSHEET_TABLE_QUERY_BOOLEAN_FIELDS) if (query[field] !== undefined) normalized[field] = Boolean(query[field]);
  if (query.growShrinkType !== undefined) normalized.growShrinkType = String(query.growShrinkType);
  if (query.autoFormatId !== undefined) normalized.autoFormatId = Number(query.autoFormatId);
  if (query.refresh != null) normalized.refresh = normalizeWorksheetTableQueryRefresh(query.refresh);
  return normalized;
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
    this.showHeaders = config.showHeaders ?? this.hasHeaders;
    this.showTotals = Boolean(config.showTotals);
    this.showBandedColumns = Boolean(config.showBandedColumns);
    this.showFilterButton = config.showFilterButton ?? true;
    this.showFirstColumn = Boolean(config.showFirstColumn);
    this.showLastColumn = Boolean(config.showLastColumn);
    this.showRowStripes = config.showRowStripes ?? this.showHeaders;
    this.style = config.style || "TableStyleMedium2";
    this.values = config.values ? config.values.map((row) => [...row]) : range.values.map((row) => [...row]);
    const configuredColumns = Array.isArray(config.columnDefinitions) ? config.columnDefinitions : Array.isArray(config.columns) ? config.columns : undefined;
    this.columnDefinitions = configuredColumns?.map((column, index) => ({
      name: String(column?.name ?? config.columnNames?.[index] ?? `Column${index + 1}`),
      calculatedColumnFormula: column?.calculatedColumnFormula ? String(column.calculatedColumnFormula) : "",
      calculatedColumnFormulaArray: Boolean(column?.calculatedColumnFormulaArray),
      totalsRowFunction: column?.totalsRowFunction ? String(column.totalsRowFunction) : "",
      totalsRowLabel: column?.totalsRowLabel ? String(column.totalsRowLabel) : "",
      totalsRowFormula: column?.totalsRowFormula ? String(column.totalsRowFormula) : "",
      totalsRowFormulaArray: Boolean(column?.totalsRowFormulaArray),
    }));
    this.columnNames = Array.isArray(config.columnNames)
      ? config.columnNames.map((value) => String(value))
      : this.columnDefinitions?.map((column) => column.name);
    this.filters = Array.isArray(config.filters) ? config.filters.map((filter) => {
      const columnIndex = Number(filter?.columnIndex ?? 0);
      if (filter?.kind === "custom") return {
        columnIndex,
        kind: "custom",
        matchAll: Boolean(filter.matchAll),
        criteria: Array.isArray(filter.criteria)
          ? filter.criteria.map((criterion) => ({ operator: String(criterion?.operator ?? ""), value: String(criterion?.value ?? "") }))
          : [],
      };
      if (filter?.kind === "dynamic") return {
        columnIndex,
        kind: "dynamic",
        type: String(filter.type ?? ""),
        ...(filter.value == null ? {} : { value: Number(filter.value) }),
        ...(filter.maxValue == null ? {} : { maxValue: Number(filter.maxValue) }),
      };
      if (filter?.kind === "top10") return {
        columnIndex,
        kind: "top10",
        top: filter.top ?? true,
        percent: Boolean(filter.percent),
        value: Number(filter.value ?? 0),
        ...(filter.filterValue == null ? {} : { filterValue: Number(filter.filterValue) }),
      };
      if (filter?.kind === "icon") return {
        columnIndex,
        kind: "icon",
        iconSet: String(filter.iconSet ?? ""),
        ...(filter.iconId == null ? {} : { iconId: Number(filter.iconId) }),
      };
      if (filter?.kind === "color") return {
        columnIndex,
        kind: "color",
        target: String(filter.target ?? ""),
        color: filter.color,
      };
      return {
        columnIndex,
        kind: "values",
        values: Array.isArray(filter?.values) ? filter.values.map((value) => String(value)) : [],
        includeBlank: Boolean(filter?.includeBlank),
        ...(Array.isArray(filter?.dateGroups) && filter.dateGroups.length ? {
          dateGroups: filter.dateGroups.map((group) => ({
            grouping: String(group?.grouping ?? ""),
            year: Number(group?.year ?? 0),
            ...(group?.month == null ? {} : { month: Number(group.month) }),
            ...(group?.day == null ? {} : { day: Number(group.day) }),
            ...(group?.hour == null ? {} : { hour: Number(group.hour) }),
            ...(group?.minute == null ? {} : { minute: Number(group.minute) }),
            ...(group?.second == null ? {} : { second: Number(group.second) }),
          })),
        } : {}),
        ...(filter?.calendarType ? { calendarType: String(filter.calendarType) } : {}),
      };
    }) : [];
    this.sortState = normalizeWorksheetTableSortState(config.sortState);
    this.queryTable = normalizeWorksheetTableQuery(config.queryTable);
    if (config.queryTableUnsupported === true)
      Object.defineProperty(this, WORKSHEET_TABLE_QUERY_UNSUPPORTED, { configurable: true, value: true, writable: true });
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
    return { kind: "table", id: this.id, sheet: this.worksheet.name, name: this.name, address: this.range, rows: this.rowCount, cols: this.columnCount, hasHeaders: this.hasHeaders, style: this.style, showFirstColumn: this.showFirstColumn, showLastColumn: this.showLastColumn, showRowStripes: this.showRowStripes, showBandedColumns: this.showBandedColumns, columnNames: this.columnNames, columnDefinitions: this.columnDefinitions, filters: this.filters, sortState: this.sortState, queryTable: this.queryTable, queryTableUnsupported: this[WORKSHEET_TABLE_QUERY_UNSUPPORTED] === true, values: this.values };
  }

  toSvg(bounds) {
    const tableBounds = parseRangeAddress(this.range);
    const frame = worksheetRangeFrame(this.worksheet, tableBounds, bounds);
    const { left, top, width, height } = frame;
    return `<rect x="${left}" y="${top}" width="${width}" height="${height}" fill="none" stroke="#0ea5e9" stroke-width="2"/><text x="${left}" y="${Math.max(12, top - 6)}" font-family="Arial" font-size="11" fill="#0284c7">${xmlEscape(this.name)}</text>`;
  }

  toJSON() { return { id: this.id, name: this.name, range: this.range, hasHeaders: this.hasHeaders, showHeaders: this.showHeaders, showTotals: this.showTotals, showBandedColumns: this.showBandedColumns, showFilterButton: this.showFilterButton, showFirstColumn: this.showFirstColumn, showLastColumn: this.showLastColumn, showRowStripes: this.showRowStripes, style: this.style, columnNames: this.columnNames, columnDefinitions: this.columnDefinitions, filters: this.filters, sortState: this.sortState, queryTable: this.queryTable, values: this.values }; }
}

class WorksheetTableCollection {
  constructor(worksheet) { this.worksheet = worksheet; this.items = []; }
  add(rangeOrConfig, hasHeaders = true, name) { const table = new WorksheetTable(this.worksheet, rangeOrConfig, hasHeaders, name); this.items.push(table); return table; }
  getItemOrNullObject(name) { return this.items.find((table) => table.name === name) || { isNullObject: true }; }
  deleteAll() { this.items = []; }
  inspectRecords() { return this.items.map((table) => table.inspectRecord()); }
  toJSON() { return this.items.map((table) => table.toJSON()); }
}

class WorksheetPivotTable {
  constructor(worksheet, config = {}) {
    this.worksheet = worksheet;
    this.id = config.id || aid("pvt");
    this.name = config.name || `PivotTable${worksheet.pivotTables.items.length + 1}`;
    this.sourceRange = workbookRangeRef(config.sourceRange || config.source || config.range || "A1");
    this.targetRange = workbookRangeRef(config.targetRange || config.destination || config.target || "A6");
    const matrix = this.sourceValues();
    const headers = config.sourceFields?.length ? config.sourceFields.map(String) : (matrix[0] || []).map((value) => String(value ?? ""));
    this.sourceFields = [...headers];
    const sourceValues = Object.fromEntries(headers.map((header, index) => [header, matrix.slice(1).map((row) => row[index])]));
    const normalized = normalizePivotConfig({ ...config, sourceValues, dateSystem: this.dateSystem }, headers);
    this.rowFields = normalized.rowFields;
    this.columnFields = normalized.columnFields;
    this.valueFields = normalized.valueFields;
    this.groupFields = normalized.groupFields;
    this.calculatedFields = normalized.calculatedFields;
    this.filters = normalized.filters;
    this.refreshPolicy = normalized.refreshPolicy;
  }

  get dateSystem() { return this.worksheet.workbook.dateSystem; }

  sourceValues() {
    const target = workbookRangeTarget(this.worksheet.workbook, this.worksheet, this.sourceRange);
    if (!target.sheet || !target.bounds) return [];
    return target.sheet.getRange(target.address).values;
  }

  computedValues() {
    return computePivotValues(this.sourceValues(), this);
  }

  inspectRecord() {
    const values = this.computedValues();
    return { kind: "pivotTable", id: this.id, sheet: this.worksheet.name, name: this.name, sourceRange: this.sourceRange.address, sourceSheet: this.sourceRange.sheetName || this.worksheet.name, targetRange: this.targetRange.address, rowFields: this.rowFields, columnFields: this.columnFields, valueFields: this.valueFields, groupFields: this.groupFields, calculatedFields: this.calculatedFields, filters: this.filters, refreshPolicy: this.refreshPolicy, values, rows: Math.max(0, values.length - 1), cols: values[0]?.length || 0 };
  }

  layoutJson(bounds) {
    const values = this.computedValues();
    const target = safeRangeBounds(this.targetRange.address) || { top: 0, left: 0, rowCount: Math.max(1, values.length), colCount: Math.max(1, values[0]?.length || 1) };
    const rowCount = Math.max(values.length, target.rowCount || 1);
    const colCount = Math.max(values[0]?.length || 0, target.colCount || 1);
    const frame = worksheetRangeFrame(this.worksheet, { top: target.top, left: target.left, bottom: target.top + rowCount - 1, right: target.left + colCount - 1, rowCount, colCount }, bounds);
    return { kind: "pivotTable", id: this.id, sheet: this.worksheet.name, name: this.name, sourceRange: this.sourceRange.address, targetRange: this.targetRange.address, rowFields: this.rowFields, columnFields: this.columnFields, valueFields: this.valueFields, groupFields: this.groupFields, calculatedFields: this.calculatedFields, filters: this.filters, refreshPolicy: this.refreshPolicy, values, bbox: [frame.left, frame.top, frame.width, frame.height] };
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

  toJSON() { return { id: this.id, name: this.name, sourceRange: this.sourceRange, sourceFields: this.sourceFields, targetRange: this.targetRange, rowFields: this.rowFields, columnFields: this.columnFields, valueFields: this.valueFields, groupFields: this.groupFields, calculatedFields: this.calculatedFields, filters: this.filters, refreshPolicy: this.refreshPolicy }; }
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
  toJSON() {
    return this.items.map((item) => {
      const { line: _line, stroke: _stroke, marker: _marker, ...rest } = item;
      const line = normalizeSpreadsheetChartSeriesLine(item);
      const marker = normalizeSpreadsheetChartSeriesMarker(item.marker);
      return { ...rest, ...(line == null ? {} : { line }), ...(marker == null ? {} : { marker }) };
    });
  }
}

function normalizeWorksheetChartAxis(value, kind) {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError(`Worksheet chart ${kind}Axis must be an object.`);
  const title = typeof value.title === "string" ? value.title : value.title?.text;
  return {
    axisType: value.axisType || (kind === "x" ? "textAxis" : "valueAxis"),
    title: {
      text: title == null ? "" : String(title),
      ...(typeof value.title === "object" && value.title?.textStyle != null ? { textStyle: value.title.textStyle } : {}),
    },
    ...(value.textStyle == null ? {} : { textStyle: value.textStyle }),
    ...(value.numberFormatCode == null ? {} : { numberFormatCode: String(value.numberFormatCode) }),
    ...(value.tickLabelInterval == null ? {} : { tickLabelInterval: Number(value.tickLabelInterval) }),
    ...(value.min == null ? {} : { min: Number(value.min) }),
    ...(value.max == null ? {} : { max: Number(value.max) }),
    ...(value.majorUnit == null ? {} : { majorUnit: Number(value.majorUnit) }),
  };
}

class WorksheetChart {
  constructor(worksheet, chartType = "bar", sourceOrConfig = {}) {
    this.worksheet = worksheet;
    this.id = sourceOrConfig.id || aid("wch");
    this.type = chartType;
    this.name = sourceOrConfig.name || `Chart ${worksheet.charts.items.length + 1}`;
    this.title = sourceOrConfig.title || "";
    if (sourceOrConfig.titleTextStyle != null && (typeof sourceOrConfig.titleTextStyle !== "object" || Array.isArray(sourceOrConfig.titleTextStyle))) throw new TypeError("Worksheet chart titleTextStyle must be an object.");
    this.titleTextStyle = { ...(sourceOrConfig.titleTextStyle || {}) };
    this.lineOptions = normalizeSpreadsheetChartLineOptions(sourceOrConfig.lineOptions) || undefined;
    this.dataLabels = normalizeSpreadsheetChartDataLabels(sourceOrConfig.dataLabels);
    this.hasLegend = sourceOrConfig.hasLegend ?? true;
    this.categories = sourceOrConfig.categories || [];
    this.position = sourceOrConfig.position || { left: 420, top: 40, width: 360, height: 220 };
    this._xAxis = undefined;
    this._yAxis = undefined;
    if (chartType !== "pie" || sourceOrConfig.xAxis != null) this.xAxis = sourceOrConfig.xAxis || {};
    if (chartType !== "pie" || sourceOrConfig.yAxis != null) this.yAxis = sourceOrConfig.yAxis || {};
    this.series = new WorksheetChartSeriesCollection(this);
    if (sourceOrConfig.series) sourceOrConfig.series.forEach((series) => Object.assign(this.series.add(series.name, series.values || []), {
      categoryFormula: series.categoryFormula,
      formula: series.formula,
      fill: series.fill,
      ...(series.line == null ? {} : { line: series.line }),
      ...(series.stroke == null ? {} : { stroke: series.stroke }),
      ...(series.marker == null ? {} : { marker: series.marker }),
    }));
    if (sourceOrConfig instanceof Range) this.setData(sourceOrConfig);
    else if (sourceOrConfig && sourceOrConfig.worksheet instanceof Worksheet) this.setData(sourceOrConfig);
  }

  get xAxis() { return this._xAxis; }
  set xAxis(value) { this._xAxis = normalizeWorksheetChartAxis(value, "x"); }
  get yAxis() { return this._yAxis; }
  set yAxis(value) { this._yAxis = normalizeWorksheetChartAxis(value, "y"); }

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
    const frame = worksheetRangeFrame(this.worksheet, { top: start.row, left: start.col, bottom: end.row, right: end.col, rowCount: end.row - start.row + 1, colCount: end.col - start.col + 1 });
    this.position = { left: frame.left, top: frame.top, width: Math.max(120, frame.width), height: Math.max(80, frame.height) };
    return this;
  }

  inspectRecord() {
    const categories = resolvedWorksheetChartCategories(this);
    const seriesItems = this.series.toJSON().map((series, index) => ({
      ...series,
      values: resolvedWorksheetChartSeriesValues(this, this.series.items[index]),
    }));
    return { kind: "drawing", drawingType: "chart", id: this.id, sheet: this.worksheet.name, name: this.name, chartType: this.type, title: this.title, titleTextStyle: this.titleTextStyle, lineOptions: this.lineOptions, dataLabels: normalizeSpreadsheetChartDataLabels(this.dataLabels), categories, series: this.series.items.length, seriesItems, xAxis: this.xAxis, yAxis: this.yAxis, bbox: [this.position.left, this.position.top, this.position.width, this.position.height], bboxUnit: "px" };
  }

  toSvg() {
    const p = this.position;
    const categories = resolvedWorksheetChartCategories(this);
    const seriesItems = this.series.items.map((series) => ({ ...series, values: resolvedWorksheetChartSeriesValues(this, series) }));
    const values = seriesItems[0]?.values || [];
    const lineOptions = normalizeSpreadsheetChartLineOptions(this.lineOptions);
    const dataLabels = normalizeSpreadsheetChartDataLabels(this.dataLabels);
    const grouping = lineOptions?.grouping || "standard";
    const lineTotals = values.map((_, pointIndex) => seriesItems.reduce((total, series) => total + (Number(series.values?.[pointIndex]) || 0), 0));
    const lineValues = seriesItems.map((series, seriesIndex) => (series.values || []).map((value, pointIndex) => {
      const raw = Number(value) || 0;
      if (grouping === "standard") return raw;
      const stacked = seriesItems.slice(0, seriesIndex + 1).reduce((total, item) => total + (Number(item.values?.[pointIndex]) || 0), 0);
      return grouping === "percentStacked" ? (lineTotals[pointIndex] === 0 ? 0 : stacked / lineTotals[pointIndex]) : stacked;
    }));
    const max = this.type === "line" ? Math.max(1, ...lineValues.flat()) : Math.max(1, ...values.map((value) => Number(value) || 0));
    const plot = { left: p.left + 28, top: p.top + 36, width: Math.max(0, p.width - 44), height: Math.max(0, p.height - 62) };
    const barW = values.length ? plot.width / values.length * 0.65 : 0;
    const gap = values.length ? plot.width / values.length * 0.35 : 0;
    const previewFill = /^#[0-9a-f]{6}$/i.test(seriesItems[0]?.fill || "") ? seriesItems[0].fill.toUpperCase() : "#38bdf8";
    const previewLine = normalizeSpreadsheetChartSeriesLine(seriesItems[0]);
    const previewStroke = previewLine?.fill || previewFill;
    const previewStrokeWidth = previewLine?.width ?? (this.type === "line" ? 2 : 0);
    const dashArray = spreadsheetChartLineDashArray(previewLine?.style);
    const strokeAttributes = ` stroke="${previewStroke}" stroke-width="${previewStrokeWidth}"${dashArray ? ` stroke-dasharray="${dashArray}"` : ""}`;
    const bars = values.map((value, index) => {
      const h = plot.height * (Number(value) || 0) / max;
      const x = plot.left + index * (barW + gap) + gap / 2;
      const y = plot.top + plot.height - h;
      const label = spreadsheetChartDataLabelText(dataLabels, categories[index], value, { seriesName: seriesItems[0]?.name });
      const placement = spreadsheetChartDataLabelSvgPlacement(dataLabels, { x, y, width: barW, height: h, baseY: plot.top + plot.height, plotTop: plot.top });
      return `<rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="${previewFill}"${previewLine == null ? "" : strokeAttributes}/>${label ? `<text x="${placement.x}" y="${placement.y}" text-anchor="${placement.textAnchor}" font-family="Arial" font-size="10" fill="#334155" data-chart-label-position="${placement.position}" data-chart-label-index="${index}">${xmlEscape(label)}</text>` : ""}`;
    }).join("");
    const previewPalette = ["#38BDF8", "#F97316", "#22C55E", "#A855F7", "#E11D48", "#0F766E"];
    const lineMarks = lineValues.map((seriesValues, seriesIndex) => {
      const series = seriesItems[seriesIndex];
      const fallback = previewPalette[seriesIndex % previewPalette.length];
      const fill = /^#[0-9a-f]{6}$/i.test(series?.fill || "") ? series.fill.toUpperCase() : fallback;
      const line = normalizeSpreadsheetChartSeriesLine(series);
      const stroke = line?.fill || fill;
      const width = line?.width ?? 2;
      const dash = spreadsheetChartLineDashArray(line?.style);
      const attributes = ` stroke="${stroke}" stroke-width="${width}"${dash ? ` stroke-dasharray="${dash}"` : ""}`;
      const points = seriesValues.map((value, index) => ({ x: plot.left + (index + 0.5) * plot.width / Math.max(1, seriesValues.length), y: plot.top + plot.height - plot.height * value / max }));
      const mark = lineOptions?.smooth === true
        ? `<path d="${spreadsheetChartSmoothLinePath(points)}" fill="none"${attributes} data-series-index="${seriesIndex}"/>`
        : `<polyline points="${points.map((point) => `${point.x},${point.y}`).join(" ")}" fill="none"${attributes} data-series-index="${seriesIndex}"/>`;
      const labels = points.map((point, index) => {
        const label = spreadsheetChartDataLabelText(dataLabels, categories[index], series?.values?.[index], { seriesName: series?.name });
        const placement = spreadsheetChartDataLabelSvgPlacement(dataLabels, { x: point.x, y: point.y, kind: "point", plotTop: plot.top });
        return label ? `<text x="${placement.x}" y="${placement.y}" text-anchor="${placement.textAnchor}" font-family="Arial" font-size="10" fill="#334155" data-chart-label-position="${placement.position}" data-chart-label-series="${seriesIndex}" data-chart-label-index="${index}">${xmlEscape(label)}</text>` : "";
      }).join("");
      return `${mark}${points.map((point) => spreadsheetChartMarkerSvg(series?.marker, point.x, point.y, stroke)).join("")}${labels}`;
    }).join("");
    const plotMarks = this.type === "line" ? lineMarks : bars;
    const xTickSize = Number(this.xAxis?.textStyle?.fontSize);
    const xTicks = Number.isFinite(xTickSize) && xTickSize > 0 && values.length ? categories.map((category, index) => `<text x="${plot.left + (index + 0.5) * plot.width / values.length}" y="${plot.top + plot.height + xTickSize + 2}" text-anchor="middle" font-family="Arial" font-size="${xTickSize}" fill="#64748b">${xmlEscape(category)}</text>`).join("") : "";
    const yTickSize = Number(this.yAxis?.textStyle?.fontSize);
    const yTicks = Number.isFinite(yTickSize) && yTickSize > 0 ? `<text x="${plot.left - 4}" y="${plot.top + yTickSize}" text-anchor="end" font-family="Arial" font-size="${yTickSize}" fill="#64748b">${max}</text><text x="${plot.left - 4}" y="${plot.top + plot.height}" text-anchor="end" font-family="Arial" font-size="${yTickSize}" fill="#64748b">0</text>` : "";
    const xTitle = this.xAxis?.title?.text ? `<text x="${plot.left + plot.width / 2}" y="${p.top + p.height - 6}" text-anchor="middle" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.xAxis.title.text)}</text>` : "";
    const yTitle = this.yAxis?.title?.text ? `<text x="${p.left + 10}" y="${plot.top + plot.height / 2}" text-anchor="middle" transform="rotate(-90 ${p.left + 10} ${plot.top + plot.height / 2})" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.yAxis.title.text)}</text>` : "";
    const titleSize = Number.isFinite(Number(this.titleTextStyle?.fontSize)) && Number(this.titleTextStyle.fontSize) > 0 ? Number(this.titleTextStyle.fontSize) : 13;
    return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#94a3b8"/><text x="${p.left + 8}" y="${p.top + 22}" font-family="Arial" font-size="${titleSize}" font-weight="700" fill="#0f172a">${xmlEscape(this.title || this.name)}</text>${plotMarks}${xTicks}${yTicks}${xTitle}${yTitle}`;
  }

  toJSON() { return { id: this.id, type: this.type, name: this.name, title: this.title, titleTextStyle: this.titleTextStyle, lineOptions: normalizeSpreadsheetChartLineOptions(this.lineOptions) || undefined, dataLabels: normalizeSpreadsheetChartDataLabels(this.dataLabels), hasLegend: this.hasLegend, categories: this.categories, position: this.position, series: this.series.toJSON(), xAxis: this.xAxis, yAxis: this.yAxis }; }
}

class WorksheetChartCollection {
  constructor(worksheet) { this.worksheet = worksheet; this.items = []; }
  add(chartType, sourceOrConfig = {}) { const chart = new WorksheetChart(this.worksheet, chartType, sourceOrConfig); this.items.push(chart); return chart; }
  getItemOrNullObject(name) { return this.items.find((chart) => chart.name === name) || { isNullObject: true }; }
  deleteAll() { this.items = []; }
  inspectRecords() { return this.items.map((chart) => chart.inspectRecord()); }
  toJSON() { return this.items.map((chart) => chart.toJSON()); }
}

function worksheetAnchorFrame(worksheet, anchor = {}) {
  if (anchor.type === "absolute" || anchor.position) {
    const extent = anchor.extent || {};
    return {
      left: 40 + Number(anchor.position?.leftPx ?? 0),
      top: 40 + Number(anchor.position?.topPx ?? 0),
      width: Number(extent.widthPx ?? 160),
      height: Number(extent.heightPx ?? 120),
    };
  }
  const from = anchor.from || { row: 0, col: 0 };
  const left = 40 + worksheetAxisOffset(worksheet, "column", Number(from.col || 0)) + Number(anchor.colOffsetPx || from.colOffsetPx || 0);
  const top = 40 + worksheetAxisOffset(worksheet, "row", Number(from.row || 0)) + Number(anchor.rowOffsetPx || from.rowOffsetPx || 0);
  if (anchor.to) {
    const right = 40 + worksheetAxisOffset(worksheet, "column", Number(anchor.to.col || 0)) + Number(anchor.to.colOffsetPx || 0);
    const bottom = 40 + worksheetAxisOffset(worksheet, "row", Number(anchor.to.row || 0)) + Number(anchor.to.rowOffsetPx || 0);
    return { left, top, width: Math.max(1, right - left), height: Math.max(1, bottom - top) };
  }
  const extent = anchor.extent || {};
  return {
    left,
    top,
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
    this.crop = config.crop;
    this.effects = config.effects;
    this.transform = config.transform;
  }

  get position() { return worksheetAnchorFrame(this.worksheet, this.anchor); }
  inspectRecord() { const p = this.position; return { kind: "drawing", drawingType: "image", id: this.id, sheet: this.worksheet.name, name: this.name, alt: this.alt, prompt: this.prompt, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", crop: this.crop, effects: this.effects, transform: this.transform }; }
  replace(config = {}) { Object.assign(this, config); return this; }
  toSvg() { const p = this.position; const filters = []; if (this.effects?.grayscale === true) filters.push("grayscale(1)"); const brightness = Number(this.effects?.brightnessPercent); const contrast = Number(this.effects?.contrastPercent); const opacity = Number(this.effects?.opacityPercent); if (Number.isFinite(brightness) && brightness >= -100 && brightness <= 100) filters.push(`brightness(${1 + brightness / 100})`); if (Number.isFinite(contrast) && contrast >= -100 && contrast <= 100) filters.push(`contrast(${1 + contrast / 100})`); const degrees = Number(this.transform?.rotationDegrees); const flipH = this.transform?.flipHorizontal === true ? -1 : 1; const flipV = this.transform?.flipVertical === true ? -1 : 1; const cx = p.left + p.width / 2; const cy = p.top + p.height / 2; const drawingTransform = (Number.isFinite(degrees) || flipH < 0 || flipV < 0) ? ` transform="translate(${cx} ${cy}) rotate(${Number.isFinite(degrees) ? degrees : 0}) scale(${flipH} ${flipV}) translate(${-cx} ${-cy})"` : ""; const visual = `${filters.length ? ` style="filter:${attrEscape(filters.join(" "))}"` : ""}${Number.isFinite(opacity) && opacity >= 0 && opacity <= 100 ? ` opacity="${opacity / 100}"` : ""}${drawingTransform}`; const image = this.dataUrl ? `<image href="${attrEscape(this.dataUrl)}" x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" preserveAspectRatio="xMidYMid meet"${visual}/>` : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#fef3c7" stroke="#f59e0b"${visual}/>`; return `${image}<text x="${p.left + 8}" y="${p.top + 20}" font-family="Arial" font-size="12" fill="#92400e">${xmlEscape(this.alt || this.prompt || this.name)}</text>`; }
  toJSON() { return { id: this.id, name: this.name, dataUrl: this.dataUrl, uri: this.uri, prompt: this.prompt, alt: this.alt, anchor: this.anchor, fit: this.fit, crop: this.crop, effects: this.effects, transform: this.transform }; }
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
  targetFrame(bounds) { const target = parseRangeAddress(this.targetRange.address); return worksheetRangeFrame(this.worksheet, target, bounds); }
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

class RangeConditionalFormatFacade {
  constructor(range) { this.range = range; }
  add(ruleType, config = {}) {
    const normalized = { ...config };
    if (String(ruleType).toLowerCase() === "containstext" && normalized.formula == null && normalized.formulas == null) {
      const text = String(normalized.text ?? "");
      if (!text) throw new Error("Range containsText conditional formatting requires text.");
      const address = makeCellAddress(this.range.rowIndex, this.range.columnIndex);
      normalized.formula = `NOT(ISERROR(SEARCH("${text.replaceAll('"', '""')}",${address})))`;
    }
    return this.range.worksheet.conditionalFormattings.add({ range: rangeToAddress(this.range.bounds), ruleType, ...normalized });
  }
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
    default: return matchesFormulaCriteria(actualValue, `${operator}${expectedValue}`);
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

function delimitedDelimiter(value = ",") {
  const delimiter = String(value);
  if (delimiter.length !== 1 || delimiter === '"' || delimiter === "\r" || delimiter === "\n") throw new Error("Delimited files require a single non-quote, non-newline delimiter.");
  return delimiter;
}

function delimitedLimit(value, fallback, name) {
  const limit = value == null ? fallback : Number(value);
  if (!Number.isInteger(limit) || limit < 1) throw new Error(`${name} must be a positive integer.`);
  return limit;
}

function coerceDelimitedCell(value, quoted, options = {}) {
  if (!options.coerceTypes || quoted || value === "") return value;
  if (/^(true|false)$/i.test(value)) return value.toLowerCase() === "true";
  if (/^[+-]?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(value)) {
    const number = Number(value);
    if (Number.isFinite(number)) return number;
  }
  return value;
}

function parseDelimitedText(text, options = {}) {
  const delimiter = delimitedDelimiter(options.delimiter ?? ",");
  const maxRows = delimitedLimit(options.maxRows, 100_000, "maxRows");
  const maxColumns = delimitedLimit(options.maxColumns, 16_384, "maxColumns");
  const source = String(text ?? "").replace(/^\uFEFF/, "");
  if (source === "") return { delimiter, rows: [], quotedCells: 0, formulaLikeCells: 0 };
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  let fieldQuoted = false;
  let closedQuote = false;
  let quotedCells = 0;
  let formulaLikeCells = 0;
  const finishField = () => {
    const value = coerceDelimitedCell(field, fieldQuoted, options);
    row.push(value);
    if (fieldQuoted) quotedCells += 1;
    if (typeof value === "string" && /^[=+\-@]/.test(value)) formulaLikeCells += 1;
    if (row.length > maxColumns) throw new Error(`Delimited input exceeds maxColumns (${maxColumns}).`);
    field = "";
    fieldQuoted = false;
    closedQuote = false;
  };
  const finishRow = () => {
    finishField();
    rows.push(row);
    if (rows.length > maxRows) throw new Error(`Delimited input exceeds maxRows (${maxRows}).`);
    row = [];
  };
  let endedWithRecordSeparator = false;
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (quoted) {
      if (character === '"') {
        if (source[index + 1] === '"') { field += '"'; index += 1; }
        else { quoted = false; closedQuote = true; }
      } else field += character;
      endedWithRecordSeparator = false;
      continue;
    }
    if (character === '"' && field === "" && !closedQuote) { quoted = true; fieldQuoted = true; endedWithRecordSeparator = false; continue; }
    if (character === delimiter) { finishField(); endedWithRecordSeparator = false; continue; }
    if (character === "\r" || character === "\n") {
      finishRow();
      if (character === "\r" && source[index + 1] === "\n") index += 1;
      endedWithRecordSeparator = true;
      continue;
    }
    if (closedQuote) throw new Error(`Delimited input has unexpected content after a closing quote at character ${index}.`);
    if (character === '"') throw new Error(`Delimited input has an unexpected quote in an unquoted field at character ${index}.`);
    field += character;
    endedWithRecordSeparator = false;
  }
  if (quoted) throw new Error("Delimited input contains an unterminated quoted field.");
  if (!endedWithRecordSeparator || field !== "" || row.length) finishRow();
  const columns = rows.reduce((maximum, values) => Math.max(maximum, values.length), 0);
  for (const values of rows) while (values.length < columns) values.push("");
  return { delimiter, rows, quotedCells, formulaLikeCells };
}

function delimitedCellText(value) {
  if (value == null) return "";
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "object") return JSON.stringify(value);
  return String(value);
}

function serializeDelimitedRows(rows, options = {}) {
  const delimiter = delimitedDelimiter(options.delimiter ?? ",");
  const lineEnding = options.lineEnding == null ? "\r\n" : String(options.lineEnding);
  if (lineEnding !== "\n" && lineEnding !== "\r\n") throw new Error("lineEnding must be LF or CRLF.");
  const quoteAll = options.quoteAll === true;
  const text = rows.map((row) => row.map((value) => {
    const raw = delimitedCellText(value);
    return quoteAll || raw.includes(delimiter) || /["\r\n]/.test(raw) ? `"${raw.replaceAll('"', '""')}"` : raw;
  }).join(delimiter)).join(lineEnding);
  return `${options.includeBom === true ? "\uFEFF" : ""}${text}${options.finalLineEnding === false || rows.length === 0 ? "" : lineEnding}`;
}

function delimitedInputBytes(input) {
  if (input instanceof FileBlob) return input.bytes;
  if (typeof input === "string") return encoder.encode(input);
  return toUint8Array(input);
}

function workbookDelimitedRows(workbook, options = {}) {
  if (!(workbook instanceof Workbook)) throw new Error("Delimited export requires a Workbook.");
  if (options.recalculate !== false) workbook.recalculate();
  const sheet = options.sheetName ? workbook.worksheets.getItem(options.sheetName) : workbook.worksheets.getActiveWorksheet();
  if (!sheet) throw new Error(`Delimited export could not find worksheet ${options.sheetName || "active worksheet"}.`);
  const range = sheet.getRange(options.range || rangeToAddress(sheet.usedBounds()));
  const values = range.values;
  if (options.formulas !== true) return { sheet, range, rows: values };
  const formulas = range.formulas;
  return { sheet, range, rows: values.map((row, rowIndex) => row.map((value, columnIndex) => formulas[rowIndex]?.[columnIndex] || value)) };
}

function workbookFrameIntersects(a, b) {
  const left = Math.max(a.left, b.left);
  const top = Math.max(a.top, b.top);
  const right = Math.min(a.left + a.width, b.left + b.width);
  const bottom = Math.min(a.top + a.height, b.top + b.height);
  return right > left && bottom > top;
}

function normalizeExcelDateSystem(value, fallback = "1900") {
  if (value == null || value === "") return fallback;
  if (value === true || String(value).trim() === "1904") return "1904";
  if (value === false || String(value).trim() === "1900") return "1900";
  throw new Error(`Unsupported Excel date system ${value}; expected 1900 or 1904.`);
}

function normalizeWorksheetVisibility(value, fallback = "visible") {
  if (value == null || value === "") return fallback;
  const normalized = String(value).replace(/[\s_-]/g, "").toLowerCase();
  if (normalized === "visible") return "visible";
  if (normalized === "hidden") return "hidden";
  if (normalized === "veryhidden") return "veryHidden";
  throw new Error(`Unsupported worksheet visibility ${value}; expected visible, hidden, or veryHidden.`);
}

const WORKBOOK_CALCULATION_MODES = new Map([
  ["auto", "automatic"],
  ["automatic", "automatic"],
  ["autonotable", "automaticExceptTables"],
  ["automaticexcepttables", "automaticExceptTables"],
  ["manual", "manual"],
]);
const WORKBOOK_CALCULATION_MAX_ITERATIONS = 1_000_000;
const WORKBOOK_CALCULATION_MAX_CHANGE = 1_000_000_000;

function normalizeWorkbookCalculation(value) {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw new Error("Workbook calculation settings must be an object.");
  const result = {};
  if (value.mode != null && value.mode !== "") {
    const mode = WORKBOOK_CALCULATION_MODES.get(String(value.mode).replace(/[\s_-]/g, "").toLowerCase());
    if (!mode) throw new Error(`Unsupported workbook calculation mode ${value.mode}; expected automatic, automaticExceptTables, or manual.`);
    result.mode = mode;
  }
  for (const field of ["calculateOnSave", "fullCalculationOnLoad", "forceFullCalculation", "fullPrecision"])
    if (value[field] !== undefined) result[field] = Boolean(value[field]);
  const sourceIteration = value.iteration;
  if (sourceIteration != null && (typeof sourceIteration !== "object" || Array.isArray(sourceIteration))) throw new Error("Workbook calculation iteration settings must be an object.");
  const iteration = {};
  const enabled = sourceIteration?.enabled ?? value.iterationEnabled ?? value.iterate;
  const maxIterations = sourceIteration?.maxIterations ?? value.maxIterations ?? value.iterateCount;
  const maxChange = sourceIteration?.maxChange ?? value.maxChange ?? value.iterateDelta;
  if (enabled !== undefined) iteration.enabled = Boolean(enabled);
  if (maxIterations !== undefined) {
    const number = Number(maxIterations);
    if (!Number.isInteger(number) || number < 1 || number > WORKBOOK_CALCULATION_MAX_ITERATIONS) throw new Error(`Workbook maximum calculation iterations must be an integer from 1 to ${WORKBOOK_CALCULATION_MAX_ITERATIONS}.`);
    iteration.maxIterations = number;
  }
  if (maxChange !== undefined) {
    const number = Number(maxChange);
    if (!Number.isFinite(number) || number <= 0 || number > WORKBOOK_CALCULATION_MAX_CHANGE) throw new Error(`Workbook maximum calculation change must be greater than zero and at most ${WORKBOOK_CALCULATION_MAX_CHANGE}.`);
    iteration.maxChange = number;
  }
  if (Object.keys(iteration).length) result.iteration = iteration;
  return result;
}

const WORKBOOK_CONNECTION_BOOLEAN_FIELDS = ["keepAlive", "background", "refreshOnLoad", "saveData"];

function normalizeWorkbookConnection(value = {}) {
  const connectionId = Number(value.connectionId ?? (typeof value.id === "number" ? value.id : String(value.id || "").match(/^connection\/(\d+)$/)?.[1]));
  const connection = {
    connectionId,
    name: String(value.name ?? ""),
    type: Number(value.type ?? 0),
    refreshedVersion: Number(value.refreshedVersion ?? 0),
  };
  if (value.description !== undefined) connection.description = String(value.description);
  for (const field of WORKBOOK_CONNECTION_BOOLEAN_FIELDS) if (value[field] !== undefined) connection[field] = Boolean(value[field]);
  if (value.intervalMinutes !== undefined) connection.intervalMinutes = Number(value.intervalMinutes);
  return connection;
}

export class Workbook {
  constructor(options = {}) {
    this.id = aid("wb");
    this.dateSystem = normalizeExcelDateSystem(options.dateSystem ?? options.date1904);
    this.theme = normalizeXlsxThemeConfig(options.theme || {});
    this.indexedColors = Array.isArray(options.indexedColors) ? [...options.indexedColors] : undefined;
    this.connections = Array.isArray(options.connections) ? options.connections.map(normalizeWorkbookConnection) : [];
    this.calculation = options.calculation === undefined ? undefined : normalizeWorkbookCalculation(options.calculation);
    this._activeWorksheetId = undefined;
    this._selectedWorksheetIds = undefined;
    this.worksheets = new WorksheetCollection(this);
    this.windows = createWorkbookWindowCollection(this);
    this.comments = new CommentsCollection(this);
    this.definedNames = new DefinedNameCollection(this);
  }

  static create(options = {}) {
    return new Workbook(options);
  }

  setDateSystem(value) {
    this.dateSystem = normalizeExcelDateSystem(value);
    return this;
  }

  setTheme(theme = {}) {
    this.theme = normalizeXlsxThemeConfig(theme);
    return this;
  }

  setCalculation(calculation) {
    this.calculation = calculation == null ? undefined : normalizeWorkbookCalculation(calculation);
    return this;
  }

  static async fromCSV(text, options = {}) {
    const workbook = Workbook.create();
    await workbook.fromCSV(text, options);
    return workbook;
  }

  async fromCSV(text, options = {}) {
    const sheet = this.worksheets.add(options.sheetName || `Sheet${this.worksheets.items.length + 1}`);
    const parsed = parseDelimitedText(text, { ...options, delimiter: options.delimiter ?? "," });
    if (parsed.rows.length) sheet.getRangeByIndexes(0, 0, parsed.rows.length, parsed.rows[0]?.length || 1).values = parsed.rows;
    return this;
  }

  recalculate() {
    if (this._recalculating) return this._lastFormulaGraph;
    this._recalculating = true;
    try {
      hydrateDeclaredDynamicArraySpills(this);
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
    const activeWorksheet = this.worksheets.items.length ? this.worksheets.getActiveWorksheet() : undefined;
    const selectedWorksheetIds = new Set(activeWorksheet ? this.worksheets.getSelectedWorksheets().map((sheet) => sheet.id) : []);
    const windows = activeWorksheet ? this.windows.toJSON() : [];
    if (kinds.has("workbook")) {
      records.push({ kind: "workbook", id: this.id, sheets: this.worksheets.items.length, windows, activeSheet: activeWorksheet?.name, selectedSheets: activeWorksheet ? this.worksheets.getSelectedWorksheets().map((sheet) => sheet.name) : [], dateSystem: this.dateSystem, date1904: this.dateSystem === "1904", theme: this.theme.name, calculation: this.calculation });
    }
    if (kinds.has("window") || kinds.has("workbookWindow")) records.push(...windows.map((window) => ({ kind: "workbookWindow", ...window })));
    if (kinds.has("theme")) records.push({ kind: "workbookTheme", id: `${this.id}/theme`, name: this.theme.name, colors: this.theme.colors });
    if (kinds.has("connection") || kinds.has("externalConnection")) records.push(...this.connections.map((connection) => ({ kind: "connection", id: `connection/${connection.connectionId}`, ...connection })));
    for (const sheet of this.worksheets) {
      if (kinds.has("sheet")) records.push({ kind: "sheet", id: sheet.id, name: sheet.name, visibility: sheet.visibility, active: sheet === activeWorksheet, selected: selectedWorksheetIds.has(sheet.id), rows: sheet.usedBounds().rowCount, cols: sheet.usedBounds().colCount, showGridLines: sheet.showGridLines, freezePanes: sheet.freezePanes.toJSON(), sortState: sheet.sortState, customColumns: sheet.columnDimensions.size, customRows: sheet.rowDimensions.size, mergedRanges: sheet.mergedRanges.length });
      if (kinds.has("table") || kinds.has("region")) records.push(sheet.tableRecord(options));
      if (kinds.has("table")) records.push(...sheet.tables.inspectRecords());
      if (kinds.has("pivotTable") || kinds.has("pivot")) records.push(...sheet.pivotTables.inspectRecords());
      if (kinds.has("drawing") || kinds.has("chart")) records.push(...sheet.charts.inspectRecords());
      if (kinds.has("drawing") || kinds.has("image")) records.push(...sheet.images.inspectRecords());
      if (kinds.has("sparkline") || kinds.has("drawing")) records.push(...sheet.sparklineGroups.inspectRecords());
      if (kinds.has("merge") || kinds.has("mergedCell")) records.push(...sheet.mergeRecords());
      if (kinds.has("dimension") || kinds.has("column") || kinds.has("row")) records.push(...sheet.dimensionRecords(kinds));
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
    if (this.dateSystem !== "1900" && this.dateSystem !== "1904") issues.push(verificationIssue("workbook", "invalidDateSystem", `Workbook date system ${this.dateSystem} is invalid; expected 1900 or 1904.`, { dateSystem: this.dateSystem }));
    try { if (this.calculation !== undefined) normalizeWorkbookCalculation(this.calculation); }
    catch (error) { issues.push(verificationIssue("workbook", "invalidCalculation", error.message, { calculation: this.calculation })); }
    if (this.worksheets.items.length === 0) issues.push(verificationIssue("workbook", "noSheets", "Workbook has no worksheets."));
    if (this.worksheets.items.length > 0 && !this.worksheets.items.some((sheet) => sheet.visibility === "visible")) issues.push(verificationIssue("workbook", "noVisibleSheets", "Workbook must contain at least one visible worksheet."));
    if (this.worksheets.items.length > 0) {
      for (const window of this.windows.items) {
        try {
          const active = window.getActiveWorksheet();
          const selected = window.getSelectedWorksheets();
          if (!selected.includes(active)) throw new Error("selected worksheets do not include the active worksheet");
        } catch (error) {
          issues.push(verificationIssue("workbook", "invalidWorkbookWindow", `Workbook window ${window.index} is invalid: ${error.message}`, { windowIndex: window.index }));
        }
      }
    }
    const connectionIds = new Set();
    for (const connection of this.connections) {
      if (!Number.isInteger(connection.connectionId) || connection.connectionId <= 0 || connectionIds.has(connection.connectionId) ||
          !String(connection.name || "").trim() || connection.type !== 5 ||
          !Number.isInteger(connection.refreshedVersion) || connection.refreshedVersion < 0 || connection.refreshedVersion > 255 ||
          connection.intervalMinutes != null && (!Number.isInteger(connection.intervalMinutes) || connection.intervalMinutes < 0 || connection.intervalMinutes > 32_767)) {
        issues.push(verificationIssue("workbook", "invalidConnection", `Workbook connection ${connection.connectionId} has invalid bounded metadata.`, { connectionId: connection.connectionId, name: connection.name }));
      }
      connectionIds.add(connection.connectionId);
    }
    for (const definedName of this.definedNames.items) {
      const refersTo = String(definedName.refersTo || "").replace(/^=/, "");
      if (!formulaRefParts(refersTo)) issues.push(verificationIssue("workbook", "invalidDefinedName", `Defined name ${definedName.name} does not reference a valid A1 range.`, { id: definedName.id, name: definedName.name, refersTo: definedName.refersTo }));
    }
    for (const sheet of this.worksheets) {
      const entries = sheet.store.entries();
      try { normalizeWorksheetVisibility(sheet.visibility); }
      catch (error) { issues.push(verificationIssue("workbook", "invalidWorksheetVisibility", error.message, { sheet: sheet.name, visibility: sheet.visibility })); }
      const frozenRows = Number(sheet.freezePanes?.rows ?? 0);
      const frozenColumns = Number(sheet.freezePanes?.columns ?? 0);
      if (!Number.isInteger(frozenRows) || frozenRows < 0 || frozenRows > XLSX_MAX_FREEZE_ROWS) issues.push(verificationIssue("workbook", "invalidFrozenRows", `Worksheet ${sheet.name} has invalid frozen row count ${sheet.freezePanes?.rows}.`, { sheet: sheet.name, rows: sheet.freezePanes?.rows }));
      if (!Number.isInteger(frozenColumns) || frozenColumns < 0 || frozenColumns > XLSX_MAX_FREEZE_COLUMNS) issues.push(verificationIssue("workbook", "invalidFrozenColumns", `Worksheet ${sheet.name} has invalid frozen column count ${sheet.freezePanes?.columns}.`, { sheet: sheet.name, columns: sheet.freezePanes?.columns }));
      for (const [column, dimension] of sheet.columnDimensions) {
        if (!Number.isInteger(column) || column < 0 || column > XLSX_MAX_FREEZE_COLUMNS) issues.push(verificationIssue("workbook", "invalidColumnDimensionIndex", `Worksheet ${sheet.name} has an invalid column dimension index ${column}.`, { sheet: sheet.name, column }));
        if (dimension.width != null && (!Number.isFinite(dimension.width) || dimension.width <= 0 || dimension.width > XLSX_MAX_COLUMN_WIDTH)) issues.push(verificationIssue("workbook", "invalidColumnWidth", `Worksheet ${sheet.name} column ${column + 1} has invalid width ${dimension.width}.`, { sheet: sheet.name, column, width: dimension.width }));
      }
      for (const [row, dimension] of sheet.rowDimensions) {
        if (!Number.isInteger(row) || row < 0 || row > XLSX_MAX_FREEZE_ROWS) issues.push(verificationIssue("workbook", "invalidRowDimensionIndex", `Worksheet ${sheet.name} has an invalid row dimension index ${row}.`, { sheet: sheet.name, row }));
        if (dimension.height != null && (!Number.isFinite(dimension.height) || dimension.height <= 0 || dimension.height > XLSX_MAX_ROW_HEIGHT)) issues.push(verificationIssue("workbook", "invalidRowHeight", `Worksheet ${sheet.name} row ${row + 1} has invalid height ${dimension.height}.`, { sheet: sheet.name, row, height: dimension.height }));
      }
      const validMerges = [];
      for (const range of sheet.mergedRanges || []) {
        let bounds;
        try { bounds = worksheetRangeBounds(range); } catch {
          issues.push(verificationIssue("workbook", "invalidMergedRange", `Worksheet ${sheet.name} has invalid merged range ${range}.`, { sheet: sheet.name, range }));
          continue;
        }
        if (bounds.rowCount * bounds.colCount <= 1) issues.push(verificationIssue("workbook", "singleCellMerge", `Worksheet ${sheet.name} merge ${range} does not span multiple cells.`, { sheet: sheet.name, range }));
        for (const existing of validMerges) if (worksheetBoundsIntersect(existing.bounds, bounds)) issues.push(verificationIssue("workbook", "overlappingMergedRanges", `Worksheet ${sheet.name} merge ${range} overlaps ${existing.range}.`, { sheet: sheet.name, range, existingRange: existing.range }));
        for (let row = bounds.top; row <= bounds.bottom; row += 1) for (let column = bounds.left; column <= bounds.right; column += 1) {
          if (row === bounds.top && column === bounds.left) continue;
          const cell = sheet.store.cells.get(makeCellAddress(row, column));
          if (cell && (cell.value != null || cell.formula)) issues.push(verificationIssue("workbook", "mergedSubordinateContent", `Worksheet ${sheet.name} merged range ${range} has content outside its upper-left cell.`, { sheet: sheet.name, range, address: makeCellAddress(row, column) }));
        }
        validMerges.push({ range, bounds });
      }
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
        const groupNames = new Set(pivot.groupFields.map((field) => field.name));
        for (const groupField of pivot.groupFields) if (!headers.includes(groupField.sourceField)) issues.push(verificationIssue("workbook", "pivotGroupSourceMissing", `Pivot table ${pivot.name || pivot.id} group field ${groupField.name} references missing source field ${groupField.sourceField}.`, { sheet: sheet.name, id: pivot.id, groupField: groupField.name, sourceField: groupField.sourceField }));
        for (const groupField of pivot.groupFields) if (groupField.supported === false) issues.push(verificationIssue("workbook", "pivotGroupFieldUnsupported", `Pivot table ${pivot.name || pivot.id} group field ${groupField.name} uses an unsupported grouping type.`, { sheet: sheet.name, id: pivot.id, groupField: groupField.name, groupBy: groupField.groupBy, error: groupField.error }));
        for (const field of [...pivot.rowFields, ...pivot.columnFields]) {
          if (field && !headers.includes(String(field)) && !groupNames.has(String(field))) issues.push(verificationIssue("workbook", "pivotFieldMissing", `Pivot table ${pivot.name || pivot.id} references missing field ${field}.`, { sheet: sheet.name, id: pivot.id, field }));
        }
        const valueFields = new Set([...headers, ...groupNames, ...pivot.calculatedFields.map((field) => field.name)]);
        for (const field of pivot.valueFields.map((item) => item.field || item.name)) if (field && !valueFields.has(String(field))) issues.push(verificationIssue("workbook", "pivotFieldMissing", `Pivot table ${pivot.name || pivot.id} references missing value field ${field}.`, { sheet: sheet.name, id: pivot.id, field }));
        for (const calculatedField of pivot.calculatedFields) for (const field of calculatedField.references) if (!headers.includes(String(field))) issues.push(verificationIssue("workbook", "pivotCalculatedFieldMissing", `Pivot table ${pivot.name || pivot.id} calculated field ${calculatedField.name} references missing source field ${field}.`, { sheet: sheet.name, id: pivot.id, calculatedField: calculatedField.name, field }));
        for (const calculatedField of pivot.calculatedFields) if (calculatedField.supported === false) issues.push(verificationIssue("workbook", "pivotCalculatedFieldUnsupported", `Pivot table ${pivot.name || pivot.id} calculated field ${calculatedField.name} uses a formula outside the supported bounded arithmetic/function subset.`, { sheet: sheet.name, id: pivot.id, calculatedField: calculatedField.name, formula: calculatedField.formula, error: calculatedField.error }));
      }
      for (const chart of sheet.charts.items) {
        if (!chart.title) issues.push(verificationIssue("workbook", "untitledChart", `Chart ${chart.name} on ${sheet.name} has no title.`, { severity: "warning", sheet: sheet.name, id: chart.id }));
        for (const [scope, style] of [["titleTextStyle", chart.titleTextStyle], ["xAxis.textStyle", chart.xAxis?.textStyle], ["yAxis.textStyle", chart.yAxis?.textStyle]]) {
          if (style == null) continue;
          const fontSize = Number(style?.fontSize);
          const unsupported = typeof style === "object" && !Array.isArray(style) ? Object.keys(style).filter((key) => key !== "fontSize" && style[key] != null) : ["non-object"];
          if (unsupported.length || style.fontSize != null && (!Number.isFinite(fontSize) || fontSize < 1 || fontSize > 4000)) issues.push(verificationIssue("workbook", "invalidChartTextStyle", `Chart ${chart.name} ${scope} must contain only a fontSize from 1 through 4000 points.`, { sheet: sheet.name, id: chart.id, scope, fontSize: style?.fontSize, unsupported }));
        }
        if (chart.titleTextStyle?.fontSize != null && !chart.title) issues.push(verificationIssue("workbook", "invalidChartTextStyle", `Chart ${chart.name} titleTextStyle requires a non-empty title.`, { sheet: sheet.name, id: chart.id, scope: "titleTextStyle" }));
        try {
          const lineOptions = normalizeSpreadsheetChartLineOptions(chart.lineOptions);
          if (lineOptions != null && chart.type !== "line") issues.push(verificationIssue("workbook", "invalidChartLineOptions", `Chart ${chart.name} lineOptions require a line chart.`, { sheet: sheet.name, id: chart.id, lineOptions }));
        } catch (error) { issues.push(verificationIssue("workbook", "invalidChartLineOptions", String(error?.message || error), { sheet: sheet.name, id: chart.id, lineOptions: chart.lineOptions })); }
        try { normalizeSpreadsheetChartDataLabels(chart.dataLabels); }
        catch (error) { issues.push(verificationIssue("workbook", "invalidChartDataLabels", String(error?.message || error), { sheet: sheet.name, id: chart.id, dataLabels: chart.dataLabels })); }
        if (chart.xAxis?.title?.textStyle != null || chart.yAxis?.title?.textStyle != null) issues.push(verificationIssue("workbook", "unsupportedChartTextStyle", `Chart ${chart.name} axis-title text styling is outside the bounded chart profile.`, { sheet: sheet.name, id: chart.id }));
        if (chart.series.items.length === 0) issues.push(verificationIssue("workbook", "emptyChart", `Chart ${chart.name} on ${sheet.name} has no data series.`, { sheet: sheet.name, id: chart.id }));
        for (const series of chart.series.items) {
          if (chart.categories.length && series.values.length && chart.categories.length !== series.values.length) issues.push(verificationIssue("workbook", "chartDataMismatch", `Chart ${chart.name} series ${series.name || "Series"} has ${series.values.length} values for ${chart.categories.length} categories.`, { sheet: sheet.name, id: chart.id, series: series.name, values: series.values.length, categories: chart.categories.length }));
          if (series.formula && !workbookRangeValid(this, sheet, series.formula)) issues.push(verificationIssue("workbook", "chartFormulaInvalid", `Chart ${chart.name} series ${series.name || "Series"} references an invalid range.`, { sheet: sheet.name, id: chart.id, formula: series.formula }));
          if (series.categoryFormula && !workbookRangeValid(this, sheet, series.categoryFormula)) issues.push(verificationIssue("workbook", "chartCategoryFormulaInvalid", `Chart ${chart.name} categories reference an invalid range.`, { sheet: sheet.name, id: chart.id, formula: series.categoryFormula }));
          if (series.fill != null && (typeof series.fill !== "string" || !/^#[0-9a-f]{6}$/i.test(series.fill))) issues.push(verificationIssue("workbook", "invalidChartSeriesFill", `Chart ${chart.name} series ${series.name || "Series"} fill must be a #RRGGBB solid color.`, { sheet: sheet.name, id: chart.id, series: series.name, fill: series.fill }));
          try { normalizeSpreadsheetChartSeriesLine(series); } catch (error) { issues.push(verificationIssue("workbook", "invalidChartSeriesLine", String(error?.message || error), { sheet: sheet.name, id: chart.id, series: series.name, line: series.line, stroke: series.stroke })); }
          try {
            const marker = normalizeSpreadsheetChartSeriesMarker(series.marker);
            if (marker != null && chart.type !== "line") issues.push(verificationIssue("workbook", "invalidChartSeriesMarker", `Chart ${chart.name} series markers require a line chart.`, { sheet: sheet.name, id: chart.id, series: series.name, marker }));
          } catch (error) { issues.push(verificationIssue("workbook", "invalidChartSeriesMarker", String(error?.message || error), { sheet: sheet.name, id: chart.id, series: series.name, marker: series.marker })); }
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
    try { planSpreadsheetThreadedComments(collectWorkbookThreadParts(this)); }
    catch (error) { issues.push(verificationIssue("workbook", "threadedCommentMetadataInvalid", error.message, { error: error.name })); }
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
      for (const ref of formulaReferences(cell.formula, sheet, address)) {
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
    const window = this.windows.items.find((item) => item.id === id);
    if (window) return window;
    const connection = this.connections.find((item) => id === item.connectionId || id === `connection/${item.connectionId}` || id === item.name);
    if (connection) return connection;
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
      activeSheet: this.worksheets.items.length ? this.worksheets.getActiveWorksheet().name : undefined,
      selectedSheets: this.worksheets.items.length ? this.worksheets.getSelectedWorksheets().map((sheet) => sheet.name) : [],
      windows: this.worksheets.items.length ? this.windows.toJSON() : [],
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
  for (const collection of ["cells", "merges", "tables", "pivots", "charts", "images", "sparklines", "rules"]) {
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
  for (const collection of ["cells", "merges", "tables", "pivots", "charts", "images", "sparklines", "rules"]) {
    const kept = keepByCollection.get(collection) || new Set();
    sliced[collection] = layout[collection].filter((_, index) => kept.has(index));
  }
  sliced.slice = { targets, search: search || undefined, before, after, matchedRecords: matches.length, returnedRecords: keep.size };
  return sliced;
}

const XLSX_MAX_FREEZE_ROWS = XLSX_MAX_ROW_INDEX;
const XLSX_MAX_FREEZE_COLUMNS = XLSX_MAX_COLUMN_INDEX;

function normalizeFreezeCount(value, maximum, axis) {
  const count = Number(value);
  if (!Number.isInteger(count) || count < 0 || count > maximum) throw new Error(`Worksheet frozen ${axis} must be an integer from 0 through ${maximum}.`);
  return count;
}

class WorksheetFreezePanes {
  constructor(worksheet) {
    this.worksheet = worksheet;
    this._rows = 0;
    this._columns = 0;
  }

  get rows() { return this._rows; }
  get columns() { return this._columns; }
  get frozen() { return this._rows > 0 || this._columns > 0; }
  get topLeftCell() { return this.frozen ? makeCellAddress(this._rows, this._columns) : undefined; }
  get activePane() {
    if (this._rows > 0 && this._columns > 0) return "bottomRight";
    if (this._rows > 0) return "bottomLeft";
    if (this._columns > 0) return "topRight";
    return undefined;
  }

  freezeRows(rowCount = 1) {
    this._rows = normalizeFreezeCount(rowCount, XLSX_MAX_FREEZE_ROWS, "row count");
    return this;
  }

  freezeColumns(columnCount = 1) {
    this._columns = normalizeFreezeCount(columnCount, XLSX_MAX_FREEZE_COLUMNS, "column count");
    return this;
  }

  unfreeze() {
    this._rows = 0;
    this._columns = 0;
    return this;
  }

  toJSON() {
    return { rows: this.rows, columns: this.columns, frozen: this.frozen, topLeftCell: this.topLeftCell, activePane: this.activePane };
  }
}

function worksheetRangeBounds(value) {
  const bounds = value instanceof Range ? value.bounds : parseRangeAddress(String(value));
  if (bounds.top < 0 || bounds.left < 0 || bounds.bottom > XLSX_MAX_FREEZE_ROWS || bounds.right > XLSX_MAX_FREEZE_COLUMNS) throw new Error(`Worksheet range ${rangeToAddress(bounds)} exceeds XLSX row/column limits.`);
  return bounds;
}

function worksheetBoundsIntersect(left, right) {
  return left.left <= right.right && left.right >= right.left && left.top <= right.bottom && left.bottom >= right.top;
}

function worksheetMergedCellInfo(sheet, row, column) {
  for (const range of sheet.mergedRanges || []) {
    let bounds;
    try { bounds = parseRangeAddress(range); } catch { continue; }
    if (row < bounds.top || row > bounds.bottom || column < bounds.left || column > bounds.right) continue;
    return { range, bounds, topLeft: row === bounds.top && column === bounds.left, parent: makeCellAddress(bounds.top, bounds.left) };
  }
  return undefined;
}

function worksheetMergeTargets(bounds, across = false) {
  if (across) return Array.from({ length: bounds.rowCount }, (_, index) => ({ top: bounds.top + index, bottom: bounds.top + index, left: bounds.left, right: bounds.right, rowCount: 1, colCount: bounds.colCount })).filter((item) => item.colCount > 1);
  return bounds.rowCount * bounds.colCount > 1 ? [bounds] : [];
}

function clearMergedSubordinateContents(sheet, bounds) {
  for (let row = bounds.top; row <= bounds.bottom; row += 1) {
    for (let column = bounds.left; column <= bounds.right; column += 1) {
      if (row === bounds.top && column === bounds.left) continue;
      const address = makeCellAddress(row, column);
      const cell = sheet.store.cells.get(address);
      if (!cell) continue;
      cell.value = null;
      cell.formula = null;
      for (const key of ["formulaType", "sharedIndex", "sharedRef", "arrayRef", "dynamicArrayRef", "spillParent", "spillAnchor", "spillRange", "spillValues", "spillError"]) delete cell[key];
    }
  }
}

function cloneCellForFill(source, sourceAddress, targetAddress) {
  const cell = structuredClone(source || { value: null, formula: null, style: {} });
  if (cell.formula) cell.formula = translateA1Formula(cell.formula, sourceAddress, targetAddress);
  for (const key of ["formulaType", "sharedIndex", "sharedRef", "arrayRef", "dynamicArrayRef", "spillParent", "spillAnchor", "spillRange", "spillValues", "spillError"]) delete cell[key];
  return cell;
}

function detachNativeFormulaTopology(sheet, address) {
  let target;
  try { target = parseCellAddress(address); } catch { return; }
  const current = sheet.store.cells.get(address);
  const sharedIndex = current?.formulaType === "shared" ? current.sharedIndex : undefined;
  const sharedRef = current?.formulaType === "shared" ? current.sharedRef : undefined;
  for (const [candidateAddress, candidate] of sheet.store.entries()) {
    let detachShared = false;
    if (sharedIndex != null && candidate.formulaType === "shared" && candidate.sharedIndex === sharedIndex) {
      detachShared = !sharedRef || !candidate.sharedRef || String(candidate.sharedRef).toUpperCase() === String(sharedRef).toUpperCase();
    }
    let detachArray = false;
    let arrayBounds;
    const arrayReference = candidate.formulaType === "array"
      ? candidate.arrayRef
      : candidate.formulaType === "dynamicArray" ? candidate.dynamicArrayRef : undefined;
    if (arrayReference) {
      try {
        arrayBounds = parseRangeAddress(arrayReference);
        detachArray = target.row >= arrayBounds.top && target.row <= arrayBounds.bottom && target.col >= arrayBounds.left && target.col <= arrayBounds.right;
      } catch {
        detachArray = candidateAddress === address;
      }
    }
    if (detachShared) {
      delete candidate.formulaType;
      delete candidate.sharedIndex;
      delete candidate.sharedRef;
    }
    if (detachArray) {
      if (candidate.formulaType === "dynamicArray" && arrayBounds) markDeclaredDynamicArrayChildren(sheet, candidateAddress, arrayReference, arrayBounds);
      delete candidate.formulaType;
      delete candidate.arrayRef;
      delete candidate.dynamicArrayRef;
    }
  }
}

export class Worksheet {
  constructor(workbook, name, options = {}) {
    this.workbook = workbook;
    this.id = aid("ws");
    this.name = name;
    this._visibility = normalizeWorksheetVisibility(options.visibility);
    this.store = new CellStore();
    this.columnDimensions = new Map();
    this.rowDimensions = new Map();
    this.mergedRanges = [];
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
    this.freezePanes = new WorksheetFreezePanes(this);
    this.showGridLines = true;
    this.sortState = undefined;
  }

  get visibility() { return this._visibility; }
  set visibility(value) {
    const visibility = normalizeWorksheetVisibility(value);
    if (visibility !== "visible" && this.workbook?.windows) {
      const memberships = worksheetWindowMemberships(this.workbook, this.id);
      const active = memberships.find((item) => item.active);
      if (active) throw new Error(`Workbook window ${active.windowIndex} active worksheet ${this.name} must remain visible; select another active worksheet first.`);
      const selected = memberships.find((item) => item.selected);
      if (selected) throw new Error(`Workbook window ${selected.windowIndex} selected worksheet ${this.name} must remain visible; change that selected worksheet group first.`);
    }
    this._visibility = visibility;
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

  mergeCells(address, across = false) {
    const bounds = worksheetRangeBounds(address);
    const targets = worksheetMergeTargets(bounds, Boolean(across));
    const existing = this.mergedRanges.map((range) => ({ range, bounds: parseRangeAddress(range) }));
    for (const target of targets) {
      const canonical = rangeToAddress(target);
      for (const item of existing) {
        if (item.range === canonical) continue;
        if (worksheetBoundsIntersect(item.bounds, target)) throw new Error(`Worksheet merge ${canonical} overlaps existing merged range ${item.range}.`);
      }
    }
    for (const target of targets) {
      const canonical = rangeToAddress(target);
      if (!this.mergedRanges.includes(canonical)) this.mergedRanges.push(canonical);
      clearMergedSubordinateContents(this, target);
    }
    this.mergedRanges.sort((left, right) => {
      const a = parseRangeAddress(left), b = parseRangeAddress(right);
      return a.top - b.top || a.left - b.left || a.bottom - b.bottom || a.right - b.right;
    });
    this.recalculate();
    return this;
  }

  unmergeCells(address) {
    const bounds = worksheetRangeBounds(address);
    this.mergedRanges = this.mergedRanges.filter((range) => !worksheetBoundsIntersect(parseRangeAddress(range), bounds));
    return this;
  }

  getUsedRange(valuesOnly = false) {
    return new Range(this, this.usedBounds(Boolean(valuesOnly)));
  }

  deleteAllDrawings() {
    this.charts.deleteAll();
    this.images.deleteAll();
    this.shapes = [];
  }

  usedBounds(valuesOnly = false) {
    const coords = this.store.entries()
      .filter(([, cell]) => !valuesOnly || cell.value != null || Boolean(cell.formula))
      .map(([address]) => parseCellAddress(address));
    if (!valuesOnly) {
      for (const range of this.mergedRanges) {
        const bounds = parseRangeAddress(range);
        coords.push({ row: bounds.top, col: bounds.left }, { row: bounds.bottom, col: bounds.right });
      }
    }
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
        dynamicArrayRef: cell.dynamicArrayRef,
        spillRange: cell.spillRange,
        spillValues: cell.spillValues,
        spillError: cell.spillError,
        precedents: graphNode?.precedents?.map((ref) => ref.key) || formulaReferences(cell.formula, this, address).map((ref) => formulaCellKey(ref.sheetName || this.name, ref.address)),
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

  dimensionRecords(kinds = new Set(["dimension"])) {
    const includeColumns = kinds.has("dimension") || kinds.has("column");
    const includeRows = kinds.has("dimension") || kinds.has("row");
    return [
      ...(includeColumns ? [...this.columnDimensions.entries()].sort((a, b) => a[0] - b[0]).map(([column, dimension]) => ({ kind: "column", sheet: this.name, index: column, column: columnNumberToLabel(column), width: dimension.width == null ? undefined : xlsxColumnWidthToCharacters(dimension.width), widthPx: worksheetColumnWidthPx(this, column, undefined, false), visibleWidthPx: worksheetColumnWidthPx(this, column), hidden: Boolean(dimension.hidden), bestFit: Boolean(dimension.bestFit) })) : []),
      ...(includeRows ? [...this.rowDimensions.entries()].sort((a, b) => a[0] - b[0]).map(([row, dimension]) => ({ kind: "row", sheet: this.name, index: row, row: row + 1, height: dimension.height, heightPx: worksheetRowHeightPx(this, row, undefined, false), visibleHeightPx: worksheetRowHeightPx(this, row), hidden: Boolean(dimension.hidden) })) : []),
    ];
  }

  mergeRecords() {
    return this.mergedRanges.map((range) => {
      const bounds = parseRangeAddress(range);
      return { kind: "mergedCell", sheet: this.name, range, address: range, topLeftCell: makeCellAddress(bounds.top, bounds.left), rows: bounds.rowCount, cols: bounds.colCount };
    });
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
    const activeWorksheet = this.workbook.worksheets.getActiveWorksheet();
    const selectedWorksheetIds = new Set(this.workbook.worksheets.getSelectedWorksheets().map((sheet) => sheet.id));
    const bounds = options.range ? parseRangeAddress(options.range) : this.usedBounds();
    const cellW = Number(options.cellWidthPx || options.cellW || xlsxColumnWidthToPixels(xlsxColumnCharactersToWidth(XLSX_DEFAULT_COLUMN_WIDTH)));
    const cellH = Number(options.cellHeightPx || options.cellH || XLSX_DEFAULT_ROW_HEIGHT * 96 / 72);
    const frameForBounds = (rangeBounds) => worksheetRangeFrame(this, rangeBounds, bounds, { cellWidthPx: cellW, cellHeightPx: cellH });
    const frameForRange = (range) => {
      try { return frameForBounds(parseRangeAddress(range)); } catch { return undefined; }
    };
    const cells = [];
    for (let r = bounds.top; r <= bounds.bottom; r += 1) {
      for (let c = bounds.left; c <= bounds.right; c += 1) {
        const address = makeCellAddress(r, c);
        const cell = this.store.get(address);
        const computed = worksheetComputedCellStyle(this, address, cell.style);
        const merged = worksheetMergedCellInfo(this, r, c);
        const frame = merged?.topLeft ? frameForBounds(merged.bounds) : frameForBounds({ top: r, bottom: r, left: c, right: c, rowCount: 1, colCount: 1 });
        cells.push({
          kind: "cell",
          sheet: this.name,
          address,
          row: r,
          col: c,
          bbox: merged && !merged.topLeft ? [frame.left, frame.top, 0, 0] : [frame.left, frame.top, frame.width, frame.height],
          hidden: Boolean(merged && !merged.topLeft || worksheetColumnDimension(this, c).hidden || worksheetRowDimension(this, r).hidden),
          mergedRange: merged?.range,
          mergedParent: merged && !merged.topLeft ? merged.parent : undefined,
          rowSpan: merged?.topLeft ? merged.bounds.rowCount : undefined,
          colSpan: merged?.topLeft ? merged.bounds.colCount : undefined,
          value: cell.value,
          displayValue: formatSpreadsheetDisplayValue(cell.value, computed.style, { dateSystem: this.workbook.dateSystem }),
          formula: cell.formula || undefined,
          spillParent: cell.spillParent,
          spillRange: cell.spillRange,
          style: cell.style && Object.keys(cell.style).length ? { ...cell.style } : undefined,
          computedStyle: computed.matches.length || Object.keys(computed.style || {}).length ? computed.style : undefined,
          conditionalFormats: computed.matches.length ? computed.matches.map((match) => ({ id: match.id, ruleType: match.ruleType, operator: match.operator, formula: match.formula, format: match.format })) : undefined,
        });
      }
    }
    const merges = this.mergedRanges.map((range) => { const mergeBounds = parseRangeAddress(range); const frame = frameForBounds(mergeBounds); return { kind: "mergedCell", sheet: this.name, range, topLeftCell: makeCellAddress(mergeBounds.top, mergeBounds.left), bbox: [frame.left, frame.top, frame.width, frame.height], rows: mergeBounds.rowCount, cols: mergeBounds.colCount }; });
    const tables = this.tables.items.map((table) => ({ kind: "table", id: table.id, sheet: this.name, name: table.name, address: table.range, bbox: Object.values(frameForRange(table.range) || {}), rows: table.rowCount, cols: table.columnCount, hasHeaders: table.hasHeaders }));
    const pivots = this.pivotTables.items.map((pivot) => pivot.layoutJson(bounds));
    const charts = this.charts.items.map((chart) => ({ kind: "chart", id: chart.id, sheet: this.name, name: chart.name, chartType: chart.type, title: chart.title, titleTextStyle: chart.titleTextStyle, bbox: [chart.position.left, chart.position.top, chart.position.width, chart.position.height], series: chart.series.items.length, seriesItems: chart.series.toJSON(), categories: chart.categories, xAxis: chart.xAxis, yAxis: chart.yAxis }));
    const images = this.images.items.map((image) => { const p = image.position; return { kind: "image", id: image.id, sheet: this.name, name: image.name, alt: image.alt, bbox: [p.left, p.top, p.width, p.height], fit: image.fit, hasDataUrl: Boolean(image.dataUrl), uri: image.uri, prompt: image.prompt }; });
    const sparklines = this.sparklineGroups.items.map((group) => { const p = group.targetFrame(bounds); return { kind: "sparkline", id: group.id, sheet: this.name, type: group.type, targetRange: group.targetRange.address, sourceData: group.sourceData.address, bbox: [p.left, p.top, p.width, p.height], values: group.values() }; });
    const rules = [...this.dataValidations.items, ...this.conditionalFormattings.items].map((rule) => ({ kind: rule.kind, id: rule.id, sheet: this.name, range: rule.range, bbox: Object.values(frameForRange(rule.range) || {}), ruleType: rule.ruleType, rule: rule.rule }));
    const drawingFrames = [...merges.map((item) => item.bbox), ...charts.map((item) => item.bbox), ...images.map((item) => item.bbox), ...sparklines.map((item) => item.bbox), ...tables.map((item) => item.bbox), ...pivots.map((item) => item.bbox)].filter((bbox) => bbox.length === 4);
    const usedFrame = frameForBounds(bounds);
    const width = Math.max(320, usedFrame.width + 80, ...drawingFrames.map((bbox) => bbox[0] + bbox[2] + 40));
    const height = Math.max(180, usedFrame.height + 80, ...drawingFrames.map((bbox) => bbox[1] + bbox[3] + 40));
    const layout = {
      kind: "worksheetLayout",
      id: this.id,
      name: this.name,
      sheet: this.name,
      bounds: { ...bounds, address: rangeToAddress(bounds) },
      unit: "px",
      origin: { left: 40, top: 40 },
      cell: { width: cellW, height: cellH },
      dimensions: this.dimensionRecords(new Set(["dimension"])),
      view: { visibility: this.visibility, active: this === activeWorksheet, selected: selectedWorksheetIds.has(this.id), showGridLines: this.showGridLines, freezePanes: this.freezePanes.toJSON() },
      width,
      height,
      cells,
      merges,
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
    const cellW = xlsxColumnWidthToPixels(xlsxColumnCharactersToWidth(XLSX_DEFAULT_COLUMN_WIDTH));
    const cellH = XLSX_DEFAULT_ROW_HEIGHT * 96 / 72;
    const imageFrames = this.images.items.map((image) => image.position);
    const sparklineFrames = this.sparklineGroups.items.map((group) => group.targetFrame(bounds));
    const pivotFrames = this.pivotTables.items.map((pivot) => { const bbox = pivot.layoutJson(bounds).bbox; return { left: bbox[0], top: bbox[1], width: bbox[2], height: bbox[3] }; });
    const usedFrame = worksheetRangeFrame(this, bounds, bounds, { cellWidthPx: cellW, cellHeightPx: cellH });
    const width = Math.max(320, usedFrame.width + 80, ...this.charts.items.map((chart) => chart.position.left + chart.position.width + 40), ...imageFrames.map((frame) => frame.left + frame.width + 40), ...sparklineFrames.map((frame) => frame.left + frame.width + 40), ...pivotFrames.map((frame) => frame.left + frame.width + 40));
    const height = Math.max(180, usedFrame.height + 80, ...this.charts.items.map((chart) => chart.position.top + chart.position.height + 40), ...imageFrames.map((frame) => frame.top + frame.height + 40), ...sparklineFrames.map((frame) => frame.top + frame.height + 40), ...pivotFrames.map((frame) => frame.top + frame.height + 40));
    const rows = [];
    const fillDefinitions = [];
    for (let r = bounds.top; r <= bounds.bottom; r++) {
      for (let c = bounds.left; c <= bounds.right; c++) {
        const address = makeCellAddress(r, c);
        const cell = this.store.get(address);
        const computed = worksheetComputedCellStyle(this, address, cell.style);
        const value = formatSpreadsheetDisplayValue(cell.value, computed.style, { dateSystem: this.workbook.dateSystem });
        const merged = worksheetMergedCellInfo(this, r, c);
        if (merged && !merged.topLeft) continue;
        const frame = worksheetRangeFrame(this, merged?.bounds || { top: r, bottom: r, left: c, right: c, rowCount: 1, colCount: 1 }, bounds, { cellWidthPx: cellW, cellHeightPx: cellH });
        const x = frame.left;
        const y = frame.top;
        if (frame.width <= 0 || frame.height <= 0) continue;
        const colorResources = { theme: this.workbook.theme, indexedColors: this.workbook.indexedColors, background: "#FFFFFF" };
        const fill = xlsxFillSvgPaint(computed.style?.fill || "white", `fill-${r}-${c}`, colorResources);
        if (fill.definition) fillDefinitions.push(fill.definition);
        const font = computed.style?.font || {};
        const fontWeight = font.bold || computed.style?.bold ? "700" : "400";
        const fontFill = xlsxColorCss(font.color || computed.style?.color || "#24292f", { ...colorResources, fallback: "#24292f" });
        const alignment = computed.style?.alignment || {};
        const textX = alignment.horizontal === "center" ? x + frame.width / 2 : ["right", "end"].includes(alignment.horizontal) ? x + frame.width - 6 : x + 6;
        const textAnchor = alignment.horizontal === "center" ? "middle" : ["right", "end"].includes(alignment.horizontal) ? "end" : "start";
        const textY = y + frame.height / 2;
        const textDecoration = [font.underline ? "underline" : "", font.strike ? "line-through" : ""].filter(Boolean).join(" ");
        const rotation = Number(alignment.textRotation || 0);
        rows.push(`<rect x="${x}" y="${y}" width="${frame.width}" height="${frame.height}" fill="${xmlEscape(fill.paint)}" stroke="#d0d7de"/>`);
        rows.push(`<text x="${textX}" y="${textY}" text-anchor="${textAnchor}" dominant-baseline="middle"${textDecoration ? ` text-decoration="${textDecoration}"` : ""}${rotation ? ` transform="rotate(${-rotation} ${textX} ${textY})"` : ""} font-family="${xmlEscape(font.name || "Arial")}" font-size="13" font-weight="${fontWeight}" fill="${xmlEscape(fontFill)}">${xmlEscape(value)}</text>`);
      }
    }
    const tableOverlays = this.tables.items.map((table) => table.toSvg(bounds)).join("");
    const pivotOverlays = this.pivotTables.items.map((pivot) => pivot.toSvg(bounds)).join("");
    const chartOverlays = this.charts.items.map((chart) => chart.toSvg()).join("");
    const imageOverlays = this.images.items.map((image) => image.toSvg()).join("");
    const sparklineOverlays = this.sparklineGroups.items.map((group) => group.toSvg(bounds)).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${fillDefinitions.length ? `<defs>${fillDefinitions.join("")}</defs>` : ""}<rect width="100%" height="100%" fill="#f6f8fa"/>${rows.join("")}${tableOverlays}${pivotOverlays}${chartOverlays}${imageOverlays}${sparklineOverlays}</svg>`;
  }
}

const RANGE_DIMENSION_FORMAT_KEYS = new Set(["columnWidth", "columnWidthPx", "rowHeight", "rowHeightPx", "columnHidden", "rowHidden"]);

function rangeFirstCellStyle(range) {
  return range.worksheet.store.get(makeCellAddress(range.bounds.top, range.bounds.left)).style || {};
}

function setWorksheetDimension(map, index, patch = {}) {
  const next = { ...(map.get(index) || {}), ...patch };
  for (const key of Object.keys(next)) if (next[key] == null || next[key] === false) delete next[key];
  if (Object.keys(next).length) map.set(index, next);
  else map.delete(index);
}

class RangeFormatFacade {
  constructor(range) { this.range = range; }

  get columnWidth() {
    const dimension = worksheetColumnDimension(this.range.worksheet, this.range.bounds.left);
    return dimension.width == null ? XLSX_DEFAULT_COLUMN_WIDTH : xlsxColumnWidthToCharacters(dimension.width);
  }
  set columnWidth(value) {
    const width = xlsxColumnCharactersToWidth(value);
    for (let column = this.range.bounds.left; column <= this.range.bounds.right; column += 1) setWorksheetDimension(this.range.worksheet.columnDimensions, column, { width, bestFit: false });
  }
  get columnWidthPx() { return worksheetColumnWidthPx(this.range.worksheet, this.range.bounds.left, undefined, false); }
  set columnWidthPx(value) {
    const width = xlsxColumnPixelsToWidth(value);
    for (let column = this.range.bounds.left; column <= this.range.bounds.right; column += 1) setWorksheetDimension(this.range.worksheet.columnDimensions, column, { width, bestFit: false });
  }
  get rowHeight() { return worksheetRowDimension(this.range.worksheet, this.range.bounds.top).height ?? XLSX_DEFAULT_ROW_HEIGHT; }
  set rowHeight(value) {
    const height = xlsxRowHeight(value);
    for (let row = this.range.bounds.top; row <= this.range.bounds.bottom; row += 1) setWorksheetDimension(this.range.worksheet.rowDimensions, row, { height });
  }
  get rowHeightPx() { return worksheetRowHeightPx(this.range.worksheet, this.range.bounds.top, undefined, false); }
  set rowHeightPx(value) { this.rowHeight = Number(value) * 72 / 96; }
  get columnHidden() { return Boolean(worksheetColumnDimension(this.range.worksheet, this.range.bounds.left).hidden); }
  set columnHidden(value) {
    for (let column = this.range.bounds.left; column <= this.range.bounds.right; column += 1) setWorksheetDimension(this.range.worksheet.columnDimensions, column, { hidden: Boolean(value) });
  }
  get rowHidden() { return Boolean(worksheetRowDimension(this.range.worksheet, this.range.bounds.top).hidden); }
  set rowHidden(value) {
    for (let row = this.range.bounds.top; row <= this.range.bounds.bottom; row += 1) setWorksheetDimension(this.range.worksheet.rowDimensions, row, { hidden: Boolean(value) });
  }

  autofitColumns() {
    for (let column = this.range.bounds.left; column <= this.range.bounds.right; column += 1) {
      let pixels = 20;
      for (let row = this.range.bounds.top; row <= this.range.bounds.bottom; row += 1) {
        const value = this.range.worksheet.store.get(makeCellAddress(row, column)).value;
        const longest = String(value ?? "").split(/\r?\n/).reduce((max, line) => Math.max(max, [...line].length), 0);
        pixels = Math.max(pixels, Math.min(1_790, longest * XLSX_MAX_DIGIT_WIDTH_PX + 12));
      }
      setWorksheetDimension(this.range.worksheet.columnDimensions, column, { width: xlsxColumnPixelsToWidth(pixels), bestFit: true });
    }
    return this.range;
  }

  autofitRows() {
    for (let row = this.range.bounds.top; row <= this.range.bounds.bottom; row += 1) {
      let pixels = 20;
      for (let column = this.range.bounds.left; column <= this.range.bounds.right; column += 1) {
        const cell = this.range.worksheet.store.get(makeCellAddress(row, column));
        const text = String(cell.value ?? "");
        const explicitLines = Math.max(1, text.split(/\r?\n/).length);
        const columnPixels = Math.max(1, worksheetColumnWidthPx(this.range.worksheet, column));
        const wrappedLines = cell.style?.alignment?.wrapText ? Math.max(explicitLines, Math.ceil((text.length * XLSX_MAX_DIGIT_WIDTH_PX + 8) / columnPixels)) : explicitLines;
        const fontPoints = Number(cell.style?.font?.size || cell.style?.fontSize || 11);
        pixels = Math.max(pixels, wrappedLines * fontPoints * 1.6);
      }
      setWorksheetDimension(this.range.worksheet.rowDimensions, row, { height: xlsxRowHeight(Math.min(XLSX_MAX_ROW_HEIGHT, pixels * 72 / 96)) });
    }
    return this.range;
  }

  toJSON() {
    return { ...rangeFirstCellStyle(this.range), columnWidth: this.columnWidth, columnWidthPx: this.columnWidthPx, rowHeight: this.rowHeight, rowHeightPx: this.rowHeightPx, columnHidden: this.columnHidden, rowHidden: this.rowHidden };
  }
}

function rangeFormatFacade(range) {
  const target = new RangeFormatFacade(range);
  return new Proxy(target, {
    get(current, property, receiver) {
      if (Reflect.has(current, property)) {
        const value = Reflect.get(current, property, receiver);
        return typeof value === "function" ? value.bind(current) : value;
      }
      return rangeFirstCellStyle(range)[property];
    },
    set(current, property, value, receiver) {
      if (Reflect.has(current, property)) return Reflect.set(current, property, value, receiver);
      range.writeStyle({ [property]: value });
      return true;
    },
    ownKeys() { return [...new Set([...Object.keys(rangeFirstCellStyle(range)), ...RANGE_DIMENSION_FORMAT_KEYS])]; },
    getOwnPropertyDescriptor() { return { enumerable: true, configurable: true }; },
  });
}

function projectedRangeFormula(worksheet, address, cell) {
  const directAnchor = cell?.spillAnchor || (cell?.spillParent ? String(cell.spillParent).slice(String(cell.spillParent).lastIndexOf("!") + 1) : undefined);
  if (directAnchor) {
    const anchorCell = worksheet.store.cells.get(directAnchor);
    if (anchorCell?.formula) {
      const reference = cell.spillRange || anchorCell.spillRange || anchorCell.dynamicArrayRef || directAnchor;
      return { anchorAddress: directAnchor, anchorCell, reference, source: "spill" };
    }
  }
  const position = parseCellAddress(address);
  for (const [anchorAddress, anchorCell] of worksheet.store.entries()) {
    if (!anchorCell.formula) continue;
    const reference = anchorCell.dynamicArrayRef || anchorCell.arrayRef;
    if (!reference) continue;
    let bounds;
    try { bounds = parseRangeAddress(reference); } catch { continue; }
    if (position.row < bounds.top || position.row > bounds.bottom || position.col < bounds.left || position.col > bounds.right) continue;
    if (address === anchorAddress) continue;
    return { anchorAddress, anchorCell, reference, source: anchorCell.formulaType === "array" ? "array" : "spill" };
  }
  return undefined;
}

function rangeFormulaInfo(worksheet, address) {
  const cell = worksheet.store.cells.get(address);
  if (cell?.formula) return { kind: "stored", formula: cell.formula, display: cell.formula, isEditable: true };
  const projected = projectedRangeFormula(worksheet, address, cell);
  if (!projected) return null;
  const reference = String(projected.reference).slice(String(projected.reference).lastIndexOf("!") + 1);
  return {
    kind: "projected",
    source: projected.source,
    display: projected.anchorCell.formula,
    anchor: `${worksheet.name}!${projected.anchorAddress}`,
    ref: `${worksheet.name}!${reference}`,
    isEditable: false,
  };
}

export class Range {
  constructor(worksheet, bounds) {
    this.worksheet = worksheet;
    this.bounds = assertWorksheetBounds(bounds);
  }

  get address() { return rangeToAddress(this.bounds); }
  get rowIndex() { return this.bounds.top; }
  get columnIndex() { return this.bounds.left; }
  get rowCount() { return this.bounds.rowCount; }
  get columnCount() { return this.bounds.colCount; }

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
      for (let c = this.bounds.left; c <= this.bounds.right; c++) row.push(this.worksheet.store.get(makeCellAddress(r, c)).formula || "");
      out.push(row);
    }
    return out;
  }

  set formulas(matrix) {
    this.writeMatrix(matrix, "formula");
  }

  get formulasR1C1() {
    const out = [];
    for (let r = this.bounds.top; r <= this.bounds.bottom; r += 1) {
      const row = [];
      for (let c = this.bounds.left; c <= this.bounds.right; c += 1) {
        const address = makeCellAddress(r, c);
        const formula = this.worksheet.store.get(address).formula;
        row.push(formula ? formulaA1ToR1C1(formula, address) : "");
      }
      out.push(row);
    }
    return out;
  }

  set formulasR1C1(matrix) {
    const rows = normalizeRangeWrite(matrix, "formulasR1C1").matrix;
    const formulas = rows.map((row, rowIndex) => row.map((formula, columnIndex) => {
      if (formula == null || formula === "") return "";
      return formulaR1C1ToA1(String(formula), makeCellAddress(this.bounds.top + rowIndex, this.bounds.left + columnIndex));
    }));
    this.writeMatrix(formulas, "formula");
  }

  get formulaInfos() {
    const out = [];
    for (let r = this.bounds.top; r <= this.bounds.bottom; r += 1) {
      const row = [];
      for (let c = this.bounds.left; c <= this.bounds.right; c += 1) row.push(rangeFormulaInfo(this.worksheet, makeCellAddress(r, c)));
      out.push(row);
    }
    return out;
  }

  get displayFormulas() {
    return this.formulaInfos.map((row) => row.map((info) => info?.display || ""));
  }

  get format() {
    return rangeFormatFacade(this);
  }

  set format(style) {
    this.writeStyle(style);
  }

  get style() { return this.format; }
  set style(value) { this.format = value; }
  setFormat(style = {}) { this.writeStyle(style); return this; }
  setNumberFormat(format) { this.writeStyle({ numberFormat: format }); return this; }

  writeStyle(style = {}) {
    const cellStyle = { ...(style || {}) };
    const numberFormats = Array.isArray(cellStyle.numberFormat)
      ? normalizeRangeWrite(cellStyle.numberFormat, "values").matrix
      : undefined;
    if (numberFormats) {
      if (this.rowCount % numberFormats.length !== 0 || this.columnCount % numberFormats[0].length !== 0) {
        throw new Error(`Range number-format matrix ${numberFormats.length}x${numberFormats[0].length} must evenly tile ${this.rowCount}x${this.columnCount}.`);
      }
      delete cellStyle.numberFormat;
    }
    const dimensions = {};
    for (const key of RANGE_DIMENSION_FORMAT_KEYS) {
      if (Object.prototype.hasOwnProperty.call(cellStyle, key)) {
        dimensions[key] = cellStyle[key];
        delete cellStyle[key];
      }
    }
    for (let r = this.bounds.top; r <= this.bounds.bottom; r++) {
      for (let c = this.bounds.left; c <= this.bounds.right; c++) {
        const cell = this.worksheet.store.get(makeCellAddress(r, c));
        cell.style = { ...(cell.style || {}), ...cellStyle };
        if (numberFormats) cell.style.numberFormat = numberFormats[(r - this.rowIndex) % numberFormats.length][(c - this.columnIndex) % numberFormats[0].length];
      }
    }
    const facade = rangeFormatFacade(this);
    for (const [key, value] of Object.entries(dimensions)) facade[key] = value;
  }

  writeMatrix(matrix, field) {
    const rows = normalizeRangeWrite(matrix, field === "formula" ? "formulas" : "values").matrix;
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < (rows[r]?.length ?? 0); c++) {
        const address = makeCellAddress(this.bounds.top + r, this.bounds.left + c);
        detachNativeFormulaTopology(this.worksheet, address);
        const cell = this.worksheet.store.get(address);
        delete cell.spillParent;
        delete cell.spillAnchor;
        delete cell.spillRange;
        delete cell.spillValues;
        delete cell.spillError;
        if (field === "formula") {
          const formula = rows[r][c] == null ? "" : String(rows[r][c]);
          cell.formula = formula || null;
          if (!formula) cell.value = null;
        } else {
          cell.value = rows[r][c];
          cell.formula = null;
        }
      }
    }
    this.worksheet.recalculate();
  }

  write(value) {
    const normalized = normalizeRangeWrite(value);
    const bounds = writtenRangeBounds(this.bounds, normalized.matrix);
    const result = new Range(this.worksheet, bounds);
    if (normalized.field === "formulasR1C1") result.formulasR1C1 = normalized.matrix;
    else if (normalized.field === "formulas") result.writeMatrix(normalized.matrix, "formula");
    else if (normalized.field === "values") result.writeMatrix(normalized.matrix, "value");
    else {
      for (let row = 0; row < normalized.matrix.length; row += 1) {
        for (let column = 0; column < normalized.matrix[row].length; column += 1) {
          const valueAtCell = normalized.matrix[row][column];
          const address = makeCellAddress(bounds.top + row, bounds.left + column);
          detachNativeFormulaTopology(this.worksheet, address);
          const cell = this.worksheet.store.get(address);
          for (const key of ["formulaType", "sharedIndex", "sharedRef", "arrayRef", "dynamicArrayRef", "spillParent", "spillAnchor", "spillRange", "spillValues", "spillError"]) delete cell[key];
          if (typeof valueAtCell === "string" && valueAtCell.startsWith("=")) {
            cell.formula = valueAtCell;
            cell.value = null;
          } else {
            cell.value = valueAtCell;
            cell.formula = null;
          }
        }
      }
      this.worksheet.recalculate();
    }
    return result;
  }

  writeValues(value) {
    this.write({ values: value });
  }

  clear(options = {}) {
    const applyTo = options?.applyTo || "all";
    if (!["contents", "formats", "all"].includes(applyTo)) throw new Error("Range.clear applyTo must be contents, formats, or all.");
    for (let r = this.bounds.top; r <= this.bounds.bottom; r += 1) {
      for (let c = this.bounds.left; c <= this.bounds.right; c += 1) {
        const address = makeCellAddress(r, c);
        if (applyTo === "all") {
          detachNativeFormulaTopology(this.worksheet, address);
          this.worksheet.store.cells.delete(address);
          continue;
        }
        const cell = this.worksheet.store.get(address);
        if (applyTo === "contents") {
          detachNativeFormulaTopology(this.worksheet, address);
          cell.value = null;
          cell.formula = null;
          for (const key of ["formulaType", "sharedIndex", "sharedRef", "arrayRef", "dynamicArrayRef", "spillParent", "spillAnchor", "spillRange", "spillValues", "spillError"]) delete cell[key];
        } else cell.style = {};
      }
    }
    this.worksheet.recalculate();
  }

  copyFrom(sourceRange, mode = "all") {
    if (!(sourceRange instanceof Range)) throw new Error("Range.copyFrom requires a source Range.");
    const copyMode = normalizeRangeCopyMode(mode);
    assertRangeCopyShape(sourceRange.bounds, this.bounds);
    const sourceCells = new Map();
    for (let row = 0; row < sourceRange.rowCount; row += 1) {
      for (let column = 0; column < sourceRange.columnCount; column += 1) {
        const address = makeCellAddress(sourceRange.rowIndex + row, sourceRange.columnIndex + column);
        sourceCells.set(`${row}:${column}`, { address, cell: structuredClone(sourceRange.worksheet.store.get(address)) });
      }
    }
    for (let row = 0; row < this.rowCount; row += 1) {
      for (let column = 0; column < this.columnCount; column += 1) {
        const source = sourceCells.get(`${row % sourceRange.rowCount}:${column % sourceRange.columnCount}`);
        const targetAddress = makeCellAddress(this.rowIndex + row, this.columnIndex + column);
        detachNativeFormulaTopology(this.worksheet, targetAddress);
        const target = this.worksheet.store.get(targetAddress);
        if (copyMode === "values") {
          target.value = structuredClone(source.cell.value);
          target.formula = null;
          for (const key of ["formulaType", "sharedIndex", "sharedRef", "arrayRef", "dynamicArrayRef", "spillParent", "spillAnchor", "spillRange", "spillValues", "spillError"]) delete target[key];
          continue;
        }
        const translatedFormula = source.cell.formula ? translateA1Formula(source.cell.formula, source.address, targetAddress) : null;
        if (copyMode === "formulas") {
          target.formula = translatedFormula;
          target.value = translatedFormula ? null : structuredClone(source.cell.value);
          for (const key of ["formulaType", "sharedIndex", "sharedRef", "arrayRef", "dynamicArrayRef", "spillParent", "spillAnchor", "spillRange", "spillValues", "spillError"]) delete target[key];
          continue;
        }
        const copy = structuredClone(source.cell);
        copy.formula = translatedFormula;
        for (const key of ["formulaType", "sharedIndex", "sharedRef", "arrayRef", "dynamicArrayRef", "spillParent", "spillAnchor", "spillRange", "spillValues", "spillError"]) delete copy[key];
        this.worksheet.store.cells.set(targetAddress, copy);
      }
    }
    this.worksheet.recalculate();
  }

  copyTo(destinationRange, mode = "all") {
    if (!(destinationRange instanceof Range)) throw new Error("Range.copyTo requires a destination Range.");
    destinationRange.copyFrom(this, mode);
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

  merge(across = false) { this.worksheet.mergeCells(this, across); return this; }
  unmerge() { this.worksheet.unmergeCells(this); return this; }
  fillDown() {
    if (this.worksheet.mergedRanges.some((range) => worksheetBoundsIntersect(parseRangeAddress(range), this.bounds))) throw new Error(`Cannot fill down range ${rangeToAddress(this.bounds)} because it intersects merged cells.`);
    const sources = new Map();
    for (let column = this.bounds.left; column <= this.bounds.right; column += 1) {
      const address = makeCellAddress(this.bounds.top, column);
      sources.set(column, { address, cell: structuredClone(this.worksheet.store.get(address)) });
    }
    for (let row = this.bounds.top + 1; row <= this.bounds.bottom; row += 1) {
      for (let column = this.bounds.left; column <= this.bounds.right; column += 1) {
        const source = sources.get(column);
        const targetAddress = makeCellAddress(row, column);
        this.worksheet.store.cells.set(targetAddress, cloneCellForFill(source.cell, source.address, targetAddress));
      }
    }
    this.worksheet.recalculate();
    return this;
  }
  fillRight() {
    if (this.worksheet.mergedRanges.some((range) => worksheetBoundsIntersect(parseRangeAddress(range), this.bounds))) throw new Error(`Cannot fill right range ${rangeToAddress(this.bounds)} because it intersects merged cells.`);
    const sources = new Map();
    for (let row = this.bounds.top; row <= this.bounds.bottom; row += 1) {
      const address = makeCellAddress(row, this.bounds.left);
      sources.set(row, { address, cell: structuredClone(this.worksheet.store.get(address)) });
    }
    for (let column = this.bounds.left + 1; column <= this.bounds.right; column += 1) {
      for (let row = this.bounds.top; row <= this.bounds.bottom; row += 1) {
        const source = sources.get(row);
        const targetAddress = makeCellAddress(row, column);
        this.worksheet.store.cells.set(targetAddress, cloneCellForFill(source.cell, source.address, targetAddress));
      }
    }
    this.worksheet.recalculate();
    return this;
  }
  getCell(row, col) { return this.getRangeByIndexes(row, col, 1, 1); }
  getRange(address) { return this.worksheet.getRange(address); }
  getRangeByIndexes(startRow, startCol, rowCount, colCount) {
    const relative = rangeBounds(startRow, startCol, startRow + rowCount - 1, startCol + colCount - 1);
    if (!Number.isInteger(startRow) || !Number.isInteger(startCol) || !Number.isInteger(rowCount) || !Number.isInteger(colCount)
      || rowCount < 1 || colCount < 1 || relative.top < 0 || relative.left < 0
      || relative.bottom >= this.rowCount || relative.right >= this.columnCount) {
      throw new Error(`Range.getRangeByIndexes(${startRow}, ${startCol}, ${rowCount}, ${colCount}) is outside the bounds of ${this.address}.`);
    }
    return this.worksheet.getRangeByIndexes(this.rowIndex + startRow, this.columnIndex + startCol, rowCount, colCount);
  }
  getRow(index) { return this.getRangeByIndexes(index, 0, 1, this.columnCount); }
  getColumn(index) { return this.getRangeByIndexes(0, index, this.rowCount, 1); }
  getCurrentRegion() {
    const populated = (row, column) => {
      const cell = this.worksheet.store.cells.get(makeCellAddress(row, column));
      return Boolean(cell && (cell.value != null || cell.formula));
    };
    return new Range(this.worksheet, currentRegionBounds(this.bounds, populated));
  }
  getOffsetRange(rowOffset, colOffset) {
    const top = this.rowIndex + Number(rowOffset);
    const left = this.columnIndex + Number(colOffset);
    const bounds = rangeBounds(top, left, top + this.rowCount - 1, left + this.columnCount - 1);
    return new Range(this.worksheet, assertWorksheetBounds(bounds, "Range.getOffsetRange"));
  }
  getResizedRange(rowCount, colCount) {
    if (!Number.isInteger(rowCount) || !Number.isInteger(colCount) || rowCount < 1 || colCount < 1) throw new Error("Range.getResizedRange requires positive integer row and column counts.");
    const bounds = rangeBounds(this.rowIndex, this.columnIndex, this.rowIndex + rowCount - 1, this.columnIndex + colCount - 1);
    return new Range(this.worksheet, assertWorksheetBounds(bounds, "Range.getResizedRange"));
  }
  offset(rowOffset, colOffset) { return this.getOffsetRange(rowOffset, colOffset); }
  resize(rowCount, colCount) { return this.getResizedRange(rowCount, colCount); }
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

function findContainingWorkbookTable(sheet, address) {
  if (!sheet || !address) return undefined;
  const cell = parseCellAddress(String(address).replaceAll("$", "").toUpperCase());
  const table = sheet.tables.items.find((item) => {
    const bounds = parseRangeAddress(item.range);
    return cell.row >= bounds.top && cell.row <= bounds.bottom && cell.col >= bounds.left && cell.col <= bounds.right;
  });
  return table ? { sheet, table } : undefined;
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

function formulaStructuredRefRange(sheet, refText = "", context = {}) {
  const parsed = parseStructuredReference(refText);
  if (!parsed) return undefined;
  const tableName = parsed.tableName;
  const found = tableName ? findWorkbookTable(sheet, tableName) : findContainingWorkbookTable(sheet, context.formulaAddress);
  const tokens = parsed.tokens;
  if (!found) return { missing: true, tableName, columnName: tokens.join(",") };
  const { table, sheet: tableSheet } = found;
  const bounds = parseRangeAddress(table.range);
  const columnCount = bounds.right - bounds.left + 1;
  const headers = Array.from({ length: columnCount }, (_, index) => table.columnNames?.[index] ?? (table.hasHeaders ? (table.values[0]?.[index] ?? `Column${index + 1}`) : `Column${index + 1}`));
  const currentRow = parsed.currentRow || (!parsed.qualified && parsed.sectionTokens.length === 0);
  const section = currentRow ? "#This Row" : parsed.sectionTokens.at(-1) || "#Data";
  const firstDataRow = bounds.top + (table.showHeaders ? 1 : 0);
  const totalsRow = table.showTotals ? bounds.bottom : undefined;
  const lastDataRow = bounds.bottom - (table.showTotals ? 1 : 0);
  let top = firstDataRow;
  let bottom = lastDataRow;
  if (currentRow) {
    const formulaCell = context.formulaAddress ? parseCellAddress(String(context.formulaAddress).replaceAll("$", "").toUpperCase()) : undefined;
    if (!formulaCell || formulaCell.row < firstDataRow || formulaCell.row > lastDataRow || formulaCell.col < bounds.left || formulaCell.col > bounds.right) return { error: "#VALUE!", tableName: table.name, sheetName: tableSheet.name, table, section };
    top = bottom = formulaCell.row;
  } else if (/^#Headers$/i.test(section)) {
    if (!table.showHeaders) return { sheetName: tableSheet.name, start: makeCellAddress(bounds.top, bounds.left), end: makeCellAddress(bounds.top - 1, bounds.left), empty: true, tableName, columnName: parsed.columnSelectors.map((item) => item.name || `${item.start}:${item.end}`).join(","), table };
    top = bottom = bounds.top;
  } else if (/^#Totals$/i.test(section)) {
    if (totalsRow == null) return { sheetName: tableSheet.name, start: makeCellAddress(bounds.bottom + 1, bounds.left), end: makeCellAddress(bounds.bottom, bounds.left), empty: true, tableName, columnName: parsed.columnSelectors.map((item) => item.name || `${item.start}:${item.end}`).join(","), table };
    top = bottom = totalsRow;
  } else if (/^#All$/i.test(section)) {
    top = bounds.top;
    bottom = bounds.bottom;
  } else if (/^#Data$/i.test(section)) {
    top = firstDataRow;
    bottom = lastDataRow;
  } else if (parsed.sectionTokens.length) {
    return { missing: true, tableName, columnName: section, sheetName: tableSheet.name };
  }
  const selected = structuredColumnIndexes(parsed.columnSelectors, headers);
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

function structuredRangeGeometry(structured) {
  const start = parseCellAddress(structured.start);
  const end = parseCellAddress(structured.end);
  const tableBounds = structured.table ? parseRangeAddress(structured.table.range) : undefined;
  const columns = structured.absoluteColumns || (tableBounds && structured.columns?.length
    ? structured.columns.map((index) => tableBounds.left + index)
    : Array.from({ length: Math.abs(end.col - start.col) + 1 }, (_, index) => Math.min(start.col, end.col) + index));
  return {
    top: Math.min(start.row, end.row),
    bottom: Math.max(start.row, end.row),
    columns,
  };
}

function formulaStructuredRefIntersection(sheet, refText = "", context = {}) {
  const text = String(refText || "").trim();
  const groups = scanStructuredReferenceIntersections(text);
  if (groups.length !== 1 || groups[0].start !== 0 || groups[0].end !== text.length) return undefined;
  const group = groups[0];
  const ranges = group.references.map((reference) => formulaStructuredRefRange(sheet, reference.text, context));
  const failure = ranges.find((range) => !range || range.error || range.missing);
  if (failure) return failure || { missing: true };
  if (ranges.some((range) => range.empty)) return { error: "#NULL!", intersectionReferences: group.references.map((reference) => reference.text) };
  const sheetNames = new Set(ranges.map((range) => range.sheetName || sheet.name));
  if (sheetNames.size !== 1) return { error: "#NULL!", intersectionReferences: group.references.map((reference) => reference.text) };
  const geometries = ranges.map(structuredRangeGeometry);
  const top = Math.max(...geometries.map((geometry) => geometry.top));
  const bottom = Math.min(...geometries.map((geometry) => geometry.bottom));
  const commonColumns = geometries[0].columns.filter((column) => geometries.slice(1).every((geometry) => geometry.columns.includes(column)));
  if (top > bottom || commonColumns.length === 0) return { error: "#NULL!", intersectionReferences: group.references.map((reference) => reference.text) };
  commonColumns.sort((left, right) => left - right);
  const sharedTable = ranges.every((range) => range.table === ranges[0].table) ? ranges[0].table : undefined;
  const sharedTableBounds = sharedTable ? parseRangeAddress(sharedTable.range) : undefined;
  const columns = sharedTableBounds ? commonColumns.map((column) => column - sharedTableBounds.left) : undefined;
  const columnNames = columns?.map((column) => sharedTable.columnNames?.[column] ?? `Column${column + 1}`);
  return {
    sheetName: [...sheetNames][0],
    start: makeCellAddress(top, commonColumns[0]),
    end: makeCellAddress(bottom, commonColumns.at(-1)),
    absoluteColumns: commonColumns,
    table: sharedTable,
    tableName: sharedTable?.name,
    columns,
    columnNames,
    columnName: columnNames?.join(","),
    section: "intersection",
    intersectionReferences: group.references.map((reference) => reference.text),
  };
}

function formulaA1RefIntersection(sheet, refText = "") {
  const parts = splitReferenceIntersectionOperands(refText);
  if (!parts) return undefined;
  const references = parts.map(formulaRefParts);
  if (references.some((reference) => !reference)) return undefined;
  const sheetNames = new Set(references.map((reference) => reference.sheetName || sheet.name));
  if (sheetNames.size !== 1) return { error: "#NULL!", intersectionReferences: parts };
  const bounds = references.map((reference) => {
    const start = parseCellAddress(reference.start);
    const end = parseCellAddress(reference.end);
    return {
      top: Math.min(start.row, end.row),
      bottom: Math.max(start.row, end.row),
      left: Math.min(start.col, end.col),
      right: Math.max(start.col, end.col),
    };
  });
  const top = Math.max(...bounds.map((item) => item.top));
  const bottom = Math.min(...bounds.map((item) => item.bottom));
  const left = Math.max(...bounds.map((item) => item.left));
  const right = Math.min(...bounds.map((item) => item.right));
  if (top > bottom || left > right) return { error: "#NULL!", intersectionReferences: parts };
  return {
    sheetName: [...sheetNames][0],
    start: makeCellAddress(top, left),
    end: makeCellAddress(bottom, right),
    absoluteColumns: Array.from({ length: right - left + 1 }, (_, index) => left + index),
    section: "intersection",
    intersectionReferences: parts,
  };
}

function scanA1ReferenceIntersections(formula = "") {
  const text = String(formula || "");
  const structured = scanStructuredReferences(text);
  const referenceRegex = /(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?/g;
  const references = [...text.matchAll(referenceRegex)]
    .map((match) => ({ text: match[0], start: match.index, end: match.index + match[0].length }))
    .filter((reference) => !structured.some((item) => reference.start >= item.start && reference.end <= item.end));
  const groups = [];
  for (let index = 0; index < references.length;) {
    const group = [references[index]];
    let cursor = index + 1;
    while (cursor < references.length && /^\s+$/.test(text.slice(group.at(-1).end, references[cursor].start))) {
      group.push(references[cursor]);
      cursor += 1;
    }
    if (group.length > 1) groups.push({ text: text.slice(group[0].start, group.at(-1).end), start: group[0].start, end: group.at(-1).end, references: group });
    index = cursor;
  }
  return groups;
}

function formulaReferences(formula, sheet, formulaAddress) {
  const raw = String(formula || "");
  const refs = [];
  const consumed = [];
  const intersectionGroups = scanStructuredReferenceIntersections(raw);
  const appendStructuredReferences = (structured, refText) => {
    if (!structured || structured.missing || structured.empty || structured.error) return;
    const start = parseCellAddress(structured.start);
    const end = parseCellAddress(structured.end);
    const cols = structuredRangeGeometry(structured).columns;
    for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
      for (const col of cols) {
        refs.push({ sheetName: structured.sheetName, address: makeCellAddress(row, col), structuredRef: refText, tableName: structured.tableName, columnName: structured.columnName, columnNames: structured.columnNames });
      }
    }
  };
  for (const group of intersectionGroups) {
    consumed.push([group.start, group.end]);
    appendStructuredReferences(formulaStructuredRefIntersection(sheet, group.text, { formulaAddress }), group.text);
  }
  for (const group of scanA1ReferenceIntersections(raw)) {
    consumed.push([group.start, group.end]);
    appendStructuredReferences(formulaA1RefIntersection(sheet, group.text), group.text);
  }
  for (const match of scanStructuredReferences(raw)) {
    if (intersectionGroups.some((group) => match.start >= group.start && match.end <= group.end)) continue;
    consumed.push([match.start, match.end]);
    const structured = formulaStructuredRefRange(sheet, match.text, { formulaAddress });
    appendStructuredReferences(structured, match.text);
  }
  const rangeRegex = /(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]+\$?\d+)\s*:\s*(\$?[A-Za-z]+\$?\d+)/g;
  for (const match of raw.matchAll(rangeRegex)) {
    if (consumed.some(([start, end]) => match.index >= start && match.index < end)) continue;
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
    for (const ref of formulaReferences(node.formula, node.sheetObject, node.address)) {
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
  const structured = formulaStructuredRefIntersection(sheet, refText, context) || formulaA1RefIntersection(sheet, refText) || formulaStructuredRefRange(sheet, refText, context);
  if (structured) {
    if (structured.error) return [[structured.error]];
    if (structured.missing) return [["#REF!"]];
    if (structured.empty) return [];
    const targetSheet = structured.sheetName ? sheet.workbook?.worksheets.getItem(structured.sheetName) : sheet;
    if (!targetSheet) return [["#REF!"]];
    const start = parseCellAddress(structured.start);
    const end = parseCellAddress(structured.end);
    const cols = structuredRangeGeometry(structured).columns;
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

function hydrateDeclaredDynamicArraySpills(workbook) {
  for (const sheet of workbook.worksheets) {
    const cells = sheet.store.entries();
    for (const [anchorAddress, anchorCell] of cells) {
      if (anchorCell.formulaType !== "dynamicArray" || !anchorCell.formula || !anchorCell.dynamicArrayRef) continue;
      let bounds;
      try {
        bounds = parseRangeAddress(anchorCell.dynamicArrayRef);
      } catch {
        continue;
      }
      if (makeCellAddress(bounds.top, bounds.left) !== anchorAddress) continue;
      markDeclaredDynamicArrayChildren(sheet, anchorAddress, anchorCell.dynamicArrayRef, bounds, cells);
    }
  }
}

function markDeclaredDynamicArrayChildren(sheet, anchorAddress, reference, bounds, cells = sheet.store.entries()) {
  const parentKey = formulaCellKey(sheet.name, anchorAddress);
  for (const [address, cell] of cells) {
    if (address === anchorAddress || cell.formula) continue;
    const position = parseCellAddress(address);
    if (position.row < bounds.top || position.row > bounds.bottom || position.col < bounds.left || position.col > bounds.right) continue;
    cell.spillParent = parentKey;
    cell.spillAnchor = anchorAddress;
    cell.spillRange = reference;
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

function statisticalFormulaNumbers(values) {
  const error = values.map(formulaErrorCode).find(Boolean);
  if (error) return { error, numbers: [] };
  return { numbers: values.filter((value) => typeof value === "number" && Number.isFinite(value)) };
}

function roundFormulaNumber(value, digits = 0, mode = "nearest") {
  const number = formulaNumber(value);
  if (formulaErrorCode(number)) return number;
  const places = Math.trunc(Number(digits) || 0);
  const factor = 10 ** Math.min(308, Math.abs(places));
  const scaled = places >= 0 ? Math.abs(number) * factor : Math.abs(number) / factor;
  const rounded = mode === "up" ? Math.ceil(scaled) : mode === "down" ? Math.floor(scaled) : Math.round(scaled + Number.EPSILON * Math.max(1, scaled));
  const result = places >= 0 ? rounded / factor : rounded * factor;
  return Object.is(number, -0) || number < 0 ? -result : result;
}

const EXCEL_1900_DATE_EPOCH_UTC = Date.UTC(1899, 11, 31);
const EXCEL_1904_DATE_EPOCH_UTC = Date.UTC(1904, 0, 1);
const EXCEL_MAX_DATE_SERIALS = { "1900": 2_958_465, "1904": 2_957_003 };

function excelFormulaDateSystem(sheet) {
  return sheet?.workbook?.dateSystem === "1904" ? "1904" : "1900";
}

function excelMaxDateSerial(dateSystem = "1900") {
  return EXCEL_MAX_DATE_SERIALS[dateSystem === "1904" ? "1904" : "1900"];
}

function excelFormulaDateNumber(value) {
  const error = formulaErrorCode(value);
  if (error) return error;
  if (value == null || value === "" || value === false) return 0;
  if (value === true) return 1;
  const number = Number(value);
  return Number.isFinite(number) ? number : "#VALUE!";
}

function excelGregorianSerial(year, month, day = 1, dateSystem = "1900") {
  const date = new Date(0);
  date.setUTCHours(0, 0, 0, 0);
  date.setUTCFullYear(year, month - 1, day);
  const epoch = dateSystem === "1904" ? EXCEL_1904_DATE_EPOCH_UTC : EXCEL_1900_DATE_EPOCH_UTC;
  const days = Math.round((date.getTime() - epoch) / 86_400_000);
  return dateSystem === "1904" ? days : days + (date.getTime() >= Date.UTC(1900, 2, 1) ? 1 : 0);
}

function excelDateSerial(yearValue, monthValue, dayValue, dateSystem = "1900") {
  let year = Math.trunc(yearValue);
  const month = Math.trunc(monthValue);
  const day = Math.trunc(dayValue);
  if (year >= 0 && year <= 1899) year += 1900;
  if (!Number.isFinite(year) || year < 0 || year > 9999 || !Number.isFinite(month) || !Number.isFinite(day)) return "#NUM!";
  const normalized = new Date(0);
  normalized.setUTCHours(0, 0, 0, 0);
  normalized.setUTCFullYear(year, month - 1, 1);
  const serial = excelGregorianSerial(normalized.getUTCFullYear(), normalized.getUTCMonth() + 1, 1, dateSystem) + day - 1;
  return serial < 0 || serial > excelMaxDateSerial(dateSystem) ? "#NUM!" : serial;
}

function excelDateParts(serialValue, dateSystem = "1900") {
  const serial = Math.floor(serialValue);
  if (!Number.isFinite(serial) || serial < 0 || serial > excelMaxDateSerial(dateSystem)) return undefined;
  if (dateSystem === "1904") {
    const date = new Date(EXCEL_1904_DATE_EPOCH_UTC + serial * 86_400_000);
    return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
  }
  if (serial === 0) return { year: 1900, month: 1, day: 0 };
  if (serial === 60) return { year: 1900, month: 2, day: 29 };
  const date = new Date(EXCEL_1900_DATE_EPOCH_UTC + (serial > 60 ? serial - 1 : serial) * 86_400_000);
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
}

function excelDateValue(value, dateSystem = "1900") {
  const parsed = parseFormulaDateText(value);
  if (!parsed) return "#VALUE!";
  if (dateSystem === "1900" && parsed.year === 1900 && parsed.month === 2 && parsed.day === 29) return 60;
  const serial = excelGregorianSerial(parsed.year, parsed.month, parsed.day, dateSystem);
  const restored = excelDateParts(serial, dateSystem);
  if (!restored || restored.year !== parsed.year || restored.month !== parsed.month || restored.day !== parsed.day) return "#VALUE!";
  return serial;
}

function excelTimeValue(value, dateSystem = "1900") {
  const time = parseFormulaTimeText(value);
  if (time) return time.dateText && formulaErrorCode(excelDateValue(time.dateText, dateSystem)) ? "#VALUE!" : time.serial;
  return formulaErrorCode(excelDateValue(value, dateSystem)) ? "#VALUE!" : 0;
}

function excelTimePart(value, part, dateSystem = "1900") {
  if (typeof value === "string") {
    const time = parseFormulaTimeText(value);
    if (time) return time.dateText && formulaErrorCode(excelDateValue(time.dateText, dateSystem)) ? "#VALUE!" : time[part];
    return formulaErrorCode(excelDateValue(value, dateSystem)) ? "#VALUE!" : 0;
  }
  const parsed = formulaTimeParts(value);
  if (parsed) return parsed[part];
  return "#VALUE!";
}

function excelDaysInMonth(year, month, dateSystem = "1900") {
  if (dateSystem === "1900" && year === 1900 && month === 2) return 29;
  const date = new Date(0);
  date.setUTCHours(0, 0, 0, 0);
  date.setUTCFullYear(year, month, 0);
  return date.getUTCDate();
}

function excelShiftMonth(serialValue, monthsValue, endOfMonth = false, dateSystem = "1900") {
  const serial = excelFormulaDateNumber(serialValue);
  const months = excelFormulaDateNumber(monthsValue);
  if (formulaErrorCode(serial)) return serial;
  if (formulaErrorCode(months)) return months;
  const parts = excelDateParts(serial, dateSystem);
  if (!parts) return "#NUM!";
  const first = new Date(0);
  first.setUTCHours(0, 0, 0, 0);
  first.setUTCFullYear(parts.year, parts.month - 1 + Math.trunc(months), 1);
  const year = first.getUTCFullYear();
  const month = first.getUTCMonth() + 1;
  const day = endOfMonth ? excelDaysInMonth(year, month, dateSystem) : Math.min(Math.max(1, parts.day), excelDaysInMonth(year, month, dateSystem));
  return excelDateSerial(year, month, day, dateSystem);
}

function excelWeekdayIndex(serial, dateSystem = "1900") {
  const day = Math.floor(serial);
  if (dateSystem === "1904") return ((day + 5) % 7 + 7) % 7;
  const adjusted = day > 60 ? day - 1 : day;
  return ((adjusted % 7) + 7) % 7;
}

function excelHolidaySet(values = [], dateSystem = "1900") {
  const error = values.map(formulaErrorCode).find(Boolean);
  if (error) return { error, holidays: new Set() };
  const holidays = new Set();
  for (const value of values) {
    if (value == null || value === "") continue;
    const serial = excelFormulaDateNumber(value);
    if (formulaErrorCode(serial)) return { error: serial, holidays: new Set() };
    const day = Math.floor(serial);
    if (day < 0 || day > excelMaxDateSerial(dateSystem)) return { error: "#NUM!", holidays: new Set() };
    holidays.add(day);
  }
  return { holidays };
}

function excelWeekendDays(value = 1, allowAllWeekend = false) {
  const error = formulaErrorCode(value);
  if (error) return { error, weekends: new Set() };
  if (typeof value === "string") {
    const weekend = value.trim();
    if (/^[01]{7}$/.test(weekend)) {
      if (weekend === "1111111" && !allowAllWeekend) return { error: "#VALUE!", weekends: new Set() };
      const weekends = new Set();
      for (let index = 0; index < 7; index += 1) if (weekend[index] === "1") weekends.add((index + 1) % 7);
      return { weekends };
    }
    if (!/^\d+(?:\.0+)?$/.test(weekend)) return { error: "#VALUE!", weekends: new Set() };
  }
  const weekendNumber = Number(value);
  if (!Number.isInteger(weekendNumber)) return { error: "#NUM!", weekends: new Set() };
  if (weekendNumber >= 1 && weekendNumber <= 7) {
    const first = weekendNumber === 1 ? 6 : weekendNumber - 2;
    return { weekends: new Set([first, (first + 1) % 7]) };
  }
  if (weekendNumber >= 11 && weekendNumber <= 17) return { weekends: new Set([weekendNumber - 11]) };
  return { error: "#NUM!", weekends: new Set() };
}

function excelBusinessDay(serial, holidays, dateSystem = "1900", weekends = new Set([0, 6])) {
  const weekday = excelWeekdayIndex(serial, dateSystem);
  return !weekends.has(weekday) && !holidays.has(serial);
}

function excelNetworkDays(startValue, endValue, holidayValues = [], dateSystem = "1900", weekendValue = 1, allowAllWeekend = false) {
  const startNumber = excelFormulaDateNumber(startValue);
  const endNumber = excelFormulaDateNumber(endValue);
  if (formulaErrorCode(startNumber)) return startNumber;
  if (formulaErrorCode(endNumber)) return endNumber;
  const start = Math.floor(startNumber), end = Math.floor(endNumber);
  if (!excelDateParts(start, dateSystem) || !excelDateParts(end, dateSystem)) return "#NUM!";
  const holidayResult = excelHolidaySet(holidayValues, dateSystem);
  if (holidayResult.error) return holidayResult.error;
  const weekendResult = excelWeekendDays(weekendValue, allowAllWeekend);
  if (weekendResult.error) return weekendResult.error;
  const direction = start <= end ? 1 : -1;
  const low = Math.min(start, end), high = Math.max(start, end);
  const total = high - low + 1;
  const fullWeeks = Math.floor(total / 7);
  let weekdays = fullWeeks * (7 - weekendResult.weekends.size);
  for (let serial = low + fullWeeks * 7; serial <= high; serial += 1) if (excelBusinessDay(serial, new Set(), dateSystem, weekendResult.weekends)) weekdays += 1;
  for (const holiday of holidayResult.holidays) if (holiday >= low && holiday <= high && !weekendResult.weekends.has(excelWeekdayIndex(holiday, dateSystem))) weekdays -= 1;
  return weekdays * direction;
}

function excelWorkday(startValue, daysValue, holidayValues = [], dateSystem = "1900", weekendValue = 1) {
  const startNumber = excelFormulaDateNumber(startValue);
  const daysNumber = excelFormulaDateNumber(daysValue);
  if (formulaErrorCode(startNumber)) return startNumber;
  if (formulaErrorCode(daysNumber)) return daysNumber;
  let serial = Math.floor(startNumber);
  const days = Math.trunc(daysNumber);
  if (!excelDateParts(serial, dateSystem) || Math.abs(days) > excelMaxDateSerial(dateSystem)) return "#NUM!";
  const holidayResult = excelHolidaySet(holidayValues, dateSystem);
  if (holidayResult.error) return holidayResult.error;
  const weekendResult = excelWeekendDays(weekendValue);
  if (weekendResult.error) return weekendResult.error;
  const direction = days < 0 ? -1 : 1;
  let remaining = Math.abs(days);
  while (remaining > 0) {
    serial += direction;
    if (!excelDateParts(serial, dateSystem)) return "#NUM!";
    if (excelBusinessDay(serial, holidayResult.holidays, dateSystem, weekendResult.weekends)) remaining -= 1;
  }
  return serial;
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

function formulaWildcardRegex(pattern) {
  let source = "";
  const text = formulaText(pattern);
  for (let index = 0; index < text.length; index++) {
    const char = text[index];
    if (char === "~" && index + 1 < text.length) source += text[++index].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    else if (char === "*") source += ".*";
    else if (char === "?") source += ".";
    else source += char.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(`^${source}$`, "i");
}

function formulaXmatchIndex(lookup, lookupValues = [], matchMode = 0, searchMode = 1) {
  const values = lookupValues.flat ? lookupValues.flat() : lookupValues;
  if (![0, -1, 1, 2].includes(matchMode) || ![1, -1, 2, -2].includes(searchMode)) return "#VALUE!";
  const indexes = Array.from({ length: values.length }, (_, index) => index);
  if (searchMode === -1 || searchMode === -2) indexes.reverse();
  const lookupNumber = Number(lookup);
  const numericLookup = Number.isFinite(lookupNumber) && formulaText(lookup).trim() !== "";
  const exact = (value) => {
    const valueNumber = Number(value);
    if (numericLookup && Number.isFinite(valueNumber) && formulaText(value).trim() !== "") return valueNumber === lookupNumber;
    return formulaText(value).toLocaleLowerCase() === formulaText(lookup).toLocaleLowerCase();
  };
  const wildcard = matchMode === 2 ? formulaWildcardRegex(lookup) : undefined;
  const found = indexes.find((index) => wildcard ? wildcard.test(formulaText(values[index])) : exact(values[index]));
  if (found != null) return found + 1;
  if (matchMode === 0 || matchMode === 2) return "#N/A";
  const comparable = indexes.map((index) => ({ index, value: values[index], number: Number(values[index]) }))
    .filter((item) => numericLookup ? Number.isFinite(item.number) : true)
    .filter((item) => matchMode < 0
      ? (numericLookup ? item.number < lookupNumber : formulaText(item.value).localeCompare(formulaText(lookup), undefined, { sensitivity: "base" }) < 0)
      : (numericLookup ? item.number > lookupNumber : formulaText(item.value).localeCompare(formulaText(lookup), undefined, { sensitivity: "base" }) > 0));
  if (!comparable.length) return "#N/A";
  comparable.sort((left, right) => {
    const delta = numericLookup ? left.number - right.number : formulaText(left.value).localeCompare(formulaText(right.value), undefined, { sensitivity: "base" });
    return matchMode < 0 ? -delta : delta;
  });
  return comparable[0].index + 1;
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
  const dateSystem = excelFormulaDateSystem(sheet);
  const scalar = (index, fallback = undefined) => {
    const value = formulaScalar(sheet, args[index], context);
    return value === undefined ? fallback : value;
  };
  const criteriaRange = (part) => {
    const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, part, context) || [[formulaScalar(sheet, part, context)]]);
    const columns = Math.max(0, ...matrix.map((row) => row.length));
    return { values: matrix.flat(), rows: matrix.length, columns, rectangular: matrix.every((row) => row.length === columns) };
  };
  const sameCriteriaShape = (left, right) => left.rectangular && right.rectangular && left.rows === right.rows && left.columns === right.columns;
  switch (fnName) {
    case "SUM":
    case "AVERAGE":
    case "MIN":
    case "MAX":
    case "COUNT":
      return aggregateFormulaValues(values(), fnName);
    case "ABS": return Math.abs(formulaNumber(scalar(0, 0)));
    case "ROUND": return roundFormulaNumber(scalar(0, 0), scalar(1, 0));
    case "ROUNDUP": return roundFormulaNumber(scalar(0, 0), scalar(1, 0), "up");
    case "ROUNDDOWN": return roundFormulaNumber(scalar(0, 0), scalar(1, 0), "down");
    case "MEDIAN": {
      const stats = statisticalFormulaNumbers(values());
      if (stats.error) return stats.error;
      if (!stats.numbers.length) return "#NUM!";
      stats.numbers.sort((left, right) => left - right);
      const middle = Math.floor(stats.numbers.length / 2);
      return stats.numbers.length % 2 ? stats.numbers[middle] : (stats.numbers[middle - 1] + stats.numbers[middle]) / 2;
    }
    case "LARGE":
    case "SMALL": {
      const stats = statisticalFormulaNumbers(values([args[0]]));
      if (stats.error) return stats.error;
      const rank = Math.trunc(formulaNumber(scalar(1, 0)));
      if (rank < 1 || rank > stats.numbers.length) return "#NUM!";
      stats.numbers.sort((left, right) => fnName === "LARGE" ? right - left : left - right);
      return stats.numbers[rank - 1];
    }
    case "RANK":
    case "RANK.EQ": {
      const number = formulaNumber(scalar(0, 0));
      if (formulaErrorCode(number)) return number;
      const stats = statisticalFormulaNumbers(values([args[1]]));
      if (stats.error) return stats.error;
      if (!stats.numbers.includes(number)) return "#N/A";
      const ascending = formulaTruthy(scalar(2, false));
      return 1 + stats.numbers.filter((value) => ascending ? value < number : value > number).length;
    }
    case "MODE":
    case "MODE.SNGL": {
      const stats = statisticalFormulaNumbers(values());
      if (stats.error) return stats.error;
      const counts = new Map();
      for (const number of stats.numbers) counts.set(number, (counts.get(number) || 0) + 1);
      const count = Math.max(0, ...counts.values());
      if (count <= 1) return "#N/A";
      return Math.min(...[...counts].filter(([, occurrences]) => occurrences === count).map(([number]) => number));
    }
    case "INT": return Math.floor(formulaNumber(scalar(0, 0)));
    case "CEILING": return Math.ceil(formulaNumber(scalar(0, 0)) / Math.max(1, formulaNumber(scalar(1, 1)))) * Math.max(1, formulaNumber(scalar(1, 1)));
    case "FLOOR": return Math.floor(formulaNumber(scalar(0, 0)) / Math.max(1, formulaNumber(scalar(1, 1)))) * Math.max(1, formulaNumber(scalar(1, 1)));
    case "DATE": {
      const parts = [0, 1, 2].map((index) => excelFormulaDateNumber(scalar(index, 0)));
      return parts.find(formulaErrorCode) || excelDateSerial(parts[0], parts[1], parts[2], dateSystem);
    }
    case "DATEVALUE": return args.length === 1 ? excelDateValue(scalar(0), dateSystem) : "#VALUE!";
    case "TIMEVALUE": return args.length === 1 ? excelTimeValue(scalar(0), dateSystem) : "#VALUE!";
    case "TIME": {
      if (args.length !== 3) return "#VALUE!";
      const parts = [0, 1, 2].map((index) => excelFormulaDateNumber(scalar(index)));
      const error = parts.find(formulaErrorCode);
      if (error) return error;
      return formulaTimeSerial(...parts) ?? "#NUM!";
    }
    case "YEAR":
    case "MONTH":
    case "DAY": {
      const serial = excelFormulaDateNumber(scalar(0, 0));
      if (formulaErrorCode(serial)) return serial;
      const parts = excelDateParts(serial, dateSystem);
      return parts ? parts[fnName.toLowerCase()] : "#NUM!";
    }
    case "HOUR": return args.length === 1 ? excelTimePart(scalar(0), "hour", dateSystem) : "#VALUE!";
    case "MINUTE": return args.length === 1 ? excelTimePart(scalar(0), "minute", dateSystem) : "#VALUE!";
    case "SECOND": return args.length === 1 ? excelTimePart(scalar(0), "second", dateSystem) : "#VALUE!";
    case "EDATE": return excelShiftMonth(scalar(0, 0), scalar(1, 0), false, dateSystem);
    case "EOMONTH": return excelShiftMonth(scalar(0, 0), scalar(1, 0), true, dateSystem);
    case "DAYS": {
      const end = excelFormulaDateNumber(scalar(0, 0));
      const start = excelFormulaDateNumber(scalar(1, 0));
      if (formulaErrorCode(end)) return end;
      if (formulaErrorCode(start)) return start;
      return excelDateParts(end, dateSystem) && excelDateParts(start, dateSystem) ? Math.floor(end) - Math.floor(start) : "#NUM!";
    }
    case "WEEKDAY": {
      const serial = excelFormulaDateNumber(scalar(0, 0));
      const returnTypeValue = excelFormulaDateNumber(scalar(1, 1));
      if (formulaErrorCode(serial)) return serial;
      if (formulaErrorCode(returnTypeValue)) return returnTypeValue;
      const returnType = Math.trunc(returnTypeValue);
      if (!excelDateParts(serial, dateSystem)) return "#NUM!";
      const weekday = excelWeekdayIndex(serial, dateSystem);
      if (returnType === 1) return weekday + 1;
      if (returnType === 2 || returnType === 11) return (weekday + 6) % 7 + 1;
      if (returnType === 3) return (weekday + 6) % 7;
      if (returnType >= 12 && returnType <= 17) return (weekday - (returnType - 10) + 7) % 7 + 1;
      return "#NUM!";
    }
    case "NETWORKDAYS": return excelNetworkDays(scalar(0, 0), scalar(1, 0), args[2] == null ? [] : values([args[2]]), dateSystem);
    case "WORKDAY": return excelWorkday(scalar(0, 0), scalar(1, 0), args[2] == null ? [] : values([args[2]]), dateSystem);
    case "NETWORKDAYS.INTL": return excelNetworkDays(scalar(0, 0), scalar(1, 0), args[3] == null ? [] : values([args[3]]), dateSystem, scalar(2, 1), true);
    case "WORKDAY.INTL": return excelWorkday(scalar(0, 0), scalar(1, 0), args[3] == null ? [] : values([args[3]]), dateSystem, scalar(2, 1));
    case "IF": return evaluateFormulaCondition(sheet, args[0], context) ? scalar(1, true) : scalar(2, false);
    case "IFERROR": { const value = scalar(0); return formulaErrorCode(value) ? scalar(1, "") : value; }
    case "IFNA": { const value = scalar(0); return formulaErrorCode(value) === "#N/A" ? scalar(1, "") : value; }
    case "AND": return args.every((arg) => evaluateFormulaCondition(sheet, arg, context));
    case "OR": return args.some((arg) => evaluateFormulaCondition(sheet, arg, context));
    case "NOT": return !evaluateFormulaCondition(sheet, args[0], context);
    case "ISNUMBER": { const value = scalar(0); return typeof value === "number" && Number.isFinite(value); }
    case "ISTEXT": { const value = scalar(0); return typeof value === "string" && !formulaErrorCode(value); }
    case "ISBLANK": { const value = scalar(0); return value == null; }
    case "ISERROR": return Boolean(formulaErrorCode(scalar(0)));
    case "ISNA": return formulaErrorCode(scalar(0)) === "#N/A";
    case "ISERR": { const error = formulaErrorCode(scalar(0)); return Boolean(error && error !== "#N/A"); }
    case "NA": return "#N/A";
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
    case "VALUE": {
      if (args.length !== 1) return "#VALUE!";
      const value = scalar(0);
      const error = formulaErrorCode(value);
      if (error) return error;
      return parseFormulaNumberText(value) ?? "#VALUE!";
    }
    case "COUNTIF": { if (args.length < 2) return "#VALUE!"; const range = values([args[0]]); const criteria = scalar(1, ""); return range.filter((value) => matchesFormulaCriteria(value, criteria)).length; }
    case "COUNTIFS": {
      if (args.length < 2 || args.length % 2 !== 0) return "#VALUE!";
      const pairs = [];
      for (let i = 0; i < args.length; i += 2) pairs.push({ range: criteriaRange(args[i]), criteria: scalar(i + 1, "") });
      const firstRange = pairs[0]?.range;
      if (!firstRange || pairs.some((pair) => !sameCriteriaShape(pair.range, firstRange))) return "#VALUE!";
      const length = firstRange.values.length;
      let count = 0;
      for (let index = 0; index < length; index++) if (pairs.every((pair) => matchesFormulaCriteria(pair.range.values[index], pair.criteria))) count += 1;
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
    case "TAKE": {
      let matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const rows = String(args[1] ?? "").trim() === "" ? undefined : Math.trunc(formulaNumber(scalar(1, 0)));
      const columns = String(args[2] ?? "").trim() === "" ? undefined : Math.trunc(formulaNumber(scalar(2, 0)));
      if (!matrix.length || rows === 0 || columns === 0 || (rows == null && columns == null)) return "#CALC!";
      if (rows != null) matrix = rows > 0 ? matrix.slice(0, rows) : matrix.slice(Math.max(0, matrix.length + rows));
      if (columns != null) matrix = matrix.map((row) => columns > 0 ? row.slice(0, columns) : row.slice(Math.max(0, row.length + columns)));
      return matrix.length && matrix.some((row) => row.length) ? matrix : "#CALC!";
    }
    case "DROP": {
      let matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const rows = String(args[1] ?? "").trim() === "" ? undefined : Math.trunc(formulaNumber(scalar(1, 0)));
      const columns = String(args[2] ?? "").trim() === "" ? undefined : Math.trunc(formulaNumber(scalar(2, 0)));
      if (!matrix.length || rows === 0 || columns === 0 || (rows == null && columns == null)) return "#CALC!";
      if (rows != null) matrix = rows >= 0 ? matrix.slice(rows) : matrix.slice(0, Math.max(0, matrix.length + rows));
      if (columns != null) matrix = matrix.map((row) => columns >= 0 ? row.slice(columns) : row.slice(0, Math.max(0, row.length + columns)));
      return matrix.length && matrix.some((row) => row.length) ? matrix : "#CALC!";
    }
    case "CHOOSECOLS": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const width = Math.max(0, ...matrix.map((row) => row.length));
      const indexes = args.slice(1).map((_, index) => Math.trunc(formulaNumber(scalar(index + 1, 0)))).map((value) => value > 0 ? value - 1 : width + value);
      if (!matrix.length || !indexes.length || indexes.some((index) => index < 0 || index >= width)) return "#VALUE!";
      return matrix.map((row) => indexes.map((index) => row[index] ?? null));
    }
    case "CHOOSEROWS": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const indexes = args.slice(1).map((_, index) => Math.trunc(formulaNumber(scalar(index + 1, 0)))).map((value) => value > 0 ? value - 1 : matrix.length + value);
      if (!matrix.length || !indexes.length || indexes.some((index) => index < 0 || index >= matrix.length)) return "#VALUE!";
      return indexes.map((index) => [...matrix[index]]);
    }
    case "TOCOL":
    case "TOROW": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const ignore = Number(scalar(1, 0));
      if (!matrix.length || !Number.isInteger(ignore) || ignore < 0 || ignore > 3) return "#VALUE!";
      const height = matrix.length;
      const width = Math.max(0, ...matrix.map((row) => row.length));
      const scanByColumn = formulaTruthy(scalar(2, false));
      const flattened = [];
      for (let outer = 0; outer < (scanByColumn ? width : height); outer++) {
        for (let inner = 0; inner < (scanByColumn ? height : width); inner++) {
          const value = scanByColumn ? matrix[inner]?.[outer] : matrix[outer]?.[inner];
          const blank = value == null || value === "";
          const error = Boolean(formulaErrorCode(value));
          if ((ignore & 1) && blank) continue;
          if ((ignore & 2) && error) continue;
          flattened.push(blank ? 0 : value);
        }
      }
      if (!flattened.length) return "#CALC!";
      return fnName === "TOCOL" ? flattened.map((value) => [value]) : [flattened];
    }
    case "WRAPROWS":
    case "WRAPCOLS": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const width = Math.max(0, ...matrix.map((row) => row.length));
      if (!matrix.length || (matrix.length > 1 && width > 1)) return "#VALUE!";
      const count = Number(scalar(1, 0));
      if (!Number.isInteger(count)) return "#VALUE!";
      if (count < 1) return "#NUM!";
      const vector = matrix.length === 1 ? [...matrix[0]] : matrix.map((row) => row[0]);
      const pad = String(args[2] ?? "").trim() === "" ? "#N/A" : scalar(2, "#N/A");
      if (fnName === "WRAPROWS") {
        const rows = Math.ceil(vector.length / count);
        return Array.from({ length: rows }, (_, row) => Array.from({ length: count }, (_, col) => {
          const index = row * count + col;
          return index < vector.length ? vector[index] : pad;
        }));
      }
      const columns = Math.ceil(vector.length / count);
      return Array.from({ length: count }, (_, row) => Array.from({ length: columns }, (_, col) => {
        const index = col * count + row;
        return index < vector.length ? vector[index] : pad;
      }));
    }
    case "HSTACK":
    case "VSTACK": {
      if (!args.length) return "#VALUE!";
      const matrices = args.map((arg, index) => normalizeFormulaMatrix(formulaRangeMatrix(sheet, arg, context) || [[scalar(index)]]));
      const normalizedValue = (value) => value == null || value === "" ? 0 : value;
      if (fnName === "HSTACK") {
        const height = Math.max(0, ...matrices.map((matrix) => matrix.length));
        return Array.from({ length: height }, (_, row) => matrices.flatMap((matrix) => {
          const width = Math.max(0, ...matrix.map((values) => values.length));
          return Array.from({ length: width }, (_, col) => row < matrix.length ? normalizedValue(matrix[row]?.[col]) : "#N/A");
        }));
      }
      const width = Math.max(0, ...matrices.flatMap((matrix) => matrix.map((row) => row.length)));
      return matrices.flatMap((matrix) => matrix.map((row) => Array.from({ length: width }, (_, col) => col < row.length ? normalizedValue(row[col]) : "#N/A")));
    }
    case "EXPAND": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      if (!matrix.length) return "#VALUE!";
      const height = matrix.length;
      const width = Math.max(0, ...matrix.map((row) => row.length));
      const rows = String(args[1] ?? "").trim() === "" ? height : Number(scalar(1, height));
      const columns = String(args[2] ?? "").trim() === "" ? width : Number(scalar(2, width));
      if (!Number.isInteger(rows) || !Number.isInteger(columns)) return "#VALUE!";
      if (rows < height || columns < width) return "#VALUE!";
      const pad = String(args[3] ?? "").trim() === "" ? "#N/A" : scalar(3, "#N/A");
      return Array.from({ length: rows }, (_, row) => Array.from({ length: columns }, (_, col) => {
        if (row >= height || col >= width) return pad;
        const value = matrix[row]?.[col];
        return value == null || value === "" ? 0 : value;
      }));
    }
    case "SUMIF": {
      if (args.length < 2) return "#VALUE!";
      const range = values([args[0]]);
      const criteria = scalar(1, "");
      const sumRange = args[2] ? values([args[2]]) : range;
      let sum = 0;
      for (let index = 0; index < range.length; index += 1) {
        if (!matchesFormulaCriteria(range[index], criteria)) continue;
        const number = formulaNumber(sumRange[index]);
        if (formulaErrorCode(number)) return number;
        sum += number;
      }
      return sum;
    }
    case "SUMIFS": {
      if (args.length < 3 || args.length % 2 === 0) return "#VALUE!";
      const sumRange = criteriaRange(args[0]);
      const pairs = [];
      for (let i = 1; i < args.length; i += 2) pairs.push({ range: criteriaRange(args[i]), criteria: scalar(i + 1, "") });
      if (pairs.some((pair) => !sameCriteriaShape(pair.range, sumRange))) return "#VALUE!";
      let sum = 0;
      for (let index = 0; index < sumRange.values.length; index += 1) {
        if (!pairs.every((pair) => matchesFormulaCriteria(pair.range.values[index], pair.criteria))) continue;
        const number = formulaNumber(sumRange.values[index]);
        if (formulaErrorCode(number)) return number;
        sum += number;
      }
      return sum;
    }
    case "AVERAGEIF": {
      if (args.length < 2) return "#VALUE!";
      const range = values([args[0]]);
      const criteria = scalar(1, "");
      const averageRange = args[2] ? values([args[2]]) : range;
      if (averageRange.length !== range.length) return "#VALUE!";
      const matched = averageRange.filter((_, index) => matchesFormulaCriteria(range[index], criteria));
      const error = matched.map(formulaErrorCode).find(Boolean);
      if (error) return error;
      const numbers = matched.filter((value) => value !== "" && value != null && Number.isFinite(Number(value))).map(Number);
      return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : "#DIV/0!";
    }
    case "AVERAGEIFS": {
      if (args.length < 3 || args.length % 2 === 0) return "#VALUE!";
      const averageRange = criteriaRange(args[0]);
      const pairs = [];
      for (let i = 1; i < args.length; i += 2) pairs.push({ range: criteriaRange(args[i]), criteria: scalar(i + 1, "") });
      if (pairs.some((pair) => !sameCriteriaShape(pair.range, averageRange))) return "#VALUE!";
      const matched = averageRange.values.filter((_, index) => pairs.every((pair) => matchesFormulaCriteria(pair.range.values[index], pair.criteria)));
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
    case "XMATCH": {
      const lookup = scalar(0, "");
      const matchValues = formulaRangeMatrix(sheet, args[1], context) || [];
      return formulaXmatchIndex(lookup, matchValues.flat(), formulaNumber(scalar(2, 0)), formulaNumber(scalar(3, 1)));
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

function evaluateFormula(sheet, formula, address, context = {}) {
  const raw = String(formula || "").trim();
  if (!raw.startsWith("=")) return raw;
  // Excel persists post-2010 worksheet functions with compatibility prefixes.
  // They are package syntax, not part of the agent-facing formula language.
  const expr = raw.slice(1).trim().replace(/_xlfn\.(?:_xlws\.)?/gi, "");
  const evaluationContext = address && !context.formulaAddress ? { ...context, formulaAddress: address } : context;
  const functionMatch = /^([A-Z][A-Z0-9.]*)\((.*)\)$/i.exec(expr);
  if (functionMatch) {
    return evaluateFormulaFunction(sheet, functionMatch[1].toUpperCase(), splitFormulaArgs(functionMatch[2]), evaluationContext);
  }
  const directReference = formulaRefParts(expr);
  if (directReference && directReference.start === directReference.end) return formulaScalar(sheet, expr, evaluationContext);

  let replacementError;
  const structuredReferences = scanStructuredReferences(expr);
  if (structuredReferences.length === 1 && structuredReferences[0].start === 0 && structuredReferences[0].end === expr.length) return formulaScalar(sheet, structuredReferences[0].text, evaluationContext);
  let structuredSafe = expr;
  for (const reference of [...structuredReferences].reverse()) {
    const value = formulaScalar(sheet, reference.text, evaluationContext);
    const error = formulaErrorCode(value);
    if (error) replacementError = error;
    structuredSafe = `${structuredSafe.slice(0, reference.start)}${formulaNumber(value)}${structuredSafe.slice(reference.end)}`;
  }
  const safe = structuredSafe.replace(/(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]+\$?\d+)/g, (match, quotedSheet, bareSheet, refAddress) => {
    const refSheetName = quotedSheet || bareSheet || undefined;
    refAddress = refAddress.replaceAll("$", "").toUpperCase();
    const targetSheet = refSheetName ? sheet.workbook?.worksheets.getItem(refSheetName) : sheet;
    const value = evaluationContext.getValue ? evaluationContext.getValue({ sheetName: refSheetName, address: refAddress }) : (targetSheet ? targetSheet.store.get(refAddress).value : "#REF!");
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
    dateSystem: workbook.dateSystem,
    activeWorksheetName: workbook.worksheets.getActiveWorksheet().name,
    selectedWorksheetNames: workbook.worksheets.getSelectedWorksheets().map((sheet) => sheet.name),
    workbookWindows: workbook.windows.items.map((window) => ({
      activeWorksheetName: window.getActiveWorksheet().name,
      selectedWorksheetNames: window.getSelectedWorksheets().map((sheet) => sheet.name),
    })),
    calculation: workbook.calculation,
    theme: workbook.theme,
    indexedColors: workbook.indexedColors,
    comments: workbook.comments.toJSON(),
    definedNames: workbook.definedNames.toJSON(),
    sheets: workbook.worksheets.items.map((sheet) => ({
      name: sheet.name,
      visibility: sheet.visibility,
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
  if (Object.prototype.hasOwnProperty.call(metadata, "calculation")) workbook.setCalculation(metadata.calculation);
  if (metadata.theme) workbook.setTheme(metadata.theme);
  if (Array.isArray(metadata.indexedColors)) workbook.indexedColors = [...metadata.indexedColors];
  workbook.definedNames.items = [];
  for (const item of metadata.definedNames || []) workbook.definedNames.add(item);
  if (metadata.comments?.self) workbook.comments.setSelf(metadata.comments.self);
  for (const threadData of metadata.comments?.threads || []) {
    const thread = new CommentThread(workbook, threadData.target, threadData.comments?.[0]?.text || "", threadData.author, { id: threadData.id, resolved: threadData.resolved, comment: threadData.comments?.[0] });
    thread.id = threadData.id || thread.id;
    thread.comments = threadData.comments || thread.comments;
    thread.resolved = Boolean(threadData.resolved);
    workbook.comments.threads.push(thread);
  }
  for (const sheetData of metadata.sheets || []) {
    const sheet = workbook.worksheets.getItem(sheetData.name);
    if (!sheet) continue;
    if (sheetData.visibility !== undefined) sheet.visibility = sheetData.visibility;
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
      const pivot = sheet.pivotTables.add({ ...pivotData, validateSource: false });
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
  if (Array.isArray(metadata.workbookWindows) && metadata.workbookWindows.length) {
    workbook.windows._clearAdditional();
    const primary = metadata.workbookWindows[0];
    if (primary.activeWorksheetName !== undefined) workbook.windows.getItemAt(0).setActiveWorksheet(primary.activeWorksheetName);
    if (Array.isArray(primary.selectedWorksheetNames)) workbook.windows.getItemAt(0).setSelectedWorksheets(primary.selectedWorksheetNames);
    for (const source of metadata.workbookWindows.slice(1)) {
      const window = workbook.windows.add({ activeWorksheet: source.activeWorksheetName });
      if (Array.isArray(source.selectedWorksheetNames)) window.setSelectedWorksheets(source.selectedWorksheetNames);
    }
  } else {
    if (metadata.activeWorksheetName !== undefined) workbook.worksheets.setActiveWorksheet(metadata.activeWorksheetName);
    if (Array.isArray(metadata.selectedWorksheetNames)) workbook.worksheets.setSelectedWorksheets(metadata.selectedWorksheetNames);
  }
}

export class SpreadsheetFile {
  static async inspectDelimited(input, options = {}) {
    const bytes = delimitedInputBytes(input);
    const maxBytes = delimitedLimit(options.maxBytes, 10 * 1024 * 1024, "maxBytes");
    if (bytes.byteLength > maxBytes) throw new Error(`Delimited input has ${bytes.byteLength} bytes and exceeds maxBytes (${maxBytes}).`);
    const inferredDelimiter = options.delimiter ?? (input instanceof FileBlob && String(input.type).split(";")[0].trim().toLowerCase() === TSV_MIME ? "\t" : ",");
    const parsed = parseDelimitedText(decoder.decode(bytes), { ...options, delimiter: inferredDelimiter });
    const columns = parsed.rows.reduce((maximum, row) => Math.max(maximum, row.length), 0);
    const summary = { kind: "delimitedFile", bytes: bytes.byteLength, delimiter: parsed.delimiter, rows: parsed.rows.length, columns, quotedCells: parsed.quotedCells, formulaLikeCells: parsed.formulaLikeCells, hasBom: bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf };
    const previewRows = Math.max(0, Math.min(parsed.rows.length, Number(options.maxPreviewRows ?? 20) || 0));
    const records = [summary, ...parsed.rows.slice(0, previewRows).map((values, index) => ({ kind: "delimitedRow", row: index + 1, columns: values.length, values }))];
    const bounded = ndjson(records, options.maxChars ?? 16_000);
    return { ok: true, issues: [], records, summary, rows: parsed.rows, ...bounded };
  }

  static async importDelimited(input, options = {}) {
    const inspection = await this.inspectDelimited(input, { ...options, maxPreviewRows: 0 });
    const workbook = Workbook.create();
    const sheet = workbook.worksheets.add(options.sheetName || "Sheet1");
    if (inspection.rows.length) sheet.getRangeByIndexes(0, 0, inspection.rows.length, inspection.summary.columns || 1).values = inspection.rows;
    return workbook;
  }

  static async exportDelimited(workbook, options = {}) {
    const selected = workbookDelimitedRows(workbook, options);
    const maxRows = delimitedLimit(options.maxRows, 100_000, "maxRows");
    const maxColumns = delimitedLimit(options.maxColumns, 16_384, "maxColumns");
    if (selected.rows.length > maxRows) throw new Error(`Delimited output has ${selected.rows.length} rows and exceeds maxRows (${maxRows}).`);
    const columns = selected.rows.reduce((maximum, row) => Math.max(maximum, row.length), 0);
    if (columns > maxColumns) throw new Error(`Delimited output has ${columns} columns and exceeds maxColumns (${maxColumns}).`);
    const text = serializeDelimitedRows(selected.rows, options);
    const bytes = encoder.encode(text);
    const maxBytes = delimitedLimit(options.maxBytes, 10 * 1024 * 1024, "maxBytes");
    if (bytes.byteLength > maxBytes) throw new Error(`Delimited output has ${bytes.byteLength} bytes and exceeds maxBytes (${maxBytes}).`);
    const delimiter = delimitedDelimiter(options.delimiter ?? ",");
    const type = options.type || (delimiter === "\t" ? TSV_MIME : CSV_MIME);
    return new FileBlob(bytes, { type, metadata: { artifactKind: "workbook", format: delimiter === "\t" ? "tsv" : "csv", delimiter, sheetName: selected.sheet.name, range: rangeToAddress(selected.range.bounds), rows: selected.rows.length, columns, formulas: options.formulas === true } });
  }

  static async importCsv(input, options = {}) { return this.importDelimited(input, { ...options, delimiter: "," }); }
  static async exportCsv(workbook, options = {}) { return this.exportDelimited(workbook, { ...options, delimiter: ",", type: CSV_MIME }); }
  static async importTsv(input, options = {}) { return this.importDelimited(input, { ...options, delimiter: "\t" }); }
  static async exportTsv(workbook, options = {}) { return this.exportDelimited(workbook, { ...options, delimiter: "\t", type: TSV_MIME }); }

  static async inspectXlsx(blobOrBuffer, options = {}) {
    return inspectOoxmlPackage(blobOrBuffer, options, XLSX_PACKAGE_CONFIG);
  }

  static async patchXlsx(blobOrBuffer, patches = [], options = {}) {
    const patched = await patchOoxmlPackage(blobOrBuffer, patches, options, XLSX_PACKAGE_CONFIG);
    return new FileBlob(patched.bytes, { type: XLSX_MIME, metadata: { artifactKind: "workbook", patchedParts: patched.patchedParts, recipesApplied: patched.recipesApplied, contentTypesUpdated: patched.contentTypesUpdated, relationshipsUpdated: patched.relationshipsUpdated, sourceReferencesUpdated: patched.sourceReferencesUpdated, validated: patched.validated, validationIssues: patched.validationIssues } });
  }

  static async exportXlsx(workbook, options = {}) {
    const { exportXlsxWithOpenChestnut } = await import("./codecs/open-chestnut.mjs");
    return exportXlsxWithOpenChestnut(workbook, options);
  }

  static async importXlsx(blobOrBuffer, options = {}) {
    const { importXlsxWithOpenChestnut } = await import("./codecs/open-chestnut.mjs");
    return importXlsxWithOpenChestnut(blobOrBuffer, options);
  }
}

function collectWorkbookThreadParts(workbook) {
  const parts = [];
  let threadPartId = 1;
  workbook.worksheets.items.forEach((sheet, sheetIndex) => {
    const activeSheetName = workbook.worksheets.getActiveWorksheet().name;
    const threads = workbook.comments.threads.filter((thread) => (thread.target.sheetName || activeSheetName) === sheet.name);
    if (threads.length) parts.push({ sheet, sheetIndex, threads, threadPartId: threadPartId++, relId: undefined });
  });
  return parts;
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
  constructor(presentation, config = {}, base = {}) {
    const normalized = normalizePresentationThemeConfig(config, base);
    this.presentation = presentation;
    this.id = config.id || "theme/default";
    this.name = normalized.name;
    this.colors = normalized.colors;
    this.fonts = normalized.fonts;
    this.textStyles = normalized.textStyles;
    this.colorMap = normalized.colorMap;
  }

  update(config = {}) {
    const normalized = normalizePresentationThemeConfig(config, this);
    Object.assign(this, normalized);
    return this;
  }

  setColors(colors = {}) { return this.update({ colors }); }
  setFonts(fonts = {}) { return this.update({ fonts }); }
  setTextStyles(textStyles = {}) { return this.update({ textStyles }); }
  setColorMap(colorMap = {}) { return this.update({ colorMap }); }
  inspectRecord() { return { kind: "theme", id: this.id, name: this.name, colors: this.colors, fonts: this.fonts, textStyles: this.textStyles, colorMap: this.colorMap }; }
  toJSON() { return { id: this.id, name: this.name, colors: this.colors, fonts: this.fonts, textStyles: this.textStyles, colorMap: this.colorMap }; }
}

function presentationThemeSemantics(theme) {
  const normalized = normalizePresentationThemeConfig(theme);
  return JSON.stringify({ name: normalized.name, colors: normalized.colors, fonts: normalized.fonts, textStyles: normalized.textStyles, colorMap: normalized.colorMap });
}

function normalizePresentationPlaceholderTransform(value, name = "Presentation placeholder transform") {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${name} must be an object.`);
  const output = {};
  if (Object.hasOwn(value, "rotationDegrees") && value.rotationDegrees != null) {
    const degrees = Number(value.rotationDegrees);
    if (!Number.isFinite(degrees) || degrees < -360 || degrees > 360) throw new RangeError(`${name}.rotationDegrees must be between -360 and 360 degrees.`);
    output.rotationDegrees = degrees;
  }
  for (const key of ["flipHorizontal", "flipVertical"]) {
    if (!Object.hasOwn(value, key) || value[key] == null) continue;
    if (typeof value[key] !== "boolean") throw new TypeError(`${name}.${key} must be a boolean.`);
    output[key] = value[key];
  }
  if (Object.keys(output).length === 0) throw new TypeError(`${name} must define rotationDegrees, flipHorizontal, or flipVertical.`);
  return output;
}

function normalizePresentationPlaceholders(value = [], idPrefix = "placeholder", options = {}) {
  if (!Array.isArray(value)) throw new TypeError("Presentation placeholders must be an array.");
  if (value.length > 128) throw new RangeError("Presentation placeholders exceed 128 entries.");
  const placeholders = value.map((placeholder, index) => {
    const position = options.allowMissingPosition && !placeholder.position && !placeholder.frame && !["left", "top", "width", "height"].some((key) => placeholder[key] != null)
      ? undefined
      : normalizeFrame(placeholder, { left: 80, top: 80 + index * 80, width: 640, height: 64 });
    const transform = normalizePresentationPlaceholderTransform(placeholder.transform, `Presentation placeholder ${placeholder.name || index + 1} transform`);
    if (transform && !position) throw new TypeError(`Presentation placeholder ${placeholder.name || index + 1} cannot define a transform without a direct position.`);
    return {
      id: placeholder.id || `${idPrefix}/${index + 1}`,
      type: placeholder.type || "body",
      idx: Number(placeholder.idx ?? index + 1),
      name: placeholder.name || `${placeholder.type || "body"} placeholder`,
      position,
      transform,
      text: placeholder.text ?? "",
      required: Boolean(placeholder.required),
      style: { ...(placeholder.style || {}) },
      paragraphStyles: normalizePresentationParagraphStyles(placeholder.paragraphStyles || placeholder.listStyles || {}),
      textBodyProperties: normalizePresentationTextBodyProperties(placeholder.textBodyProperties || placeholder.bodyProperties || {}),
    };
  });
  if (placeholders.some((placeholder) => !Number.isInteger(placeholder.idx) || placeholder.idx < 0 || placeholder.idx > 4_294_967_295)) throw new RangeError("Presentation placeholder idx must be an unsigned 32-bit integer.");
  if (new Set(placeholders.map((placeholder) => `${placeholder.type}:${placeholder.idx}`)).size !== placeholders.length) throw new Error("Presentation placeholder type/idx pairs must be unique.");
  return placeholders;
}

function clonePresentationParagraphStyles(styles = {}) {
  return Object.fromEntries(Object.entries(styles).map(([level, style]) => [Number(level), { ...style, style: { ...(style.style || {}) } }]));
}

function mergePresentationParagraphStyles(base = {}, overrides = {}) {
  const result = clonePresentationParagraphStyles(base);
  for (const [level, style] of Object.entries(overrides || {})) {
    const inherited = { ...(result[Number(level)] || {}) };
    if (["bulletCharacter", "bulletImage", "autoNumber", "bulletNone"].some((field) => Object.hasOwn(style, field))) {
      delete inherited.bulletCharacter;
      delete inherited.bulletImage;
      delete inherited.autoNumber;
      delete inherited.bulletNone;
    }
    for (const fields of [["bulletFont", "bulletFontFollowText"], ["bulletColor", "bulletColorFollowText"], ["bulletSize", "bulletSizePercent", "bulletSizeFollowText"]]) {
      if (!fields.some((field) => Object.hasOwn(style, field))) continue;
      for (const field of fields) delete inherited[field];
    }
    result[Number(level)] = { ...inherited, ...style, style: { ...(inherited.style || {}), ...(style.style || {}) } };
  }
  return result;
}

function normalizePresentationMasterParagraphStyles(value = {}) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation master textParagraphStyles must be an object.");
  return Object.fromEntries(["title", "body", "other"].map((kind) => [kind, normalizePresentationParagraphStyles(value[kind] || {})]));
}

function presentationPlaceholderTextStyleKind(type = "body") {
  if (["title", "ctrTitle"].includes(type)) return "title";
  if (["body", "subTitle", "obj", "chart", "tbl", "clipArt", "dgm", "media", "pic"].includes(type)) return "body";
  return "other";
}

class PresentationSlideMaster {
  constructor(presentation, config = {}) {
    this.presentation = presentation;
    this.configured = Object.keys(config).length > 0;
    this.id = config.id || "master/default";
    this.name = config.name || "Default Master";
    this.theme = config.theme ? new PresentationTheme(presentation, { ...config.theme, id: config.theme.id || `${this.id}/theme` }, presentation.theme) : undefined;
    Object.defineProperty(this, "_backgroundClearRequested", { value: false, writable: true });
    this.background = Object.hasOwn(config, "background")
      ? normalizePresentationBackground(config.background)
      : normalizePresentationBackground(presentation.theme.colors.bg1);
    this.placeholders = normalizePresentationPlaceholders(config.placeholders || [], `${this.id}/ph`);
    this.textParagraphStyles = normalizePresentationMasterParagraphStyles(config.textParagraphStyles || {});
  }

  update(config = {}) {
    if (Object.keys(config).length > 0) this.configured = true;
    const previousId = this.id;
    if (config.id) this.id = String(config.id);
    if (this.theme?.id === `${previousId}/theme`) this.theme.id = `${this.id}/theme`;
    if (config.name) this.name = String(config.name);
    if (Object.hasOwn(config, "theme")) this.theme = config.theme ? new PresentationTheme(this.presentation, { ...config.theme, id: config.theme.id || `${this.id}/theme` }, this.presentation.theme) : undefined;
    if (Object.hasOwn(config, "background")) {
      this.background = config.background == null ? undefined : normalizePresentationBackground(config.background, this.background);
      this._backgroundClearRequested = false;
    }
    if (config.placeholders) this.placeholders = normalizePresentationPlaceholders(config.placeholders, `${this.id}/ph`);
    if (config.textParagraphStyles) this.textParagraphStyles = normalizePresentationMasterParagraphStyles(config.textParagraphStyles);
    return this;
  }

  setBackground(background) { this.configured = true; this.background = normalizePresentationBackground(background, this.background); this._backgroundClearRequested = false; return this; }
  clearBackground() { this.configured = true; this.background = undefined; this._backgroundClearRequested = true; return this; }
  setTheme(theme) { this.configured = true; this.theme = theme ? new PresentationTheme(this.presentation, { ...theme, id: theme.id || `${this.id}/theme` }, this.presentation.theme) : undefined; return this; }
  effectiveTheme() { return this.theme || this.presentation.theme; }
  effectiveBackground() { return this.background || normalizePresentationBackground(this.effectiveTheme().colors.bg1, "#ffffff"); }
  paragraphStylesForPlaceholder(type) { return this.textParagraphStyles[presentationPlaceholderTextStyleKind(type)] || {}; }
  inspectRecord() { const theme = this.effectiveTheme(); return { kind: "slideMaster", id: this.id, name: this.name, background: this.background, effectiveBackground: this.effectiveBackground(), placeholders: this.placeholders.length, placeholderTypes: this.placeholders.map((placeholder) => placeholder.type), textParagraphStyleLevels: Object.fromEntries(Object.entries(this.textParagraphStyles).map(([kind, styles]) => [kind, Object.keys(styles).length])), hasThemeOverride: Boolean(this.theme), themeId: theme.id, themeName: theme.name }; }
  toJSON() { return { id: this.id, name: this.name, background: this.background, theme: this.theme?.toJSON(), placeholders: this.placeholders.map((placeholder) => ({ ...placeholder })), textParagraphStyles: normalizePresentationMasterParagraphStyles(this.textParagraphStyles) }; }
}

class PresentationSlideMasterCollection {
  constructor(presentation) { this.presentation = presentation; this.items = []; }
  add(config = {}) {
    if (this.items.length >= 64) throw new RangeError("Presentation masters exceed 64 entries.");
    const master = config instanceof PresentationSlideMaster ? config : new PresentationSlideMaster(this.presentation, config);
    if (this.items.some((item) => item.id === master.id)) throw new Error(`Duplicate presentation master ID ${master.id}.`);
    master.presentation = this.presentation;
    if (master.theme) master.theme.presentation = this.presentation;
    this.items.push(master);
    return master;
  }
  getItem(idOrName) { return this.items.find((master) => master.id === idOrName || master.name === idOrName); }
  get count() { return this.items.length; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class SlideLayoutTemplate {
  constructor(presentation, config = {}) {
    this.presentation = presentation;
    this.id = config.id || aid("lo");
    this.name = config.name || "Blank";
    this.type = config.type || "blank";
    this.masterId = config.masterId || presentation.master.id;
    Object.defineProperty(this, "_backgroundClearRequested", { value: false, writable: true });
    this.background = config.background ? normalizePresentationBackground(config.background) : undefined;
    this.placeholders = normalizePresentationPlaceholders(config.placeholders || [], `${this.id}/ph`, { allowMissingPosition: true });
  }

  effectiveMaster() { return this.presentation.masters.getItem(this.masterId); }
  effectiveTheme() { return this.effectiveMaster()?.effectiveTheme() || this.presentation.theme; }
  setBackground(background) { this.background = normalizePresentationBackground(background, this.background); this._backgroundClearRequested = false; return this; }
  clearBackground() { this.background = undefined; this._backgroundClearRequested = true; return this; }
  effectivePlaceholders() {
    const master = this.effectiveMaster();
    return mergePresentationPlaceholders(master?.placeholders || [], this.placeholders).map((placeholder) => ({
      ...placeholder,
      paragraphStyles: mergePresentationParagraphStyles(master?.paragraphStylesForPlaceholder(placeholder.type), placeholder.paragraphStyles),
    }));
  }
  effectiveBackground() { return this.background || this.effectiveMaster()?.effectiveBackground() || normalizePresentationBackground(this.presentation.theme.colors.bg1, "#ffffff"); }

  apply(slide) {
    slide.layoutId = this.id;
    const placeholders = this.effectivePlaceholders();
    return placeholders.map((placeholder) => {
      const shape = slide.shapes.add({
        id: placeholder.id,
        name: placeholder.name,
        geometry: "rect",
        position: placeholder.position,
        transform: placeholder.transform,
        fill: "transparent",
        line: { fill: "transparent", width: 0 },
        text: placeholder.text,
        textBodyProperties: placeholder.textBodyProperties,
        placeholder: { layoutId: this.id, type: placeholder.type, name: placeholder.name, required: placeholder.required, idx: placeholder.idx },
      });
      shape.text.style = { ...placeholder.style };
      shape.text.inheritedParagraphStyles = Object.fromEntries(Object.entries(placeholder.paragraphStyles || {}).map(([level, style]) => [level, { ...style, style: { ...(style.style || {}) } }]));
      return shape;
    });
  }

  inspectRecord() { return { kind: "layoutTemplate", id: this.id, name: this.name, type: this.type, masterId: this.masterId, themeId: this.effectiveTheme().id, background: this.background, effectiveBackground: this.effectiveBackground(), placeholders: this.placeholders.length, effectivePlaceholders: this.effectivePlaceholders().length, placeholderTypes: this.effectivePlaceholders().map((placeholder) => placeholder.type) }; }
  toJSON() { return { id: this.id, name: this.name, type: this.type, masterId: this.masterId, background: this.background, placeholders: this.placeholders.map((placeholder) => ({ ...placeholder })) }; }
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
    this.commentFormat = options.commentFormat || "legacy";
    this.theme = new PresentationTheme(this, options.theme || {});
    this.masters = new PresentationSlideMasterCollection(this);
    const masterConfigs = Array.isArray(options.masters) && options.masters.length ? options.masters : [options.master || {}];
    for (const master of masterConfigs) this.masters.add(master);
    this.layouts = new SlideLayoutCollection(this);
    for (const layout of options.layouts || []) this.layouts.add(layout);
    this.slides = new SlideCollection(this);
    this.customShows = new PresentationCustomShowCollection(this);
  }

  static create(options = {}) { return new Presentation(options); }
  get master() { return this.masters.items[0]; }
  set master(value) {
    const master = value instanceof PresentationSlideMaster ? value : new PresentationSlideMaster(this, value || {});
    master.presentation = this;
    if (master.theme) master.theme.presentation = this;
    if (this.masters.items.length) this.masters.items[0] = master;
    else this.masters.items.push(master);
  }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["deck", "slide", "textbox", "shape", "nativeObject", "layout"]);
    const records = [];
    if (kinds.has("deck")) records.push({ kind: "deck", id: this.id, slides: this.slides.count, customShows: this.customShows.count });
    if (kinds.has("theme")) records.push(this.theme.inspectRecord());
    if (kinds.has("slideMaster") || kinds.has("master")) records.push(...this.masters.items.map((master) => master.inspectRecord()));
    if (kinds.has("layout") || kinds.has("layoutTemplate")) records.push(...this.layouts.inspectRecords());
    if (kinds.has("customShow")) records.push(...this.customShows.items.map((show) => show.inspectRecord()));
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
    try { planPresentationCustomShows(this); }
    catch (error) { issues.push(verificationIssue("presentation", "invalidCustomShow", error.message)); }
    if (this.commentFormat === "modern" || this.slides.items.some((slide) => slide.comments.items.some((thread) => thread.nativeFormat === "modern"))) {
      try { planPresentationModernComments(this.slides.items); }
      catch (error) { issues.push(verificationIssue("presentation", "invalidModernCommentMetadata", error.message)); }
    }
    const duplicateMasterIds = this.masters.items.map((master) => master.id).filter((id, index, ids) => ids.indexOf(id) !== index);
    for (const masterId of new Set(duplicateMasterIds)) issues.push(verificationIssue("presentation", "duplicateMasterId", `Presentation contains duplicate master ID ${masterId}.`, { masterId }));
    const knownMasterIds = new Set(this.masters.items.map((master) => master.id));
    for (const layout of this.layouts.items) if (!knownMasterIds.has(layout.masterId)) issues.push(verificationIssue("presentation", "missingMaster", `Layout ${layout.name || layout.id} references missing master ${layout.masterId}.`, { id: layout.id, masterId: layout.masterId }));
    issues.push(...this.validateLayout(options).issues.map((issue) => ({ ...issue, artifactKind: "presentation" })));
    for (const slide of this.slides) {
      const slideElements = presentationSlideElements(slide);
      if (slide.layoutId && !this.layouts.getItem(slide.layoutId)) issues.push(verificationIssue("presentation", "missingLayout", `Slide ${slide.index + 1} references missing layout ${slide.layoutId}.`, { slide: slide.index + 1, layoutId: slide.layoutId }));
      for (const shape of slideElements.filter((element) => element instanceof Shape)) {
        if (shape.placeholder?.required && !shape.text.value.trim()) issues.push(verificationIssue("presentation", "placeholderMissingContent", `Required ${shape.placeholder.type || "placeholder"} placeholder ${shape.name || shape.id} on slide ${slide.index + 1} is empty.`, { slide: slide.index + 1, id: shape.id, placeholder: shape.placeholder }));
      }
      for (const table of slideElements.filter((element) => element instanceof TableElement)) {
        if (!table.rows || !table.columns || table.values.length === 0 || table.values.every((row) => row.every((cell) => String(cell ?? "").trim() === ""))) issues.push(verificationIssue("presentation", "emptyTable", `Table ${table.name || table.id} on slide ${slide.index + 1} has no visible cell data.`, { slide: slide.index + 1, id: table.id }));
        if (table.values.length !== table.rows) issues.push(verificationIssue("presentation", "tableDataMismatch", `Table ${table.name || table.id} declares ${table.rows} rows but has ${table.values.length} value rows.`, { slide: slide.index + 1, id: table.id, rows: table.rows, valueRows: table.values.length }));
        if (table.values.some((row) => row.length !== table.columns)) issues.push(verificationIssue("presentation", "raggedTableRows", `Table ${table.name || table.id} has rows that do not match its declared column count.`, { slide: slide.index + 1, id: table.id, columns: table.columns, rowLengths: table.values.map((row) => row.length) }));
      }
      for (const chart of slideElements.filter((element) => element instanceof ChartElement)) {
        if (!/^(bar|line|pie|combo)$/i.test(chart.chartType)) issues.push(verificationIssue("presentation", "unsupportedChartType", `Chart ${chart.name || chart.id} uses unsupported chart type ${chart.chartType}.`, { severity: "warning", slide: slide.index + 1, id: chart.id, chartType: chart.chartType }));
        if (!chart.series.length) issues.push(verificationIssue("presentation", "emptyChart", `Chart ${chart.name || chart.id} on slide ${slide.index + 1} has no data series.`, { slide: slide.index + 1, id: chart.id }));
        for (const series of chart.series) {
          const values = Array.isArray(series.values) ? series.values : [];
          if (chart.categories.length && values.length && chart.categories.length !== values.length) issues.push(verificationIssue("presentation", "chartDataMismatch", `Chart ${chart.name || chart.id} series ${series.name || "Series"} has ${values.length} values for ${chart.categories.length} categories.`, { slide: slide.index + 1, id: chart.id, series: series.name, values: values.length, categories: chart.categories.length }));
          if (values.some((value) => value !== "" && value != null && !Number.isFinite(Number(value)))) issues.push(verificationIssue("presentation", "chartDataNonNumeric", `Chart ${chart.name || chart.id} series ${series.name || "Series"} contains non-numeric values.`, { slide: slide.index + 1, id: chart.id, series: series.name }));
        }
      }
      for (const image of slideElements.filter((element) => element instanceof ImageElement)) {
        if (!image.dataUrl && !image.uri && !image.prompt) issues.push(verificationIssue("presentation", "emptyImage", `Image ${image.name || image.id} on slide ${slide.index + 1} has no dataUrl, uri, or prompt.`, { slide: slide.index + 1, id: image.id }));
        if (image.dataUrl && !imageDataFromDataUrl(image.dataUrl)) issues.push(verificationIssue("presentation", "invalidImageDataUrl", `Image ${image.name || image.id} on slide ${slide.index + 1} has an unsupported data URL.`, { slide: slide.index + 1, id: image.id }));
      }
      for (const object of slideElements.filter((element) => element instanceof NativePresentationObject)) {
        if (!object.rawXml) issues.push(verificationIssue("presentation", "nativeObjectMarkupMissing", `Native ${object.nativeKind} object ${object.name || object.id} on slide ${slide.index + 1} has no preserved markup.`, { slide: slide.index + 1, id: object.id, nativeKind: object.nativeKind }));
        const partPaths = new Set(object.parts.map((part) => part.path));
        const sourcePart = object.sourcePart || `ppt/slides/slide${slide.index + 1}.xml`;
        for (const relationship of object.rootRelationships) {
          if (relationship.targetMode?.toLowerCase() === "external") continue;
          const target = ooxmlSafePartPath(ooxmlResolveRelationshipTarget(sourcePart, relationship.target), "PPTX");
          if (!partPaths.has(target)) issues.push(verificationIssue("presentation", "nativeObjectPartMissing", `Native ${object.nativeKind} object ${object.name || object.id} is missing relationship target ${target}.`, { slide: slide.index + 1, id: object.id, relationshipId: relationship.id, target }));
        }
        for (const part of object.parts) for (const relationship of part.relationships || []) {
          if (relationship.targetMode?.toLowerCase() === "external") continue;
          const target = ooxmlSafePartPath(ooxmlResolveRelationshipTarget(part.path, relationship.target), "PPTX");
          if (!partPaths.has(target)) issues.push(verificationIssue("presentation", "nativeObjectPartMissing", `Native ${object.nativeKind} object ${object.name || object.id} is missing recursive relationship target ${target}.`, { slide: slide.index + 1, id: object.id, sourcePart: part.path, relationshipId: relationship.id, target }));
        }
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
    const master = this.masters.getItem(id);
    if (master) return master;
    const layout = this.layouts.getItem(id);
    if (layout) return layout;
    const customShow = this.customShows.getItem(id);
    if (customShow) return customShow;
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
    return { id: this.id, slideSize: this.slideSize, theme: this.theme.toJSON(), master: this.master.toJSON(), masters: this.masters.items.map((master) => master.toJSON()), layouts: this.layouts.items.map((layout) => layout.toJSON()), slides: this.slides.items.map((slide) => slide.toProto()) };
  }
}

class ShapeCollection {
  constructor(slide, owner) { this.slide = slide; this.owner = owner; this.items = []; }
  add(config = {}) { const shape = new Shape(this.slide, config); shape.parentGroup = this.owner; this.items.push(shape); this.owner?._rememberChild?.(shape); return shape; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class ElementCollection {
  constructor(slide, ElementClass, owner) { this.slide = slide; this.ElementClass = ElementClass; this.owner = owner; this.items = []; }
  add(...args) { const element = new this.ElementClass(this.slide, ...args); element.parentGroup = this.owner; this.items.push(element); this.owner?._rememberChild?.(element); return element; }
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
  const paragraphs = typeof element.text.effectiveParagraphs === "function" ? element.text.effectiveParagraphs() : normalizePresentationParagraphs(text);
  const requiredHeight = paragraphs.reduce((height, paragraph) => {
    const paragraphFontSize = Math.max(element.text.style.fontSize || 24, ...paragraph.runs.map((run) => run.style?.fontSize || 0));
    const availableWidth = Math.max(1, frame.width - 18 - Math.max(0, paragraph.marginLeft || paragraph.level * 24));
    const charsPerLine = Math.max(1, Math.floor(availableWidth / (paragraphFontSize * 0.55)));
    const requiredLines = presentationParagraphsText([paragraph]).split("\n").reduce((lines, line) => lines + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);
    const spacing = paragraph.lineSpacing || element.text.style.lineSpacing || 1.2;
    const lineHeight = spacing > 10 ? spacing : paragraphFontSize * spacing;
    return height + (paragraph.spaceBefore ?? paragraphFontSize * (paragraph.spaceBeforePercent || 0)) + requiredLines * lineHeight + (paragraph.spaceAfter ?? paragraphFontSize * (paragraph.spaceAfterPercent || 0));
  }, 12);
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

function tableOverflowIssues(slide, tableElement, frame = tableElement.position) {
  const issues = [];
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
    this.nativeFormat = config.nativeFormat;
    this.nativeAnchor = config.nativeAnchor;
    this.position = config.position;
    this.comments = (config.comments || [{ author: this.author, text: String(text ?? ""), created: this.created }]).map((comment) => ({ ...comment, author: comment.author || this.author, text: String(comment.text ?? ""), created: comment.created || this.created }));
  }

  addReply(text, config = {}) {
    this.comments.push({ ...config, author: config.author || this.author, text: String(text ?? ""), created: config.created || new Date(0).toISOString() });
    return this;
  }

  resolve() { this.resolved = true; return this; }
  reopen() { this.resolved = false; return this; }

  inspectRecord() {
    return { kind: "comment", id: this.id, slide: this.slide.index + 1, targetId: this.targetId, author: this.author, resolved: this.resolved, nativeFormat: this.nativeFormat, nativeAnchor: this.nativeAnchor, nativeCommentIds: this.comments.map((comment) => comment.nativeId).filter(Boolean), replies: Math.max(0, this.comments.length - 1), textPreview: this.comments.map((comment) => comment.text).join("\n").slice(0, 300) };
  }

  toJSON() { return { id: this.id, targetId: this.targetId, author: this.author, resolved: this.resolved, created: this.created, nativeFormat: this.nativeFormat, nativeAnchor: this.nativeAnchor, position: this.position, comments: this.comments.map((comment) => ({ ...comment })) }; }
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
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
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
    return { kind: "connector", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, connectorType: this.connectorType, start: this.start, end: this.end, startTargetId: this.startTargetId, endTargetId: this.endTargetId, line: this.line };
  }

  layoutJson() { return { kind: "connector", id: this.id, name: this.name, connectorType: this.connectorType, start: this.start, end: this.end, startTargetId: this.startTargetId, endTargetId: this.endTargetId, line: this.line, frame: this.position }; }

  toSvg() {
    const stroke = resolveColorToken(this.line?.fill || this.line?.color || "#334155", "#334155");
    const width = this.line?.width ?? 2;
    const markerId = `${this.id.replace(/[^A-Za-z0-9_-]/g, "")}-arrow`;
    const marker = this.line?.endArrow ? `<defs><marker id="${markerId}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" fill="${xmlEscape(stroke)}"/></marker></defs>` : "";
    return `${marker}<line x1="${this.start.x}" y1="${this.start.y}" x2="${this.end.x}" y2="${this.end.y}" stroke="${xmlEscape(stroke)}" stroke-width="${width}" marker-end="${this.line?.endArrow ? `url(#${markerId})` : ""}"/>`;
  }

}

class NativePresentationObject {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.kind = "nativeObject";
    this.id = config.id || aid("no");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.name = config.name || "";
    this.nativeKind = config.nativeKind || "graphicFrame";
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 1, height: 1 });
    this.rawXml = String(config.rawXml || "");
    this.sourcePart = config.sourcePart;
    Object.defineProperty(this, "editable", { enumerable: true, value: false, writable: false });
    this.relationshipReferences = (config.relationshipReferences || []).map((reference) => ({ ...reference }));
    this.rootRelationships = (config.rootRelationships || []).map((relationship) => ({ ...relationship }));
    this.parts = (config.parts || []).map((part) => ({ ...part, bytes: new Uint8Array(part.bytes), relationships: (part.relationships || []).map((relationship) => ({ ...relationship })) }));
    this.oleWorkbook = config.oleWorkbook ? Object.freeze({
      partPath: String(config.oleWorkbook.partPath || ""),
      contentType: String(config.oleWorkbook.contentType || ""),
      sourceSha256: String(config.oleWorkbook.sourceSha256 || "").toLowerCase(),
      relationshipId: String(config.oleWorkbook.relationshipId || ""),
    }) : undefined;
  }

  setName(value) {
    if (!this.editable) throw new Error(`Native ${this.nativeKind} object ${this.id} is read-only.`);
    const name = String(value ?? "");
    if (name.length > 1_024) throw new RangeError("Native presentation object names cannot exceed 1024 characters.");
    this.name = name;
    return this;
  }

  setPosition(value = {}) {
    if (!this.editable) throw new Error(`Native ${this.nativeKind} object ${this.id} is read-only.`);
    this.position = normalizeFrame({ position: { ...this.position, ...value } }, this.position);
    return this;
  }

  embeddedWorkbookPart() {
    if (!this.oleWorkbook) throw new Error(`Native ${this.nativeKind} object ${this.id} has no embedded XLSX workbook.`);
    const matches = this.parts.filter((part) => part.path === this.oleWorkbook.partPath && part.contentType === this.oleWorkbook.contentType);
    if (matches.length !== 1) throw new Error(`Native ${this.nativeKind} object ${this.id} no longer resolves to one embedded XLSX workbook part.`);
    return matches[0];
  }

  getEmbeddedWorkbook() {
    const part = this.embeddedWorkbookPart();
    return new FileBlob(Uint8Array.from(part.bytes), {
      type: this.oleWorkbook.contentType,
      metadata: { artifactKind: "workbook", source: "presentationOleObject", partPath: this.oleWorkbook.partPath, sourceSha256: this.oleWorkbook.sourceSha256 },
    });
  }

  replaceEmbeddedWorkbook(_input) {
    throw new Error(`Native ${this.nativeKind} object ${this.id} is source-bound and read-only in OpenChestnut 0.2.`);
  }

  inspectRecord() {
    const frame = this.parentGroup ? this.parentGroup.absoluteChildFrame(this) : this.position;
    const editableFields = [];
    return {
      kind: "nativeObject",
      id: this.id,
      slide: this.slide.index + 1,
      name: this.name || undefined,
      nativeKind: this.nativeKind,
      nativeId: this.nativeId,
      creationId: this.creationId,
      sourcePart: this.sourcePart,
      relationships: this.rootRelationships.length,
      preservedParts: this.parts.length,
      relationshipReferences: this.relationshipReferences.map(({ attribute, id, namespaceUri }) => ({ attribute, id, namespaceUri })),
      nativeRelationships: this.rootRelationships.map(({ id, type, target, targetMode }) => ({ id, type, target, targetMode })),
      nativeParts: this.parts.map((part) => ({ path: part.path, contentType: part.contentType, relationships: part.relationships.length })),
      embeddedWorkbook: this.oleWorkbook ? { partPath: this.oleWorkbook.partPath, contentType: this.oleWorkbook.contentType, bytes: this.embeddedWorkbookPart().bytes.length, sourceSha256: this.oleWorkbook.sourceSha256 } : undefined,
      bbox: [frame.left, frame.top, frame.width, frame.height],
      bboxUnit: "px",
      editable: false,
      editableFields,
    };
  }

  layoutJson() {
    return { kind: "nativeObject", id: this.id, name: this.name, nativeKind: this.nativeKind, frame: this.position, relationships: this.rootRelationships.length, preservedParts: this.parts.length, embeddedWorkbook: this.oleWorkbook ? { partPath: this.oleWorkbook.partPath, contentType: this.oleWorkbook.contentType, bytes: this.embeddedWorkbookPart().bytes.length } : undefined, editable: false, editableFields: [] };
  }

  toSvg() {
    const p = this.position;
    if (!(p.width > 1 && p.height > 1)) return `<g data-native-object-id="${attrEscape(this.id)}" data-native-kind="${attrEscape(this.nativeKind)}"/>`;
    const label = this.name || this.nativeKind;
    return `<g data-native-object-id="${attrEscape(this.id)}" data-native-kind="${attrEscape(this.nativeKind)}"><rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#f8fafc" fill-opacity="0.72" stroke="#64748b" stroke-dasharray="6 4"/><text x="${p.left + 8}" y="${p.top + 20}" font-family="Arial" font-size="12" fill="#475569">${xmlEscape(label)}</text></g>`;
  }

}

const GroupShape = createPresentationGroupShapeClass({
  createId: aid,
  createShapeCollection: (slide, owner) => new ShapeCollection(slide, owner),
  createConnectorCollection: (slide, owner) => new ElementCollection(slide, ConnectorElement, owner),
  createGroupCollection: (slide, owner, GroupClass) => new ElementCollection(slide, GroupClass, owner),
  createTableCollection: (slide, owner) => new ElementCollection(slide, TableElement, owner),
  createChartCollection: (slide, owner) => new ElementCollection(slide, ChartElement, owner),
  createImageCollection: (slide, owner) => new ElementCollection(slide, ImageElement, owner),
  createNativeObjectCollection: (slide, owner) => new ElementCollection(slide, NativePresentationObject, owner),
  isShape: (element) => element instanceof Shape,
  isConnector: (element) => element instanceof ConnectorElement,
  isGroup: (element) => element instanceof GroupShape,
  isTable: (element) => element instanceof TableElement,
  isChart: (element) => element instanceof ChartElement,
  isImage: (element) => element instanceof ImageElement,
  isNativeObject: (element) => element instanceof NativePresentationObject,
  elementKind: (element) => presentationElementKind(element),
  validateChildLayout: (element, frame) => element instanceof TableElement ? tableOverflowIssues(element.slide, element, frame) : [],
  createTextRange: (element, id) => createTextRange(element, id, { parentKind: "shape" }),
  textRangeRecord,
  elementLabel,
});
export { GroupShape };
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
    this.groups = new ElementCollection(this, GroupShape);
    this.nativeObjects = new ElementCollection(this, NativePresentationObject);
    this.comments = new SlideCommentCollection(this);
    this.layoutId = options.layoutId || options.layout?.id || (typeof options.layout === "string" ? options.layout : undefined);
    this.speakerNotes = { text: String(options.notes || options.speakerNotes?.text || "") };
    this.background = options.background ? normalizePresentationBackground(options.background) : {};
  }

  get index() { return this.presentation.slides.items.indexOf(this); }
  get frame() { return { left: 0, top: 0, ...this.presentation.slideSize }; }

  addNotes(text) { this.speakerNotes.text = String(text ?? ""); return this.speakerNotes; }
  addComment(target, text, config = {}) { return this.comments.addThread(target, text, config); }
  addConnector(config = {}) { return this.connectors.add(config); }
  addGroup(config = {}) { return this.groups.add(config); }
  applyLayout(layoutOrName) { const layout = typeof layoutOrName === "string" ? this.presentation.layouts.getItem(layoutOrName) : layoutOrName; if (!layout) throw new Error(`Unknown slide layout: ${layoutOrName}`); return layout.apply(this); }
  effectiveBackground() { const layout = this.presentation.layouts.getItem(this.layoutId); return this.background.fill ? this.background : layout?.effectiveBackground() || this.presentation.master.effectiveBackground(); }
  effectiveTheme() { const layout = this.presentation.layouts.getItem(this.layoutId); return layout?.effectiveTheme() || this.presentation.master.effectiveTheme(); }

  inspectRecords(kinds) {
    const records = [];
    if (kinds.has("layout")) { const layout = this.presentation.layouts.getItem(this.layoutId); records.push({ kind: "layout", layoutId: this.layoutId || `${this.id}/layout`, name: layout?.name || "Blank", type: layout?.type || "blank", masterId: layout?.masterId, themeId: this.effectiveTheme().id, placeholders: layout?.placeholders.length || 0 }); }
    if (kinds.has("slide")) records.push({ kind: "slide", id: this.id, slide: this.index + 1, title: this.title(), textShapes: this.shapes.items.filter((s) => s.text.value).length, tables: this.tables.items.length, charts: this.charts.items.length, images: this.images.items.length, connectors: this.connectors.items.length, groups: this.groups.items.length, nativeObjects: this.nativeObjects.items.length, comments: this.comments.items.length, hasNotes: Boolean(this.speakerNotes.text) });
    for (const shape of this.shapes) {
      if (kinds.has("textbox") && shape.text.value) records.push(shape.inspectRecord("textbox"));
      else if (kinds.has("shape")) records.push(shape.inspectRecord("shape"));
      if (kinds.has("textRange") && shape.text.value) records.push(textRangeRecord(shape, { parentKind: "shape", record: { slide: this.index + 1, bbox: [shape.position.left, shape.position.top, shape.position.width, shape.position.height], bboxUnit: "px" } }));
    }
    if (kinds.has("table")) records.push(...this.tables.items.map((table) => table.inspectRecord()));
    if (kinds.has("chart")) records.push(...this.charts.items.map((chart) => chart.inspectRecord()));
    if (kinds.has("image")) records.push(...this.images.items.map((image) => image.inspectRecord()));
    if (kinds.has("connector")) records.push(...this.connectors.items.map((connector) => connector.inspectRecord()));
    if (kinds.has("nativeObject") || kinds.has("native")) records.push(...this.nativeObjects.items.map((object) => object.inspectRecord()));
    for (const nativeKind of ["contentPart", "oleObject", "diagram", "graphicFrame"]) if (kinds.has(nativeKind)) records.push(...this.nativeObjects.items.filter((object) => object.nativeKind === nativeKind).map((object) => object.inspectRecord()));
    for (const group of this.groups) records.push(...group.inspectRecords(kinds));
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
    const direct = [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.connectors.items, ...this.groups.items, ...this.nativeObjects.items, ...this.comments.items].find((element) => element.id === id);
    if (direct) return direct;
    for (const group of this.groups) {
      const nested = group.resolve(id);
      if (nested) return nested;
    }
    return undefined;
  }

  validateLayout(options = {}) {
    const issues = [];
    const slideFrame = this.frame;
    const elements = [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.groups.items, ...this.nativeObjects.items];
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
    for (const group of this.groups) issues.push(...group.validateLayout());
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
    const elements = [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.connectors.items, ...this.groups.items, ...this.nativeObjects.items].map((element) => {
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
      slide: { id: this.id, slide: this.index + 1, frame: this.frame, background: this.effectiveBackground(), notes: this.speakerNotes.text || undefined },
      elements,
    }, options);
  }

  toSvg() {
    const { width, height } = this.presentation.slideSize;
    const elements = [...this.connectors.items, ...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.groups.items, ...this.nativeObjects.items].map((element) => element.toSvg()).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="${xmlEscape(resolvePresentationBackgroundColor(this.effectiveBackground(), this.effectiveTheme()))}"/>${elements}</svg>`;
  }

  toProto() { return { id: this.id, layoutId: this.layoutId, background: this.background.fill ? this.background : undefined, notes: this.speakerNotes.text || undefined, comments: this.comments.items.map((comment) => comment.toJSON()), elements: [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.connectors.items, ...this.nativeObjects.items].map((element) => element.layoutJson()), groups: this.groups.items.map((group) => group.toProto()) }; }

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
  constructor(text = "", bodyProperties, { defaultBodyProperties = false } = {}) { this._paragraphs = normalizePresentationParagraphs(text); this.style = {}; this.inheritedParagraphStyles = {}; this.bodyProperties = normalizePresentationTextBodyProperties(bodyProperties, { defaults: defaultBodyProperties }); }
  get value() { return presentationParagraphsText(this._paragraphs); }
  set value(text) { this._paragraphs = normalizePresentationParagraphs(text); }
  get paragraphs() { return normalizePresentationParagraphs(this._paragraphs); }
  set paragraphs(value) { this._paragraphs = normalizePresentationParagraphs(value); }
  effectiveParagraphs() { return inheritPresentationParagraphs(this._paragraphs, this.inheritedParagraphStyles); }
  set(text) { this._paragraphs = normalizePresentationParagraphs(text); return this; }
  replace(search, replacement) { replacePresentationParagraphText(this._paragraphs, search, replacement); return this; }
  toString() { return this.value; }
}

export class Shape {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("sh");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.geometry = config.geometry || "rect";
    this.customPaths = normalizePresentationCustomPaths(config.customPaths, { geometry: this.geometry });
    this.name = config.name || "";
    this.position = config.position || { left: 0, top: 0, width: 160, height: 80 };
    this.transform = config.transform == null ? undefined : normalizePresentationPlaceholderTransform(config.transform, `Presentation shape ${this.name || this.id} transform`);
    this.fill = config.fill || "transparent";
    this.line = config.line || { fill: "#334155", width: 1 };
    this.borderRadius = config.borderRadius;
    this.shadow = config.shadow ? { ...config.shadow } : undefined;
    this.placeholder = config.placeholder;
    this._text = new TextFrame(config.text ?? "", config.textBodyProperties, { defaultBodyProperties: config.textBodyProperties === undefined });
    this._text.style = { ...(config.textStyle || config.style?.text || {}) };
  }

  get text() { return this._text; }
  set text(value) { this._text.set(value); }

  inspectRecord(kind = "shape") {
    const p = this.position;
    const paragraphs = this.text.effectiveParagraphs();
    return { kind, id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, text: this.text.value || undefined, textPreview: this.text.value || undefined, textChars: this.text.value.length || undefined, textLines: this.text.value ? this.text.value.split("\n").length : undefined, paragraphs: presentationParagraphsNeedSerialization(paragraphs) ? paragraphs : undefined, bodyProperties: this.text.bodyProperties, customPathCount: this.customPaths.length || undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", transform: this.transform, shadow: this.shadow, placeholder: this.placeholder || undefined };
  }

  layoutJson() { const paragraphs = this.text.effectiveParagraphs(); return { kind: this.text.value ? "textbox" : "shape", id: this.id, name: this.name, geometry: this.geometry, customPaths: this.customPaths.length ? this.customPaths : undefined, frame: this.position, transform: this.transform, text: this.text.value, paragraphs: presentationParagraphsNeedSerialization(paragraphs) ? paragraphs : undefined, bodyProperties: this.text.bodyProperties, placeholder: this.placeholder, style: { fill: this.fill, line: this.line, borderRadius: this.borderRadius, shadow: this.shadow, text: this.text.style } }; }

  toSvg() {
    const p = this.position;
    const fill = typeof this.fill === "string" ? resolveColorToken(this.fill, this.fill) : this.fill?.color || "transparent";
    const stroke = resolveColorToken(this.line?.fill || this.line?.color || "#334155", "#334155");
    const sw = this.line?.width ?? 1;
    const visual = this.geometry === "custom"
      ? `<g fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}">${presentationCustomPathsSvg(this.customPaths, p, { escape: xmlEscape })}</g>`
      : this.geometry === "ellipse"
      ? `<ellipse cx="${p.left + p.width / 2}" cy="${p.top + p.height / 2}" rx="${p.width / 2}" ry="${p.height / 2}" fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}"/>`
      : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" rx="${this.borderRadius ? 12 : 0}" fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}"/>`;
    const text = this.text.value ? presentationParagraphsSvg(this.text.effectiveParagraphs(), p, this.text.style, { escape: xmlEscape }) : "";
    if (!this.transform) return visual + text;
    const cx = p.left + p.width / 2;
    const cy = p.top + p.height / 2;
    const rotation = Number(this.transform.rotationDegrees || 0);
    const flipHorizontal = this.transform.flipHorizontal === true ? -1 : 1;
    const flipVertical = this.transform.flipVertical === true ? -1 : 1;
    return `<g transform="translate(${cx} ${cy}) rotate(${rotation}) scale(${flipHorizontal} ${flipVertical}) translate(${-cx} ${-cy})">${visual}${text}</g>`;
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
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
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
    return { kind: "table", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, rows: this.rows, cols: this.columns, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", values: this.values };
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

}

function normalizeChartSeries(seriesItems = [], chartType = "bar") {
  return (seriesItems || []).map((series, index) => {
    const values = (series.values || series.data || []).map((value) => value);
    const style = normalizePresentationChartSeriesStyle(series, values.length);
    const seriesChartType = chartType === "combo" ? String(series.chartType || series.type || "").toLowerCase() : undefined;
    if (chartType === "combo" && !new Set(["bar", "line"]).has(seriesChartType)) throw new TypeError("Presentation combo chart series chartType must be bar or line.");
    const rawAxisGroup = series.axisGroup ?? series.axis ?? (series.secondaryAxis === true ? "secondary" : "primary");
    const axisGroup = normalizePresentationChartAxisGroup(rawAxisGroup === "y2" ? "secondary" : rawAxisGroup === "y1" ? "primary" : String(rawAxisGroup).toLowerCase(), seriesChartType || chartType);
    return {
      name: series.name || `Series ${index + 1}`,
      values,
      categories: series.categories,
      color: style.color || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4],
      ...(style.line ? { line: style.line } : {}),
      ...(style.points.length ? { points: style.points } : {}),
      ...(style.marker ? { marker: style.marker } : {}),
      ...(style.smooth == null ? {} : { smooth: style.smooth }),
      ...(series.dataLabels === undefined ? {} : { dataLabels: normalizePresentationChartDataLabels(series.dataLabels) }),
      ...((series.trendlines ?? series.trendline) == null ? {} : { trendlines: normalizePresentationChartTrendlines(series.trendlines ?? series.trendline, values.length, seriesChartType || chartType) }),
      ...(series.errorBars == null ? {} : { errorBars: normalizePresentationChartErrorBars(series.errorBars, seriesChartType || chartType, values.length) }),
      ...(seriesChartType ? { chartType: seriesChartType } : {}),
      ...(axisGroup === "secondary" ? { axisGroup } : {}),
    };
  });
}

function normalizeChartAxes(config = {}, hasSecondary = false) {
  const axes = config.axes || {};
  const axisTitles = config.axisTitles || {};
  const secondary = axes.secondary || {};
  const secondaryAxisTitles = axisTitles.secondary || config.secondaryAxisTitles || {};
  return {
    category: { ...(axes.category || axes.x || {}), title: axes.category?.title || axes.x?.title || axisTitles.category || axisTitles.x || config.categoryAxisTitle || config.xAxisTitle || "" },
    value: { ...(axes.value || axes.y || {}), title: axes.value?.title || axes.y?.title || axisTitles.value || axisTitles.y || config.valueAxisTitle || config.yAxisTitle || "" },
    ...(hasSecondary ? {
      secondary: {
        category: { ...(secondary.category || secondary.x || axes.secondaryCategory || {}), title: secondary.category?.title || secondary.x?.title || axes.secondaryCategory?.title || secondaryAxisTitles.category || secondaryAxisTitles.x || config.secondaryCategoryAxisTitle || config.secondaryXAxisTitle || "" },
        value: { ...(secondary.value || secondary.y || axes.secondaryValue || axes.y2 || {}), title: secondary.value?.title || secondary.y?.title || axes.secondaryValue?.title || axes.y2?.title || secondaryAxisTitles.value || secondaryAxisTitles.y || config.secondaryValueAxisTitle || config.secondaryYAxisTitle || "" },
      },
    } : {}),
  };
}

function normalizeChartLegend(config = {}, seriesLength = 0) {
  const raw = config.legend;
  if (raw === false || config.hasLegend === false) return { visible: false, position: "r" };
  if (typeof raw === "string") return { visible: true, position: raw };
  return { visible: raw?.visible ?? config.hasLegend ?? seriesLength > 1, position: raw?.position || config.legendPosition || "r" };
}

function normalizeChartDataLabels(config = {}) {
  const raw = config.dataLabels ?? config.labels ?? {};
  if (raw === true || raw === false) return normalizePresentationChartDataLabels(raw);
  return normalizePresentationChartDataLabels({
    ...raw,
    showValue: raw.showValue ?? config.showValues,
    showCategoryName: raw.showCategoryName ?? raw.showCategory ?? config.showCategoryLabels,
  });
}

function pieSlicePath(cx, cy, radius, startAngle, endAngle) {
  const startX = cx + radius * Math.cos(startAngle);
  const startY = cy + radius * Math.sin(startAngle);
  const endX = cx + radius * Math.cos(endAngle);
  const endY = cy + radius * Math.sin(endAngle);
  const largeArc = endAngle - startAngle > Math.PI ? 1 : 0;
  return `M ${cx} ${cy} L ${startX} ${startY} A ${radius} ${radius} 0 ${largeArc} 1 ${endX} ${endY} Z`;
}

function presentationChartMarkerSvg(marker, x, y, color) {
  if (!marker || marker.symbol === "none") return "";
  const size = Math.max(2, Number(marker.size) || 5);
  const radius = size / 2;
  const stroke = xmlEscape(color);
  if (marker.symbol === "square") return `<rect x="${x - radius}" y="${y - radius}" width="${size}" height="${size}" fill="${stroke}"/>`;
  if (marker.symbol === "diamond") return `<path d="M ${x} ${y - radius} L ${x + radius} ${y} L ${x} ${y + radius} L ${x - radius} ${y} Z" fill="${stroke}"/>`;
  if (marker.symbol === "triangle") return `<path d="M ${x} ${y - radius} L ${x + radius} ${y + radius} L ${x - radius} ${y + radius} Z" fill="${stroke}"/>`;
  if (marker.symbol === "x") return `<path d="M ${x - radius} ${y - radius} L ${x + radius} ${y + radius} M ${x + radius} ${y - radius} L ${x - radius} ${y + radius}" fill="none" stroke="${stroke}" stroke-width="1.5"/>`;
  if (marker.symbol === "plus") return `<path d="M ${x - radius} ${y} L ${x + radius} ${y} M ${x} ${y - radius} L ${x} ${y + radius}" fill="none" stroke="${stroke}" stroke-width="1.5"/>`;
  if (marker.symbol === "dash") return `<line x1="${x - radius}" y1="${y}" x2="${x + radius}" y2="${y}" stroke="${stroke}" stroke-width="2"/>`;
  return `<circle cx="${x}" cy="${y}" r="${marker.symbol === "dot" ? Math.max(1, radius / 2) : radius}" fill="${stroke}"/>`;
}

function presentationChartDataLabelText(dataLabels, category, value) {
  if (!dataLabels?.showValue && !dataLabels?.showCategoryName) return "";
  if (dataLabels.showValue && dataLabels.showCategoryName) return `${category}: ${value}`;
  return dataLabels.showCategoryName ? String(category ?? "") : String(value ?? "");
}

function presentationChartErrorBarsSvg(series, points, plot, max) {
  const errorBars = series.errorBars;
  if (!errorBars || !points.length) return "";
  const numericValues = (series.values || []).map(Number).filter(Number.isFinite);
  const mean = numericValues.reduce((sum, value) => sum + value, 0) / Math.max(1, numericValues.length);
  const deviation = Math.sqrt(numericValues.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(1, numericValues.length));
  const magnitudeFor = (value, index, side) => errorBars.valueType === "cust" ? Number(errorBars[`${side}Values`]?.[index]) || 0
    : errorBars.valueType === "percentage" ? Math.abs(Number(value) || 0) * (errorBars.value || 0) / 100
    : errorBars.valueType === "stdDev" ? deviation * (errorBars.value || 1)
      : errorBars.valueType === "stdErr" ? deviation / Math.sqrt(Math.max(1, numericValues.length))
        : errorBars.value || 0;
  const attributes = presentationChartLineSvgAttributes(errorBars.line || { fill: series.color || "#475569", width: 1, style: "solid" });
  return points.map((point, index) => {
    const pointIndex = point.index ?? index;
    const scale = (errorBars.direction === "x" ? plot.width : plot.height) / Math.max(1, max);
    const minus = errorBars.type !== "plus" ? magnitudeFor(series.values?.[pointIndex], pointIndex, "minus") * scale : 0;
    const plus = errorBars.type !== "minus" ? magnitudeFor(series.values?.[pointIndex], pointIndex, "plus") * scale : 0;
    const x1 = errorBars.direction === "x" ? point.x - minus : point.x;
    const x2 = errorBars.direction === "x" ? point.x + plus : point.x;
    const y1 = errorBars.direction === "y" ? point.y + minus : point.y;
    const y2 = errorBars.direction === "y" ? point.y - plus : point.y;
    const caps = errorBars.noEndCap ? "" : errorBars.direction === "x"
      ? `${minus > 0 ? `<line x1="${x1}" y1="${point.y - 4}" x2="${x1}" y2="${point.y + 4}"${attributes}/>` : ""}${plus > 0 ? `<line x1="${x2}" y1="${point.y - 4}" x2="${x2}" y2="${point.y + 4}"${attributes}/>` : ""}`
      : `${minus > 0 ? `<line x1="${point.x - 4}" y1="${y1}" x2="${point.x + 4}" y2="${y1}"${attributes}/>` : ""}${plus > 0 ? `<line x1="${point.x - 4}" y1="${y2}" x2="${point.x + 4}" y2="${y2}"${attributes}/>` : ""}`;
    return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}"${attributes}/>${caps}`;
  }).join("");
}

export class ChartElement {
  constructor(slide, chartType = "bar", config = {}) {
    this.slide = slide;
    this.id = config.id || aid("ch");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.name = config.name || "";
    this.chartType = String(chartType || config.chartType || "bar").toLowerCase();
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 360, height: 220 });
    this.title = config.title || "";
    this.categories = config.categories || [];
    this.series = normalizeChartSeries(config.series || [], this.chartType);
    this.externalData = normalizePresentationChartExternalData(config.externalData ?? config.sourceWorkbook);
    if (presentationChartUsesFormulaReferences(this) && !this.externalData) throw new TypeError("Presentation chart formula references require externalData with an embedded workbook or external workbook URI.");
    if (this.chartType === "combo" && (!this.series.some((series) => series.chartType === "bar") || !this.series.some((series) => series.chartType === "line"))) throw new TypeError("Presentation combo chart requires at least one bar series and one line series.");
    const hasSecondary = this.series.some((series) => series.axisGroup === "secondary");
    const hasConfiguredSecondaryAxes = Boolean(config.axes?.secondary || config.axes?.secondaryCategory || config.axes?.secondaryValue || config.axes?.y2 || config.secondaryAxisTitles || config.secondaryCategoryAxisTitle || config.secondaryValueAxisTitle || config.secondaryXAxisTitle || config.secondaryYAxisTitle);
    if (hasConfiguredSecondaryAxes && !hasSecondary) throw new TypeError("Presentation secondary axes require at least one chart series with axisGroup secondary.");
    if (hasSecondary && !this.series.some((series) => series.axisGroup !== "secondary")) throw new TypeError("Presentation secondary-axis charts require at least one primary-axis series.");
    this.axes = normalizeChartAxes(config, hasSecondary);
    this.legend = normalizeChartLegend(config, this.series.length);
    this.hasLegend = this.legend.visible;
    this.dataLabels = normalizeChartDataLabels(config);
    Object.assign(this, normalizePresentationChartStyle(this.chartType, config));
  }

  inspectRecord() {
    const p = this.position;
    return { kind: "chart", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, chartType: this.chartType, title: this.title, categories: this.categories, series: this.series.length, seriesDetails: this.series, axes: this.axes, legend: this.legend, dataLabels: this.dataLabels, externalData: this.externalData ? { embedded: Boolean(this.externalData.bytes), uri: this.externalData.uri, autoUpdate: this.externalData.autoUpdate, bytes: this.externalData.bytes?.byteLength } : undefined, styleId: this.styleId, varyColors: this.varyColors, barOptions: ["bar", "combo"].includes(this.chartType) ? this.barOptions : undefined, lineOptions: ["line", "combo"].includes(this.chartType) ? this.lineOptions : undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px" };
  }

  layoutJson() { return { kind: "chart", id: this.id, name: this.name, chartType: this.chartType, title: this.title, frame: this.position, categories: this.categories, series: this.series, axes: this.axes, legend: this.legend, dataLabels: this.dataLabels, externalData: this.externalData ? { embedded: Boolean(this.externalData.bytes), uri: this.externalData.uri, autoUpdate: this.externalData.autoUpdate, bytes: this.externalData.bytes?.byteLength } : undefined, styleId: this.styleId, varyColors: this.varyColors, barOptions: ["bar", "combo"].includes(this.chartType) ? this.barOptions : undefined, lineOptions: ["line", "combo"].includes(this.chartType) ? this.lineOptions : undefined }; }

  toSvg() {
    const p = this.position;
    const categories = this.categories.length ? this.categories : Array.from({ length: Math.max(0, ...this.series.map((series) => series.values?.length || 0)) }, (_, index) => String(index + 1));
    const barSeries = this.chartType === "combo" ? this.series.filter((series) => series.chartType === "bar") : this.chartType === "bar" ? this.series : [];
    const lineSeries = this.chartType === "combo" ? this.series.filter((series) => series.chartType === "line") : this.chartType === "line" ? this.series : [];
    const stackedBars = barSeries.length > 0 && this.barOptions.grouping !== "clustered";
    const stackedLines = lineSeries.length > 0 && this.lineOptions.grouping !== "standard";
    const forAxisGroup = (series, axisGroup) => series.filter((item) => (item.axisGroup || "primary") === axisGroup);
    const stackedTotals = (series) => categories.map((_, categoryIndex) => series.reduce((sum, item) => sum + Math.max(0, Number(item.values?.[categoryIndex]) || 0), 0));
    const barByAxis = { primary: forAxisGroup(barSeries, "primary"), secondary: forAxisGroup(barSeries, "secondary") };
    const lineByAxis = { primary: forAxisGroup(lineSeries, "primary"), secondary: forAxisGroup(lineSeries, "secondary") };
    const barStackedMax = { primary: stackedTotals(barByAxis.primary), secondary: stackedTotals(barByAxis.secondary) };
    const lineStackedMax = { primary: stackedTotals(lineByAxis.primary), secondary: stackedTotals(lineByAxis.secondary) };
    const groupMax = (series, stacked, stackedValues, percentStacked) => percentStacked
      ? 1
      : Math.max(0, ...(stacked ? stackedValues : series.flatMap((item) => item.values || []).map((value) => Math.max(0, Number(value) || 0))));
    const barMax = {
      primary: groupMax(barByAxis.primary, stackedBars, barStackedMax.primary, this.barOptions?.grouping === "percentStacked"),
      secondary: groupMax(barByAxis.secondary, stackedBars, barStackedMax.secondary, this.barOptions?.grouping === "percentStacked"),
    };
    const lineMax = {
      primary: groupMax(lineByAxis.primary, stackedLines, lineStackedMax.primary, this.lineOptions?.grouping === "percentStacked"),
      secondary: groupMax(lineByAxis.secondary, stackedLines, lineStackedMax.secondary, this.lineOptions?.grouping === "percentStacked"),
    };
    const maxForAxisGroup = (axisGroup) => Math.max(
      1,
      barMax[axisGroup],
      lineMax[axisGroup],
    );
    const primaryMax = maxForAxisGroup("primary");
    const secondaryMax = maxForAxisGroup("secondary");
    const hasSecondary = this.series.some((series) => series.axisGroup === "secondary");
    const plot = { left: p.left + 42, top: p.top + 42, width: Math.max(0, p.width - 72), height: Math.max(0, p.height - 82) };
    const title = `<text x="${p.left + 12}" y="${p.top + 24}" font-family="Arial" font-size="16" font-weight="700" fill="#0f172a">${xmlEscape(this.title || this.chartType)}</text>`;
    const axes = `<line x1="${plot.left}" y1="${plot.top + plot.height}" x2="${plot.left + plot.width}" y2="${plot.top + plot.height}" stroke="#94a3b8"/><line x1="${plot.left}" y1="${plot.top}" x2="${plot.left}" y2="${plot.top + plot.height}" stroke="#94a3b8"/>${hasSecondary ? `<line x1="${plot.left}" y1="${plot.top}" x2="${plot.left + plot.width}" y2="${plot.top}" stroke="#64748b"/><line x1="${plot.left + plot.width}" y1="${plot.top}" x2="${plot.left + plot.width}" y2="${plot.top + plot.height}" stroke="#64748b"/>` : ""}${this.axes.category.title ? `<text x="${plot.left + plot.width / 2 - 24}" y="${p.top + p.height - 4}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.category.title)}</text>` : ""}${this.axes.value.title ? `<text x="${p.left + 8}" y="${plot.top + 10}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.value.title)}</text>` : ""}${this.axes.secondary?.category?.title ? `<text x="${plot.left + plot.width / 2 - 24}" y="${plot.top - 4}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.secondary.category.title)}</text>` : ""}${this.axes.secondary?.value?.title ? `<text x="${plot.left + plot.width - 2}" y="${plot.top + 10}" text-anchor="end" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.secondary.value.title)}</text>` : ""}`;
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
        const point = series.points?.find((item) => item.idx === index);
        const color = resolveColorToken(point?.fill || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4], "#0ea5e9");
        const effectiveLabels = series.dataLabels || this.dataLabels;
        const labelText = presentationChartDataLabelText(effectiveLabels, categories[index], value);
        const label = labelText ? `<text x="${cx + (radius + 8) * Math.cos((angle + next) / 2)}" y="${cy + (radius + 8) * Math.sin((angle + next) / 2)}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(labelText)}</text>` : "";
        const path = `<path d="${pieSlicePath(cx, cy, radius, angle, next)}" fill="${xmlEscape(color)}"${presentationChartLineSvgAttributes(point?.line || series.line) || ' stroke="#ffffff"'}/>${label}`;
        angle = next;
        return path;
      }).join("");
      const categoryLegend = categories.map((category, index) => `<rect x="${p.left + p.width - 82}" y="${p.top + 18 + index * 16}" width="10" height="10" fill="${xmlEscape(resolveColorToken(series.points?.find((item) => item.idx === index)?.fill || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4], "#0ea5e9"))}"/><text x="${p.left + p.width - 68}" y="${p.top + 27 + index * 16}" font-family="Arial" font-size="10" fill="#334155">${xmlEscape(category)}</text>`).join("");
      return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>${title}${slices}${this.legend.visible ? categoryLegend : ""}`;
    }
    const lineBody = lineSeries.map((series, seriesIndex) => {
        const axisGroup = series.axisGroup || "primary";
        const seriesMax = axisGroup === "secondary" ? secondaryMax : primaryMax;
        const points = (series.values || []).map((value, index) => {
          const stackedValue = stackedLines ? lineSeries.slice(0, seriesIndex + 1).filter((item) => (item.axisGroup || "primary") === axisGroup).reduce((sum, item) => sum + Math.max(0, Number(item.values?.[index]) || 0), 0) : Number(value) || 0;
          const plottedValue = this.lineOptions.grouping === "percentStacked" ? stackedValue / (lineStackedMax[axisGroup][index] || 1) : stackedValue;
          const x = plot.left + (categories.length <= 1 ? plot.width / 2 : (index / Math.max(1, categories.length - 1)) * plot.width);
          const y = plot.top + plot.height - (plottedValue / seriesMax) * plot.height;
          return { x, y, index };
        });
        const color = resolveColorToken(series.line?.fill || series.color, series.color);
        const smooth = series.smooth ?? this.lineOptions.smooth;
        const strokeAttributes = presentationChartLineSvgAttributes(series.line) || ` stroke="${xmlEscape(color)}" stroke-width="2"`;
        const line = smooth && points.length > 2
          ? `<path d="M ${points[0].x} ${points[0].y} ${points.slice(1, -1).map((point, index) => { const next = points[index + 2]; return `Q ${point.x} ${point.y} ${(point.x + next.x) / 2} ${(point.y + next.y) / 2}`; }).join(" ")} T ${points.at(-1).x} ${points.at(-1).y}" fill="none"${strokeAttributes}/>`
          : `<polyline points="${points.map((point) => `${point.x},${point.y}`).join(" ")}" fill="none"${strokeAttributes}/>`;
        const marker = series.marker || this.lineOptions.marker;
        const effectiveLabels = series.dataLabels || this.dataLabels;
        const labels = points.map((point, index) => {
          const label = presentationChartDataLabelText(effectiveLabels, categories[index], series.values?.[index]);
          return label ? `<text x="${point.x + 4}" y="${point.y - 4}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(label)}</text>` : "";
        }).join("");
        return `${line}${presentationChartErrorBarsSvg(series, points, plot, seriesMax)}${points.map((point, index) => presentationChartMarkerSvg(marker, point.x, point.y, resolveColorToken(series.points?.find((item) => item.idx === index)?.fill || color, color))).join("")}${labels}`;
      }).join("");
    const horizontal = barSeries.length > 0 && this.barOptions.direction === "bar";
    const barBody = (() => {
      const groupExtent = categories.length ? (horizontal ? plot.height : plot.width) / categories.length : 0;
      const gapRatio = Math.max(0.12, 100 / (100 + this.barOptions.gapWidth));
      const barExtent = stackedBars ? groupExtent * gapRatio : groupExtent * gapRatio / Math.max(1, barSeries.length);
      const offsets = { primary: categories.map(() => 0), secondary: categories.map(() => 0) };
      return barSeries.flatMap((series, seriesIndex) => (series.values || []).map((rawValue, categoryIndex) => {
        const axisGroup = series.axisGroup || "primary";
        const seriesMax = axisGroup === "secondary" ? secondaryMax : primaryMax;
        const total = barStackedMax[axisGroup][categoryIndex] || 1;
        const value = Math.max(0, Number(rawValue) || 0);
        const ratio = this.barOptions.grouping === "percentStacked" ? value / total : value / seriesMax;
        const offset = offsets[axisGroup][categoryIndex];
        offsets[axisGroup][categoryIndex] += ratio;
        const point = series.points?.find((item) => item.idx === categoryIndex);
        const color = xmlEscape(resolveColorToken(point?.fill || series.color, series.color));
        const stroke = presentationChartLineSvgAttributes(point?.line || series.line);
        const labelText = presentationChartDataLabelText(series.dataLabels || this.dataLabels, categories[categoryIndex], rawValue);
        if (horizontal) {
          const width = plot.width * ratio;
          const x = plot.left + (stackedBars ? plot.width * offset : 0);
          const y = plot.top + categoryIndex * groupExtent + (stackedBars ? (groupExtent - barExtent) / 2 : (groupExtent - barExtent * barSeries.length) / 2 + seriesIndex * barExtent);
          const label = labelText ? `<text x="${x + width + 3}" y="${y + barExtent - 2}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(labelText)}</text>` : "";
          const errorBars = presentationChartErrorBarsSvg(series, [{ x: x + width, y: y + Math.max(1, barExtent - 2) / 2, index: categoryIndex }], plot, seriesMax);
          return `<rect x="${x}" y="${y}" width="${width}" height="${Math.max(1, barExtent - 2)}" fill="${color}"${stroke}/>${errorBars}${label}`;
        }
        const height = plot.height * ratio;
        const x = plot.left + categoryIndex * groupExtent + (stackedBars ? (groupExtent - barExtent) / 2 : (groupExtent - barExtent * barSeries.length) / 2 + seriesIndex * barExtent);
        const y = plot.top + plot.height - height - (stackedBars ? plot.height * offset : 0);
        const label = labelText ? `<text x="${x}" y="${y - 4}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(labelText)}</text>` : "";
        const errorBars = presentationChartErrorBarsSvg(series, [{ x: x + Math.max(1, barExtent - 2) / 2, y, index: categoryIndex }], plot, seriesMax);
        return `<rect x="${x}" y="${y}" width="${Math.max(1, barExtent - 2)}" height="${height}" fill="${color}"${stroke}/>${errorBars}${label}`;
      })).join("");
    })();
    const trendlineBody = `${barSeries.map((series) => presentationChartTrendlinesSvg(series, plot, series.axisGroup === "secondary" ? secondaryMax : primaryMax, categories.length, { horizontal, centered: true })).join("")}${lineSeries.map((series) => presentationChartTrendlinesSvg(series, plot, series.axisGroup === "secondary" ? secondaryMax : primaryMax, categories.length)).join("")}`;
    const body = `${barBody}${lineBody}${trendlineBody}`;
    const labels = this.chartType === "bar" && horizontal
      ? categories.map((category, index) => `<text x="${plot.left - 4}" y="${plot.top + (index + 0.6) * (plot.height / Math.max(1, categories.length))}" text-anchor="end" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(category)}</text>`).join("")
      : categories.map((category, index) => `<text x="${plot.left + index * (plot.width / Math.max(1, categories.length))}" y="${p.top + p.height - 18}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(category)}</text>`).join("");
    return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>${title}${axes}${body}${labels}${legend}`;
  }

}

export class ImageElement {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("im");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
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
    this.transform = config.transform == null ? undefined : normalizePresentationPlaceholderTransform(config.transform, `Presentation image ${this.name || this.id} transform`);
  }

  get frame() { return this.position; }
  set frame(value) { this.position = normalizeFrame(value, this.position); }
  replace(config = {}) { Object.assign(this, config); }

  inspectRecord() {
    const p = this.position;
    return { kind: "image", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, alt: this.alt || undefined, prompt: this.prompt || undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", fit: this.fit, transform: this.transform };
  }

  layoutJson() { return { kind: "image", id: this.id, name: this.name, frame: this.position, alt: this.alt, prompt: this.prompt, uri: this.uri, dataUrl: this.dataUrl, fit: this.fit, geometry: this.geometry, borderRadius: this.borderRadius, transform: this.transform }; }

  toSvg() {
    const p = this.position;
    const label = this.alt || this.prompt || this.uri || "image";
    const cx = p.left + p.width / 2;
    const cy = p.top + p.height / 2;
    const rotation = Number(this.transform?.rotationDegrees || 0);
    const flipHorizontal = this.transform?.flipHorizontal === true ? -1 : 1;
    const flipVertical = this.transform?.flipVertical === true ? -1 : 1;
    const transform = this.transform ? ` transform="translate(${cx} ${cy}) rotate(${rotation}) scale(${flipHorizontal} ${flipVertical}) translate(${-cx} ${-cy})"` : "";
    if (this.dataUrl) return `<image href="${attrEscape(this.dataUrl)}" x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" preserveAspectRatio="${this.fit === "stretch" ? "none" : "xMidYMid meet"}"${transform}/>`;
    const rect = this.geometry === "ellipse"
      ? `<ellipse cx="${p.left + p.width / 2}" cy="${p.top + p.height / 2}" rx="${p.width / 2}" ry="${p.height / 2}" fill="#e0f2fe" stroke="#0284c7"/>`
      : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" rx="${this.borderRadius ? 12 : 0}" fill="#e0f2fe" stroke="#0284c7"/>`;
    const fallback = `${rect}<text x="${p.left + 12}" y="${p.top + 28}" font-family="Arial" font-size="14" fill="#075985">${xmlEscape(label)}</text>`;
    return transform ? `<g${transform}>${fallback}</g>` : fallback;
  }

}

export class PresentationFile {
  static async inspectPptx(blobOrBuffer, options = {}) {
    return inspectOoxmlPackage(blobOrBuffer, options, PPTX_PACKAGE_CONFIG);
  }

  static async patchPptx(blobOrBuffer, patches = [], options = {}) {
    const patched = await patchOoxmlPackage(blobOrBuffer, patches, options, PPTX_PACKAGE_CONFIG);
    return new FileBlob(patched.bytes, { type: PPTX_MIME, metadata: { artifactKind: "presentation", patchedParts: patched.patchedParts, recipesApplied: patched.recipesApplied, contentTypesUpdated: patched.contentTypesUpdated, relationshipsUpdated: patched.relationshipsUpdated, sourceReferencesUpdated: patched.sourceReferencesUpdated, validated: patched.validated, validationIssues: patched.validationIssues } });
  }

  static async exportPptx(presentation, options = {}) {
    const { exportPptxWithOpenChestnut } = await import("./codecs/open-chestnut.mjs");
    return exportPptxWithOpenChestnut(presentation, options);
  }

  static async importPptx(blobOrBuffer, options = {}) {
    const { importPptxWithOpenChestnut } = await import("./codecs/open-chestnut.mjs");
    return importPptxWithOpenChestnut(blobOrBuffer, options);
  }
}

function presentationElementKind(element) {
  if (element instanceof NativePresentationObject) return "nativeObject";
  if (element instanceof ConnectorElement) return "connector";
  if (element instanceof GroupShape) return "groupShape";
  if (element instanceof TableElement) return "table";
  if (element instanceof ChartElement) return "chart";
  if (element instanceof ImageElement) return "image";
  return "shape";
}

function presentationSlideElements(slide) {
  const direct = [...slide.connectors.items, ...slide.shapes.items, ...slide.tables.items, ...slide.charts.items, ...slide.images.items, ...slide.groups.items, ...slide.nativeObjects.items];
  return direct.flatMap((element) => element instanceof GroupShape ? element.allElements() : [element]);
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
  cascade(id, seen = new Set()) {
    const style = this.get(id);
    if (!style || seen.has(style.id)) return [];
    seen.add(style.id);
    const parentId = style.basedOn || style.parent || style.extends;
    return [...(parentId ? this.cascade(parentId, seen) : []), style];
  }
  values() { return [...this.items.values()]; }
}

class DocumentTableCell {
  constructor(table, row, column) {
    this.table = table;
    this.kind = "tableCell";
    this.tableId = table.id;
    this.row = Number(row);
    this.column = Number(column);
    this.id = `${table.id}/cell/${this.row}/${this.column}`;
  }
  _record() { return this.table.cells?.find((cell) => cell.row === this.row && cell.column === this.column); }
  get value() { return this.table.values[this.row]?.[this.column] ?? ""; }
  set value(value) { this.table.ensureCell(this.row, this.column); this.table.values[this.row][this.column] = value; }
  get gridColumn() { return this._record()?.gridColumn ?? this.column; }
  get columnSpan() { return this._record()?.columnSpan ?? 1; }
  get rowSpan() { return this._record()?.rowSpan ?? 1; }
  get verticalMerge() { return this._record()?.verticalMerge ?? "none"; }
  get editable() { return this._record()?.editable ?? true; }
  inspectRecord() { return { kind: this.kind, id: this.id, tableId: this.tableId, row: this.row, column: this.column, gridColumn: this.gridColumn, columnSpan: this.columnSpan, rowSpan: this.rowSpan, verticalMerge: this.verticalMerge, editable: this.editable, value: this.value }; }
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
    this.cells = Array.isArray(config.cells) ? config.cells.map((cell) => {
      const verticalMerge = String(cell.verticalMerge || "none");
      return {
        row: Math.max(0, Math.round(Number(cell.row) || 0)),
        column: Math.max(0, Math.round(Number(cell.column) || 0)),
        gridColumn: Math.max(0, Math.round(Number(cell.gridColumn) || 0)),
        columnSpan: Math.max(1, Math.round(Number(cell.columnSpan) || 1)),
        rowSpan: Math.max(0, Math.round(Number(cell.rowSpan ?? (verticalMerge === "continue" ? 0 : 1)) || 0)),
        verticalMerge,
        editable: verticalMerge === "continue" ? false : cell.editable !== false,
      };
    }) : undefined;
    const derivedGridColumns = this.cells?.reduce((maximum, cell) => Math.max(maximum, cell.gridColumn + cell.columnSpan), 0) || 0;
    this.gridColumns = Math.max(0, Math.round(Number(config.gridColumns ?? Math.max(this.columns, derivedGridColumns))));
    const formattingColumns = this.cells?.length ? this.gridColumns : this.columns;
    this.widthDxa = Math.round(Number(config.widthDxa ?? 9360));
    this.indentDxa = Math.round(Number(config.indentDxa ?? 120));
    this.columnWidthsDxa = Array.isArray(config.columnWidthsDxa)
      ? config.columnWidthsDxa.map((value) => Math.round(Number(value)))
      : documentTableDefaultColumnWidths(formattingColumns, this.widthDxa);
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
  inspectRecord(index) { return { kind: "table", id: this.id, index, name: this.name || undefined, rows: this.rows, cols: this.columns, gridColumns: this.gridColumns, cells: this.cells, styleId: this.styleId, widthDxa: this.widthDxa, indentDxa: this.indentDxa, columnWidthsDxa: this.columnWidthsDxa, cellMarginsDxa: this.cellMarginsDxa, borderColor: this.borderColor, borderSize: this.borderSize, headerFill: this.headerFill, values: this.values }; }
  toProto() { return { kind: "table", id: this.id, name: this.name, styleId: this.styleId, gridColumns: this.gridColumns, cells: this.cells, widthDxa: this.widthDxa, indentDxa: this.indentDxa, columnWidthsDxa: this.columnWidthsDxa, cellMarginsDxa: this.cellMarginsDxa, borderColor: this.borderColor, borderSize: this.borderSize, headerFill: this.headerFill, values: this.values }; }
}

function normalizeDocumentRuns(text, config = {}, theme = {}) {
  const runs = (config.runs || config.textRuns || []).map((run) => ({ text: String(run.text ?? run.value ?? ""), style: normalizeDocxRunStyle(run.style || run.textStyle || {}, theme) })).filter((run) => run.text.length > 0);
  if (runs.length) return runs;
  const rawText = String(text ?? "");
  return rawText ? [{ text: rawText, style: normalizeDocxRunStyle({}, theme) }] : [];
}

function documentEffectiveRunStyle(document, block, run) {
  const characterStyleId = run.style?.runStyleId;
  const cascade = [document.defaultRunStyle, ...document.styles.cascade(block.styleId), ...(characterStyleId ? document.styles.cascade(characterStyleId) : []), run.style || {}];
  return effectiveDocxRunStyle(mergeDocxRunStyleCascade(cascade, document.theme), run.text, document.theme);
}

function documentRunsNeedSerialization(runs = []) {
  return runs.length > 1 || runs.some((run) => Object.keys(run.style || {}).length > 0);
}

class DocumentParagraphBlock {
  constructor(document, text, config = {}) {
    this.document = document;
    this.kind = "paragraph";
    this.id = config.id || aid("dp");
    this.runs = normalizeDocumentRuns(text, config, document.theme);
    this.text = this.runs.map((run) => run.text).join("") || String(text ?? "");
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
    this.paragraphFormat = { ...(config.paragraphFormat || config.formatting || {}) };
  }

  inspectRecord(index) { return { kind: "paragraph", id: this.id, index, name: this.name || undefined, styleId: this.styleId, paragraphFormat: Object.keys(this.paragraphFormat).length ? this.paragraphFormat : undefined, text: this.text, textChars: this.text.length, runs: documentRunsNeedSerialization(this.runs) ? this.runs : undefined }; }
  toProto() { return { kind: "paragraph", id: this.id, name: this.name, styleId: this.styleId, paragraphFormat: Object.keys(this.paragraphFormat).length ? this.paragraphFormat : undefined, text: this.text, runs: documentRunsNeedSerialization(this.runs) ? this.runs : undefined }; }
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
    const configuredFormat = config.numberFormat || config.numFmt;
    this.listType = config.listType || config.type || (configuredFormat === "bullet" ? "bullet" : configuredFormat ? "number" : "bullet");
    this.numberFormat = configuredFormat || (this.listType === "number" ? "decimal" : "bullet");
    this.start = Number(config.start ?? 1);
    this.levelText = config.levelText || config.lvlText || (this.listType === "number" ? `%${this.level + 1}.` : "•");
    this.pictureBullet = normalizeDocumentPictureBullet(config.pictureBullet ?? config.bulletImage ?? config.bullet?.image);
    if (this.pictureBullet && (this.listType !== "bullet" || this.numberFormat !== "bullet")) throw new TypeError("DOCX picture bullet requires listType and numberFormat to be bullet.");
    this.numberingId = config.numberingId ?? config.numId;
    this.abstractNumberingId = config.abstractNumberingId ?? config.abstractNumId;
    this.numberingStyleId = config.numberingStyleId ?? config.numStyleId;
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "listItem", id: this.id, index, name: this.name || undefined, styleId: this.styleId, listType: this.listType, level: this.level, numberFormat: this.numberFormat, start: this.start, levelText: this.levelText, pictureBullet: this.pictureBullet, numberingId: this.numberingId, abstractNumberingId: this.abstractNumberingId, numberingStyleId: this.numberingStyleId, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: "listItem", id: this.id, name: this.name, styleId: this.styleId, listType: this.listType, level: this.level, numberFormat: this.numberFormat, start: this.start, levelText: this.levelText, pictureBullet: this.pictureBullet, numberingId: this.numberingId, abstractNumberingId: this.abstractNumberingId, numberingStyleId: this.numberingStyleId, text: this.text }; }
}

class DocumentHyperlinkBlock {
  constructor(document, text, url, config = {}) {
    this.document = document;
    this.kind = "hyperlink";
    this.id = config.id || aid("dhl");
    this.text = String(text ?? "");
    const target = typeof url === "object" && url ? url : config.anchor || config.bookmark;
    const rawUrl = typeof url === "string" ? url : config.url;
    this.anchor = String(target?.name || target || config.anchor?.name || config.anchor || config.bookmarkName || (String(rawUrl || "").startsWith("#") ? String(rawUrl).slice(1) : "")).trim() || undefined;
    this.url = this.anchor ? "" : String(rawUrl || "");
    this.relationshipId = config.relationshipId || config.relId;
    this.tooltip = config.tooltip;
    this.history = config.history !== false;
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "hyperlink", id: this.id, index, name: this.name || undefined, styleId: this.styleId, relationshipId: this.relationshipId, text: this.text, url: this.url || undefined, anchor: this.anchor, tooltip: this.tooltip, history: this.history, textChars: this.text.length }; }
  toProto() { return { kind: "hyperlink", id: this.id, name: this.name, styleId: this.styleId, relationshipId: this.relationshipId, text: this.text, url: this.url || undefined, anchor: this.anchor, tooltip: this.tooltip, history: this.history }; }
}

function documentBookmarkEndpoint(value) {
  if (!value) return undefined;
  const raw = typeof value === "object" ? value : undefined;
  const id = typeof value === "string" ? value : raw?.id;
  const cellMatch = /^(.+)\/cell\/(\d+)\/(\d+)$/.exec(String(id || ""));
  if (cellMatch || raw?.kind === "tableCell" || raw?.type === "tableCell" || raw?.tableId !== undefined) {
    const tableId = String(raw?.tableId || raw?.table?.id || cellMatch?.[1] || "");
    const row = Number(raw?.row ?? raw?.rowIndex ?? cellMatch?.[2]);
    const column = Number(raw?.column ?? raw?.columnIndex ?? cellMatch?.[3]);
    return { type: "tableCell", id: `${tableId}/cell/${row}/${column}`, tableId, row, column };
  }
  return { type: "block", id: String(id || raw?.blockId || value), blockId: String(raw?.blockId || id || value) };
}

function documentBookmarkEndpointBlockId(endpoint) {
  return endpoint?.type === "tableCell" ? endpoint.tableId : endpoint?.blockId;
}

function documentBookmarkEndpointOrder(endpoint, blockIndexes) {
  const blockIndex = blockIndexes.get(documentBookmarkEndpointBlockId(endpoint));
  return endpoint?.type === "tableCell" ? [blockIndex, 1, endpoint.row, endpoint.column] : [blockIndex, 0, 0, 0];
}

function compareDocumentBookmarkEndpoints(left, right, blockIndexes) {
  const a = documentBookmarkEndpointOrder(left, blockIndexes);
  const b = documentBookmarkEndpointOrder(right, blockIndexes);
  for (let index = 0; index < a.length; index += 1) if (a[index] !== b[index]) return a[index] - b[index];
  return 0;
}

class DocumentBookmark {
  constructor(document, target, name, config = {}) {
    this.document = document;
    this.kind = "bookmark";
    this.id = config.id || aid("dbm");
    this.name = String(name ?? config.name ?? "");
    this.target = documentBookmarkEndpoint(target ?? config.target ?? config.targetId);
    const endTarget = config.endTarget ?? config.end ?? config.endTargetId ?? this.target;
    this.endTarget = documentBookmarkEndpoint(endTarget);
    this.targetId = this.target?.id;
    this.endTargetId = this.endTarget?.id || this.targetId;
    this.nativeId = config.nativeId === undefined ? undefined : Number(config.nativeId);
  }

  inspectRecord() { return { kind: "bookmark", id: this.id, name: this.name, targetId: this.targetId, endTargetId: this.endTargetId, target: this.target, endTarget: this.endTarget, nativeId: this.nativeId }; }
  toProto() { return { kind: "bookmark", id: this.id, name: this.name, targetId: this.targetId, endTargetId: this.endTargetId, target: this.target, endTarget: this.endTarget, nativeId: this.nativeId }; }
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

class DocumentBibliographySource {
  constructor(document, config = {}) {
    this.document = document;
    Object.assign(this, normalizeDocxBibliographySource(config, document.bibliographySources.length));
  }

  inspectRecord(index) { return { ...this.toProto(), index }; }
  toProto() { return Object.fromEntries(Object.entries(this).filter(([key, value]) => key !== "document" && value !== undefined)); }
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
    this.referenceType = ["default", "first", "even"].includes(config.referenceType || config.type) ? (config.referenceType || config.type) : "default";
    this.relationshipId = config.relationshipId || config.relId;
    this.partPath = config.partPath;
    this.sectionIndex = config.sectionIndex === undefined || config.sectionIndex === null ? undefined : Number(config.sectionIndex);
    this.variantActive = config.variantActive ?? config.activateVariant ?? config.active;
    this.fieldInstruction = config.fieldInstruction || config.field;
  }

  inspectRecord(index) { return { kind: this.kind, id: this.id, index, name: this.name || undefined, styleId: this.styleId, referenceType: this.referenceType, variantActive: this.variantActive, relationshipId: this.relationshipId, partPath: this.partPath, sectionIndex: this.sectionIndex, fieldInstruction: this.fieldInstruction, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: this.kind, id: this.id, name: this.name, styleId: this.styleId, referenceType: this.referenceType, variantActive: this.variantActive, relationshipId: this.relationshipId, partPath: this.partPath, sectionIndex: this.sectionIndex, fieldInstruction: this.fieldInstruction, text: this.text }; }
}

function documentCommentInitials(author) {
  const words = String(author || "User").trim().split(/\s+/).filter(Boolean);
  return (words.length > 1 ? words.map((word) => word[0]).join("") : (words[0] || "U").slice(0, 2)).slice(0, 4).toUpperCase();
}

class DocumentComment {
  constructor(document, targetId, text, config = {}) {
    this.document = document;
    this.kind = "comment";
    this.id = config.id || aid("dc");
    this.targetId = targetId;
    this.author = config.author || "User";
    this.initials = config.initials || documentCommentInitials(this.author);
    this.date = config.date || undefined;
    this.text = String(text ?? "");
    this.resolved = Boolean(config.resolved);
    this.parentId = typeof (config.parentId ?? config.replyToId ?? config.replyTo) === "object" ? (config.parentId ?? config.replyToId ?? config.replyTo)?.id : (config.parentId ?? config.replyToId ?? config.replyTo);
    this.paraId = config.paraId ? String(config.paraId).toUpperCase() : undefined;
    this.durableId = config.durableId ? String(config.durableId).toUpperCase() : undefined;
    this.dateUtc = config.dateUtc;
    this.person = config.person ? { ...config.person } : (config.providerId || config.userId ? { providerId: config.providerId, userId: config.userId } : undefined);
    this.intelligentPlaceholder = Boolean(config.intelligentPlaceholder);
  }

  inspectRecord() { return { kind: "comment", id: this.id, targetId: this.targetId, parentId: this.parentId, paraId: this.paraId, durableId: this.durableId, author: this.author, initials: this.initials, date: this.date, dateUtc: this.dateUtc, person: this.person, intelligentPlaceholder: this.intelligentPlaceholder || undefined, resolved: this.resolved, textPreview: this.text.slice(0, 300) }; }
  toProto() { return { kind: "comment", id: this.id, targetId: this.targetId, parentId: this.parentId, paraId: this.paraId, durableId: this.durableId, author: this.author, initials: this.initials, date: this.date, dateUtc: this.dateUtc, person: this.person, intelligentPlaceholder: this.intelligentPlaceholder || undefined, text: this.text, resolved: this.resolved }; }
}

function documentBlockHeight(document, block, pageWidth = 612, margin = 72) {
  if (block.kind === "table") return Math.max(24, block.rows * 24 + 16);
  if (block.kind === "image") return Math.max(32, Math.min(360, Number(block.heightPx) || 160)) + 20;
  if (block.kind === "section") return 34;
  if (block.kind === "change") return 22;
  const style = document.styles.effective(block.styleId) || document.styles.get("Normal") || {};
  const runSizes = block.kind === "paragraph" ? (block.runs || []).map((run) => documentEffectiveRunStyle(document, block, run).effectiveFontSize).filter(Number.isFinite) : [];
  const fontSize = Math.max(10, Math.max(style.fontSize || 22, ...runSizes) / 2);
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
  const headerFooterPlan = planDocxHeaderFooterSections(document);
  const pages = [];
  const elements = [];
  let page = 1;
  let sectionIndex = 0;
  let pageInSection = 1;
  let y = margin;
  const ensurePage = () => {
    if (!pages.find((item) => item.page === page)) {
      const headerFooter = resolveDocxPageHeaderFooter(headerFooterPlan, sectionIndex, pageInSection);
      pages.push({ id: `${document.id}/page/${page}`, page, width: pageWidth, height: pageHeight, margin, sectionIndex, sectionIndexes: [sectionIndex], ...headerFooter });
    }
  };
  ensurePage();
  for (const block of document.blocks) {
    const height = documentBlockHeight(document, block, pageWidth, margin);
    if (y + height > pageHeight - margin && y > margin) { page += 1; pageInSection += 1; y = margin; ensurePage(); }
    const textPreview = documentBlockLayoutText(block).slice(0, 120);
    const comments = document.comments.filter((comment) => comment.targetId === block.id).map((comment) => comment.id);
    const effectiveStyle = block.styleId ? document.styles.effective(block.styleId) : undefined;
    const runs = block.kind === "paragraph" && block.runs?.length ? block.runs.map((run) => ({ text: run.text, style: documentEffectiveRunStyle(document, block, run) })) : undefined;
    elements.push({ kind: "layoutElement", id: block.id, layoutId: `${block.id}/layout`, blockKind: block.kind, name: block.name || undefined, textRangeId: ("text" in block || "display" in block) ? `${block.id}/text` : undefined, commentIds: comments.length ? comments : undefined, page, bbox: [margin, y, pageWidth - margin * 2, height], styleId: block.styleId, effectiveStyle, runs, textPreview });
    y += height;
    if (block.kind === "section") {
      sectionIndex += 1;
      pageInSection = 1;
      if (block.breakType === "nextPage") { page += 1; y = margin; ensurePage(); }
      else {
        const currentPage = pages.find((item) => item.page === page);
        if (currentPage && !currentPage.sectionIndexes.includes(sectionIndex)) currentPage.sectionIndexes.push(sectionIndex);
        if (currentPage) currentPage.continuousSectionBoundary = true;
      }
    }
  }
  return documentLayoutSlice(document, { schema: "open-office-artifact.document-layout/v1", unit: "px", document: { id: document.id, name: document.name, designPreset: document.designPreset, theme: document.theme }, pages, elements }, options);
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
    this.theme = normalizeDocxThemeConfig(options.theme || {});
    this.defaultRunStyle = normalizeDocxRunStyle(options.defaultRunStyle || {}, this.theme);
    this.settings = normalizeDocxSettings(options.settings || {});
    this.styles = new DocumentStyleCollection(options.styles || {});
    this.blocks = [];
    this.bookmarks = [];
    this.comments = [];
    this.bibliography = {
      selectedStyle: String(options.bibliography?.selectedStyle || ""),
      styleName: String(options.bibliography?.styleName || ""),
      uri: String(options.bibliography?.uri || ""),
    };
    this.bibliographySources = [];
    this.headers = [];
    this.footers = [];
    this.sectionSettings = [];
    const preservesEvenOddActivation = Object.prototype.hasOwnProperty.call(options.settings || {}, "evenAndOddHeaders");
    for (const source of options.bibliographySources || options.bibliography?.sources || []) this.addBibliographySource(source);
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
      else if (!block.kind || block.kind === "paragraph") this.addParagraph(block.text ?? "", block);
      else throw new TypeError(`Unsupported document block kind ${block.kind}.`);
    }
    this.sectionSettings = normalizeDocxSectionSettings(options.sectionSettings || [], this.blocks);
    for (const header of options.headers || []) this.addHeader(header.text, { ...header, _restore: true });
    for (const footer of options.footers || []) this.addFooter(footer.text, { ...footer, _restore: true });
    if (!preservesEvenOddActivation && [...this.headers, ...this.footers].some((block) => block.referenceType === "even" && block.variantActive !== false)) this.settings = normalizeDocxSettings({ ...this.settings, evenAndOddHeaders: true });
    for (const bookmark of options.bookmarks || []) this.addBookmark(bookmark.targetId, bookmark.name, bookmark);
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
  addBibliographySource(config = {}) { const source = new DocumentBibliographySource(this, config); this.bibliographySources.push(source); return source; }
  addCitation(text, metadata = {}, config = {}) {
    const block = new DocumentCitationBlock(this, text, metadata, config);
    const explicitTag = block.metadata.tag ?? block.metadata.bibliographyTag;
    const seed = explicitTag || block.metadata.source || block.metadata.title || block.name || `Source${this.bibliographySources.length + 1}`;
    const baseTag = explicitTag ? String(explicitTag).trim() : String(seed).replace(/[^A-Za-z0-9_.:-]+/g, "").slice(0, 255) || `Source${this.bibliographySources.length + 1}`;
    let tag = baseTag;
    if (!explicitTag) for (let suffix = 2; this.bibliographySources.some((source) => source.tag === tag && source.title !== (block.metadata.title || block.metadata.source || block.text)); suffix += 1) tag = `${baseTag.slice(0, 250)}${suffix}`;
    block.metadata.tag = tag;
    if (!this.bibliographySources.some((source) => source.tag === tag)) {
      const sourceConfig = { ...block.metadata, ...(block.metadata.bibliography || block.metadata.sourceData || {}), tag, title: block.metadata.title || block.metadata.source || block.text, sourceType: block.metadata.sourceType || (block.metadata.url ? "InternetSite" : "Misc") };
      this.addBibliographySource(sourceConfig);
    }
    this.blocks.push(block);
    const bookmarkName = block.metadata.bookmark || `OpenOfficeCitation_${block.id.replace(/[^A-Za-z0-9_]/g, "_")}`;
    const bookmark = this.addBookmark(block, bookmarkName, { id: block.metadata.bookmarkId || `${block.id}/bookmark`, nativeId: block.metadata.bookmarkNativeId });
    block.metadata.bookmark = bookmark.name;
    block.metadata.bookmarkId = bookmark.id;
    return block;
  }
  addImage(config = {}) { const block = new DocumentImageBlock(this, config); this.blocks.push(block); return block; }
  addSection(config = {}) { const block = new DocumentSectionBlock(this, config); this.blocks.push(block); return block; }
  addPageBreakSection(config = {}) { return this.addSection({ ...config, breakType: "nextPage" }); }
  addChange(changeType, text, config = {}) { const block = new DocumentChangeBlock(this, changeType, text, config); this.blocks.push(block); return block; }
  addInsertion(text, config = {}) { return this.addChange("insert", text, config); }
  addDeletion(text, config = {}) { return this.addChange("delete", text, config); }
  addTable(config = {}) { const block = new DocumentTableBlock(this, config); this.blocks.push(block); return block; }
  addHeader(text, config = {}) { const block = new DocumentHeaderFooterBlock(this, "header", text, config); this.headers.push(block); if (!config._restore && block.referenceType === "even" && block.variantActive !== false) this.settings = normalizeDocxSettings({ ...this.settings, evenAndOddHeaders: true }); return block; }
  addFooter(text, config = {}) { const block = new DocumentHeaderFooterBlock(this, "footer", text, config); this.footers.push(block); if (!config._restore && block.referenceType === "even" && block.variantActive !== false) this.settings = normalizeDocxSettings({ ...this.settings, evenAndOddHeaders: true }); return block; }
  addBookmark(target, name, config = {}) {
    const targetEndpoint = documentBookmarkEndpoint(target ?? config.target ?? config.targetId);
    const targetId = targetEndpoint?.id;
    const bookmarkName = String(name || config.name || "");
    const existing = this.bookmarks.find((bookmark) => bookmark.name === bookmarkName);
    if (existing && existing.targetId === targetId) {
      if (config.nativeId !== undefined) existing.nativeId = Number(config.nativeId);
      const configuredEnd = config.endTarget ?? config.end ?? config.endTargetId;
      if (configuredEnd) {
        existing.endTarget = documentBookmarkEndpoint(configuredEnd);
        existing.endTargetId = existing.endTarget?.id;
      }
      return existing;
    }
    const bookmark = new DocumentBookmark(this, target, name, config);
    this.bookmarks.push(bookmark);
    return bookmark;
  }
  addComment(target, text, config = {}) { const targetId = typeof target === "string" ? target : target?.id; const comment = new DocumentComment(this, targetId, text, config); this.comments.push(comment); return comment; }
  replyToComment(parent, text, config = {}) { const comment = typeof parent === "string" ? this.comments.find((item) => item.id === parent) : parent; if (!comment || !this.comments.includes(comment)) throw new Error(`Unknown parent document comment: ${typeof parent === "string" ? parent : parent?.id}`); return this.addComment(comment.targetId, text, { ...config, parentId: comment.id }); }
  setSettings(settings = {}) { this.settings = normalizeDocxSettings({ ...this.settings, ...settings }); return this; }
  setSectionSettings(sectionIndex, settings = {}) { this.sectionSettings = normalizeDocxSectionSettings([...this.sectionSettings.filter((entry) => entry.sectionIndex !== Number(sectionIndex)), { sectionIndex: Number(sectionIndex), ...settings }], this.blocks); return this; }
  resolve(id) {
    const token = String(id || "");
    if (token.endsWith("/text")) return documentTextRange(this, token);
    const cellMatch = /^(.+)\/cell\/(\d+)\/(\d+)$/.exec(token);
    if (cellMatch) {
      const table = this.blocks.find((block) => block.kind === "table" && block.id === cellMatch[1]);
      const row = Number(cellMatch[2]);
      const column = Number(cellMatch[3]);
      if (table && row < table.rows && column < table.columns) return table.getCell(row, column);
      return undefined;
    }
    return token === `${this.id}/settings` ? this.settings : token === `${this.id}/theme` ? this.theme : this.id === token ? this : this.blocks.find((block) => block.id === token) || this.headers.find((block) => block.id === token) || this.footers.find((block) => block.id === token) || this.bookmarks.find((bookmark) => bookmark.id === token || bookmark.name === token) || this.comments.find((comment) => comment.id === token) || this.bibliographySources.find((source) => source.id === token || source.tag === token) || this.styles.get(token);
  }

  toProto() { return { id: this.id, name: this.name, designPreset: this.designPreset, theme: this.theme, defaultRunStyle: this.defaultRunStyle, settings: this.settings, bibliography: this.bibliography, bibliographySources: this.bibliographySources.map((source) => source.toProto()), sectionSettings: this.sectionSettings, styles: Object.fromEntries(this.styles.values().map((style) => [style.id, style])), blocks: this.blocks.map((block) => block.toProto()), headers: this.headers.map((block) => block.toProto()), footers: this.footers.map((block) => block.toProto()), bookmarks: this.bookmarks.map((bookmark) => bookmark.toProto()), comments: this.comments.map((comment) => comment.toProto()) }; }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["paragraph", "table", "listItem", "hyperlink", "field", "citation", "bibliographySource", "image", "section", "change", "bookmark", "comment", "header", "footer"]);
    const records = [];
    if (kinds.has("document")) records.push({ kind: "document", id: this.id, name: this.name, blocks: this.blocks.length, sections: this.blocks.filter((block) => block.kind === "section").length + 1, bookmarks: this.bookmarks.length, bibliographySources: this.bibliographySources.length, designPreset: this.designPreset, defaultRunStyle: this.defaultRunStyle, settings: this.settings, sectionSettings: this.sectionSettings });
    if (kinds.has("theme")) records.push({ kind: "theme", id: `${this.id}/theme`, ...this.theme });
    if (kinds.has("settings")) records.push({ kind: "settings", id: `${this.id}/settings`, ...this.settings });
    if (kinds.has("layout")) records.push(...documentLayoutRecords(this, options));
    this.blocks.forEach((block, index) => { if (kinds.has(block.kind)) records.push(documentInspectRecord(this, block, index)); });
    if (kinds.has("tableCell")) for (const table of this.blocks.filter((block) => block.kind === "table")) for (let row = 0; row < table.rows; row += 1) for (let column = 0; column < table.columns; column += 1) records.push(table.getCell(row, column).inspectRecord());
    if (kinds.has("header")) records.push(...this.headers.map((block, index) => documentInspectRecord(this, block, index)));
    if (kinds.has("footer")) records.push(...this.footers.map((block, index) => documentInspectRecord(this, block, index)));
    if (kinds.has("bookmark")) records.push(...this.bookmarks.map((bookmark) => bookmark.inspectRecord()));
    if (kinds.has("bibliographySource")) records.push(...this.bibliographySources.map((source, index) => source.inspectRecord(index)));
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
    const bookmarkByName = new Map();
    for (const bookmark of this.bookmarks) {
      if (!bookmark.name || bookmark.name.length > 40) issues.push(verificationIssue("document", "invalidBookmarkName", `Bookmark ${bookmark.id} name must contain 1 to 40 characters.`, { id: bookmark.id, name: bookmark.name }));
      if (bookmarkByName.has(bookmark.name)) issues.push(verificationIssue("document", "duplicateBookmarkName", `Bookmark ${bookmark.id} duplicates name ${bookmark.name}.`, { id: bookmark.id, name: bookmark.name }));
      else bookmarkByName.set(bookmark.name, bookmark);
    }
    const bibliographyByTag = new Map();
    for (const source of this.bibliographySources) {
      try { normalizeDocxBibliographySource(source); }
      catch (error) { issues.push(verificationIssue("document", "invalidBibliographySource", error.message, { id: source.id, tag: source.tag })); }
      if (bibliographyByTag.has(source.tag)) issues.push(verificationIssue("document", "duplicateBibliographyTag", `Bibliography source ${source.id} duplicates tag ${source.tag}.`, { id: source.id, tag: source.tag }));
      else bibliographyByTag.set(source.tag, source);
    }
    for (const block of this.blocks) {
      if (block.kind === "paragraph") {
        for (const run of block.runs || []) {
          if (run.style?.runStyleId && !knownStyleIds.has(run.style.runStyleId)) issues.push(verificationIssue("document", "unknownRunStyle", `Paragraph ${block.id} references missing character style ${run.style.runStyleId}.`, { severity: "warning", id: block.id, runStyleId: run.style.runStyleId }));
        }
      }
      if (block.kind === "paragraph" && /^\s*([-*•]|\d+[.)])\s+/.test(block.text)) {
        issues.push(verificationIssue("document", "fakeList", `Paragraph ${block.id} looks like a fake list item; use addListItem instead.`, { id: block.id }));
      }
      if (block.kind === "hyperlink" && block.anchor) {
        if (block.url) issues.push(verificationIssue("document", "ambiguousHyperlink", `Hyperlink ${block.id} cannot combine an external URL with an internal anchor.`, { id: block.id, url: block.url, anchor: block.anchor }));
        if (!bookmarkByName.has(block.anchor)) issues.push(verificationIssue("document", "missingHyperlinkAnchor", `Hyperlink ${block.id} targets missing bookmark ${block.anchor}.`, { id: block.id, anchor: block.anchor }));
        if (block.anchor.length > 255) issues.push(verificationIssue("document", "invalidHyperlinkAnchor", `Hyperlink ${block.id} anchor exceeds 255 characters.`, { id: block.id, anchor: block.anchor }));
      } else if (block.kind === "hyperlink" && !/^https?:\/\//.test(block.url)) {
        issues.push(verificationIssue("document", "invalidHyperlink", `Hyperlink ${block.id} is missing an absolute http(s) URL or internal bookmark anchor.`, { id: block.id, url: block.url }));
      }
      if (block.kind === "hyperlink" && block.tooltip && String(block.tooltip).length > 260) issues.push(verificationIssue("document", "invalidHyperlinkTooltip", `Hyperlink ${block.id} tooltip exceeds 260 characters.`, { id: block.id }));
      if (block.kind === "field" && !block.instruction.trim()) {
        issues.push(verificationIssue("document", "emptyField", `Field ${block.id} is missing an instruction.`, { id: block.id }));
      }
      if (block.kind === "citation" && block.metadata?.url && !/^https?:\/\//.test(String(block.metadata.url))) {
        issues.push(verificationIssue("document", "invalidCitationUrl", `Citation ${block.id} has a non-http(s) URL.`, { severity: "warning", id: block.id, url: block.metadata.url }));
      }
      if (block.kind === "citation" && !bibliographyByTag.has(block.metadata?.tag)) issues.push(verificationIssue("document", "missingCitationSource", `Citation ${block.id} references missing bibliography source ${block.metadata?.tag || "(none)"}.`, { id: block.id, tag: block.metadata?.tag }));
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
      if (block.kind === "listItem") {
        if (!Number.isInteger(block.level) || block.level < 0 || block.level > 8) issues.push(verificationIssue("document", "invalidListLevel", `List item ${block.id} level must be an integer from 0 through 8.`, { id: block.id, level: block.level }));
        if (!Number.isInteger(block.start) || block.start < 1) issues.push(verificationIssue("document", "invalidListStart", `List item ${block.id} start must be a positive integer.`, { id: block.id, start: block.start }));
        if (!String(block.numberFormat || "").trim()) issues.push(verificationIssue("document", "invalidListNumberFormat", `List item ${block.id} is missing numberFormat.`, { id: block.id, numberFormat: block.numberFormat }));
        try {
          const pictureBullet = normalizeDocumentPictureBullet(block.pictureBullet);
          if (pictureBullet && (block.listType !== "bullet" || block.numberFormat !== "bullet")) issues.push(verificationIssue("document", "invalidPictureBulletListType", `List item ${block.id} picture bullet requires bullet list semantics.`, { id: block.id, listType: block.listType, numberFormat: block.numberFormat }));
        } catch (error) {
          issues.push(verificationIssue("document", "invalidPictureBullet", `List item ${block.id} has an invalid picture bullet: ${error.message}`, { id: block.id }));
        }
      }
      if (block.kind === "table") {
        if (!block.rows || !block.columns) issues.push(verificationIssue("document", "emptyTable", `Table ${block.id} has no rows or columns.`, { id: block.id, rows: block.rows, columns: block.columns }));
        if (block.columns > 12) issues.push(verificationIssue("document", "wideTable", `Table ${block.id} has ${block.columns} columns and may not fit the page.`, { severity: "warning", id: block.id, columns: block.columns }));
        if (!Number.isFinite(block.widthDxa) || block.widthDxa <= 0) issues.push(verificationIssue("document", "invalidTableWidth", `Table ${block.id} has an invalid width.`, { id: block.id, widthDxa: block.widthDxa }));
        if (!Number.isFinite(block.indentDxa) || block.indentDxa < 0) issues.push(verificationIssue("document", "invalidTableIndent", `Table ${block.id} has an invalid indent.`, { id: block.id, indentDxa: block.indentDxa }));
        const formattingColumns = block.cells?.length ? block.gridColumns : block.columns;
        if (!Array.isArray(block.columnWidthsDxa) || block.columnWidthsDxa.length !== formattingColumns) issues.push(verificationIssue("document", "invalidTableColumnWidths", `Table ${block.id} needs one column width per logical grid column.`, { id: block.id, columns: formattingColumns, columnWidthsDxa: block.columnWidthsDxa }));
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
    const blockIndexes = new Map(this.blocks.map((block, index) => [block.id, index]));
    const bookmarkNativeIds = new Set();
    for (const bookmark of this.bookmarks) {
      const validateEndpoint = (endpoint, end = false) => {
        const blockId = documentBookmarkEndpointBlockId(endpoint);
        const block = this.blocks.find((item) => item.id === blockId);
        if (!block) {
          issues.push(verificationIssue("document", end ? "missingBookmarkEndTarget" : "missingBookmarkTarget", `Bookmark ${bookmark.id} points at missing ${end ? "end" : "start"} block ${blockId}.`, { id: bookmark.id, targetId: endpoint?.id }));
          return false;
        }
        if (endpoint?.type === "tableCell") {
          if (block.kind !== "table" || !Number.isInteger(endpoint.row) || !Number.isInteger(endpoint.column) || endpoint.row < 0 || endpoint.column < 0 || endpoint.row >= block.rows || endpoint.column >= block.columns) {
            issues.push(verificationIssue("document", "invalidBookmarkTableCell", `Bookmark ${bookmark.id} points at invalid table cell ${endpoint.id}.`, { id: bookmark.id, targetId: endpoint.id, tableId: endpoint.tableId, row: endpoint.row, column: endpoint.column }));
            return false;
          }
        } else if (block.kind === "table") {
          issues.push(verificationIssue("document", "unsupportedBookmarkTarget", `Bookmark ${bookmark.id} must target a specific table cell, not the table block.`, { id: bookmark.id, targetId: endpoint?.id }));
          return false;
        }
        return true;
      };
      const startValid = validateEndpoint(bookmark.target, false);
      const endValid = validateEndpoint(bookmark.endTarget, true);
      if (startValid && endValid && compareDocumentBookmarkEndpoints(bookmark.target, bookmark.endTarget, blockIndexes) > 0) issues.push(verificationIssue("document", "reversedBookmarkRange", `Bookmark ${bookmark.id} ends before it starts.`, { id: bookmark.id, targetId: bookmark.targetId, endTargetId: bookmark.endTargetId }));
      if (bookmark.nativeId !== undefined) {
        if (!Number.isInteger(bookmark.nativeId) || bookmark.nativeId === -1) issues.push(verificationIssue("document", "invalidBookmarkNativeId", `Bookmark ${bookmark.id} has invalid native ID ${bookmark.nativeId}.`, { id: bookmark.id, nativeId: bookmark.nativeId }));
        else if (bookmarkNativeIds.has(bookmark.nativeId)) issues.push(verificationIssue("document", "duplicateBookmarkNativeId", `Bookmark ${bookmark.id} duplicates native ID ${bookmark.nativeId}.`, { id: bookmark.id, nativeId: bookmark.nativeId }));
        bookmarkNativeIds.add(bookmark.nativeId);
      }
    }
    const finalSectionIndex = this.blocks.filter((block) => block.kind === "section").length;
    for (const block of [...this.headers, ...this.footers]) {
      if (block.sectionIndex !== undefined && (!Number.isInteger(block.sectionIndex) || block.sectionIndex < 0 || block.sectionIndex > finalSectionIndex)) {
        issues.push(verificationIssue("document", "invalidHeaderFooterSection", `${block.kind} ${block.id} targets invalid section index ${block.sectionIndex}; expected 0 through ${finalSectionIndex}.`, { id: block.id, kind: block.kind, sectionIndex: block.sectionIndex, finalSectionIndex }));
      }
    }
    const commentParaIds = new Set();
    const commentDurableIds = new Set();
    for (const comment of this.comments) {
      if (!blockIds.has(comment.targetId)) issues.push(verificationIssue("document", "danglingComment", `Comment ${comment.id} points at a missing block.`, { id: comment.id, targetId: comment.targetId }));
      if (comment.date && Number.isNaN(Date.parse(comment.date))) issues.push(verificationIssue("document", "invalidCommentDate", `Comment ${comment.id} has an invalid date.`, { id: comment.id, date: comment.date }));
      const parent = comment.parentId ? this.comments.find((item) => item.id === comment.parentId) : undefined;
      if (comment.parentId && !parent) issues.push(verificationIssue("document", "missingCommentParent", `Comment ${comment.id} points at missing parent comment ${comment.parentId}.`, { id: comment.id, parentId: comment.parentId }));
      if (parent && parent.targetId !== comment.targetId) issues.push(verificationIssue("document", "commentParentTargetMismatch", `Comment ${comment.id} and parent ${parent.id} target different blocks.`, { id: comment.id, parentId: parent.id, targetId: comment.targetId, parentTargetId: parent.targetId }));
      if (comment.paraId && !/^[0-9A-Fa-f]{8}$/.test(comment.paraId)) issues.push(verificationIssue("document", "invalidCommentParaId", `Comment ${comment.id} has invalid paraId ${comment.paraId}.`, { id: comment.id, paraId: comment.paraId }));
      else if (comment.paraId && commentParaIds.has(comment.paraId.toUpperCase())) issues.push(verificationIssue("document", "duplicateCommentParaId", `Comment ${comment.id} duplicates paraId ${comment.paraId}.`, { id: comment.id, paraId: comment.paraId }));
      if (comment.paraId) commentParaIds.add(comment.paraId.toUpperCase());
      const durableNumber = /^[0-9A-Fa-f]{8}$/.test(comment.durableId || "") ? Number.parseInt(comment.durableId, 16) : undefined;
      if (comment.durableId && (!durableNumber || durableNumber >= 0x7FFFFFFF)) issues.push(verificationIssue("document", "invalidCommentDurableId", `Comment ${comment.id} has invalid durableId ${comment.durableId}.`, { id: comment.id, durableId: comment.durableId }));
      else if (comment.durableId && commentDurableIds.has(comment.durableId.toUpperCase())) issues.push(verificationIssue("document", "duplicateCommentDurableId", `Comment ${comment.id} duplicates durableId ${comment.durableId}.`, { id: comment.id, durableId: comment.durableId }));
      if (comment.durableId) commentDurableIds.add(comment.durableId.toUpperCase());
      if (comment.dateUtc && Number.isNaN(Date.parse(comment.dateUtc))) issues.push(verificationIssue("document", "invalidCommentDateUtc", `Comment ${comment.id} has an invalid UTC date.`, { id: comment.id, dateUtc: comment.dateUtc }));
      if (comment.parentId && comment.intelligentPlaceholder) issues.push(verificationIssue("document", "invalidCommentReplyPlaceholder", `Reply comment ${comment.id} must not be an intelligent placeholder.`, { id: comment.id, parentId: comment.parentId }));
      if (comment.person) {
        const providerId = String(comment.person.providerId || "");
        const userId = String(comment.person.userId || "");
        if (!providerId || providerId.length > 100 || !userId || userId.length > 300) issues.push(verificationIssue("document", "invalidCommentPerson", `Comment ${comment.id} has invalid person presence metadata.`, { id: comment.id, providerId, userId }));
      }
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
    const firstPageHeaderFooter = resolveDocxPageHeaderFooter(planDocxHeaderFooterSections(this), 0, 1);
    const firstPageHeaderIds = new Set(firstPageHeaderFooter.headers);
    const firstPageFooterIds = new Set(firstPageHeaderFooter.footers);
    for (const header of this.headers.filter((block) => firstPageHeaderIds.has(block.id))) {
      parts.push(`<text x="${margin}" y="${Math.max(28, y - 36)}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(header.text)}</text>`);
    }
    let listCounters = new Map();
    for (const block of this.blocks) {
      if (block.kind === "paragraph") {
        const style = this.styles.effective(block.styleId) || this.styles.get("Normal");
        const fontSize = Math.max(10, (style.fontSize || 22) / 2);
        const runs = block.runs?.length ? block.runs : [{ text: block.text, style: {} }];
        const tspans = runs.map((run, index) => {
          const runStyle = documentEffectiveRunStyle(this, block, run);
          return `<tspan${index ? "" : ` x=\"${margin}\"`} font-family="${xmlEscape(runStyle.effectiveFontFamily || runStyle.fontFamily || "Arial")}" font-size="${Math.max(5, (runStyle.effectiveFontSize || style.fontSize || 22) / 2)}" font-style="${runStyle.effectiveItalic ? "italic" : "normal"}" font-weight="${runStyle.effectiveBold ? "700" : "400"}" fill="${xmlEscape(runStyle.effectiveColor || "#111827")}">${xmlEscape(run.text)}</tspan>`;
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
        if (block.pictureBullet?.dataUrl) {
          const markerWidth = Math.max(4, Number(block.pictureBullet.widthPt) || fontSize);
          const markerHeight = Math.max(4, Number(block.pictureBullet.heightPt) || fontSize);
          parts.push(`<image data-picture-bullet="embedded" href="${attrEscape(block.pictureBullet.dataUrl)}" x="${x}" y="${y - markerHeight}" width="${markerWidth}" height="${markerHeight}" preserveAspectRatio="xMidYMid meet"/>`);
        } else if (block.pictureBullet?.uri) {
          const markerSize = Math.max(4, Number(block.pictureBullet.widthPt) || fontSize);
          parts.push(`<rect data-picture-bullet="external" data-uri="${attrEscape(block.pictureBullet.uri)}" x="${x}" y="${y - markerSize}" width="${markerSize}" height="${markerSize}" rx="2" fill="#dbeafe" stroke="#2563eb"/>`);
        } else {
          parts.push(`<text x="${x}" y="${y}" font-family="${xmlEscape(style.fontFamily || "Arial")}" font-size="${fontSize}" font-style="${style.italic ? "italic" : "normal"}" font-weight="${style.bold ? "700" : "400"}" fill="${xmlEscape(style.color || "#111827")}">${xmlEscape(marker)}</text>`);
        }
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
    for (const footer of this.footers.filter((block) => firstPageFooterIds.has(block.id))) {
      parts.push(`<text x="${margin}" y="${height - 36}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(footer.text)}</text>`);
    }
    return new FileBlob(`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${parts.join("")}</svg>`, { type: "image/svg+xml" });
  }
}

function ooxmlSafePartPath(partPath, family = "OOXML") {
  const raw = String(partPath || "").replaceAll("\\", "/").trim();
  if (!raw || raw.startsWith("/") || raw.includes("\0")) throw new Error(`Unsafe ${family} part path: ${partPath}`);
  const normalized = path.posix.normalize(raw).replace(/^\.\//, "");
  if (!normalized || normalized === "." || normalized.startsWith("../") || normalized.includes("/../") || normalized === "..") throw new Error(`Unsafe ${family} part path: ${partPath}`);
  if (normalized.length > 1024) throw new Error(`Unsafe ${family} part path: path exceeds 1024 characters`);
  return normalized;
}

function ooxmlXmlAttributes(tag = "") {
  return Object.fromEntries([...String(tag).matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])(.*?)\2/g)].map((match) => [match[1], decodeXml(match[3])]));
}

function ooxmlContentTypeMaps(xml = "") {
  const defaults = new Map();
  const overrides = new Map();
  for (const match of String(xml).matchAll(/<Default\b[^>]*\/?\s*>/g)) {
    const attrs = ooxmlXmlAttributes(match[0]);
    if (attrs.Extension && attrs.ContentType) defaults.set(String(attrs.Extension).toLowerCase(), attrs.ContentType);
  }
  for (const match of String(xml).matchAll(/<Override\b[^>]*\/?\s*>/g)) {
    const attrs = ooxmlXmlAttributes(match[0]);
    if (attrs.PartName && attrs.ContentType) overrides.set(String(attrs.PartName).replace(/^\//, ""), attrs.ContentType);
  }
  return { defaults, overrides };
}

function ooxmlFallbackContentType(partPath) {
  if (partPath.endsWith(".rels")) return "application/vnd.openxmlformats-package.relationships+xml";
  if (partPath.endsWith(".xml")) return "application/xml";
  if (partPath.endsWith(".json")) return "application/json";
  return imageContentTypeFromExtension(path.posix.extname(partPath).slice(1));
}

function ooxmlPartExtension(partPath) {
  if (partPath.endsWith(".rels")) return "rels";
  return path.posix.extname(partPath).slice(1).toLowerCase();
}

const OOXML_RELATIONSHIP_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OOXML_MICROSOFT_RELATIONSHIP_BASE = "http://schemas.microsoft.com/office/2017/10/relationships";
const OOXML_COMMON_PART_RECIPES = {
  image: { relationshipType: `${OOXML_RELATIONSHIP_BASE}/image` },
  chart: { contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/chart` },
  theme: { contentType: "application/vnd.openxmlformats-officedocument.theme+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/theme` },
  customxml: { contentType: "application/xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/customXml` },
};
const OOXML_FAMILY_PART_RECIPES = {
  DOCX: {
    header: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/header` },
    footer: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/footer` },
    comments: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/comments` },
    commentsextended: { contentType: DOCX_COMMENTS_EXTENDED_CONTENT_TYPE, relationshipType: DOCX_COMMENTS_EXTENDED_RELATIONSHIP_TYPE },
    commentsids: { contentType: DOCX_COMMENTS_IDS_CONTENT_TYPE, relationshipType: DOCX_COMMENTS_IDS_RELATIONSHIP_TYPE },
    commentsextensible: { contentType: DOCX_COMMENTS_EXTENSIBLE_CONTENT_TYPE, relationshipType: DOCX_COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE },
    people: { contentType: DOCX_PEOPLE_CONTENT_TYPE, relationshipType: DOCX_PEOPLE_RELATIONSHIP_TYPE },
    numbering: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/numbering` },
    styles: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/styles` },
    settings: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/settings` },
  },
  XLSX: {
    drawing: { contentType: "application/vnd.openxmlformats-officedocument.drawing+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/drawing` },
    worksheet: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/worksheet` },
    styles: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/styles` },
    sharedstrings: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/sharedStrings` },
    comments: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/comments` },
    table: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/table` },
    pivottable: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/pivotTable` },
    pivotcachedefinition: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/pivotCacheDefinition` },
    pivotcacherecords: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/pivotCacheRecords` },
    threadedcomments: { contentType: "application/vnd.ms-excel.threadedcomments+xml", relationshipType: `${OOXML_MICROSOFT_RELATIONSHIP_BASE}/threadedComment` },
    person: { contentType: "application/vnd.ms-excel.person+xml", relationshipType: `${OOXML_MICROSOFT_RELATIONSHIP_BASE}/person` },
  },
  PPTX: {
    slide: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/slide` },
    slidelayout: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/slideLayout` },
    slidemaster: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/slideMaster` },
    notesslide: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/notesSlide` },
    comments: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.comments+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/comments` },
    commentauthors: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/commentAuthors` },
    moderncomments: { contentType: PPTX_MODERN_COMMENT_CONTENT_TYPE, relationshipType: PPTX_MODERN_COMMENT_RELATIONSHIP_TYPE },
    modernauthors: { contentType: PPTX_MODERN_AUTHOR_CONTENT_TYPE, relationshipType: PPTX_MODERN_AUTHOR_RELATIONSHIP_TYPE },
  },
};

function ooxmlRecipeKind(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function applyOoxmlPartRecipe(patch, family) {
  const rawRecipe = patch.recipe ?? patch.partRecipe ?? patch.partKind;
  if (rawRecipe == null) return patch;
  const recipe = typeof rawRecipe === "object" ? rawRecipe : { kind: rawRecipe };
  const kind = ooxmlRecipeKind(recipe.kind || recipe.name || recipe.type);
  const spec = OOXML_FAMILY_PART_RECIPES[family]?.[kind] || OOXML_COMMON_PART_RECIPES[kind];
  if (!spec) {
    const supported = [...Object.keys(OOXML_COMMON_PART_RECIPES), ...Object.keys(OOXML_FAMILY_PART_RECIPES[family] || {})].sort().join(", ");
    throw new Error(`${family} OOXML part recipe ${recipe.kind || recipe.name || recipe.type || "(missing)"} is unsupported. Supported recipes: ${supported}.`);
  }
  const derivedRelationship = recipe.relationship === false ? undefined : {
    source: recipe.source ?? recipe.sourcePart,
    id: recipe.id ?? recipe.relationshipId,
    type: spec.relationshipType,
    target: recipe.target,
    targetMode: recipe.targetMode,
  };
  let relationship = patch.relationship;
  let relationships = patch.relationships;
  if (relationship) relationship = { ...derivedRelationship, ...relationship, type: relationship.type || spec.relationshipType };
  else if (relationships) relationships = relationships.map((item) => ({ ...derivedRelationship, ...item, type: item.type || spec.relationshipType }));
  else if (derivedRelationship?.source !== undefined) relationship = derivedRelationship;
  return { ...patch, contentType: patch.contentType || patch.mimeType || patch.type || spec.contentType, relationship, relationships, sourceReference: patch.sourceReference ?? recipe.sourceReference, recipeKind: kind };
}

function ooxmlRelationshipSource(partPath) {
  if (partPath === "_rels/.rels") return "";
  const match = /^(?:(.*)\/)?_rels\/([^/]+)\.rels$/.exec(partPath);
  if (!match) return undefined;
  return match[1] ? `${match[1]}/${match[2]}` : match[2];
}

function ooxmlResolveRelationshipTarget(source, rawTarget) {
  const target = String(rawTarget || "").split("#")[0];
  if (target.startsWith("/")) return target.slice(1);
  const sourceDir = source ? path.posix.dirname(source) : "";
  return path.posix.normalize(path.posix.join(sourceDir === "." ? "" : sourceDir, target));
}

const OOXML_RELATIONSHIP_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);

function ooxmlRelationshipReferences(xml = "") {
  const source = String(xml || "");
  const tags = [...source.matchAll(/<[^!?][^>]*>/g)].map((match) => match[0]);
  const references = [];
  const namespaceStack = [new Map()];
  for (const tag of tags) {
    if (/^<\//.test(tag)) { if (namespaceStack.length > 1) namespaceStack.pop(); continue; }
    const namespaces = new Map(namespaceStack.at(-1));
    for (const match of tag.matchAll(/\bxmlns:([A-Za-z_][\w.-]*)\s*=\s*(["'])(.*?)\2/g)) namespaces.set(match[1], decodeXml(match[3]));
    for (const [prefix, namespace] of namespaces) {
      if (!OOXML_RELATIONSHIP_NAMESPACES.has(namespace)) continue;
      const escapedPrefix = prefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const pattern = new RegExp(`\\b${escapedPrefix}:(id|embed|link|dm|lo|qs|cs)\\s*=\\s*(["'])(.*?)\\2`, "g");
      for (const match of tag.matchAll(pattern)) references.push({ attribute: `${prefix}:${match[1]}`, id: decodeXml(match[3]) });
    }
    if (!/\/\s*>$/.test(tag)) namespaceStack.push(namespaces);
  }
  return references.filter((reference) => reference.id);
}

function ooxmlPackageIssues(files, bytesByPath, contentTypes, family) {
  const paths = new Set(files.map((file) => file.name));
  const issues = [];
  if (!paths.has("[Content_Types].xml")) issues.push({ kind: "ooxmlIssue", family, type: "missingContentTypes", severity: "error", message: `${family} package is missing [Content_Types].xml.` });
  if (!paths.has("_rels/.rels")) issues.push({ kind: "ooxmlIssue", family, type: "missingRootRelationships", severity: "error", message: `${family} package is missing _rels/.rels.` });
  for (const partPath of contentTypes.overrides.keys()) {
    if (!paths.has(partPath)) issues.push({ kind: "ooxmlIssue", family, type: "contentTypeTargetNotFound", severity: "error", path: partPath, message: `${family} [Content_Types].xml declares missing part ${partPath}.` });
  }
  for (const partPath of paths) {
    if (partPath === "[Content_Types].xml") continue;
    const extension = ooxmlPartExtension(partPath);
    if (!contentTypes.overrides.has(partPath) && !contentTypes.defaults.has(extension)) {
      issues.push({ kind: "ooxmlIssue", family, type: "missingContentType", severity: "error", path: partPath, message: `${family} part ${partPath} has no [Content_Types].xml declaration.` });
    }
  }
  const relationshipsBySource = new Map();
  for (const [partPath, bytes] of bytesByPath) {
    if (!partPath.endsWith(".rels")) continue;
    const source = ooxmlRelationshipSource(partPath);
    if (source == null) continue;
    const xml = decoder.decode(bytes);
    const relationshipEntries = ooxmlRelationshipEntries(xml);
    relationshipsBySource.set(source, { path: partPath, ids: new Set(relationshipEntries.map((entry) => entry.attrs.Id).filter(Boolean)) });
    if (source && !paths.has(source)) issues.push({ kind: "ooxmlIssue", family, type: "relationshipSourceNotFound", severity: "error", path: partPath, source, message: `${family} relationship part ${partPath} belongs to missing source part ${source}.` });
    const relationshipIds = new Set();
    for (const { attrs } of relationshipEntries) {
      if (attrs.Id && relationshipIds.has(attrs.Id)) issues.push({ kind: "ooxmlIssue", family, type: "duplicateRelationshipId", severity: "error", path: partPath, relationshipId: attrs.Id, message: `${family} relationship part ${partPath} contains duplicate Id ${attrs.Id}.` });
      if (attrs.Id) relationshipIds.add(attrs.Id);
      if (String(attrs.TargetMode || "").toLowerCase() === "external") continue;
      if (!attrs.Target) {
        issues.push({ kind: "ooxmlIssue", family, type: "relationshipTargetMissing", severity: "error", path: partPath, relationshipId: attrs.Id, message: `${family} relationship ${attrs.Id || "(unknown)"} in ${partPath} has no target.` });
        continue;
      }
      const target = ooxmlResolveRelationshipTarget(source, attrs.Target);
      if (!paths.has(target)) issues.push({ kind: "ooxmlIssue", family, type: "relationshipTargetNotFound", severity: "error", path: partPath, relationshipId: attrs.Id, target, message: `${family} relationship ${attrs.Id || "(unknown)"} in ${partPath} targets missing part ${target}.` });
    }
  }
  for (const [partPath, bytes] of bytesByPath) {
    if (!partPath.endsWith(".xml") || partPath === "[Content_Types].xml") continue;
    const references = ooxmlRelationshipReferences(decoder.decode(bytes));
    if (!references.length) continue;
    const relationshipPart = relationshipsBySource.get(partPath);
    if (!relationshipPart) {
      issues.push({ kind: "ooxmlIssue", family, type: "relationshipReferencePartNotFound", severity: "error", path: partPath, relationshipIds: [...new Set(references.map((reference) => reference.id))], message: `${family} source part ${partPath} contains relationship references but has no corresponding .rels part.` });
      continue;
    }
    const seen = new Set();
    for (const reference of references) {
      const key = `${reference.attribute}\u0000${reference.id}`;
      if (seen.has(key) || relationshipPart.ids.has(reference.id)) continue;
      seen.add(key);
      issues.push({ kind: "ooxmlIssue", family, type: "relationshipReferenceIdNotFound", severity: "error", path: partPath, relationshipsPath: relationshipPart.path, relationshipId: reference.id, referenceAttribute: reference.attribute, message: `${family} source part ${partPath} references missing relationship Id ${reference.id} through ${reference.attribute}.` });
    }
  }
  return issues;
}

async function ooxmlPackageRecords(zip, options = {}, config = {}) {
  const includeText = Boolean(options.includeText || options.preview || options.includeXml);
  const maxPreviewChars = Math.max(0, Number(options.maxPreviewChars ?? 400) || 0);
  const files = Object.values(zip.files).filter((file) => !file.dir).sort((a, b) => a.name.localeCompare(b.name));
  const family = config.family || "OOXML";
  const maxParts = Math.max(1, Number(options.maxParts ?? 5000) || 5000);
  const maxPartBytes = Math.max(1, Number(options.maxPartBytes ?? 64 * 1024 * 1024) || 64 * 1024 * 1024);
  const maxTotalBytes = Math.max(1, Number(options.maxTotalBytes ?? 256 * 1024 * 1024) || 256 * 1024 * 1024);
  if (files.length > maxParts) throw new Error(`${family} package has ${files.length} parts; maxParts is ${maxParts}.`);
  const safePaths = new Map();
  let declaredTotalBytes = 0;
  for (const file of files) {
    const partPath = ooxmlSafePartPath(file.name, family);
    safePaths.set(file, partPath);
    const declaredSize = Number(file._data?.uncompressedSize);
    if (!Number.isFinite(declaredSize)) continue;
    if (declaredSize > maxPartBytes) throw new Error(`${family} part ${partPath} exceeds maxPartBytes (${maxPartBytes}).`);
    declaredTotalBytes += declaredSize;
    if (declaredTotalBytes > maxTotalBytes) throw new Error(`${family} package exceeds maxTotalBytes (${maxTotalBytes}).`);
  }
  const contentTypesEntry = zip.file("[Content_Types].xml");
  const contentTypesText = contentTypesEntry ? await contentTypesEntry.async("text") : "";
  const contentTypes = ooxmlContentTypeMaps(contentTypesText);
  const counts = Object.fromEntries(Object.entries(config.counts || {}).map(([name, pattern]) => [name, files.filter((file) => pattern.test(file.name)).length]));
  const records = [{ kind: config.packageKind || "ooxmlPackage", family, parts: files.length, ...counts }];
  const bytesByPath = new Map();
  let totalBytes = 0;
  for (const file of files) {
    const partPath = safePaths.get(file);
    const bytes = await file.async("uint8array");
    if (bytes.byteLength > maxPartBytes) throw new Error(`${family} part ${partPath} exceeds maxPartBytes (${maxPartBytes}).`);
    totalBytes += bytes.byteLength;
    if (totalBytes > maxTotalBytes) throw new Error(`${family} package exceeds maxTotalBytes (${maxTotalBytes}).`);
    const extension = ooxmlPartExtension(partPath);
    const contentType = contentTypes.overrides.get(partPath) || contentTypes.defaults.get(extension) || ooxmlFallbackContentType(partPath);
    const record = { kind: config.partKind || "ooxmlPart", path: partPath, size: bytes.byteLength, contentType };
    if (includeText && /\.(xml|json|rels)$/i.test(partPath)) record.textPreview = decoder.decode(bytes).slice(0, maxPreviewChars);
    records.push(record);
    bytesByPath.set(partPath, bytes);
  }
  const issues = ooxmlPackageIssues(files, bytesByPath, contentTypes, family);
  const semanticIssues = config.semanticIssues ? await config.semanticIssues({ bytesByPath, contentTypes, family }) : [];
  issues.push(...semanticIssues);
  records[0].uncompressedBytes = totalBytes;
  records[0].relationshipReferences = [...bytesByPath].reduce((count, [partPath, bytes]) => count + (partPath.endsWith(".xml") && partPath !== "[Content_Types].xml" ? ooxmlRelationshipReferences(decoder.decode(bytes)).length : 0), 0);
  records[0].relationshipReferenceIssues = issues.filter((issue) => issue.type === "relationshipReferencePartNotFound" || issue.type === "relationshipReferenceIdNotFound").length;
  if (config.semanticIssues) {
    records[0].semanticValidation = true;
    records[0].semanticIssues = semanticIssues.length;
  }
  records[0].ok = issues.length === 0;
  records[0].issues = issues.length;
  records.push(...issues);
  return records;
}

function ooxmlPatchData(patch, family = "OOXML") {
  if (patch.json !== undefined) return encoder.encode(JSON.stringify(patch.json, null, 2));
  if (patch.text !== undefined || patch.xml !== undefined) return encoder.encode(String(patch.text ?? patch.xml));
  if (patch.bytes !== undefined || patch.data !== undefined || patch.buffer !== undefined) return toUint8Array(patch.bytes ?? patch.data ?? patch.buffer);
  if (patch.content !== undefined) return typeof patch.content === "string" ? encoder.encode(patch.content) : toUint8Array(patch.content);
  if (patch.remove || patch.delete) return undefined;
  throw new Error(`${family} patch for ${patch.path || patch.part || "unknown part"} has no content or remove flag.`);
}

function ooxmlRemoveContentTypeOverride(xml, partPath) {
  return String(xml).replace(/<Override\b[^>]*\/?\s*>/g, (tag) => {
    const partName = ooxmlXmlAttributes(tag).PartName;
    return String(partName || "").replace(/^\//, "") === partPath ? "" : tag;
  });
}

function ooxmlInferredPatchContentType(patch, partPath) {
  const explicit = patch.contentType || patch.mimeType || patch.type;
  if (explicit) return String(explicit);
  const extension = ooxmlPartExtension(partPath);
  if (patch.json !== undefined || extension === "json") return "application/json";
  if (patch.xml !== undefined || extension === "xml" || extension === "rels") return extension === "rels" ? "application/vnd.openxmlformats-package.relationships+xml" : "application/xml";
  if (patch.text !== undefined || extension === "txt") return "text/plain";
  const imageType = imageContentTypeFromExtension(extension);
  return imageType === "application/octet-stream" ? undefined : imageType;
}

async function syncOoxmlPatchContentTypes(zip, normalizedPatches, options, family) {
  if (options.syncContentTypes === false) return 0;
  const entry = zip.file("[Content_Types].xml");
  if (!entry) throw new Error(`${family} package is missing [Content_Types].xml; cannot synchronize patch content types.`);
  let xml = await entry.async("text");
  let updates = 0;
  for (const { patch, partPath } of normalizedPatches) {
    if (partPath === "[Content_Types].xml") continue;
    const withoutOverride = ooxmlRemoveContentTypeOverride(xml, partPath);
    const removedOverride = withoutOverride !== xml;
    if (patch.remove || patch.delete) {
      if (removedOverride) { xml = withoutOverride; updates += 1; }
      continue;
    }
    const requestedType = ooxmlInferredPatchContentType(patch, partPath);
    const declarations = ooxmlContentTypeMaps(xml);
    const extension = ooxmlPartExtension(partPath);
    const existingType = declarations.overrides.get(partPath) || declarations.defaults.get(extension);
    if (!requestedType || (existingType && !patch.contentType && !patch.mimeType && !patch.type)) continue;
    if (declarations.defaults.get(extension) === requestedType && !patch.contentType && !patch.mimeType && !patch.type) {
      if (removedOverride) { xml = withoutOverride; updates += 1; }
      continue;
    }
    xml = withoutOverride.replace(/<\/Types>\s*$/, `<Override PartName="/${attrEscape(partPath)}" ContentType="${attrEscape(requestedType)}"/></Types>`);
    updates += 1;
  }
  if (updates) zip.file("[Content_Types].xml", xml);
  return updates;
}

function ooxmlRelationshipPartPath(source, family) {
  if (!source) return "_rels/.rels";
  const safeSource = ooxmlSafePartPath(source, family);
  const dir = path.posix.dirname(safeSource);
  return `${dir === "." ? "" : `${dir}/`}_rels/${path.posix.basename(safeSource)}.rels`;
}

function ooxmlRelationshipEntries(xml = "") {
  return [...String(xml).matchAll(/<Relationship\b[^>]*\/?\s*>/g)].map((match) => ({
    tag: match[0],
    attrs: ooxmlXmlAttributes(match[0]),
  }));
}

function ooxmlEmptyRelationshipsXml() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
}

async function syncOoxmlPatchRelationships(zip, normalizedPatches, options, family) {
  if (options.syncRelationships === false) return 0;
  let updates = 0;
  const removedParts = new Set(normalizedPatches.filter(({ patch }) => patch.remove || patch.delete).map(({ partPath }) => partPath));
  const removedRelationshipIds = new Map();
  for (const removedPart of removedParts) {
    const outgoingRelationships = ooxmlRelationshipPartPath(removedPart, family);
    if (zip.file(outgoingRelationships)) { zip.remove(outgoingRelationships); updates += 1; }
  }
  for (const file of Object.values(zip.files).filter((item) => !item.dir && item.name.endsWith(".rels"))) {
    const relsPath = ooxmlSafePartPath(file.name, family);
    const source = ooxmlRelationshipSource(relsPath);
    if (source == null) continue;
    const xml = await file.async("text");
    const next = xml.replace(/<Relationship\b[^>]*\/?\s*>/g, (tag) => {
      const entry = ooxmlRelationshipEntries(tag)[0];
      if (!entry || String(entry.attrs.TargetMode || "").toLowerCase() === "external") return tag;
      const resolvedTarget = ooxmlResolveRelationshipTarget(source, entry.attrs.Target);
      if (!removedParts.has(resolvedTarget)) return tag;
      const key = `${source}\u0000${resolvedTarget}`;
      if (entry.attrs.Id) removedRelationshipIds.set(key, [...(removedRelationshipIds.get(key) || []), entry.attrs.Id]);
      return "";
    });
    if (next !== xml) { zip.file(relsPath, next); updates += 1; }
  }
  for (const { patch, partPath } of normalizedPatches) {
    const relationshipConfigs = patch.relationships || (patch.relationship ? [patch.relationship] : []);
    for (const relationship of relationshipConfigs) {
      const source = relationship.source || relationship.sourcePart || "";
      if (source && !zip.file(ooxmlSafePartPath(source, family))) throw new Error(`${family} relationship source part not found: ${source}`);
      const relsPath = ooxmlRelationshipPartPath(source, family);
      const relsEntry = zip.file(relsPath);
      let xml = relsEntry ? await relsEntry.async("text") : ooxmlEmptyRelationshipsXml();
      const entries = ooxmlRelationshipEntries(xml);
      const target = relationship.target || (source ? path.posix.relative(path.posix.dirname(source), partPath) : partPath);
      const remove = relationship.remove === true || relationship.delete === true || patch.remove || patch.delete;
      const matches = (entry) => (relationship.id && entry.attrs.Id === relationship.id) || (!relationship.id && ooxmlResolveRelationshipTarget(source, entry.attrs.Target) === partPath);
      const existingById = relationship.id ? entries.find((entry) => entry.attrs.Id === relationship.id) : undefined;
      if (!remove && existingById && ooxmlResolveRelationshipTarget(source, existingById.attrs.Target) !== partPath && relationship.replaceExisting !== true && relationship.replace !== true) throw new Error(`${family} relationship Id ${relationship.id} in ${relsPath} already targets ${existingById.attrs.Target}; pass replaceExisting:true only for an intentional rebind.`);
      relationship.resolvedIds = [...new Set([...entries.filter(matches).map((entry) => entry.attrs.Id).filter(Boolean), ...(removedRelationshipIds.get(`${source}\u0000${partPath}`) || [])])];
      const withoutMatch = xml.replace(/<Relationship\b[^>]*\/?\s*>/g, (tag) => {
        const entry = ooxmlRelationshipEntries(tag)[0];
        return entry && matches(entry) ? "" : tag;
      });
      if (remove) {
        if (withoutMatch !== xml) { zip.file(relsPath, withoutMatch); updates += 1; }
        continue;
      }
      if (!relationship.type) throw new Error(`${family} relationship for ${partPath} requires type.`);
      const usedIds = new Set(entries.map((entry) => entry.attrs.Id));
      let id = relationship.id;
      if (!id) { let index = 1; while (usedIds.has(`rId${index}`)) index += 1; id = `rId${index}`; }
      relationship.resolvedId = id;
      const targetMode = relationship.targetMode ? ` TargetMode="${attrEscape(relationship.targetMode)}"` : "";
      const tag = `<Relationship Id="${attrEscape(id)}" Type="${attrEscape(relationship.type)}" Target="${attrEscape(target)}"${targetMode}/>`;
      const next = withoutMatch.replace(/<\/Relationships>\s*$/, `${tag}</Relationships>`);
      zip.file(relsPath, next);
      updates += 1;
    }
  }
  return updates;
}

function ooxmlTagRelationshipId(tag = "") {
  return Object.entries(ooxmlXmlAttributes(tag)).find(([name]) => /:(?:id|embed|link)$/.test(name))?.[1];
}

function ooxmlReferenceConfig(value) {
  if (value === true) return {};
  if (value && typeof value === "object") return value;
  return undefined;
}

async function syncOoxmlSourceReferences(zip, normalizedPatches, options, family) {
  if (options.syncSourceReferences === false) return 0;
  let updates = 0;
  for (const { patch, partPath } of normalizedPatches) {
    const config = ooxmlReferenceConfig(patch.sourceReference);
    if (!config) continue;
    if (!supportsOoxmlSourceReference(family, patch.recipeKind)) throw new Error(`${family} sourceReference is not supported for recipe ${patch.recipeKind || "(missing)"}. Supported recipes: ${supportedOoxmlSourceReferenceSummary()}.`);
    const relationship = patch.relationship || patch.relationships?.[0];
    const source = relationship?.source || relationship?.sourcePart;
    if (!source) throw new Error(`${family} ${patch.recipeKind} sourceReference requires recipe.source.`);
    const safeSource = ooxmlSafePartPath(source, family);
    const sourceEntry = zip.file(safeSource);
    if (!sourceEntry) throw new Error(`${family} sourceReference source part not found: ${safeSource}`);
    const remove = patch.remove === true || patch.delete === true;
    const resolvedIds = new Set([...(relationship.resolvedIds || []), relationship.id].filter(Boolean));
    const addId = remove ? undefined : relationship.resolvedId || relationship.id;
    if (!remove && !addId) throw new Error(`${family} ${patch.recipeKind} sourceReference could not resolve a relationship Id.`);
    const validatesTarget = (family === "DOCX" && ["comments", "numbering", "settings"].includes(patch.recipeKind)) || (family === "PPTX" && patch.recipeKind === "chart");
    if (!remove && validatesTarget) {
      const targetEntry = zip.file(partPath);
      if (!targetEntry) throw new Error(`${family} sourceReference target part not found: ${partPath}`);
      const targetXml = await targetEntry.async("text");
      if (family === "DOCX" && patch.recipeKind === "settings") {
        const nextTarget = mutateOoxmlSourceReferenceTarget({ family, recipeKind: patch.recipeKind, targetXml, config });
        if (nextTarget !== targetXml) { zip.file(partPath, nextTarget); updates += 1; }
        continue;
      }
      validateOoxmlSourceReferenceTarget({ family, recipeKind: patch.recipeKind, targetXml, config });
    }
    const xml = await sourceEntry.async("text");
    const next = mutateOoxmlSourceReference({ family, recipeKind: patch.recipeKind, xml, relationshipIds: resolvedIds, addId, config });
    if (next !== xml) { zip.file(safeSource, next); updates += 1; }
  }
  return updates;
}

async function inspectOoxmlPackage(blobOrBuffer, options = {}, config = {}) {
  const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
  const zip = await JSZip.loadAsync(bytes);
  const records = await ooxmlPackageRecords(zip, options, config);
  const partKind = config.partKind || "ooxmlPart";
  return { ok: records[0].ok, issues: records.filter((record) => record.kind === "ooxmlIssue"), parts: records.filter((record) => record.kind === partKind), records, ...ndjson(records, options.maxChars ?? Infinity) };
}

async function patchOoxmlPackage(blobOrBuffer, patches = [], options = {}, config = {}) {
  const family = config.family || "OOXML";
  const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
  const zip = await JSZip.loadAsync(bytes);
  const list = Array.isArray(patches) ? patches : Object.entries(patches || {}).map(([partPath, content]) => (
    content && typeof content === "object" && !(content instanceof Uint8Array) && !(content instanceof ArrayBuffer) && !ArrayBuffer.isView(content)
      ? { path: partPath, ...content }
      : { path: partPath, content }
  ));
  const preparedList = list.map((patch) => applyOoxmlPartRecipe(patch, family));
  const maxPatchBytes = Math.max(1, Number(options.maxPatchBytes ?? 5 * 1024 * 1024) || 5 * 1024 * 1024);
  const maxParts = Math.max(1, Number(options.maxParts ?? 5000) || 5000);
  const existingParts = new Set(Object.values(zip.files).filter((file) => !file.dir).map((file) => ooxmlSafePartPath(file.name, family)));
  const normalizedPatches = preparedList.map((patch) => ({ patch, partPath: ooxmlSafePartPath(patch.path || patch.part || patch.name, family) }));
  const resultingParts = new Set(existingParts);
  for (const { patch, partPath } of normalizedPatches) {
    if (patch.remove || patch.delete) resultingParts.delete(partPath);
    else resultingParts.add(partPath);
  }
  if (resultingParts.size > maxParts) throw new Error(`${family} patch would create ${resultingParts.size} parts; maxParts is ${maxParts}.`);
  for (const { patch, partPath } of normalizedPatches) {
    if (patch.remove || patch.delete) { zip.remove(partPath); continue; }
    const data = ooxmlPatchData(patch, family);
    if (data?.byteLength > maxPatchBytes) throw new Error(`${family} patch for ${partPath} exceeds maxPatchBytes (${maxPatchBytes}).`);
    zip.file(partPath, data);
  }
  const contentTypesUpdated = await syncOoxmlPatchContentTypes(zip, normalizedPatches, options, family);
  const relationshipsUpdated = await syncOoxmlPatchRelationships(zip, normalizedPatches, options, family);
  const sourceReferencesUpdated = await syncOoxmlSourceReferences(zip, normalizedPatches, options, family);
  const validated = options.validateResult !== false;
  let validationIssues = [];
  if (validated) {
    const records = await ooxmlPackageRecords(zip, { maxParts, maxPartBytes: options.maxPartBytes, maxTotalBytes: options.maxTotalBytes }, { ...config, family });
    validationIssues = records.filter((record) => record.kind === "ooxmlIssue");
    if (validationIssues.length) {
      const summary = validationIssues.slice(0, 5).map((issue) => `${issue.type}${issue.path ? `:${issue.path}` : ""}`).join(", ");
      throw new Error(`${family} patch produced an invalid OOXML package (${validationIssues.length} issue${validationIssues.length === 1 ? "" : "s"}): ${summary}. Pass { validateResult: false } only when intentionally constructing an invalid fixture.`);
    }
  }
  return { bytes: await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), patchedParts: preparedList.length, recipesApplied: preparedList.filter((patch) => patch.recipeKind).length, contentTypesUpdated, relationshipsUpdated, sourceReferencesUpdated, validated, validationIssues: validationIssues.length };
}

export class DocumentFile {
  static async inspectDocx(blobOrBuffer, options = {}) {
    return inspectOoxmlPackage(blobOrBuffer, options, DOCX_PACKAGE_CONFIG);
  }

  static async patchDocx(blobOrBuffer, patches = [], options = {}) {
    const patched = await patchOoxmlPackage(blobOrBuffer, patches, options, DOCX_PACKAGE_CONFIG);
    return new FileBlob(patched.bytes, { type: DOCX_MIME, metadata: { artifactKind: "document", patchedParts: patched.patchedParts, recipesApplied: patched.recipesApplied, contentTypesUpdated: patched.contentTypesUpdated, relationshipsUpdated: patched.relationshipsUpdated, sourceReferencesUpdated: patched.sourceReferencesUpdated, validated: patched.validated, validationIssues: patched.validationIssues } });
  }

  static async exportDocx(document, options = {}) {
    const { exportDocxWithOpenChestnut } = await import("./codecs/open-chestnut.mjs");
    return exportDocxWithOpenChestnut(document, options);
  }

  static async importDocx(blobOrBuffer, options = {}) {
    const { importDocxWithOpenChestnut } = await import("./codecs/open-chestnut.mjs");
    return importDocxWithOpenChestnut(blobOrBuffer, options);
  }
}
