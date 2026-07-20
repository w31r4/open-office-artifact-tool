import { create, fromBinary, toBinary } from "@bufbuild/protobuf";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { Workbook } from "../spreadsheet/index.mjs";
import { DocumentModel } from "../document/index.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { isXmlSafeText } from "../shared/xml.mjs";
import { XLSX_THEME_COLOR_NAMES, normalizeXlsxStyle, normalizeXlsxThemeConfig } from "../spreadsheet/ooxml-styles.mjs";
import { deterministicSpreadsheetGuid } from "../spreadsheet/ooxml-threaded-comments.mjs";
import { normalizeDataBarConfig, normalizeIconSetConfig } from "../spreadsheet/conditional-formats.mjs";
import {
  ArtifactFamily,
  CellFormulaKind,
  CodecOperation,
  CodecRequestSchema,
  CodecResponseSchema,
  DocumentChangeType,
  DocumentHeaderFooterReference,
  DocumentNoteKind,
  DocumentSectionBreak,
  DocumentStyleType,
  DocumentTableVerticalMerge,
  SpreadsheetCalculationMode,
  SpreadsheetWorksheetVisibility,
  WorkbookDateSystem,
} from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";
import { presentationEnvelope, presentationFromEnvelope } from "./open-chestnut-presentation.mjs";
import { spreadsheetChartFromWire, spreadsheetChartSnapshot, wireWorksheetCharts } from "./open-chestnut-spreadsheet-charts.mjs";
import { spreadsheetImageFromWire, spreadsheetImageSnapshot, wireWorksheetImages } from "./open-chestnut-spreadsheet-images.mjs";
import { spreadsheetSparklineFromWire, spreadsheetSparklineSnapshot, wireWorksheetSparklines } from "./open-chestnut-spreadsheet-sparklines.mjs";
import { hydrateWorksheetDataTable, wireWorksheetDataTables } from "./open-chestnut-spreadsheet-data-tables.mjs";
import { hydrateWorkbookPivots, wireWorksheetPivots } from "./open-chestnut-spreadsheet-pivots.mjs";

export { OpenChestnutCodecError } from "./open-chestnut-error.mjs";

export const OPEN_CHESTNUT_PROTOCOL_VERSION = 2;

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const RUNTIME_URL = new URL("../../runtime/open-chestnut/main.mjs", import.meta.url);
const MANIFEST_URL = new URL("../../runtime/open-chestnut/manifest.json", import.meta.url);
const WORKBOOK_STATE = Symbol.for("open-office-artifact-tool.open-chestnut-state");
const TABLE_STATE = Symbol.for("open-office-artifact-tool.open-chestnut-table-state");
const DOCUMENT_STATE = Symbol.for("open-office-artifact-tool.open-chestnut-document-state");

function assertTrustedImportedState(state, family) {
  if (!state) return;
  const sourceHash = String(state.source?.packageSha256 || "").toLowerCase();
  const snapshot = state.opaqueOpc?.sourcePackage;
  const snapshotHash = String(snapshot?.sha256 || "").toLowerCase();
  if (!sourceHash || !snapshotHash || sourceHash !== snapshotHash || !snapshot?.data?.length) {
    throw new OpenChestnutCodecError(`${family} source-bound export requires its validated source package snapshot.`, [], { code: "missing_source_package" });
  }
}
const MAX_XLSX_NUMBER_FORMAT_CODE_LENGTH = 4096;
const MAX_XLSX_FORMULA_LENGTH = 8192;
const MAX_XLSX_FORMULA_TOPOLOGY_CELLS = 1_048_576;
const XLSX_FORMULA_METADATA_KEYS = new Set([
  "formulaType", "sharedIndex", "sharedRef", "arrayRef", "dynamicArrayRef",
  "spillParent", "spillAnchor", "spillRange", "spillValues", "spillError",
]);
const XLSX_NUMBER_FORMAT_STYLE_KEYS = new Set(["numberFormat", "numFmt"]);
const EXCEL_ERRORS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#UNKNOWN!", "#CONNECT!", "#CYCLE!"]);
const XLSX_THEME_WIRE_FIELDS = [
  ["dk1", "dk1Rgb"], ["lt1", "lt1Rgb"], ["dk2", "dk2Rgb"], ["lt2", "lt2Rgb"],
  ["accent1", "accent1Rgb"], ["accent2", "accent2Rgb"], ["accent3", "accent3Rgb"],
  ["accent4", "accent4Rgb"], ["accent5", "accent5Rgb"], ["accent6", "accent6Rgb"],
  ["hlink", "hlinkRgb"], ["folHlink", "folHlinkRgb"],
];

let runtimePromise;

async function runtime() {
  if (!runtimePromise) {
    runtimePromise = import(RUNTIME_URL.href)
      .then((module) => module.loadOpenChestnut())
      .catch((error) => {
        runtimePromise = undefined;
        throw new OpenChestnutCodecError("Bundled OpenChestnut runtime could not be loaded.", [], { code: "runtime_unavailable", cause: error });
      });
  }
  return runtimePromise;
}

function uint64(value, name) {
  if (value == null) return 0n;
  if (typeof value === "bigint") {
    if (value < 0n) throw new TypeError(`${name} must be a non-negative integer.`);
    return value;
  }
  if (!Number.isSafeInteger(value) || value < 0) throw new TypeError(`${name} must be a non-negative safe integer or bigint.`);
  return BigInt(value);
}

function uint32(value, name) {
  if (value == null) return 0;
  if (!Number.isInteger(value) || value < 0 || value > 0xffff_ffff) throw new TypeError(`${name} must be an unsigned 32-bit integer.`);
  return value;
}

function codecLimits(limits = {}) {
  return {
    maxInputBytes: uint64(limits.maxInputBytes, "maxInputBytes"),
    maxUncompressedBytes: uint64(limits.maxUncompressedBytes, "maxUncompressedBytes"),
    maxParts: uint32(limits.maxParts, "maxParts"),
    maxSheets: uint32(limits.maxSheets, "maxSheets"),
    maxCells: uint64(limits.maxCells, "maxCells"),
    maxCompressionRatio: uint32(limits.maxCompressionRatio, "maxCompressionRatio"),
  };
}

function publicWorksheetVisibility(value) {
  if (value === SpreadsheetWorksheetVisibility.HIDDEN) return "hidden";
  if (value === SpreadsheetWorksheetVisibility.VERY_HIDDEN) return "veryHidden";
  return "visible";
}

function wireWorksheetVisibility(value) {
  if (value === "hidden") return SpreadsheetWorksheetVisibility.HIDDEN;
  if (value === "veryHidden") return SpreadsheetWorksheetVisibility.VERY_HIDDEN;
  if (value === "visible") return SpreadsheetWorksheetVisibility.VISIBLE;
  throw new OpenChestnutCodecError(`Unsupported worksheet visibility ${value}; expected visible, hidden, or veryHidden.`, [], { code: "invalid_worksheet_visibility" });
}

function worksheetMetadataSnapshot(sheet) {
  return { name: sheet.name, visibility: sheet.visibility };
}

function wireWorksheetMetadata(sheet, slot) {
  const unchanged = slot && JSON.stringify(worksheetMetadataSnapshot(sheet)) === JSON.stringify(slot.publicSnapshot);
  return {
    visibility: unchanged ? slot.wire.visibility : wireWorksheetVisibility(sheet.visibility),
    source: slot?.wire.source,
  };
}

function workbookViewSnapshots(workbook) {
  return workbook.windows.items.map((window) => ({
    activeWorksheetId: window.getActiveWorksheet().id,
    selectedWorksheetIds: window.getSelectedWorksheets().map((sheet) => sheet.id),
  }));
}

function wireWorkbookViews(workbook, state) {
  const slots = state?.viewSlots || [];
  if (state && slots.length === 0 && !workbook._activeWorksheetId && workbook.windows.count === 1)
    return { view: undefined, additionalViews: [] };
  const snapshots = workbookViewSnapshots(workbook);
  const views = snapshots.map((snapshot, index) => {
    const slot = slots[index];
    if (slot && JSON.stringify(snapshot) === JSON.stringify(slot.publicSnapshot)) return slot.wire;
    return { ...snapshot, source: slot?.wire.source };
  });
  return { view: views[0], additionalViews: views.slice(1) };
}

function bytesFrom(value) {
  if (value instanceof Uint8Array) return value;
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  throw new TypeError("Expected FileBlob, Uint8Array, ArrayBuffer, or ArrayBuffer view.");
}

async function inputBytes(value) {
  if (value instanceof FileBlob) return new Uint8Array(await value.arrayBuffer());
  return bytesFrom(value);
}

function responseFailure(response) {
  const message = response.diagnostics.length
    ? response.diagnostics.map((item) => `${item.code}: ${item.message}`).join("\n")
    : "OpenChestnut codec returned an unspecified failure.";
  return new OpenChestnutCodecError(message, response.diagnostics);
}

export async function invokeOpenChestnut(request) {
  if (Object.hasOwn(request || {}, "allowLossy") || Object.hasOwn(request || {}, "allow_lossy")) {
    throw new TypeError("invokeOpenChestnut no longer accepts allowLossy/allow_lossy; opaque Office content without a validated source package always fails closed.");
  }
  const loaded = await runtime();
  const wireRequest = create(CodecRequestSchema, request);
  const wireResponse = bytesFrom(loaded.invoke(toBinary(CodecRequestSchema, wireRequest)));
  const response = fromBinary(CodecResponseSchema, wireResponse);
  if (!response.ok) throw responseFailure(response);
  return response;
}

function assertCodecOptions(options, allowed, apiName) {
  if (options == null || typeof options !== "object" || Array.isArray(options)) throw new TypeError(`${apiName} options must be an object.`);
  const unsupported = Object.keys(options).filter((key) => !allowed.has(key));
  if (unsupported.length) throw new TypeError(`${apiName} does not accept option${unsupported.length === 1 ? "" : "s"} ${unsupported.join(", ")}. OpenChestnut is the only Office codec and lossy fallback is unavailable.`);
}

function cellCoordinates(address) {
  const match = /^([A-Z]{1,3})([1-9]\d*)$/i.exec(String(address));
  if (!match) throw new OpenChestnutCodecError(`Cell address ${address} is not valid A1 notation.`, [], { code: "invalid_cell_address" });
  let column = 0;
  for (const character of match[1].toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

function cellAddress(row, column) {
  let number = Number(column) + 1;
  let label = "";
  while (number > 0) {
    number -= 1;
    label = String.fromCharCode(65 + number % 26) + label;
    number = Math.floor(number / 26);
  }
  return `${label}${Number(row) + 1}`;
}

function formulaRangeBounds(reference, location) {
  const pieces = String(reference || "").split(":");
  if (pieces.length < 1 || pieces.length > 2 || !pieces[0]) throw new OpenChestnutCodecError(`Cell ${location} formula reference ${reference || "(empty)"} is not a bounded A1 range.`, [], { code: "invalid_cell_formula" });
  let first;
  let second;
  try {
    first = cellCoordinates(pieces[0].replaceAll("$", ""));
    second = pieces[1] ? cellCoordinates(pieces[1].replaceAll("$", "")) : first;
  } catch {
    throw new OpenChestnutCodecError(`Cell ${location} formula reference ${reference} is invalid.`, [], { code: "invalid_cell_formula" });
  }
  for (const coordinate of [first, second]) {
    if (coordinate.row >= 1_048_576 || coordinate.column >= 16_384) throw new OpenChestnutCodecError(`Cell ${location} formula reference ${reference} exceeds XLSX limits.`, [], { code: "invalid_cell_formula" });
  }
  if (first.row > second.row || first.column > second.column) throw new OpenChestnutCodecError(`Cell ${location} formula reference ${reference} must be top-left to bottom-right.`, [], { code: "invalid_cell_formula" });
  return { top: first.row, left: first.column, bottom: second.row, right: second.column, cellCount: (second.row - first.row + 1) * (second.column - first.column + 1) };
}

function partialSharedFormulaRanges(diagnostics) {
  const byWorksheetName = new Map();
  for (const diagnostic of diagnostics || []) {
    if (diagnostic?.code !== "partial_shared_formula_preserved") continue;
    const worksheetName = String(diagnostic.sourcePath || "");
    const reference = String(diagnostic.sourceIdentity || "");
    if (!worksheetName || !reference) {
      throw new OpenChestnutCodecError("OpenChestnut returned a partial shared-formula diagnostic without a worksheet name and range identity.", [], { code: "invalid_open_chestnut_diagnostic" });
    }
    const bounds = formulaRangeBounds(reference, `${worksheetName} partial shared formula`);
    const ranges = byWorksheetName.get(worksheetName) || [];
    ranges.push({ reference, bounds });
    byWorksheetName.set(worksheetName, ranges);
  }
  return byWorksheetName;
}

function normalizedFormula(value) {
  const formula = String(value || "");
  return formula && !formula.startsWith("=") ? `=${formula}` : formula;
}

function validateFormulaText(value, location, required = false) {
  const formula = normalizedFormula(value);
  const body = formula.startsWith("=") ? formula.slice(1) : formula;
  if (required && !body.trim()) throw new OpenChestnutCodecError(`Cell ${location} requires non-empty formula text.`, [], { code: "invalid_cell_formula" });
  if (body.length > MAX_XLSX_FORMULA_LENGTH || /\p{Cc}/u.test(body)) throw new OpenChestnutCodecError(`Cell ${location} formula is outside the bounded XLSX formula profile.`, [], { code: "invalid_cell_formula" });
  return formula;
}

function translateSharedFormula(value, source, target) {
  const formula = normalizedFormula(value);
  const rowOffset = target.row - source.row;
  const columnOffset = target.column - source.column;
  const protectedParts = [];
  const protectedFormula = formula.replace(/"(?:[^"]|"")*"|\[[^\]]*\]/g, (part) => {
    const token = `\uE000${protectedParts.length}\uE001`;
    protectedParts.push(part);
    return token;
  });
  const shifted = protectedFormula.replace(/(?<![A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_. ]*))!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])/g, (match, quotedSheet, bareSheet, absoluteColumn, columnText, absoluteRow, rowText, offset, sourceText) => {
    if (/^\s*\(/.test(sourceText.slice(offset + match.length))) return match;
    const coordinate = cellCoordinates(`${columnText}${rowText}`);
    const column = absoluteColumn ? coordinate.column : coordinate.column + columnOffset;
    const row = absoluteRow ? coordinate.row : coordinate.row + rowOffset;
    const prefix = quotedSheet != null ? `'${quotedSheet}'!` : bareSheet != null ? `${bareSheet}!` : "";
    if (column < 0 || column >= 16_384 || row < 0 || row >= 1_048_576) return `${prefix}#REF!`;
    return `${prefix}${absoluteColumn || ""}${cellAddress(row, column).replace(/\d+$/, `${absoluteRow || ""}${row + 1}`)}`;
  });
  return shifted.replace(/\uE000(\d+)\uE001/g, (_match, index) => protectedParts[Number(index)] || "");
}

function cellFormulaMetadata(address, cell) {
  const location = address;
  const type = cell.formulaType == null ? "" : String(cell.formulaType);
  if (!type && cell.formula && cell.spillError) throw new OpenChestnutCodecError(`Cell ${location} blocked dynamic array cannot be exported through the current bounded OpenChestnut slice.`, [], { code: "unsupported_dynamic_array_edit" });
  const inferredDynamic = !type && Boolean(cell.formula) && Boolean(cell.spillRange) && !cell.spillError;
  if (!type && !inferredDynamic && [cell.sharedIndex, cell.sharedRef, cell.arrayRef, cell.dynamicArrayRef].every((value) => value == null || value === "")) return undefined;
  if (type === "shared") {
    if (!Number.isInteger(cell.sharedIndex) || cell.sharedIndex < 0 || cell.sharedIndex > 0xffff_ffff) throw new OpenChestnutCodecError(`Cell ${location} shared formula requires an unsigned sharedIndex.`, [], { code: "invalid_cell_formula" });
    const reference = String(cell.sharedRef || "");
    formulaRangeBounds(reference, location);
    validateFormulaText(cell.formula, location, true);
    if (cell.arrayRef != null) throw new OpenChestnutCodecError(`Cell ${location} shared formula must not set arrayRef.`, [], { code: "invalid_cell_formula" });
    return { kind: CellFormulaKind.SHARED, sharedIndex: cell.sharedIndex, reference };
  }
  if (type === "array") {
    const reference = String(cell.arrayRef || "");
    formulaRangeBounds(reference, location);
    validateFormulaText(cell.formula, location, true);
    if (cell.sharedIndex != null || cell.sharedRef != null) throw new OpenChestnutCodecError(`Cell ${location} legacy array formula must not set shared metadata.`, [], { code: "invalid_cell_formula" });
    return { kind: CellFormulaKind.ARRAY, sharedIndex: 0, reference };
  }
  if (type === "dynamicArray" || inferredDynamic) {
    const reference = String(cell.dynamicArrayRef || cell.spillRange || "");
    const bounds = formulaRangeBounds(reference, location);
    const anchor = cellCoordinates(address);
    if (anchor.row !== bounds.top || anchor.column !== bounds.left) throw new OpenChestnutCodecError(`Cell ${location} dynamic array formula must be the top-left anchor of ${reference}.`, [], { code: "invalid_cell_formula" });
    validateFormulaText(cell.formula, location, true);
    if (cell.spillError) throw new OpenChestnutCodecError(`Cell ${location} blocked dynamic array cannot be exported through the current bounded OpenChestnut slice.`, [], { code: "unsupported_dynamic_array_edit" });
    if (cell.sharedIndex != null || cell.sharedRef != null || cell.arrayRef != null) throw new OpenChestnutCodecError(`Cell ${location} dynamic array formula must not set shared or legacy-array metadata.`, [], { code: "invalid_cell_formula" });
    return { kind: CellFormulaKind.DYNAMIC_ARRAY, sharedIndex: 0, reference };
  }
  throw new OpenChestnutCodecError(`Cell ${location} formula type ${type || "unspecified"} is outside the OpenChestnut XLSX formula slice.`, [], { code: "unsupported_cell_formula" });
}

function validateFormulaTopology(cells, sheetName) {
  const byCoordinate = new Map(cells.map((cell) => [`${cell.row}:${cell.column}`, cell]));
  if (byCoordinate.size !== cells.length) throw new OpenChestnutCodecError(`Worksheet ${sheetName} contains duplicate cell coordinates.`, [], { code: "duplicate_cell" });
  const sharedGroups = new Map();
  for (const cell of cells) {
    validateFormulaText(cell.formula, `${sheetName}!${cellAddress(cell.row, cell.column)}`, Boolean(cell.formulaMetadata) && cell.formulaMetadata.kind !== CellFormulaKind.DATA_TABLE);
    if (cell.formulaMetadata?.kind === CellFormulaKind.SHARED) {
      const key = cell.formulaMetadata.sharedIndex;
      if (!sharedGroups.has(key)) sharedGroups.set(key, []);
      sharedGroups.get(key).push(cell);
    }
  }
  for (const [index, members] of sharedGroups) {
    const references = new Set(members.map((cell) => cell.formulaMetadata.reference.toUpperCase()));
    if (references.size !== 1) throw new OpenChestnutCodecError(`Worksheet ${sheetName} shared formula si=${index} has inconsistent references.`, [], { code: "invalid_cell_formula" });
    const reference = members[0].formulaMetadata.reference;
    const bounds = formulaRangeBounds(reference, `${sheetName}!${cellAddress(members[0].row, members[0].column)}`);
    const memberMap = new Map(members.map((cell) => [`${cell.row}:${cell.column}`, cell]));
    const expectedCount = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
    if (memberMap.size !== expectedCount) throw new OpenChestnutCodecError(`Worksheet ${sheetName} shared formula si=${index} declares ${reference} with ${expectedCount} cells but contains ${memberMap.size} members.`, [], { code: "invalid_cell_formula" });
    const master = memberMap.get(`${bounds.top}:${bounds.left}`);
    if (!master) throw new OpenChestnutCodecError(`Worksheet ${sheetName} shared formula si=${index} is missing its top-left master.`, [], { code: "invalid_cell_formula" });
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      for (let column = bounds.left; column <= bounds.right; column += 1) {
        const member = memberMap.get(`${row}:${column}`);
        if (!member) throw new OpenChestnutCodecError(`Worksheet ${sheetName} shared formula si=${index} is missing ${cellAddress(row, column)}.`, [], { code: "invalid_cell_formula" });
        const expected = translateSharedFormula(master.formula, { row: bounds.top, column: bounds.left }, { row, column });
        if (normalizedFormula(member.formula) !== expected) throw new OpenChestnutCodecError(`Cell ${sheetName}!${cellAddress(row, column)} expanded shared formula must be ${expected}.`, [], { code: "invalid_cell_formula" });
      }
    }
  }
  const occupied = new Map();
  const sharedRoots = [...sharedGroups.values()].map((members) => members[0]);
  const topologyRoots = [...sharedRoots, ...cells.filter((item) => [CellFormulaKind.ARRAY, CellFormulaKind.DYNAMIC_ARRAY, CellFormulaKind.DATA_TABLE].includes(item.formulaMetadata?.kind))];
  let topologyCellCount = 0;
  for (const cell of topologyRoots) {
    const metadata = cell.formulaMetadata;
    const bounds = formulaRangeBounds(metadata.reference, `${sheetName}!${cellAddress(cell.row, cell.column)}`);
    topologyCellCount += bounds.cellCount;
    if (topologyCellCount > MAX_XLSX_FORMULA_TOPOLOGY_CELLS) throw new OpenChestnutCodecError(`Cell ${sheetName}!${cellAddress(cell.row, cell.column)} native formula topology exceeds ${MAX_XLSX_FORMULA_TOPOLOGY_CELLS} cells.`, [], { code: "invalid_cell_formula" });
    const dynamic = metadata.kind === CellFormulaKind.DYNAMIC_ARRAY;
    const array = metadata.kind === CellFormulaKind.ARRAY || dynamic;
    const dataTable = metadata.kind === CellFormulaKind.DATA_TABLE;
    const owner = metadata.kind === CellFormulaKind.SHARED ? `shared:${metadata.sharedIndex}` : `${dataTable ? "data-table" : dynamic ? "dynamic" : "array"}:${cell.row}:${cell.column}`;
    if ((array || dataTable) && (cell.row !== bounds.top || cell.column !== bounds.left)) throw new OpenChestnutCodecError(`Cell ${sheetName}!${cellAddress(cell.row, cell.column)} ${dataTable ? "data table" : `${dynamic ? "dynamic" : "legacy"} array`} formula must be the top-left anchor of ${metadata.reference}.`, [], { code: "invalid_cell_formula" });
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      for (let column = bounds.left; column <= bounds.right; column += 1) {
        const key = `${row}:${column}`;
        if (occupied.has(key) && occupied.get(key) !== owner) throw new OpenChestnutCodecError(`Cell ${sheetName}!${cellAddress(cell.row, cell.column)} formula range ${metadata.reference} overlaps another native formula range.`, [], { code: "invalid_cell_formula" });
        occupied.set(key, owner);
        const nested = byCoordinate.get(key);
        if ((array || dataTable) && (row !== cell.row || column !== cell.column) && nested?.formula) throw new OpenChestnutCodecError(`Cell ${sheetName}!${cellAddress(row, column)} must not contain another formula inside ${dataTable ? "data table" : `${dynamic ? "dynamic" : "legacy"} array`} range ${metadata.reference}.`, [], { code: "invalid_cell_formula" });
      }
    }
  }
}

function itemCount(collection) {
  return Array.isArray(collection?.items) ? collection.items.length : 0;
}

function numberFormatCode(value, address) {
  if (value == null || value === "") return "";
  if (typeof value !== "string") throw new OpenChestnutCodecError(`Cell ${address} number format must be a string.`, [], { code: "invalid_cell_number_format" });
  if (/^general$/i.test(value)) return "";
  if (value.length > MAX_XLSX_NUMBER_FORMAT_CODE_LENGTH) throw new OpenChestnutCodecError(`Cell ${address} number format exceeds ${MAX_XLSX_NUMBER_FORMAT_CODE_LENGTH} characters.`, [], { code: "invalid_cell_number_format" });
  if (/\p{Cc}/u.test(value)) throw new OpenChestnutCodecError(`Cell ${address} number format contains a control character.`, [], { code: "invalid_cell_number_format" });
  return value;
}

function cellNumberFormatCode(cell, address) {
  return numberFormatCode(cell?.style?.numberFormat ?? cell?.style?.numFmt, address);
}

function invalidCellStyle(address, message, cause) {
  throw new OpenChestnutCodecError(`Cell ${address} ${message}`, [], { code: "invalid_cell_style", cause });
}

function wireSpreadsheetColor(value, address, component) {
  if (value == null) return undefined;
  if (typeof value === "string") {
    const rgb = value.replace(/^#/, "").slice(-6).toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(rgb)) invalidCellStyle(address, `${component} color must be six/eight-digit RGB or a supported symbolic color.`);
    return { source: { case: "rgb", value: rgb } };
  }
  const tint = value.tint == null ? undefined : Number(value.tint);
  if (value.theme != null) return { source: { case: "theme", value: Number(value.theme) }, tint };
  if (value.indexed != null) return { source: { case: "indexed", value: Number(value.indexed) }, tint };
  if (value.auto === true) return { source: { case: "automatic", value: true }, tint };
  if (value.rgb != null) return { ...wireSpreadsheetColor(value.rgb, address, component), tint };
  invalidCellStyle(address, `${component} color has no supported source.`);
}

function wireBorderEdge(edge, address, name) {
  if (!edge?.style) return undefined;
  return { style: String(edge.style), color: wireSpreadsheetColor(edge.color, address, `${name} border`) };
}

function wireCellStyle(style, address) {
  const keys = Object.keys(style || {}).filter((key) => style[key] != null && !XLSX_NUMBER_FORMAT_STYLE_KEYS.has(key));
  if (keys.length === 0) return undefined;
  let normalized;
  try {
    normalized = normalizeXlsxStyle(style);
  } catch (cause) {
    invalidCellStyle(address, `has invalid static formatting: ${cause.message}`, cause);
  }
  const font = normalized.font;
  const fill = typeof normalized.fill === "string" ? { patternType: "solid", foreground: normalized.fill } : normalized.fill;
  const border = normalized.border;
  const uniformEdge = border?.style ? { style: border.style, color: border.color } : undefined;
  const edge = (name) => wireBorderEdge(uniformEdge || border?.[name], address, name);
  return {
    font: {
      bold: font.bold,
      italic: font.italic,
      underline: font.underline ? String(font.underline === true ? "single" : font.underline) : undefined,
      strike: font.strike,
      color: wireSpreadsheetColor(font.color, address, "font"),
      sizePoints: font.size,
      name: font.name,
    },
    fill: fill ? {
      patternType: fill.patternType,
      foreground: wireSpreadsheetColor(fill.foreground, address, "fill foreground"),
      background: wireSpreadsheetColor(fill.background, address, "fill background"),
    } : undefined,
    border: border ? {
      left: edge("left"), right: edge("right"), top: edge("top"), bottom: edge("bottom"),
      diagonal: edge("diagonal"), start: edge("start"), end: edge("end"),
      horizontal: edge("horizontal"), vertical: edge("vertical"),
      diagonalUp: border.diagonalUp,
      diagonalDown: border.diagonalDown,
      outline: border.outline,
    } : undefined,
    alignment: normalized.alignment ? {
      horizontal: normalized.alignment.horizontal,
      vertical: normalized.alignment.vertical,
      wrapText: normalized.alignment.wrapText,
      textRotation: normalized.alignment.textRotation,
      indent: normalized.alignment.indent,
      shrinkToFit: normalized.alignment.shrinkToFit,
      readingOrder: normalized.alignment.readingOrder,
    } : undefined,
    protection: normalized.protection ? {
      locked: normalized.protection.locked,
      hidden: normalized.protection.hidden,
    } : undefined,
  };
}

function spreadsheetColorFromWire(color) {
  if (!color?.source?.case) return undefined;
  const tint = color.tint == null || color.tint === 0 ? {} : { tint: color.tint };
  if (color.source.case === "rgb") return color.tint == null || color.tint === 0
    ? `#${String(color.source.value).slice(-6).toUpperCase()}`
    : { rgb: `#${String(color.source.value).slice(-6).toUpperCase()}`, ...tint };
  if (color.source.case === "theme") return { theme: color.source.value, ...tint };
  if (color.source.case === "indexed") return { indexed: color.source.value, ...tint };
  if (color.source.case === "automatic") return { auto: true, ...tint };
  return undefined;
}

function borderEdgeFromWire(edge) {
  if (!edge?.style) return undefined;
  return { style: edge.style, color: spreadsheetColorFromWire(edge.color) || "#000000" };
}

function cellStyleFromWire(source) {
  if (!source) return undefined;
  const style = {};
  if (source.font) {
    style.font = {
      ...(source.font.bold == null ? {} : { bold: source.font.bold }),
      ...(source.font.italic == null ? {} : { italic: source.font.italic }),
      ...(source.font.underline == null ? {} : { underline: source.font.underline }),
      ...(source.font.strike == null ? {} : { strike: source.font.strike }),
      ...(source.font.color ? { color: spreadsheetColorFromWire(source.font.color) } : {}),
      ...(source.font.sizePoints == null ? {} : { size: source.font.sizePoints }),
      ...(source.font.name == null ? {} : { name: source.font.name }),
    };
  }
  if (source.fill) {
    const foreground = spreadsheetColorFromWire(source.fill.foreground);
    const background = spreadsheetColorFromWire(source.fill.background);
    style.fill = source.fill.patternType === "solid" && typeof foreground === "string" && !background
      ? foreground
      : { patternType: source.fill.patternType || "none", ...(foreground ? { foreground } : {}), ...(background ? { background } : {}) };
  }
  if (source.border) {
    const border = {};
    for (const name of ["left", "right", "top", "bottom", "diagonal", "start", "end", "horizontal", "vertical"]) {
      const value = borderEdgeFromWire(source.border[name]);
      if (value) border[name] = value;
    }
    for (const [wire, model] of [["diagonalUp", "diagonalUp"], ["diagonalDown", "diagonalDown"], ["outline", "outline"]]) {
      if (source.border[wire] != null) border[model] = source.border[wire];
    }
    const perimeter = [border.left, border.right, border.top, border.bottom];
    const samePerimeter = perimeter.every(Boolean) && perimeter.every((candidate) => JSON.stringify(candidate) === JSON.stringify(perimeter[0]));
    const hasExtras = border.diagonal || border.start || border.end || border.horizontal || border.vertical || border.diagonalUp != null || border.diagonalDown != null || border.outline != null;
    style.border = samePerimeter && !hasExtras ? perimeter[0] : border;
  }
  if (source.alignment) {
    style.alignment = Object.fromEntries(Object.entries({
      horizontal: source.alignment.horizontal,
      vertical: source.alignment.vertical,
      wrapText: source.alignment.wrapText,
      textRotation: source.alignment.textRotation,
      indent: source.alignment.indent,
      shrinkToFit: source.alignment.shrinkToFit,
      readingOrder: source.alignment.readingOrder,
    }).filter(([, value]) => value != null));
  }
  if (source.protection) {
    style.protection = Object.fromEntries(Object.entries({ locked: source.protection.locked, hidden: source.protection.hidden }).filter(([, value]) => value != null));
  }
  return Object.keys(style).length ? style : undefined;
}

function dynamicArrayCellSnapshot(cell) {
  return {
    formula: cell.formula == null ? null : String(cell.formula),
    formulaType: cell.formulaType == null ? null : String(cell.formulaType),
    dynamicArrayRef: cell.dynamicArrayRef == null ? null : String(cell.dynamicArrayRef),
    spillRange: cell.spillRange == null ? null : String(cell.spillRange),
    spillParent: cell.spillParent == null ? null : String(cell.spillParent),
    spillAnchor: cell.spillAnchor == null ? null : String(cell.spillAnchor),
    spillError: cell.spillError == null ? null : String(cell.spillError),
  };
}

function sourceBoundFormulaCellSnapshot(cell) {
  return {
    value: cell.value instanceof Date ? { type: "date", value: cell.value.getTime() } : cell.value ?? null,
    formula: cell.formula == null ? null : String(cell.formula),
    formulaType: cell.formulaType == null ? null : String(cell.formulaType),
    sharedIndex: cell.sharedIndex == null ? null : Number(cell.sharedIndex),
    sharedRef: cell.sharedRef == null ? null : String(cell.sharedRef),
    arrayRef: cell.arrayRef == null ? null : String(cell.arrayRef),
    dynamicArrayRef: cell.dynamicArrayRef == null ? null : String(cell.dynamicArrayRef),
    spillParent: cell.spillParent == null ? null : String(cell.spillParent),
    spillAnchor: cell.spillAnchor == null ? null : String(cell.spillAnchor),
    spillRange: cell.spillRange == null ? null : String(cell.spillRange),
    spillError: cell.spillError == null ? null : String(cell.spillError),
    style: cell.style || {},
  };
}

function isDynamicArrayCell(cell) {
  return cell?.formulaType === "dynamicArray" || cell?.dynamicArrayRef != null ||
    (Boolean(cell?.formula) && cell?.spillRange != null && !cell?.spillError);
}

function unsupportedWorkbookFeatures(workbook, state) {
  const unsupported = [];
  if (workbook.indexedColors?.length) unsupported.push("custom indexed colors");
  if (workbook.connections?.length && !state) unsupported.push("source-free workbook connections");
  for (const sheet of workbook.worksheets?.items || []) {
    const prefix = `worksheet ${sheet.name}`;
    if (sheet.shapes?.length) unsupported.push(`${prefix} shapes`);
    for (const [address, cell] of sheet.store?.entries?.() || []) {
      const dynamicSlot = state?.dynamicArraySlotsBySheet?.get(sheet.id)?.get(address);
      if (isDynamicArrayCell(cell) && !dynamicSlot) unsupported.push(`${prefix} source-free dynamic array at ${address}`);
      if (cell.style && Object.keys(cell.style).some((key) => cell.style[key] != null)) wireCellStyle(cell.style, `${sheet.name}!${address}`);
      const metadata = Object.keys(cell).filter((key) => !["value", "formula", "style"].includes(key) && !XLSX_FORMULA_METADATA_KEYS.has(key));
      if (metadata.length) unsupported.push(`${prefix} advanced formula metadata at ${address}`);
    }
  }
  return unsupported;
}

const XLSX_DATA_VALIDATION_TYPES = new Set(["list", "whole", "decimal", "date", "time", "textLength", "custom"]);
const XLSX_CONDITIONAL_FORMAT_TYPES = new Set(["cellIs", "expression", "containsText", "colorScale", "dataBar", "iconSet"]);
const BRACED_GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/i;

function wireDataValidation(item, sheetName) {
  const rule = item?.rule || item || {};
  const type = String(rule.type || item?.type || "");
  if (!XLSX_DATA_VALIDATION_TYPES.has(type)) throw new OpenChestnutCodecError(`Worksheet ${sheetName} data validation ${item?.id || "(unnamed)"} uses unsupported type ${type || "(empty)"}.`, [], { code: "unsupported_data_validation" });
  const values = Array.isArray(rule.values) ? rule.values.map(String) : [];
  if (values.length && rule.formula1 != null) throw new OpenChestnutCodecError(`Worksheet ${sheetName} data validation ${item?.id || "(unnamed)"} cannot combine values and formula1.`, [], { code: "invalid_data_validation" });
  return {
    id: String(item?.id || ""),
    range: String(item?.range || "A1"),
    type,
    operator: String(rule.operator || item?.operator || ""),
    formula1: rule.formula1 == null ? "" : String(rule.formula1),
    formula2: rule.formula2 == null ? "" : String(rule.formula2),
    values,
  };
}

function publicDataValidation(item) {
  return {
    id: item.id || undefined,
    range: item.range || "A1",
    rule: {
      type: item.type,
      ...(item.operator ? { operator: item.operator } : {}),
      ...(item.formula1 ? { formula1: item.formula1 } : {}),
      ...(item.formula2 ? { formula2: item.formula2 } : {}),
      ...(item.values?.length ? { values: [...item.values] } : {}),
    },
  };
}

function wireConditionalFormat(item, sheetName, index) {
  const rule = item?.rule || item || {};
  const ruleType = String(item?.ruleType || rule.ruleType || rule.type || "expression");
  if (!XLSX_CONDITIONAL_FORMAT_TYPES.has(ruleType)) throw new OpenChestnutCodecError(`Worksheet ${sheetName} conditional format ${item?.id || "(unnamed)"} uses unsupported type ${ruleType}.`, [], { code: "unsupported_conditional_format" });
  const rawFormulas = item?.formulas ?? rule.formulas ?? item?.formula ?? item?.expression ?? rule.formula ?? rule.expression;
  const formulas = (Array.isArray(rawFormulas) ? rawFormulas : rawFormulas == null ? [] : [rawFormulas]).map(String);
  const colors = (item?.colors || rule.colors || []).map((color, colorIndex) => wireSpreadsheetColor(color, `${sheetName}!${item?.range || "A1"}`, `conditional format color ${colorIndex + 1}`));
  let dataBar;
  let iconSet;
  try {
    if (ruleType === "dataBar") {
      const profile = normalizeDataBarConfig(rule, `Worksheet ${sheetName} dataBar ${item?.id || "(unnamed)"}`);
      dataBar = {
        color: wireSpreadsheetColor(profile.color, `${sheetName}!${item?.range || "A1"}`, "data-bar"),
        thresholds: profile.thresholds.map((threshold) => ({ ...threshold })),
        ...(profile.showValue == null ? {} : { showValue: profile.showValue }),
        ...(profile.gradient == null ? {} : { gradient: profile.gradient }),
      };
    }
    if (ruleType === "iconSet") {
      const profile = normalizeIconSetConfig(rule, `Worksheet ${sheetName} iconSet ${item?.id || "(unnamed)"}`);
      iconSet = {
        iconSet: profile.iconSet,
        thresholds: profile.thresholds.map((threshold) => ({ ...threshold })),
        ...(profile.showValue == null ? {} : { showValue: profile.showValue }),
        ...(profile.reverse == null ? {} : { reverse: profile.reverse }),
      };
    }
  } catch (cause) {
    throw new OpenChestnutCodecError(cause.message, [], { code: "invalid_conditional_format", cause });
  }
  if (ruleType === "colorScale" && ![2, 3].includes(colors.length)) throw new OpenChestnutCodecError(`Worksheet ${sheetName} colorScale ${item?.id || "(unnamed)"} requires two or three colors.`, [], { code: "invalid_conditional_format" });
  if (!["colorScale", "dataBar", "iconSet", "containsText"].includes(ruleType) && formulas.length === 0) throw new OpenChestnutCodecError(`Worksheet ${sheetName} conditional format ${item?.id || "(unnamed)"} requires a formula.`, [], { code: "invalid_conditional_format" });
  if (ruleType !== "colorScale" && colors.length) throw new OpenChestnutCodecError(`Worksheet ${sheetName} conditional format ${item?.id || "(unnamed)"} colors are valid only for colorScale.`, [], { code: "invalid_conditional_format" });
  if ((dataBar || iconSet) && (formulas.length || item?.format || rule.format || item?.operator || rule.operator || item?.text || rule.text)) throw new OpenChestnutCodecError(`Worksheet ${sheetName} ${ruleType} ${item?.id || "(unnamed)"} cannot combine visual metadata with formulas, operators, text, or differential formatting.`, [], { code: "invalid_conditional_format" });
  return {
    id: String(item?.id || ""),
    range: String(item?.range || "A1"),
    ruleType,
    operator: String(item?.operator || rule.operator || ""),
    formulas,
    text: String(item?.text || rule.text || ""),
    format: wireCellStyle(item?.format || rule.format || {}, `${sheetName}!${item?.range || "A1"}`),
    colors,
    priority: Number.isInteger(item?.priority) && item.priority > 0 ? item.priority : index + 1,
    dataBar,
    iconSet,
  };
}

function publicConditionalFormat(item) {
  const formulas = [...(item.formulas || [])];
  const thresholds = (source) => (source?.thresholds || []).map((threshold) => ({ type: threshold.type, ...(threshold.value == null ? {} : { value: threshold.value }) }));
  return {
    id: item.id || undefined,
    range: item.range || "A1",
    ruleType: item.ruleType,
    ...(item.operator ? { operator: item.operator } : {}),
    ...(formulas.length === 1 ? { formula: formulas[0] } : formulas.length ? { formulas } : {}),
    ...(item.text ? { text: item.text } : {}),
    ...(item.format ? { format: cellStyleFromWire(item.format) || {} } : {}),
    ...(item.colors?.length ? { colors: item.colors.map(spreadsheetColorFromWire) } : {}),
    ...(item.dataBar ? {
      color: spreadsheetColorFromWire(item.dataBar.color),
      thresholds: thresholds(item.dataBar),
      ...(item.dataBar.showValue == null ? {} : { showValue: item.dataBar.showValue }),
      ...(item.dataBar.gradient == null ? {} : { gradient: item.dataBar.gradient }),
    } : {}),
    ...(item.iconSet ? {
      iconSet: item.iconSet.iconSet,
      thresholds: thresholds(item.iconSet),
      ...(item.iconSet.showValue == null ? {} : { showValue: item.iconSet.showValue }),
      ...(item.iconSet.reverse == null ? {} : { reverse: item.iconSet.reverse }),
    } : {}),
    ...(item.priority ? { priority: item.priority } : {}),
  };
}

function publicThreadedComment(item) {
  return {
    ...(item.nativeCommentId || item.id ? { id: item.nativeCommentId || item.id } : {}),
    ...(item.personId ? { personId: item.personId } : {}),
    ...(item.dateTime ? { date: item.dateTime } : {}),
    ...(item.parentNativeCommentId ? { parentId: item.parentNativeCommentId } : {}),
    author: item.author || "User",
    person: {
      displayName: item.author || "User",
      ...(item.userId ? { userId: item.userId } : {}),
      ...(item.providerId ? { providerId: item.providerId } : {}),
    },
    done: Boolean(item.resolved),
  };
}

function wireThreadedComments(workbook, sheet) {
  const activeSheetName = workbook.worksheets.getActiveWorksheet().name;
  return (workbook.comments?.threads || []).filter((thread) => (thread.target?.sheetName || activeSheetName) === sheet.name).flatMap((thread) => {
    const address = String(thread.target?.address || "").toUpperCase();
    if (!/^[A-Z]{1,3}[1-9]\d*$/.test(address)) throw new OpenChestnutCodecError(`Worksheet ${sheet.name} threaded comment ${thread.id} must target one cell.`, [], { code: "invalid_threaded_comment_target" });
    const comments = thread.comments || [];
    if (!comments.length) throw new OpenChestnutCodecError(`Worksheet ${sheet.name} threaded comment ${thread.id} has no root comment.`, [], { code: "invalid_spreadsheet_threaded_comment" });
    const nativeIds = comments.map((comment, index) => BRACED_GUID.test(comment.id || "")
      ? String(comment.id).toUpperCase()
      : deterministicSpreadsheetGuid(`open-chestnut:${sheet.id}:${thread.id}:${address}:${index}`));
    const rootModelIds = new Set([String(thread.id || ""), String(comments[0]?.id || ""), nativeIds[0]].filter(Boolean));
    return comments.map((comment, index) => {
      if (index > 0 && comment.parentId != null && !rootModelIds.has(String(comment.parentId))) {
        throw new OpenChestnutCodecError(`Worksheet ${sheet.name} threaded comment ${thread.id} contains a nested or branched reply graph.`, [], { code: "unsupported_threaded_comment_reply_topology" });
      }
      const person = comment.person || {};
      return {
        id: String(index === 0 ? thread.id || "" : comment.modelId || comment.id || `${thread.id}/reply/${index}`),
        cellReference: address,
        nativeCommentId: nativeIds[index],
        text: String(comment.text ?? ""),
        personId: BRACED_GUID.test(comment.personId || person.id || "") ? String(comment.personId || person.id).toUpperCase() : "",
        author: String(person.displayName || comment.author || thread.author || "User"),
        userId: String(person.userId ?? comment.userId ?? ""),
        providerId: String(person.providerId ?? comment.providerId ?? ""),
        dateTime: comment.date ? new Date(comment.date).toISOString() : "1970-01-01T00:00:00.000Z",
        resolved: Boolean(comment.done ?? thread.resolved),
        ...(index > 0 ? { parentNativeCommentId: nativeIds[0] } : {}),
      };
    });
  });
}

function wireWorkbookTheme(theme, source) {
  let normalized;
  try {
    normalized = normalizeXlsxThemeConfig(theme);
  } catch (cause) {
    throw new OpenChestnutCodecError(`Workbook theme is invalid: ${cause.message}`, [], { code: "invalid_workbook_theme", cause });
  }
  return {
    name: normalized.name,
    ...Object.fromEntries(XLSX_THEME_WIRE_FIELDS.map(([model, wire]) => [wire, normalized.colors[model].replace(/^#/, "").toUpperCase()])),
    source,
  };
}

function workbookThemeFromWire(theme) {
  if (!theme || XLSX_THEME_WIRE_FIELDS.some(([, wire]) => !/^[0-9A-Fa-f]{6}$/.test(theme[wire] || ""))) return undefined;
  return normalizeXlsxThemeConfig({
    name: theme.name,
    colors: Object.fromEntries(XLSX_THEME_WIRE_FIELDS.map(([model, wire]) => [model, `#${theme[wire]}`])),
  });
}

function sameWorkbookTheme(left, right) {
  const a = normalizeXlsxThemeConfig(left);
  const b = normalizeXlsxThemeConfig(right);
  return a.name === b.name && XLSX_THEME_COLOR_NAMES.every((name) => a.colors[name] === b.colors[name]);
}

const WORKBOOK_CONNECTION_BOOLEAN_FIELDS = ["keepAlive", "background", "refreshOnLoad", "saveData"];

function publicWorkbookConnection(value) {
  const connection = {
    connectionId: Number(value?.connectionId ?? 0),
    name: String(value?.name ?? ""),
    type: Number(value?.type ?? 0),
    refreshedVersion: Number(value?.refreshedVersion ?? 0),
  };
  if (value?.description !== undefined) connection.description = String(value.description);
  for (const field of WORKBOOK_CONNECTION_BOOLEAN_FIELDS) if (value?.[field] !== undefined) connection[field] = Boolean(value[field]);
  if (value?.intervalMinutes !== undefined) connection.intervalMinutes = Number(value.intervalMinutes);
  return connection;
}

function connectionSnapshot(value) {
  return publicWorkbookConnection(value);
}

function wireWorkbookConnections(workbook, state) {
  const remaining = new Set(workbook.connections || []);
  const output = [];
  for (const slot of state?.connectionSlots || []) {
    if (!remaining.delete(slot.connection)) {
      throw new OpenChestnutCodecError(`Workbook cannot remove imported connection ${slot.connection.connectionId} in the bounded OpenChestnut slice.`, [], { code: "invalid_workbook_connection" });
    }
    if (JSON.stringify(connectionSnapshot(slot.connection)) !== JSON.stringify(slot.publicSnapshot)) {
      throw new OpenChestnutCodecError(`Imported workbook connection ${slot.connection.connectionId} is source-bound and read-only in OpenChestnut 0.2.`, [], { code: "unsupported_workbook_connection_edit" });
    }
    output.push(slot.wire);
  }
  if (remaining.size) {
    throw new OpenChestnutCodecError("OpenChestnut 0.2 cannot author workbook connections; imported connections are source-bound and read-only.", [], { code: "unsupported_workbook_connection_edit" });
  }
  return output;
}

function publicWorkbookDefinedName(value) {
  const definedName = {
    id: String(value?.id ?? ""),
    name: String(value?.name ?? ""),
    refersTo: String(value?.refersTo ?? ""),
  };
  if (value?.scopeSheetName !== undefined) definedName.scope = String(value.scopeSheetName);
  if (value?.comment !== undefined) definedName.comment = String(value.comment);
  if (value?.hidden !== undefined) definedName.hidden = Boolean(value.hidden);
  return definedName;
}

function definedNameSnapshot(value) {
  return {
    id: String(value?.id ?? ""),
    name: String(value?.name ?? ""),
    refersTo: String(value?.refersTo ?? ""),
    ...(value?.scope !== undefined ? { scope: String(value.scope) } : {}),
    ...(value?.comment !== undefined ? { comment: String(value.comment) } : {}),
    ...(value?.hidden !== undefined ? { hidden: Boolean(value.hidden) } : {}),
  };
}

function wireWorkbookDefinedName(value, source) {
  const publicValue = definedNameSnapshot(value);
  return {
    id: publicValue.id,
    name: publicValue.name,
    refersTo: publicValue.refersTo,
    ...(publicValue.scope !== undefined ? { scopeSheetName: publicValue.scope } : {}),
    ...(publicValue.comment !== undefined ? { comment: publicValue.comment } : {}),
    ...(publicValue.hidden !== undefined ? { hidden: publicValue.hidden } : {}),
    source,
  };
}

function wireWorkbookDefinedNames(workbook, state) {
  const remaining = new Set(workbook.definedNames?.items || []);
  const output = [];
  for (const slot of state?.definedNameSlots || []) {
    if (!remaining.delete(slot.definedName)) {
      throw new OpenChestnutCodecError(`Workbook cannot remove imported defined name ${slot.definedName.name} in the bounded OpenChestnut slice.`, [], { code: "invalid_workbook_defined_name" });
    }
    output.push(JSON.stringify(definedNameSnapshot(slot.definedName)) === JSON.stringify(slot.publicSnapshot)
      ? slot.wire
      : wireWorkbookDefinedName(slot.definedName, slot.wire.source));
  }
  output.push(...[...remaining].map((definedName) => wireWorkbookDefinedName(definedName)));
  return output;
}

function publicWorkbookCalculation(value) {
  if (!value) return undefined;
  const calculation = {};
  if (value.mode !== undefined) {
    calculation.mode = value.mode === SpreadsheetCalculationMode.AUTOMATIC ? "automatic"
      : value.mode === SpreadsheetCalculationMode.AUTOMATIC_EXCEPT_TABLES ? "automaticExceptTables"
      : value.mode === SpreadsheetCalculationMode.MANUAL ? "manual" : undefined;
    if (!calculation.mode) throw new OpenChestnutCodecError("OpenChestnut returned an unsupported workbook calculation mode.", [], { code: "invalid_workbook_calculation" });
  }
  for (const [wireField, publicField] of [["calculateOnSave", "calculateOnSave"], ["fullCalculationOnLoad", "fullCalculationOnLoad"], ["forceFullCalculation", "forceFullCalculation"], ["fullPrecision", "fullPrecision"]])
    if (value[wireField] !== undefined) calculation[publicField] = Boolean(value[wireField]);
  const iteration = {};
  if (value.iterationEnabled !== undefined) iteration.enabled = Boolean(value.iterationEnabled);
  if (value.maxIterations !== undefined) iteration.maxIterations = value.maxIterations;
  if (value.maxChange !== undefined) iteration.maxChange = value.maxChange;
  if (Object.keys(iteration).length) calculation.iteration = iteration;
  return calculation;
}

function calculationSnapshot(value) {
  if (value === undefined) return undefined;
  return {
    ...(value.mode !== undefined ? { mode: value.mode } : {}),
    ...(value.calculateOnSave !== undefined ? { calculateOnSave: Boolean(value.calculateOnSave) } : {}),
    ...(value.fullCalculationOnLoad !== undefined ? { fullCalculationOnLoad: Boolean(value.fullCalculationOnLoad) } : {}),
    ...(value.forceFullCalculation !== undefined ? { forceFullCalculation: Boolean(value.forceFullCalculation) } : {}),
    ...(value.iteration ? { iteration: {
      ...(value.iteration.enabled !== undefined ? { enabled: Boolean(value.iteration.enabled) } : {}),
      ...(value.iteration.maxIterations !== undefined ? { maxIterations: Number(value.iteration.maxIterations) } : {}),
      ...(value.iteration.maxChange !== undefined ? { maxChange: Number(value.iteration.maxChange) } : {}),
    } } : {}),
    ...(value.fullPrecision !== undefined ? { fullPrecision: Boolean(value.fullPrecision) } : {}),
  };
}

function wireWorkbookCalculation(value, source) {
  if (value === undefined) return undefined;
  const calculation = calculationSnapshot(value);
  const mode = calculation.mode === "automatic" ? SpreadsheetCalculationMode.AUTOMATIC
    : calculation.mode === "automaticExceptTables" ? SpreadsheetCalculationMode.AUTOMATIC_EXCEPT_TABLES
    : calculation.mode === "manual" ? SpreadsheetCalculationMode.MANUAL : undefined;
  if (calculation.mode !== undefined && mode === undefined) throw new OpenChestnutCodecError(`Unsupported workbook calculation mode ${calculation.mode}.`, [], { code: "invalid_workbook_calculation" });
  return {
    ...(mode !== undefined ? { mode } : {}),
    ...(calculation.calculateOnSave !== undefined ? { calculateOnSave: calculation.calculateOnSave } : {}),
    ...(calculation.fullCalculationOnLoad !== undefined ? { fullCalculationOnLoad: calculation.fullCalculationOnLoad } : {}),
    ...(calculation.forceFullCalculation !== undefined ? { forceFullCalculation: calculation.forceFullCalculation } : {}),
    ...(calculation.iteration?.enabled !== undefined ? { iterationEnabled: calculation.iteration.enabled } : {}),
    ...(calculation.iteration?.maxIterations !== undefined ? { maxIterations: calculation.iteration.maxIterations } : {}),
    ...(calculation.iteration?.maxChange !== undefined ? { maxChange: calculation.iteration.maxChange } : {}),
    ...(calculation.fullPrecision !== undefined ? { fullPrecision: calculation.fullPrecision } : {}),
    source,
  };
}

function wireWorkbookCalculationForExport(workbook, state) {
  const slot = state?.calculationSlot;
  if (!slot) return wireWorkbookCalculation(workbook.calculation);
  if (workbook.calculation === undefined) throw new OpenChestnutCodecError("Workbook cannot remove imported calculation properties in the bounded OpenChestnut slice.", [], { code: "invalid_workbook_calculation" });
  return JSON.stringify(calculationSnapshot(workbook.calculation)) === JSON.stringify(slot.publicSnapshot)
    ? slot.wire
    : wireWorkbookCalculation(workbook.calculation, slot.wire.source);
}

function tableColumnNames(table) {
  let bounds;
  try {
    bounds = formulaRangeBounds(table.range, table.name || "worksheet table");
  } catch (cause) {
    throw new OpenChestnutCodecError(`Worksheet table ${table.name || "(unnamed)"} has an invalid range: ${cause.message}`, [], { code: "invalid_worksheet_table", cause });
  }
  const count = bounds.right - bounds.left + 1;
  if (Array.isArray(table.columnNames)) return table.columnNames.map((value) => String(value));
  const headers = table.showHeaders !== false && table.values?.[0] ? table.values[0] : [];
  return Array.from({ length: count }, (_value, index) => String(headers[index] ?? "").trim() || `Column${index + 1}`);
}

function tableColumnDefinitions(table, names) {
  if (!Array.isArray(table.columnDefinitions)) return undefined;
  return names.map((name, index) => {
    const column = table.columnDefinitions[index] || {};
    return {
      name,
      calculatedColumnFormula: column.calculatedColumnFormula ? String(column.calculatedColumnFormula) : "",
      calculatedColumnFormulaArray: Boolean(column.calculatedColumnFormulaArray),
      totalsRowFunction: column.totalsRowFunction ? String(column.totalsRowFunction) : "",
      totalsRowLabel: column.totalsRowLabel ? String(column.totalsRowLabel) : "",
      totalsRowFormula: column.totalsRowFormula ? String(column.totalsRowFormula) : "",
      totalsRowFormulaArray: Boolean(column.totalsRowFormulaArray),
    };
  });
}

function tableFilters(table) {
  if (!Array.isArray(table.filters)) return [];
  return table.filters.map((filter) => {
    const columnIndex = Number(filter?.columnIndex ?? 0);
    if (filter?.kind === "custom") {
      return {
        columnIndex,
        criteria: {
          case: "custom",
          value: {
            matchAll: Boolean(filter.matchAll),
            criteria: Array.isArray(filter.criteria)
              ? filter.criteria.map((criterion) => ({ operator: String(criterion?.operator ?? ""), value: String(criterion?.value ?? "") }))
              : [],
          },
        },
      };
    }
    if (filter?.kind === "dynamic") {
      return {
        columnIndex,
        criteria: {
          case: "dynamic",
          value: {
            type: String(filter.type ?? ""),
            value: filter.value == null ? undefined : Number(filter.value),
            maxValue: filter.maxValue == null ? undefined : Number(filter.maxValue),
          },
        },
      };
    }
    if (filter?.kind === "top10") {
      return {
        columnIndex,
        criteria: {
          case: "top10",
          value: {
            top: filter.top ?? true,
            percent: Boolean(filter.percent),
            value: Number(filter.value ?? 0),
            filterValue: filter.filterValue == null ? undefined : Number(filter.filterValue),
          },
        },
      };
    }
    if (filter?.kind === "icon") {
      return {
        columnIndex,
        criteria: {
          case: "icon",
          value: {
            iconSet: String(filter.iconSet ?? ""),
            iconId: filter.iconId == null ? undefined : Number(filter.iconId),
          },
        },
      };
    }
    if (filter?.kind === "color") {
      return {
        columnIndex,
        criteria: { case: "color", value: wireTableColor(filter, `table ${table.name} filter column ${columnIndex}`) },
      };
    }
    return {
      columnIndex,
      criteria: {
        case: "values",
        value: {
          values: Array.isArray(filter?.values) ? filter.values.map((value) => String(value)) : [],
          includeBlank: Boolean(filter?.includeBlank),
          dateGroups: Array.isArray(filter?.dateGroups) ? filter.dateGroups.map((group) => ({
            grouping: String(group?.grouping ?? ""),
            year: Number(group?.year ?? 0),
            month: group?.month == null ? undefined : Number(group.month),
            day: group?.day == null ? undefined : Number(group.day),
            hour: group?.hour == null ? undefined : Number(group.hour),
            minute: group?.minute == null ? undefined : Number(group.minute),
            second: group?.second == null ? undefined : Number(group.second),
          })) : [],
          calendarType: filter?.calendarType ? String(filter.calendarType) : "",
        },
      },
    };
  });
}

function publicTableFilter(filter) {
  if (filter?.criteria?.case === "custom") {
    return {
      columnIndex: Number(filter.columnIndex ?? 0),
      kind: "custom",
      matchAll: Boolean(filter.criteria.value?.matchAll),
      criteria: (filter.criteria.value?.criteria || []).map((criterion) => ({ operator: criterion.operator, value: criterion.value })),
    };
  }
  if (filter?.criteria?.case === "dynamic") {
    return {
      columnIndex: Number(filter.columnIndex ?? 0),
      kind: "dynamic",
      type: filter.criteria.value?.type || "",
      ...(filter.criteria.value?.value == null ? {} : { value: filter.criteria.value.value }),
      ...(filter.criteria.value?.maxValue == null ? {} : { maxValue: filter.criteria.value.maxValue }),
    };
  }
  if (filter?.criteria?.case === "top10") {
    return {
      columnIndex: Number(filter.columnIndex ?? 0),
      kind: "top10",
      top: Boolean(filter.criteria.value?.top),
      percent: Boolean(filter.criteria.value?.percent),
      value: Number(filter.criteria.value?.value ?? 0),
      ...(filter.criteria.value?.filterValue == null ? {} : { filterValue: filter.criteria.value.filterValue }),
    };
  }
  if (filter?.criteria?.case === "icon") {
    return {
      columnIndex: Number(filter.columnIndex ?? 0),
      kind: "icon",
      iconSet: filter.criteria.value?.iconSet || "",
      ...(filter.criteria.value?.iconId == null ? {} : { iconId: filter.criteria.value.iconId }),
    };
  }
  if (filter?.criteria?.case === "color") {
    return {
      columnIndex: Number(filter.columnIndex ?? 0),
      kind: "color",
      ...publicTableColor(filter.criteria.value),
    };
  }
  return {
    columnIndex: Number(filter?.columnIndex ?? 0),
    kind: "values",
    values: [...(filter?.criteria?.value?.values || [])],
    includeBlank: Boolean(filter?.criteria?.value?.includeBlank),
    ...((filter?.criteria?.value?.dateGroups || []).length ? {
      dateGroups: filter.criteria.value.dateGroups.map((group) => ({
        grouping: group.grouping,
        year: Number(group.year ?? 0),
        ...(group.month == null ? {} : { month: group.month }),
        ...(group.day == null ? {} : { day: group.day }),
        ...(group.hour == null ? {} : { hour: group.hour }),
        ...(group.minute == null ? {} : { minute: group.minute }),
        ...(group.second == null ? {} : { second: group.second }),
      })),
    } : {}),
    ...(filter?.criteria?.value?.calendarType ? { calendarType: filter.criteria.value.calendarType } : {}),
  };
}

function wireTableSortState(sort, address) {
  if (!sort) return undefined;
  return {
    reference: String(sort.reference ?? ""),
    caseSensitive: Boolean(sort.caseSensitive),
    ...(sort.sortMethod == null ? {} : { sortMethod: String(sort.sortMethod) }),
    ...(sort.columnSort == null ? {} : { columnSort: Boolean(sort.columnSort) }),
    conditions: Array.isArray(sort.conditions)
      ? sort.conditions.map((condition) => ({
          reference: String(condition?.reference ?? ""),
          descending: Boolean(condition?.descending),
          ...((condition?.kind === "icon" || condition?.iconSet) ? {
            icon: {
              iconSet: String(condition.iconSet ?? ""),
              iconId: condition.iconId == null ? undefined : Number(condition.iconId),
            },
          } : condition?.kind === "color" ? {
            color: wireTableColor(condition, `${address} sort ${condition.reference}`),
          } : condition?.customList == null ? {} : { customList: String(condition.customList) }),
        }))
      : [],
  };
}

function tableSortState(table) {
  return wireTableSortState(table?.sortState, `table ${table?.name || "(unnamed)"}`);
}

function publicTableSortState(sort) {
  if (!sort) return undefined;
  return {
    reference: sort.reference,
    caseSensitive: Boolean(sort.caseSensitive),
    ...(sort.sortMethod == null ? {} : { sortMethod: sort.sortMethod }),
    ...(sort.columnSort == null ? {} : { columnSort: Boolean(sort.columnSort) }),
    conditions: (sort.conditions || []).map((condition) => ({
      reference: condition.reference,
      descending: Boolean(condition.descending),
      ...(condition.icon ? {
        kind: "icon",
        iconSet: condition.icon.iconSet,
        ...(condition.icon.iconId == null ? {} : { iconId: condition.icon.iconId }),
      } : condition.color ? {
        kind: "color",
        ...publicTableColor(condition.color),
      } : condition.customList == null ? {} : { customList: condition.customList }),
    })),
  };
}

function wireTableColor(value, address) {
  const target = value?.target;
  if (target !== "cell" && target !== "font") {
    throw new OpenChestnutCodecError(`Worksheet ${address} color target must be 'cell' or 'font'.`, [], { code: "invalid_worksheet_table" });
  }
  const color = wireSpreadsheetColor(value.color, address, `${target} color`);
  if (!color) throw new OpenChestnutCodecError(`Worksheet ${address} must provide a color.`, [], { code: "invalid_worksheet_table" });
  return { target: { case: target === "cell" ? "cellColor" : "fontColor", value: true }, color };
}

function publicTableColor(value) {
  return {
    target: value?.target?.case === "cellColor" ? "cell" : "font",
    color: spreadsheetColorFromWire(value?.color),
  };
}

const TABLE_QUERY_BOOLEAN_FIELDS = [
  "headers", "rowNumbers", "disableRefresh", "backgroundRefresh", "firstBackgroundRefresh", "refreshOnLoad",
  "fillFormulas", "removeDataOnSave", "disableEdit", "preserveFormatting", "adjustColumnWidth", "intermediate",
  "applyNumberFormats", "applyBorderFormats", "applyFontFormats", "applyPatternFormats", "applyAlignmentFormats",
  "applyWidthHeightFormats",
];

const TABLE_QUERY_REFRESH_BOOLEAN_FIELDS = ["preserveSortFilterLayout", "fieldIdWrapped", "headersInLastRefresh"];
const TABLE_QUERY_REFRESH_UINT_FIELDS = ["minimumVersion", "nextId", "unboundColumnsLeft", "unboundColumnsRight"];
const TABLE_QUERY_FIELD_BOOLEAN_FIELDS = ["dataBound", "rowNumbers", "fillFormulas", "clipped"];

function publicTableQueryField(value) {
  const field = { id: Number(value?.id ?? 0) };
  if (value?.name !== undefined) field.name = String(value.name);
  for (const name of TABLE_QUERY_FIELD_BOOLEAN_FIELDS) if (value?.[name] !== undefined) field[name] = Boolean(value[name]);
  if (value?.tableColumnId !== undefined) field.tableColumnId = Number(value.tableColumnId);
  return field;
}

function publicTableQueryRefresh(value) {
  if (!value) return undefined;
  const refresh = { fields: Array.isArray(value.fields) ? value.fields.map(publicTableQueryField) : [] };
  for (const field of TABLE_QUERY_REFRESH_BOOLEAN_FIELDS) if (value[field] !== undefined) refresh[field] = Boolean(value[field]);
  for (const field of TABLE_QUERY_REFRESH_UINT_FIELDS) if (value[field] !== undefined) refresh[field] = Number(value[field]);
  if (Array.isArray(value.deletedFieldNames) && value.deletedFieldNames.length)
    refresh.deletedFieldNames = value.deletedFieldNames.map((name) => String(name));
  if (value.sortState) refresh.sortState = publicTableSortState(value.sortState);
  return refresh;
}

function publicTableQuery(value) {
  if (!value) return undefined;
  const query = { name: String(value.name ?? ""), connectionId: Number(value.connectionId ?? 0) };
  for (const field of TABLE_QUERY_BOOLEAN_FIELDS) if (value[field] !== undefined) query[field] = Boolean(value[field]);
  if (value.growShrinkType !== undefined) query.growShrinkType = String(value.growShrinkType);
  if (value.autoFormatId !== undefined) query.autoFormatId = Number(value.autoFormatId);
  if (value.refresh) query.refresh = publicTableQueryRefresh(value.refresh);
  return query;
}

function wireTableQuery(table) {
  const state = table[TABLE_STATE];
  const query = publicTableQuery(table.queryTable);
  if (state) {
    if (JSON.stringify(query) !== JSON.stringify(state.querySnapshot)) {
      throw new OpenChestnutCodecError(`Imported query table ${table.name} is source-bound and read-only in OpenChestnut 0.2.`, [], { code: "unsupported_query_table_edit" });
    }
    return state.wire?.queryTable;
  }
  if (query) {
    throw new OpenChestnutCodecError(`OpenChestnut 0.2 cannot author query table ${table.name}; imported query tables are source-bound and read-only.`, [], { code: "unsupported_query_table_edit" });
  }
  return undefined;
}

function tableSnapshot(table) {
  const columnNames = tableColumnNames(table);
  return {
    id: table.id,
    name: table.name,
    reference: table.range,
    hasHeaders: table.showHeaders !== false,
    showTotals: Boolean(table.showTotals),
    showFilterButton: table.showFilterButton !== false,
    styleName: table.style || "TableStyleMedium2",
    showFirstColumn: Boolean(table.showFirstColumn),
    showLastColumn: Boolean(table.showLastColumn),
    showRowStripes: table.showRowStripes ?? table.showHeaders !== false,
    showColumnStripes: Boolean(table.showBandedColumns),
    columnNames,
    columns: tableColumnDefinitions(table, columnNames),
    filters: tableFilters(table),
    sortState: tableSortState(table),
    queryTable: publicTableQuery(table.queryTable),
  };
}

function sameTableSnapshot(table, snapshot) {
  return JSON.stringify(tableSnapshot(table)) === JSON.stringify(snapshot);
}

function wireWorksheetTable(table) {
  return { ...tableSnapshot(table), queryTable: wireTableQuery(table), source: table[TABLE_STATE]?.wire?.source };
}

function wireWorksheetTables(sheet, state) {
  const remaining = new Set(sheet.tables?.items || []);
  const output = [];
  for (const slot of state?.slots || []) {
    if (!slot.table) {
      output.push(slot.wire);
      continue;
    }
    if (!remaining.delete(slot.table)) {
      throw new OpenChestnutCodecError(`Worksheet ${sheet.name} cannot remove imported table ${slot.table.name} in the bounded OpenChestnut slice.`, [], { code: "invalid_worksheet_table" });
    }
    output.push(sameTableSnapshot(slot.table, slot.publicSnapshot) ? slot.wire : wireWorksheetTable(slot.table));
  }
  output.push(...[...remaining].map(wireWorksheetTable));
  return output;
}

function excelSerialFromDate(value, dateSystem, address) {
  const milliseconds = value.getTime();
  if (!Number.isFinite(milliseconds)) throw new OpenChestnutCodecError(`Cell ${address} has an invalid Date value.`, [], { code: "invalid_cell_date" });
  const dayMilliseconds = 86_400_000;
  if (dateSystem === "1904") return (milliseconds - Date.UTC(1904, 0, 1)) / dayMilliseconds;
  const serial = (milliseconds - Date.UTC(1899, 11, 31)) / dayMilliseconds;
  return milliseconds >= Date.UTC(1900, 2, 1) ? serial + 1 : serial;
}

function wireCell(address, cell, dateSystem) {
  const coordinates = cellCoordinates(address);
  const dateValue = cell.value instanceof Date;
  const target = {
    row: coordinates.row,
    column: coordinates.column,
    formula: cell.formula ? String(cell.formula) : "",
    formulaMetadata: cellFormulaMetadata(address, cell),
    numberFormatCode: cellNumberFormatCode(cell, address) || (dateValue ? "yyyy-mm-dd hh:mm:ss" : ""),
    style: wireCellStyle(cell.style, address),
    value: { case: undefined },
  };
  if (cell.value == null) return target;
  if (dateValue) {
    target.value = { case: "numberValue", value: excelSerialFromDate(cell.value, dateSystem, address) };
  } else if (typeof cell.value === "string") {
    target.value = EXCEL_ERRORS.has(cell.value) ? { case: "errorValue", value: cell.value } : { case: "stringValue", value: cell.value };
  } else if (typeof cell.value === "number") {
    if (!Number.isFinite(cell.value)) throw new OpenChestnutCodecError(`Cell ${address} has a non-finite numeric value.`, [], { code: "non_finite_cell_value" });
    target.value = { case: "numberValue", value: cell.value };
  } else if (typeof cell.value === "boolean") {
    target.value = { case: "boolValue", value: cell.value };
  } else {
    throw new OpenChestnutCodecError(`Cell ${address} has unsupported ${cell.value?.constructor?.name || typeof cell.value} content.`, [], { code: "unsupported_cell_value" });
  }
  return target;
}

function workbookEnvelope(workbook) {
  if (!(workbook instanceof Workbook)) throw new TypeError("exportXlsxWithOpenChestnut expects a Workbook instance.");
  if (!workbook.worksheets?.items?.length) throw new OpenChestnutCodecError("Workbook must contain at least one worksheet.", [], { code: "missing_worksheets" });
  if (!workbook.worksheets.items.some((sheet) => sheet.visibility === "visible")) throw new OpenChestnutCodecError("Workbook must contain at least one visible worksheet.", [], { code: "missing_visible_worksheet" });
  const state = workbook[WORKBOOK_STATE];
  assertTrustedImportedState(state, "XLSX");
  const unsupported = unsupportedWorkbookFeatures(workbook, state);
  if (unsupported.length) {
    throw new OpenChestnutCodecError(`OpenChestnut cannot encode these XLSX features: ${unsupported.slice(0, 8).join(", ")}${unsupported.length > 8 ? `, and ${unsupported.length - 8} more` : ""}. This operation fails closed; preserve them only through a validated source-bound package.`, [], { code: "unsupported_workbook_features" });
  }
  const theme = state?.themeWire && sameWorkbookTheme(workbook.theme, state.publicTheme)
    ? state.themeWire
    : wireWorkbookTheme(workbook.theme, state?.themeWire?.source);
  const views = wireWorkbookViews(workbook, state);
  const assets = new Map();
  const worksheets = workbook.worksheets.items.map((sheet) => {
    const metadata = wireWorksheetMetadata(sheet, state?.worksheetSlots?.get(sheet.id));
    const cells = (() => {
      const dynamicSlots = state?.dynamicArraySlotsBySheet?.get(sheet.id) || new Map();
      const sourceBoundFormulaSlots = state?.sourceBoundFormulaSlotsBySheet?.get(sheet.id) || new Map();
      const entries = sheet.store?.entries?.() || [];
      const byAddress = new Map(entries);
      for (const [address, slot] of dynamicSlots) {
        if (byAddress.get(address) !== slot.cell || JSON.stringify(dynamicArrayCellSnapshot(slot.cell)) !== JSON.stringify(slot.publicSnapshot)) {
          throw new OpenChestnutCodecError(`Imported dynamic array ${sheet.name}!${address} is source-bound and read-only in OpenChestnut 0.2.`, [], { code: "unsupported_dynamic_array_edit" });
        }
      }
      for (const [address, slot] of sourceBoundFormulaSlots) {
        if (byAddress.get(address) !== slot.cell || JSON.stringify(sourceBoundFormulaCellSnapshot(slot.cell)) !== JSON.stringify(slot.publicSnapshot)) {
          throw new OpenChestnutCodecError(`Imported partial shared formula ${sheet.name}!${address} is source-bound and read-only in OpenChestnut 0.2.`, [], { code: "unsupported_cell_formula_edit" });
        }
      }
      const output = entries
        .filter(([, cell]) => cell.value != null || cell.formula || cell.formulaType || Object.keys(cell.style || {}).some((key) => cell.style[key] != null))
        .map(([address, cell]) => sourceBoundFormulaSlots.get(address)?.wire || dynamicSlots.get(address)?.wire || wireCell(address, cell, workbook.dateSystem));
      wireWorksheetDataTables(sheet, state?.dataTablesBySheet?.get(sheet.id), output);
      validateFormulaTopology(output, sheet.name);
      return output;
    })();
    const pivotTables = wireWorksheetPivots(workbook, sheet, state?.pivotsBySheet?.get(sheet.id), cells);
    return {
      id: sheet.id,
      name: sheet.name,
      visibility: metadata.visibility,
      source: metadata.source,
      showGridLines: sheet.showGridLines !== false,
      freezePane: {
        rows: sheet.freezePanes?.rows || 0,
        columns: sheet.freezePanes?.columns || 0,
        topLeftCell: sheet.freezePanes?.topLeftCell || "",
        activePane: sheet.freezePanes?.activePane || "",
      },
      columnDimensions: [...(sheet.columnDimensions || new Map())].map(([column, dimension]) => ({ column, width: dimension.width || 0, hidden: Boolean(dimension.hidden), bestFit: Boolean(dimension.bestFit) })),
      rowDimensions: [...(sheet.rowDimensions || new Map())].map(([row, dimension]) => ({ row, height: dimension.height || 0, hidden: Boolean(dimension.hidden) })),
      mergedRanges: [...(sheet.mergedRanges || [])],
      sortState: wireTableSortState(sheet.sortState, `worksheet ${sheet.name}`),
      tables: wireWorksheetTables(sheet, state?.tablesBySheet?.get(sheet.id)),
      images: wireWorksheetImages(sheet, state?.imagesBySheet?.get(sheet.id), assets),
      charts: wireWorksheetCharts(sheet, state?.chartsBySheet?.get(sheet.id)),
      sparklineGroups: wireWorksheetSparklines(sheet, state?.sparklinesBySheet?.get(sheet.id)),
      pivotTables,
      dataValidations: (sheet.dataValidations?.items || []).map((item) => wireDataValidation(item, sheet.name)),
      conditionalFormats: (sheet.conditionalFormattings?.items || []).map((item, index) => wireConditionalFormat(item, sheet.name, index)),
      threadedComments: wireThreadedComments(workbook, sheet),
      cells,
    };
  });
  return {
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    family: ArtifactFamily.WORKBOOK,
    source: state?.source,
    opaqueOpc: state?.opaqueOpc,
    assets: [...assets.values()],
    diagnostics: state?.diagnostics || [],
    payload: {
      case: "workbook",
      value: {
        id: workbook.id,
        dateSystem: workbook.dateSystem === "1904" ? WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1904 : WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1900,
        theme,
        connections: wireWorkbookConnections(workbook, state),
        definedNames: wireWorkbookDefinedNames(workbook, state),
        calculation: wireWorkbookCalculationForExport(workbook, state),
        view: views.view,
        additionalViews: views.additionalViews,
        worksheets,
      },
    },
  };
}

export async function exportXlsxWithOpenChestnut(workbook, options = {}) {
  assertCodecOptions(options, new Set(["limits", "recalculate"]), "exportXlsxWithOpenChestnut");
  if (!(workbook instanceof Workbook)) throw new TypeError("exportXlsxWithOpenChestnut expects a Workbook instance.");
  if (options.recalculate !== false) workbook.recalculate();
  const response = await invokeOpenChestnut({
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    operation: CodecOperation.EXPORT_XLSX,
    family: ArtifactFamily.WORKBOOK,
    artifact: workbookEnvelope(workbook),
    limits: codecLimits(options.limits),
  });
  return new FileBlob(response.file, {
    type: XLSX_MIME,
    metadata: { artifactKind: "workbook", codec: "open-chestnut", diagnostics: response.diagnostics },
  });
}

function workbookFromEnvelope(envelope) {
  if (envelope.family !== ArtifactFamily.WORKBOOK || envelope.payload.case !== "workbook") {
    throw new OpenChestnutCodecError("OpenChestnut response does not contain a workbook artifact.", [], { code: "invalid_workbook_artifact" });
  }
  const source = envelope.payload.value;
  const importedTheme = workbookThemeFromWire(source.theme);
  const importedConnections = (source.connections || []).map(publicWorkbookConnection);
  const importedCalculation = publicWorkbookCalculation(source.calculation);
  const workbook = Workbook.create({
    dateSystem: source.dateSystem === WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1904 ? "1904" : "1900",
    ...(importedTheme ? { theme: importedTheme } : {}),
    connections: importedConnections,
    ...(importedCalculation !== undefined ? { calculation: importedCalculation } : {}),
  });
  workbook.id = source.id || workbook.id;
  const tablesBySheet = new Map();
  const imagesBySheet = new Map();
  const chartsBySheet = new Map();
  const sparklinesBySheet = new Map();
  let pivotsBySheet;
  const dataTablesBySheet = new Map();
  const dynamicArraySlotsBySheet = new Map();
  const sourceBoundFormulaSlotsBySheet = new Map();
  const worksheetSlots = new Map();
  const partialSharedFormulaRangesBySheetName = partialSharedFormulaRanges(envelope.diagnostics);
  const assets = new Map((envelope.assets || []).map((asset) => [asset.id, asset]));
  const connectionSlots = (source.connections || []).map((wire, index) => ({
    wire,
    connection: workbook.connections[index],
    publicSnapshot: connectionSnapshot(workbook.connections[index]),
  }));
  for (const sourceSheet of source.worksheets) {
    const sheet = workbook.worksheets.add(sourceSheet.name, { visibility: publicWorksheetVisibility(sourceSheet.visibility) });
    sheet.id = sourceSheet.id || sheet.id;
    worksheetSlots.set(sheet.id, { wire: sourceSheet, publicSnapshot: worksheetMetadataSnapshot(sheet) });
    sheet.showGridLines = sourceSheet.showGridLines;
    if (sourceSheet.freezePane) {
      sheet.freezePanes.freezeRows(sourceSheet.freezePane.rows);
      sheet.freezePanes.freezeColumns(sourceSheet.freezePane.columns);
    }
    for (const dimension of sourceSheet.columnDimensions) sheet.columnDimensions.set(dimension.column, { width: dimension.width || undefined, hidden: dimension.hidden, bestFit: dimension.bestFit });
    for (const dimension of sourceSheet.rowDimensions) sheet.rowDimensions.set(dimension.row, { height: dimension.height || undefined, hidden: dimension.hidden });
    sheet.mergedRanges = [...sourceSheet.mergedRanges];
    sheet.sortState = publicTableSortState(sourceSheet.sortState);
    sheet.dataValidations.items = (sourceSheet.dataValidations || []).map(publicDataValidation);
    sheet.conditionalFormattings.items = (sourceSheet.conditionalFormats || []).map(publicConditionalFormat);
    const sourceComments = sourceSheet.threadedComments || [];
    const rootComments = sourceComments.filter((item) => !item.parentNativeCommentId);
    const consumedReplies = new Set();
    for (const sourceComment of rootComments) {
      const thread = workbook.comments.addThread(
        { sheetName: sheet.name, address: sourceComment.cellReference },
        sourceComment.text,
        {
          id: sourceComment.id || undefined,
          author: sourceComment.author || "User",
          resolved: sourceComment.resolved,
          comment: publicThreadedComment(sourceComment),
        },
      );
      if (sourceComment.id) thread.id = sourceComment.id;
      for (const reply of sourceComments.filter((item) => item.parentNativeCommentId === sourceComment.nativeCommentId)) {
        thread.addReply(reply.text, publicThreadedComment(reply));
        consumedReplies.add(reply);
      }
    }
    if (rootComments.length + consumedReplies.size !== sourceComments.length) {
      throw new OpenChestnutCodecError(`Worksheet ${sheet.name} contains an unsupported threaded-comment reply graph.`, [], { code: "unsupported_threaded_comment_reply_topology" });
    }
    const dynamicArraySlots = new Map();
    const sourceBoundFormulaSlots = new Map();
    const partialSharedFormulaRanges = partialSharedFormulaRangesBySheetName.get(sourceSheet.name) || [];
    const dataTableSlots = [];
    for (const sourceCell of sourceSheet.cells) {
      const address = cellAddress(sourceCell.row, sourceCell.column);
      const cell = sheet.store.get(address);
      cell.formula = sourceCell.formula || null;
      if (sourceCell.formulaMetadata?.kind === CellFormulaKind.SHARED) {
        cell.formulaType = "shared";
        cell.sharedIndex = sourceCell.formulaMetadata.sharedIndex;
        cell.sharedRef = sourceCell.formulaMetadata.reference;
      } else if (sourceCell.formulaMetadata?.kind === CellFormulaKind.ARRAY) {
        cell.formulaType = "array";
        cell.arrayRef = sourceCell.formulaMetadata.reference;
      } else if (sourceCell.formulaMetadata?.kind === CellFormulaKind.DYNAMIC_ARRAY) {
        cell.formulaType = "dynamicArray";
        cell.dynamicArrayRef = sourceCell.formulaMetadata.reference;
      }
      const staticStyle = cellStyleFromWire(sourceCell.style);
      if (staticStyle || sourceCell.numberFormatCode) cell.style = { ...(staticStyle || {}), ...(sourceCell.numberFormatCode ? { numberFormat: sourceCell.numberFormatCode } : {}) };
      switch (sourceCell.value.case) {
        case "stringValue": cell.value = sourceCell.value.value; break;
        case "numberValue": cell.value = sourceCell.value.value; break;
        case "boolValue": cell.value = sourceCell.value.value; break;
        case "errorValue": cell.value = sourceCell.value.value; break;
        default: cell.value = null;
      }
      if (partialSharedFormulaRanges.some(({ bounds }) =>
        sourceCell.row >= bounds.top && sourceCell.row <= bounds.bottom &&
        sourceCell.column >= bounds.left && sourceCell.column <= bounds.right)) {
        sourceBoundFormulaSlots.set(address, { wire: sourceCell, cell, publicSnapshot: sourceBoundFormulaCellSnapshot(cell) });
      }
      if (sourceCell.formulaMetadata?.kind === CellFormulaKind.DYNAMIC_ARRAY) {
        dynamicArraySlots.set(address, { wire: sourceCell, cell, publicSnapshot: dynamicArrayCellSnapshot(cell) });
      } else if (sourceCell.formulaMetadata?.kind === CellFormulaKind.DATA_TABLE) {
        dataTableSlots.push(hydrateWorksheetDataTable(sheet, sourceCell));
      }
    }
    dynamicArraySlotsBySheet.set(sheet.id, dynamicArraySlots);
    if (partialSharedFormulaRanges.length && sourceBoundFormulaSlots.size === 0) {
      throw new OpenChestnutCodecError(`OpenChestnut reported partial shared formulas for ${sourceSheet.name} but returned no matching cells.`, [], { code: "invalid_open_chestnut_diagnostic" });
    }
    sourceBoundFormulaSlotsBySheet.set(sheet.id, sourceBoundFormulaSlots);
    dataTablesBySheet.set(sheet.id, { slots: dataTableSlots });
    const slots = [];
    for (const sourceTable of sourceSheet.tables || []) {
      if (!sourceTable.name || !sourceTable.reference || !sourceTable.columnNames?.length) {
        slots.push({ wire: sourceTable });
        continue;
      }
      const table = sheet.tables.add({
        id: sourceTable.id,
        range: sourceTable.reference,
        name: sourceTable.name,
        hasHeaders: sourceTable.hasHeaders,
        showTotals: sourceTable.showTotals,
        showFilterButton: sourceTable.showFilterButton,
        showBandedColumns: sourceTable.showColumnStripes,
        style: sourceTable.styleName,
        columnNames: [...sourceTable.columnNames],
        columnDefinitions: sourceTable.columns?.length ? sourceTable.columns.map((column) => ({ ...column })) : undefined,
        filters: sourceTable.filters?.map(publicTableFilter),
        sortState: publicTableSortState(sourceTable.sortState),
        queryTable: publicTableQuery(sourceTable.queryTable),
      });
      table.showHeaders = sourceTable.hasHeaders;
      table.showFirstColumn = sourceTable.showFirstColumn;
      table.showLastColumn = sourceTable.showLastColumn;
      table.showRowStripes = sourceTable.showRowStripes;
      const publicSnapshot = tableSnapshot(table);
      Object.defineProperty(table, TABLE_STATE, { configurable: true, value: { wire: sourceTable, querySnapshot: publicTableQuery(table.queryTable) }, writable: true });
      slots.push({ wire: sourceTable, table, publicSnapshot });
    }
    tablesBySheet.set(sheet.id, { slots });
    const imageSlots = [];
    for (const sourceImage of sourceSheet.images || []) {
      const image = spreadsheetImageFromWire(sheet, sourceImage, assets);
      imageSlots.push({ wire: sourceImage, image, publicSnapshot: spreadsheetImageSnapshot(image) });
    }
    imagesBySheet.set(sheet.id, { slots: imageSlots });
    const chartSlots = [];
    for (const sourceChart of sourceSheet.charts || []) {
      const chart = spreadsheetChartFromWire(sheet, sourceChart);
      chartSlots.push({ wire: sourceChart, chart, publicSnapshot: spreadsheetChartSnapshot(chart) });
    }
    chartsBySheet.set(sheet.id, { slots: chartSlots });
    const sparklineSlots = [];
    for (const sourceSparkline of sourceSheet.sparklineGroups || []) {
      const group = spreadsheetSparklineFromWire(sheet, sourceSparkline);
      sparklineSlots.push({ wire: sourceSparkline, group, publicSnapshot: spreadsheetSparklineSnapshot(group) });
    }
    sparklinesBySheet.set(sheet.id, { slots: sparklineSlots });
  }
  for (const sourceDefinedName of source.definedNames || []) workbook.definedNames.add(publicWorkbookDefinedName(sourceDefinedName));
  pivotsBySheet = hydrateWorkbookPivots(workbook, source.worksheets);
  const sourceViews = source.view ? [source.view, ...(source.additionalViews || [])] : [];
  if (sourceViews.length) {
    workbook.windows.getItemAt(0).setActiveWorksheet(sourceViews[0].activeWorksheetId);
    if (sourceViews[0].selectedWorksheetIds?.length) workbook.windows.getItemAt(0).setSelectedWorksheets(sourceViews[0].selectedWorksheetIds);
    for (const sourceView of sourceViews.slice(1)) {
      const window = workbook.windows.add({ activeWorksheet: sourceView.activeWorksheetId });
      if (sourceView.selectedWorksheetIds?.length) window.setSelectedWorksheets(sourceView.selectedWorksheetIds);
    }
  }
  const definedNameSlots = (source.definedNames || []).map((wire, index) => ({
    wire,
    definedName: workbook.definedNames.items[index],
    publicSnapshot: definedNameSnapshot(workbook.definedNames.items[index]),
  }));
  const calculationSlot = source.calculation ? { wire: source.calculation, publicSnapshot: calculationSnapshot(workbook.calculation) } : undefined;
  const snapshots = sourceViews.length ? workbookViewSnapshots(workbook) : [];
  const viewSlots = sourceViews.map((wire, index) => ({ wire, publicSnapshot: snapshots[index] }));
  Object.defineProperty(workbook, WORKBOOK_STATE, {
    configurable: true,
    value: {
      source: envelope.source,
      opaqueOpc: envelope.opaqueOpc,
      diagnostics: envelope.diagnostics,
      themeWire: source.theme,
      publicTheme: normalizeXlsxThemeConfig(workbook.theme),
      connectionSlots,
      definedNameSlots,
      calculationSlot,
      viewSlots,
      worksheetSlots,
      dynamicArraySlotsBySheet,
      sourceBoundFormulaSlotsBySheet,
      tablesBySheet,
      imagesBySheet,
      chartsBySheet,
      sparklinesBySheet,
      pivotsBySheet,
      dataTablesBySheet,
    },
    writable: true,
  });
  return workbook;
}

export async function importXlsxWithOpenChestnut(input, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "importXlsxWithOpenChestnut");
  const response = await invokeOpenChestnut({
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    operation: CodecOperation.IMPORT_XLSX,
    family: ArtifactFamily.WORKBOOK,
    file: await inputBytes(input),
    limits: codecLimits(options.limits),
  });
  return workbookFromEnvelope(response.artifact);
}

const DOCUMENT_RUN_STYLE_KEYS = new Set(["runStyleId", "bold", "italic", "underline", "fontFamily", "fontSize", "color", "characterSpacing", "characterSpacingTwips"]);
const DOCUMENT_RUN_DERIVED_STYLE_KEYS = new Set(["resolvedColor", "resolvedFontFamily", "resolvedFontFamilyEastAsia", "resolvedFontFamilyComplexScript"]);
const DOCUMENT_BIBLIOGRAPHY_FIELD_KEYS = [
  "title", "year", "city", "stateProvince", "countryRegion", "publisher", "bookTitle", "journalName", "periodicalTitle", "publicationTitle", "internetSiteTitle",
  "conferenceName", "institution", "department", "volume", "issue", "pages", "edition", "numberVolumes", "chapterNumber", "standardNumber", "shortTitle", "comments", "medium",
  "month", "day", "yearAccessed", "monthAccessed", "dayAccessed", "url", "guid", "lcid", "reporter", "caseNumber", "abbreviatedCaseNumber", "court", "patentNumber", "patentType",
  "broadcaster", "broadcastTitle", "station", "theater", "productionCompany", "distributor", "recordingNumber", "albumTitle", "thesisType", "version", "referenceOrder",
];
const DOCUMENT_CITATION_TAG = /^[A-Za-z0-9_.:-]{1,255}$/;
const DOCUMENT_FIELD_COMMANDS = new Set(["PAGE", "NUMPAGES", "SECTION", "SECTIONPAGES", "DATE", "TIME", "CREATEDATE", "SAVEDATE", "PRINTDATE", "AUTHOR", "TITLE", "SUBJECT", "COMMENTS", "FILENAME", "FILESIZE", "NUMWORDS", "NUMCHARS"]);
const DOCUMENT_INLINE_FIELD_INSTRUCTION = /^(?:SEQ [A-Za-z][A-Za-z0-9_]{0,39} \\[*] ARABIC|(?:REF|PAGEREF) [A-Za-z][A-Za-z0-9_]{0,39} \\h)$/;

function documentRgb(value, label) {
  if (value == null || value === "") return undefined;
  const rgb = String(value).replace(/^#/, "").toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(rgb)) throw new OpenChestnutCodecError(`${label} color must be a six-digit RGB value.`, [], { code: "invalid_document_formatting" });
  return rgb;
}

function documentRunFormatting(style = {}, label = "Document run") {
  const unsupported = Object.keys(style).filter((key) => !DOCUMENT_RUN_STYLE_KEYS.has(key) && !DOCUMENT_RUN_DERIVED_STYLE_KEYS.has(key));
  if (unsupported.length) throw new OpenChestnutCodecError(`${label} uses unsupported run style fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_document_features" });
  const formatting = {};
  if (Object.hasOwn(style, "fontFamily")) formatting.fontFamily = String(style.fontFamily || "");
  if (Object.hasOwn(style, "fontSize")) {
    const points = Number(style.fontSize);
    if (!Number.isFinite(points) || points <= 0 || points > 1_638) throw new OpenChestnutCodecError(`${label} fontSize must be greater than 0 and no more than 1638 points.`, [], { code: "invalid_document_formatting" });
    formatting.fontSizeHalfPoints = uint32(Math.round(points * 2), `${label} fontSize`);
  }
  if (Object.hasOwn(style, "color")) formatting.colorRgb = documentRgb(style.color, label);
  if (Object.hasOwn(style, "characterSpacing") || Object.hasOwn(style, "characterSpacingTwips")) {
    const value = Number(style.characterSpacingTwips ?? style.characterSpacing);
    if (!Number.isInteger(value) || value < -31_680 || value > 31_680) throw new OpenChestnutCodecError(`${label} character spacing must be an integer from -31680 through 31680 twips.`, [], { code: "invalid_document_formatting" });
    formatting.characterSpacingTwips = value;
  }
  for (const key of ["bold", "italic"]) if (Object.hasOwn(style, key)) formatting[key] = Boolean(style[key]);
  if (Object.hasOwn(style, "underline")) formatting.underline = style.underline === true || style.underline === "single";
  return Object.keys(formatting).length ? formatting : undefined;
}

function publicDocumentRunFormatting(formatting) {
  if (!formatting) return {};
  return {
    ...(formatting.fontFamily !== undefined ? { fontFamily: formatting.fontFamily } : {}),
    ...(formatting.fontSizeHalfPoints !== undefined ? { fontSize: formatting.fontSizeHalfPoints / 2 } : {}),
    ...(formatting.colorRgb !== undefined ? { color: `#${formatting.colorRgb}` } : {}),
    ...(formatting.characterSpacingTwips !== undefined ? { characterSpacingTwips: formatting.characterSpacingTwips } : {}),
    ...(formatting.bold !== undefined ? { bold: formatting.bold } : {}),
    ...(formatting.italic !== undefined ? { italic: formatting.italic } : {}),
    ...(formatting.underline !== undefined ? { underline: formatting.underline } : {}),
  };
}

function documentParagraphFormatting(block) {
  const value = block?.paragraphFormat || block?.formatting || {};
  const result = {};
  const text = (model, wire) => { if (value[model] != null) result[wire] = String(value[model]); };
  const integer = (model, wire) => {
    if (value[model] == null) return;
    const number = Number(value[model]);
    if (!Number.isInteger(number) || number < -1_000_000 || number > 1_000_000) throw new OpenChestnutCodecError(`Document paragraph ${block.id} ${model} must be a bounded integer.`, [], { code: "invalid_document_formatting" });
    result[wire] = number;
  };
  text("alignment", "alignment");
  for (const [model, wire] of [["leftIndentTwips", "leftIndentTwips"], ["rightIndentTwips", "rightIndentTwips"], ["firstLineIndentTwips", "firstLineIndentTwips"], ["hangingIndentTwips", "hangingIndentTwips"], ["spaceBeforeTwips", "spaceBeforeTwips"], ["spaceAfterTwips", "spaceAfterTwips"], ["lineSpacingTwips", "lineSpacingTwips"]]) integer(model, wire);
  text("lineSpacingRule", "lineSpacingRule");
  if (value.keepNext != null) result.keepNext = Boolean(value.keepNext);
  if (value.pageBreakBefore != null) result.pageBreakBefore = Boolean(value.pageBreakBefore);
  return Object.keys(result).length ? result : undefined;
}

function publicDocumentParagraphFormatting(value) {
  if (!value) return undefined;
  const result = {};
  for (const key of ["alignment", "leftIndentTwips", "rightIndentTwips", "firstLineIndentTwips", "hangingIndentTwips", "spaceBeforeTwips", "spaceAfterTwips", "lineSpacingTwips", "lineSpacingRule", "keepNext", "pageBreakBefore"]) {
    if (value[key] !== undefined) result[key] = value[key];
  }
  return Object.keys(result).length ? result : undefined;
}

function planDocumentTextContentControls(document) {
  const controls = document.blocks.flatMap((block) => block.kind === "paragraph"
    ? block.runs.filter((run) => run.contentControl).map((run) => ({ block, run, control: run.contentControl }))
    : []);
  const used = new Set();
  for (const { block, control } of controls) {
    if (!control.id || !String(control.id).trim()) throw new OpenChestnutCodecError(`Document block ${block.id} content control requires a non-empty model ID.`, [], { code: "invalid_document_content_control" });
    const nativeId = control.nativeId == null ? undefined : Number(control.nativeId);
    if (nativeId === undefined) continue;
    if (!Number.isInteger(nativeId) || nativeId < 1 || nativeId > 0x7fffffff || used.has(nativeId)) throw new OpenChestnutCodecError(`Document block ${block.id} content control ${control.id} has an invalid or duplicate nativeId.`, [], { code: "invalid_document_content_control" });
    used.add(nativeId);
  }
  const result = new Map();
  let next = 1;
  for (const { run, control } of controls) {
    if (control.nativeId != null) {
      result.set(run, Number(control.nativeId));
      continue;
    }
    while (used.has(next)) next += 1;
    if (next > 0x7fffffff) throw new OpenChestnutCodecError("Document content controls exhausted the positive native ID range.", [], { code: "invalid_document_content_control" });
    result.set(run, next);
    used.add(next);
    next += 1;
  }
  return result;
}

function wireDocumentTextContentControl(control, nativeId, blockId) {
  const id = String(control?.id || "").trim();
  const tag = String(control?.tag || "").trim();
  const alias = String(control?.alias ?? tag);
  if (!id || !tag || tag.length > 64 || alias.length > 255 || /[\u0000-\u001f\u007f]/.test(tag + alias)) throw new OpenChestnutCodecError(`Document block ${blockId} has an invalid plain-text content control.`, [], { code: "invalid_document_content_control" });
  return { id, tag, alias, nativeId };
}

function documentRun(run, blockId, contentControlNativeId) {
  const style = run.style || {};
  const formatting = documentRunFormatting(style, `Document block ${blockId}`);
  const inlineInstruction = run.inlineField ? String(run.inlineField.instruction || "").trim() : undefined;
  if (inlineInstruction !== undefined && !DOCUMENT_INLINE_FIELD_INSTRUCTION.test(inlineInstruction)) {
    throw new OpenChestnutCodecError(`Document block ${blockId} inline field must be canonical SEQ <label> \\* ARABIC, REF <bookmark> \\h, or PAGEREF <bookmark> \\h.`, [], { code: "invalid_document_inline_field" });
  }
  if (run.contentControl && inlineInstruction !== undefined) throw new OpenChestnutCodecError(`Document block ${blockId} run cannot combine a content control and an inline field.`, [], { code: "invalid_document_inline_field" });
  const bookmarkName = inlineInstruction === undefined ? "" : String(run.inlineField?.bookmarkName || "").trim();
  let bookmarkNativeId = "";
  if (run.inlineField?.bookmarkNativeId !== undefined) {
    const value = Number(run.inlineField.bookmarkNativeId);
    if (!Number.isInteger(value) || value < 0 || value > 0xffffffff) throw new OpenChestnutCodecError(`Document block ${blockId} inline field bookmarkNativeId must be an unsigned 32-bit integer.`, [], { code: "invalid_document_inline_field" });
    bookmarkNativeId = String(value);
  }
  if (bookmarkName && (!/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(bookmarkName) || !inlineInstruction.startsWith("SEQ "))) {
    throw new OpenChestnutCodecError(`Document block ${blockId} may bookmark only a canonical SEQ cached result with a valid Word bookmark name.`, [], { code: "invalid_document_inline_field" });
  }
  if (!bookmarkName && bookmarkNativeId) throw new OpenChestnutCodecError(`Document block ${blockId} inline field bookmarkNativeId requires bookmarkName.`, [], { code: "invalid_document_inline_field" });
  return {
    text: String(run.text ?? ""),
    styleId: style.runStyleId || "",
    bold: style.bold === true,
    italic: style.italic === true,
    underline: style.underline === true || style.underline === "single",
    formatting,
    textContentControl: run.contentControl ? wireDocumentTextContentControl(run.contentControl, contentControlNativeId, blockId) : undefined,
    inlineField: inlineInstruction === undefined ? undefined : { instruction: inlineInstruction, bookmarkName, bookmarkNativeId },
  };
}

function documentContentControlTopology(runs = []) {
  return runs.flatMap((run, index) => run.textContentControl || run.contentControl
    ? [{ index, nativeId: Number((run.textContentControl || run.contentControl).nativeId) }]
    : []);
}

function assertDocumentContentControlTopology(block, original) {
  if (!original || original.content.case !== "paragraph") return;
  const requested = documentContentControlTopology(block.runs || []);
  const source = documentContentControlTopology(original.content.value.runs || []);
  if (JSON.stringify(requested) !== JSON.stringify(source)) {
    throw new OpenChestnutCodecError(`Imported document paragraph ${block.id} plain-text content-control topology is source-bound.`, [], { code: "document_content_control_topology_changed" });
  }
}

function documentInlineFieldTopology(runs = []) {
  return runs.flatMap((run, index) => {
    const field = run.inlineField || run.field;
    return field ? [{
      index,
      instruction: String(field.instruction || "").trim(),
      bookmarkName: String(field.bookmarkName || ""),
      bookmarkNativeId: field.bookmarkNativeId === undefined || field.bookmarkNativeId === "" ? undefined : Number(field.bookmarkNativeId),
    }] : [];
  });
}

function assertDocumentInlineFieldTopology(block, original) {
  if (!original || original.content.case !== "paragraph") return;
  const requested = documentInlineFieldTopology(block.runs || []);
  const source = documentInlineFieldTopology(original.content.value.runs || []);
  if (JSON.stringify(requested) !== JSON.stringify(source)) {
    throw new OpenChestnutCodecError(`Imported document paragraph ${block.id} inline-field positions and instructions are source-bound.`, [], { code: "document_inline_field_topology_changed" });
  }
}

function sameTableValues(block, original) {
  return JSON.stringify(block.values || []) === JSON.stringify((original.content.value.rows || []).map((row) => [...row.cells]));
}

function documentTableCells(table) {
  const mergeName = (value) => value === DocumentTableVerticalMerge.RESTART
    ? "restart"
    : value === DocumentTableVerticalMerge.CONTINUE ? "continue" : "none";
  return table.rows.flatMap((row, rowIndex) => row.richCells.map((cell, column) => ({
    row: rowIndex,
    column,
    gridColumn: cell.gridColumn,
    columnSpan: cell.columnSpan || 1,
    rowSpan: cell.rowSpan,
    verticalMerge: mergeName(cell.verticalMerge),
    editable: cell.editable,
    textPatchable: cell.textPatchable,
  })));
}

function sameDocumentTableGeometry(block, table) {
  if (block.gridColumns !== table.gridColumns) return false;
  const sourceCells = documentTableCells(table);
  if (!Array.isArray(block.cells) || block.cells.length !== sourceCells.length) return false;
  return block.cells.every((cell, index) => {
    const source = sourceCells[index];
    return cell.row === source.row && cell.column === source.column &&
      cell.gridColumn === source.gridColumn && cell.columnSpan === source.columnSpan &&
      cell.rowSpan === source.rowSpan && cell.verticalMerge === source.verticalMerge &&
      cell.editable === source.editable && cell.textPatchable === source.textPatchable;
  });
}

function wireDocumentTableTextPatches(block, source) {
  const patches = Array.isArray(block.textPatches) ? block.textPatches : [];
  if (!patches.length) return [];
  if (!source) throw new OpenChestnutCodecError(`Document table ${block.id} text patches require a validated imported source.`, [], { code: "unsupported_document_edit" });
  if (patches.length > 10_000) throw new OpenChestnutCodecError(`Document table ${block.id} exceeds 10,000 source text patches.`, [], { code: "invalid_document_table" });
  return patches.map((patch) => {
    const row = Number(patch.row);
    const column = Number(patch.column);
    const sourceRow = source.rows?.[row];
    const sourceCell = sourceRow?.richCells?.[column];
    if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || column < 0 || !sourceCell) {
      throw new OpenChestnutCodecError(`Document table ${block.id} text patch ${row},${column} is outside the source cell matrix.`, [], { code: "invalid_document_table" });
    }
    if (!sourceCell.textPatchable) {
      throw new OpenChestnutCodecError(`Document table ${block.id} cell ${row},${column} does not advertise source-bound text replacement capability.`, [], { code: "unsupported_document_edit" });
    }
    const search = String(patch.search ?? "");
    const replacement = String(patch.replacement ?? "");
    if (!search || search.length > 1_000_000 || replacement.length > 1_000_000 || !isXmlSafeText(search) || !isXmlSafeText(replacement)) {
      throw new OpenChestnutCodecError(`Document table ${block.id} cell ${row},${column} text patch requires bounded XML-safe strings.`, [], { code: "invalid_document_table" });
    }
    const sourceText = String(sourceRow.cells[column] ?? "");
    return {
      row,
      column,
      search,
      replacement,
      sourceTextSha256: createHash("sha256").update(sourceText, "utf8").digest("hex"),
    };
  });
}

function authoredDocumentTableGeometry(block) {
  const invalid = (message) => {
    throw new OpenChestnutCodecError(`Document table ${block.id} ${message}`, [], { code: "invalid_document_table" });
  };
  if (!Array.isArray(block.cells) || block.cells.length === 0) invalid("requires one explicit geometry record for every physical cell.");
  if (!Number.isInteger(block.gridColumns) || block.gridColumns < 1 || block.gridColumns > 4_096) {
    invalid("gridColumns must be an integer from 1 through 4096.");
  }

  const records = new Map();
  for (const cell of block.cells) {
    if (!Number.isInteger(cell.row) || !Number.isInteger(cell.column) || cell.row < 0 || cell.column < 0 ||
        cell.row >= block.values.length || cell.column >= (block.values[cell.row]?.length || 0)) {
      invalid(`cell ${cell.row},${cell.column} does not identify a physical value cell.`);
    }
    const key = `${cell.row}:${cell.column}`;
    if (records.has(key)) invalid(`contains duplicate geometry for cell ${cell.row},${cell.column}.`);
    records.set(key, cell);
  }

  const rows = block.values.map((values, rowIndex) => {
    if (values.length === 0) invalid(`row ${rowIndex} has no physical cells.`);
    let cursor;
    const richCells = values.map((_value, column) => {
      const source = records.get(`${rowIndex}:${column}`);
      if (!source) invalid(`is missing geometry for cell ${rowIndex},${column}.`);
      if (!Number.isInteger(source.gridColumn) || source.gridColumn < 0 || source.gridColumn > 4_096 ||
          !Number.isInteger(source.columnSpan) || source.columnSpan < 1 || source.columnSpan > 4_096) {
        invalid(`cell ${rowIndex},${column} has invalid bounded grid geometry.`);
      }
      if (cursor !== undefined && source.gridColumn !== cursor) {
        invalid(`cell ${rowIndex},${column} must begin at grid column ${cursor}, not ${source.gridColumn}.`);
      }
      const end = source.gridColumn + source.columnSpan;
      if (end > block.gridColumns) invalid(`cell ${rowIndex},${column} extends beyond gridColumns ${block.gridColumns}.`);
      cursor = end;
      const verticalMerge = String(source.verticalMerge || "none");
      const merge = verticalMerge === "restart"
        ? DocumentTableVerticalMerge.RESTART
        : verticalMerge === "continue" ? DocumentTableVerticalMerge.CONTINUE
          : verticalMerge === "none" ? DocumentTableVerticalMerge.UNSPECIFIED : undefined;
      if (merge === undefined) invalid(`cell ${rowIndex},${column} has unsupported verticalMerge ${verticalMerge}.`);
      const rowSpan = Number(source.rowSpan);
      if (!Number.isInteger(rowSpan) || rowSpan < 0 || rowSpan > 4_096) invalid(`cell ${rowIndex},${column} has invalid rowSpan.`);
      if (verticalMerge === "continue") {
        if (rowSpan !== 0 || String(values[column] ?? "") !== "") invalid(`continuation cell ${rowIndex},${column} must have rowSpan 0 and empty text.`);
      } else {
        if (rowSpan < 1 || source.editable === false) invalid(`origin cell ${rowIndex},${column} must have a positive rowSpan and remain editable.`);
        if (verticalMerge === "none" && rowSpan !== 1) invalid(`unmerged cell ${rowIndex},${column} must have rowSpan 1.`);
      }
      return {
        gridColumn: source.gridColumn,
        columnSpan: source.columnSpan,
        rowSpan,
        verticalMerge: merge,
        editable: verticalMerge !== "continue",
      };
    });
    const gridBefore = richCells[0].gridColumn;
    const gridAfter = block.gridColumns - cursor;
    return { cells: values.map((value) => String(value ?? "")), richCells, gridBefore, gridAfter };
  });
  if (records.size !== block.values.reduce((total, row) => total + row.length, 0)) invalid("contains geometry outside the physical value matrix.");

  let active = new Map();
  const finish = (group) => {
    if (group.seen !== group.expected) invalid(`merge origin ${group.row},${group.column} declares rowSpan ${group.expected} but spans ${group.seen} rows.`);
  };
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const continued = new Map();
    for (let column = 0; column < rows[rowIndex].richCells.length; column += 1) {
      const cell = rows[rowIndex].richCells[column];
      const key = `${cell.gridColumn}:${cell.columnSpan}`;
      if (cell.verticalMerge === DocumentTableVerticalMerge.CONTINUE) {
        const group = active.get(key);
        if (!group) invalid(`continuation cell ${rowIndex},${column} has no matching restart in the preceding row.`);
        group.seen += 1;
        continued.set(key, group);
      } else if (cell.verticalMerge === DocumentTableVerticalMerge.RESTART) {
        continued.set(key, { row: rowIndex, column, expected: cell.rowSpan, seen: 1 });
      }
    }
    for (const [key, group] of active) if (!continued.has(key) || continued.get(key) !== group) finish(group);
    active = continued;
  }
  for (const group of active.values()) finish(group);
  return { gridColumns: block.gridColumns, rows };
}

function defaultDocumentTableColumnWidths(columns, widthDxa = 9360) {
  const count = Math.max(1, Number(columns) || 1);
  const base = Math.floor(widthDxa / count);
  return Array.from({ length: count }, (_value, index) => base + (index < widthDxa - base * count ? 1 : 0));
}

function documentTableFormatting(block, logicalColumns) {
  const invalid = (message) => {
    throw new OpenChestnutCodecError(`Document table ${block.id} ${message}`, [], { code: "invalid_document_table" });
  };
  const dxa = (value, name, { positive = false } = {}) => {
    if (!Number.isInteger(value) || value < (positive ? 1 : 0) || value > 1_000_000) {
      invalid(`${name} must be an integer from ${positive ? 1 : 0} through 1000000.`);
    }
    return value;
  };
  const widthDxa = dxa(block.widthDxa, "widthDxa", { positive: true });
  const indentDxa = dxa(block.indentDxa, "indentDxa");
  if (!Number.isInteger(logicalColumns) || logicalColumns < 1 || logicalColumns > 4_096) {
    invalid("requires between 1 and 4096 logical formatting columns.");
  }
  if (!Array.isArray(block.columnWidthsDxa) || block.columnWidthsDxa.length !== logicalColumns) {
    invalid(`columnWidthsDxa must contain one width for each of ${logicalColumns} logical grid columns.`);
  }
  const columnWidthsDxa = block.columnWidthsDxa.map((value, index) => dxa(value, `columnWidthsDxa[${index}]`, { positive: true }));
  if (columnWidthsDxa.reduce((sum, value) => sum + value, 0) !== widthDxa) {
    invalid("columnWidthsDxa must sum exactly to widthDxa.");
  }
  const margins = block.cellMarginsDxa;
  if (!margins || typeof margins !== "object") invalid("cellMarginsDxa must define top, bottom, start, and end margins.");
  const cellMarginsDxa = {
    top: dxa(margins.top, "cellMarginsDxa.top"),
    bottom: dxa(margins.bottom, "cellMarginsDxa.bottom"),
    start: dxa(margins.start, "cellMarginsDxa.start"),
    end: dxa(margins.end, "cellMarginsDxa.end"),
  };
  const borderColor = String(block.borderColor ?? "");
  const headerFill = String(block.headerFill ?? "");
  if (!/^[0-9A-F]{6}$/.test(borderColor)) invalid("borderColor must be a six-digit uppercase RGB value.");
  if (!/^[0-9A-F]{6}$/.test(headerFill)) invalid("headerFill must be a six-digit uppercase RGB value.");
  const borderSize = block.borderSize;
  if (!Number.isInteger(borderSize) || borderSize < 0 || borderSize > 96 || borderSize === 1) {
    invalid("borderSize must be zero or an integer from 2 through 96 eighths of a point.");
  }
  return { widthDxa, indentDxa, columnWidthsDxa, cellMarginsDxa, borderColor, borderSize, headerFill };
}

function documentTableFormattingConfig(table) {
  const logicalColumns = table.gridColumns || Math.max(1, ...table.rows.map((row) => row.cells.length));
  const formatting = table.formatting;
  if (formatting) {
    return {
      widthDxa: formatting.widthDxa,
      indentDxa: formatting.indentDxa,
      columnWidthsDxa: [...formatting.columnWidthsDxa],
      cellMarginsDxa: { ...formatting.cellMarginsDxa },
      borderColor: formatting.borderColor,
      borderSize: formatting.borderSize,
      headerFill: formatting.headerFill,
    };
  }
  return {
    widthDxa: 9360,
    indentDxa: 120,
    columnWidthsDxa: defaultDocumentTableColumnWidths(logicalColumns),
    cellMarginsDxa: { top: 80, bottom: 80, start: 120, end: 120 },
    borderColor: "D9D9D9",
    borderSize: 4,
    headerFill: "F2F4F7",
  };
}

function sameDocumentTableFormatting(block, table) {
  const expected = documentTableFormattingConfig(table);
  return block.widthDxa === expected.widthDxa && block.indentDxa === expected.indentDxa &&
    JSON.stringify(block.columnWidthsDxa) === JSON.stringify(expected.columnWidthsDxa) &&
    block.cellMarginsDxa?.top === expected.cellMarginsDxa.top &&
    block.cellMarginsDxa?.bottom === expected.cellMarginsDxa.bottom &&
    block.cellMarginsDxa?.start === expected.cellMarginsDxa.start &&
    block.cellMarginsDxa?.end === expected.cellMarginsDxa.end &&
    block.borderColor === expected.borderColor && block.borderSize === expected.borderSize &&
    block.headerFill === expected.headerFill;
}

function sameDocumentNumbering(block, paragraph) {
  const numbering = paragraph.numbering;
  if (!numbering || block.kind !== "listItem") return false;
  const numberFormat = numbering.numberFormat || "decimal";
  return block.text === paragraph.text &&
    block.listType === (numberFormat === "bullet" ? "bullet" : "number") &&
    block.numberFormat === numberFormat &&
    block.level === numbering.level &&
    block.start === (numbering.start || 1) &&
    block.levelText === (numbering.levelText || (numberFormat === "bullet" ? "•" : `%${numbering.level + 1}.`)) &&
    block.numberingId === numbering.numberingId &&
    block.abstractNumberingId === numbering.abstractNumberingId &&
    (block.numberingStyleId || "") === (numbering.numberingStyleId || "");
}

function sameDocumentNumberingIdentity(block, numbering) {
  return numbering && block.kind === "listItem" &&
    block.level === numbering.level &&
    block.numberingId === numbering.numberingId &&
    block.abstractNumberingId === numbering.abstractNumberingId &&
    (block.numberingStyleId || "") === (numbering.numberingStyleId || "");
}

function editedDocumentNumbering(block, source) {
  if (!sameDocumentNumberingIdentity(block, source)) {
    throw new OpenChestnutCodecError(`Document list item ${block.id} numbering identity, level, and style linkage are source-bound.`, [], { code: "unsupported_document_edit" });
  }
  const numberFormat = String(block.numberFormat ?? "");
  const levelText = String(block.levelText ?? "");
  const start = uint32(block.start, `Document list item ${block.id} start`);
  if (numberFormat.length > 128) {
    throw new OpenChestnutCodecError(`Document list item ${block.id} numberFormat exceeds 128 characters.`, [], { code: "invalid_document_numbering" });
  }
  if (levelText.length > 1_024) {
    throw new OpenChestnutCodecError(`Document list item ${block.id} levelText exceeds 1024 characters.`, [], { code: "invalid_document_numbering" });
  }
  const listType = numberFormat === "bullet" ? "bullet" : "number";
  if (block.listType !== listType) {
    throw new OpenChestnutCodecError(`Document list item ${block.id} listType must be ${listType} for numberFormat ${numberFormat || "(empty)"}.`, [], { code: "invalid_document_numbering" });
  }
  return { ...source, numberFormat, start, levelText };
}

function directDocumentNumberingPlan(document) {
  const groups = new Map();
  const usedNumberingIds = new Set();
  const usedAbstractIds = new Set();
  const result = new Map();
  const invalid = (message) => {
    throw new OpenChestnutCodecError(message, [], { code: "invalid_document_numbering" });
  };
  const integer = (value, name, { positive = false } = {}) => {
    const normalized = typeof value === "string" && /^\d+$/.test(value) ? Number(value) : value;
    if (!Number.isInteger(normalized) || normalized < (positive ? 1 : 0) || normalized > 0x7fff_ffff) {
      invalid(`${name} must be ${positive ? "a positive" : "a non-negative"} WordprocessingML signed integer.`);
    }
    return normalized;
  };
  const sameDefinition = (left, right) => left.numberFormat === right.numberFormat && left.start === right.start && left.levelText === right.levelText;

  for (const block of document.blocks.filter((item) => item.kind === "listItem")) {
    if (block.pictureBullet) {
      throw new OpenChestnutCodecError(`The DOCX WebAssembly vertical slice cannot directly author picture bullet ${block.id}.`, [], { code: "unsupported_document_features" });
    }
    if (block.numberingStyleId) {
      throw new OpenChestnutCodecError(`The DOCX WebAssembly vertical slice cannot directly author style-linked numbering for list item ${block.id}.`, [], { code: "unsupported_document_features" });
    }
    const level = integer(block.level, `Document list item ${block.id} level`);
    if (level > 8) invalid(`Document list item ${block.id} level must be between 0 and 8.`);
    const start = integer(block.start, `Document list item ${block.id} start`, { positive: true });
    const numberFormat = String(block.numberFormat || "");
    const levelText = String(block.levelText || "");
    if (!numberFormat || numberFormat.length > 128) invalid(`Document list item ${block.id} numberFormat must contain 1 through 128 characters.`);
    if (!levelText || levelText.length > 1_024) invalid(`Document list item ${block.id} levelText must contain 1 through 1024 characters.`);
    const expectedListType = numberFormat === "bullet" ? "bullet" : "number";
    if (block.listType !== expectedListType) invalid(`Document list item ${block.id} listType must be ${expectedListType} for numberFormat ${numberFormat}.`);

    const explicitNumberingId = block.numberingId == null ? undefined : integer(block.numberingId, `Document list item ${block.id} numberingId`, { positive: true });
    const explicitAbstractId = block.abstractNumberingId == null ? undefined : integer(block.abstractNumberingId, `Document list item ${block.id} abstractNumberingId`);
    if (explicitNumberingId != null) usedNumberingIds.add(explicitNumberingId);
    if (explicitAbstractId != null) usedAbstractIds.add(explicitAbstractId);
    const key = explicitNumberingId == null ? `default:${block.listType}` : `native:${explicitNumberingId}`;
    if (!groups.has(key)) groups.set(key, { blocks: [], definitions: new Map(), explicitNumberingId, abstractIds: new Set() });
    const group = groups.get(key);
    if (group.explicitNumberingId !== explicitNumberingId) invalid(`Document numbering group ${key} has conflicting numbering IDs.`);
    if (explicitAbstractId != null) group.abstractIds.add(explicitAbstractId);
    const definition = { numberFormat, start, levelText };
    const existing = group.definitions.get(level);
    if (existing && !sameDefinition(existing, definition)) invalid(`Document numbering ${explicitNumberingId ?? key} level ${level} has conflicting definitions.`);
    group.definitions.set(level, definition);
    group.blocks.push({ block, level, definition });
  }

  const allocate = (used, start = 1) => {
    let candidate = start;
    while (used.has(candidate)) candidate += 1;
    if (candidate > 0x7fff_ffff) invalid("Document numbering ID space is exhausted.");
    used.add(candidate);
    return candidate;
  };
  const sharedDefinitions = new Map();
  for (const [key, group] of groups) {
    if (group.abstractIds.size > 1) invalid(`Document numbering ${group.explicitNumberingId ?? key} references conflicting abstract numbering IDs.`);
    const numberingId = group.explicitNumberingId ?? allocate(usedNumberingIds);
    const abstractNumberingId = group.abstractIds.size ? [...group.abstractIds][0] : allocate(usedAbstractIds);
    for (const [level, definition] of group.definitions) {
      const definitionKey = `${abstractNumberingId}:${level}`;
      const existing = sharedDefinitions.get(definitionKey);
      if (existing && !sameDefinition(existing, definition)) invalid(`Document abstract numbering ${abstractNumberingId} level ${level} has conflicting definitions.`);
      sharedDefinitions.set(definitionKey, definition);
    }
    for (const { block, level, definition } of group.blocks) {
      result.set(block, { numberingId, abstractNumberingId, level, ...definition });
    }
  }
  return result;
}

function sameDocumentHyperlink(block, source) {
  if (block.kind !== "hyperlink" || block.text !== source.text) return false;
  if (block.styleId !== (source.styleId || "Normal")) return false;
  if ((block.relationshipId || "") !== (source.relationshipId || "")) return false;
  if ((block.tooltip ?? undefined) !== source.tooltip) return false;
  if (block.history !== (source.history ?? true)) return false;
  if (source.target.case === "externalUri") return !block.anchor && block.url === source.target.value;
  if (source.target.case === "internalAnchor") return block.anchor === source.target.value && !block.url;
  return false;
}

function documentHyperlink(block, original) {
  const source = original?.content.case === "hyperlink" ? original.content.value : undefined;
  if (source && (block.relationshipId || "") !== (source.relationshipId || "")) {
    throw new OpenChestnutCodecError(`Document hyperlink ${block.id} relationshipId is a source locator and cannot be edited directly.`, [], { code: "unsupported_document_edit" });
  }
  const text = String(block.text ?? "");
  if (text.length > 1_000_000) throw new OpenChestnutCodecError(`Document hyperlink ${block.id} text exceeds 1,000,000 characters.`, [], { code: "invalid_document_hyperlink" });
  const anchor = String(block.anchor || "").trim();
  const url = String(block.url || "");
  let target;
  if (anchor) {
    if (anchor.length > 255 || [...anchor].some((character) => /[\u0000-\u001f\u007f]/.test(character))) {
      throw new OpenChestnutCodecError(`Document hyperlink ${block.id} anchor must contain 1 through 255 characters without controls.`, [], { code: "invalid_document_hyperlink" });
    }
    target = { case: "internalAnchor", value: anchor };
  } else {
    let parsed;
    try { parsed = new URL(url); } catch { parsed = undefined; }
    if (!parsed || !new Set(["http:", "https:"]).has(parsed.protocol) || url.length > 4_096 || /[\u0000-\u001f\u007f]/.test(url)) {
      throw new OpenChestnutCodecError(`Document hyperlink ${block.id} URI must be an absolute http(s) URI of at most 4096 characters without controls.`, [], { code: "invalid_document_hyperlink" });
    }
    target = { case: "externalUri", value: url };
  }
  if (block.tooltip != null && String(block.tooltip).length > 260) {
    throw new OpenChestnutCodecError(`Document hyperlink ${block.id} tooltip exceeds 260 characters.`, [], { code: "invalid_document_hyperlink" });
  }
  const originalHistory = source?.history;
  const history = source && block.history === (originalHistory ?? true) ? originalHistory : block.history;
  return {
    text,
    target,
    relationshipId: source?.relationshipId || "",
    tooltip: block.tooltip == null ? undefined : String(block.tooltip),
    history,
  };
}

function documentField(block, original) {
  if (original?.source?.editable === false) {
    throw new OpenChestnutCodecError(`Document field ${block.id} is source-preserved but its instruction or result topology is not editable.`, [], { code: "unsupported_document_edit" });
  }
  const instruction = String(block.instruction ?? "");
  const display = String(block.display ?? "");
  if (!instruction.trim() || instruction.length > 8_192 || /[\u0000-\u001f\u007f]/.test(instruction)) {
    throw new OpenChestnutCodecError(`Document field ${block.id} instruction must contain 1 through 8192 characters without controls.`, [], { code: "invalid_document_field" });
  }
  const command = /^[A-Za-z]+/.exec(instruction.trimStart())?.[0]?.toUpperCase();
  const complex = Boolean(block.complex);
  if (!command || (complex ? command !== "TOC" : !DOCUMENT_FIELD_COMMANDS.has(command))) {
    throw new OpenChestnutCodecError(`Document field ${block.id} command ${command || "(missing)"} is outside the bounded editable field catalog.`, [], { code: "invalid_document_field" });
  }
  if (complex && !/^TOC \\o "[1-9]-[1-9]"(?: \\h)?(?: \\z)?(?: \\u)?$/.test(instruction)) {
    throw new OpenChestnutCodecError(`Document field ${block.id} complex TOC instruction is outside the canonical bounded profile.`, [], { code: "invalid_document_field" });
  }
  if (display.length > 1_000_000) throw new OpenChestnutCodecError(`Document field ${block.id} display text exceeds 1,000,000 characters.`, [], { code: "invalid_document_field" });
  return { instruction, display, complex };
}

function documentCommentSnapshot(comment) {
  return {
    id: comment.id,
    targetId: comment.targetId,
    author: comment.author,
    initials: comment.initials,
    date: comment.date,
    text: comment.text,
    resolved: comment.resolved,
    parentId: comment.parentId,
    paraId: comment.paraId,
    durableId: comment.durableId,
    dateUtc: comment.dateUtc,
    person: comment.person,
    intelligentPlaceholder: comment.intelligentPlaceholder,
  };
}

const DOCUMENT_COMMENT_HEX_ID = /^[0-9A-F]{8}$/;

function validateDocumentCommentThreads(document) {
  const byId = new Map();
  for (const comment of document.comments) {
    const id = String(comment.id || "");
    if (!id || byId.has(id)) {
      throw new OpenChestnutCodecError("Document comments require unique, non-empty IDs.", [], { code: "invalid_document_comment" });
    }
    byId.set(id, comment);
  }
  for (const comment of document.comments) {
    if (comment.parentId) {
      const parent = byId.get(String(comment.parentId));
      if (!parent) {
        throw new OpenChestnutCodecError(`Document comment ${comment.id} references missing parent ${comment.parentId}.`, [], { code: "invalid_document_comment_thread" });
      }
      if (parent.parentId) {
        throw new OpenChestnutCodecError(`Document comment ${comment.id} is a nested reply; OpenChestnut supports roots plus direct replies only.`, [], { code: "unsupported_document_comment_thread" });
      }
      if (parent.targetId !== comment.targetId) {
        throw new OpenChestnutCodecError(`Document comment ${comment.id} and root ${parent.id} must target the same block.`, [], { code: "invalid_document_comment_thread" });
      }
      if (comment.intelligentPlaceholder) {
        throw new OpenChestnutCodecError(`Document reply ${comment.id} cannot be an intelligent placeholder.`, [], { code: "invalid_document_comment_thread" });
      }
    }
    for (const [name, value] of [["paraId", comment.paraId], ["durableId", comment.durableId]]) {
      if (value != null && value !== "" && !DOCUMENT_COMMENT_HEX_ID.test(String(value).toUpperCase())) {
        throw new OpenChestnutCodecError(`Document comment ${comment.id} ${name} must contain exactly eight hexadecimal digits.`, [], { code: "invalid_document_comment" });
      }
    }
    if (comment.durableId) {
      const durableNumber = Number.parseInt(comment.durableId, 16);
      if (durableNumber <= 0 || durableNumber >= 0x7FFFFFFF) {
        throw new OpenChestnutCodecError(`Document comment ${comment.id} durableId must be between 00000001 and 7FFFFFFE.`, [], { code: "invalid_document_comment" });
      }
    }
    if (comment.dateUtc != null) {
      const dateUtc = String(comment.dateUtc);
      if (!dateUtc || dateUtc.length > 64 || Number.isNaN(Date.parse(dateUtc))) {
        throw new OpenChestnutCodecError(`Document comment ${comment.id} dateUtc must be an ISO 8601 date-time of at most 64 characters.`, [], { code: "invalid_document_comment" });
      }
    }
    if (comment.person) {
      const providerId = String(comment.person.providerId ?? "");
      const userId = String(comment.person.userId ?? "");
      if (!providerId || !userId || providerId.length > 100 || userId.length > 300) {
        throw new OpenChestnutCodecError(`Document comment ${comment.id} person requires providerId of 1 through 100 characters and userId of 1 through 300 characters.`, [], { code: "invalid_document_comment" });
      }
    }
  }
  const commentsByAuthor = new Map();
  for (const comment of document.comments) {
    const author = String(comment.author || "");
    if (!commentsByAuthor.has(author)) commentsByAuthor.set(author, []);
    commentsByAuthor.get(author).push(comment);
  }
  for (const comments of commentsByAuthor.values()) {
    const profiles = new Set(comments.map((comment) => comment.person
      ? `${comment.person.providerId}\u0000${comment.person.userId}`
      : ""));
    if (profiles.size > 1) {
      throw new OpenChestnutCodecError(`Document comment author ${comments[0].author} has inconsistent people metadata.`, [], { code: "invalid_document_comment" });
    }
  }
}

function documentBookmarkSnapshot(bookmark) {
  return {
    id: bookmark.id,
    name: bookmark.name,
    targetId: bookmark.targetId,
    endTargetId: bookmark.endTargetId,
    nativeId: bookmark.nativeId,
  };
}

function wireDocumentBibliographySource(source) {
  const tag = String(source?.tag || "");
  if (!DOCUMENT_CITATION_TAG.test(tag)) {
    throw new OpenChestnutCodecError(`Document bibliography source ${source?.id || "(unknown)"} tag must contain 1 through 255 ASCII letters, digits, periods, underscores, colons, or hyphens.`, [], { code: "invalid_document_bibliography" });
  }
  const authors = (source.authors || []).map((author) => ({
    first: String(author?.first || ""),
    middle: String(author?.middle || ""),
    last: String(author?.last || ""),
  }));
  const fields = Object.fromEntries(DOCUMENT_BIBLIOGRAPHY_FIELD_KEYS.flatMap((key) =>
    source[key] === undefined || source[key] === null || source[key] === "" ? [] : [[key, String(source[key])]]));
  return {
    id: String(source.id || `bibliography/${tag}`),
    tag,
    sourceType: String(source.sourceType || "Misc"),
    authors,
    corporateAuthor: String(source.corporateAuthor || ""),
    fields,
  };
}

function publicDocumentBibliographySource(source) {
  return {
    id: source.id,
    tag: source.tag,
    sourceType: source.sourceType,
    authors: (source.authors || []).map((author) => ({ first: author.first, middle: author.middle, last: author.last })),
    corporateAuthor: source.corporateAuthor || undefined,
    ...(source.fields || {}),
  };
}

function wireDocumentBibliography(document, original) {
  const settings = {
    selectedStyle: String(document.bibliography?.selectedStyle || ""),
    styleName: String(document.bibliography?.styleName || ""),
    uri: String(document.bibliography?.uri || ""),
  };
  const sources = document.bibliographySources.map(wireDocumentBibliographySource);
  if (!sources.length && !Object.values(settings).some(Boolean)) return undefined;
  if (original) {
    if (sources.length !== original.sources.length) {
      throw new OpenChestnutCodecError(`Source-preserving DOCX export requires the original ${original.sources.length}-source bibliography topology; the document contains ${sources.length} sources.`, [], { code: "document_bibliography_topology_changed" });
    }
    for (let index = 0; index < sources.length; index += 1) {
      if (sources[index].id !== original.sources[index].id || sources[index].tag !== original.sources[index].tag) {
        throw new OpenChestnutCodecError(`Imported document bibliography source ${index} ID, tag, and order are source-bound.`, [], { code: "unsupported_document_bibliography_edit" });
      }
    }
  }
  return { ...settings, sources, source: original?.source };
}

function wireDocumentCitation(block, original) {
  const tag = String(block.metadata?.tag ?? block.metadata?.bibliographyTag ?? "");
  if (!DOCUMENT_CITATION_TAG.test(tag)) {
    throw new OpenChestnutCodecError(`Document citation ${block.id} tag must contain 1 through 255 ASCII letters, digits, periods, underscores, colons, or hyphens.`, [], { code: "invalid_document_citation" });
  }
  if (original && tag !== original.tag) {
    throw new OpenChestnutCodecError(`Imported document citation ${block.id} source tag is source-bound.`, [], { code: "unsupported_document_edit" });
  }
  const display = String(block.text ?? "");
  if (display.length > 1_000_000) {
    throw new OpenChestnutCodecError(`Document citation ${block.id} display text exceeds 1,000,000 characters.`, [], { code: "invalid_document_citation" });
  }
  return { tag, display };
}

function documentNoteSnapshot(note) {
  return {
    id: note.id,
    kind: note.kind,
    targetId: note.targetId,
    text: note.text,
    nativeId: note.nativeId,
  };
}

function wireDocumentNoteKind(value) {
  if (value === "footnote") return DocumentNoteKind.FOOTNOTE;
  if (value === "endnote") return DocumentNoteKind.ENDNOTE;
  return DocumentNoteKind.UNSPECIFIED;
}

function publicDocumentNoteKind(value) {
  if (value === DocumentNoteKind.FOOTNOTE) return "footnote";
  if (value === DocumentNoteKind.ENDNOTE) return "endnote";
  throw new OpenChestnutCodecError(`Document note kind ${value} is invalid.`, [], { code: "invalid_document_note" });
}

function documentNote(note, slot, document) {
  const kind = String(note.kind || "");
  const targetId = String(note.targetId || "");
  const text = String(note.text ?? "");
  const nativeId = note.nativeId === undefined ? "" : String(note.nativeId);
  if (slot) {
    const original = slot.publicSnapshot;
    if (note.id !== original.id || kind !== original.kind || targetId !== original.targetId || note.nativeId !== original.nativeId) {
      throw new OpenChestnutCodecError(`Imported document ${kind || "note"} ${note.id} identity, kind, target, and native ID are source-bound.`, [], { code: "unsupported_document_note_edit" });
    }
    if (text === original.text) return slot.wire;
    if (slot.wire.source?.editable !== true) {
      throw new OpenChestnutCodecError(`Imported document ${kind} ${note.id} body topology is preserved but not editable.`, [], { code: "unsupported_document_note_edit" });
    }
    validateDocumentNoteText(note, text);
    return { ...slot.wire, text };
  }
  if (!new Set(["footnote", "endnote"]).has(kind)) {
    throw new OpenChestnutCodecError(`Document note ${note.id} kind must be footnote or endnote.`, [], { code: "invalid_document_note" });
  }
  const target = document.blocks.find((block) => block.id === targetId);
  if (!target || !new Set(["paragraph", "listItem"]).has(target.kind)) {
    throw new OpenChestnutCodecError(`Document ${kind} ${note.id} target must be a paragraph or list item.`, [], { code: "invalid_document_note" });
  }
  validateDocumentNoteText(note, text);
  if (nativeId && (!/^\d+$/.test(nativeId) || Number(nativeId) < 1 || Number(nativeId) > 2_147_483_647)) {
    throw new OpenChestnutCodecError(`Document ${kind} ${note.id} nativeId must be a positive 32-bit integer when present.`, [], { code: "invalid_document_note" });
  }
  return {
    id: String(note.id || ""),
    kind: wireDocumentNoteKind(kind),
    targetBlockId: targetId,
    text,
    nativeId,
  };
}

function validateDocumentNoteText(note, text) {
  if (!text.length || text.length > 1_000_000 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(text)) {
    throw new OpenChestnutCodecError(`Document ${note.kind || "note"} ${note.id} text must contain 1 through 1,000,000 XML-safe characters.`, [], { code: "invalid_document_note" });
  }
}

function documentBookmark(bookmark, slot, document) {
  if (slot) {
    if (JSON.stringify(documentBookmarkSnapshot(bookmark)) !== JSON.stringify(slot.publicSnapshot)) {
      throw new OpenChestnutCodecError(`Imported document bookmark ${bookmark.id} identity, name, and target are source-bound in protocol 2.`, [], { code: "unsupported_document_bookmark_edit" });
    }
    return slot.wire;
  }
  const name = String(bookmark.name || "");
  if (!/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(name)) {
    throw new OpenChestnutCodecError(`Document bookmark ${bookmark.id} name must start with an ASCII letter and contain only letters, digits, or underscores (maximum 40 characters).`, [], { code: "invalid_document_bookmark" });
  }
  if (!bookmark.targetId || bookmark.targetId !== bookmark.endTargetId) {
    throw new OpenChestnutCodecError(`Document bookmark ${bookmark.id} must wrap exactly one block in protocol 2.`, [], { code: "invalid_document_bookmark" });
  }
  const target = document.blocks.find((block) => block.id === bookmark.targetId);
  if (!target || !new Set(["paragraph", "hyperlink", "field", "citation", "change", "image"]).has(target.kind)) {
    throw new OpenChestnutCodecError(`Document bookmark ${bookmark.id} target must be a paragraph, hyperlink, field, citation, tracked change, or image block.`, [], { code: "invalid_document_bookmark" });
  }
  let nativeId = "";
  if (bookmark.nativeId !== undefined) {
    const value = Number(bookmark.nativeId);
    if (!Number.isInteger(value) || value < 0 || value > 4_294_967_295) {
      throw new OpenChestnutCodecError(`Document bookmark ${bookmark.id} nativeId must be an unsigned 32-bit integer when present.`, [], { code: "invalid_document_bookmark" });
    }
    nativeId = String(value);
  }
  return {
    id: String(bookmark.id || ""),
    name,
    targetBlockId: bookmark.targetId,
    endTargetBlockId: bookmark.endTargetId,
    nativeId,
  };
}

function documentComment(comment, slot) {
  if (slot && (comment.id !== slot.wire.id || comment.targetId !== slot.wire.targetBlockId)) {
    throw new OpenChestnutCodecError(`Document comment ${comment.id} identity and target are source-bound.`, [], { code: "unsupported_document_comment_edit" });
  }
  if (slot) {
    const immutable = {
      parentId: slot.wire.parentCommentId || undefined,
      paraId: slot.wire.paragraphId || undefined,
      durableId: slot.wire.durableId || undefined,
      dateUtc: slot.wire.dateUtc,
      person: slot.wire.person ? { providerId: slot.wire.person.providerId, userId: slot.wire.person.userId } : undefined,
      intelligentPlaceholder: Boolean(slot.wire.intelligentPlaceholder),
    };
    const requested = {
      parentId: comment.parentId || undefined,
      paraId: comment.paraId || undefined,
      durableId: comment.durableId || undefined,
      dateUtc: comment.dateUtc,
      person: comment.person ? { providerId: String(comment.person.providerId ?? ""), userId: String(comment.person.userId ?? "") } : undefined,
      intelligentPlaceholder: Boolean(comment.intelligentPlaceholder),
    };
    if (JSON.stringify(requested) !== JSON.stringify(immutable)) {
      throw new OpenChestnutCodecError(`Document comment ${comment.id} parent, paragraph/durable identity, UTC/person metadata, and intelligent-placeholder state are source-bound.`, [], { code: "unsupported_document_comment_edit" });
    }
  }
  if (slot && JSON.stringify(documentCommentSnapshot(comment)) === JSON.stringify(slot.publicSnapshot)) return slot.wire;
  const author = String(comment.author ?? "");
  const initials = slot && comment.initials === slot.publicSnapshot.initials
    ? slot.wire.initials
    : comment.initials == null ? undefined : String(comment.initials);
  const text = String(comment.text ?? "");
  if (!author || author.length > 255) throw new OpenChestnutCodecError(`Document comment ${comment.id} author must contain 1 through 255 characters.`, [], { code: "invalid_document_comment" });
  if (initials !== undefined && (!initials || initials.length > 9)) throw new OpenChestnutCodecError(`Document comment ${comment.id} initials must contain 1 through 9 characters when present.`, [], { code: "invalid_document_comment" });
  if (text.length > 1_000_000) throw new OpenChestnutCodecError(`Document comment ${comment.id} text exceeds 1,000,000 characters.`, [], { code: "invalid_document_comment" });
  let createdAt;
  if (comment.date != null) {
    createdAt = String(comment.date);
    if (createdAt.length > 64 || Number.isNaN(Date.parse(createdAt))) throw new OpenChestnutCodecError(`Document comment ${comment.id} date must be an ISO 8601 date-time of at most 64 characters.`, [], { code: "invalid_document_comment" });
  }
  const modern = Boolean(comment.parentId || comment.paraId || comment.durableId || comment.dateUtc || comment.person || comment.intelligentPlaceholder || comment._resolvedSpecified || slot?.wire.resolved !== undefined);
  return {
    id: slot?.wire.id || comment.id,
    targetBlockId: comment.targetId,
    author,
    text,
    initials,
    createdAt,
    source: slot?.wire.source,
    parentCommentId: comment.parentId || "",
    resolved: modern ? Boolean(comment.resolved) : undefined,
    paragraphId: comment.paraId ? String(comment.paraId).toUpperCase() : "",
    durableId: comment.durableId ? String(comment.durableId).toUpperCase() : "",
    dateUtc: comment.dateUtc == null ? undefined : String(comment.dateUtc),
    person: comment.person ? { providerId: String(comment.person.providerId), userId: String(comment.person.userId) } : undefined,
    intelligentPlaceholder: modern && comment.intelligentPlaceholder ? true : undefined,
  };
}

function documentStyleType(value) {
  if (value === "character") return DocumentStyleType.CHARACTER;
  if (value === "table") return DocumentStyleType.TABLE;
  return DocumentStyleType.PARAGRAPH;
}

function publicDocumentStyleType(value) {
  if (value === DocumentStyleType.CHARACTER) return "character";
  if (value === DocumentStyleType.TABLE) return "table";
  return "paragraph";
}

function wireDocumentStyle(style) {
  const runSource = Object.fromEntries([...DOCUMENT_RUN_STYLE_KEYS].filter((key) => key !== "runStyleId" && Object.hasOwn(style, key)).map((key) => [key, style[key]]));
  return {
    id: String(style.id || ""),
    name: String(style.name || style.id || ""),
    type: documentStyleType(style.type),
    basedOn: String(style.basedOn || style.parent || style.extends || ""),
    runFormat: documentRunFormatting(runSource, `Document style ${style.id || "(unnamed)"}`),
    paragraphFormat: documentParagraphFormatting({ id: style.id || "(unnamed)", paragraphFormat: style.paragraphFormat || style }),
  };
}

function publicDocumentStyle(style) {
  return {
    id: style.id,
    name: style.name || style.id,
    type: publicDocumentStyleType(style.type),
    ...(style.basedOn ? { basedOn: style.basedOn } : {}),
    ...publicDocumentRunFormatting(style.runFormat),
    ...(publicDocumentParagraphFormatting(style.paragraphFormat) || {}),
  };
}

function headerFooterReference(value) {
  if (value === "first") return DocumentHeaderFooterReference.FIRST;
  if (value === "even") return DocumentHeaderFooterReference.EVEN;
  return DocumentHeaderFooterReference.DEFAULT;
}

function publicHeaderFooterReference(value) {
  if (value === DocumentHeaderFooterReference.FIRST) return "first";
  if (value === DocumentHeaderFooterReference.EVEN) return "even";
  return "default";
}

function wireHeaderFooter(block) {
  const instruction = String(block.fieldInstruction || block.field || "");
  if (instruction && !DOCUMENT_FIELD_COMMANDS.has(instruction.trim().split(/\s+/)[0].toUpperCase())) throw new OpenChestnutCodecError(`Document ${block.kind} ${block.id} uses unsupported field ${instruction}.`, [], { code: "invalid_document_field" });
  return {
    id: String(block.id || ""),
    name: String(block.name || block.kind || ""),
    styleId: String(block.styleId || "Normal"),
    text: String(block.text || ""),
    reference: headerFooterReference(block.referenceType),
    sectionIndex: block.sectionIndex == null ? undefined : uint32(block.sectionIndex, `Document ${block.kind} ${block.id} sectionIndex`),
    relationshipId: String(block.relationshipId || ""),
    partPath: String(block.partPath || ""),
    variantActive: block.variantActive == null ? undefined : Boolean(block.variantActive),
    fieldInstruction: instruction,
  };
}

function publicHeaderFooter(block) {
  return {
    id: block.id || undefined,
    name: block.name || undefined,
    styleId: block.styleId || "Normal",
    text: block.text || "",
    referenceType: publicHeaderFooterReference(block.reference),
    sectionIndex: block.sectionIndex,
    relationshipId: block.relationshipId || undefined,
    partPath: block.partPath || undefined,
    variantActive: block.variantActive,
    fieldInstruction: block.fieldInstruction || undefined,
  };
}

function documentSectionBreak(value) {
  if (value === "continuous") return DocumentSectionBreak.CONTINUOUS;
  if (value === "evenPage") return DocumentSectionBreak.EVEN_PAGE;
  if (value === "oddPage") return DocumentSectionBreak.ODD_PAGE;
  return DocumentSectionBreak.NEXT_PAGE;
}

function publicDocumentSectionBreak(value) {
  if (value === DocumentSectionBreak.CONTINUOUS) return "continuous";
  if (value === DocumentSectionBreak.EVEN_PAGE) return "evenPage";
  if (value === DocumentSectionBreak.ODD_PAGE) return "oddPage";
  return "nextPage";
}

function documentChangeType(value) {
  if (value === "insert") return DocumentChangeType.INSERT;
  if (value === "delete") return DocumentChangeType.DELETE;
  throw new OpenChestnutCodecError(`Document tracked-change type ${value || "(empty)"} must be insert or delete.`, [], { code: "invalid_document_change" });
}

function publicDocumentChangeType(value) {
  if (value === DocumentChangeType.INSERT) return "insert";
  if (value === DocumentChangeType.DELETE) return "delete";
  throw new OpenChestnutCodecError("Document tracked-change wire type must be insert or delete.", [], { code: "invalid_document_change" });
}

function wireDocumentChange(block) {
  const text = String(block.text ?? "");
  const author = String(block.author ?? "");
  const date = block.date == null || block.date === "" ? undefined : String(block.date);
  if (text.length > 1_000_000) throw new OpenChestnutCodecError(`Document tracked change ${block.id} text exceeds 1,000,000 characters.`, [], { code: "invalid_document_change" });
  if (!author.trim() || author.length > 255 || /[\u0000-\u001f\u007f]/.test(author)) throw new OpenChestnutCodecError(`Document tracked change ${block.id} requires an author of at most 255 characters without controls.`, [], { code: "invalid_document_change" });
  if (date !== undefined && Number.isNaN(Date.parse(date))) throw new OpenChestnutCodecError(`Document tracked change ${block.id} date must be an ISO 8601 timestamp.`, [], { code: "invalid_document_change" });
  return { type: documentChangeType(block.changeType), text, author, date };
}

function documentImage(block, assets) {
  if (!block.dataUrl) throw new OpenChestnutCodecError(`Document image ${block.id} requires embedded PNG or JPEG data.`, [], { code: "unsupported_document_image" });
  const match = /^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+/=\s]+)$/i.exec(String(block.dataUrl));
  if (!match) throw new OpenChestnutCodecError(`Document image ${block.id} must use a base64 PNG/JPEG data URL.`, [], { code: "unsupported_document_image" });
  const bytes = new Uint8Array(Buffer.from(match[2].replace(/\s/g, ""), "base64"));
  if (!bytes.length) throw new OpenChestnutCodecError(`Document image ${block.id} contains no image bytes.`, [], { code: "invalid_document_image" });
  const contentType = match[1].toLowerCase();
  const sha256 = createHash("sha256").update(bytes).digest("hex");
  const assetId = `asset/document/image/${sha256}`;
  assets.set(assetId, { id: assetId, fileName: `${sha256}.${contentType === "image/png" ? "png" : "jpg"}`, contentType, data: bytes, sha256 });
  const widthEmu = Math.round(Number(block.widthPx) * 9_525);
  const heightEmu = Math.round(Number(block.heightPx) * 9_525);
  if (!Number.isSafeInteger(widthEmu) || !Number.isSafeInteger(heightEmu) || widthEmu <= 0 || heightEmu <= 0) throw new OpenChestnutCodecError(`Document image ${block.id} dimensions must be positive bounded pixels.`, [], { code: "invalid_document_image" });
  return { assetId, altText: String(block.alt || block.name || "image"), widthEmu, heightEmu };
}

function wireDocumentSection(block) {
  const page = block.pageSize || {};
  const margins = block.margins || {};
  return {
    breakType: documentSectionBreak(block.breakType),
    pageWidthTwips: uint32(Math.round(Number(page.widthTwips)), `Document section ${block.id} page width`),
    pageHeightTwips: uint32(Math.round(Number(page.heightTwips)), `Document section ${block.id} page height`),
    landscape: block.orientation === "landscape",
    marginTopTwips: uint32(Math.round(Number(margins.top)), `Document section ${block.id} top margin`),
    marginRightTwips: uint32(Math.round(Number(margins.right)), `Document section ${block.id} right margin`),
    marginBottomTwips: uint32(Math.round(Number(margins.bottom)), `Document section ${block.id} bottom margin`),
    marginLeftTwips: uint32(Math.round(Number(margins.left)), `Document section ${block.id} left margin`),
  };
}

function unchangedSourceBlock(block, original) {
  switch (original.content.case) {
    case "paragraph": {
      if (original.content.value.numbering) {
        return block.styleId === (original.styleId || "Normal") && sameDocumentNumbering(block, original.content.value);
      }
      if (block.kind !== "paragraph" || block.text !== original.content.value.text || block.styleId !== (original.styleId || "Normal")) return false;
      if (original.source?.editable !== false) return false;
      return block.runs.every((run) => Object.keys(run.style || {}).length === 0);
    }
    case "table": {
      if (block.kind !== "table" || block.textPatches?.length || !sameTableValues(block, original) ||
          !sameDocumentTableGeometry(block, original.content.value) ||
          !sameDocumentTableFormatting(block, original.content.value)) return false;
      return block.styleId === original.styleId || (!original.styleId && block.styleId === "TableGrid");
    }
    case "hyperlink":
      return sameDocumentHyperlink(block, original.content.value);
    case "field":
      return block.kind === "field" && block.styleId === (original.styleId || "Normal") && block.instruction === original.content.value.instruction && block.display === original.content.value.display && Boolean(block.complex) === Boolean(original.content.value.complex);
    case "citation":
      return block.kind === "citation" && block.styleId === (original.styleId || "Normal") &&
        String(block.metadata?.tag || "") === original.content.value.tag && block.text === original.content.value.display;
    case "change": {
      if (block.kind !== "change" || block.styleId !== (original.styleId || "Normal")) return false;
      const value = original.content.value;
      return block.changeType === publicDocumentChangeType(value.type) && block.text === value.text && block.author === value.author && (block.date || undefined) === value.date;
    }
    case "section": {
      if (block.kind !== "section") return false;
      const value = wireDocumentSection(block);
      return JSON.stringify(value) === JSON.stringify(original.content.value);
    }
    case "opaque":
      return block.kind === "paragraph" && block.text === original.content.value.text && block.runs.every((run) => Object.keys(run.style || {}).length === 0);
    default:
      return false;
  }
}

function documentBlockSnapshot(block) {
  return JSON.stringify({
    proto: typeof block?.toProto === "function" ? block.toProto() : {
      id: block?.id,
      name: block?.name,
      kind: block?.kind,
      styleId: block?.styleId,
      text: block?.text,
    },
    runs: Array.isArray(block?.runs)
      ? block.runs.map((run) => ({ text: String(run?.text ?? ""), style: { ...(run?.style || {}) }, contentControl: run?.contentControl ? { ...run.contentControl } : undefined, inlineField: run?.inlineField ? { ...run.inlineField } : undefined }))
      : undefined,
  });
}

function patchedSourceParagraphBlock(block, original) {
  if (block.kind !== "paragraph") return undefined;
  const patches = Array.isArray(block.textPatches) ? block.textPatches : [];
  if (!patches.length) return undefined;
  if (original?.content.case !== "paragraph" || original.source?.textPatchable !== true || original.source?.editable !== false) {
    throw new OpenChestnutCodecError(`Document paragraph ${block.id} text patches require a non-editable imported paragraph that advertises textPatchable.`, [], { code: "unsupported_document_edit" });
  }
  if (patches.length > 10_000) throw new OpenChestnutCodecError(`Document paragraph ${block.id} exceeds 10,000 source text patches.`, [], { code: "invalid_document_text_patch" });
  let expected = String(original.content.value.text ?? "");
  const sourceTextSha256 = createHash("sha256").update(expected, "utf8").digest("hex");
  const wirePatches = patches.map((patch) => {
    const search = String(patch.search ?? "");
    const replacement = String(patch.replacement ?? "");
    if (!search || search.length > 1_000_000 || replacement.length > 1_000_000 || !isXmlSafeText(search) || !isXmlSafeText(replacement)) {
      throw new OpenChestnutCodecError(`Document paragraph ${block.id} text patch requires bounded XML-safe strings.`, [], { code: "invalid_document_text_patch" });
    }
    const first = expected.indexOf(search);
    if (first < 0 || expected.indexOf(search, first + 1) >= 0) {
      throw new OpenChestnutCodecError(`Document paragraph ${block.id} text patch requires exactly one visible match.`, [], { code: "unsupported_document_edit" });
    }
    expected = expected.replace(search, replacement);
    return { search, replacement, sourceTextSha256 };
  });
  const baselineFormat = publicDocumentParagraphFormatting(original.content.value.formatting) || {};
  const plainSyntheticRuns = block.runs.length === (expected ? 1 : 0) && block.runs.every((run) =>
    !run.contentControl && !run.inlineField && Object.keys(run.style || {}).length === 0 && run.text === expected);
  if (block.kind !== "paragraph" || block.id !== original.id || block.name !== (original.name || "") ||
      block.styleId !== (original.styleId || "Normal") || block.text !== expected || !plainSyntheticRuns ||
      JSON.stringify(block.paragraphFormat || {}) !== JSON.stringify(baselineFormat)) {
    throw new OpenChestnutCodecError(`Document paragraph ${block.id} cannot combine a native text patch with other semantic or formatting edits.`, [], { code: "unsupported_document_edit" });
  }
  const { $typeName: _typeName, ...sourceBlock } = original;
  return { ...sourceBlock, textPatches: wirePatches };
}

function documentBlock(block, original, directNumbering, assets, contentControlNativeIds) {
  const patchedParagraph = patchedSourceParagraphBlock(block, original);
  if (patchedParagraph) return patchedParagraph;
  if (original && unchangedSourceBlock(block, original)) return original;
  const common = {
    id: original?.id || block.id,
    name: block.name || original?.name || "",
    styleId: block.styleId || original?.styleId || "",
    source: original?.source,
  };
  if (block.kind === "paragraph") {
    assertDocumentContentControlTopology(block, original);
    assertDocumentInlineFieldTopology(block, original);
    return {
      ...common,
      content: {
        case: "paragraph",
        value: { text: block.text, runs: block.runs.map((run) => documentRun(run, block.id, contentControlNativeIds.get(run))), formatting: documentParagraphFormatting(block) },
      },
    };
  }
  if (block.kind === "listItem") {
    const source = original?.content.case === "paragraph" ? original.content.value : undefined;
    if (!source?.numbering) {
      if (!directNumbering) {
        throw new OpenChestnutCodecError(`The DOCX WebAssembly vertical slice could not plan a numbering-definition graph for list item ${block.id}.`, [], { code: "invalid_document_numbering" });
      }
      const text = String(block.text ?? "");
      if (text.length > 1_000_000) throw new OpenChestnutCodecError(`Document list item ${block.id} text exceeds 1,000,000 characters.`, [], { code: "invalid_document_numbering" });
      return {
        ...common,
        content: { case: "paragraph", value: { text, numbering: directNumbering } },
      };
    }
    if (original.source?.editable === false) {
      throw new OpenChestnutCodecError(`Document list item ${block.id} is source-preserved but its paragraph topology is not editable.`, [], { code: "unsupported_document_edit" });
    }
    const numbering = editedDocumentNumbering(block, source.numbering);
    const text = String(block.text ?? "");
    if (text.length > 1_000_000) throw new OpenChestnutCodecError(`Document list item ${block.id} text exceeds 1,000,000 characters.`, [], { code: "invalid_document_numbering" });
    return {
      ...common,
      content: {
        case: "paragraph",
        value: {
          text,
          runs: source.runs.map((run) => ({ ...run, text })),
          numbering,
        },
      },
    };
  }
  if (block.kind === "table") {
    const source = original?.content.case === "table" ? original.content.value : undefined;
    const authored = !source && Array.isArray(block.cells) ? authoredDocumentTableGeometry(block) : undefined;
    if (source && !sameDocumentTableGeometry(block, source)) {
      throw new OpenChestnutCodecError(`Document table ${block.id} grid, span, merge, and per-cell editability metadata are source-bound.`, [], { code: "unsupported_document_edit" });
    }
    const formattingChanged = source && !sameDocumentTableFormatting(block, source);
    if (formattingChanged && !source.formatting) {
      throw new OpenChestnutCodecError(`Document table ${block.id} direct formatting can change only when OpenChestnut recognized the complete bounded profile during import.`, [], { code: "unsupported_document_edit" });
    }
    if (source) {
      for (let rowIndex = 0; rowIndex < source.rows.length; rowIndex += 1) {
        for (let cellIndex = 0; cellIndex < source.rows[rowIndex].cells.length; cellIndex += 1) {
          if (String(block.values?.[rowIndex]?.[cellIndex] ?? "") !== source.rows[rowIndex].cells[cellIndex] &&
              source.rows[rowIndex].richCells[cellIndex]?.editable === false) {
            throw new OpenChestnutCodecError(`Document table ${block.id} cell ${rowIndex},${cellIndex} is a vertical continuation or complex source cell and cannot be edited.`, [], { code: "unsupported_document_edit" });
          }
        }
      }
    }
    const textPatches = wireDocumentTableTextPatches(block, source);
    return {
      ...common,
      content: {
        case: "table",
        value: {
          ...(source ? { gridColumns: source.gridColumns } : authored ? { gridColumns: authored.gridColumns } : {}),
          ...(source ? (source.formatting ? {
            formatting: formattingChanged
              ? documentTableFormatting(block, source.gridColumns || Math.max(1, ...source.rows.map((row) => row.cells.length)))
              : { ...source.formatting, columnWidthsDxa: [...source.formatting.columnWidthsDxa], cellMarginsDxa: { ...source.formatting.cellMarginsDxa } },
          } : {}) : {
            formatting: documentTableFormatting(block, authored?.gridColumns || Math.max(1, block.columns)),
          }),
          rows: authored?.rows || (block.values || []).map((cells, rowIndex) => ({
            cells: cells.map((value) => String(value ?? "")),
            ...(source ? {
              richCells: source.rows[rowIndex]?.richCells.map((cell) => ({ ...cell })) || [],
              gridBefore: source.rows[rowIndex]?.gridBefore || 0,
              gridAfter: source.rows[rowIndex]?.gridAfter || 0,
            } : {}),
          })),
          textPatches,
        },
      },
    };
  }
  if (block.kind === "hyperlink") {
    return {
      ...common,
      content: { case: "hyperlink", value: documentHyperlink(block, original) },
    };
  }
  if (block.kind === "field") {
    return {
      ...common,
      content: { case: "field", value: documentField(block, original) },
    };
  }
  if (block.kind === "citation") {
    return {
      ...common,
      content: { case: "citation", value: wireDocumentCitation(block, original?.content.case === "citation" ? original.content.value : undefined) },
    };
  }
  if (block.kind === "change") {
    return {
      ...common,
      content: { case: "change", value: wireDocumentChange(block) },
    };
  }
  if (block.kind === "image") {
    return {
      ...common,
      content: { case: "image", value: documentImage(block, assets) },
    };
  }
  if (block.kind === "section") {
    return {
      ...common,
      content: { case: "section", value: wireDocumentSection(block) },
    };
  }
  throw new OpenChestnutCodecError(`The DOCX WebAssembly vertical slice cannot author document block kind ${block.kind}.`, [], { code: "unsupported_document_features" });
}

function unsupportedDocumentCollections(document) {
  const unsupported = [];
  if (document.settings?.trackRevisions) unsupported.push("revision tracking");
  if (document.settings?.mirrorMargins) unsupported.push("mirrored margins");
  if (document.settings?.documentProtection != null) unsupported.push("document protection");
  return unsupported;
}

function documentEnvelope(document) {
  if (!(document instanceof DocumentModel)) throw new TypeError("exportDocxWithOpenChestnut expects a DocumentModel instance.");
  const state = document[DOCUMENT_STATE];
  assertTrustedImportedState(state, "DOCX");
  const unsupported = unsupportedDocumentCollections(document);
  if (unsupported.length) {
    throw new OpenChestnutCodecError(`OpenChestnut cannot author or edit these DOCX features: ${unsupported.join(", ")}. This operation fails closed; preserve imported instances only through their validated source-bound package.`, [], { code: "unsupported_document_features" });
  }
  if (state && state.blocks.length !== document.blocks.length) {
    throw new OpenChestnutCodecError(`Source-preserving DOCX export requires the original ${state.blocks.length}-block topology; the document contains ${document.blocks.length} blocks.`, [], { code: "document_topology_changed" });
  }
  if (state && state.comments.length !== document.comments.length) {
    throw new OpenChestnutCodecError(`Source-preserving DOCX export requires the original ${state.comments.length}-comment topology; the document contains ${document.comments.length} comments.`, [], { code: "document_comment_topology_changed" });
  }
  validateDocumentCommentThreads(document);
  if (state && state.bookmarks.length !== document.bookmarks.length) {
    throw new OpenChestnutCodecError(`Source-preserving DOCX export requires the original ${state.bookmarks.length}-bookmark topology; the document contains ${document.bookmarks.length} bookmarks.`, [], { code: "document_bookmark_topology_changed" });
  }
  if (state && state.notes.length !== document.notes.length) {
    throw new OpenChestnutCodecError(`Source-preserving DOCX export requires the original ${state.notes.length}-note topology; the document contains ${document.notes.length} notes.`, [], { code: "document_note_topology_changed" });
  }
  if (state && Boolean(state.bibliography) !== Boolean(document.bibliographySources.length || Object.values(document.bibliography || {}).some(Boolean))) {
    throw new OpenChestnutCodecError("Source-preserving DOCX export cannot add or remove the modeled bibliography catalog.", [], { code: "document_bibliography_topology_changed" });
  }
  if (state && (state.headers.length !== document.headers.length || state.footers.length !== document.footers.length)) {
    throw new OpenChestnutCodecError("Source-preserving DOCX export requires the original header/footer topology.", [], { code: "document_header_footer_topology_changed" });
  }
  for (const slot of state?.readOnlyBlockSlots || []) {
    if (document.blocks[slot.index] !== slot.block || documentBlockSnapshot(slot.block) !== slot.publicSnapshot) {
      throw new OpenChestnutCodecError(`Imported document block ${slot.wire.id} is source-bound and read-only in OpenChestnut 0.2.`, [], { code: "unsupported_document_edit" });
    }
  }
  const directNumbering = state ? undefined : directDocumentNumberingPlan(document);
  const contentControlNativeIds = planDocumentTextContentControls(document);
  const assets = new Map((state?.assets || []).map((asset) => [asset.id, asset]));
  const defaultRunSource = Object.fromEntries([...DOCUMENT_RUN_STYLE_KEYS].filter((key) => key !== "runStyleId" && Object.hasOwn(document.defaultRunStyle || {}, key)).map((key) => [key, document.defaultRunStyle[key]]));
  const blocks = document.blocks.map((block, index) => documentBlock(block, state?.blocks[index], directNumbering?.get(block), assets, contentControlNativeIds));
  return {
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    family: ArtifactFamily.DOCUMENT,
    source: state?.source,
    opaqueOpc: state?.opaqueOpc,
    assets: [...assets.values()],
    diagnostics: state?.diagnostics || [],
    payload: {
      case: "document",
      value: {
        id: document.id,
        name: document.name,
        blocks,
        comments: document.comments.map((comment, index) => documentComment(comment, state?.comments[index])),
        bookmarks: document.bookmarks.map((bookmark, index) => documentBookmark(bookmark, state?.bookmarks[index], document)),
        notes: document.notes.map((note, index) => documentNote(note, state?.notes[index], document)),
        styles: document.styles.values().map(wireDocumentStyle),
        defaultRunStyle: documentRunFormatting(defaultRunSource, "Document default run style"),
        headers: document.headers.map(wireHeaderFooter),
        footers: document.footers.map(wireHeaderFooter),
        evenAndOddHeaders: Boolean(document.settings?.evenAndOddHeaders),
        updateFields: Boolean(document.settings?.updateFields),
        sectionSettings: (document.sectionSettings || []).map((settings) => ({
          sectionIndex: uint32(settings.sectionIndex, "Document section settings index"),
          differentFirstPage: settings.differentFirstPage == null ? undefined : Boolean(settings.differentFirstPage),
        })),
        bibliography: wireDocumentBibliography(document, state?.bibliography),
      },
    },
  };
}

export async function exportDocxWithOpenChestnut(document, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "exportDocxWithOpenChestnut");
  const response = await invokeOpenChestnut({
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    operation: CodecOperation.EXPORT_DOCX,
    family: ArtifactFamily.DOCUMENT,
    artifact: documentEnvelope(document),
    limits: codecLimits(options.limits),
  });
  return new FileBlob(response.file, {
    type: DOCX_MIME,
    metadata: { artifactKind: "document", codec: "open-chestnut", diagnostics: response.diagnostics },
  });
}

function documentFromEnvelope(envelope) {
  if (envelope.family !== ArtifactFamily.DOCUMENT || envelope.payload.case !== "document") {
    throw new OpenChestnutCodecError("OpenChestnut response does not contain a document artifact.", [], { code: "invalid_document_artifact" });
  }
  const source = envelope.payload.value;
  const assets = new Map((envelope.assets || []).map((asset) => [asset.id, asset]));
  const styles = Object.fromEntries((source.styles || []).map((style) => [style.id, publicDocumentStyle(style)]));
  if (!(source.styles || []).length) for (const block of source.blocks) {
    if (block.styleId) styles[block.styleId] = { id: block.styleId, name: block.styleId, type: block.content.case === "table" ? "table" : "paragraph" };
    if (block.content.case === "paragraph") for (const run of block.content.value.runs) if (run.styleId) styles[run.styleId] = { id: run.styleId, name: run.styleId, type: "character" };
  }
  const blocks = source.blocks.map((block) => {
    switch (block.content.case) {
      case "paragraph":
        if (block.content.value.numbering) {
          const paragraph = block.content.value;
          const numbering = paragraph.numbering;
          const numberFormat = numbering.numberFormat || "decimal";
          return {
            kind: "listItem",
            id: block.id,
            name: block.name,
            styleId: block.styleId || "Normal",
            text: paragraph.text,
            listType: numberFormat === "bullet" ? "bullet" : "number",
            numberFormat,
            level: numbering.level,
            start: numbering.start || 1,
            levelText: numbering.levelText || (numberFormat === "bullet" ? "•" : `%${numbering.level + 1}.`),
            numberingId: numbering.numberingId,
            abstractNumberingId: numbering.abstractNumberingId,
            numberingStyleId: numbering.numberingStyleId || undefined,
          };
        }
        return {
          kind: "paragraph",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          textEditable: block.source?.editable !== false,
          textPatchable: block.source?.textPatchable === true,
          textPatches: [],
          text: block.content.value.text,
          paragraphFormat: publicDocumentParagraphFormatting(block.content.value.formatting),
          runs: block.content.value.runs.length ? block.content.value.runs.map((run) => ({
            text: run.text,
            style: {
              ...(run.styleId ? { runStyleId: run.styleId } : {}),
              ...publicDocumentRunFormatting(run.formatting),
              ...(!run.formatting && run.bold ? { bold: true } : {}),
              ...(!run.formatting && run.italic ? { italic: true } : {}),
              ...(!run.formatting && run.underline ? { underline: true } : {}),
            },
            ...(run.textContentControl ? { contentControl: {
              id: run.textContentControl.id,
              tag: run.textContentControl.tag,
              alias: run.textContentControl.alias,
              nativeId: run.textContentControl.nativeId,
            } } : {}),
            ...(run.inlineField ? { inlineField: {
              instruction: run.inlineField.instruction,
              ...(run.inlineField.bookmarkName ? { bookmarkName: run.inlineField.bookmarkName } : {}),
              ...(run.inlineField.bookmarkNativeId !== "" ? { bookmarkNativeId: Number(run.inlineField.bookmarkNativeId) } : {}),
            } } : {}),
          })) : undefined,
        };
      case "table":
        {
          const formatting = documentTableFormattingConfig(block.content.value);
        return {
          kind: "table",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "TableGrid",
          values: block.content.value.rows.map((row) => [...row.cells]),
          gridColumns: block.content.value.gridColumns,
          cells: documentTableCells(block.content.value),
          textPatches: [],
          ...formatting,
        };
        }
      case "hyperlink": {
        const hyperlink = block.content.value;
        return {
          kind: "hyperlink",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          text: hyperlink.text,
          url: hyperlink.target.case === "externalUri" ? hyperlink.target.value : undefined,
          anchor: hyperlink.target.case === "internalAnchor" ? hyperlink.target.value : undefined,
          relationshipId: hyperlink.relationshipId || undefined,
          tooltip: hyperlink.tooltip,
          history: hyperlink.history,
        };
      }
      case "field":
        return {
          kind: "field",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          instruction: block.content.value.instruction,
          display: block.content.value.display,
          complex: Boolean(block.content.value.complex),
        };
      case "citation":
        return {
          kind: "citation",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          text: block.content.value.display,
          metadata: { tag: block.content.value.tag },
          _restore: true,
        };
      case "change": {
        const change = block.content.value;
        return {
          kind: "change",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          changeType: publicDocumentChangeType(change.type),
          text: change.text,
          author: change.author,
          date: change.date,
          _restore: true,
        };
      }
      case "image": {
        const image = block.content.value;
        const asset = assets.get(image.assetId);
        if (!asset || !new Set(["image/png", "image/jpeg"]).has(asset.contentType)) throw new OpenChestnutCodecError(`Document image ${block.id} references a missing or unsupported asset.`, [], { code: "invalid_document_asset" });
        return {
          kind: "image",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "Normal",
          dataUrl: `data:${asset.contentType};base64,${Buffer.from(asset.data).toString("base64")}`,
          alt: image.altText,
          widthPx: Number(image.widthEmu) / 9_525,
          heightPx: Number(image.heightEmu) / 9_525,
        };
      }
      case "section": {
        const section = block.content.value;
        return {
          kind: "section",
          id: block.id,
          name: block.name,
          breakType: publicDocumentSectionBreak(section.breakType),
          orientation: section.landscape ? "landscape" : "portrait",
          pageSize: { widthTwips: section.pageWidthTwips, heightTwips: section.pageHeightTwips },
          margins: { top: section.marginTopTwips, right: section.marginRightTwips, bottom: section.marginBottomTwips, left: section.marginLeftTwips },
        };
      }
      case "opaque":
        return {
          kind: "paragraph",
          id: block.id,
          name: block.name || `Preserved ${block.content.value.elementName}`,
          styleId: "Normal",
          textEditable: false,
          textPatchable: false,
          text: block.content.value.text,
        };
      default:
        throw new OpenChestnutCodecError(`Document block ${block.id} has no supported wire content.`, [], { code: "invalid_document_artifact" });
    }
  });
  const comments = source.comments.map((comment) => ({
    id: comment.id,
    targetId: comment.targetBlockId,
    parentId: comment.parentCommentId || undefined,
    author: comment.author,
    initials: comment.initials,
    date: comment.createdAt,
    text: comment.text,
    resolved: comment.resolved,
    paraId: comment.paragraphId || undefined,
    durableId: comment.durableId || undefined,
    dateUtc: comment.dateUtc,
    person: comment.person ? { providerId: comment.person.providerId, userId: comment.person.userId } : undefined,
    intelligentPlaceholder: comment.intelligentPlaceholder ?? false,
  }));
  const bookmarks = (source.bookmarks || []).map((bookmark) => ({
    id: bookmark.id,
    name: bookmark.name,
    targetId: bookmark.targetBlockId,
    endTargetId: bookmark.endTargetBlockId,
    nativeId: bookmark.nativeId === "" ? undefined : Number(bookmark.nativeId),
  }));
  const notes = (source.notes || []).map((note) => ({
    id: note.id,
    kind: publicDocumentNoteKind(note.kind),
    targetId: note.targetBlockId,
    text: note.text,
    nativeId: note.nativeId === "" ? undefined : Number(note.nativeId),
  }));
  const document = DocumentModel.create({
    name: source.name || "Imported document",
    styles,
    defaultRunStyle: publicDocumentRunFormatting(source.defaultRunStyle),
    blocks,
    comments,
    bookmarks,
    notes,
    bibliography: source.bibliography ? {
      selectedStyle: source.bibliography.selectedStyle,
      styleName: source.bibliography.styleName,
      uri: source.bibliography.uri,
    } : undefined,
    bibliographySources: (source.bibliography?.sources || []).map(publicDocumentBibliographySource),
    headers: (source.headers || []).map(publicHeaderFooter),
    footers: (source.footers || []).map(publicHeaderFooter),
    settings: { evenAndOddHeaders: Boolean(source.evenAndOddHeaders), updateFields: Boolean(source.updateFields) },
    sectionSettings: (source.sectionSettings || []).map((settings) => ({ sectionIndex: settings.sectionIndex, differentFirstPage: settings.differentFirstPage })),
  });
  document.id = source.id || document.id;
  const commentSlots = source.comments.map((wire, index) => ({
    wire,
    publicSnapshot: documentCommentSnapshot(document.comments[index]),
  }));
  const bookmarkSlots = (source.bookmarks || []).map((wire, index) => ({
    wire,
    publicSnapshot: documentBookmarkSnapshot(document.bookmarks[index]),
  }));
  const noteSlots = (source.notes || []).map((wire, index) => ({
    wire,
    publicSnapshot: documentNoteSnapshot(document.notes[index]),
  }));
  const readOnlyBlockSlots = source.blocks.flatMap((wire, index) => {
    if (wire.content.case !== "opaque" && (wire.source?.editable !== false || wire.source?.textPatchable === true)) return [];
    const block = document.blocks[index];
    return [{ wire, index, block, publicSnapshot: documentBlockSnapshot(block) }];
  });
  Object.defineProperty(document, DOCUMENT_STATE, {
    configurable: true,
    value: { source: envelope.source, opaqueOpc: envelope.opaqueOpc, diagnostics: envelope.diagnostics, assets: envelope.assets || [], blocks: source.blocks, readOnlyBlockSlots, comments: commentSlots, bookmarks: bookmarkSlots, notes: noteSlots, bibliography: source.bibliography, headers: source.headers || [], footers: source.footers || [] },
    writable: true,
  });
  return document;
}

export async function importDocxWithOpenChestnut(input, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "importDocxWithOpenChestnut");
  const response = await invokeOpenChestnut({
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    operation: CodecOperation.IMPORT_DOCX,
    family: ArtifactFamily.DOCUMENT,
    file: await inputBytes(input),
    limits: codecLimits(options.limits),
  });
  return documentFromEnvelope(response.artifact);
}

export async function exportPptxWithOpenChestnut(presentation, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "exportPptxWithOpenChestnut");
  const response = await invokeOpenChestnut({
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    operation: CodecOperation.EXPORT_PPTX,
    family: ArtifactFamily.PRESENTATION,
    artifact: presentationEnvelope(presentation, OPEN_CHESTNUT_PROTOCOL_VERSION),
    limits: codecLimits(options.limits),
  });
  return new FileBlob(response.file, {
    type: PPTX_MIME,
    metadata: { artifactKind: "presentation", codec: "open-chestnut", diagnostics: response.diagnostics },
  });
}

export async function importPptxWithOpenChestnut(input, options = {}) {
  assertCodecOptions(options, new Set(["limits"]), "importPptxWithOpenChestnut");
  const response = await invokeOpenChestnut({
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    operation: CodecOperation.IMPORT_PPTX,
    family: ArtifactFamily.PRESENTATION,
    file: await inputBytes(input),
    limits: codecLimits(options.limits),
  });
  return presentationFromEnvelope(response.artifact);
}

export async function openChestnutStatus() {
  const [loaded, manifestText] = await Promise.all([runtime(), readFile(MANIFEST_URL, "utf8")]);
  const manifest = JSON.parse(manifestText);
  return { available: true, protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION, assemblyName: loaded.assemblyName, manifest };
}
