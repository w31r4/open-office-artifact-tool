import { create, fromBinary, toBinary } from "@bufbuild/protobuf";
import { readFile } from "node:fs/promises";
import { FileBlob, Workbook } from "../index.mjs";
import {
  ArtifactFamily,
  CodecOperation,
  CodecRequestSchema,
  CodecResponseSchema,
  WorkbookDateSystem,
} from "../generated/open_office/artifact/v1/office_artifact_pb.js";

export const OPEN_XML_WASM_PROTOCOL_VERSION = 1;

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const RUNTIME_URL = new URL("../../runtime/openxml-wasm/main.mjs", import.meta.url);
const MANIFEST_URL = new URL("../../runtime/openxml-wasm/manifest.json", import.meta.url);
const WORKBOOK_STATE = Symbol.for("open-office-artifact-tool.openxml-wasm-state");
const EXCEL_ERRORS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#UNKNOWN!", "#CONNECT!", "#CYCLE!"]);
const DEFAULT_THEME_COLORS = {
  dk1: "#000000", lt1: "#FFFFFF", dk2: "#1F497D", lt2: "#EEECE1",
  accent1: "#4F81BD", accent2: "#C0504D", accent3: "#9BBB59", accent4: "#8064A2",
  accent5: "#4BACC6", accent6: "#F79646", hlink: "#0000FF", folHlink: "#800080",
};

let runtimePromise;

export class OpenXmlWasmCodecError extends Error {
  constructor(message, diagnostics = [], options = {}) {
    super(message, options);
    this.name = "OpenXmlWasmCodecError";
    this.code = diagnostics[0]?.code || options.code || "openxml_wasm_codec_error";
    this.diagnostics = diagnostics;
  }
}

async function runtime() {
  if (!runtimePromise) {
    runtimePromise = import(RUNTIME_URL.href)
      .then((module) => module.loadOpenXmlWasm())
      .catch((error) => {
        runtimePromise = undefined;
        throw new OpenXmlWasmCodecError("Bundled OpenXML WebAssembly runtime could not be loaded.", [], { code: "runtime_unavailable", cause: error });
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
    : "OpenXML WebAssembly codec returned an unspecified failure.";
  return new OpenXmlWasmCodecError(message, response.diagnostics);
}

export async function invokeOpenXmlWasm(request) {
  const loaded = await runtime();
  const wireRequest = create(CodecRequestSchema, request);
  const wireResponse = bytesFrom(loaded.invoke(toBinary(CodecRequestSchema, wireRequest)));
  const response = fromBinary(CodecResponseSchema, wireResponse);
  if (!response.ok) throw responseFailure(response);
  return response;
}

function cellCoordinates(address) {
  const match = /^([A-Z]{1,3})([1-9]\d*)$/i.exec(String(address));
  if (!match) throw new OpenXmlWasmCodecError(`Cell address ${address} is not valid A1 notation.`, [], { code: "invalid_cell_address" });
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

function itemCount(collection) {
  return Array.isArray(collection?.items) ? collection.items.length : 0;
}

function unsupportedWorkbookFeatures(workbook) {
  const unsupported = [];
  if (workbook.definedNames?.items?.length) unsupported.push("defined names");
  if (workbook.comments?.threads?.length) unsupported.push("threaded comments");
  if (workbook.indexedColors?.length) unsupported.push("custom indexed colors");
  if (workbook.theme?.name !== "Office Clean Room" || Object.entries(DEFAULT_THEME_COLORS).some(([key, value]) => workbook.theme?.colors?.[key] !== value)) unsupported.push("custom workbook theme");
  for (const sheet of workbook.worksheets?.items || []) {
    const prefix = `worksheet ${sheet.name}`;
    if (itemCount(sheet.tables)) unsupported.push(`${prefix} tables`);
    if (itemCount(sheet.pivotTables)) unsupported.push(`${prefix} pivot tables`);
    if (itemCount(sheet.charts)) unsupported.push(`${prefix} charts`);
    if (itemCount(sheet.images)) unsupported.push(`${prefix} images`);
    if (itemCount(sheet.sparklineGroups)) unsupported.push(`${prefix} sparklines`);
    if (sheet.shapes?.length) unsupported.push(`${prefix} shapes`);
    if (sheet.dataValidations?.items?.length) unsupported.push(`${prefix} data validations`);
    if (sheet.conditionalFormattings?.items?.length) unsupported.push(`${prefix} conditional formatting`);
    for (const [address, cell] of sheet.store?.entries?.() || []) {
      if (cell.style && Object.keys(cell.style).length) unsupported.push(`${prefix} styled cell ${address}`);
      const metadata = Object.keys(cell).filter((key) => !["value", "formula", "style"].includes(key));
      if (metadata.length) unsupported.push(`${prefix} advanced formula metadata at ${address}`);
    }
  }
  return unsupported;
}

function wireCell(address, cell) {
  const coordinates = cellCoordinates(address);
  const target = {
    row: coordinates.row,
    column: coordinates.column,
    formula: cell.formula ? String(cell.formula) : "",
    value: { case: undefined },
  };
  if (cell.value == null) return target;
  if (typeof cell.value === "string") {
    target.value = EXCEL_ERRORS.has(cell.value) ? { case: "errorValue", value: cell.value } : { case: "stringValue", value: cell.value };
  } else if (typeof cell.value === "number") {
    if (!Number.isFinite(cell.value)) throw new OpenXmlWasmCodecError(`Cell ${address} has a non-finite numeric value.`, [], { code: "non_finite_cell_value" });
    target.value = { case: "numberValue", value: cell.value };
  } else if (typeof cell.value === "boolean") {
    target.value = { case: "boolValue", value: cell.value };
  } else {
    throw new OpenXmlWasmCodecError(`Cell ${address} has unsupported ${cell.value?.constructor?.name || typeof cell.value} content.`, [], { code: "unsupported_cell_value" });
  }
  return target;
}

function workbookEnvelope(workbook) {
  if (!(workbook instanceof Workbook)) throw new TypeError("exportXlsxWithOpenXmlWasm expects a Workbook instance.");
  if (!workbook.worksheets?.items?.length) throw new OpenXmlWasmCodecError("Workbook must contain at least one worksheet.", [], { code: "missing_worksheets" });
  const unsupported = unsupportedWorkbookFeatures(workbook);
  if (unsupported.length) {
    throw new OpenXmlWasmCodecError(`The XLSX WebAssembly vertical slice cannot encode: ${unsupported.slice(0, 8).join(", ")}${unsupported.length > 8 ? `, and ${unsupported.length - 8} more` : ""}. Use SpreadsheetFile.exportXlsx until parity reaches these features.`, [], { code: "unsupported_workbook_features" });
  }
  const state = workbook[WORKBOOK_STATE];
  return {
    protocolVersion: OPEN_XML_WASM_PROTOCOL_VERSION,
    family: ArtifactFamily.WORKBOOK,
    source: state?.source,
    opaqueOpc: state?.opaqueOpc,
    diagnostics: state?.diagnostics || [],
    payload: {
      case: "workbook",
      value: {
        id: workbook.id,
        dateSystem: workbook.dateSystem === "1904" ? WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1904 : WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1900,
        worksheets: workbook.worksheets.items.map((sheet) => ({
          id: sheet.id,
          name: sheet.name,
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
          cells: (sheet.store?.entries?.() || []).filter(([, cell]) => cell.value != null || cell.formula).map(([address, cell]) => wireCell(address, cell)),
        })),
      },
    },
  };
}

export async function exportXlsxWithOpenXmlWasm(workbook, options = {}) {
  if (!(workbook instanceof Workbook)) throw new TypeError("exportXlsxWithOpenXmlWasm expects a Workbook instance.");
  if (options.recalculate !== false) workbook.recalculate();
  const response = await invokeOpenXmlWasm({
    protocolVersion: OPEN_XML_WASM_PROTOCOL_VERSION,
    operation: CodecOperation.EXPORT_XLSX,
    family: ArtifactFamily.WORKBOOK,
    artifact: workbookEnvelope(workbook),
    limits: codecLimits(options.limits),
    allowLossy: options.allowLossy === true,
  });
  return new FileBlob(response.file, {
    type: XLSX_MIME,
    metadata: { artifactKind: "workbook", codec: "openxml-wasm", diagnostics: response.diagnostics },
  });
}

function workbookFromEnvelope(envelope) {
  if (envelope.family !== ArtifactFamily.WORKBOOK || envelope.payload.case !== "workbook") {
    throw new OpenXmlWasmCodecError("OpenXML WebAssembly response does not contain a workbook artifact.", [], { code: "invalid_workbook_artifact" });
  }
  const source = envelope.payload.value;
  const workbook = Workbook.create({ dateSystem: source.dateSystem === WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1904 ? "1904" : "1900" });
  workbook.id = source.id || workbook.id;
  for (const sourceSheet of source.worksheets) {
    const sheet = workbook.worksheets.add(sourceSheet.name);
    sheet.id = sourceSheet.id || sheet.id;
    sheet.showGridLines = sourceSheet.showGridLines;
    if (sourceSheet.freezePane) {
      sheet.freezePanes.freezeRows(sourceSheet.freezePane.rows);
      sheet.freezePanes.freezeColumns(sourceSheet.freezePane.columns);
    }
    for (const dimension of sourceSheet.columnDimensions) sheet.columnDimensions.set(dimension.column, { width: dimension.width || undefined, hidden: dimension.hidden, bestFit: dimension.bestFit });
    for (const dimension of sourceSheet.rowDimensions) sheet.rowDimensions.set(dimension.row, { height: dimension.height || undefined, hidden: dimension.hidden });
    sheet.mergedRanges = [...sourceSheet.mergedRanges];
    for (const sourceCell of sourceSheet.cells) {
      const cell = sheet.store.get(cellAddress(sourceCell.row, sourceCell.column));
      cell.formula = sourceCell.formula || null;
      switch (sourceCell.value.case) {
        case "stringValue": cell.value = sourceCell.value.value; break;
        case "numberValue": cell.value = sourceCell.value.value; break;
        case "boolValue": cell.value = sourceCell.value.value; break;
        case "errorValue": cell.value = sourceCell.value.value; break;
        default: cell.value = null;
      }
    }
  }
  Object.defineProperty(workbook, WORKBOOK_STATE, {
    configurable: true,
    value: { source: envelope.source, opaqueOpc: envelope.opaqueOpc, diagnostics: envelope.diagnostics },
    writable: true,
  });
  return workbook;
}

export async function importXlsxWithOpenXmlWasm(input, options = {}) {
  const response = await invokeOpenXmlWasm({
    protocolVersion: OPEN_XML_WASM_PROTOCOL_VERSION,
    operation: CodecOperation.IMPORT_XLSX,
    family: ArtifactFamily.WORKBOOK,
    file: await inputBytes(input),
    limits: codecLimits(options.limits),
  });
  return workbookFromEnvelope(response.artifact);
}

export async function openXmlWasmStatus() {
  const [loaded, manifestText] = await Promise.all([runtime(), readFile(MANIFEST_URL, "utf8")]);
  const manifest = JSON.parse(manifestText);
  return { available: true, protocolVersion: OPEN_XML_WASM_PROTOCOL_VERSION, assemblyName: loaded.assemblyName, manifest };
}
