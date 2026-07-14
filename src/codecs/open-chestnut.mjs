import { create, fromBinary, toBinary } from "@bufbuild/protobuf";
import { readFile } from "node:fs/promises";
import { DocumentModel, FileBlob, Workbook } from "../index.mjs";
import { XLSX_THEME_COLOR_NAMES, normalizeXlsxStyle, normalizeXlsxThemeConfig } from "../spreadsheet/ooxml-styles.mjs";
import {
  ArtifactFamily,
  CellFormulaKind,
  CodecOperation,
  CodecRequestSchema,
  CodecResponseSchema,
  DocumentTableVerticalMerge,
  WorkbookDateSystem,
} from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";
import { presentationEnvelope, presentationFromEnvelope } from "./open-chestnut-presentation.mjs";

export { OpenChestnutCodecError } from "./open-chestnut-error.mjs";

export const OPEN_CHESTNUT_PROTOCOL_VERSION = 1;

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const RUNTIME_URL = new URL("../../runtime/open-chestnut/main.mjs", import.meta.url);
const MANIFEST_URL = new URL("../../runtime/open-chestnut/manifest.json", import.meta.url);
const WORKBOOK_STATE = Symbol.for("open-office-artifact-tool.open-chestnut-state");
const TABLE_STATE = Symbol.for("open-office-artifact-tool.open-chestnut-table-state");
const DOCUMENT_STATE = Symbol.for("open-office-artifact-tool.open-chestnut-document-state");
const MAX_XLSX_NUMBER_FORMAT_CODE_LENGTH = 4096;
const MAX_XLSX_FORMULA_LENGTH = 8192;
const MAX_XLSX_FORMULA_TOPOLOGY_CELLS = 1_048_576;
const XLSX_FORMULA_METADATA_KEYS = new Set(["formulaType", "sharedIndex", "sharedRef", "arrayRef"]);
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
  const loaded = await runtime();
  const wireRequest = create(CodecRequestSchema, request);
  const wireResponse = bytesFrom(loaded.invoke(toBinary(CodecRequestSchema, wireRequest)));
  const response = fromBinary(CodecResponseSchema, wireResponse);
  if (!response.ok) throw responseFailure(response);
  return response;
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
  if (!type && [cell.sharedIndex, cell.sharedRef, cell.arrayRef].every((value) => value == null || value === "")) return undefined;
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
  throw new OpenChestnutCodecError(`Cell ${location} formula type ${type || "unspecified"} is outside the OpenChestnut XLSX formula slice.`, [], { code: "unsupported_cell_formula" });
}

function validateFormulaTopology(cells, sheetName) {
  const byCoordinate = new Map(cells.map((cell) => [`${cell.row}:${cell.column}`, cell]));
  if (byCoordinate.size !== cells.length) throw new OpenChestnutCodecError(`Worksheet ${sheetName} contains duplicate cell coordinates.`, [], { code: "duplicate_cell" });
  const sharedGroups = new Map();
  for (const cell of cells) {
    validateFormulaText(cell.formula, `${sheetName}!${cellAddress(cell.row, cell.column)}`, Boolean(cell.formulaMetadata));
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
  const topologyRoots = [...sharedRoots, ...cells.filter((item) => item.formulaMetadata?.kind === CellFormulaKind.ARRAY)];
  let topologyCellCount = 0;
  for (const cell of topologyRoots) {
    const metadata = cell.formulaMetadata;
    const bounds = formulaRangeBounds(metadata.reference, `${sheetName}!${cellAddress(cell.row, cell.column)}`);
    topologyCellCount += bounds.cellCount;
    if (topologyCellCount > MAX_XLSX_FORMULA_TOPOLOGY_CELLS) throw new OpenChestnutCodecError(`Cell ${sheetName}!${cellAddress(cell.row, cell.column)} native formula topology exceeds ${MAX_XLSX_FORMULA_TOPOLOGY_CELLS} cells.`, [], { code: "invalid_cell_formula" });
    const owner = metadata.kind === CellFormulaKind.SHARED ? `shared:${metadata.sharedIndex}` : `array:${cell.row}:${cell.column}`;
    if (metadata.kind === CellFormulaKind.ARRAY && (cell.row !== bounds.top || cell.column !== bounds.left)) throw new OpenChestnutCodecError(`Cell ${sheetName}!${cellAddress(cell.row, cell.column)} legacy array formula must be the top-left anchor of ${metadata.reference}.`, [], { code: "invalid_cell_formula" });
    for (let row = bounds.top; row <= bounds.bottom; row += 1) {
      for (let column = bounds.left; column <= bounds.right; column += 1) {
        const key = `${row}:${column}`;
        if (occupied.has(key) && occupied.get(key) !== owner) throw new OpenChestnutCodecError(`Cell ${sheetName}!${cellAddress(cell.row, cell.column)} formula range ${metadata.reference} overlaps another native formula range.`, [], { code: "invalid_cell_formula" });
        occupied.set(key, owner);
        const nested = byCoordinate.get(key);
        if (metadata.kind === CellFormulaKind.ARRAY && (row !== cell.row || column !== cell.column) && nested?.formula) throw new OpenChestnutCodecError(`Cell ${sheetName}!${cellAddress(row, column)} must not contain another formula inside legacy array range ${metadata.reference}.`, [], { code: "invalid_cell_formula" });
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

function unsupportedWorkbookFeatures(workbook) {
  const unsupported = [];
  if (workbook.definedNames?.items?.length) unsupported.push("defined names");
  if (workbook.comments?.threads?.length) unsupported.push("threaded comments");
  if (workbook.indexedColors?.length) unsupported.push("custom indexed colors");
  for (const sheet of workbook.worksheets?.items || []) {
    const prefix = `worksheet ${sheet.name}`;
    if (itemCount(sheet.pivotTables)) unsupported.push(`${prefix} pivot tables`);
    if (itemCount(sheet.charts)) unsupported.push(`${prefix} charts`);
    if (itemCount(sheet.images)) unsupported.push(`${prefix} images`);
    if (itemCount(sheet.sparklineGroups)) unsupported.push(`${prefix} sparklines`);
    if (sheet.shapes?.length) unsupported.push(`${prefix} shapes`);
    if (sheet.dataValidations?.items?.length) unsupported.push(`${prefix} data validations`);
    if (sheet.conditionalFormattings?.items?.length) unsupported.push(`${prefix} conditional formatting`);
    for (const [address, cell] of sheet.store?.entries?.() || []) {
      if (cell.style && Object.keys(cell.style).some((key) => cell.style[key] != null)) wireCellStyle(cell.style, `${sheet.name}!${address}`);
      const metadata = Object.keys(cell).filter((key) => !["value", "formula", "style"].includes(key) && !XLSX_FORMULA_METADATA_KEYS.has(key));
      if (metadata.length) unsupported.push(`${prefix} advanced formula metadata at ${address}`);
    }
  }
  return unsupported;
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

function wireWorkbookConnection(value, source) {
  return { ...publicWorkbookConnection(value), source };
}

function wireWorkbookConnections(workbook, state) {
  const remaining = new Set(workbook.connections || []);
  const output = [];
  for (const slot of state?.connectionSlots || []) {
    if (!remaining.delete(slot.connection)) {
      throw new OpenChestnutCodecError(`Workbook cannot remove imported connection ${slot.connection.connectionId} in the bounded OpenChestnut slice.`, [], { code: "invalid_workbook_connection" });
    }
    output.push(JSON.stringify(connectionSnapshot(slot.connection)) === JSON.stringify(slot.publicSnapshot)
      ? slot.wire
      : wireWorkbookConnection(slot.connection, slot.wire.source));
  }
  output.push(...[...remaining].map((connection) => wireWorkbookConnection(connection)));
  return output;
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
  const query = publicTableQuery(table.queryTable);
  if (!query) return undefined;
  if (table.queryTable?.refresh) {
    query.refresh = publicTableQueryRefresh(table.queryTable.refresh);
    if (table.queryTable.refresh.sortState)
      query.refresh.sortState = wireTableSortState(table.queryTable.refresh.sortState, `table ${table.name} query refresh`);
  }
  return { ...query, source: table[TABLE_STATE]?.wire?.queryTable?.source };
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

function wireCell(address, cell) {
  const coordinates = cellCoordinates(address);
  const target = {
    row: coordinates.row,
    column: coordinates.column,
    formula: cell.formula ? String(cell.formula) : "",
    formulaMetadata: cellFormulaMetadata(address, cell),
    numberFormatCode: cellNumberFormatCode(cell, address),
    style: wireCellStyle(cell.style, address),
    value: { case: undefined },
  };
  if (cell.value == null) return target;
  if (typeof cell.value === "string") {
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
  const unsupported = unsupportedWorkbookFeatures(workbook);
  if (unsupported.length) {
    throw new OpenChestnutCodecError(`The XLSX WebAssembly vertical slice cannot encode: ${unsupported.slice(0, 8).join(", ")}${unsupported.length > 8 ? `, and ${unsupported.length - 8} more` : ""}. Use SpreadsheetFile.exportXlsx until parity reaches these features.`, [], { code: "unsupported_workbook_features" });
  }
  const state = workbook[WORKBOOK_STATE];
  const theme = state?.themeWire && sameWorkbookTheme(workbook.theme, state.publicTheme)
    ? state.themeWire
    : wireWorkbookTheme(workbook.theme, state?.themeWire?.source);
  return {
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    family: ArtifactFamily.WORKBOOK,
    source: state?.source,
    opaqueOpc: state?.opaqueOpc,
    diagnostics: state?.diagnostics || [],
    payload: {
      case: "workbook",
      value: {
        id: workbook.id,
        dateSystem: workbook.dateSystem === "1904" ? WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1904 : WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1900,
        theme,
        connections: wireWorkbookConnections(workbook, state),
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
          tables: wireWorksheetTables(sheet, state?.tablesBySheet?.get(sheet.id)),
          cells: (() => {
            const cells = (sheet.store?.entries?.() || []).filter(([, cell]) => cell.value != null || cell.formula || cell.formulaType || Object.keys(cell.style || {}).some((key) => cell.style[key] != null)).map(([address, cell]) => wireCell(address, cell));
            validateFormulaTopology(cells, sheet.name);
            return cells;
          })(),
        })),
      },
    },
  };
}

export async function exportXlsxWithOpenChestnut(workbook, options = {}) {
  if (!(workbook instanceof Workbook)) throw new TypeError("exportXlsxWithOpenChestnut expects a Workbook instance.");
  if (options.recalculate !== false) workbook.recalculate();
  const response = await invokeOpenChestnut({
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    operation: CodecOperation.EXPORT_XLSX,
    family: ArtifactFamily.WORKBOOK,
    artifact: workbookEnvelope(workbook),
    limits: codecLimits(options.limits),
    allowLossy: options.allowLossy === true,
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
  const workbook = Workbook.create({
    dateSystem: source.dateSystem === WorkbookDateSystem.WORKBOOK_DATE_SYSTEM_1904 ? "1904" : "1900",
    ...(importedTheme ? { theme: importedTheme } : {}),
    connections: importedConnections,
  });
  workbook.id = source.id || workbook.id;
  const tablesBySheet = new Map();
  const connectionSlots = (source.connections || []).map((wire, index) => ({
    wire,
    connection: workbook.connections[index],
    publicSnapshot: connectionSnapshot(workbook.connections[index]),
  }));
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
      if (sourceCell.formulaMetadata?.kind === CellFormulaKind.SHARED) {
        cell.formulaType = "shared";
        cell.sharedIndex = sourceCell.formulaMetadata.sharedIndex;
        cell.sharedRef = sourceCell.formulaMetadata.reference;
      } else if (sourceCell.formulaMetadata?.kind === CellFormulaKind.ARRAY) {
        cell.formulaType = "array";
        cell.arrayRef = sourceCell.formulaMetadata.reference;
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
    }
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
      Object.defineProperty(table, TABLE_STATE, { configurable: true, value: { wire: sourceTable }, writable: true });
      slots.push({ wire: sourceTable, table, publicSnapshot });
    }
    tablesBySheet.set(sheet.id, { slots });
  }
  Object.defineProperty(workbook, WORKBOOK_STATE, {
    configurable: true,
    value: {
      source: envelope.source,
      opaqueOpc: envelope.opaqueOpc,
      diagnostics: envelope.diagnostics,
      themeWire: source.theme,
      publicTheme: normalizeXlsxThemeConfig(workbook.theme),
      connectionSlots,
      tablesBySheet,
    },
    writable: true,
  });
  return workbook;
}

export async function importXlsxWithOpenChestnut(input, options = {}) {
  const response = await invokeOpenChestnut({
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    operation: CodecOperation.IMPORT_XLSX,
    family: ArtifactFamily.WORKBOOK,
    file: await inputBytes(input),
    limits: codecLimits(options.limits),
  });
  return workbookFromEnvelope(response.artifact);
}

const DOCUMENT_RUN_STYLE_KEYS = new Set(["runStyleId", "bold", "italic", "underline"]);
const DOCUMENT_FIELD_COMMANDS = new Set(["PAGE", "NUMPAGES", "SECTION", "SECTIONPAGES", "DATE", "TIME", "CREATEDATE", "SAVEDATE", "PRINTDATE", "AUTHOR", "TITLE", "SUBJECT", "COMMENTS", "FILENAME", "FILESIZE", "NUMWORDS", "NUMCHARS"]);

function documentRun(run, blockId) {
  const style = run.style || {};
  const unsupported = Object.keys(style).filter((key) => !DOCUMENT_RUN_STYLE_KEYS.has(key));
  if (unsupported.length) {
    throw new OpenChestnutCodecError(`Document block ${blockId} uses unsupported run style fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_document_features" });
  }
  return {
    text: String(run.text ?? ""),
    styleId: style.runStyleId || "",
    bold: style.bold === true,
    italic: style.italic === true,
    underline: style.underline === true || style.underline === "single",
  };
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
      cell.editable === source.editable;
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
  if (!command || !DOCUMENT_FIELD_COMMANDS.has(command)) {
    throw new OpenChestnutCodecError(`Document field ${block.id} command ${command || "(missing)"} is outside the bounded editable field catalog.`, [], { code: "invalid_document_field" });
  }
  if (display.length > 1_000_000) throw new OpenChestnutCodecError(`Document field ${block.id} display text exceeds 1,000,000 characters.`, [], { code: "invalid_document_field" });
  return { instruction, display };
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
      if (block.kind !== "table" || !sameTableValues(block, original) ||
          !sameDocumentTableGeometry(block, original.content.value) ||
          !sameDocumentTableFormatting(block, original.content.value)) return false;
      return block.styleId === original.styleId || (!original.styleId && block.styleId === "TableGrid");
    }
    case "hyperlink":
      return sameDocumentHyperlink(block, original.content.value);
    case "field":
      return block.kind === "field" && block.styleId === (original.styleId || "Normal") && block.instruction === original.content.value.instruction && block.display === original.content.value.display;
    case "opaque":
      return block.kind === "paragraph" && block.text === original.content.value.text && block.runs.every((run) => Object.keys(run.style || {}).length === 0);
    default:
      return false;
  }
}

function documentBlock(block, original, directNumbering) {
  if (original && unchangedSourceBlock(block, original)) return original;
  const common = {
    id: original?.id || block.id,
    name: block.name || original?.name || "",
    styleId: block.styleId || original?.styleId || "",
    source: original?.source,
  };
  if (block.kind === "paragraph") {
    return {
      ...common,
      content: {
        case: "paragraph",
        value: { text: block.text, runs: block.runs.map((run) => documentRun(run, block.id)) },
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
  throw new OpenChestnutCodecError(`The DOCX WebAssembly vertical slice cannot author document block kind ${block.kind}.`, [], { code: "unsupported_document_features" });
}

function unsupportedDocumentCollections(document) {
  const unsupported = [];
  for (const [label, value] of [
    ["bookmarks", document.bookmarks],
    ["comments", document.comments],
    ["bibliography sources", document.bibliographySources],
    ["headers", document.headers],
    ["footers", document.footers],
    ["section settings", document.sectionSettings],
  ]) {
    if (value?.length) unsupported.push(label);
  }
  if (document.bibliography && Object.values(document.bibliography).some(Boolean)) unsupported.push("bibliography settings");
  return unsupported;
}

function documentEnvelope(document) {
  if (!(document instanceof DocumentModel)) throw new TypeError("exportDocxWithOpenChestnut expects a DocumentModel instance.");
  const state = document[DOCUMENT_STATE];
  if (!state) {
    const unsupported = unsupportedDocumentCollections(document);
    if (unsupported.length) {
      throw new OpenChestnutCodecError(`The DOCX WebAssembly vertical slice cannot author: ${unsupported.join(", ")}. Use DocumentFile.exportDocx until parity reaches these features.`, [], { code: "unsupported_document_features" });
    }
  }
  if (state && state.blocks.length !== document.blocks.length) {
    throw new OpenChestnutCodecError(`Source-preserving DOCX export requires the original ${state.blocks.length}-block topology; the document contains ${document.blocks.length} blocks.`, [], { code: "document_topology_changed" });
  }
  const directNumbering = state ? undefined : directDocumentNumberingPlan(document);
  return {
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    family: ArtifactFamily.DOCUMENT,
    source: state?.source,
    opaqueOpc: state?.opaqueOpc,
    diagnostics: state?.diagnostics || [],
    payload: {
      case: "document",
      value: {
        id: document.id,
        name: document.name,
        blocks: document.blocks.map((block, index) => documentBlock(block, state?.blocks[index], directNumbering?.get(block))),
      },
    },
  };
}

export async function exportDocxWithOpenChestnut(document, options = {}) {
  const response = await invokeOpenChestnut({
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    operation: CodecOperation.EXPORT_DOCX,
    family: ArtifactFamily.DOCUMENT,
    artifact: documentEnvelope(document),
    limits: codecLimits(options.limits),
    allowLossy: options.allowLossy === true,
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
  const styles = {};
  for (const block of source.blocks) {
    if (block.styleId) styles[block.styleId] = { id: block.styleId, name: block.styleId, type: block.content.case === "table" ? "table" : "paragraph" };
    if (block.content.case === "paragraph") {
      for (const run of block.content.value.runs) {
        if (run.styleId) styles[run.styleId] = { id: run.styleId, name: run.styleId, type: "character" };
      }
    }
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
          text: block.content.value.text,
          runs: block.content.value.runs.length ? block.content.value.runs.map((run) => ({
            text: run.text,
            style: {
              ...(run.styleId ? { runStyleId: run.styleId } : {}),
              ...(run.bold ? { bold: true } : {}),
              ...(run.italic ? { italic: true } : {}),
              ...(run.underline ? { underline: true } : {}),
            },
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
        };
      case "opaque":
        return {
          kind: "paragraph",
          id: block.id,
          name: block.name || `Preserved ${block.content.value.elementName}`,
          styleId: "Normal",
          text: block.content.value.text,
        };
      default:
        throw new OpenChestnutCodecError(`Document block ${block.id} has no supported wire content.`, [], { code: "invalid_document_artifact" });
    }
  });
  const document = DocumentModel.create({ name: source.name || "Imported document", styles, blocks });
  document.id = source.id || document.id;
  Object.defineProperty(document, DOCUMENT_STATE, {
    configurable: true,
    value: { source: envelope.source, opaqueOpc: envelope.opaqueOpc, diagnostics: envelope.diagnostics, blocks: source.blocks },
    writable: true,
  });
  return document;
}

export async function importDocxWithOpenChestnut(input, options = {}) {
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
  const response = await invokeOpenChestnut({
    protocolVersion: OPEN_CHESTNUT_PROTOCOL_VERSION,
    operation: CodecOperation.EXPORT_PPTX,
    family: ArtifactFamily.PRESENTATION,
    artifact: presentationEnvelope(presentation, OPEN_CHESTNUT_PROTOCOL_VERSION),
    limits: codecLimits(options.limits),
    allowLossy: options.allowLossy === true,
  });
  return new FileBlob(response.file, {
    type: PPTX_MIME,
    metadata: { artifactKind: "presentation", codec: "open-chestnut", diagnostics: response.diagnostics },
  });
}

export async function importPptxWithOpenChestnut(input, options = {}) {
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
