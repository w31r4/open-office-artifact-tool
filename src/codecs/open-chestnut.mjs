import { create, fromBinary, toBinary } from "@bufbuild/protobuf";
import { readFile } from "node:fs/promises";
import { DocumentModel, FileBlob, Workbook } from "../index.mjs";
import {
  ArtifactFamily,
  CodecOperation,
  CodecRequestSchema,
  CodecResponseSchema,
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
const DOCUMENT_STATE = Symbol.for("open-office-artifact-tool.open-chestnut-document-state");
const EXCEL_ERRORS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#UNKNOWN!", "#CONNECT!", "#CYCLE!"]);
const DEFAULT_THEME_COLORS = {
  dk1: "#000000", lt1: "#FFFFFF", dk2: "#1F497D", lt2: "#EEECE1",
  accent1: "#4F81BD", accent2: "#C0504D", accent3: "#9BBB59", accent4: "#8064A2",
  accent5: "#4BACC6", accent6: "#F79646", hlink: "#0000FF", folHlink: "#800080",
};

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
      if (block.kind !== "paragraph" || block.text !== original.content.value.text || block.styleId !== (original.styleId || "Normal")) return false;
      if (original.source?.editable !== false) return false;
      return block.runs.every((run) => Object.keys(run.style || {}).length === 0);
    }
    case "table": {
      if (block.kind !== "table" || !sameTableValues(block, original)) return false;
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

function documentBlock(block, original) {
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
  if (block.kind === "table") {
    return {
      ...common,
      content: {
        case: "table",
        value: { rows: (block.values || []).map((cells) => ({ cells: cells.map((value) => String(value ?? "")) })) },
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
        blocks: document.blocks.map((block, index) => documentBlock(block, state?.blocks[index])),
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
        return {
          kind: "table",
          id: block.id,
          name: block.name,
          styleId: block.styleId || "TableGrid",
          values: block.content.value.rows.map((row) => [...row.cells]),
        };
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
