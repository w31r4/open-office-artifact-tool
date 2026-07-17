import { inspectOoxmlPackage, patchOoxmlPackage } from "../ooxml/package.mjs";
import { matchesFormulaCriteria } from "./formula-criteria.mjs";
import { normalizeSpreadsheetChartSeriesLine } from "./chart-line-style.mjs";
import { normalizeSpreadsheetChartLineOptions } from "./chart-line-options.mjs";
import { normalizeSpreadsheetChartSeriesMarker } from "./chart-marker-style.mjs";
import { normalizeSpreadsheetChartDataLabels } from "./chart-data-labels.mjs";
import { resolvedWorksheetChartCategories, resolvedWorksheetChartSeriesBubbleSizes, resolvedWorksheetChartSeriesValues, resolvedWorksheetChartSeriesXValues } from "./chart-source-data.mjs";
import { renderWorksheetChartSvg } from "./chart-preview.mjs";
import { WorksheetDataTableCollection } from "./data-tables.mjs";
import { computePivotValues, normalizePivotConfig } from "./pivots.mjs";
import { formatSpreadsheetDisplayValue, normalizeXlsxColor, normalizeXlsxThemeConfig, xlsxColorCss, xlsxFillSvgPaint } from "./ooxml-styles.mjs";
import { planSpreadsheetThreadedComments, validateSpreadsheetThreadedCommentPackageSemantics } from "./ooxml-threaded-comments.mjs";
import { parseStructuredReference, scanStructuredReferenceIntersections, scanStructuredReferences, splitReferenceIntersectionOperands } from "./structured-references.mjs";
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
} from "./range-addressing.mjs";
import {
  assertRangeCopyShape,
  currentRegionBounds,
  normalizeRangeCopyMode,
  normalizeRangeWrite,
  writtenRangeBounds,
} from "./range-operations.mjs";
import { createSpreadsheetSparklineClasses } from "./sparklines.mjs";
import { formulaTimeParts, formulaTimeSerial, parseFormulaDateText, parseFormulaNumberText, parseFormulaTimeText } from "./formula-coercion.mjs";
import { calculateDb, calculateDdb, calculateIpmt, calculateIrr, calculateNpv, calculatePmt, calculatePpmt, calculateRate, calculateSln, calculateXirr, calculateXnpv } from "./financial-formulas.mjs";
import { createWorkbookWindowCollection, worksheetWindowMemberships } from "./workbook-windows.mjs";
import { decoder, encoder, toUint8Array } from "../shared/binary.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { officeFontFamilies } from "../shared/font-design-metrics.mjs";
import { aid } from "../shared/ids.mjs";
import { imageDataFromDataUrl } from "../shared/images.mjs";
import { filterInspectRecords, inspectRecordMatchesTarget, inspectTargetTokens, ndjson, normalizeKinds, verificationIssue, verificationResult } from "../shared/inspection.mjs";
import { fileBlobFromRenderOutput, LAYOUT_MIME, renderTypeForOptions } from "../shared/render-output.mjs";
import { attrEscape, xmlEscape } from "../shared/xml.mjs";
import { queryHelpRecords } from "../help/index.mjs";

export { WorksheetDataTableCollection } from "./data-tables.mjs";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const XLSX_DYNAMIC_ARRAY_METADATA_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
const XLSX_DYNAMIC_ARRAY_METADATA_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
const XLSX_DYNAMIC_ARRAY_METADATA_EXTENSION_URI = "{BDBB8CDC-FA1E-496E-A857-3C3F30C029C3}";
const CSV_MIME = "text/csv";
const TSV_MIME = "text/tab-separated-values";
const XLSX_PACKAGE_CONFIG = {
  family: "XLSX",
  packageKind: "xlsxPackage",
  partKind: "xlsxPart",
  counts: { sheets: /^xl\/worksheets\/sheet\d+\.xml$/ },
  semanticIssues: validateSpreadsheetThreadedCommentPackageSemantics,
};

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

const WORKSHEET_NUMERIC_X_CHART_TYPES = new Set(["scatter", "bubble"]);

class WorksheetChartSeriesCollection {
  constructor(chart) { this.chart = chart; this.items = []; }
  add(name, values = []) { const series = { name, values, xValues: undefined, bubbleSizes: undefined, categoryFormula: undefined, xFormula: undefined, formula: undefined, bubbleSizeFormula: undefined, fill: undefined }; this.items.push(series); return series; }
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

function normalizeWorksheetChartAxis(value, kind, chartType) {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError(`Worksheet chart ${kind}Axis must be an object.`);
  const title = typeof value.title === "string" ? value.title : value.title?.text;
  return {
    axisType: value.axisType || (kind === "x" && !WORKSHEET_NUMERIC_X_CHART_TYPES.has(chartType) ? "textAxis" : "valueAxis"),
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
    const circular = chartType === "pie" || chartType === "doughnut";
    if (!circular || sourceOrConfig.xAxis != null) this.xAxis = sourceOrConfig.xAxis || {};
    if (!circular || sourceOrConfig.yAxis != null) this.yAxis = sourceOrConfig.yAxis || {};
    this.series = new WorksheetChartSeriesCollection(this);
    if (sourceOrConfig.series) sourceOrConfig.series.forEach((series) => Object.assign(this.series.add(series.name, series.values || []), {
      xValues: series.xValues == null ? undefined : [...series.xValues],
      bubbleSizes: series.bubbleSizes == null ? undefined : [...series.bubbleSizes],
      categoryFormula: series.categoryFormula,
      xFormula: series.xFormula,
      formula: series.formula,
      bubbleSizeFormula: series.bubbleSizeFormula,
      fill: series.fill,
      ...(series.line == null ? {} : { line: series.line }),
      ...(series.stroke == null ? {} : { stroke: series.stroke }),
      ...(series.marker == null ? {} : { marker: series.marker }),
    }));
    if (sourceOrConfig instanceof Range) this.setData(sourceOrConfig);
    else if (sourceOrConfig && sourceOrConfig.worksheet instanceof Worksheet) this.setData(sourceOrConfig);
  }

  get xAxis() { return this._xAxis; }
  set xAxis(value) { this._xAxis = normalizeWorksheetChartAxis(value, "x", this.type); }
  get yAxis() { return this._yAxis; }
  set yAxis(value) { this._yAxis = normalizeWorksheetChartAxis(value, "y", this.type); }

  setData(range) {
    const values = range.values;
    if (!values.length || !values[0]?.length) return this;
    const header = values[0];
    const dataRows = values.slice(1);
    if (this.type === "bubble") {
      if (header.length !== 3) throw new TypeError(`Worksheet bubble chart range shortcut requires exactly three columns ordered X | Y | Size; received ${header.length}. Use explicit series configuration for multiple series.`);
      const numericColumn = (column, label, positive = false) => dataRows.map((row, index) => {
        const raw = row[column];
        const number = Number(raw);
        if (raw == null || raw === "" || !Number.isFinite(number) || positive && number <= 0) {
          throw new TypeError(`Worksheet bubble chart ${label} value at data row ${index + 1} must be ${positive ? "finite and positive" : "finite"}.`);
        }
        return number;
      });
      const xValues = numericColumn(0, "X");
      const yValues = numericColumn(1, "Y");
      const bubbleSizes = numericColumn(2, "Size", true);
      const formulaFor = (column) => `'${range.worksheet.name}'!${makeCellAddress(range.bounds.top + 1, range.bounds.left + column)}:${makeCellAddress(range.bounds.bottom, range.bounds.left + column)}`;
      this.categories = [];
      this.series.items = [];
      const series = this.series.add(String(header[1] || "Series 1"), yValues);
      series.xValues = xValues;
      series.bubbleSizes = bubbleSizes;
      series.xFormula = formulaFor(0);
      series.formula = formulaFor(1);
      series.bubbleSizeFormula = formulaFor(2);
      return this;
    }
    const scatter = this.type === "scatter";
    const xValues = scatter ? dataRows.map((row) => Number(row[0])) : undefined;
    this.categories = scatter ? [] : dataRows.map((row) => String(row[0] ?? ""));
    this.series.items = [];
    for (let column = 1; column < header.length; column++) {
      const series = this.series.add(String(header[column] || `Series ${column}`), dataRows.map((row) => Number(row[column]) || 0));
      const start = makeCellAddress(range.bounds.top + 1, range.bounds.left + column);
      const end = makeCellAddress(range.bounds.bottom, range.bounds.left + column);
      series.formula = `'${range.worksheet.name}'!${start}:${end}`;
      const horizontalFormula = `'${range.worksheet.name}'!${makeCellAddress(range.bounds.top + 1, range.bounds.left)}:${makeCellAddress(range.bounds.bottom, range.bounds.left)}`;
      if (scatter) {
        series.xValues = [...xValues];
        series.xFormula = horizontalFormula;
      } else series.categoryFormula = horizontalFormula;
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
      ...(WORKSHEET_NUMERIC_X_CHART_TYPES.has(this.type) ? { xValues: resolvedWorksheetChartSeriesXValues(this, this.series.items[index]) } : {}),
      ...(this.type === "bubble" ? { bubbleSizes: resolvedWorksheetChartSeriesBubbleSizes(this, this.series.items[index]) } : {}),
      values: resolvedWorksheetChartSeriesValues(this, this.series.items[index]),
    }));
    return { kind: "drawing", drawingType: "chart", id: this.id, sheet: this.worksheet.name, name: this.name, chartType: this.type, title: this.title, titleTextStyle: this.titleTextStyle, lineOptions: this.lineOptions, dataLabels: normalizeSpreadsheetChartDataLabels(this.dataLabels), categories, series: this.series.items.length, seriesItems, xAxis: this.xAxis, yAxis: this.yAxis, bbox: [this.position.left, this.position.top, this.position.width, this.position.height], bboxUnit: "px" };
  }

  toSvg() { return renderWorksheetChartSvg(this); }

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

const { RangeSparklineFacade, SparklineGroupCollection } = createSpreadsheetSparklineClasses({ workbookRangeRef, worksheetRangeFrame });

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

  get fontFamilies() {
    const cells = this.worksheets.items.flatMap((sheet) => sheet.store.entries().map(([, cell]) => cell));
    const conditionalFormats = this.worksheets.items.flatMap((sheet) => sheet.conditionalFormattings.items.map((rule) => rule.format));
    return officeFontFamilies([cells, conditionalFormats], ["Aptos"]);
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
      if (kinds.has("dataTable")) records.push(...sheet.dataTables.inspectRecords());
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
          if (WORKSHEET_NUMERIC_X_CHART_TYPES.has(chart.type)) {
            if (chart.categories.length) issues.push(verificationIssue("workbook", "chartDataMismatch", `${chart.type === "bubble" ? "Bubble" : "Scatter"} chart ${chart.name} must use per-series numeric xValues instead of shared categories.`, { sheet: sheet.name, id: chart.id, categories: chart.categories.length }));
            if ((series.xValues?.length || 0) !== series.values.length) issues.push(verificationIssue("workbook", "chartDataMismatch", `${chart.type === "bubble" ? "Bubble" : "Scatter"} chart ${chart.name} series ${series.name || "Series"} has ${series.xValues?.length || 0} x values for ${series.values.length} y values.`, { sheet: sheet.name, id: chart.id, series: series.name, xValues: series.xValues?.length || 0, yValues: series.values.length }));
            if (series.categoryFormula) issues.push(verificationIssue("workbook", "chartDataMismatch", `${chart.type === "bubble" ? "Bubble" : "Scatter"} chart ${chart.name} series ${series.name || "Series"} must use xFormula rather than categoryFormula.`, { sheet: sheet.name, id: chart.id, series: series.name, categoryFormula: series.categoryFormula }));
            if (chart.type === "bubble") {
              if ((series.bubbleSizes?.length || 0) !== series.values.length) issues.push(verificationIssue("workbook", "chartDataMismatch", `Bubble chart ${chart.name} series ${series.name || "Series"} has ${series.bubbleSizes?.length || 0} sizes for ${series.values.length} y values.`, { sheet: sheet.name, id: chart.id, series: series.name, bubbleSizes: series.bubbleSizes?.length || 0, yValues: series.values.length }));
              if ((series.bubbleSizes || []).some((value) => !Number.isFinite(Number(value)) || Number(value) <= 0)) issues.push(verificationIssue("workbook", "chartDataMismatch", `Bubble chart ${chart.name} series ${series.name || "Series"} requires finite positive sizes.`, { sheet: sheet.name, id: chart.id, series: series.name, bubbleSizes: series.bubbleSizes }));
            } else if ((series.bubbleSizes?.length || 0) > 0 || series.bubbleSizeFormula) issues.push(verificationIssue("workbook", "chartDataMismatch", `Scatter chart ${chart.name} series ${series.name || "Series"} cannot carry bubble sizes.`, { sheet: sheet.name, id: chart.id, series: series.name }));
          } else if (chart.categories.length && series.values.length && chart.categories.length !== series.values.length) issues.push(verificationIssue("workbook", "chartDataMismatch", `Chart ${chart.name} series ${series.name || "Series"} has ${series.values.length} values for ${chart.categories.length} categories.`, { sheet: sheet.name, id: chart.id, series: series.name, values: series.values.length, categories: chart.categories.length }));
          if (series.formula && !workbookRangeValid(this, sheet, series.formula)) issues.push(verificationIssue("workbook", "chartFormulaInvalid", `Chart ${chart.name} series ${series.name || "Series"} references an invalid range.`, { sheet: sheet.name, id: chart.id, formula: series.formula }));
          if (series.xFormula && !workbookRangeValid(this, sheet, series.xFormula)) issues.push(verificationIssue("workbook", "chartXFormulaInvalid", `Numeric-X chart ${chart.name} series ${series.name || "Series"} references an invalid x-value range.`, { sheet: sheet.name, id: chart.id, formula: series.xFormula }));
          if (series.bubbleSizeFormula && !workbookRangeValid(this, sheet, series.bubbleSizeFormula)) issues.push(verificationIssue("workbook", "chartBubbleSizeFormulaInvalid", `Bubble chart ${chart.name} series ${series.name || "Series"} references an invalid size range.`, { sheet: sheet.name, id: chart.id, formula: series.bubbleSizeFormula }));
          if (series.categoryFormula && !workbookRangeValid(this, sheet, series.categoryFormula)) issues.push(verificationIssue("workbook", "chartCategoryFormulaInvalid", `Chart ${chart.name} categories reference an invalid range.`, { sheet: sheet.name, id: chart.id, formula: series.categoryFormula }));
          if (series.fill != null && (typeof series.fill !== "string" || !/^#[0-9a-f]{6}$/i.test(series.fill))) issues.push(verificationIssue("workbook", "invalidChartSeriesFill", `Chart ${chart.name} series ${series.name || "Series"} fill must be a #RRGGBB solid color.`, { sheet: sheet.name, id: chart.id, series: series.name, fill: series.fill }));
          try {
            const line = normalizeSpreadsheetChartSeriesLine(series);
            if (chart.type === "scatter" && line != null) issues.push(verificationIssue("workbook", "unsupportedChartSeriesLine", `Marker-only scatter chart ${chart.name} series ${series.name || "Series"} cannot carry line/stroke; use marker.line to style the marker border.`, { sheet: sheet.name, id: chart.id, series: series.name, line }));
          } catch (error) { issues.push(verificationIssue("workbook", "invalidChartSeriesLine", String(error?.message || error), { sheet: sheet.name, id: chart.id, series: series.name, line: series.line, stroke: series.stroke })); }
          try {
            const marker = normalizeSpreadsheetChartSeriesMarker(series.marker);
            if (marker != null && chart.type !== "line" && chart.type !== "scatter") issues.push(verificationIssue("workbook", "invalidChartSeriesMarker", `Chart ${chart.name} series markers require a line or scatter chart.`, { sheet: sheet.name, id: chart.id, series: series.name, marker }));
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
    return ndjson(queryHelpRecords("workbook", query, options), options.maxChars ?? Infinity);
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
    this.dataTables = new WorksheetDataTableCollection(this);
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

const FORMULA_NUMERIC_LITERAL = /^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

function formulaScalar(sheet, expr, context = {}) {
  const text = String(expr ?? "").trim();
  const quoted = formulaUnquote(text);
  if (quoted !== undefined) return quoted;
  const error = formulaErrorCode(text);
  if (error) return error;
  if (FORMULA_NUMERIC_LITERAL.test(text)) return Number(text);
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

function formulaArithmeticNumberLiteral(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return "0";
  const text = String(number);
  if (!/[eE]/.test(text)) return text;
  const negative = text.startsWith("-");
  const [mantissa, exponentText] = (negative ? text.slice(1) : text).toLowerCase().split("e");
  const exponent = Number(exponentText);
  const digits = mantissa.replace(".", "");
  const decimalIndex = 1 + exponent;
  let decimal;
  if (decimalIndex <= 0) decimal = `0.${"0".repeat(-decimalIndex)}${digits}`;
  else if (decimalIndex >= digits.length) decimal = `${digits}${"0".repeat(decimalIndex - digits.length)}`;
  else decimal = `${digits.slice(0, decimalIndex)}.${digits.slice(decimalIndex)}`;
  return negative ? `-${decimal}` : decimal;
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
  const financialHelpers = {
    errorCode: formulaErrorCode,
    dateNumber: (value) => {
      if (!(value instanceof Date)) return excelFormulaDateNumber(value);
      if (!Number.isFinite(value.getTime())) return "#VALUE!";
      return excelGregorianSerial(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate(), dateSystem);
    },
    isValidDate: (serial) => Boolean(excelDateParts(serial, dateSystem)),
  };
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
    case "SLN": return args.length === 3
      ? calculateSln({ cost: scalar(0), salvage: scalar(1), life: scalar(2) }, financialHelpers)
      : "#VALUE!";
    case "DB": return args.length >= 4 && args.length <= 5
      ? calculateDb({ cost: scalar(0), salvage: scalar(1), life: scalar(2), period: scalar(3), month: scalar(4, 12) }, financialHelpers)
      : "#VALUE!";
    case "DDB": return args.length >= 4 && args.length <= 5
      ? calculateDdb({ cost: scalar(0), salvage: scalar(1), life: scalar(2), period: scalar(3), factor: scalar(4, 2) }, financialHelpers)
      : "#VALUE!";
    case "PMT": return args.length >= 3 && args.length <= 5
      ? calculatePmt({ rate: scalar(0), nper: scalar(1), pv: scalar(2), fv: scalar(3, 0), type: scalar(4, 0) }, financialHelpers)
      : "#VALUE!";
    case "IPMT": return args.length >= 4 && args.length <= 6
      ? calculateIpmt({ rate: scalar(0), per: scalar(1), nper: scalar(2), pv: scalar(3), fv: scalar(4, 0), type: scalar(5, 0) }, financialHelpers)
      : "#VALUE!";
    case "PPMT": return args.length >= 4 && args.length <= 6
      ? calculatePpmt({ rate: scalar(0), per: scalar(1), nper: scalar(2), pv: scalar(3), fv: scalar(4, 0), type: scalar(5, 0) }, financialHelpers)
      : "#VALUE!";
    case "RATE": return args.length >= 3 && args.length <= 6
      ? calculateRate({ nper: scalar(0), pmt: scalar(1), pv: scalar(2), fv: scalar(3, 0), type: scalar(4, 0), guess: args.length === 6 ? scalar(5) : undefined }, financialHelpers)
      : "#VALUE!";
    case "NPV": return args.length >= 2
      ? calculateNpv({ rate: scalar(0), cashFlows: values(args.slice(1)) }, financialHelpers)
      : "#VALUE!";
    case "XNPV": return args.length === 3
      ? calculateXnpv({ rate: scalar(0), cashFlows: values([args[1]]), dates: values([args[2]]) }, financialHelpers)
      : "#VALUE!";
    case "IRR": return args.length >= 1 && args.length <= 2
      ? calculateIrr({ cashFlows: values([args[0]]), guess: args.length === 2 ? scalar(1) : undefined }, financialHelpers)
      : "#VALUE!";
    case "XIRR": return args.length >= 2 && args.length <= 3
      ? calculateXirr({ cashFlows: values([args[0]]), dates: values([args[1]]), guess: args.length === 3 ? scalar(2) : undefined }, financialHelpers)
      : "#VALUE!";
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
  if (FORMULA_NUMERIC_LITERAL.test(expr)) return Number(expr);
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
    // Parenthesize numeric replacements so a negative referenced value cannot
    // accidentally combine with a binary minus as JavaScript's `--` token.
    // Scientific notation is expanded because the constrained evaluator only
    // accepts decimal arithmetic literals after reference substitution.
    return `(${formulaArithmeticNumberLiteral(value)})`;
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
      dataTables: sheet.dataTables.toJSON(),
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
    sheet.dataTables._definitions = [];
    for (const dataTable of sheetData.dataTables || []) sheet.dataTables._hydrate(dataTable);
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
    const { exportXlsxWithOpenChestnut } = await import("../codecs/open-chestnut.mjs");
    return exportXlsxWithOpenChestnut(workbook, options);
  }

  static async importXlsx(blobOrBuffer, options = {}) {
    const { importXlsxWithOpenChestnut } = await import("../codecs/open-chestnut.mjs");
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
