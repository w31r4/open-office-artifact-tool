import {
  XLSX_MAX_COLUMN_INDEX,
  XLSX_MAX_ROW_INDEX,
  assertWorksheetBounds,
  makeCellAddress,
  parseCellAddress,
  parseRangeAddress,
  rangeToAddress,
} from "./range-addressing.mjs";

function normalizedLocalCell(worksheet, value, label) {
  const raw = String(value ?? "").trim();
  if (!raw) return undefined;
  if (raw.includes(":")) throw new Error(`Data table ${label} must be a single cell address, not a range: ${raw}`);
  let address = raw;
  if (raw.includes("!")) {
    const separator = raw.lastIndexOf("!");
    const sheetToken = raw.slice(0, separator).trim();
    const sheetName = sheetToken.startsWith("'") && sheetToken.endsWith("'")
      ? sheetToken.slice(1, -1).replace(/''/g, "'")
      : sheetToken;
    if (sheetName !== worksheet.name) throw new Error(`Data table ${label} must be on sheet "${worksheet.name}", got "${sheetName}"`);
    address = raw.slice(separator + 1);
  }
  const coordinate = parseCellAddress(address);
  if (coordinate.row < 0 || coordinate.row > XLSX_MAX_ROW_INDEX || coordinate.col < 0 || coordinate.col > XLSX_MAX_COLUMN_INDEX) {
    throw new Error(`Data table ${label} is outside the worksheet: ${raw}`);
  }
  return makeCellAddress(coordinate.row, coordinate.col);
}

function cloneDefinition(definition) {
  return {
    range: { ...definition.range },
    formulaRef: definition.formulaRef,
    anchor: { ...definition.anchor },
    ...(definition.rowInput ? { rowInput: definition.rowInput } : {}),
    ...(definition.columnInput ? { columnInput: definition.columnInput } : {}),
    rowOriented: definition.rowOriented,
    twoVariable: definition.twoVariable,
    displayFormula: definition.displayFormula,
  };
}

function overlaps(left, right) {
  return left.startRow <= right.endRow && right.startRow <= left.endRow
    && left.startCol <= right.endCol && right.startCol <= left.endCol;
}

function definitionFromResultRange(worksheet, bounds, options = {}) {
  const normalized = assertWorksheetBounds(bounds, "Data table result range");
  const rowInput = normalizedLocalCell(worksheet, options.rowInput, "rowInput");
  const columnInput = normalizedLocalCell(worksheet, options.columnInput, "columnInput");
  if (!rowInput && !columnInput) throw new Error("Data tables require at least a rowInput or columnInput address");
  const twoVariable = Boolean(rowInput && columnInput);
  const rowOriented = Boolean(rowInput && !columnInput);
  const formulaRef = rangeToAddress(normalized);
  const displayInputs = [rowInput, columnInput].filter(Boolean).join(",");
  return {
    range: {
      startRow: normalized.top,
      startCol: normalized.left,
      endRow: normalized.bottom,
      endCol: normalized.right,
    },
    formulaRef,
    anchor: { row: normalized.top, col: normalized.left, address: makeCellAddress(normalized.top, normalized.left) },
    ...(rowInput ? { rowInput } : {}),
    ...(columnInput ? { columnInput } : {}),
    rowOriented,
    twoVariable,
    displayFormula: `{=TABLE(${displayInputs})}`,
  };
}

export class WorksheetDataTableCollection {
  constructor(worksheet) {
    this.worksheet = worksheet;
    this._definitions = [];
  }

  add(rangeOrAddress, options = {}) {
    const range = typeof rangeOrAddress === "string" ? this.worksheet.getRange(rangeOrAddress) : rangeOrAddress;
    if (!range || range.worksheet !== this.worksheet || !range.bounds) throw new TypeError("Data table range must belong to its worksheet.");
    const bounds = assertWorksheetBounds(range.bounds, "Data table range");
    if (bounds.rowCount < 2 || bounds.colCount < 2) {
      throw new Error(`Data table range must include the formula cell plus at least one row and one column of results: ${rangeToAddress(bounds)}`);
    }
    const formulaAddress = makeCellAddress(bounds.top, bounds.left);
    if (!this.worksheet.store.get(formulaAddress).formula) {
      throw new Error(`Data table range needs a top-left formula cell. "${formulaAddress}" is empty.`);
    }
    const definition = definitionFromResultRange(this.worksheet, {
      top: bounds.top + 1,
      left: bounds.left + 1,
      bottom: bounds.bottom,
      right: bounds.right,
    }, options);
    this._append(definition);
    return undefined;
  }

  __getDefinitions() {
    return this._definitions.map(cloneDefinition);
  }

  inspectRecords() {
    return this._definitions.map((definition) => ({
      kind: "dataTable",
      sheet: this.worksheet.name,
      range: definition.formulaRef,
      anchor: definition.anchor.address,
      rowInput: definition.rowInput,
      columnInput: definition.columnInput,
      rowOriented: definition.rowOriented,
      twoVariable: definition.twoVariable,
      displayFormula: definition.displayFormula,
    }));
  }

  toJSON() {
    return this.__getDefinitions();
  }

  _hydrate(wire) {
    const bounds = parseRangeAddress(wire.reference || wire.formulaRef);
    const definition = definitionFromResultRange(this.worksheet, bounds, {
      rowInput: wire.rowInput,
      columnInput: wire.columnInput,
    });
    if (definition.rowOriented !== Boolean(wire.rowOriented) || definition.twoVariable !== Boolean(wire.twoVariable)) {
      throw new Error(`Imported data table ${wire.reference || wire.formulaRef} has inconsistent orientation metadata.`);
    }
    this._append(definition);
    return definition;
  }

  _append(definition) {
    const conflict = this._definitions.find((candidate) => overlaps(candidate.range, definition.range));
    if (conflict) throw new Error(`Data table result range ${definition.formulaRef} overlaps ${conflict.formulaRef}.`);
    this._definitions.push(definition);
  }
}
