import {
  assertWorksheetBounds,
  rangeBounds,
  XLSX_MAX_COLUMN_INDEX,
  XLSX_MAX_ROW_INDEX,
} from "./range-addressing.mjs";

export const RANGE_WRITE_FIELDS = Object.freeze(["values", "formulas", "formulasR1C1"]);
export const RANGE_COPY_MODES = Object.freeze(["values", "formulas", "all"]);

export function normalizeRangeMatrix(value) {
  if (!Array.isArray(value)) return [[value]];
  if (value.length === 0) throw new Error("Range write requires at least one value.");
  const rows = value.every((item) => Array.isArray(item)) ? value : [value];
  const width = Math.max(0, ...rows.map((row) => row.length));
  if (width === 0) throw new Error("Range write requires at least one column.");
  return rows.map((row) => Array.from({ length: width }, (_, index) => row[index] ?? null));
}

export function normalizeRangeWrite(value, forcedField) {
  if (forcedField) return { field: forcedField, matrix: normalizeRangeMatrix(value) };
  if (value && typeof value === "object" && !Array.isArray(value)) {
    const fields = RANGE_WRITE_FIELDS.filter((field) => Object.prototype.hasOwnProperty.call(value, field));
    if (fields.length !== 1) {
      const listed = fields.length ? fields.join(", ") : "none";
      throw new Error(`Range.write(payload) expects exactly one field; got ${listed}.`);
    }
    return { field: fields[0], matrix: normalizeRangeMatrix(value[fields[0]]) };
  }
  return { field: "mixed", matrix: normalizeRangeMatrix(value) };
}

export function writtenRangeBounds(anchor, matrix) {
  return assertWorksheetBounds(rangeBounds(
    anchor.top,
    anchor.left,
    anchor.top + matrix.length - 1,
    anchor.left + matrix[0].length - 1,
  ), "Range.write result");
}

export function normalizeRangeCopyMode(value = "all") {
  const mode = String(value || "all");
  if (!RANGE_COPY_MODES.includes(mode)) throw new Error(`Range copy mode must be one of: ${RANGE_COPY_MODES.join(", ")}.`);
  return mode;
}

export function assertRangeCopyShape(source, destination) {
  const tiles = destination.rowCount % source.rowCount === 0 && destination.colCount % source.colCount === 0;
  if (!tiles) {
    throw new Error(
      `Range.copyFrom requires source and destination to match shape (or source to evenly tile the destination). `
      + `Source is ${source.rowCount}x${source.colCount}, destination is ${destination.rowCount}x${destination.colCount}.`,
    );
  }
}

export function currentRegionBounds(initial, hasContent) {
  let current = rangeBounds(initial.top, initial.left, initial.bottom, initial.right);
  let changed = true;
  const rowHasContent = (row) => {
    for (let column = current.left; column <= current.right; column += 1) if (hasContent(row, column)) return true;
    return false;
  };
  const columnHasContent = (column) => {
    for (let row = current.top; row <= current.bottom; row += 1) if (hasContent(row, column)) return true;
    return false;
  };
  while (changed) {
    changed = false;
    if (current.top > 0 && rowHasContent(current.top - 1)) { current = rangeBounds(current.top - 1, current.left, current.bottom, current.right); changed = true; }
    if (current.bottom < XLSX_MAX_ROW_INDEX && rowHasContent(current.bottom + 1)) { current = rangeBounds(current.top, current.left, current.bottom + 1, current.right); changed = true; }
    if (current.left > 0 && columnHasContent(current.left - 1)) { current = rangeBounds(current.top, current.left - 1, current.bottom, current.right); changed = true; }
    if (current.right < XLSX_MAX_COLUMN_INDEX && columnHasContent(current.right + 1)) { current = rangeBounds(current.top, current.left, current.bottom, current.right + 1); changed = true; }
  }
  return current;
}
