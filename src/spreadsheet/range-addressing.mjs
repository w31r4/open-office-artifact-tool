export const XLSX_MAX_ROW_INDEX = 1_048_575;
export const XLSX_MAX_COLUMN_INDEX = 16_383;

export function columnNumberToLabel(index) {
  let value = Number(index) + 1;
  if (!Number.isInteger(value) || value < 1) throw new Error(`Invalid column index: ${index}`);
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

export function columnLabelToNumber(label) {
  const text = String(label || "").toUpperCase();
  if (!/^[A-Z]+$/.test(text)) throw new Error(`Invalid column label: ${label}`);
  let value = 0;
  for (const character of text) value = value * 26 + (character.charCodeAt(0) - 64);
  return value - 1;
}

export function parseCellAddress(address) {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(String(address).trim());
  if (!match) throw new Error(`Invalid cell address: ${address}`);
  return { col: columnLabelToNumber(match[1]), row: Number(match[2]) - 1 };
}

export function makeCellAddress(row, col) {
  return `${columnNumberToLabel(col)}${Number(row) + 1}`;
}

export function rangeBounds(top, left, bottom, right) {
  return {
    top,
    left,
    bottom,
    right,
    rowCount: bottom - top + 1,
    colCount: right - left + 1,
  };
}

export function parseRangeAddress(address) {
  const raw = String(address || "A1").trim();
  const withoutSheet = raw.includes("!") ? raw.slice(raw.lastIndexOf("!") + 1) : raw;
  const [startRaw, endRaw = startRaw] = withoutSheet.split(":");
  const start = parseCellAddress(startRaw);
  const end = parseCellAddress(endRaw);
  return rangeBounds(
    Math.min(start.row, end.row),
    Math.min(start.col, end.col),
    Math.max(start.row, end.row),
    Math.max(start.col, end.col),
  );
}

export function rangeToAddress(bounds) {
  const start = makeCellAddress(bounds.top, bounds.left);
  const end = makeCellAddress(bounds.bottom, bounds.right);
  return start === end ? start : `${start}:${end}`;
}

export function assertWorksheetBounds(bounds, context = "Range") {
  const valid = Number.isInteger(bounds?.top) && Number.isInteger(bounds?.left)
    && Number.isInteger(bounds?.bottom) && Number.isInteger(bounds?.right)
    && bounds.top >= 0 && bounds.left >= 0
    && bounds.bottom >= bounds.top && bounds.right >= bounds.left
    && bounds.bottom <= XLSX_MAX_ROW_INDEX && bounds.right <= XLSX_MAX_COLUMN_INDEX;
  if (!valid) throw new Error(`${context} is outside the worksheet.`);
  return rangeBounds(bounds.top, bounds.left, bounds.bottom, bounds.right);
}

function protectFormulaSegments(formula, includeStructuredReferences = false) {
  const protectedParts = [];
  const pattern = includeStructuredReferences
    ? /"(?:[^"]|"")*"|(?!R(?:\d+)?C\[|R\[|C\[)[A-Za-z_][A-Za-z0-9_.]*\[(?:[^\]]|\]\])*\]/gi
    : /"(?:[^"]|"")*"|\[[^\]]*\]/g;
  const text = String(formula || "").replace(pattern, (part) => {
    const token = `\uE000${protectedParts.length}\uE001`;
    protectedParts.push(part);
    return token;
  });
  return {
    text,
    restore(value) {
      return String(value).replace(/\uE000(\d+)\uE001/g, (_, index) => protectedParts[Number(index)] || "");
    },
  };
}

function sheetPrefix(quotedSheet, bareSheet) {
  if (quotedSheet != null) return `'${quotedSheet}'!`;
  if (bareSheet != null) return `${bareSheet}!`;
  return "";
}

function transformA1References(formula, transform) {
  const protectedFormula = protectFormulaSegments(formula);
  const reference = /(^|[^A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_. ]*))!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_.(])/g;
  const transformed = protectedFormula.text.replace(
    reference,
    (match, leading, quotedSheet, bareSheet, absoluteColumn, columnText, absoluteRow, rowText) => {
      const column = columnLabelToNumber(columnText);
      const row = Number(rowText) - 1;
      if (column > XLSX_MAX_COLUMN_INDEX || row > XLSX_MAX_ROW_INDEX) return match;
      return `${leading}${sheetPrefix(quotedSheet, bareSheet)}${transform({
        row,
        column,
        absoluteRow: Boolean(absoluteRow),
        absoluteColumn: Boolean(absoluteColumn),
      })}`;
    },
  );
  return protectedFormula.restore(transformed);
}

export function translateA1Formula(formula, sourceAddress, targetAddress) {
  const source = typeof sourceAddress === "string" ? parseCellAddress(sourceAddress) : sourceAddress;
  const target = typeof targetAddress === "string" ? parseCellAddress(targetAddress) : targetAddress;
  return transformA1References(formula, ({ row, column, absoluteRow, absoluteColumn }) => {
    const shiftedRow = absoluteRow ? row : row + target.row - source.row;
    const shiftedColumn = absoluteColumn ? column : column + target.col - source.col;
    if (shiftedRow < 0 || shiftedColumn < 0 || shiftedRow > XLSX_MAX_ROW_INDEX || shiftedColumn > XLSX_MAX_COLUMN_INDEX) return "#REF!";
    return `${absoluteColumn ? "$" : ""}${columnNumberToLabel(shiftedColumn)}${absoluteRow ? "$" : ""}${shiftedRow + 1}`;
  });
}

export function formulaA1ToR1C1(formula, baseAddress) {
  const base = typeof baseAddress === "string" ? parseCellAddress(baseAddress) : baseAddress;
  return transformA1References(formula, ({ row, column, absoluteRow, absoluteColumn }) => {
    const rowToken = absoluteRow ? `R${row + 1}` : row === base.row ? "R" : `R[${row - base.row}]`;
    const columnToken = absoluteColumn ? `C${column + 1}` : column === base.col ? "C" : `C[${column - base.col}]`;
    return `${rowToken}${columnToken}`;
  });
}

function r1c1Coordinate(token, base, maximum, axis) {
  if (!token) return { value: base, absolute: false };
  if (token.startsWith("[")) {
    const offset = Number(token.slice(1, -1));
    const value = base + offset;
    if (!Number.isInteger(offset) || value < 0 || value > maximum) throw new Error(`R1C1 ${axis} reference resolves outside worksheet.`);
    return { value, absolute: false };
  }
  const value = Number(token) - 1;
  if (!Number.isInteger(value) || value < 0 || value > maximum) throw new Error(`R1C1 ${axis} reference resolves outside worksheet.`);
  return { value, absolute: true };
}

export function formulaR1C1ToA1(formula, baseAddress) {
  const base = typeof baseAddress === "string" ? parseCellAddress(baseAddress) : baseAddress;
  const protectedFormula = protectFormulaSegments(formula, true);
  const reference = /(^|[^A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_. ]*))!)?R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_.(])/gi;
  const transformed = protectedFormula.text.replace(reference, (match, leading, quotedSheet, bareSheet, rowToken, columnToken) => {
    const row = r1c1Coordinate(rowToken, base.row, XLSX_MAX_ROW_INDEX, "row");
    const column = r1c1Coordinate(columnToken, base.col, XLSX_MAX_COLUMN_INDEX, "column");
    return `${leading}${sheetPrefix(quotedSheet, bareSheet)}${column.absolute ? "$" : ""}${columnNumberToLabel(column.value)}${row.absolute ? "$" : ""}${row.value + 1}`;
  });
  return protectedFormula.restore(transformed);
}
