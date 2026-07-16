import { aid } from "../shared/ids.mjs";
import { attrEscape } from "../shared/xml.mjs";
import { makeCellAddress, parseRangeAddress, rangeToAddress } from "./range-addressing.mjs";
import { xlsxColorCss } from "./ooxml-styles.mjs";

const DEFAULT_AXIS = Object.freeze({ minMode: 0, maxMode: 0, showAxis: true, rightToLeft: false });
const DEFAULT_MARKERS = Object.freeze({ show: false, high: false, low: false, first: false, last: false, negative: false });

function sourceWorksheet(group) {
  return group.sourceData.sheetName
    ? group.worksheet.workbook.worksheets.getItem(group.sourceData.sheetName)
    : group.worksheet;
}

function mapping(target, source) {
  return {
    targetAddress: makeCellAddress(target.row, target.col),
    sourceAddress: rangeToAddress(source),
  };
}

// A native sparkline group is a list of formula/target-cell pairs. The public
// facade uses two ranges, so only the reversible row- and column-oriented
// rectangular profiles are projected here and in OpenChestnut.
export function sparklineRangeMappings(group) {
  const target = parseRangeAddress(group.targetRange.address);
  const source = parseRangeAddress(group.sourceData.address);
  if (target.rowCount > 1 && target.colCount > 1) return [];
  if (target.rowCount === 1 && target.colCount === 1) {
    if (source.rowCount > 1 && source.colCount > 1) return [];
    return [mapping({ row: target.top, col: target.left }, source)];
  }
  if (target.rowCount > 1) {
    if (source.rowCount !== target.rowCount) return [];
    return Array.from({ length: target.rowCount }, (_value, index) => mapping(
      { row: target.top + index, col: target.left },
      { top: source.top + index, bottom: source.top + index, left: source.left, right: source.right },
    ));
  }
  if (source.colCount !== target.colCount) return [];
  return Array.from({ length: target.colCount }, (_value, index) => mapping(
    { row: target.top, col: target.left + index },
    { top: source.top, bottom: source.bottom, left: source.left + index, right: source.left + index },
  ));
}

function color(group, value, fallback) {
  return xlsxColorCss(value ?? fallback, { theme: group.worksheet.workbook.theme, fallback });
}

function numericValues(group, sourceAddress) {
  const sheet = sourceWorksheet(group);
  if (!sheet) return [];
  return sheet.getRange(sourceAddress).values.flat().map((value) => Number(value)).filter(Number.isFinite);
}

function markerCircles(group, points, values) {
  if (!group.markers.show && !group.markers.high && !group.markers.low && !group.markers.first && !group.markers.last && !group.markers.negative) return "";
  const high = Math.max(...values), low = Math.min(...values);
  return points.map(([x, y], index) => {
    const value = values[index];
    const selected = group.markers.show || group.markers.high && value === high || group.markers.low && value === low ||
      group.markers.first && index === 0 || group.markers.last && index === values.length - 1 || group.markers.negative && value < 0;
    if (!selected) return "";
    const fill = value < 0 && group.markers.negative ? color(group, group.negativeColor, "#DC2626")
      : index === 0 && group.markers.first ? color(group, group.firstMarkerColor, group.seriesColor)
      : index === values.length - 1 && group.markers.last ? color(group, group.lastMarkerColor, group.seriesColor)
      : value === high && group.markers.high ? color(group, group.highMarkerColor, group.seriesColor)
      : value === low && group.markers.low ? color(group, group.lowMarkerColor, group.seriesColor)
      : color(group, group.markersColor, group.seriesColor);
    return `<circle cx="${x}" cy="${y}" r="2.2" fill="${attrEscape(fill)}"/>`;
  }).join("");
}

function renderSparkline(group, frame, values) {
  const left = frame.left + 3, top = frame.top + 3, width = Math.max(1, frame.width - 6), height = Math.max(1, frame.height - 6);
  if (!values.length) return `<rect x="${left}" y="${top}" width="${width}" height="${height}" fill="none" stroke="#38BDF8" stroke-dasharray="3 2"/>`;
  const minimum = group.axis.manualMin == null ? Math.min(...values, 0) : Number(group.axis.manualMin);
  const maximum = group.axis.manualMax == null ? Math.max(...values, 0) : Number(group.axis.manualMax);
  const span = Math.max(1e-12, maximum - minimum);
  const baseline = top + height - ((0 - minimum) / span) * height;
  const axis = group.axis.showAxis && minimum < 0 && maximum > 0
    ? `<line x1="${left}" y1="${baseline}" x2="${left + width}" y2="${baseline}" stroke="${attrEscape(color(group, group.axisColor, "#64748B"))}" stroke-width="0.75"/>`
    : "";
  const series = color(group, group.seriesColor, "#0EA5E9");
  const negative = color(group, group.negativeColor, "#DC2626");
  if (group.type === "column" || group.type === "stacked") {
    const slot = width / values.length, barWidth = Math.max(1, slot * 0.62);
    const bars = values.map((value, index) => {
      const x = left + index * slot + (slot - barWidth) / 2;
      if (group.type === "stacked") {
        const positive = value >= 0, y = positive ? top : top + height / 2;
        return `<rect x="${x}" y="${y}" width="${barWidth}" height="${height / 2}" fill="${attrEscape(positive ? series : negative)}"/>`;
      }
      const y = top + height - ((value - minimum) / span) * height;
      return `<rect x="${x}" y="${Math.min(y, baseline)}" width="${barWidth}" height="${Math.max(1, Math.abs(baseline - y))}" fill="${attrEscape(value < 0 ? negative : series)}"/>`;
    }).join("");
    return `${axis}${bars}`;
  }
  const points = values.map((value, index) => [
    left + (values.length === 1 ? width / 2 : index * width / (values.length - 1)),
    top + height - ((value - minimum) / span) * height,
  ]);
  const pointText = points.map(([x, y]) => `${x},${y}`).join(" ");
  return `${axis}<polyline points="${pointText}" fill="none" stroke="${attrEscape(series)}" stroke-width="${group.lineWeight}"/>${markerCircles(group, points, values)}`;
}

export function createSpreadsheetSparklineClasses({ workbookRangeRef, worksheetRangeFrame }) {
  class SparklineGroup {
    constructor(worksheet, config = {}) {
      this.worksheet = worksheet;
      this.id = config.id || aid("sp");
      this.type = config.type || "line";
      this.targetRange = workbookRangeRef(config.targetRange || config.locationRange || "A1");
      this.sourceData = workbookRangeRef(config.sourceData || "A1");
      this.dateAxisRange = config.dateAxisRange ? workbookRangeRef(config.dateAxisRange) : undefined;
      this.seriesColor = config.seriesColor || "#0EA5E9";
      this.negativeColor = config.negativeColor;
      this.axisColor = config.axisColor;
      this.markersColor = config.markersColor;
      this.firstMarkerColor = config.firstMarkerColor;
      this.lastMarkerColor = config.lastMarkerColor;
      this.highMarkerColor = config.highMarkerColor;
      this.lowMarkerColor = config.lowMarkerColor;
      this.markers = { ...DEFAULT_MARKERS, ...(config.markers || {}) };
      this.axis = { ...DEFAULT_AXIS, ...(config.axis || {}) };
      this.lineWeight = config.lineWeight ?? 1;
      this.displayHidden = Boolean(config.displayHidden);
      this.displayEmptyCellsAs = config.displayEmptyCellsAs ?? 2;
    }

    get locationRange() { return this.targetRange; }
    set locationRange(value) { this.targetRange = workbookRangeRef(value); }
    get rightToLeft() { return Boolean(this.axis.rightToLeft); }
    set rightToLeft(value) { this.axis.rightToLeft = Boolean(value); }
    get displayXAxis() { return this.axis.showAxis !== false; }
    set displayXAxis(value) { this.axis.showAxis = Boolean(value); }
    get sparklineCount() { return sparklineRangeMappings(this).length; }

    delete() { this.worksheet.sparklineGroups.items = this.worksheet.sparklineGroups.items.filter((group) => group !== this); }
    inspectRecord() { return { kind: "sparkline", id: this.id, sheet: this.worksheet.name, type: this.type, targetRange: this.targetRange.address, sourceData: this.sourceData.address, dateAxisRange: this.dateAxisRange?.address, seriesColor: this.seriesColor, sparklineCount: this.sparklineCount }; }
    values() { return numericValues(this, this.sourceData.address); }
    targetFrame(bounds) { return worksheetRangeFrame(this.worksheet, parseRangeAddress(this.targetRange.address), bounds); }
    toSvg(bounds) {
      const mappings = sparklineRangeMappings(this);
      if (!mappings.length) {
        const frame = this.targetFrame(bounds);
        return `<rect x="${frame.left}" y="${frame.top}" width="${frame.width}" height="${frame.height}" fill="none" stroke="#EF4444" stroke-dasharray="3 2"/>`;
      }
      return mappings.map((item) => renderSparkline(
        this,
        worksheetRangeFrame(this.worksheet, parseRangeAddress(item.targetAddress), bounds),
        numericValues(this, item.sourceAddress),
      )).join("");
    }
    toJSON() { return { id: this.id, type: this.type, targetRange: this.targetRange, sourceData: this.sourceData, dateAxisRange: this.dateAxisRange, seriesColor: this.seriesColor, negativeColor: this.negativeColor, axisColor: this.axisColor, markersColor: this.markersColor, firstMarkerColor: this.firstMarkerColor, lastMarkerColor: this.lastMarkerColor, highMarkerColor: this.highMarkerColor, lowMarkerColor: this.lowMarkerColor, markers: this.markers, axis: this.axis, lineWeight: this.lineWeight, displayHidden: this.displayHidden, displayEmptyCellsAs: this.displayEmptyCellsAs }; }
  }

  class SparklineGroupCollection {
    constructor(worksheet) { this.worksheet = worksheet; this.items = []; }
    add(config = {}) { const group = new SparklineGroup(this.worksheet, config); this.items.push(group); return group; }
    getAll() { return [...this.items]; }
    clear() { this.items = []; }
    deleteAll() { this.clear(); }
    delete(groupOrId) { this.items = this.items.filter((group) => group !== groupOrId && group.id !== groupOrId); }
    inspectRecords() { return this.items.map((group) => group.inspectRecord()); }
    toJSON() { return this.items.map((group) => group.toJSON()); }
  }

  class RangeSparklineFacade {
    constructor(range) { this.range = range; }
    add(type, sourceData, config = {}) { return this.range.worksheet.sparklineGroups.add({ ...config, type, targetRange: this.range, sourceData }); }
  }

  return { RangeSparklineFacade, SparklineGroup, SparklineGroupCollection };
}
