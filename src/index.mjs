import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PDF_MIME = "application/pdf";
const LAYOUT_MIME = "application/vnd.open-office-artifact.layout+json";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
let nextId = 1;

function aid(prefix) {
  return `${prefix}/${(nextId++).toString(36).padStart(4, "0")}`;
}

function xmlEscape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function attrEscape(value) {
  return xmlEscape(value).replaceAll('"', "&quot;");
}

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

function columnNumberToLabel(index) {
  let n = index + 1;
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function columnLabelToNumber(label) {
  let n = 0;
  for (const ch of String(label).toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

function parseCellAddress(address) {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(String(address).trim());
  if (!match) throw new Error(`Invalid cell address: ${address}`);
  return { col: columnLabelToNumber(match[1]), row: Number(match[2]) - 1 };
}

function makeCellAddress(row, col) {
  return `${columnNumberToLabel(col)}${row + 1}`;
}

function parseRangeAddress(address) {
  const raw = String(address || "A1").trim();
  const withoutSheet = raw.includes("!") ? raw.slice(raw.lastIndexOf("!") + 1) : raw;
  const [startRaw, endRaw = startRaw] = withoutSheet.split(":");
  const start = parseCellAddress(startRaw);
  const end = parseCellAddress(endRaw);
  const top = Math.min(start.row, end.row);
  const left = Math.min(start.col, end.col);
  const bottom = Math.max(start.row, end.row);
  const right = Math.max(start.col, end.col);
  return { top, left, bottom, right, rowCount: bottom - top + 1, colCount: right - left + 1 };
}

function rangeToAddress(bounds) {
  const start = makeCellAddress(bounds.top, bounds.left);
  const end = makeCellAddress(bounds.bottom, bounds.right);
  return start === end ? start : `${start}:${end}`;
}

function ndjson(records, maxChars = Infinity) {
  const lines = records.map((record) => JSON.stringify(record));
  let text = lines.join("\n");
  let truncated = false;
  if (text.length > maxChars) {
    truncated = true;
    const kept = [];
    let chars = 0;
    for (const line of lines) {
      if (chars + line.length + 1 > maxChars) break;
      kept.push(line);
      chars += line.length + 1;
    }
    kept.push(JSON.stringify({ kind: "notice", message: `Truncated: omitted ${lines.length - kept.length} lines. Increase maxChars or narrow query.` }));
    text = kept.join("\n");
  }
  return { ndjson: text, truncated };
}

function verificationResult(artifactKind, issues, options = {}) {
  const result = {
    artifactKind,
    ok: issues.length === 0,
    issues,
    ...ndjson(issues, options.maxChars ?? Infinity),
  };
  return result;
}

function verificationIssue(artifactKind, type, message, details = {}) {
  return { kind: "verificationIssue", artifactKind, type, severity: details.severity || "error", message, ...details };
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

export async function renderArtifact(artifact, options = {}) {
  const artifactKind = inferArtifactKind(artifact);
  if (!artifact || (typeof artifact.render !== "function" && typeof artifact.export !== "function")) {
    throw new Error("Artifact does not expose a render() or export() method.");
  }
  const renderer = typeof artifact.render === "function" ? artifact.render.bind(artifact) : artifact.export.bind(artifact);
  const blob = await renderer(options);
  if (!(blob instanceof FileBlob)) {
    return new FileBlob(await blob.arrayBuffer?.() ?? String(blob), {
      type: blob.type || "application/octet-stream",
      metadata: { artifactKind, format: options.format || blob.type || "unknown" },
    });
  }
  blob.metadata = {
    ...(blob.metadata || {}),
    artifactKind,
    format: options.format || blob.metadata?.format || blob.type,
    page: options.page,
    pageIndex: options.pageIndex,
    slide: options.slide,
    sheetName: options.sheetName,
    range: options.range,
  };
  return blob;
}

export const HELP_CATALOG = [
  { artifactKind: "workbook", kind: "api", name: "Workbook.create", summary: "Create an empty workbook; add worksheets before editing." },
  { artifactKind: "workbook", kind: "api", name: "SpreadsheetFile.importXlsx", summary: "Load an XLSX file into a Workbook facade." },
  { artifactKind: "workbook", kind: "api", name: "SpreadsheetFile.exportXlsx", summary: "Serialize a Workbook facade to an XLSX FileBlob." },
  { artifactKind: "workbook", kind: "api", name: "worksheet.getRange", summary: "Select an A1 range for values, formulas, formatting, merge, fill, and copy operations." },
  { artifactKind: "workbook", kind: "api", name: "workbook.inspect", summary: "Emit bounded NDJSON records for workbook, sheets, tables, formulas, matches, comments, validations, conditional formats, and drawings." },
  { artifactKind: "workbook", kind: "api", name: "workbook.render", summary: "Return a lightweight SVG preview for a sheet or range in the current clean-room MVP." },
  { artifactKind: "workbook", kind: "api", name: "workbook.verify", summary: "Return bounded QA issues for sheets, formulas, tables, charts, and comments." },
  { artifactKind: "workbook", kind: "api", name: "workbook.trace", summary: "Return a formula precedent tree and bounded NDJSON trace for a target cell." },
  { artifactKind: "workbook", kind: "api", name: "range.dataValidation", summary: "Assign a validation rule to a range or use sheet.dataValidations.add({ range, rule })." },
  { artifactKind: "workbook", kind: "api", name: "range.conditionalFormats.add", summary: "Add a conditional formatting rule to a range; addCustom(expression, format) creates expression rules." },
  { artifactKind: "workbook", kind: "api", name: "workbook.comments.addThread", summary: "Create threaded comments after comments.setSelf({ displayName }); resolve with wb.resolve('th/...')." },
  { artifactKind: "workbook", kind: "api", name: "sheet.tables.add", summary: "Create an inspectable worksheet table over an A1 range with rows.add, getDataRows, getHeaderRowRange, style, and visibility toggles." },
  { artifactKind: "workbook", kind: "api", name: "sheet.charts.add", summary: "Create an inspectable worksheet chart from a range or config; setData(range) infers categories and series formulas." },
  { artifactKind: "workbook", kind: "api", name: "sheet.images.add", summary: "Create an inspectable worksheet image placeholder from a data URL, URI, or prompt with 0-based cell anchors and pixel extents." },
  { artifactKind: "workbook", kind: "api", name: "sheet.sparklineGroups.add", summary: "Create line/column/stacked sparklines from sourceData into a targetRange; range.sparklines.add is a shorthand." },
  { artifactKind: "workbook", kind: "formula", name: "fx.SUM", category: "math-trig", summary: "Sum numeric values across arguments and ranges.", examples: ["=SUM(A1:A10)"] },
  { artifactKind: "workbook", kind: "formula", name: "fx.PMT", category: "financial", summary: "Calculate a loan payment for constant payments and constant interest rate.", examples: ["=PMT(rate,nper,pv)"], notes: ["Catalog entry only in MVP; full financial formula evaluation is roadmap."] },

  { artifactKind: "presentation", kind: "api", name: "Presentation.create", summary: "Create a deck with a default or explicit slide size." },
  { artifactKind: "presentation", kind: "api", name: "presentation.inspect", summary: "Emit NDJSON for deck, slides, textboxes, shapes, tables, charts, images, notes, and layout." },
  { artifactKind: "presentation", kind: "api", name: "presentation.resolve", summary: "Map stable inspect anchor IDs back to editable facade objects." },
  { artifactKind: "presentation", kind: "api", name: "presentation.export", summary: "Export a slide preview, deck montage, or layout JSON." },
  { artifactKind: "presentation", kind: "api", name: "presentation.validateLayout", summary: "Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow." },
  { artifactKind: "presentation", kind: "api", name: "presentation.verify", summary: "Return presentation QA issues, currently including layout validation issues." },
  { artifactKind: "presentation", kind: "api", name: "slide.shapes.add", summary: "Add a shape/textbox with geometry, position, fill, line, and text." },
  { artifactKind: "presentation", kind: "api", name: "slide.compose", summary: "Materialize a clean-room compose tree with row, column, grid, layers, box, paragraph, shape, table, chart, image, and rule nodes into editable slide objects." },
  { artifactKind: "presentation", kind: "api", name: "slide.autoLayout", summary: "Place existing shapes inside a frame using horizontal or vertical flow, gap, padding, and alignment options." },
  { artifactKind: "presentation", kind: "api", name: "slide.tables.add", summary: "Add an inspectable native-style table facade with rows, columns, values, cells, layout JSON, and SVG/PPTX placeholder output." },
  { artifactKind: "presentation", kind: "api", name: "slide.charts.add", summary: "Add an inspectable chart facade with chartType, title, categories, series, layout JSON, SVG preview, and PPTX placeholder output." },
  { artifactKind: "presentation", kind: "api", name: "slide.images.add", summary: "Add an inspectable image facade with alt text, prompt/URI/data URL metadata, fit, frame, layout JSON, SVG preview, and PPTX placeholder output." },
  { artifactKind: "presentation", kind: "api", name: "compose.column", summary: "Create a vertical compose container. Use width/height fill, hug, or fixed pixels; gap and padding are in pixels." },
  { artifactKind: "presentation", kind: "api", name: "compose.paragraph", summary: "Create an editable text block with name, className/style text tokens, and stable inspect output." },

  { artifactKind: "document", kind: "api", name: "DocumentModel.create", summary: "Create a document with paragraph, list, table, header/footer, style, and comment blocks." },
  { artifactKind: "document", kind: "api", name: "document.addParagraph", summary: "Append a styled paragraph block and return an inspectable/resolveable paragraph object." },
  { artifactKind: "document", kind: "api", name: "document.addListItem", summary: "Append a real numbered or bulleted list item backed by DOCX numbering definitions." },
  { artifactKind: "document", kind: "api", name: "document.addHeader", summary: "Add header text exported as a DOCX header part and referenced from section properties." },
  { artifactKind: "document", kind: "api", name: "document.addFooter", summary: "Add footer text exported as a DOCX footer part and referenced from section properties." },
  { artifactKind: "document", kind: "api", name: "document.addHyperlink", summary: "Append an external hyperlink backed by a DOCX relationship and w:hyperlink element." },
  { artifactKind: "document", kind: "api", name: "document.addField", summary: "Append a Word field block exported as w:fldSimple with instruction text such as PAGE, REF, PAGEREF, or TOC." },
  { artifactKind: "document", kind: "api", name: "document.addCitation", summary: "Append a citation block with visible text and structured metadata preserved through clean-room DOCX metadata." },
  { artifactKind: "document", kind: "api", name: "document.addImage", summary: "Append an inspectable image block; dataUrl images export as native DOCX media parts with DrawingML inline pictures." },
  { artifactKind: "document", kind: "api", name: "document.addChange", summary: "Append a tracked insertion or deletion block backed by native DOCX w:ins/w:del revision markup." },
  { artifactKind: "document", kind: "api", name: "document.addInsertion", summary: "Append a tracked insertion with author/date metadata and native DOCX w:ins export." },
  { artifactKind: "document", kind: "api", name: "document.addDeletion", summary: "Append a tracked deletion with author/date metadata and native DOCX w:del/w:delText export." },
  { artifactKind: "document", kind: "api", name: "document.addTable", summary: "Append a Word-style table block with rows, columns, cell values, and style metadata." },
  { artifactKind: "document", kind: "api", name: "document.addComment", summary: "Attach a comment to a paragraph or table block using a stable target ID." },
  { artifactKind: "document", kind: "api", name: "document.verify", summary: "Return QA issues for fake lists, invalid links, dangling comments, and table cells that look like prose blocks." },
  { artifactKind: "document", kind: "api", name: "DocumentFile.exportDocx", summary: "Export DocumentModel to a DOCX package with document.xml, styles.xml, comments.xml, numbering.xml, header/footer parts, hyperlinks, fields, citations, and metadata." },

  { artifactKind: "pdf", kind: "api", name: "PdfArtifact.create", summary: "Create a modeled PDF artifact with pages, text, and table regions." },
  { artifactKind: "pdf", kind: "api", name: "pdf.extractText", summary: "Extract modeled text across all pages or a selected page." },
  { artifactKind: "pdf", kind: "api", name: "pdf.extractTables", summary: "Extract modeled table values and bounding boxes across all pages or a selected page." },
  { artifactKind: "pdf", kind: "api", name: "pdf.inspect", summary: "Emit bounded NDJSON for pages, text, and table records." },
  { artifactKind: "pdf", kind: "api", name: "pdf.render", summary: "Render a modeled PDF page to SVG in the current clean-room MVP." },
  { artifactKind: "pdf", kind: "api", name: "pdf.verify", summary: "Return QA issues for empty pages, Unicode dashes, malformed tables, and out-of-bounds table boxes." },
  { artifactKind: "pdf", kind: "api", name: "PdfFile.exportPdf", summary: "Export a modeled PDF artifact to a minimal PDF with visible text/table rows and embedded clean-room metadata." },
  { artifactKind: "pdf", kind: "api", name: "PdfFile.importPdf", summary: "Import clean-room generated PDFs from metadata or heuristically extract visible text and pipe-delimited table rows." },

  { artifactKind: "shared", kind: "api", name: "verifyArtifact", summary: "Run an artifact's verify() method and return a bounded NDJSON QA report." },
  { artifactKind: "shared", kind: "api", name: "renderArtifact", summary: "Render an artifact through its render/export method and attach normalized FileBlob metadata." },
];

export function helpArtifact(artifactOrKind = "*", query = "*", options = {}) {
  const artifactKind = typeof artifactOrKind === "string" ? artifactOrKind : inferArtifactKind(artifactOrKind);
  const q = String(query || "*").toLowerCase();
  const search = String(options.search || "").toLowerCase();
  const records = HELP_CATALOG.filter((item) => {
    const kindMatch = artifactKind === "*" || item.artifactKind === artifactKind || (artifactKind === "unknown" && item.artifactKind === "shared");
    if (!kindMatch) return false;
    const haystack = `${item.name}\n${item.summary}\n${item.category || ""}\n${(item.examples || []).join("\n")}\n${(item.notes || []).join("\n")}`.toLowerCase();
    const queryMatch = q === "*" || item.name.toLowerCase().includes(q.replace("fx.", "")) || haystack.includes(q);
    const searchMatch = !search || haystack.includes(search);
    return queryMatch && searchMatch;
  });
  return ndjson(records, options.maxChars ?? Infinity);
}

function normalizeKinds(kind, fallback) {
  if (!kind) return new Set(fallback);
  return new Set(String(kind).split(",").map((k) => k.trim()).filter(Boolean));
}

function toUint8Array(data) {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  if (typeof data === "string") return encoder.encode(data);
  throw new Error("Unsupported binary payload");
}

export function node(type, props = {}, children = []) {
  const normalizedProps = props && !Array.isArray(props) && typeof props === "object" ? props : {};
  const rawChildren = Array.isArray(props) || typeof props === "string" || typeof props === "number" ? props : children;
  return { type, props: normalizedProps, children: normalizeComposeChildren(rawChildren) };
}

export const row = (props = {}, children = []) => node("row", props, children);
export const column = (props = {}, children = []) => node("column", props, children);
export const grid = (props = {}, children = []) => node("grid", props, children);
export const layers = (props = {}, children = []) => node("layers", props, children);
export const box = (props = {}, children = []) => node("box", props, children);
export const paragraph = (props = {}, children = []) => node("paragraph", props, children);
export const run = (props = {}, children = []) => node("run", props, children);
export const shape = (props = {}, children = []) => node("shape", props, children);
export const image = (props = {}, children = []) => node("image", props, children);
export const table = (props = {}, children = []) => node("table", props, children);
export const chart = (props = {}, children = []) => node("chart", props, children);
export const rule = (props = {}, children = []) => node("rule", props, children);

function normalizeComposeChildren(children) {
  if (children == null || children === false) return [];
  if (!Array.isArray(children)) return [children];
  return children.flatMap((child) => normalizeComposeChildren(child));
}

function isComposeNode(value) {
  return value && typeof value === "object" && typeof value.type === "string" && Array.isArray(value.children);
}

function textFromComposeChildren(children) {
  return normalizeComposeChildren(children).map((child) => {
    if (typeof child === "string" || typeof child === "number") return String(child);
    if (isComposeNode(child)) return textFromComposeChildren(child.children);
    return "";
  }).join("");
}

function normalizePadding(padding = {}) {
  if (typeof padding === "number") return { top: padding, right: padding, bottom: padding, left: padding };
  return {
    top: padding.top ?? padding.y ?? 0,
    right: padding.right ?? padding.x ?? 0,
    bottom: padding.bottom ?? padding.y ?? 0,
    left: padding.left ?? padding.x ?? 0,
  };
}

function innerFrame(frame, padding) {
  return {
    left: frame.left + padding.left,
    top: frame.top + padding.top,
    width: Math.max(0, frame.width - padding.left - padding.right),
    height: Math.max(0, frame.height - padding.top - padding.bottom),
  };
}

function resolveColorToken(value, fallback = "transparent") {
  if (!value) return fallback;
  const raw = String(value);
  const base = raw.split("/")[0];
  const tokens = {
    white: "#ffffff",
    black: "#000000",
    transparent: "transparent",
    "slate-50": "#f8fafc",
    "slate-100": "#f1f5f9",
    "slate-200": "#e2e8f0",
    "slate-300": "#cbd5e1",
    "slate-500": "#64748b",
    "slate-600": "#475569",
    "slate-700": "#334155",
    "slate-900": "#0f172a",
    "slate-950": "#020617",
    "sky-50": "#f0f9ff",
    "sky-100": "#e0f2fe",
    "sky-500": "#0ea5e9",
    "sky-600": "#0284c7",
    accent1: "#156082",
    tx1: "#1f1f1f",
    bg1: "#ffffff",
  };
  return tokens[base] || raw;
}

function parseTextStyle(props = {}) {
  const style = {};
  const className = String(props.className || "");
  for (const token of className.split(/\s+/).filter(Boolean)) {
    if (token === "font-bold") style.bold = true;
    if (token === "font-semibold") style.bold = true;
    if (token.startsWith("text-")) {
      const sizeMap = { xs: 12, sm: 14, base: 16, lg: 18, xl: 20, "2xl": 24, "3xl": 30, "4xl": 36, "5xl": 48, "6xl": 60 };
      const key = token.slice(5);
      if (sizeMap[key]) style.fontSize = sizeMap[key];
      else if (/^\[\d+px\]$/.test(key)) style.fontSize = Number(key.slice(1, -3));
      else style.color = resolveColorToken(key, style.color);
    }
    if (token.startsWith("leading-")) {
      const key = token.slice(8);
      if (key === "tight") style.lineSpacing = 1.1;
      else if (key === "relaxed") style.lineSpacing = 1.35;
      else if (/^\[\d+(?:\.\d+)?\]$/.test(key)) style.lineSpacing = Number(key.slice(1, -1));
    }
  }
  if (typeof props.style === "string") {
    const font = /font:\s*(\d+)\s+(\d+)px/i.exec(props.style);
    if (font) { style.bold = Number(font[1]) >= 600; style.fontSize = Number(font[2]); }
    const color = /color:\s*([^;]+)/i.exec(props.style);
    if (color) style.color = resolveColorToken(color[1].trim(), style.color);
    const leading = /leading:\s*([\d.]+)/i.exec(props.style);
    if (leading) style.lineSpacing = Number(leading[1]);
  } else if (props.style && typeof props.style === "object") {
    Object.assign(style, props.style);
  }
  return style;
}

function styleFromClassName(className = "") {
  const style = {};
  for (const token of String(className).split(/\s+/).filter(Boolean)) {
    if (token.startsWith("bg-")) style.fill = resolveColorToken(token.slice(3), style.fill);
    if (token.startsWith("rounded-")) style.borderRadius = token;
  }
  return style;
}

function composeIntrinsicSize(composeNode) {
  const props = composeNode.props || {};
  const text = textFromComposeChildren(composeNode.children);
  const textStyle = parseTextStyle(props);
  const fontSize = textStyle.fontSize || 20;
  const lineCount = Math.max(1, text.split(/\r?\n/).length);
  return {
    width: typeof props.width === "number" ? props.width : Math.max(160, Math.min(720, text.length * fontSize * 0.55 + 24)),
    height: typeof props.height === "number" ? props.height : Math.max(36, lineCount * fontSize * (textStyle.lineSpacing || 1.2) + 18),
  };
}

function composeChildFrames(children, frame, direction, gap) {
  const mainSize = direction === "row" ? "width" : "height";
  const crossSize = direction === "row" ? "height" : "width";
  const fixed = children.reduce((sum, child) => {
    const prop = child.props?.[mainSize];
    if (typeof prop === "number") return sum + prop;
    if (prop === "hug") return sum + composeIntrinsicSize(child)[mainSize];
    return sum;
  }, 0);
  const fillCount = children.filter((child) => child.props?.[mainSize] === "fill" || child.props?.[mainSize] == null).length || 1;
  const available = Math.max(0, frame[mainSize] - fixed - gap * Math.max(0, children.length - 1));
  const fillSize = available / fillCount;
  let cursor = direction === "row" ? frame.left : frame.top;
  return children.map((child) => {
    const intrinsic = composeIntrinsicSize(child);
    const main = typeof child.props?.[mainSize] === "number" ? child.props[mainSize] : child.props?.[mainSize] === "hug" ? intrinsic[mainSize] : fillSize;
    const cross = typeof child.props?.[crossSize] === "number" ? child.props[crossSize] : child.props?.[crossSize] === "hug" ? intrinsic[crossSize] : frame[crossSize];
    const childFrame = direction === "row"
      ? { left: cursor, top: frame.top, width: main, height: cross }
      : { left: frame.left, top: cursor, width: cross, height: main };
    cursor += main + gap;
    return childFrame;
  });
}

function normalizeTrack(track) {
  if (typeof track === "number") return { mode: "fixed", value: track };
  if (typeof track === "string") return track === "fixed" ? { mode: "fixed", value: 0 } : { mode: "fr", value: 1 };
  return { mode: track?.mode || "fr", value: Number(track?.value ?? 1) };
}

function resolveGridTracks(total, tracks, fallbackCount, gap) {
  const normalized = (tracks?.length ? tracks : Array.from({ length: fallbackCount }, () => ({ mode: "fr", value: 1 }))).map(normalizeTrack);
  const fixed = normalized.reduce((sum, track) => track.mode === "fixed" ? sum + track.value : sum, 0);
  const fr = normalized.reduce((sum, track) => track.mode === "fr" ? sum + Math.max(0, track.value) : sum, 0) || 1;
  const available = Math.max(0, total - fixed - gap * Math.max(0, normalized.length - 1));
  return normalized.map((track) => track.mode === "fixed" ? track.value : available * Math.max(0, track.value) / fr);
}

function gridChildFrame(frame, columns, rows, columnGap, rowGap, columnIndex, rowIndex, columnSpan = 1, rowSpan = 1) {
  const left = frame.left + columns.slice(0, columnIndex).reduce((sum, value) => sum + value, 0) + columnGap * columnIndex;
  const top = frame.top + rows.slice(0, rowIndex).reduce((sum, value) => sum + value, 0) + rowGap * rowIndex;
  const width = columns.slice(columnIndex, columnIndex + columnSpan).reduce((sum, value) => sum + value, 0) + columnGap * Math.max(0, columnSpan - 1);
  const height = rows.slice(rowIndex, rowIndex + rowSpan).reduce((sum, value) => sum + value, 0) + rowGap * Math.max(0, rowSpan - 1);
  return { left, top, width, height };
}

function materializeComposeNode(slide, composeNode, frame) {
  if (typeof composeNode === "string" || typeof composeNode === "number") {
    return materializeComposeNode(slide, paragraph({}, [String(composeNode)]), frame);
  }
  if (!isComposeNode(composeNode)) return [];
  const props = composeNode.props || {};
  const children = normalizeComposeChildren(composeNode.children).filter((child) => child !== null && child !== undefined && child !== false);
  const type = composeNode.type;
  if (type === "row" || type === "column") {
    const pad = normalizePadding(props.padding);
    const inner = innerFrame(frame, pad);
    const childFrames = composeChildFrames(children.filter(isComposeNode), inner, type, Number(props.gap || 0));
    return children.filter(isComposeNode).flatMap((child, index) => materializeComposeNode(slide, child, childFrames[index]));
  }
  if (type === "grid") {
    const gridChildren = children.filter(isComposeNode);
    const pad = normalizePadding(props.padding);
    const inner = innerFrame(frame, pad);
    const columnGap = Number(props.columnGap ?? props.gap ?? 0);
    const rowGap = Number(props.rowGap ?? props.gap ?? 0);
    const fallbackColumns = Math.max(1, props.columns?.length || Math.ceil(Math.sqrt(gridChildren.length || 1)));
    const columns = resolveGridTracks(inner.width, props.columns, fallbackColumns, columnGap);
    const fallbackRows = Math.max(1, props.rows?.length || Math.ceil((gridChildren.length || 1) / columns.length));
    const rows = resolveGridTracks(inner.height, props.rows, fallbackRows, rowGap);
    return gridChildren.flatMap((child, index) => {
      const columnIndex = Math.min(columns.length - 1, Number(child.props?.column ?? child.props?.col ?? (index % columns.length)));
      const rowIndex = Math.min(rows.length - 1, Number(child.props?.row ?? Math.floor(index / columns.length)));
      const columnSpan = Math.min(columns.length - columnIndex, Math.max(1, Number(child.props?.columnSpan ?? 1)));
      const rowSpan = Math.min(rows.length - rowIndex, Math.max(1, Number(child.props?.rowSpan ?? 1)));
      return materializeComposeNode(slide, child, gridChildFrame(inner, columns, rows, columnGap, rowGap, columnIndex, rowIndex, columnSpan, rowSpan));
    });
  }
  if (type === "layers") {
    const pad = normalizePadding(props.padding);
    const inner = innerFrame(frame, pad);
    return children.filter(isComposeNode).flatMap((child) => materializeComposeNode(slide, child, {
      left: inner.left,
      top: inner.top,
      width: typeof child.props?.width === "number" ? child.props.width : inner.width,
      height: typeof child.props?.height === "number" ? child.props.height : inner.height,
    }));
  }
  if (type === "box") {
    const classStyle = styleFromClassName(props.className);
    const surface = slide.shapes.add({
      id: props.id,
      name: props.name,
      geometry: props.geometry || "roundRect",
      position: frame,
      fill: props.fill || classStyle.fill || "transparent",
      line: props.line || { fill: "transparent", width: 0 },
      borderRadius: props.borderRadius || classStyle.borderRadius,
    });
    const pad = normalizePadding(props.padding ?? { x: 0, y: 0 });
    return [surface, ...children.filter(isComposeNode).flatMap((child) => materializeComposeNode(slide, child, innerFrame(frame, pad)))];
  }
  if (type === "paragraph") {
    const shape = slide.shapes.add({
      id: props.id,
      name: props.name,
      geometry: "textbox",
      position: frame,
      fill: "transparent",
      line: { fill: "transparent", width: 0 },
      text: textFromComposeChildren(children),
    });
    shape.text.style = parseTextStyle(props);
    return [shape];
  }
  if (type === "shape") {
    const classStyle = styleFromClassName(props.className);
    const shape = slide.shapes.add({
      ...props,
      position: frame,
      fill: props.fill || classStyle.fill || "transparent",
      text: textFromComposeChildren(children) || props.text,
    });
    shape.text.style = parseTextStyle(props);
    return [shape];
  }
  if (type === "table") {
    return [slide.tables.add({ ...props, position: frame })];
  }
  if (type === "chart") {
    return [slide.charts.add(props.chartType || props.type || "bar", { ...props, position: frame })];
  }
  if (type === "image") {
    return [slide.images.add({ ...props, position: frame, alt: props.alt || textFromComposeChildren(children) || props.name })];
  }
  if (type === "rule") {
    const horizontal = (props.width ?? frame.width) >= (props.height ?? props.weight ?? 2);
    const shape = slide.shapes.add({
      id: props.id,
      name: props.name,
      geometry: "rect",
      position: { left: frame.left, top: frame.top, width: horizontal ? frame.width : Number(props.weight || 2), height: horizontal ? Number(props.weight || 2) : frame.height },
      fill: props.stroke || "#0f172a",
      line: { fill: props.stroke || "#0f172a", width: 0 },
    });
    return [shape];
  }
  return children.filter(isComposeNode).flatMap((child) => materializeComposeNode(slide, child, frame));
}

export class FileBlob {
  constructor(data, options = {}) {
    this.bytes = toUint8Array(data ?? new Uint8Array());
    this.type = options.type || "application/octet-stream";
    this.metadata = options.metadata || {};
  }

  async arrayBuffer() {
    return this.bytes.buffer.slice(this.bytes.byteOffset, this.bytes.byteOffset + this.bytes.byteLength);
  }

  async text() {
    return decoder.decode(this.bytes);
  }

  async save(filePath) {
    await fs.mkdir(path.dirname(path.resolve(filePath)), { recursive: true });
    await fs.writeFile(filePath, this.bytes);
    return filePath;
  }

  static async load(filePath, options = {}) {
    const bytes = await fs.readFile(filePath);
    return new FileBlob(bytes, options);
  }
}

class WorksheetCollection {
  constructor(workbook) {
    this.workbook = workbook;
    this.items = [];
  }

  add(name = `Sheet${this.items.length + 1}`) {
    const worksheet = new Worksheet(this.workbook, name);
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

  getActiveWorksheet() {
    if (!this.items[0]) throw new Error("Workbook has no worksheets. Add one first.");
    return this.items[0];
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
  if (rangeOrRef?.sheetName !== undefined && rangeOrRef?.address) return { sheetName: rangeOrRef.sheetName, address: rangeOrRef.address };
  if (rangeOrRef?.cell) return workbookRangeRef(rangeOrRef.cell);
  if (rangeOrRef?.range) return workbookRangeRef(rangeOrRef.range);
  return { sheetName: undefined, address: "A1" };
}

class CommentThread {
  constructor(workbook, target, text, author) {
    this.workbook = workbook;
    this.id = aid("th");
    this.target = workbookRangeRef(target);
    this.author = author || workbook.comments.self?.displayName || "User";
    this.comments = [{ author: this.author, text: String(text ?? "") }];
    this.resolved = false;
  }

  addReply(text) {
    this.comments.push({ author: this.workbook.comments.self?.displayName || this.author, text: String(text ?? "") });
    return this;
  }

  resolve() { this.resolved = true; return this; }
  reopen() { this.resolved = false; return this; }

  inspectRecord() {
    return { kind: "thread", id: this.id, sheet: this.target.sheetName, address: this.target.address, author: this.author, resolved: this.resolved, replies: Math.max(0, this.comments.length - 1), textPreview: this.comments.map((comment) => comment.text).join("\n").slice(0, 300) };
  }

  toJSON() { return { id: this.id, target: this.target, author: this.author, comments: this.comments, resolved: this.resolved }; }
}

class CommentsCollection {
  constructor(workbook) { this.workbook = workbook; this.self = undefined; this.threads = []; }
  setSelf(self) { this.self = { displayName: self?.displayName || "User" }; return this.self; }
  addThread(target, text) { const thread = new CommentThread(this.workbook, target, text, this.self?.displayName); this.threads.push(thread); return thread; }
  getItem(id) { return this.threads.find((thread) => thread.id === id); }
  toJSON() { return { self: this.self, threads: this.threads.map((thread) => thread.toJSON()) }; }
}

class WorksheetRuleCollection {
  constructor(worksheet, kind) { this.worksheet = worksheet; this.kind = kind; this.items = []; }
  add(config = {}) { const record = { id: aid(this.kind === "dataValidation" ? "dv" : "cf"), kind: this.kind, sheet: this.worksheet.name, ...config }; this.items.push(record); return record; }
  deleteAll() { this.items = []; }
  clear() { this.deleteAll(); }
  inspectRecords() { return this.items.map((item) => ({ ...item })); }
  toJSON() { return this.items.map((item) => ({ ...item })); }
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
    this.showHeaders = this.hasHeaders;
    this.showTotals = Boolean(config.showTotals);
    this.showBandedColumns = Boolean(config.showBandedColumns);
    this.showFilterButton = config.showFilterButton ?? true;
    this.style = config.style || "TableStyleMedium2";
    this.values = config.values ? config.values.map((row) => [...row]) : range.values.map((row) => [...row]);
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
    return { kind: "table", id: this.id, sheet: this.worksheet.name, name: this.name, address: this.range, rows: this.rowCount, cols: this.columnCount, hasHeaders: this.hasHeaders, style: this.style, values: this.values };
  }

  toSvg(bounds) {
    const tableBounds = parseRangeAddress(this.range);
    const left = 40 + (tableBounds.left - bounds.left) * 96;
    const top = 40 + (tableBounds.top - bounds.top) * 28;
    const width = Math.max(96, this.columnCount * 96);
    const height = Math.max(28, this.rowCount * 28);
    return `<rect x="${left}" y="${top}" width="${width}" height="${height}" fill="none" stroke="#0ea5e9" stroke-width="2"/><text x="${left}" y="${Math.max(12, top - 6)}" font-family="Arial" font-size="11" fill="#0284c7">${xmlEscape(this.name)}</text>`;
  }

  toJSON() { return { id: this.id, name: this.name, range: this.range, hasHeaders: this.hasHeaders, showHeaders: this.showHeaders, showTotals: this.showTotals, showBandedColumns: this.showBandedColumns, showFilterButton: this.showFilterButton, style: this.style, values: this.values }; }
}

class WorksheetTableCollection {
  constructor(worksheet) { this.worksheet = worksheet; this.items = []; }
  add(rangeOrConfig, hasHeaders = true, name) { const table = new WorksheetTable(this.worksheet, rangeOrConfig, hasHeaders, name); this.items.push(table); return table; }
  getItemOrNullObject(name) { return this.items.find((table) => table.name === name) || { isNullObject: true }; }
  deleteAll() { this.items = []; }
  inspectRecords() { return this.items.map((table) => table.inspectRecord()); }
  toJSON() { return this.items.map((table) => table.toJSON()); }
}

class WorksheetChartSeriesCollection {
  constructor(chart) { this.chart = chart; this.items = []; }
  add(name, values = []) { const series = { name, values, categoryFormula: undefined, formula: undefined, fill: undefined }; this.items.push(series); return series; }
  getItemAt(index) { return this.items[index]; }
  toJSON() { return this.items.map((item) => ({ ...item })); }
}

class WorksheetChart {
  constructor(worksheet, chartType = "bar", sourceOrConfig = {}) {
    this.worksheet = worksheet;
    this.id = sourceOrConfig.id || aid("wch");
    this.type = chartType;
    this.name = sourceOrConfig.name || `Chart ${worksheet.charts.items.length + 1}`;
    this.title = sourceOrConfig.title || "";
    this.hasLegend = sourceOrConfig.hasLegend ?? true;
    this.categories = sourceOrConfig.categories || [];
    this.position = sourceOrConfig.position || { left: 420, top: 40, width: 360, height: 220 };
    this.series = new WorksheetChartSeriesCollection(this);
    if (sourceOrConfig.series) sourceOrConfig.series.forEach((series) => this.series.add(series.name, series.values || []));
    if (sourceOrConfig instanceof Range) this.setData(sourceOrConfig);
    else if (sourceOrConfig && sourceOrConfig.worksheet instanceof Worksheet) this.setData(sourceOrConfig);
  }

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
    this.position = { left: 40 + start.col * 96, top: 40 + start.row * 28, width: Math.max(120, (end.col - start.col + 1) * 96), height: Math.max(80, (end.row - start.row + 1) * 28) };
    return this;
  }

  inspectRecord() { return { kind: "drawing", drawingType: "chart", id: this.id, sheet: this.worksheet.name, name: this.name, chartType: this.type, title: this.title, categories: this.categories, series: this.series.items.length, bbox: [this.position.left, this.position.top, this.position.width, this.position.height], bboxUnit: "px" }; }

  toSvg() {
    const p = this.position;
    const values = this.series.items[0]?.values || [];
    const max = Math.max(1, ...values.map((value) => Number(value) || 0));
    const plot = { left: p.left + 28, top: p.top + 36, width: Math.max(0, p.width - 44), height: Math.max(0, p.height - 62) };
    const barW = values.length ? plot.width / values.length * 0.65 : 0;
    const gap = values.length ? plot.width / values.length * 0.35 : 0;
    const bars = values.map((value, index) => {
      const h = plot.height * (Number(value) || 0) / max;
      return `<rect x="${plot.left + index * (barW + gap) + gap / 2}" y="${plot.top + plot.height - h}" width="${barW}" height="${h}" fill="#38bdf8"/>`;
    }).join("");
    return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#94a3b8"/><text x="${p.left + 8}" y="${p.top + 22}" font-family="Arial" font-size="13" font-weight="700" fill="#0f172a">${xmlEscape(this.title || this.name)}</text>${bars}`;
  }

  toJSON() { return { id: this.id, type: this.type, name: this.name, title: this.title, hasLegend: this.hasLegend, categories: this.categories, position: this.position, series: this.series.toJSON() }; }
}

class WorksheetChartCollection {
  constructor(worksheet) { this.worksheet = worksheet; this.items = []; }
  add(chartType, sourceOrConfig = {}) { const chart = new WorksheetChart(this.worksheet, chartType, sourceOrConfig); this.items.push(chart); return chart; }
  getItemOrNullObject(name) { return this.items.find((chart) => chart.name === name) || { isNullObject: true }; }
  deleteAll() { this.items = []; }
  inspectRecords() { return this.items.map((chart) => chart.inspectRecord()); }
  toJSON() { return this.items.map((chart) => chart.toJSON()); }
}

function worksheetAnchorFrame(anchor = {}) {
  const from = anchor.from || { row: 0, col: 0 };
  const extent = anchor.extent || {};
  return {
    left: 40 + Number(from.col || 0) * 96 + Number(anchor.colOffsetPx || from.colOffsetPx || 0),
    top: 40 + Number(from.row || 0) * 28 + Number(anchor.rowOffsetPx || from.rowOffsetPx || 0),
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
  }

  get position() { return worksheetAnchorFrame(this.anchor); }
  inspectRecord() { const p = this.position; return { kind: "drawing", drawingType: "image", id: this.id, sheet: this.worksheet.name, name: this.name, alt: this.alt, prompt: this.prompt, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px" }; }
  replace(config = {}) { Object.assign(this, config); return this; }
  toSvg() { const p = this.position; const image = this.dataUrl ? `<image href="${attrEscape(this.dataUrl)}" x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" preserveAspectRatio="xMidYMid meet"/>` : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#fef3c7" stroke="#f59e0b"/>`; return `${image}<text x="${p.left + 8}" y="${p.top + 20}" font-family="Arial" font-size="12" fill="#92400e">${xmlEscape(this.alt || this.prompt || this.name)}</text>`; }
  toJSON() { return { id: this.id, name: this.name, dataUrl: this.dataUrl, uri: this.uri, prompt: this.prompt, alt: this.alt, anchor: this.anchor, fit: this.fit }; }
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
  targetFrame(bounds) { const target = parseRangeAddress(this.targetRange.address); return { left: 40 + (target.left - bounds.left) * 96, top: 40 + (target.top - bounds.top) * 28, width: Math.max(96, target.colCount * 96), height: Math.max(28, target.rowCount * 28) }; }
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
  add(ruleType, config = {}) { return this.range.worksheet.conditionalFormattings.add({ range: rangeToAddress(this.range.bounds), ruleType, ...config }); }
  addCustom(expression, format = {}) { return this.add("expression", { formula: expression, format }); }
  deleteAll() { this.range.worksheet.conditionalFormattings.items = this.range.worksheet.conditionalFormattings.items.filter((item) => item.range !== rangeToAddress(this.range.bounds)); }
  clear() { this.deleteAll(); }
}

export class Workbook {
  constructor() {
    this.id = aid("wb");
    this.worksheets = new WorksheetCollection(this);
    this.comments = new CommentsCollection(this);
  }

  static create() {
    return new Workbook();
  }

  static async fromCSV(text, options = {}) {
    const workbook = Workbook.create();
    await workbook.fromCSV(text, options);
    return workbook;
  }

  async fromCSV(text, options = {}) {
    const sheet = this.worksheets.add(options.sheetName || `Sheet${this.worksheets.items.length + 1}`);
    const rows = String(text).trim().split(/\r?\n/).map((line) => line.split(",").map((cell) => cell.trim()));
    sheet.getRangeByIndexes(0, 0, rows.length, rows[0]?.length || 1).values = rows;
    return this;
  }

  recalculate() {
    for (const sheet of this.worksheets) sheet.recalculate();
  }

  inspect(options = {}) {
    this.recalculate();
    const kinds = normalizeKinds(options.kind, ["workbook", "sheet", "table", "formula"]);
    const records = [];
    if (kinds.has("workbook")) records.push({ kind: "workbook", id: this.id, sheets: this.worksheets.items.length });
    for (const sheet of this.worksheets) {
      if (kinds.has("sheet")) records.push({ kind: "sheet", id: sheet.id, name: sheet.name, rows: sheet.usedBounds().rowCount, cols: sheet.usedBounds().colCount });
      if (kinds.has("table") || kinds.has("region")) records.push(sheet.tableRecord(options));
      if (kinds.has("table")) records.push(...sheet.tables.inspectRecords());
      if (kinds.has("drawing") || kinds.has("chart")) records.push(...sheet.charts.inspectRecords());
      if (kinds.has("drawing") || kinds.has("image")) records.push(...sheet.images.inspectRecords());
      if (kinds.has("sparkline") || kinds.has("drawing")) records.push(...sheet.sparklineGroups.inspectRecords());
      if (kinds.has("formula")) records.push(...sheet.formulaRecords(options));
      if (kinds.has("match")) records.push(...sheet.matchRecords(options));
      if (kinds.has("dataValidation")) records.push(...sheet.dataValidations.inspectRecords());
      if (kinds.has("conditionalFormat")) records.push(...sheet.conditionalFormattings.inspectRecords());
    }
    if (kinds.has("thread")) records.push(...this.comments.threads.map((thread) => thread.inspectRecord()));
    const search = String(options.search || options.searchTerm || "").trim().toLowerCase();
    const filtered = search ? records.filter((record) => JSON.stringify(record).toLowerCase().includes(search)) : records;
    return ndjson(filtered.filter(Boolean), options.maxChars ?? Infinity);
  }

  verify(options = {}) {
    this.recalculate();
    const issues = [];
    if (this.worksheets.items.length === 0) issues.push(verificationIssue("workbook", "noSheets", "Workbook has no worksheets."));
    for (const sheet of this.worksheets) {
      const entries = sheet.store.entries();
      if (entries.length === 0) issues.push(verificationIssue("workbook", "emptySheet", `Worksheet ${sheet.name} has no populated cells.`, { severity: "warning", sheet: sheet.name }));
      for (const [address, cell] of entries) {
        const value = String(cell.value ?? "");
        if (/^#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A)/.test(value)) {
          issues.push(verificationIssue("workbook", "formulaError", `Formula error ${value} at ${sheet.name}!${address}.`, { sheet: sheet.name, address, value }));
        }
      }
      for (const table of sheet.tables.items) {
        if (!table.name) issues.push(verificationIssue("workbook", "unnamedTable", `A worksheet table on ${sheet.name} is missing a stable name.`, { sheet: sheet.name, id: table.id }));
      }
      for (const chart of sheet.charts.items) {
        if (!chart.title) issues.push(verificationIssue("workbook", "untitledChart", `Chart ${chart.name} on ${sheet.name} has no title.`, { severity: "warning", sheet: sheet.name, id: chart.id }));
        if (chart.series.items.length === 0) issues.push(verificationIssue("workbook", "emptyChart", `Chart ${chart.name} on ${sheet.name} has no data series.`, { sheet: sheet.name, id: chart.id }));
      }
    }
    for (const thread of this.comments.threads) {
      if (!thread.target?.address) issues.push(verificationIssue("workbook", "unanchoredComment", `Comment thread ${thread.id} is missing a target cell.`, { id: thread.id }));
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
      for (const ref of formulaReferences(cell.formula)) {
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
    const thread = this.comments.getItem(id);
    if (thread) return thread;
    for (const sheet of this.worksheets) {
      if (sheet.id === id) return sheet;
      const table = sheet.tables.items.find((item) => item.id === id);
      if (table) return table;
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

  async render(options = {}) {
    const sheet = this.worksheets.getItem(options.sheetName) || this.worksheets.getActiveWorksheet();
    return new FileBlob(sheet.toSvg(options), { type: "image/svg+xml" });
  }
}

export class Worksheet {
  constructor(workbook, name) {
    this.workbook = workbook;
    this.id = aid("ws");
    this.name = name;
    this.store = new CellStore();
    this.charts = new WorksheetChartCollection(this);
    this.shapes = [];
    this.images = new WorksheetImageCollection(this);
    this.tables = new WorksheetTableCollection(this);
    this.sparklineGroups = new SparklineGroupCollection(this);
    this.sparklines = this.sparklineGroups;
    this.dataValidations = new WorksheetRuleCollection(this, "dataValidation");
    this.conditionalFormattings = new WorksheetRuleCollection(this, "conditionalFormat");
    this.freezePanes = { freezeRows() {}, freezeColumns() {}, unfreeze() {} };
    this.showGridLines = true;
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

  getUsedRange() {
    return new Range(this, this.usedBounds());
  }

  deleteAllDrawings() {
    this.charts.deleteAll();
    this.images.deleteAll();
    this.shapes = [];
  }

  usedBounds() {
    const coords = this.store.entries().map(([address]) => parseCellAddress(address));
    if (coords.length === 0) return { top: 0, left: 0, bottom: 0, right: 0, rowCount: 1, colCount: 1 };
    const top = Math.min(...coords.map((c) => c.row));
    const left = Math.min(...coords.map((c) => c.col));
    const bottom = Math.max(...coords.map((c) => c.row));
    const right = Math.max(...coords.map((c) => c.col));
    return { top, left, bottom, right, rowCount: bottom - top + 1, colCount: right - left + 1 };
  }

  recalculate() {
    for (const [address, cell] of this.store.entries()) {
      if (cell.formula) cell.value = evaluateFormula(this, cell.formula, address);
    }
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
      records.push({ kind: "formula", sheet: this.name, address, formula: cell.formula, value: cell.value });
    }
    return records;
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

  toSvg() {
    this.recalculate();
    const bounds = this.usedBounds();
    const cellW = 96;
    const cellH = 28;
    const imageFrames = this.images.items.map((image) => image.position);
    const sparklineFrames = this.sparklineGroups.items.map((group) => group.targetFrame(bounds));
    const width = Math.max(320, bounds.colCount * cellW + 80, ...this.charts.items.map((chart) => chart.position.left + chart.position.width + 40), ...imageFrames.map((frame) => frame.left + frame.width + 40), ...sparklineFrames.map((frame) => frame.left + frame.width + 40));
    const height = Math.max(180, bounds.rowCount * cellH + 80, ...this.charts.items.map((chart) => chart.position.top + chart.position.height + 40), ...imageFrames.map((frame) => frame.top + frame.height + 40), ...sparklineFrames.map((frame) => frame.top + frame.height + 40));
    const rows = [];
    for (let r = bounds.top; r <= bounds.bottom; r++) {
      for (let c = bounds.left; c <= bounds.right; c++) {
        const address = makeCellAddress(r, c);
        const value = this.store.get(address).value ?? "";
        const x = 40 + (c - bounds.left) * cellW;
        const y = 40 + (r - bounds.top) * cellH;
        rows.push(`<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" fill="white" stroke="#d0d7de"/>`);
        rows.push(`<text x="${x + 6}" y="${y + 18}" font-family="Arial" font-size="13" fill="#24292f">${xmlEscape(value)}</text>`);
      }
    }
    const tableOverlays = this.tables.items.map((table) => table.toSvg(bounds)).join("");
    const chartOverlays = this.charts.items.map((chart) => chart.toSvg()).join("");
    const imageOverlays = this.images.items.map((image) => image.toSvg()).join("");
    const sparklineOverlays = this.sparklineGroups.items.map((group) => group.toSvg(bounds)).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="#f6f8fa"/>${rows.join("")}${tableOverlays}${chartOverlays}${imageOverlays}${sparklineOverlays}</svg>`;
  }
}

export class Range {
  constructor(worksheet, bounds) {
    this.worksheet = worksheet;
    this.bounds = bounds;
  }

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
      for (let c = this.bounds.left; c <= this.bounds.right; c++) row.push(this.worksheet.store.get(makeCellAddress(r, c)).formula);
      out.push(row);
    }
    return out;
  }

  set formulas(matrix) {
    this.writeMatrix(matrix, "formula");
  }

  writeMatrix(matrix, field) {
    const rows = Array.isArray(matrix) ? matrix : [[matrix]];
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < (rows[r]?.length ?? 0); c++) {
        const address = makeCellAddress(this.bounds.top + r, this.bounds.left + c);
        const cell = this.worksheet.store.get(address);
        cell[field] = rows[r][c];
      }
    }
    this.worksheet.recalculate();
  }

  clear() {
    for (let r = this.bounds.top; r <= this.bounds.bottom; r++) for (let c = this.bounds.left; c <= this.bounds.right; c++) this.worksheet.store.cells.delete(makeCellAddress(r, c));
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

  merge() {}
  unmerge() {}
  fillDown() {}
  fillRight() {}
  getCell(row, col) { return this.worksheet.getCell(this.bounds.top + row, this.bounds.left + col); }
  getRange(address) { return this.worksheet.getRange(address); }
  getOffsetRange(rowOffset, colOffset) { return this.worksheet.getRangeByIndexes(this.bounds.top + rowOffset, this.bounds.left + colOffset, this.bounds.rowCount, this.bounds.colCount); }
  getResizedRange(rowCount, colCount) { return this.worksheet.getRangeByIndexes(this.bounds.top, this.bounds.left, rowCount, colCount); }
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

function formulaReferences(formula) {
  const raw = String(formula || "");
  const refs = [];
  const rangeRegex = /(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]+\$?\d+)\s*:\s*(\$?[A-Za-z]+\$?\d+)/g;
  const consumed = [];
  for (const match of raw.matchAll(rangeRegex)) {
    consumed.push(match.index);
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
    if (consumed.some((index) => match.index >= index && match.index <= index + 64)) continue;
    refs.push({ sheetName: match[1] || match[2] || undefined, address: match[3].replaceAll("$", "").toUpperCase() });
  }
  const seen = new Set();
  return refs.filter((ref) => {
    const key = `${ref.sheetName || ""}!${ref.address}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function evaluateFormula(sheet, formula) {
  const raw = String(formula || "").trim();
  if (!raw.startsWith("=")) return raw;
  const expr = raw.slice(1);
  const sumMatch = /^SUM\(([^)]+)\)$/i.exec(expr);
  if (sumMatch) {
    const range = sheet.getRange(sumMatch[1]);
    return range.values.flat().reduce((acc, value) => acc + (Number(value) || 0), 0);
  }
  const safe = expr.replace(/\b\$?[A-Za-z]+\$?\d+\b/g, (ref) => Number(sheet.store.get(ref.replaceAll("$", "").toUpperCase()).value) || 0);
  if (!/^[0-9+\-*/().\s]+$/.test(safe)) return `#NAME?`;
  try {
    return Function(`"use strict"; return (${safe});`)();
  } catch {
    return "#VALUE!";
  }
}

function workbookMetadata(workbook) {
  return {
    version: 1,
    comments: workbook.comments.toJSON(),
    sheets: workbook.worksheets.items.map((sheet) => ({
      name: sheet.name,
      dataValidations: sheet.dataValidations.toJSON(),
      conditionalFormattings: sheet.conditionalFormattings.toJSON(),
      tables: sheet.tables.toJSON(),
      charts: sheet.charts.toJSON(),
      images: sheet.images.toJSON(),
      sparklineGroups: sheet.sparklineGroups.toJSON(),
    })),
  };
}

function applyWorkbookMetadata(workbook, metadata = {}) {
  if (metadata.comments?.self) workbook.comments.setSelf(metadata.comments.self);
  for (const threadData of metadata.comments?.threads || []) {
    const thread = new CommentThread(workbook, threadData.target, threadData.comments?.[0]?.text || "", threadData.author);
    thread.id = threadData.id || thread.id;
    thread.comments = threadData.comments || thread.comments;
    thread.resolved = Boolean(threadData.resolved);
    workbook.comments.threads.push(thread);
  }
  for (const sheetData of metadata.sheets || []) {
    const sheet = workbook.worksheets.getItem(sheetData.name);
    if (!sheet) continue;
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
}

export class SpreadsheetFile {
  static async exportXlsx(workbook) {
    workbook.recalculate();
    const zip = new JSZip();
    const tableParts = collectWorkbookTableParts(workbook);
    const imageParts = collectWorkbookImageParts(workbook);
    const chartParts = collectWorkbookChartParts(workbook, imageParts);
    zip.file("[Content_Types].xml", xlsxContentTypes(workbook.worksheets.items.length, tableParts, imageParts, chartParts));
    zip.file("_rels/.rels", relsXml([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "xl/workbook.xml" }]));
    zip.file("xl/workbook.xml", workbookXml(workbook));
    zip.file("xl/_rels/workbook.xml.rels", workbookRelsXml(workbook.worksheets.items.length));
    zip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>`);
    zip.file("customXml/open-office-artifact.json", JSON.stringify(workbookMetadata(workbook), null, 2));
    workbook.worksheets.items.forEach((sheet, index) => {
      const sheetTableParts = tableParts.filter((part) => part.sheetIndex === index);
      const sheetImageParts = imageParts.filter((part) => part.sheetIndex === index);
      const sheetChartParts = chartParts.filter((part) => part.sheetIndex === index);
      const drawingRelId = sheetImageParts.length || sheetChartParts.length ? `rId${sheetTableParts.length + 1}` : undefined;
      zip.file(`xl/worksheets/sheet${index + 1}.xml`, worksheetXml(sheet, sheetTableParts, drawingRelId));
      if (sheetTableParts.length || sheetImageParts.length || sheetChartParts.length) zip.file(`xl/worksheets/_rels/sheet${index + 1}.xml.rels`, worksheetRelsXml(sheetTableParts, drawingRelId ? { relId: drawingRelId, target: `../drawings/drawing${index + 1}.xml` } : undefined));
      if (sheetImageParts.length || sheetChartParts.length) {
        zip.file(`xl/drawings/drawing${index + 1}.xml`, drawingXml(sheetImageParts, sheetChartParts));
        zip.file(`xl/drawings/_rels/drawing${index + 1}.xml.rels`, drawingRelsXml(sheetImageParts, sheetChartParts));
      }
    });
    tableParts.forEach((part) => zip.file(`xl/tables/table${part.tablePartId}.xml`, tableXml(part.table, part.tablePartId)));
    imageParts.forEach((part) => zip.file(`xl/media/image${part.imagePartId}.${part.extension}`, part.bytes));
    chartParts.forEach((part) => zip.file(`xl/charts/chart${part.chartPartId}.xml`, xlsxChartXml(part.chart)));
    const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return new FileBlob(bytes, { type: XLSX_MIME });
  }

  static async importXlsx(blobOrBuffer) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const workbook = Workbook.create();
    const workbookText = await zip.file("xl/workbook.xml")?.async("text");
    const sheetNames = [...String(workbookText || "").matchAll(/<sheet[^>]*name="([^"]+)"[^>]*sheetId="(\d+)"/g)].map((m) => ({ name: decodeXml(m[1]), index: Number(m[2]) }));
    for (const { name, index } of sheetNames.length ? sheetNames : [{ name: "Sheet1", index: 1 }]) {
      const sheet = workbook.worksheets.add(name);
      const xml = await zip.file(`xl/worksheets/sheet${index}.xml`)?.async("text");
      if (xml) parseWorksheetXml(sheet, xml);
    }
    const metadataText = await zip.file("customXml/open-office-artifact.json")?.async("text");
    if (metadataText) applyWorkbookMetadata(workbook, JSON.parse(metadataText));
    workbook.recalculate();
    return workbook;
  }
}

function collectWorkbookTableParts(workbook) {
  const parts = [];
  let tablePartId = 1;
  workbook.worksheets.items.forEach((sheet, sheetIndex) => {
    sheet.tables.items.forEach((table, tableIndex) => {
      parts.push({ sheet, sheetIndex, table, tableIndex, tablePartId: tablePartId++, relId: `rId${tableIndex + 1}` });
    });
  });
  return parts;
}

function imageDataFromDataUrl(dataUrl) {
  const match = /^data:([^;,]+);base64,(.+)$/i.exec(String(dataUrl || ""));
  if (!match) return undefined;
  const contentType = match[1].toLowerCase();
  const extension = contentType === "image/jpeg" ? "jpg" : contentType.split("/")[1] || "bin";
  return { contentType, extension: extension === "svg+xml" ? "svg" : extension, bytes: Buffer.from(match[2], "base64") };
}

function imageContentTypeDefaults(imageParts = []) {
  return [
    ["png", "image/png"],
    ["jpg", "image/jpeg"],
    ["jpeg", "image/jpeg"],
    ["gif", "image/gif"],
    ["svg", "image/svg+xml"],
  ].filter(([extension]) => imageParts.some((part) => part.extension === extension)).map(([extension, contentType]) => `<Default Extension="${extension}" ContentType="${contentType}"/>`).join("");
}

function imageContentTypeFromExtension(extension) {
  const normalized = String(extension || "").toLowerCase().replace(/^\./, "");
  if (normalized === "jpg" || normalized === "jpeg") return "image/jpeg";
  if (normalized === "png") return "image/png";
  if (normalized === "gif") return "image/gif";
  if (normalized === "svg" || normalized === "svg+xml") return "image/svg+xml";
  return "application/octet-stream";
}

function collectWorkbookImageParts(workbook) {
  const parts = [];
  let imagePartId = 1;
  workbook.worksheets.items.forEach((sheet, sheetIndex) => {
    let drawingRelIndex = 1;
    sheet.images.items.forEach((image, imageIndex) => {
      const data = imageDataFromDataUrl(image.dataUrl);
      if (!data) return;
      parts.push({
        sheet,
        sheetIndex,
        image,
        imageIndex,
        imagePartId: imagePartId++,
        drawingRelId: `rId${drawingRelIndex++}`,
        ...data,
      });
    });
  });
  return parts;
}

function collectWorkbookChartParts(workbook, imageParts = []) {
  const parts = [];
  let chartPartId = 1;
  workbook.worksheets.items.forEach((sheet, sheetIndex) => {
    let drawingRelIndex = imageParts.filter((part) => part.sheetIndex === sheetIndex).length + 1;
    sheet.charts.items.forEach((chart, chartIndex) => {
      parts.push({ sheet, sheetIndex, chart, chartIndex, chartPartId: chartPartId++, drawingRelId: `rId${drawingRelIndex++}` });
    });
  });
  return parts;
}

function xlsxContentTypes(sheetCount, tableParts = [], imageParts = [], chartParts = []) {
  const sheets = Array.from({ length: sheetCount }, (_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("");
  const tables = tableParts.map((part) => `<Override PartName="/xl/tables/table${part.tablePartId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`).join("");
  const imageDefaults = [
    ["png", "image/png"],
    ["jpg", "image/jpeg"],
    ["jpeg", "image/jpeg"],
    ["gif", "image/gif"],
    ["svg", "image/svg+xml"],
  ].filter(([extension]) => imageParts.some((part) => part.extension === extension)).map(([extension, contentType]) => `<Default Extension="${extension}" ContentType="${contentType}"/>`).join("");
  const charts = chartParts.map((part) => `<Override PartName="/xl/charts/chart${part.chartPartId}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="json" ContentType="application/json"/>${imageDefaults}<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets}${tables}${charts}</Types>`;
}

function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map((rel) => `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${attrEscape(rel.target)}"${rel.targetMode ? ` TargetMode="${attrEscape(rel.targetMode)}"` : ""}/>`).join("")}</Relationships>`;
}

function parseRelsXml(xml) {
  return [...String(xml || "").matchAll(/<Relationship\b([^>]*)\/>/g)].map((match) => {
    const attrs = match[1] || "";
    return {
      id: decodeXml(/\bId="([^"]*)"/.exec(attrs)?.[1] || ""),
      type: decodeXml(/\bType="([^"]*)"/.exec(attrs)?.[1] || ""),
      target: decodeXml(/\bTarget="([^"]*)"/.exec(attrs)?.[1] || ""),
      targetMode: decodeXml(/\bTargetMode="([^"]*)"/.exec(attrs)?.[1] || ""),
    };
  }).filter((rel) => rel.id);
}

function workbookXml(workbook) {
  const sheets = workbook.worksheets.items.map((sheet, i) => `<sheet name="${attrEscape(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets}</sheets></workbook>`;
}

function workbookRelsXml(sheetCount) {
  const rels = Array.from({ length: sheetCount }, (_, i) => ({ id: `rId${i + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", target: `worksheets/sheet${i + 1}.xml` }));
  rels.push({ id: `rId${sheetCount + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", target: "styles.xml" });
  return relsXml(rels);
}

function worksheetRelsXml(tableParts, drawingRel) {
  const rels = tableParts.map((part) => ({
    id: part.relId,
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
    target: `../tables/table${part.tablePartId}.xml`,
  }));
  if (drawingRel) rels.push({ id: drawingRel.relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", target: drawingRel.target });
  return relsXml(rels);
}

function drawingRelsXml(imageParts, chartParts = []) {
  return relsXml([
    ...imageParts.map((part) => ({
      id: part.drawingRelId,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      target: `../media/image${part.imagePartId}.${part.extension}`,
    })),
    ...chartParts.map((part) => ({
      id: part.drawingRelId,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
      target: `../charts/chart${part.chartPartId}.xml`,
    })),
  ]);
}

function drawingXml(imageParts, chartParts = []) {
  const imageAnchors = imageParts.map((part, index) => {
    const from = part.image.anchor?.from || { row: 0, col: 0 };
    const extent = part.image.anchor?.extent || {};
    const widthPx = Number(extent.widthPx || part.image.anchor?.widthPx || 160);
    const heightPx = Number(extent.heightPx || part.image.anchor?.heightPx || 120);
    const cx = Math.round(widthPx * 9525);
    const cy = Math.round(heightPx * 9525);
    return `<xdr:oneCellAnchor><xdr:from><xdr:col>${Number(from.col || 0)}</xdr:col><xdr:colOff>${Math.round(Number(from.colOffsetPx || 0) * 9525)}</xdr:colOff><xdr:row>${Number(from.row || 0)}</xdr:row><xdr:rowOff>${Math.round(Number(from.rowOffsetPx || 0) * 9525)}</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${index + 2}" name="${attrEscape(part.image.name || `Image ${index + 1}`)}" descr="${attrEscape(part.image.alt || "")}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${part.drawingRelId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
  }).join("");
  const chartAnchors = chartParts.map((part, index) => {
    const p = part.chart.position;
    const col = Math.max(0, Math.floor((p.left - 40) / 96));
    const row = Math.max(0, Math.floor((p.top - 40) / 28));
    const colOff = Math.max(0, Math.round((p.left - 40 - col * 96) * 9525));
    const rowOff = Math.max(0, Math.round((p.top - 40 - row * 28) * 9525));
    const cx = Math.round(p.width * 9525);
    const cy = Math.round(p.height * 9525);
    const id = imageParts.length + index + 2;
    return `<xdr:oneCellAnchor><xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${attrEscape(part.chart.name || `Chart ${index + 1}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="${part.drawingRelId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${imageAnchors}${chartAnchors}</xdr:wsDr>`;
}

function sanitizeTableColumnName(value, index, seen) {
  const base = String(value ?? "").trim() || `Column${index + 1}`;
  let name = base;
  let suffix = 2;
  while (seen.has(name)) name = `${base}_${suffix++}`;
  seen.add(name);
  return name;
}

function tableXml(table, tablePartId) {
  const ref = table.range;
  const seen = new Set();
  const headers = table.showHeaders && table.values[0]
    ? table.values[0]
    : Array.from({ length: table.columnCount || 1 }, (_, index) => `Column${index + 1}`);
  const columns = Array.from({ length: table.columnCount || headers.length || 1 }, (_, index) => {
    const name = sanitizeTableColumnName(headers[index], index, seen);
    return `<tableColumn id="${index + 1}" name="${attrEscape(name)}"/>`;
  }).join("");
  const headerRowCount = table.showHeaders ? 1 : 0;
  const totalsRowShown = table.showTotals ? 1 : 0;
  const autoFilter = table.showFilterButton ? `<autoFilter ref="${attrEscape(ref)}"/>` : "";
  const styleName = table.style || "TableStyleMedium2";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${tablePartId}" name="${attrEscape(table.name)}" displayName="${attrEscape(table.name)}" ref="${attrEscape(ref)}" headerRowCount="${headerRowCount}" totalsRowShown="${totalsRowShown}">${autoFilter}<tableColumns count="${table.columnCount || headers.length || 1}">${columns}</tableColumns><tableStyleInfo name="${attrEscape(styleName)}" showFirstColumn="0" showLastColumn="0" showRowStripes="${table.showHeaders ? 1 : 0}" showColumnStripes="${table.showBandedColumns ? 1 : 0}"/></table>`;
}

function xlsxChartXml(chart) {
  const chartType = chart.type || chart.chartType || "bar";
  const chartElementName = chartType === "line" ? "lineChart" : "barChart";
  const grouping = chartType === "line" ? "<c:grouping val=\"standard\"/>" : "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>";
  const seriesItems = chart.series?.items || chart.series || [];
  const seriesXml = (seriesItems.length ? seriesItems : [{ name: chart.title || "Series", values: [] }]).map((series, index) => {
    const values = series.values || [];
    const categories = chart.categories || values.map((_, i) => String(i + 1));
    const catPts = categories.map((category, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${xmlEscape(category)}</c:v></c:pt>`).join("");
    const valPts = values.map((value, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${Number(value) || 0}</c:v></c:pt>`).join("");
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:v>${xmlEscape(series.name || `Series ${index + 1}`)}</c:v></c:tx><c:cat><c:strLit><c:ptCount val="${categories.length}"/>${catPts}</c:strLit></c:cat><c:val><c:numLit><c:ptCount val="${values.length}"/>${valPts}</c:numLit></c:val></c:ser>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(chart.title || chartType)}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:layout/><c:${chartElementName}>${grouping}${seriesXml}<c:axId val="1"/><c:axId val="2"/></c:${chartElementName}><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:layout/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

function dataValidationsXml(sheet) {
  if (sheet.dataValidations.items.length === 0) return "";
  const rules = sheet.dataValidations.items.map((item) => {
    const rule = item.rule || item;
    const type = rule.type || "custom";
    const operator = rule.operator ? ` operator="${attrEscape(rule.operator)}"` : "";
    const valuesFormula = Array.isArray(rule.values) ? `"${rule.values.map((value) => String(value).replaceAll('"', '""')).join(",")}"` : undefined;
    const formula1 = rule.formula1 ?? valuesFormula;
    const formula2 = rule.formula2;
    return `<dataValidation type="${attrEscape(type)}"${operator} allowBlank="1" sqref="${attrEscape(item.range || "A1")}">${formula1 != null ? `<formula1>${xmlEscape(formula1)}</formula1>` : ""}${formula2 != null ? `<formula2>${xmlEscape(formula2)}</formula2>` : ""}</dataValidation>`;
  }).join("");
  return `<dataValidations count="${sheet.dataValidations.items.length}">${rules}</dataValidations>`;
}

function conditionalFormattingXml(sheet) {
  return sheet.conditionalFormattings.items.map((item, index) => {
    const ruleType = item.ruleType || "expression";
    const type = ruleType === "cellIs" || ruleType === "CellValue" ? "cellIs" : ruleType === "containsText" ? "containsText" : "expression";
    const operator = item.operator ? ` operator="${attrEscape(item.operator)}"` : "";
    const formula = Array.isArray(item.formula) ? item.formula[0] : item.formula || item.expression || "TRUE";
    return `<conditionalFormatting sqref="${attrEscape(item.range || "A1")}"><cfRule type="${attrEscape(type)}" priority="${index + 1}"${operator}><formula>${xmlEscape(formula)}</formula></cfRule></conditionalFormatting>`;
  }).join("");
}

function worksheetXml(sheet, tableParts = [], drawingRelId) {
  const rows = new Map();
  for (const [address, cell] of sheet.store.entries()) {
    const { row, col } = parseCellAddress(address);
    if (!rows.has(row)) rows.set(row, []);
    rows.get(row).push({ address, col, cell });
  }
  const rowXml = [...rows.entries()].sort((a, b) => a[0] - b[0]).map(([row, cells]) => `<row r="${row + 1}">${cells.sort((a, b) => a.col - b.col).map(({ address, cell }) => cellXml(address, cell)).join("")}</row>`).join("");
  const tablePartsXml = tableParts.length ? `<tableParts count="${tableParts.length}">${tableParts.map((part) => `<tablePart r:id="${part.relId}"/>`).join("")}</tableParts>` : "";
  const drawingXml = drawingRelId ? `<drawing r:id="${drawingRelId}"/>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>${rowXml}</sheetData>${conditionalFormattingXml(sheet)}${dataValidationsXml(sheet)}${drawingXml}${tablePartsXml}</worksheet>`;
}

function cellXml(address, cell) {
  const f = cell.formula ? `<f>${xmlEscape(String(cell.formula).replace(/^=/, ""))}</f>` : "";
  if (typeof cell.value === "number") return `<c r="${address}">${f}<v>${cell.value}</v></c>`;
  if (typeof cell.value === "boolean") return `<c r="${address}" t="b">${f}<v>${cell.value ? 1 : 0}</v></c>`;
  if (cell.value == null && !f) return "";
  if (cell.value == null) return `<c r="${address}">${f}</c>`;
  return `<c r="${address}" t="inlineStr">${f}<is><t>${xmlEscape(cell.value)}</t></is></c>`;
}

function parseWorksheetXml(sheet, xml) {
  for (const match of xml.matchAll(/<c\s+([^>]*)>([\s\S]*?)<\/c>/g)) {
    const attrs = match[1];
    const body = match[2];
    const address = /r="([^"]+)"/.exec(attrs)?.[1];
    if (!address) continue;
    const cell = sheet.store.get(address);
    const formula = /<f[^>]*>([\s\S]*?)<\/f>/.exec(body)?.[1];
    if (formula) cell.formula = `=${decodeXml(formula)}`;
    const text = /<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/is>/.exec(body)?.[1];
    const value = /<v[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1];
    if (text !== undefined) cell.value = decodeXml(text);
    else if (value !== undefined) cell.value = Number.isFinite(Number(value)) ? Number(value) : decodeXml(value);
  }
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

export class Presentation {
  constructor(options = {}) {
    this.id = aid("pr");
    this.slideSize = options.slideSize || { width: 1280, height: 720 };
    this.slides = new SlideCollection(this);
  }

  static create(options = {}) { return new Presentation(options); }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["deck", "slide", "textbox", "shape", "layout"]);
    const records = [];
    if (kinds.has("deck")) records.push({ kind: "deck", id: this.id, slides: this.slides.count });
    for (const slide of this.slides) records.push(...slide.inspectRecords(kinds));
    const search = String(options.search || "").trim().toLowerCase();
    const filtered = search
      ? records.filter((record) => JSON.stringify(record).toLowerCase().includes(search))
      : records;
    return ndjson(filtered, options.maxChars ?? Infinity);
  }

  validateLayout(options = {}) {
    const issues = this.slides.items.flatMap((slide) => slide.validateLayout(options).issues);
    return { ok: issues.length === 0, issues, ...ndjson(issues, options.maxChars ?? Infinity) };
  }

  verify(options = {}) {
    const issues = [];
    if (this.slides.items.length === 0) issues.push(verificationIssue("presentation", "noSlides", "Presentation has no slides."));
    issues.push(...this.validateLayout(options).issues.map((issue) => ({ ...issue, artifactKind: "presentation" })));
    return verificationResult("presentation", issues, options);
  }

  resolve(id) {
    if (id === this.id) return this;
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
    const slide = options.slide || this.slides.getItem(0) || this.slides.add();
    if (options.format === "layout") return slide.export({ format: "layout" });
    return slide.export(options);
  }

  toProto() {
    return { id: this.id, slideSize: this.slideSize, slides: this.slides.items.map((slide) => slide.toProto()) };
  }
}

class ShapeCollection {
  constructor(slide) { this.slide = slide; this.items = []; }
  add(config = {}) { const shape = new Shape(this.slide, config); this.items.push(shape); return shape; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class ElementCollection {
  constructor(slide, ElementClass) { this.slide = slide; this.ElementClass = ElementClass; this.items = []; }
  add(...args) { const element = new this.ElementClass(this.slide, ...args); this.items.push(element); return element; }
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
  const fontSize = element.text.style.fontSize || 24;
  const lineSpacing = element.text.style.lineSpacing || 1.2;
  const availableWidth = Math.max(1, frame.width - 18);
  const charsPerLine = Math.max(1, Math.floor(availableWidth / (fontSize * 0.55)));
  const requiredLines = String(text).split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);
  const requiredHeight = requiredLines * fontSize * lineSpacing + 12;
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

function tableOverflowIssues(slide, tableElement) {
  const issues = [];
  const frame = tableElement.position;
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

export class Slide {
  constructor(presentation, options = {}) {
    this.presentation = presentation;
    this.id = aid("sl");
    this.name = options.name || "";
    this.shapes = new ShapeCollection(this);
    this.images = new ElementCollection(this, ImageElement);
    this.tables = new ElementCollection(this, TableElement);
    this.charts = new ElementCollection(this, ChartElement);
    this.speakerNotes = { text: "" };
    this.background = { fill: "white" };
  }

  get index() { return this.presentation.slides.items.indexOf(this); }
  get frame() { return { left: 0, top: 0, ...this.presentation.slideSize }; }

  inspectRecords(kinds) {
    const records = [];
    if (kinds.has("layout")) records.push({ kind: "layout", layoutId: `${this.id}/layout`, name: "Blank", type: "blank" });
    if (kinds.has("slide")) records.push({ kind: "slide", id: this.id, slide: this.index + 1, title: this.title(), textShapes: this.shapes.items.filter((s) => s.text.value).length, tables: this.tables.items.length, charts: this.charts.items.length, images: this.images.items.length });
    for (const shape of this.shapes) {
      if (kinds.has("textbox") && shape.text.value) records.push(shape.inspectRecord("textbox"));
      else if (kinds.has("shape")) records.push(shape.inspectRecord("shape"));
    }
    if (kinds.has("table")) records.push(...this.tables.items.map((table) => table.inspectRecord()));
    if (kinds.has("chart")) records.push(...this.charts.items.map((chart) => chart.inspectRecord()));
    if (kinds.has("image")) records.push(...this.images.items.map((image) => image.inspectRecord()));
    return records;
  }

  title() { return this.shapes.items.find((shape) => shape.text.value)?.text.value || this.charts.items[0]?.title || ""; }
  resolve(id) { return [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items].find((element) => element.id === id); }

  validateLayout(options = {}) {
    const issues = [];
    const slideFrame = this.frame;
    const elements = [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items];
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
    if (options.format === "layout") return new FileBlob(JSON.stringify(this.layoutJson(), null, 2), { type: LAYOUT_MIME });
    return new FileBlob(this.toSvg(), { type: "image/svg+xml" });
  }

  layoutJson() {
    return {
      schema: "open-office-artifact.layout/v1",
      unit: "px",
      slide: { id: this.id, slide: this.index + 1, frame: this.frame },
      elements: [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items].map((element) => element.layoutJson()),
    };
  }

  toSvg() {
    const { width, height } = this.presentation.slideSize;
    const elements = [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items].map((element) => element.toSvg()).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="${xmlEscape(this.background.fill || "white")}"/>${elements}</svg>`;
  }

  toProto() { return { id: this.id, elements: [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items].map((element) => element.layoutJson()) }; }

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

class TextFrame {
  constructor(text = "") { this.value = text; this.style = {}; }
  set(text) { this.value = String(text ?? ""); }
  replace(search, replacement) { this.value = this.value.replace(search, replacement); }
  toString() { return this.value; }
}

export class Shape {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("sh");
    this.geometry = config.geometry || "rect";
    this.name = config.name || "";
    this.position = config.position || { left: 0, top: 0, width: 160, height: 80 };
    this.fill = config.fill || "transparent";
    this.line = config.line || { fill: "#334155", width: 1 };
    this.borderRadius = config.borderRadius;
    this._text = new TextFrame(config.text || "");
  }

  get text() { return this._text; }
  set text(value) { this._text.set(value); }

  inspectRecord(kind = "shape") {
    const p = this.position;
    return { kind, id: this.id, slide: this.slide.index + 1, name: this.name || undefined, text: this.text.value || undefined, textPreview: this.text.value || undefined, textChars: this.text.value.length || undefined, textLines: this.text.value ? this.text.value.split(/\r?\n/).length : undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px" };
  }

  layoutJson() { return { kind: this.text.value ? "textbox" : "shape", id: this.id, name: this.name, geometry: this.geometry, frame: this.position, text: this.text.value, style: { fill: this.fill, line: this.line, borderRadius: this.borderRadius, text: this.text.style } }; }

  toSvg() {
    const p = this.position;
    const fill = typeof this.fill === "string" ? resolveColorToken(this.fill, this.fill) : this.fill?.color || "transparent";
    const stroke = resolveColorToken(this.line?.fill || this.line?.color || "#334155", "#334155");
    const sw = this.line?.width ?? 1;
    const rect = this.geometry === "ellipse"
      ? `<ellipse cx="${p.left + p.width / 2}" cy="${p.top + p.height / 2}" rx="${p.width / 2}" ry="${p.height / 2}" fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}"/>`
      : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" rx="${this.borderRadius ? 12 : 0}" fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}"/>`;
    const text = this.text.value ? `<text x="${p.left + 12}" y="${p.top + 36}" font-family="Arial" font-size="${this.text.style.fontSize || 24}" font-weight="${this.text.style.bold ? "700" : "400"}" fill="${xmlEscape(this.text.style.color || "#0f172a")}">${xmlEscape(this.text.value)}</text>` : "";
    return rect + text;
  }

  toPptxShape(index) {
    return pptxTextShapeXml(index, this.name || this.id, this.geometry, this.position, this.text.value);
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
    return { kind: "table", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, rows: this.rows, cols: this.columns, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", values: this.values };
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

  toPptxShape(index) { return pptxTableXml(index, this); }
}

export class ChartElement {
  constructor(slide, chartType = "bar", config = {}) {
    this.slide = slide;
    this.id = config.id || aid("ch");
    this.name = config.name || "";
    this.chartType = chartType || config.chartType || "bar";
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 360, height: 220 });
    this.title = config.title || "";
    this.categories = config.categories || [];
    this.series = config.series || [];
    this.hasLegend = config.hasLegend ?? this.series.length > 1;
  }

  inspectRecord() {
    const p = this.position;
    return { kind: "chart", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, chartType: this.chartType, title: this.title, series: this.series.length, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px" };
  }

  layoutJson() { return { kind: "chart", id: this.id, name: this.name, chartType: this.chartType, title: this.title, frame: this.position, categories: this.categories, series: this.series }; }

  toSvg() {
    const p = this.position;
    const values = this.series[0]?.values || [];
    const max = Math.max(1, ...values.map((value) => Number(value) || 0));
    const plot = { left: p.left + 36, top: p.top + 42, width: Math.max(0, p.width - 56), height: Math.max(0, p.height - 72) };
    const barW = values.length ? plot.width / values.length * 0.65 : 0;
    const gap = values.length ? plot.width / values.length * 0.35 : 0;
    const bars = values.map((value, index) => {
      const h = plot.height * (Number(value) || 0) / max;
      const x = plot.left + index * (barW + gap) + gap / 2;
      const y = plot.top + plot.height - h;
      return `<rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="#0ea5e9"/>`;
    }).join("");
    const labels = this.categories.map((category, index) => `<text x="${plot.left + index * (barW + gap)}" y="${p.top + p.height - 12}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(category)}</text>`).join("");
    return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/><text x="${p.left + 12}" y="${p.top + 24}" font-family="Arial" font-size="16" font-weight="700" fill="#0f172a">${xmlEscape(this.title || this.chartType)}</text>${bars}${labels}`;
  }

  toPptxShape(index, relId) {
    if (relId) return pptxChartFrameXml(index, this.name || this.id, this.position, relId);
    return pptxTextShapeXml(index, this.name || this.id, "rect", this.position, `${this.title || this.chartType}\n${this.series.map((series) => `${series.name || "Series"}: ${(series.values || []).join(", ")}`).join("\n")}`);
  }
}

export class ImageElement {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("im");
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
  }

  get frame() { return this.position; }
  set frame(value) { this.position = normalizeFrame(value, this.position); }
  replace(config = {}) { Object.assign(this, config); }

  inspectRecord() {
    const p = this.position;
    return { kind: "image", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, alt: this.alt || undefined, prompt: this.prompt || undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", fit: this.fit };
  }

  layoutJson() { return { kind: "image", id: this.id, name: this.name, frame: this.position, alt: this.alt, prompt: this.prompt, uri: this.uri, dataUrl: this.dataUrl, fit: this.fit, geometry: this.geometry, borderRadius: this.borderRadius }; }

  toSvg() {
    const p = this.position;
    const label = this.alt || this.prompt || this.uri || "image";
    if (this.dataUrl) return `<image href="${attrEscape(this.dataUrl)}" x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" preserveAspectRatio="xMidYMid meet"/><text x="${p.left + 12}" y="${p.top + 28}" font-family="Arial" font-size="14" fill="#075985">${xmlEscape(label)}</text>`;
    const rect = this.geometry === "ellipse"
      ? `<ellipse cx="${p.left + p.width / 2}" cy="${p.top + p.height / 2}" rx="${p.width / 2}" ry="${p.height / 2}" fill="#e0f2fe" stroke="#0284c7"/>`
      : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" rx="${this.borderRadius ? 12 : 0}" fill="#e0f2fe" stroke="#0284c7"/>`;
    return `${rect}<text x="${p.left + 12}" y="${p.top + 28}" font-family="Arial" font-size="14" fill="#075985">${xmlEscape(label)}</text>`;
  }

  toPptxShape(index, relId) {
    if (!relId) return pptxTextShapeXml(index, this.name || this.id, this.geometry === "ellipse" ? "ellipse" : "rect", this.position, this.alt || this.prompt || "Image");
    return pptxPictureXml(index, this.name || this.id, this.alt || this.prompt || "", this.position, relId);
  }
}

export class PresentationFile {
  static async exportPptx(presentation) {
    const zip = new JSZip();
    const imageParts = collectPresentationImageParts(presentation);
    const chartParts = collectPresentationChartParts(presentation, imageParts);
    zip.file("[Content_Types].xml", pptxContentTypes(presentation.slides.count, imageParts, chartParts));
    zip.file("_rels/.rels", relsXml([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "ppt/presentation.xml" }]));
    zip.file("ppt/presentation.xml", presentationXml(presentation));
    zip.file("ppt/_rels/presentation.xml.rels", relsXml(presentation.slides.items.map((_, i) => ({ id: `rId${i + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", target: `slides/slide${i + 1}.xml` }))));
    presentation.slides.items.forEach((slide, i) => {
      const slideImageParts = imageParts.filter((part) => part.slideIndex === i);
      const slideChartParts = chartParts.filter((part) => part.slideIndex === i);
      zip.file(`ppt/slides/slide${i + 1}.xml`, slideXml(slide, slideImageParts, slideChartParts));
      if (slideImageParts.length || slideChartParts.length) zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, pptxSlideRelsXml(slideImageParts, slideChartParts));
    });
    imageParts.forEach((part) => zip.file(`ppt/media/image${part.imagePartId}.${part.extension}`, part.bytes));
    chartParts.forEach((part) => zip.file(`ppt/charts/chart${part.chartPartId}.xml`, pptxChartXml(part.chart)));
    const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return new FileBlob(bytes, { type: PPTX_MIME });
  }

  static async importPptx(blobOrBuffer) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const presentation = Presentation.create();
    const slideFiles = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort();
    for (const file of slideFiles) {
      const slide = presentation.slides.add();
      parseSlideXml(slide, await zip.file(file).async("text"));
    }
    return presentation;
  }
}

function collectPresentationImageParts(presentation) {
  const parts = [];
  let imagePartId = 1;
  presentation.slides.items.forEach((slide, slideIndex) => {
    let relIndex = 1;
    slide.images.items.forEach((image, imageIndex) => {
      const data = imageDataFromDataUrl(image.dataUrl);
      if (!data) return;
      parts.push({
        slide,
        slideIndex,
        image,
        imageIndex,
        imagePartId: imagePartId++,
        slideRelId: `rId${relIndex++}`,
        ...data,
      });
    });
  });
  return parts;
}

function collectPresentationChartParts(presentation, imageParts = []) {
  const parts = [];
  let chartPartId = 1;
  presentation.slides.items.forEach((slide, slideIndex) => {
    let relIndex = imageParts.filter((part) => part.slideIndex === slideIndex).length + 1;
    slide.charts.items.forEach((chart, chartIndex) => {
      parts.push({ slide, slideIndex, chart, chartIndex, chartPartId: chartPartId++, slideRelId: `rId${relIndex++}` });
    });
  });
  return parts;
}

function pptxContentTypes(slideCount, imageParts = [], chartParts = []) {
  const slides = Array.from({ length: slideCount }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("");
  const charts = chartParts.map((part) => `<Override PartName="/ppt/charts/chart${part.chartPartId}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join("");
  const imageDefaults = [
    ["png", "image/png"],
    ["jpg", "image/jpeg"],
    ["jpeg", "image/jpeg"],
    ["gif", "image/gif"],
    ["svg", "image/svg+xml"],
  ].filter(([extension]) => imageParts.some((part) => part.extension === extension)).map(([extension, contentType]) => `<Default Extension="${extension}" ContentType="${contentType}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${imageDefaults}<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${slides}${charts}</Types>`;
}

function pptxSlideRelsXml(imageParts, chartParts = []) {
  return relsXml([
    ...imageParts.map((part) => ({ id: part.slideRelId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: `../media/image${part.imagePartId}.${part.extension}` })),
    ...chartParts.map((part) => ({ id: part.slideRelId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", target: `../charts/chart${part.chartPartId}.xml` })),
  ]);
}

function presentationXml(presentation) {
  const ids = presentation.slides.items.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="12192000" cy="6858000"/><p:sldIdLst>${ids}</p:sldIdLst></p:presentation>`;
}

function pptxTextShapeXml(index, name, geometry, position, text = "") {
  const p = position;
  const x = Math.round(p.left * 9525), y = Math.round(p.top * 9525), cx = Math.round(p.width * 9525), cy = Math.round(p.height * 9525);
  const paragraphs = String(text || "").split(/\r?\n/).map((line) => `<a:p><a:r><a:t>${xmlEscape(line)}</a:t></a:r></a:p>`).join("");
  return `<p:sp><p:nvSpPr><p:cNvPr id="${index + 2}" name="${attrEscape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${attrEscape(geometry === "textbox" ? "rect" : geometry)}"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs || "<a:p/>"}</p:txBody></p:sp>`;
}

function pptxPictureXml(index, name, alt, position, relId) {
  const p = position;
  const x = Math.round(p.left * 9525), y = Math.round(p.top * 9525), cx = Math.round(p.width * 9525), cy = Math.round(p.height * 9525);
  return `<p:pic><p:nvPicPr><p:cNvPr id="${index + 2}" name="${attrEscape(name)}" descr="${attrEscape(alt)}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

function pptxTableXml(index, table) {
  const p = table.position;
  const x = Math.round(p.left * 9525), y = Math.round(p.top * 9525), cx = Math.round(p.width * 9525), cy = Math.round(p.height * 9525);
  const colWidth = Math.max(1, Math.floor(cx / Math.max(1, table.columns)));
  const rowHeight = Math.max(1, Math.floor(cy / Math.max(1, table.rows)));
  const grid = Array.from({ length: Math.max(1, table.columns) }, () => `<a:gridCol w="${colWidth}"/>`).join("");
  const rows = Array.from({ length: Math.max(1, table.rows) }, (_, rowIndex) => {
    const cells = Array.from({ length: Math.max(1, table.columns) }, (_, colIndex) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(table.values[rowIndex]?.[colIndex] ?? "")}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`).join("");
    return `<a:tr h="${rowHeight}">${cells}</a:tr>`;
  }).join("");
  const firstRow = table.styleOptions.headerRow ? 1 : 0;
  const bandRow = table.styleOptions.bandedRows ? 1 : 0;
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${index + 2}" name="${attrEscape(table.name || table.id)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="${firstRow}" bandRow="${bandRow}"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

function pptxChartFrameXml(index, name, position, relId) {
  const p = position;
  const x = Math.round(p.left * 9525), y = Math.round(p.top * 9525), cx = Math.round(p.width * 9525), cy = Math.round(p.height * 9525);
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${index + 2}" name="${attrEscape(name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${relId}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

function pptxChartXml(chart) {
  const chartElementName = chart.chartType === "line" ? "lineChart" : "barChart";
  const grouping = chart.chartType === "line" ? "<c:grouping val=\"standard\"/>" : "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>";
  const seriesXml = (chart.series.length ? chart.series : [{ name: chart.title || "Series", values: [] }]).map((series, index) => {
    const values = series.values || [];
    const categories = series.categories || chart.categories || values.map((_, i) => String(i + 1));
    const catPts = categories.map((category, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${xmlEscape(category)}</c:v></c:pt>`).join("");
    const valPts = values.map((value, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${Number(value) || 0}</c:v></c:pt>`).join("");
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:v>${xmlEscape(series.name || `Series ${index + 1}`)}</c:v></c:tx><c:cat><c:strLit><c:ptCount val="${categories.length}"/>${catPts}</c:strLit></c:cat><c:val><c:numLit><c:ptCount val="${values.length}"/>${valPts}</c:numLit></c:val></c:ser>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(chart.title || chart.chartType)}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:layout/><c:${chartElementName}>${grouping}${seriesXml}<c:axId val="1"/><c:axId val="2"/></c:${chartElementName}><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:layout/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

function slideXml(slide, imageParts = [], chartParts = []) {
  const imageRelById = new Map(imageParts.map((part) => [part.image.id, part.slideRelId]));
  const chartRelById = new Map(chartParts.map((part) => [part.chart.id, part.slideRelId]));
  const elements = [...slide.shapes.items, ...slide.tables.items, ...slide.charts.items, ...slide.images.items];
  const shapes = elements.map((element, index) => element.toPptxShape(index, imageRelById.get(element.id) || chartRelById.get(element.id))).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld></p:sld>`;
}

function parseSlideXml(slide, xml) {
  for (const match of xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const part = match[0];
    const name = decodeXml(/<p:cNvPr[^>]*name="([^"]*)"/.exec(part)?.[1] || "");
    const text = [...part.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decodeXml(m[1])).join("");
    const off = /<a:off[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/.exec(part);
    const ext = /<a:ext[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(part);
    const shape = slide.shapes.add({ name, position: { left: off ? Number(off[1]) / 9525 : 0, top: off ? Number(off[2]) / 9525 : 0, width: ext ? Number(ext[1]) / 9525 : 160, height: ext ? Number(ext[2]) / 9525 : 80 } });
    shape.text = text;
  }
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
  values() { return [...this.items.values()]; }
}

class DocumentTableCell {
  constructor(table, row, column) { this.table = table; this.row = row; this.column = column; }
  get value() { return this.table.values[this.row]?.[this.column] ?? ""; }
  set value(value) { this.table.ensureCell(this.row, this.column); this.table.values[this.row][this.column] = value; }
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
  }

  ensureCell(row, column) { while (this.values.length <= row) this.values.push([]); while (this.values[row].length <= column) this.values[row].push(""); this.rows = this.values.length; this.columns = Math.max(this.columns, column + 1); }
  getCell(row, column) { return new DocumentTableCell(this, row, column); }
  inspectRecord(index) { return { kind: "table", id: this.id, index, name: this.name || undefined, rows: this.rows, cols: this.columns, styleId: this.styleId, values: this.values }; }
  toProto() { return { kind: "table", id: this.id, name: this.name, styleId: this.styleId, values: this.values }; }
}

class DocumentParagraphBlock {
  constructor(document, text, config = {}) {
    this.document = document;
    this.kind = "paragraph";
    this.id = config.id || aid("dp");
    this.text = String(text ?? "");
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "paragraph", id: this.id, index, name: this.name || undefined, styleId: this.styleId, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: "paragraph", id: this.id, name: this.name, styleId: this.styleId, text: this.text }; }
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
    this.listType = config.listType || config.type || "bullet";
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "listItem", id: this.id, index, name: this.name || undefined, styleId: this.styleId, listType: this.listType, level: this.level, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: "listItem", id: this.id, name: this.name, styleId: this.styleId, listType: this.listType, level: this.level, text: this.text }; }
}

class DocumentHyperlinkBlock {
  constructor(document, text, url, config = {}) {
    this.document = document;
    this.kind = "hyperlink";
    this.id = config.id || aid("dhl");
    this.text = String(text ?? "");
    this.url = String(url ?? config.url ?? "");
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "hyperlink", id: this.id, index, name: this.name || undefined, styleId: this.styleId, text: this.text, url: this.url, textChars: this.text.length }; }
  toProto() { return { kind: "hyperlink", id: this.id, name: this.name, styleId: this.styleId, text: this.text, url: this.url }; }
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

class DocumentHeaderFooterBlock {
  constructor(document, kind, text, config = {}) {
    this.document = document;
    this.kind = kind;
    this.id = config.id || aid(kind === "header" ? "dh" : "df");
    this.text = String(text ?? "");
    this.name = config.name || kind;
    this.styleId = config.styleId || "Normal";
  }

  inspectRecord(index) { return { kind: this.kind, id: this.id, index, name: this.name || undefined, styleId: this.styleId, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: this.kind, id: this.id, name: this.name, styleId: this.styleId, text: this.text }; }
}

class DocumentComment {
  constructor(document, targetId, text, config = {}) {
    this.document = document;
    this.kind = "comment";
    this.id = config.id || aid("dc");
    this.targetId = targetId;
    this.author = config.author || "User";
    this.text = String(text ?? "");
    this.resolved = Boolean(config.resolved);
  }

  inspectRecord() { return { kind: "comment", id: this.id, targetId: this.targetId, author: this.author, resolved: this.resolved, textPreview: this.text.slice(0, 300) }; }
  toProto() { return { kind: "comment", id: this.id, targetId: this.targetId, author: this.author, text: this.text, resolved: this.resolved }; }
}

export class DocumentModel {
  constructor(options = {}) {
    this.id = aid("doc");
    this.name = options.name || "New document";
    this.styles = new DocumentStyleCollection(options.styles || {});
    this.blocks = [];
    this.comments = [];
    this.headers = [];
    this.footers = [];
    const sourceBlocks = options.blocks || (options.paragraphs ? options.paragraphs.map((text, index) => ({ kind: "paragraph", text, styleId: index === 0 ? "Title" : "Normal" })) : [{ kind: "paragraph", text: "Start writing here...", styleId: "Normal" }]);
    for (const block of sourceBlocks) {
      if (block.kind === "table") this.addTable(block);
      else if (block.kind === "listItem") this.addListItem(block.text ?? "", block);
      else if (block.kind === "hyperlink") this.addHyperlink(block.text ?? "", block.url, block);
      else if (block.kind === "field") this.addField(block.instruction, block.display, block);
      else if (block.kind === "citation") this.addCitation(block.text ?? "", block.metadata || {}, block);
      else if (block.kind === "image") this.addImage(block);
      else if (block.kind === "change") this.addChange(block.changeType || block.type, block.text ?? "", block);
      else this.addParagraph(block.text ?? "", block);
    }
    for (const header of options.headers || []) this.addHeader(header.text, header);
    for (const footer of options.footers || []) this.addFooter(footer.text, footer);
    for (const comment of options.comments || []) this.addComment(comment.targetId, comment.text, comment);
  }

  static create(options = {}) { return new DocumentModel(options); }
  get paragraphs() { return this.blocks.filter((block) => block.kind === "paragraph").map((block) => block.text); }

  addParagraph(text, config = {}) { const block = new DocumentParagraphBlock(this, text, config); this.blocks.push(block); return block; }
  addListItem(text, config = {}) { const block = new DocumentListItemBlock(this, text, config); this.blocks.push(block); return block; }
  addList(items = [], config = {}) { return items.map((item) => this.addListItem(typeof item === "string" ? item : item.text, { ...config, ...(typeof item === "string" ? {} : item) })); }
  addHyperlink(text, url, config = {}) { const block = new DocumentHyperlinkBlock(this, text, url, config); this.blocks.push(block); return block; }
  addField(instruction, display, config = {}) { const block = new DocumentFieldBlock(this, instruction, display, config); this.blocks.push(block); return block; }
  addCitation(text, metadata = {}, config = {}) { const block = new DocumentCitationBlock(this, text, metadata, config); this.blocks.push(block); return block; }
  addImage(config = {}) { const block = new DocumentImageBlock(this, config); this.blocks.push(block); return block; }
  addChange(changeType, text, config = {}) { const block = new DocumentChangeBlock(this, changeType, text, config); this.blocks.push(block); return block; }
  addInsertion(text, config = {}) { return this.addChange("insert", text, config); }
  addDeletion(text, config = {}) { return this.addChange("delete", text, config); }
  addTable(config = {}) { const block = new DocumentTableBlock(this, config); this.blocks.push(block); return block; }
  addHeader(text, config = {}) { const block = new DocumentHeaderFooterBlock(this, "header", text, config); this.headers.push(block); return block; }
  addFooter(text, config = {}) { const block = new DocumentHeaderFooterBlock(this, "footer", text, config); this.footers.push(block); return block; }
  addComment(target, text, config = {}) { const targetId = typeof target === "string" ? target : target?.id; const comment = new DocumentComment(this, targetId, text, config); this.comments.push(comment); return comment; }
  resolve(id) { return this.id === id ? this : this.blocks.find((block) => block.id === id) || this.headers.find((block) => block.id === id) || this.footers.find((block) => block.id === id) || this.comments.find((comment) => comment.id === id) || this.styles.get(id); }

  toProto() { return { id: this.id, name: this.name, styles: Object.fromEntries(this.styles.values().map((style) => [style.id, style])), blocks: this.blocks.map((block) => block.toProto()), headers: this.headers.map((block) => block.toProto()), footers: this.footers.map((block) => block.toProto()), comments: this.comments.map((comment) => comment.toProto()) }; }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["paragraph", "table", "listItem", "hyperlink", "field", "citation", "image", "change", "comment", "header", "footer"]);
    const records = [];
    this.blocks.forEach((block, index) => { if (kinds.has(block.kind)) records.push(block.inspectRecord(index)); });
    if (kinds.has("header")) records.push(...this.headers.map((block, index) => block.inspectRecord(index)));
    if (kinds.has("footer")) records.push(...this.footers.map((block, index) => block.inspectRecord(index)));
    if (kinds.has("comment")) records.push(...this.comments.map((comment) => comment.inspectRecord()));
    if (kinds.has("style")) records.push(...this.styles.values().map((style) => ({ kind: "style", ...style })));
    const search = String(options.search || "").trim().toLowerCase();
    return ndjson(search ? records.filter((record) => JSON.stringify(record).toLowerCase().includes(search)) : records, options.maxChars ?? Infinity);
  }

  verify(options = {}) {
    const issues = [];
    if (this.blocks.length === 0) issues.push(verificationIssue("document", "emptyDocument", "Document has no body blocks."));
    for (const block of this.blocks) {
      if (block.kind === "paragraph" && /^\s*([-*•]|\d+[.)])\s+/.test(block.text)) {
        issues.push(verificationIssue("document", "fakeList", `Paragraph ${block.id} looks like a fake list item; use addListItem instead.`, { id: block.id }));
      }
      if (block.kind === "hyperlink" && !/^https?:\/\//.test(block.url)) {
        issues.push(verificationIssue("document", "invalidHyperlink", `Hyperlink ${block.id} is missing an absolute http(s) URL.`, { id: block.id, url: block.url }));
      }
      if (block.kind === "field" && !block.instruction.trim()) {
        issues.push(verificationIssue("document", "emptyField", `Field ${block.id} is missing an instruction.`, { id: block.id }));
      }
      if (block.kind === "image") {
        if (!block.dataUrl && !block.uri && !block.prompt) issues.push(verificationIssue("document", "emptyImage", `Image ${block.id} has no dataUrl, uri, or prompt.`, { id: block.id }));
        if (block.dataUrl && !imageDataFromDataUrl(block.dataUrl)) issues.push(verificationIssue("document", "invalidImageDataUrl", `Image ${block.id} has an unsupported data URL.`, { id: block.id }));
      }
      if (block.kind === "table") {
        for (const row of block.values) for (const cell of row) {
          if (String(cell ?? "").length > 240) issues.push(verificationIssue("document", "tableCellTooLong", `Table ${block.id} contains paragraph-like cell content.`, { severity: "warning", id: block.id }));
        }
      }
    }
    const blockIds = new Set(this.blocks.map((block) => block.id));
    for (const comment of this.comments) {
      if (!blockIds.has(comment.targetId)) issues.push(verificationIssue("document", "danglingComment", `Comment ${comment.id} points at a missing block.`, { id: comment.id, targetId: comment.targetId }));
    }
    return verificationResult("document", issues, options);
  }

  help(query = "*", options = {}) {
    return helpArtifact("document", query, options);
  }

  async render(options = {}) {
    const width = 612;
    const margin = 72;
    let y = 72;
    const parts = [`<rect width="100%" height="100%" fill="white"/>`];
    for (const header of this.headers) {
      parts.push(`<text x="${margin}" y="${Math.max(28, y - 36)}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(header.text)}</text>`);
    }
    let listCounters = new Map();
    for (const block of this.blocks) {
      if (block.kind === "paragraph") {
        const style = this.styles.get(block.styleId) || this.styles.get("Normal");
        const fontSize = Math.max(10, (style.fontSize || 22) / 2);
        parts.push(`<text x="${margin}" y="${y}" font-family="${xmlEscape(style.fontFamily || "Arial")}" font-size="${fontSize}" font-weight="${style.bold ? "700" : "400"}" fill="#111827">${xmlEscape(block.text)}</text>`);
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
        if (block.dataUrl) parts.push(`<image href="${attrEscape(block.dataUrl)}" x="${margin}" y="${y}" width="${imageWidth}" height="${imageHeight}" preserveAspectRatio="xMidYMid meet"/>`);
        else parts.push(`<rect x="${margin}" y="${y}" width="${imageWidth}" height="${imageHeight}" fill="#fef3c7" stroke="#f59e0b"/>`);
        parts.push(`<text x="${margin + 8}" y="${y + 18}" font-family="Arial" font-size="11" fill="#92400e">${xmlEscape(block.alt || block.prompt || block.uri || block.name || "image")}</text>`);
        y += imageHeight + 20;
      } else if (block.kind === "change") {
        const marker = block.changeType === "delete" ? "-" : "+";
        const fill = block.changeType === "delete" ? "#dc2626" : "#047857";
        const decoration = block.changeType === "delete" ? " text-decoration=\"line-through\"" : "";
        parts.push(`<text x="${margin}" y="${y}" font-family="Arial" font-size="10" fill="#64748b">tracked ${xmlEscape(block.changeType)} by ${xmlEscape(block.author)}</text>`);
        parts.push(`<text x="${margin + 92}" y="${y}" font-family="Arial" font-size="11" fill="${fill}"${decoration}>${xmlEscape(marker)} ${xmlEscape(block.text)}</text>`);
        y += 22;
      } else if (block.kind === "listItem") {
        const key = `${block.listType}:${block.level}`;
        const next = (listCounters.get(key) || 0) + 1;
        listCounters.set(key, next);
        const marker = block.listType === "number" ? `${next}.` : "•";
        const x = margin + block.level * 24;
        parts.push(`<text x="${x}" y="${y}" font-family="Arial" font-size="11" fill="#111827">${xmlEscape(marker)}</text>`);
        parts.push(`<text x="${x + 22}" y="${y}" font-family="Arial" font-size="11" fill="#111827">${xmlEscape(block.text)}</text>`);
        y += 20;
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
    for (const footer of this.footers) {
      parts.push(`<text x="${margin}" y="${height - 36}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(footer.text)}</text>`);
    }
    return new FileBlob(`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${parts.join("")}</svg>`, { type: "image/svg+xml" });
  }
}

function docxContentTypes({ hasComments, hasHeader, hasFooter, hasNumbering, imageParts = [] }) {
  const imageDefaults = imageContentTypeDefaults(imageParts);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="json" ContentType="application/json"/>${imageDefaults}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${hasNumbering ? `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` : ""}${hasComments ? `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` : ""}${hasHeader ? `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` : ""}${hasFooter ? `<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>` : ""}</Types>`;
}

function docxStylesXml(document) {
  const styles = document.styles.values().map((style) => `<w:style w:type="paragraph" w:styleId="${attrEscape(style.id)}"><w:name w:val="${attrEscape(style.name || style.id)}"/><w:rPr>${style.bold ? "<w:b/>" : ""}<w:sz w:val="${Math.round(style.fontSize || 22)}"/><w:rFonts w:ascii="${attrEscape(style.fontFamily || "Aptos")}" w:hAnsi="${attrEscape(style.fontFamily || "Aptos")}"/></w:rPr></w:style>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${styles}</w:styles>`;
}

function docxNumberingXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>`;
}

function docxHeaderFooterXml(kind, blocks) {
  const tag = kind === "header" ? "hdr" : "ftr";
  const body = blocks.map((block) => `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr><w:r><w:t>${xmlEscape(block.text)}</w:t></w:r></w:p>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${body}</w:${tag}>`;
}

function docxParagraphXml(block, commentIndexes) {
  const commentStart = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeStart w:id="${id}"/>`).join("") : "";
  const commentEnd = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeEnd w:id="${id}"/>`).join("") : "";
  const refs = commentIndexes.length ? commentIndexes.map((id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`).join("") : "";
  const numPr = block.kind === "listItem" ? `<w:numPr><w:ilvl w:val="${Math.max(0, block.level || 0)}"/><w:numId w:val="${block.listType === "number" ? 2 : 1}"/></w:numPr>` : "";
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/>${numPr}</w:pPr>${commentStart}<w:r><w:t>${xmlEscape(block.text)}</w:t></w:r>${commentEnd}${refs}</w:p>`;
}

function docxRevisionId(block) {
  const raw = String(block?.id || block?.name || block?.text || "1");
  const sum = raw.split("").reduce((acc, ch) => ((acc * 31) + ch.charCodeAt(0)) >>> 0, 17);
  return Math.max(1, sum % 2147483647);
}

function docxChangeXml(block, commentIndexes = []) {
  const commentStart = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeStart w:id="${id}"/>`).join("") : "";
  const commentEnd = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeEnd w:id="${id}"/>`).join("") : "";
  const refs = commentIndexes.length ? commentIndexes.map((id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`).join("") : "";
  const tag = block.changeType === "delete" ? "del" : "ins";
  const textTag = block.changeType === "delete" ? "delText" : "t";
  const date = block.date ? ` w:date="${attrEscape(block.date)}"` : "";
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr>${commentStart}<w:${tag} w:id="${docxRevisionId(block)}" w:author="${attrEscape(block.author || "User")}"${date}><w:r><w:${textTag}>${xmlEscape(block.text)}</w:${textTag}></w:r></w:${tag}>${commentEnd}${refs}</w:p>`;
}

function collectDocxImageParts(document) {
  const parts = [];
  let imagePartId = 1;
  let relIndex = 1;
  for (const block of document.blocks.filter((item) => item.kind === "image")) {
    const data = imageDataFromDataUrl(block.dataUrl);
    if (!data) continue;
    parts.push({ image: block, imagePartId: imagePartId++, relId: `rIdImage${relIndex++}`, ...data });
  }
  return parts;
}

function docxImageXml(block, relId, commentIndexes = []) {
  const commentStart = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeStart w:id="${id}"/>`).join("") : "";
  const commentEnd = commentIndexes.length ? commentIndexes.map((id) => `<w:commentRangeEnd w:id="${id}"/>`).join("") : "";
  const refs = commentIndexes.length ? commentIndexes.map((id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`).join("") : "";
  if (!relId) return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr>${commentStart}<w:r><w:t>${xmlEscape(block.alt || block.prompt || block.uri || block.name || "image")}</w:t></w:r>${commentEnd}${refs}</w:p>`;
  const cx = Math.round(Math.max(16, Number(block.widthPx) || 240) * 9525);
  const cy = Math.round(Math.max(16, Number(block.heightPx) || 160) * 9525);
  const docPrId = docxRevisionId(block);
  const name = block.name || `Image ${docPrId}`;
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr>${commentStart}<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${docPrId}" name="${attrEscape(name)}" descr="${attrEscape(block.alt || "")}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${docPrId}" name="${attrEscape(name)}" descr="${attrEscape(block.alt || "")}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>${commentEnd}${refs}</w:p>`;
}

function docxHyperlinkXml(block, relId) {
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr><w:hyperlink r:id="${relId}"><w:r><w:rPr><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t>${xmlEscape(block.text)}</w:t></w:r></w:hyperlink></w:p>`;
}

function docxFieldXml(block) {
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr><w:fldSimple w:instr="${attrEscape(block.instruction)}"><w:r><w:t>${xmlEscape(block.display)}</w:t></w:r></w:fldSimple></w:p>`;
}

function docxCitationXml(block) {
  const label = block.metadata?.source ? `${block.text} (${block.metadata.source})` : block.text;
  return `<w:p><w:pPr><w:pStyle w:val="${attrEscape(block.styleId || "Normal")}"/></w:pPr><w:bookmarkStart w:id="${Math.abs(block.id.split("").reduce((sum, ch) => sum + ch.charCodeAt(0), 0))}" w:name="${attrEscape(`OpenOfficeCitation_${block.id.replace(/[^A-Za-z0-9_]/g, "_")}`)}"/><w:r><w:t>${xmlEscape(label)}</w:t></w:r><w:bookmarkEnd w:id="${Math.abs(block.id.split("").reduce((sum, ch) => sum + ch.charCodeAt(0), 0))}"/></w:p>`;
}

function docxTableXml(block) {
  const grid = Array.from({ length: block.columns || 1 }, () => `<w:gridCol w:w="${Math.floor(9360 / Math.max(1, block.columns || 1))}"/>`).join("");
  const rows = block.values.map((row) => `<w:tr>${Array.from({ length: block.columns || row.length || 1 }, (_, column) => `<w:tc><w:tcPr><w:tcW w:w="${Math.floor(9360 / Math.max(1, block.columns || 1))}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${xmlEscape(row[column] ?? "")}</w:t></w:r></w:p></w:tc>`).join("")}</w:tr>`).join("");
  return `<w:tbl><w:tblPr><w:tblStyle w:val="${attrEscape(block.styleId || "TableGrid")}"/><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
}

function docxDocumentXml(document, relIds = {}) {
  const commentIndex = new Map(document.comments.map((comment, index) => [comment, index]));
  const body = document.blocks.map((block) => {
    if (block.kind === "table") return docxTableXml(block);
    if (block.kind === "hyperlink") return docxHyperlinkXml(block, relIds.hyperlinks?.get(block.id));
    if (block.kind === "field") return docxFieldXml(block);
    if (block.kind === "citation") return docxCitationXml(block);
    const indexes = document.comments.filter((comment) => comment.targetId === block.id).map((comment) => commentIndex.get(comment));
    if (block.kind === "image") return docxImageXml(block, relIds.images?.get(block.id), indexes);
    if (block.kind === "change") return docxChangeXml(block, indexes);
    return docxParagraphXml(block, indexes);
  }).join("");
  const refs = `${relIds.header ? `<w:headerReference w:type="default" r:id="${relIds.header}"/>` : ""}${relIds.footer ? `<w:footerReference w:type="default" r:id="${relIds.footer}"/>` : ""}`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body}<w:sectPr>${refs}</w:sectPr></w:body></w:document>`;
}

function docxCommentsXml(document) {
  const comments = document.comments.map((comment, index) => `<w:comment w:id="${index}" w:author="${attrEscape(comment.author)}"><w:p><w:r><w:t>${xmlEscape(comment.text)}</w:t></w:r></w:p></w:comment>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${comments}</w:comments>`;
}

function parseDocxParagraph(part, commentTextById, imageByRelId = new Map()) {
  const styleId = /<w:pStyle[^>]*w:val="([^"]+)"/.exec(part)?.[1] || "Normal";
  const commentIds = [...part.matchAll(/<w:commentRangeStart[^>]*w:id="(\d+)"/g)].map((match) => match[1]);
  const imageRelId = /<a:blip[^>]*r:embed="([^"]+)"/.exec(part)?.[1];
  if (imageRelId) {
    const image = imageByRelId.get(imageRelId) || {};
    const docPr = /<wp:docPr\b([^>]*)\/>/.exec(part)?.[1] || /<pic:cNvPr\b([^>]*)\/>/.exec(part)?.[1] || "";
    const extent = /<wp:extent[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(part);
    const name = decodeXml(/\bname="([^"]*)"/.exec(docPr)?.[1] || image.name || "");
    const alt = decodeXml(/\bdescr="([^"]*)"/.exec(docPr)?.[1] || image.alt || name || "image");
    return { block: { kind: "image", name, alt, dataUrl: image.dataUrl, uri: image.uri, widthPx: extent ? Math.round(Number(extent[1]) / 9525) : image.widthPx, heightPx: extent ? Math.round(Number(extent[2]) / 9525) : image.heightPx, styleId }, commentIds };
  }
  const changeMatch = /<w:(ins|del)\b([^>]*)>([\s\S]*?)<\/w:\1>/.exec(part);
  if (changeMatch) {
    const changeType = changeMatch[1] === "del" ? "delete" : "insert";
    const attrs = changeMatch[2] || "";
    const inner = changeMatch[3] || "";
    const text = decodeXml([...inner.matchAll(/<w:(?:t|delText)[^>]*>([\s\S]*?)<\/w:(?:t|delText)>/g)].map((t) => t[1]).join(""));
    const author = decodeXml(/w:author="([^"]*)"/.exec(attrs)?.[1] || "User");
    const date = decodeXml(/w:date="([^"]*)"/.exec(attrs)?.[1] || "");
    return { block: { kind: "change", changeType, text, styleId, author, date }, commentIds };
  }
  const text = decodeXml([...part.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join(""));
  const numId = /<w:numId[^>]*w:val="(\d+)"/.exec(part)?.[1];
  const level = Number(/<w:ilvl[^>]*w:val="(\d+)"/.exec(part)?.[1] || 0);
  if (numId) return { block: { kind: "listItem", text, styleId, level, listType: numId === "2" ? "number" : "bullet" }, commentIds };
  return { block: { kind: "paragraph", text, styleId }, commentIds };
}

function parseDocxTable(part) {
  const values = [...part.matchAll(/<w:tr[\s\S]*?<\/w:tr>/g)].map((rowMatch) => [...rowMatch[0].matchAll(/<w:tc[\s\S]*?<\/w:tc>/g)].map((cellMatch) => decodeXml([...cellMatch[0].matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join(""))));
  return { kind: "table", values, rows: values.length, columns: Math.max(0, ...values.map((row) => row.length)) };
}

function parseHeaderFooterXml(xml) {
  return [...String(xml || "").matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((match) => ({ text: decodeXml([...match[0].matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join("")) })).filter((item) => item.text.length > 0);
}

export class DocumentFile {
  static async exportDocx(document) {
    const zip = new JSZip();
    const imageParts = collectDocxImageParts(document);
    const hasNumbering = document.blocks.some((block) => block.kind === "listItem");
    const hasHeader = document.headers.length > 0;
    const hasFooter = document.footers.length > 0;
    zip.file("[Content_Types].xml", docxContentTypes({ hasComments: document.comments.length > 0, hasHeader, hasFooter, hasNumbering, imageParts }));
    zip.file("_rels/.rels", relsXml([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "word/document.xml" }]));
    const docRels = [{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", target: "styles.xml" }];
    const relIds = {};
    if (hasNumbering) docRels.push({ id: `rId${docRels.length + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", target: "numbering.xml" });
    if (document.comments.length) docRels.push({ id: `rId${docRels.length + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", target: "comments.xml" });
    if (hasHeader) { relIds.header = `rId${docRels.length + 1}`; docRels.push({ id: relIds.header, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", target: "header1.xml" }); }
    if (hasFooter) { relIds.footer = `rId${docRels.length + 1}`; docRels.push({ id: relIds.footer, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", target: "footer1.xml" }); }
    relIds.hyperlinks = new Map();
    for (const block of document.blocks.filter((item) => item.kind === "hyperlink")) {
      const relId = `rId${docRels.length + 1}`;
      relIds.hyperlinks.set(block.id, relId);
      docRels.push({ id: relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", target: block.url, targetMode: "External" });
    }
    relIds.images = new Map();
    for (const part of imageParts) {
      relIds.images.set(part.image.id, part.relId);
      docRels.push({ id: part.relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: `media/image${part.imagePartId}.${part.extension}` });
    }
    zip.file("word/_rels/document.xml.rels", relsXml(docRels));
    zip.file("word/styles.xml", docxStylesXml(document));
    if (hasNumbering) zip.file("word/numbering.xml", docxNumberingXml());
    if (document.comments.length) zip.file("word/comments.xml", docxCommentsXml(document));
    if (hasHeader) zip.file("word/header1.xml", docxHeaderFooterXml("header", document.headers));
    if (hasFooter) zip.file("word/footer1.xml", docxHeaderFooterXml("footer", document.footers));
    imageParts.forEach((part) => zip.file(`word/media/image${part.imagePartId}.${part.extension}`, part.bytes));
    zip.file("word/open-office-artifact.json", JSON.stringify(document.toProto(), null, 2));
    zip.file("word/document.xml", docxDocumentXml(document, relIds));
    return new FileBlob(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: DOCX_MIME });
  }

  static async importDocx(blobOrBuffer) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const zip = await JSZip.loadAsync(bytes);
    const metadataText = await zip.file("word/open-office-artifact.json")?.async("text");
    if (metadataText) return DocumentModel.create(JSON.parse(metadataText));
    const xml = await zip.file("word/document.xml")?.async("text");
    const commentsXml = await zip.file("word/comments.xml")?.async("text");
    const commentTextById = new Map([...String(commentsXml || "").matchAll(/<w:comment[^>]*w:id="(\d+)"[^>]*>([\s\S]*?)<\/w:comment>/g)].map((match) => [match[1], decodeXml([...match[2].matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((t) => t[1]).join(""))]));
    const relsText = await zip.file("word/_rels/document.xml.rels")?.async("text");
    const imageByRelId = new Map();
    for (const rel of parseRelsXml(relsText).filter((item) => item.type.endsWith("/image"))) {
      const target = rel.target.replace(/^\//, "");
      const packagePath = target.startsWith("word/") ? target : path.posix.normalize(`word/${target}`).replace(/^\.\//, "");
      const bytes = await zip.file(packagePath)?.async("uint8array");
      const extension = /\.([A-Za-z0-9+]+)$/.exec(target)?.[1] || "bin";
      imageByRelId.set(rel.id, bytes ? { dataUrl: `data:${imageContentTypeFromExtension(extension)};base64,${Buffer.from(bytes).toString("base64")}` } : { uri: rel.target });
    }
    const blocks = [];
    const pendingComments = [];
    for (const match of String(xml || "").matchAll(/<w:tbl[\s\S]*?<\/w:tbl>|<w:p[\s\S]*?<\/w:p>/g)) {
      const part = match[0];
      if (part.startsWith("<w:tbl")) blocks.push(parseDocxTable(part));
      else {
        const parsed = parseDocxParagraph(part, commentTextById, imageByRelId);
        blocks.push(parsed.block);
        for (const commentId of parsed.commentIds) pendingComments.push({ blockIndex: blocks.length - 1, text: commentTextById.get(commentId) || "" });
      }
    }
    const document = DocumentModel.create({ blocks: blocks.length ? blocks : [{ kind: "paragraph", text: "" }] });
    for (const header of parseHeaderFooterXml(await zip.file("word/header1.xml")?.async("text"))) document.addHeader(header.text, header);
    for (const footer of parseHeaderFooterXml(await zip.file("word/footer1.xml")?.async("text"))) document.addFooter(footer.text, footer);
    pendingComments.forEach((comment) => document.addComment(document.blocks[comment.blockIndex]?.id, comment.text));
    return document;
  }
}

class PdfTable {
  constructor(page, config = {}) {
    this.page = page;
    this.id = config.id || aid("ptb");
    this.name = config.name || "";
    this.values = (config.values || [[]]).map((row) => [...row]);
    this.bbox = config.bbox || [72, 140, 468, Math.max(24, this.values.length * 24)];
  }

  inspectRecord(pageIndex) { return { kind: "table", id: this.id, page: pageIndex + 1, name: this.name || undefined, rows: this.values.length, cols: Math.max(0, ...this.values.map((row) => row.length)), bbox: this.bbox, values: this.values }; }
  toJSON() { return { id: this.id, name: this.name, values: this.values, bbox: this.bbox }; }
}

class PdfPage {
  constructor(artifact, config = {}) {
    this.artifact = artifact;
    this.id = config.id || aid("pg");
    this.text = String(config.text || "");
    this.width = config.width || 612;
    this.height = config.height || 792;
    this.tables = (config.tables || []).map((table) => new PdfTable(this, table));
  }

  addTable(config = {}) { const table = new PdfTable(this, config); this.tables.push(table); return table; }
  inspectRecord(index) { return { kind: "page", id: this.id, page: index + 1, width: this.width, height: this.height, textPreview: this.text.slice(0, 300), textChars: this.text.length, tables: this.tables.length }; }
  textRecord(index) { return { kind: "text", id: `${this.id}/text`, page: index + 1, text: this.text, textChars: this.text.length }; }
  toJSON() { return { id: this.id, text: this.text, width: this.width, height: this.height, tables: this.tables.map((table) => table.toJSON()) }; }
}

export class PdfArtifact {
  constructor(options = {}) {
    this.id = options.id || aid("pdf");
    const pages = options.pages || [{ text: options.text || "", tables: options.tables || [] }];
    this.pages = pages.map((page) => new PdfPage(this, page));
  }

  static create(options = {}) { return new PdfArtifact(options); }
  addPage(config = {}) { const page = new PdfPage(this, config); this.pages.push(page); return page; }
  addTable(config = {}) { return (this.pages[0] || this.addPage()).addTable(config); }
  extractText(options = {}) { const pages = options.page == null ? this.pages : [this.pages[Number(options.page) - 1]].filter(Boolean); return pages.map((page) => page.text).join("\n\n"); }
  extractTables(options = {}) { const pages = options.page == null ? this.pages : [this.pages[Number(options.page) - 1]].filter(Boolean); return pages.flatMap((page, index) => page.tables.map((table) => ({ page: options.page || index + 1, id: table.id, name: table.name, values: table.values, bbox: table.bbox }))); }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["page", "text", "table"]);
    const records = [];
    this.pages.forEach((page, index) => {
      if (kinds.has("page")) records.push(page.inspectRecord(index));
      if (kinds.has("text")) records.push(page.textRecord(index));
      if (kinds.has("table")) records.push(...page.tables.map((table) => table.inspectRecord(index)));
    });
    const search = String(options.search || "").trim().toLowerCase();
    return ndjson(search ? records.filter((record) => JSON.stringify(record).toLowerCase().includes(search)) : records, options.maxChars ?? Infinity);
  }

  verify(options = {}) {
    const issues = [];
    if (this.pages.length === 0) issues.push(verificationIssue("pdf", "noPages", "PDF artifact has no pages."));
    this.pages.forEach((page, pageIndex) => {
      if (!page.text.trim() && page.tables.length === 0) issues.push(verificationIssue("pdf", "emptyPage", `PDF page ${pageIndex + 1} has no modeled text or tables.`, { page: pageIndex + 1 }));
      if (/[\u2010-\u2015]/.test(page.text)) issues.push(verificationIssue("pdf", "unicodeDash", `PDF page ${pageIndex + 1} contains a Unicode dash; use ASCII hyphen for compatibility.`, { page: pageIndex + 1 }));
      for (const table of page.tables) {
        if (!table.values.length || !table.values[0]?.length) issues.push(verificationIssue("pdf", "emptyTable", `PDF table ${table.id} on page ${pageIndex + 1} has no cells.`, { page: pageIndex + 1, id: table.id }));
        const [left, top, width, height] = table.bbox || [];
        if (left < 0 || top < 0 || width <= 0 || height <= 0 || left + width > page.width || top + height > page.height) {
          issues.push(verificationIssue("pdf", "tableOutOfBounds", `PDF table ${table.id} extends outside page ${pageIndex + 1}.`, { page: pageIndex + 1, id: table.id, bbox: table.bbox }));
        }
      }
    });
    return verificationResult("pdf", issues, options);
  }

  help(query = "*", options = {}) { return helpArtifact("pdf", query, options); }

  async render(options = {}) { return new FileBlob(pdfPageSvg(this.pages[options.pageIndex || 0] || new PdfPage(this)), { type: "image/svg+xml" }); }
  toJSON() { return { id: this.id, pages: this.pages.map((page) => page.toJSON()) }; }
}

export class PdfFile {
  static async exportPdf(artifact) {
    return new FileBlob(buildMinimalPdf(artifact), { type: PDF_MIME });
  }

  static async importPdf(blobOrBuffer) {
    const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
    const text = decoder.decode(bytes);
    const metadata = /%OPEN_OFFICE_ARTIFACT ([A-Za-z0-9+/=]+)/.exec(text)?.[1];
    if (metadata) return PdfArtifact.create(JSON.parse(Buffer.from(metadata, "base64").toString("utf8")));
    const strings = [...text.matchAll(/\(([^()]*)\)\s*Tj/g)].map((m) => m[1].replaceAll("\\)", ")").replaceAll("\\(", "("));
    const plain = strings.join("\n");
    const tableRows = strings.filter((line) => line.includes("|")).map((line) => line.split("|").map((cell) => cell.trim()));
    return PdfArtifact.create({ pages: [{ text: plain, tables: tableRows.length ? [{ name: "extracted-table", values: tableRows }] : [] }] });
  }
}

function pdfPageSvg(page) {
  const width = page.width || 612;
  const height = page.height || 792;
  const lines = String(page.text || "").split(/\r?\n/).filter(Boolean);
  const text = lines.map((line, index) => `<text x="72" y="${96 + index * 20}" font-family="Helvetica" font-size="14" fill="#111827">${xmlEscape(line)}</text>`).join("");
  const tables = (page.tables || []).map((table) => {
    const [left, top, tableWidth, tableHeight] = table.bbox;
    const rows = table.values.length;
    const cols = Math.max(1, ...table.values.map((row) => row.length));
    const cellW = tableWidth / cols;
    const cellH = tableHeight / Math.max(1, rows);
    const cells = [];
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const x = left + c * cellW;
        const y = top + r * cellH;
        cells.push(`<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" fill="${r === 0 ? "#f1f5f9" : "#ffffff"}" stroke="#cbd5e1"/>`);
        cells.push(`<text x="${x + 5}" y="${y + Math.min(16, cellH - 4)}" font-family="Helvetica" font-size="11" fill="#111827">${xmlEscape(table.values[r]?.[c] ?? "")}</text>`);
      }
    }
    return cells.join("");
  }).join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="white"/>${text}${tables}</svg>`;
}

function escapePdfString(text) {
  return String(text).replaceAll("\\", "\\\\").replaceAll("(", "\\(").replaceAll(")", "\\)");
}

function buildMinimalPdf(artifact) {
  const page = artifact.pages[0] || new PdfPage(artifact);
  const tableLines = page.tables.flatMap((table) => table.values.map((row) => row.join(" | ")));
  const lines = [...String(page.text || "").split(/\r?\n/), ...tableLines].filter(Boolean);
  const content = `BT\n/F1 18 Tf\n72 720 Td\n${lines.map((line, index) => `${index === 0 ? "" : "0 -24 Td\n"}(${escapePdfString(line)}) Tj`).join("\n")}\nET`;
  const metadata = Buffer.from(JSON.stringify(artifact.toJSON()), "utf8").toString("base64");
  const objects = [
    `1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n`,
    `2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n`,
    `3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${page.width || 612} ${page.height || 792}] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>\nendobj\n`,
    `4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n`,
    `5 0 obj\n<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}\nendstream\nendobj\n`,
  ];
  let pdf = `%PDF-1.4\n%OPEN_OFFICE_ARTIFACT ${metadata}\n`;
  const offsets = [0];
  for (const obj of objects) { offsets.push(Buffer.byteLength(pdf)); pdf += obj; }
  const xref = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (let i = 1; i < offsets.length; i++) pdf += `${String(offsets[i]).padStart(10, "0")} 00000 n \n`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return encoder.encode(pdf);
}
