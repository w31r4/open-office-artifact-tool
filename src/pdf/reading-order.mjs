function normalizedId(value) {
  if (value && typeof value === "object") return String(value.id ?? "").trim();
  return String(value ?? "").trim();
}

export function pdfPageBodyTextLines(page = {}) {
  const positioned = new Set((page.textItems || []).map((item) => String(item.text || "").trim()).filter(Boolean));
  return String(page.text || "").split(/\r?\n/).map((line) => line.trim()).filter((line) => line && !positioned.has(line));
}

export function pdfReadingOrderEntries(page = {}) {
  const entries = [];
  const bodyLines = pdfPageBodyTextLines(page);
  if (bodyLines.length) entries.push({ id: `${page.id}/text`, kind: "text", label: bodyLines[0] });
  for (const item of page.textItems || []) if (String(item.text || "").trim()) entries.push({ id: String(item.id), kind: "textItem", label: String(item.text) });
  for (const table of page.tables || []) entries.push({ id: String(table.id), kind: "table", label: table.name || "Data table" });
  for (const image of page.images || []) if (!image.decorative) entries.push({ id: String(image.id), kind: "image", label: image.alt || image.name || "Image" });
  for (const chart of page.charts || []) if (!chart.decorative) entries.push({ id: String(chart.id), kind: "chart", label: chart.alt || chart.title || chart.name || "Chart" });
  return entries;
}

export function normalizePdfReadingOrder(order) {
  if (order == null) return undefined;
  if (!Array.isArray(order)) throw new TypeError("PDF readingOrder must be an array of target IDs or objects with an id property.");
  return order.map(normalizedId);
}

export function analyzePdfReadingOrder(page = {}) {
  const entries = pdfReadingOrderEntries(page);
  const byId = new Map();
  const ambiguousIds = new Set();
  for (const entry of entries) {
    if (byId.has(entry.id)) ambiguousIds.add(entry.id);
    else byId.set(entry.id, entry);
  }
  const explicit = Array.isArray(page.readingOrder);
  const declaredIds = explicit ? page.readingOrder.map(normalizedId) : entries.map((entry) => entry.id);
  const seen = new Set();
  const duplicateIds = [];
  const unknownIds = [];
  for (const id of declaredIds) {
    if (seen.has(id) && !duplicateIds.includes(id)) duplicateIds.push(id);
    seen.add(id);
    if (!byId.has(id) && !unknownIds.includes(id)) unknownIds.push(id);
  }
  const missingIds = explicit ? entries.map((entry) => entry.id).filter((id) => !seen.has(id)) : [];
  const errors = [
    ...[...ambiguousIds].map((id) => ({ code: "readingOrderAmbiguous", id, message: `Target ID ${id} identifies more than one semantic page item.` })),
    ...duplicateIds.map((id) => ({ code: "readingOrderDuplicate", id, message: `Target ID ${id} appears more than once.` })),
    ...unknownIds.map((id) => ({ code: "readingOrderUnknown", id, message: `Target ID ${id || "(empty)"} does not identify semantic content on this page.` })),
    ...missingIds.map((id) => ({ code: "readingOrderMissing", id, message: `Semantic target ${id} is missing from the explicit order.` })),
  ];
  return { explicit, declaredIds, entries, byId, errors, valid: errors.length === 0 };
}

export function resolvePdfReadingOrder(page = {}, options = {}) {
  const analysis = analyzePdfReadingOrder(page);
  if (options.strict !== false && !analysis.valid) {
    throw new Error(`Invalid PDF reading order for page ${page.id || "(unknown)"}: ${analysis.errors.map((error) => `${error.code}: ${error.message}`).join(" ")}`);
  }
  return analysis.declaredIds.map((id) => analysis.byId.get(id)).filter(Boolean);
}

export function pdfReadingOrderInspectRecords(page = {}, pageIndex = 0) {
  const analysis = analyzePdfReadingOrder(page);
  return analysis.declaredIds.map((targetId, index) => {
    const target = analysis.byId.get(targetId);
    return {
      kind: "readingOrder",
      id: `${page.id}/reading-order/${index + 1}`,
      page: pageIndex + 1,
      index: index + 1,
      targetId,
      targetKind: target?.kind || "unknown",
      label: target?.label,
      explicit: analysis.explicit,
      valid: analysis.valid,
    };
  });
}

function decodePdfLiteral(value) {
  return value.replace(/\\([nrtbf()\\])/g, (_, escape) => ({ n: "\n", r: "\r", t: "\t", b: "\b", f: "\f", "(": "(", ")": ")", "\\": "\\" })[escape]).replace(/\\([0-7]{1,3})/g, (_, octal) => String.fromCharCode(Number.parseInt(octal, 8)));
}

function decodePdfHex(value) {
  const bytes = Buffer.from(value.length % 2 ? `${value}0` : value, "hex");
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    let result = "";
    for (let index = 2; index + 1 < bytes.length; index += 2) result += String.fromCharCode((bytes[index] << 8) | bytes[index + 1]);
    return result;
  }
  return bytes.toString("utf8");
}

export function inspectPdfReadingOrderIds(pdfText = "") {
  const objects = new Map();
  for (const match of String(pdfText).matchAll(/(?:^|\n)(\d+)\s+0\s+obj\s*([\s\S]*?)\s*endobj/g)) objects.set(Number(match[1]), match[2]);
  const rootBody = [...objects.values()].find((body) => /\/Type\s*\/StructTreeRoot\b/.test(body));
  const kids = /\/K\s*\[([^\]]*)\]/.exec(rootBody || "")?.[1] || "";
  const ids = [];
  for (const reference of kids.matchAll(/(\d+)\s+0\s+R/g)) {
    const body = objects.get(Number(reference[1])) || "";
    const token = /\/ID\s*(\((?:\\.|[^\\)])*\)|<([A-Fa-f0-9]+)>)/.exec(body);
    if (!token) continue;
    ids.push(token[2] == null ? decodePdfLiteral(token[1].slice(1, -1)) : decodePdfHex(token[2]));
  }
  return ids;
}
