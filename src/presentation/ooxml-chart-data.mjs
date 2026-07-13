import path from "node:path";

export const PRESENTATION_CHART_EXTERNAL_DATA_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
export const PRESENTATION_CHART_EXTERNAL_DATA_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";

const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const MAX_WORKBOOK_BYTES = 64 * 1024 * 1024;
const decoder = new TextDecoder();

function issue(type, message, detail = {}) {
  return { kind: "ooxmlIssue", family: "PPTX", type, severity: "error", ...detail, message };
}

function xmlAttributes(tag = "") {
  return Object.fromEntries([...String(tag).matchAll(/([A-Za-z_][\w.:-]*)\s*=\s*"([^"]*)"/g)].map((match) => [match[1], match[2]
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&")]));
}

function relationshipPartPath(source) {
  const directory = path.posix.dirname(source);
  const base = path.posix.basename(source);
  return path.posix.join(directory === "." ? "" : directory, "_rels", `${base}.rels`);
}

function resolveTarget(source, rawTarget) {
  const target = String(rawTarget || "").split("#")[0];
  if (target.startsWith("/")) return target.slice(1);
  const directory = path.posix.dirname(source);
  return path.posix.normalize(path.posix.join(directory === "." ? "" : directory, target));
}

function externalDataTag(xml = "") {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?externalData\\b[^>]*>`, "i").exec(String(xml))?.[0]
    || new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?externalData\\b[^>]*/>`, "i").exec(String(xml))?.[0]
    || "";
}

function relationshipId(tag = "") {
  const attrs = xmlAttributes(tag);
  return attrs["r:id"] || Object.entries(attrs).find(([name]) => name.endsWith(":id"))?.[1] || attrs.id;
}

function autoUpdateFromChartXml(xml = "") {
  const block = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?externalData\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?externalData>`, "i").exec(String(xml))?.[1] || "";
  const tag = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?autoUpdate\\b[^>]*/?>`, "i").exec(block)?.[0];
  if (!tag) return true;
  const value = xmlAttributes(tag).val;
  return value == null || !new Set(["0", "false", "off", "no"]).has(String(value).toLowerCase());
}

function bytesFrom(value) {
  if (value == null) return undefined;
  if (value instanceof Uint8Array) return new Uint8Array(value);
  if (value instanceof ArrayBuffer) return new Uint8Array(value.slice(0));
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength));
  if (value.bytes instanceof Uint8Array) return new Uint8Array(value.bytes);
  if (value.bytes instanceof ArrayBuffer) return new Uint8Array(value.bytes.slice(0));
  return undefined;
}

function workbookDataUrlBytes(value) {
  const match = /^data:application\/(?:vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet|octet-stream);base64,([A-Za-z0-9+/=\s]+)$/i.exec(String(value || ""));
  if (!match) return undefined;
  return new Uint8Array(Buffer.from(match[1].replace(/\s+/g, ""), "base64"));
}

function normalizeExternalUri(value) {
  const uri = String(value || "").trim();
  if (!uri || uri.length > 4096) throw new RangeError("presentation chart externalData uri must contain 1 to 4096 characters.");
  let parsed;
  try { parsed = new URL(uri); } catch { throw new TypeError("presentation chart externalData uri must be an absolute http, https, or file URL."); }
  if (!new Set(["http:", "https:", "file:"]).has(parsed.protocol)) throw new TypeError("presentation chart externalData uri must use http, https, or file.");
  return uri;
}

export function normalizePresentationChartExternalData(value) {
  if (value == null || value === false) return undefined;
  const config = typeof value === "object" && !ArrayBuffer.isView(value) && !(value instanceof ArrayBuffer) ? value : { workbook: value };
  const autoUpdate = Boolean(config.autoUpdate);
  if (config.uri != null || config.url != null) {
    if (config.workbook != null || config.dataUrl != null || config.bytes != null) throw new TypeError("presentation chart externalData must use either uri or an embedded workbook, not both.");
    return { uri: normalizeExternalUri(config.uri ?? config.url), autoUpdate };
  }
  const source = config.workbook ?? config.dataUrl ?? config.bytes ?? value;
  const bytes = typeof source === "string" ? workbookDataUrlBytes(source) : bytesFrom(source);
  if (!bytes) throw new TypeError("presentation chart externalData requires workbook bytes, FileBlob, ArrayBuffer, Uint8Array, an XLSX data URL, or uri.");
  if (bytes.byteLength < 4 || bytes.byteLength > MAX_WORKBOOK_BYTES) throw new RangeError(`presentation chart externalData workbook must contain 4 to ${MAX_WORKBOOK_BYTES} bytes.`);
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new TypeError("presentation chart externalData workbook must be an OOXML ZIP package.");
  return { bytes, autoUpdate };
}

export function presentationChartUsesFormulaReferences(chart) {
  return Boolean(chart?.series?.some((series) => series.errorBars?.plusFormula || series.errorBars?.minusFormula));
}

export function planPresentationChartExternalDataParts(chartParts = []) {
  let workbookPartId = 1;
  return chartParts.flatMap((chartPart) => {
    const externalData = chartPart.chart.externalData;
    if (!externalData) return [];
    if (externalData.uri) return [{ chartPart, externalData, relationshipId: "rId1", target: externalData.uri, targetMode: "External" }];
    const outputPath = `ppt/embeddings/Microsoft_Excel_Worksheet${workbookPartId++}.xlsx`;
    return [{ chartPart, externalData, relationshipId: "rId1", outputPath, target: `../embeddings/${path.posix.basename(outputPath)}` }];
  });
}

export function presentationChartExternalDataContentTypesXml(parts = []) {
  return parts.filter((part) => part.outputPath).map((part) => `<Override PartName="/${part.outputPath}" ContentType="${PRESENTATION_CHART_EXTERNAL_DATA_CONTENT_TYPE}"/>`).join("");
}

export function presentationChartExternalDataRelationship(part) {
  return {
    id: part.relationshipId,
    type: PRESENTATION_CHART_EXTERNAL_DATA_RELATIONSHIP_TYPE,
    target: part.target,
    ...(part.targetMode ? { targetMode: part.targetMode } : {}),
  };
}

export async function validatePresentationChartExternalDataWorkbooks(parts = [], inspectXlsx) {
  if (typeof inspectXlsx !== "function") throw new TypeError("Presentation chart embedded-workbook validation requires an XLSX inspector.");
  for (const part of parts.filter((candidate) => candidate.outputPath)) {
    let inspection;
    try {
      inspection = await inspectXlsx(part.externalData.bytes);
    } catch (error) {
      throw new TypeError(`Presentation chart externalData workbook is not a valid XLSX package: ${error.message}`, { cause: error });
    }
    if (!inspection?.ok) {
      const issueTypes = [...new Set((inspection?.issues || []).map((entry) => entry.type).filter(Boolean))].slice(0, 8);
      throw new TypeError(`Presentation chart externalData workbook is not a valid XLSX package${issueTypes.length ? ` (${issueTypes.join(", ")})` : ""}.`);
    }
  }
}

export async function parsePresentationChartExternalData(options = {}) {
  const tag = externalDataTag(options.chartXml);
  if (!tag) return undefined;
  const id = relationshipId(tag);
  const relationship = (options.relationships || []).find((item) => item.id === id);
  if (!id || !relationship) throw new Error("Presentation chart externalData references a missing relationship.");
  if (!String(relationship.type || "").endsWith("/package")) throw new Error("Presentation chart externalData relationship must use the package relationship type.");
  const autoUpdate = autoUpdateFromChartXml(options.chartXml);
  if (String(relationship.targetMode || "").toLowerCase() === "external") return normalizePresentationChartExternalData({ uri: relationship.target, autoUpdate });
  const target = options.resolveTarget(options.chartPath, relationship.target);
  const bytes = await options.readPart(target);
  if (!bytes) throw new Error(`Presentation chart externalData workbook part is missing: ${target}.`);
  return normalizePresentationChartExternalData({ workbook: bytes, autoUpdate });
}

function chartRelationshipEntries(bytesByPath, chartPath) {
  const relsPath = relationshipPartPath(chartPath);
  const xml = bytesByPath.has(relsPath) ? decoder.decode(bytesByPath.get(relsPath)) : "";
  return [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*\/?>/g)].map((match) => {
    const attrs = xmlAttributes(match[0]);
    return { path: relsPath, id: attrs.Id, type: attrs.Type, target: attrs.Target, targetMode: attrs.TargetMode };
  });
}

function declaredContentType(contentTypes, partPath) {
  const extension = path.posix.extname(partPath).slice(1).toLowerCase();
  return contentTypes.overrides.get(partPath) || contentTypes.defaults.get(extension);
}

export function validatePresentationChartExternalDataPackageSemantics({ bytesByPath, contentTypes }) {
  const issues = [];
  const chartPaths = [...bytesByPath.keys()].filter((partPath) => declaredContentType(contentTypes, partPath) === CHART_CONTENT_TYPE);
  for (const chartPath of chartPaths) {
    const chartXml = decoder.decode(bytesByPath.get(chartPath));
    const tag = externalDataTag(chartXml);
    const id = relationshipId(tag);
    const entries = chartRelationshipEntries(bytesByPath, chartPath);
    const packageEntries = entries.filter((entry) => String(entry.type || "").endsWith("/package"));
    const hasFormulaReferences = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?errBars\\b[\\s\\S]*?<(?:[A-Za-z_][\\w.-]*:)?numRef\\b[\\s\\S]*?<(?:[A-Za-z_][\\w.-]*:)?f\\b`, "i").test(chartXml);
    if (hasFormulaReferences && !tag) issues.push(issue("pptxChartFormulaExternalDataMissing", `PPTX chart ${chartPath} uses formula-backed error bars without externalData.`, { path: chartPath }));
    if (tag) {
      const entry = entries.find((candidate) => candidate.id === id);
      if (!id || !entry) issues.push(issue("pptxChartExternalDataRelationshipMissing", `PPTX chart ${chartPath} externalData references a missing relationship.`, { path: chartPath, relationshipId: id }));
      else if (!String(entry.type || "").endsWith("/package")) issues.push(issue("pptxChartExternalDataRelationshipTypeInvalid", `PPTX chart ${chartPath} externalData relationship ${id} must use the package relationship type.`, { path: entry.path, source: chartPath, relationshipId: id, relationshipType: entry.type }));
      else if (String(entry.targetMode || "").toLowerCase() !== "external") {
        const target = resolveTarget(chartPath, entry.target);
        if (bytesByPath.has(target)) {
          const contentType = declaredContentType(contentTypes, target);
          if (contentType !== PRESENTATION_CHART_EXTERNAL_DATA_CONTENT_TYPE) issues.push(issue("pptxChartExternalDataContentTypeInvalid", `PPTX embedded chart workbook ${target} must use the XLSX content type.`, { path: target, source: chartPath, relationshipId: id, contentType, expectedContentType: PRESENTATION_CHART_EXTERNAL_DATA_CONTENT_TYPE }));
          const bytes = bytesByPath.get(target);
          if (bytes.byteLength < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) issues.push(issue("pptxChartExternalDataWorkbookInvalid", `PPTX embedded chart workbook ${target} must be an OOXML ZIP package.`, { path: target, source: chartPath, relationshipId: id }));
        }
      }
    }
    for (const entry of packageEntries) if (!tag || entry.id !== id) issues.push(issue("pptxChartExternalDataRelationshipOrphaned", `PPTX chart ${chartPath} package relationship ${entry.id || "(unknown)"} is not referenced by externalData.`, { path: entry.path, source: chartPath, relationshipId: entry.id }));
  }
  return issues;
}
