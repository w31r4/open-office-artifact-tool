import crypto from "node:crypto";
import fs from "node:fs/promises";

import mupdf from "mupdf";

import { toUint8Array } from "../shared/binary.mjs";
import { FileBlob } from "../shared/file-blob.mjs";

export const MUPDF_VERSION = "1.28.0";

const PDF_MIME = "application/pdf";
const DEFAULT_LIMITS = Object.freeze({
  maxBytes: 256 * 1024 * 1024,
  maxPages: 10_000,
  maxObjects: 1_000_000,
  maxAnnotations: 100_000,
  maxLinks: 100_000,
  maxImages: 10_000,
  maxImagePixels: 40_000_000,
  maxTotalImagePixels: 100_000_000,
  maxRenderPixels: 40_000_000,
  maxTotalImageBytes: 64 * 1024 * 1024,
});

const METADATA_KEYS = Object.freeze({
  author: mupdf.Document.META_INFO_AUTHOR,
  title: mupdf.Document.META_INFO_TITLE,
  subject: mupdf.Document.META_INFO_SUBJECT,
  keywords: mupdf.Document.META_INFO_KEYWORDS,
  creator: mupdf.Document.META_INFO_CREATOR,
  producer: mupdf.Document.META_INFO_PRODUCER,
  creationDate: mupdf.Document.META_INFO_CREATIONDATE,
  modificationDate: mupdf.Document.META_INFO_MODIFICATIONDATE,
});

const INCREMENTAL_DESTRUCTIVE_OPERATIONS = new Set([
  "add_text_annotation",
  "delete_annotation",
  "update_annotation",
  "delete_page",
  "delete_embedded_file",
  "add_link",
  "delete_link",
  "update_link",
  "redact_text",
  "redact_rect",
]);

const PAGE_ROTATIONS = new Set([0, 90, 180, 270]);
const ANNOTATION_ID_PREFIX = "mupdf-annotation";
const ANNOTATION_EXPECTATION_FIELDS = new Set(["type", "contents", "name", "author", "subject", "rect"]);
const ANNOTATION_PATCH_FIELDS = new Set(["contents", "author", "subject"]);
const WIDGET_ID_PREFIX = "mupdf-widget";
const FORM_FIELD_ID_PREFIX = "mupdf-form-field";
const FORM_FIELD_EXPECTATION_FIELDS = new Set(["name", "type", "value", "readOnly", "options", "exportOptions", "widgets"]);
const LINK_ID_PREFIX = "mupdf-link";
const LINK_EXPECTATION_FIELDS = new Set(["url", "bbox", "external"]);
const LINK_PATCH_FIELDS = new Set(["url"]);
const PAGE_EXPECTATION_FIELDS = new Set(["bbox", "rotation"]);
const SAFE_NATIVE_LINK_PROTOCOLS = new Set(["http:", "https:", "mailto:"]);
const TEXT_ANNOTATION_ANCHOR_SIZE = 20;

function limitsFor(options = {}) {
  const topLevel = Object.fromEntries(Object.keys(DEFAULT_LIMITS)
    .filter((name) => options[name] !== undefined)
    .map((name) => [name, options[name]]));
  const limits = { ...DEFAULT_LIMITS, ...topLevel, ...(options.limits || {}) };
  for (const [name, value] of Object.entries(limits)) {
    if (!Number.isFinite(value) || value <= 0) throw new Error(`PDF limit ${name} must be a positive finite number.`);
  }
  return limits;
}

async function bytesFromInput(input, limits) {
  if (typeof input === "string") {
    const stat = await fs.stat(input);
    if (stat.size > limits.maxBytes) throw new Error(`PDF exceeds maxBytes (${stat.size} > ${limits.maxBytes}).`);
    return new Uint8Array(await fs.readFile(input));
  }
  if (input instanceof FileBlob || (input && typeof input.arrayBuffer === "function")) {
    if (Number.isFinite(input.size) && input.size > limits.maxBytes) throw new Error(`PDF exceeds maxBytes (${input.size} > ${limits.maxBytes}).`);
    return new Uint8Array(await input.arrayBuffer());
  }
  const bytes = toUint8Array(input);
  if (bytes.byteLength > limits.maxBytes) throw new Error(`PDF exceeds maxBytes (${bytes.byteLength} > ${limits.maxBytes}).`);
  return bytes;
}

function copyBytes(value) {
  const bytes = value?.asUint8Array ? value.asUint8Array() : toUint8Array(value);
  return new Uint8Array(bytes);
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function pdfRectToBbox(rect) {
  const [x0, y0, x1, y1] = rect.map(Number);
  return [x0, y0, Math.max(0, x1 - x0), Math.max(0, y1 - y0)];
}

function bboxToPdfRect(bbox, label = "rectangle") {
  if (!Array.isArray(bbox) || bbox.length !== 4 || !bbox.every((value) => Number.isFinite(Number(value)))) {
    throw new Error(`${label} must be [x, y, width, height].`);
  }
  const [x, y, width, height] = bbox.map(Number);
  if (width <= 0 || height <= 0) throw new Error(`${label} width and height must be positive.`);
  return [x, y, x + width, y + height];
}

function pdfRectContains(outer, inner) {
  return inner[0] >= outer[0]
    && inner[1] >= outer[1]
    && inner[2] <= outer[2]
    && inner[3] <= outer[3];
}

function rawPageBox(page, name) {
  let pageObject;
  let box;
  try {
    pageObject = page.getObject();
    box = pageObject.getInheritable(name);
    if (!box.isArray() || box.length !== 4) return undefined;
    const values = [];
    for (let index = 0; index < box.length; index += 1) {
      let value;
      try {
        value = box.get(index);
        if (!value.isNumber()) return undefined;
        const number = value.asNumber();
        if (!Number.isFinite(number)) return undefined;
        values.push(number);
      } finally {
        if (value && value !== mupdf.PDFObject.Null) value.destroy();
      }
    }
    return values;
  } finally {
    if (box && box !== mupdf.PDFObject.Null) box.destroy();
    pageObject?.destroy();
  }
}

function rawPageRotation(page) {
  let pageObject;
  let rotation;
  try {
    pageObject = page.getObject();
    rotation = pageObject.getInheritable("Rotate");
    if (rotation.isNull()) return 0;
    if (!rotation.isNumber()) return undefined;
    const value = rotation.asNumber();
    return Number.isInteger(value) ? ((value % 360) + 360) % 360 : undefined;
  } finally {
    if (rotation && rotation !== mupdf.PDFObject.Null) rotation.destroy();
    pageObject?.destroy();
  }
}

function pdfRectsEqual(left, right, tolerance = 0.001) {
  return left.length === right.length
    && left.every((value, index) => Math.abs(value - right[index]) <= tolerance);
}

function hitRect(quads) {
  const points = quads.flatMap((quad) => Array.isArray(quad) && quad.length === 8 ? [quad] : quad);
  const xs = points.flatMap((quad) => [quad[0], quad[2], quad[4], quad[6]]).map(Number);
  const ys = points.flatMap((quad) => [quad[1], quad[3], quad[5], quad[7]]).map(Number);
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
}

function pageIndexFor(document, operation = {}, fallback = undefined) {
  const raw = operation.pageIndex ?? (operation.page == null ? fallback : Number(operation.page) - 1);
  const index = Number(raw);
  if (!Number.isInteger(index) || index < 0 || index >= document.countPages()) {
    throw new Error(`page must identify an existing 1-based page between 1 and ${document.countPages()}.`);
  }
  return index;
}

function pageFor(document, operation, fallback) {
  const index = pageIndexFor(document, operation, fallback);
  return { index, page: document.loadPage(index) };
}

function metadataFor(document) {
  return Object.fromEntries(Object.entries(METADATA_KEYS)
    .map(([name, key]) => [name, document.getMetaData(key)])
    .filter(([, value]) => value !== undefined && value !== ""));
}

function annotationXref(annotation) {
  let object;
  try {
    object = annotation.getObject();
    return pdfObjectXref(object);
  } catch {
    return undefined;
  } finally {
    object?.destroy();
  }
}

function pdfObjectXref(object) {
  try {
    if (!object?.isIndirect()) return undefined;
    const xref = object.asIndirect();
    return Number.isSafeInteger(xref) && xref > 0 ? xref : undefined;
  } catch {
    return undefined;
  }
}

function annotationId(pageNumber, xref) {
  return `${ANNOTATION_ID_PREFIX}-${pageNumber}-${xref}`;
}

function widgetId(pageNumber, xref) {
  return `${WIDGET_ID_PREFIX}-${pageNumber}-${xref}`;
}

function formFieldId(xref) {
  return `${FORM_FIELD_ID_PREFIX}-${xref}`;
}

function parseFormFieldId(value, operationName = "update_form_field") {
  const match = new RegExp(`^${FORM_FIELD_ID_PREFIX}-(\\d+)$`, "u").exec(String(value || ""));
  if (!match) throw new Error(`${operationName} formFieldId must be a ${FORM_FIELD_ID_PREFIX}-<xref> locator returned by PdfFile.inspectPdf.`);
  const xref = Number(match[1]);
  if (!Number.isSafeInteger(xref) || xref < 1) {
    throw new Error(`${operationName} formFieldId contains an invalid xref number.`);
  }
  return { xref };
}

function parseAnnotationId(value, operationName = "delete_annotation") {
  const match = new RegExp(`^${ANNOTATION_ID_PREFIX}-(\\d+)-(\\d+)$`, "u").exec(String(value || ""));
  if (!match) throw new Error(`${operationName} annotationId must be a ${ANNOTATION_ID_PREFIX}-<page>-<xref> locator returned by PdfFile.inspectPdf.`);
  const page = Number(match[1]);
  const xref = Number(match[2]);
  if (!Number.isSafeInteger(page) || page < 1 || !Number.isSafeInteger(xref) || xref < 1) {
    throw new Error(`${operationName} annotationId contains an invalid page or xref number.`);
  }
  return { page, xref };
}

function nativeAnnotation(annotation, pageNumber) {
  const xref = annotationXref(annotation);
  const record = {
    id: xref ? annotationId(pageNumber, xref) : undefined,
    xref,
    type: annotation.getType(),
    rect: annotation.hasRect() ? pdfRectToBbox(annotation.getRect()) : undefined,
    contents: annotation.getContents() || undefined,
    name: annotation.getName() || undefined,
    author: annotation.hasAuthor() ? annotation.getAuthor() || undefined : undefined,
    subject: annotation.hasSubject() ? annotation.getSubject() || undefined : undefined,
    flags: annotation.getFlags(),
  };
  return Object.fromEntries(Object.entries(record).filter(([, value]) => value !== undefined));
}

function consumeAnnotationBudget(budget, count, pageNumber) {
  budget.used += count;
  if (budget.used > budget.maxAnnotations) {
    throw new Error(`MuPDF annotations exceed maxAnnotations (${budget.used} > ${budget.maxAnnotations}) while reading page ${pageNumber}.`);
  }
}

function annotationExpectation(value, operationName = "delete_annotation") {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error(`${operationName} expected must be an object with at least one source snapshot field.`);
  }
  for (const name of Object.keys(value)) {
    if (!ANNOTATION_EXPECTATION_FIELDS.has(name)) {
      throw new Error(`${operationName} expected contains unsupported snapshot field: ${name}.`);
    }
  }
  const expected = {};
  for (const name of ["type", "contents", "name", "author", "subject"]) {
    if (value[name] === undefined) continue;
    if (typeof value[name] !== "string") throw new Error(`${operationName} expected.${name} must be a string.`);
    expected[name] = value[name];
  }
  if (value.rect !== undefined) {
    bboxToPdfRect(value.rect, `${operationName} expected.rect`);
    expected.rect = value.rect.map(Number);
  }
  if (!Object.keys(expected).length) {
    throw new Error(`${operationName} expected must include at least one of type, contents, name, author, subject, or rect.`);
  }
  return expected;
}

function annotationExpectationMismatch(actual, expected) {
  for (const name of ["type", "contents", "name", "author", "subject"]) {
    if (expected[name] !== undefined && actual[name] !== expected[name]) return name;
  }
  if (expected.rect !== undefined) {
    if (!actual.rect || !pdfRectsEqual(bboxToPdfRect(actual.rect), bboxToPdfRect(expected.rect))) return "rect";
  }
  return undefined;
}

function annotationPatch(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("update_annotation patch must be an object with at least one mutable field.");
  }
  for (const name of Object.keys(value)) {
    if (!ANNOTATION_PATCH_FIELDS.has(name)) {
      throw new Error(`update_annotation patch contains unsupported field: ${name}.`);
    }
  }
  const patch = {};
  for (const name of ["contents", "author", "subject"]) {
    if (value[name] === undefined) continue;
    if (typeof value[name] !== "string" || !value[name].length) {
      throw new Error(`update_annotation patch.${name} must be a non-empty string.`);
    }
    patch[name] = value[name];
  }
  if (!Object.keys(patch).length) {
    throw new Error("update_annotation patch must include at least one of contents, author, or subject.");
  }
  return patch;
}

function linkId(pageNumber, record) {
  const fingerprint = sha256(Buffer.from(JSON.stringify({
    page: pageNumber,
    url: record.url,
    bbox: record.bbox,
    external: record.external,
  }), "utf8"));
  return `${LINK_ID_PREFIX}-${pageNumber}-${fingerprint}`;
}

function parseLinkId(value, operationName = "delete_link") {
  const match = new RegExp(`^${LINK_ID_PREFIX}-(\\d+)-([a-f0-9]{64})$`, "u").exec(String(value || ""));
  if (!match) throw new Error(`${operationName} linkId must be a ${LINK_ID_PREFIX}-<page>-<fingerprint> locator returned by PdfFile.inspectPdf.`);
  const page = Number(match[1]);
  if (!Number.isSafeInteger(page) || page < 1) throw new Error(`${operationName} linkId contains an invalid page number.`);
  return { page, fingerprint: match[2] };
}

function linkExpectation(value, operationName = "delete_link") {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error(`${operationName} expected must be an object with at least one source snapshot field.`);
  }
  for (const name of Object.keys(value)) {
    if (!LINK_EXPECTATION_FIELDS.has(name)) {
      throw new Error(`${operationName} expected contains unsupported snapshot field: ${name}.`);
    }
  }
  const expected = {};
  if (value.url !== undefined) {
    if (typeof value.url !== "string") throw new Error(`${operationName} expected.url must be a string.`);
    expected.url = value.url;
  }
  if (value.bbox !== undefined) {
    bboxToPdfRect(value.bbox, `${operationName} expected.bbox`);
    expected.bbox = value.bbox.map(Number);
  }
  if (value.external !== undefined) {
    if (typeof value.external !== "boolean") throw new Error(`${operationName} expected.external must be a boolean.`);
    expected.external = value.external;
  }
  if (!Object.keys(expected).length) {
    throw new Error(`${operationName} expected must include at least one of url, bbox, or external.`);
  }
  return expected;
}

function linkExpectationMismatch(actual, expected) {
  if (expected.url !== undefined && actual.url !== expected.url) return "url";
  if (expected.external !== undefined && actual.external !== expected.external) return "external";
  if (expected.bbox !== undefined && !pdfRectsEqual(bboxToPdfRect(actual.bbox), bboxToPdfRect(expected.bbox))) return "bbox";
  return undefined;
}

function linkPatch(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("update_link patch must be an object containing url.");
  }
  for (const name of Object.keys(value)) {
    if (!LINK_PATCH_FIELDS.has(name)) {
      throw new Error(`update_link patch contains unsupported field: ${name}.`);
    }
  }
  return { url: nativeLinkUri(value.url, "update_link patch.url") };
}

function nativeLinkUri(value, label = "link url") {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error(`${label} must be a non-empty string.`);
  }
  const url = value.trim();
  if (url.startsWith("#")) {
    if (url.length === 1 || /\s|[\u0000-\u001f\u007f]/u.test(url)) throw new Error(`${label} must be a safe internal destination or an absolute http, https, or mailto URL.`);
    return url;
  }
  let parsed;
  try { parsed = new URL(url); } catch { throw new Error(`${label} must be a safe internal destination or an absolute http, https, or mailto URL.`); }
  if (!SAFE_NATIVE_LINK_PROTOCOLS.has(parsed.protocol)) {
    throw new Error(`${label} uses unsupported protocol ${parsed.protocol || "(none)"}; use an internal destination or an absolute http, https, or mailto URL.`);
  }
  return url;
}

function pageExpectation(value, operationName) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error(`${operationName} expectedPage must be an inspect-returned object with bbox and rotation.`);
  }
  for (const name of Object.keys(value)) {
    if (!PAGE_EXPECTATION_FIELDS.has(name)) {
      throw new Error(`${operationName} expectedPage contains unsupported snapshot field: ${name}.`);
    }
  }
  if (value.bbox === undefined || value.rotation === undefined) {
    throw new Error(`${operationName} expectedPage must include both bbox and rotation from the inspected mupdfPage record.`);
  }
  const bbox = bboxToPdfRect(value.bbox, `${operationName} expectedPage.bbox`);
  const rotation = Number(value.rotation);
  if (!Number.isInteger(rotation) || !PAGE_ROTATIONS.has(rotation)) {
    throw new Error(`${operationName} expectedPage.rotation must be 0, 90, 180, or 270.`);
  }
  return { bbox: pdfRectToBbox(bbox), rotation };
}

function nativePageSnapshot(page, operationName) {
  const bbox = pdfRectToBbox(page.getBounds("CropBox"));
  const rotation = rawPageRotation(page);
  if (rotation === undefined || !PAGE_ROTATIONS.has(rotation)) {
    throw new Error(`${operationName} requires a finite right-angle inherited Rotate value.`);
  }
  return { bbox, rotation };
}

function pageExpectationMismatch(actual, expected) {
  if (!pdfRectsEqual(bboxToPdfRect(actual.bbox), bboxToPdfRect(expected.bbox))) return "bbox";
  if (actual.rotation !== expected.rotation) return "rotation";
  return undefined;
}

function textAnnotationPoint(value) {
  if (!Array.isArray(value) || value.length !== 2 || !value.every((coordinate) => Number.isFinite(Number(coordinate)))) {
    throw new Error("add_text_annotation point must be [x, y] in the inspected raw unrotated PDF page coordinates.");
  }
  return value.map(Number);
}

function textAnnotationRequest(operation) {
  if (operation.text !== undefined) {
    throw new Error("add_text_annotation uses contents, not the legacy text alias.");
  }
  if (operation.bbox !== undefined || operation.rect !== undefined) {
    throw new Error("add_text_annotation uses one source-bound point, not a requested rectangle; MuPDF owns the native Text-note icon rectangle.");
  }
  if (typeof operation.contents !== "string" || !operation.contents.length) {
    throw new Error("add_text_annotation requires non-empty contents.");
  }
  const request = { contents: operation.contents, point: textAnnotationPoint(operation.point) };
  for (const name of ["author", "subject"]) {
    if (operation[name] === undefined) continue;
    if (typeof operation[name] !== "string" || !operation[name].length) {
      throw new Error(`add_text_annotation ${name} must be a non-empty string when supplied.`);
    }
    request[name] = operation[name];
  }
  return request;
}

function annotationAddedAtPointMismatch(actual, request) {
  if (actual.type !== "Text") return "type";
  for (const name of ["contents", "author", "subject"]) {
    if (request[name] !== undefined && actual[name] !== request[name]) return name;
  }
  if (!actual.rect || Math.abs(actual.rect[0] - request.point[0]) > 0.001 || Math.abs(actual.rect[1] - request.point[1]) > 0.001) {
    return "point";
  }
  return undefined;
}

function widgetIdentity(widget) {
  let object;
  let parent;
  try {
    object = widget.getObject();
    const xref = pdfObjectXref(object);
    if (!xref) return {};
    parent = object.get("Parent");
    if (parent.isNull()) return { xref, fieldXref: xref };
    const fieldXref = pdfObjectXref(parent);
    return fieldXref ? { xref, fieldXref } : { xref };
  } catch {
    return {};
  } finally {
    parent?.destroy();
    object?.destroy();
  }
}

function nativeWidget(widget, pageNumber) {
  const { xref, fieldXref } = widgetIdentity(widget);
  const record = {
    id: xref ? widgetId(pageNumber, xref) : undefined,
    xref,
    formFieldId: fieldXref ? formFieldId(fieldXref) : undefined,
    fieldXref,
    type: widget.getFieldType(),
    name: widget.getName(),
    label: widget.getLabel() || undefined,
    value: widget.getValue(),
    rect: widget.hasRect() ? pdfRectToBbox(widget.getRect()) : undefined,
    readOnly: widget.isReadOnly(),
    options: widget.isChoice() ? widget.getOptions() : undefined,
    exportOptions: widget.isChoice() ? widget.getOptions(true) : undefined,
    maxLength: widget.isText() ? widget.getMaxLen() || undefined : undefined,
    password: widget.isText() && widget.isPassword() ? true : undefined,
  };
  return Object.fromEntries(Object.entries(record).filter(([, value]) => value !== undefined));
}

function sameStringArrays(left, right) {
  return Array.isArray(left)
    && Array.isArray(right)
    && left.length === right.length
    && left.every((value, index) => value === right[index]);
}

function widgetSnapshot(widget) {
  return {
    id: widget.id,
    page: widget.page,
    xref: widget.xref,
    rect: widget.rect,
  };
}

function nativeFormFieldRecord(widgets) {
  const ordered = [...widgets].sort((left, right) => left.page - right.page || left.xref - right.xref);
  const first = ordered[0];
  if (!first?.formFieldId || !first.fieldXref) return undefined;
  const sameField = ordered.every((widget) => widget.formFieldId === first.formFieldId
    && widget.name === first.name
    && widget.type === first.type
    && widget.value === first.value
    && widget.readOnly === first.readOnly
    && sameStringArrays(widget.options || [], first.options || [])
    && sameStringArrays(widget.exportOptions || [], first.exportOptions || []));
  const record = {
    kind: "mupdfFormField",
    id: first.formFieldId,
    xref: first.fieldXref,
    name: first.name,
    type: first.type,
    label: ordered.every((widget) => widget.label === first.label) ? first.label : undefined,
    value: first.value,
    readOnly: first.readOnly,
    options: first.options,
    exportOptions: first.exportOptions,
    widgets: ordered.map(widgetSnapshot),
  };
  if (!sameField) {
    return {
      ...record,
      unsupportedReason: "Widgets sharing this form-field locator disagree on semantic state; route the original PDF to the explicit pypdf form workflow.",
    };
  }
  return {
    ...record,
    snapshot: Object.fromEntries(Object.entries({
      name: record.name,
      type: record.type,
      value: record.value,
      readOnly: record.readOnly,
      options: record.options,
      exportOptions: record.exportOptions,
      widgets: record.widgets,
    }).filter(([, value]) => value !== undefined)),
  };
}

function collectNativeFormFields(widgetRecords) {
  const grouped = new Map();
  for (const widget of widgetRecords) {
    if (!widget.formFieldId) continue;
    const current = grouped.get(widget.formFieldId) || [];
    current.push(widget);
    grouped.set(widget.formFieldId, current);
  }
  return [...grouped.values()]
    .map(nativeFormFieldRecord)
    .filter(Boolean)
    .sort((left, right) => left.id.localeCompare(right.id));
}

function formFieldExpectation(value, operationName = "update_form_field") {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error(`${operationName} expected must be the snapshot returned by a mupdfFormField inspection record.`);
  }
  for (const name of Object.keys(value)) {
    if (!FORM_FIELD_EXPECTATION_FIELDS.has(name)) {
      throw new Error(`${operationName} expected contains unsupported snapshot field: ${name}.`);
    }
  }
  for (const name of ["name", "type", "value"]) {
    if (typeof value[name] !== "string") throw new Error(`${operationName} expected.${name} must be a string.`);
  }
  if (typeof value.readOnly !== "boolean") throw new Error(`${operationName} expected.readOnly must be a boolean.`);
  if (!Array.isArray(value.widgets) || value.widgets.length !== 1) {
    throw new Error(`${operationName} expected.widgets must contain exactly one inspect-returned widget; shared-widget fields require the explicit pypdf form workflow.`);
  }
  const widgets = value.widgets.map((widget, index) => {
    if (!widget || typeof widget !== "object" || Array.isArray(widget)) {
      throw new Error(`${operationName} expected.widgets[${index}] must be an inspect-returned widget snapshot.`);
    }
    const allowed = new Set(["id", "page", "xref", "rect"]);
    for (const name of Object.keys(widget)) {
      if (!allowed.has(name)) throw new Error(`${operationName} expected.widgets[${index}] contains unsupported snapshot field: ${name}.`);
    }
    const page = Number(widget.page);
    const xref = Number(widget.xref);
    if (!Number.isSafeInteger(page) || page < 1 || !Number.isSafeInteger(xref) || xref < 1 || widget.id !== widgetId(page, xref)) {
      throw new Error(`${operationName} expected.widgets[${index}] must preserve one valid mupdf-widget-<page>-<xref> locator.`);
    }
    const rect = bboxToPdfRect(widget.rect, `${operationName} expected.widgets[${index}].rect`);
    return { id: widget.id, page, xref, rect: pdfRectToBbox(rect) };
  });
  const expected = {
    name: value.name,
    type: value.type,
    value: value.value,
    readOnly: value.readOnly,
    widgets,
  };
  for (const name of ["options", "exportOptions"]) {
    if (value[name] === undefined) continue;
    if (!Array.isArray(value[name]) || value[name].some((option) => typeof option !== "string")) {
      throw new Error(`${operationName} expected.${name} must be an array of strings.`);
    }
    expected[name] = [...value[name]];
  }
  return expected;
}

function formFieldExpectationMismatch(actual, expected) {
  for (const name of ["name", "type", "value", "readOnly"]) {
    if (actual[name] !== expected[name]) return name;
  }
  for (const name of ["options", "exportOptions"]) {
    if (expected[name] !== undefined && !sameStringArrays(actual[name] || [], expected[name])) return name;
  }
  if (!sameStringArrays(actual.widgets.map((widget) => widget.id), expected.widgets.map((widget) => widget.id))) return "widgets";
  for (let index = 0; index < actual.widgets.length; index += 1) {
    const left = actual.widgets[index];
    const right = expected.widgets[index];
    if (left.page !== right.page || left.xref !== right.xref || !pdfRectsEqual(bboxToPdfRect(left.rect), bboxToPdfRect(right.rect))) return "widgets";
  }
  return undefined;
}

function collectNativeFormField(document, fieldXref, limits = DEFAULT_LIMITS) {
  const widgets = [];
  let widgetCount = 0;
  for (let index = 0; index < document.countPages(); index += 1) {
    const page = document.loadPage(index);
    const pageWidgets = page.getWidgets();
    try {
      widgetCount += pageWidgets.length;
      if (widgetCount > limits.maxAnnotations) {
        throw new Error(`MuPDF annotations/widgets exceed maxAnnotations (${widgetCount} > ${limits.maxAnnotations}) while locating a source-bound form field.`);
      }
      for (const widget of pageWidgets) {
        const record = nativeWidget(widget, index + 1);
        if (record.fieldXref === fieldXref) widgets.push({ page: index + 1, ...record });
      }
    } finally {
      pageWidgets.forEach((widget) => widget.destroy());
      page.destroy();
    }
  }
  return nativeFormFieldRecord(widgets);
}

function withSingleNativeFormWidget(document, field, mutate) {
  const reference = field.widgets[0];
  const page = document.loadPage(reference.page - 1);
  const widgets = page.getWidgets();
  let target;
  try {
    target = widgets.find((widget) => {
      const record = nativeWidget(widget, reference.page);
      return record.id === reference.id && record.formFieldId === field.id;
    });
    if (!target) {
      throw new Error(`update_form_field could not resolve ${reference.id} from ${field.id}; re-inspect the current source PDF before retrying.`);
    }
    return mutate({ page, target });
  } finally {
    widgets.forEach((widget) => widget.destroy());
    page.destroy();
  }
}

function checkboxIsChecked(value) {
  return !/^\s*(?:|off)\s*$/iu.test(String(value ?? ""));
}

function sourceBoundFormFieldValue(widget, field, operation) {
  if (!Object.hasOwn(operation, "value")) throw new Error("update_form_field requires a value property.");
  if (widget.isText()) {
    if (widget.isPassword()) throw new Error(`update_form_field refuses password field ${field.name}; use an explicit specialist provider after a security review.`);
    if (typeof operation.value !== "string") throw new Error(`update_form_field text field ${field.name} requires a string value.`);
    const maxLength = widget.getMaxLen();
    if (maxLength > 0 && operation.value.length > maxLength) {
      throw new Error(`update_form_field text value exceeds the source field maxLength (${operation.value.length} > ${maxLength}).`);
    }
    const accepted = widget.setTextValue(operation.value);
    if (!accepted) throw new Error(`MuPDF rejected value for source-bound text field ${field.name}.`);
    return { kind: "text", value: operation.value };
  }
  if (widget.isChoice()) {
    if (widget.isListBox() || (widget.getFieldFlags() & mupdf.PDFWidget.CH_FIELD_IS_MULTI_SELECT)) {
      throw new Error(`update_form_field does not support list or multi-select field ${field.name}; use the explicit pypdf form workflow.`);
    }
    if (typeof operation.value !== "string") throw new Error(`update_form_field choice field ${field.name} requires a string value.`);
    const options = widget.getOptions();
    const exportOptions = widget.getOptions(true);
    if (!sameStringArrays(options, exportOptions)) {
      throw new Error(`update_form_field refuses choice field ${field.name} because its display and export values differ; use the explicit pypdf form workflow.`);
    }
    if (!options.includes(operation.value)) {
      throw new Error(`update_form_field choice value ${JSON.stringify(operation.value)} is not an inspected option for ${field.name}.`);
    }
    const accepted = widget.setChoiceValue(operation.value);
    if (!accepted) throw new Error(`MuPDF rejected value for source-bound choice field ${field.name}.`);
    return { kind: "choice", value: operation.value };
  }
  if (widget.isCheckbox()) {
    const desired = normalizeCheckboxValue(operation.value, field.name, "update_form_field");
    const accepted = desired === checkboxIsChecked(widget.getValue()) ? 1 : widget.toggle();
    if (!accepted) throw new Error(`MuPDF rejected value for source-bound checkbox field ${field.name}.`);
    return { kind: "checkbox", value: desired };
  }
  if (widget.isRadioButton()) {
    throw new Error(`update_form_field does not expose a trustworthy radio export-value mapping for field ${field.name}; use the explicit pypdf form workflow.`);
  }
  throw new Error(`update_form_field does not support widget type ${widget.getFieldType()} for field ${field.name}.`);
}

function sourceBoundFormValueMismatch(field, requested) {
  if (requested.kind === "checkbox") return checkboxIsChecked(field.value) === requested.value ? undefined : "value";
  return field.value === requested.value ? undefined : "value";
}

function applySourceBoundFormFieldUpdate(document, operation, context = {}) {
  if (!/^[a-f0-9]{64}$/u.test(String(operation.sourceSha256 || "")) || operation.sourceSha256 !== context.sourceSha256) {
    throw new Error("update_form_field sourceSha256 must exactly match PdfFile.inspectPdf(...).summary.sourceSha256 for the current input bytes.");
  }
  const locator = parseFormFieldId(operation.formFieldId, "update_form_field");
  const expected = formFieldExpectation(operation.expected, "update_form_field");
  const matched = collectNativeFormField(document, locator.xref, context.limits);
  if (!matched) {
    throw new Error(`update_form_field could not find source-bound form field ${operation.formFieldId}. Re-inspect the current source PDF before retrying.`);
  }
  if (matched.id !== operation.formFieldId || !matched.snapshot) {
    throw new Error(`update_form_field locator ${operation.formFieldId} did not resolve to one consistent inspectable form field.`);
  }
  const mismatch = formFieldExpectationMismatch(matched.snapshot, expected);
  if (mismatch) {
    throw new Error(`update_form_field precondition ${mismatch} did not match ${operation.formFieldId}; refusing a stale or ambiguous mutation.`);
  }
  if (matched.readOnly) throw new Error(`update_form_field refuses read-only form field ${matched.name}.`);
  if (matched.widgets.length !== 1) {
    throw new Error(`update_form_field refuses shared-widget form field ${matched.name}; use the explicit pypdf form workflow.`);
  }
  const requested = withSingleNativeFormWidget(document, matched, ({ page, target }) => {
    const value = sourceBoundFormFieldValue(target, matched, operation);
    target.update();
    page.update();
    return value;
  });
  const updated = collectNativeFormField(document, locator.xref, context.limits);
  if (!updated || updated.id !== matched.id || !updated.snapshot || updated.widgets.length !== 1) {
    throw new Error(`MuPDF did not retain one uniquely addressable source-bound form field ${operation.formFieldId}; refusing to save an ambiguous update.`);
  }
  const structureMismatch = formFieldExpectationMismatch(updated.snapshot, { ...matched.snapshot, value: updated.snapshot.value });
  if (structureMismatch) {
    throw new Error(`MuPDF changed form-field structure ${structureMismatch} while updating ${operation.formFieldId}; refusing to save an ambiguous update.`);
  }
  const valueMismatch = sourceBoundFormValueMismatch(updated, requested);
  if (valueMismatch) {
    throw new Error(`MuPDF did not preserve update_form_field ${valueMismatch} for ${operation.formFieldId}; refusing to save an ambiguous update.`);
  }
  return {
    type: "update_form_field",
    formFieldId: matched.id,
    xref: locator.xref,
    matched: matched.snapshot,
    value: requested.value,
    updated: updated.snapshot,
  };
}

function nativeLink(link, pageNumber) {
  const record = { url: String(link.getURI() || ""), bbox: pdfRectToBbox(link.getBounds()), external: link.isExternal() };
  return { id: linkId(pageNumber, record), ...record };
}

function consumeLinkBudget(budget, count, pageNumber) {
  budget.used += count;
  if (budget.used > budget.maxLinks) {
    throw new Error(`MuPDF links exceed maxLinks (${budget.used} > ${budget.maxLinks}) while reading page ${pageNumber}.`);
  }
}

function structuredPage(page, pageNumber, options, imageBudget, annotationBudget, linkBudget) {
  const structured = page.toStructuredText(options.includeImages === false ? "preserve-whitespace" : "preserve-whitespace,preserve-images");
  try {
    let json;
    try { json = JSON.parse(structured.asJSON()); } catch (error) { throw new Error(`MuPDF structured-text JSON failed on page ${pageNumber}: ${error.message}`); }
    const textItems = [];
    for (const block of json.blocks || []) {
      if (block.type !== "text") continue;
      for (const line of block.lines || []) {
        const bbox = line.bbox ? [line.bbox.x, line.bbox.y, line.bbox.w, line.bbox.h].map(Number) : [0, 0, 0, 0];
        textItems.push({
          text: String(line.text || ""),
          bbox,
          fontName: line.font?.name || line.font?.family,
          fontSize: Number(line.font?.size) || undefined,
          bold: line.font?.weight === "bold",
          italic: line.font?.style === "italic",
        });
      }
    }
    const images = [];
    if (options.includeImages !== false) {
      structured.walk({
        onImageBlock(bbox, transform, image) {
          let pixmap;
          let mask;
          let colorSpace;
          try {
            const pixelWidth = image.getWidth();
            const pixelHeight = image.getHeight();
            const pixels = pixelWidth * pixelHeight;
            imageBudget.usedImages += 1;
            imageBudget.usedPixels += pixels;
            if (imageBudget.usedImages > imageBudget.maxImages) throw new Error(`MuPDF extracted images exceed maxImages (${imageBudget.usedImages} > ${imageBudget.maxImages}).`);
            if (pixels > imageBudget.maxImagePixels) throw new Error(`MuPDF image on page ${pageNumber} exceeds maxImagePixels (${pixels} > ${imageBudget.maxImagePixels}).`);
            if (imageBudget.usedPixels > imageBudget.maxTotalImagePixels) throw new Error(`MuPDF extracted images exceed maxTotalImagePixels (${imageBudget.usedPixels} > ${imageBudget.maxTotalImagePixels}).`);
            mask = image.getMask();
            colorSpace = image.getColorSpace();
            pixmap = image.toPixmap();
            const png = pixmap.asPNG();
            const base64 = Buffer.from(png).toString("base64");
            imageBudget.usedBytes += Buffer.byteLength(base64, "ascii");
            if (imageBudget.usedBytes > imageBudget.maxTotalImageBytes) throw new Error(`MuPDF retained image data exceeds maxTotalImageBytes (${imageBudget.usedBytes} > ${imageBudget.maxTotalImageBytes}).`);
            images.push({
              name: `mupdf-image-${pageNumber}-${images.length + 1}`,
              alt: `Imported raster placement ${images.length + 1} on page ${pageNumber}`,
              dataUrl: `data:image/png;base64,${base64}`,
              bbox: pdfRectToBbox(bbox),
              pixelWidth,
              pixelHeight,
              isMask: image.getImageMask(),
              hasSoftMask: Boolean(mask),
              colorSpace: colorSpace?.getName(),
              sourceOperator: "mupdf-structured-text-image",
              transform: [...transform],
            });
          } finally {
            colorSpace?.destroy();
            mask?.destroy();
            pixmap?.destroy();
            image.destroy();
          }
        },
      });
    }
    const bounds = page.getBounds();
    const mediaBox = rawPageBox(page, "MediaBox");
    const cropBox = rawPageBox(page, "CropBox") || mediaBox;
    const annotationObjects = page.getAnnotations();
    const widgetObjects = page.getWidgets();
    const linkObjects = page.getLinks();
    try {
      consumeAnnotationBudget(annotationBudget, annotationObjects.length + widgetObjects.length, pageNumber);
      consumeLinkBudget(linkBudget, linkObjects.length, pageNumber);
      const annotations = annotationObjects.map((annotation) => nativeAnnotation(annotation, pageNumber));
      const widgets = widgetObjects.map((widget) => nativeWidget(widget, pageNumber));
      const links = linkObjects.map((link) => nativeLink(link, pageNumber));
      return {
        width: bounds[2] - bounds[0],
        height: bounds[3] - bounds[1],
        mediaBox: mediaBox && pdfRectToBbox(mediaBox),
        cropBox: cropBox && pdfRectToBbox(cropBox),
        rotation: rawPageRotation(page),
        text: structured.asText().replace(/\s+$/u, ""),
        textItems,
        images,
        links: links.map((link) => ({ id: link.id, text: link.url, url: link.url, bbox: link.bbox })),
        native: { annotations, widgets, links },
      };
    } finally {
      annotationObjects.forEach((annotation) => annotation.destroy());
      widgetObjects.forEach((widget) => widget.destroy());
      linkObjects.forEach((link) => link.destroy());
    }
  } finally {
    structured.destroy();
    page.destroy();
  }
}

export async function openPdfWithMuPdf(input, options = {}) {
  const limits = limitsFor(options);
  const bytes = await bytesFromInput(input, limits);
  if (bytes.byteLength > limits.maxBytes) throw new Error(`PDF exceeds maxBytes (${bytes.byteLength} > ${limits.maxBytes}).`);
  let document;
  try { document = new mupdf.PDFDocument(bytes); } catch (error) { throw new Error(`MuPDF could not open the PDF: ${error.message}`); }
  try {
    if (document.needsPassword()) {
      if (!options.password) throw new Error("PDF requires a password; provide options.password.");
      if (!document.authenticatePassword(String(options.password))) throw new Error("MuPDF rejected the supplied PDF password.");
    }
    const pageCount = document.countPages();
    const objectCount = document.countObjects();
    if (pageCount > limits.maxPages) throw new Error(`PDF exceeds maxPages (${pageCount} > ${limits.maxPages}).`);
    if (objectCount > limits.maxObjects) throw new Error(`PDF exceeds maxObjects (${objectCount} > ${limits.maxObjects}).`);
    return { document, bytes, limits };
  } catch (error) {
    document.destroy();
    throw error;
  }
}

export async function parsePdfWithMuPdf(input, options = {}) {
  const context = input && typeof input === "object" && ("bytes" in input || "input" in input) ? input : undefined;
  const effectiveOptions = { ...(context?.options || {}), ...options };
  const { document, limits } = await openPdfWithMuPdf(context?.bytes || context?.input || input, effectiveOptions);
  const imageBudget = { maxImages: limits.maxImages, maxImagePixels: limits.maxImagePixels, maxTotalImagePixels: limits.maxTotalImagePixels, maxTotalImageBytes: limits.maxTotalImageBytes, usedImages: 0, usedPixels: 0, usedBytes: 0 };
  const annotationBudget = { maxAnnotations: limits.maxAnnotations, used: 0 };
  const linkBudget = { maxLinks: limits.maxLinks, used: 0 };
  try {
    const pages = Array.from({ length: document.countPages() }, (_, index) => structuredPage(document.loadPage(index), index + 1, effectiveOptions, imageBudget, annotationBudget, linkBudget));
    return {
      parser: "mupdf",
      metadata: {
        parser: "mupdf",
        provider: "mupdf",
        providerVersion: MUPDF_VERSION,
        pdf: metadataFor(document),
        repaired: document.wasRepaired(),
        objectCount: document.countObjects(),
      },
      pages,
    };
  } finally {
    document.destroy();
  }
}

export function createMuPdfParser(defaultOptions = {}) {
  return async (context = {}) => parsePdfWithMuPdf(context, { ...defaultOptions, ...(context.options || {}) });
}

export async function inspectPdfWithMuPdf(input, options = {}) {
  const { document, bytes, limits } = await openPdfWithMuPdf(input, options);
  try {
    const annotationBudget = { maxAnnotations: limits.maxAnnotations, used: 0 };
    const linkBudget = { maxLinks: limits.maxLinks, used: 0 };
    const pageRecords = Array.from({ length: document.countPages() }, (_, index) => {
      const page = document.loadPage(index);
      const structured = page.toStructuredText("preserve-whitespace");
      const annotations = page.getAnnotations();
      const widgets = page.getWidgets();
      const links = page.getLinks();
      try {
        const text = structured.asText();
        const mediaBox = rawPageBox(page, "MediaBox");
        const cropBox = rawPageBox(page, "CropBox") || mediaBox;
        consumeAnnotationBudget(annotationBudget, annotations.length + widgets.length, index + 1);
        consumeLinkBudget(linkBudget, links.length, index + 1);
        const nativeWidgets = widgets.map((widget) => nativeWidget(widget, index + 1));
        return [
          {
            kind: "mupdfPage",
            page: index + 1,
            bbox: pdfRectToBbox(page.getBounds("CropBox")),
            mediaBox: mediaBox && pdfRectToBbox(mediaBox),
            cropBox: cropBox && pdfRectToBbox(cropBox),
            rotation: rawPageRotation(page),
            textChars: text.length,
            annotations: annotations.length,
            widgets: widgets.length,
            links: links.length,
          },
          ...annotations.map((annotation) => ({
            kind: "mupdfAnnotation",
            page: index + 1,
            ...nativeAnnotation(annotation, index + 1),
          })),
          ...nativeWidgets.map((widget) => ({
            kind: "mupdfWidget",
            page: index + 1,
            ...widget,
          })),
          ...links.map((link) => ({
            kind: "mupdfLink",
            page: index + 1,
            ...nativeLink(link, index + 1),
          })),
        ];
      } finally {
        annotations.forEach((annotation) => annotation.destroy());
        widgets.forEach((widget) => widget.destroy());
        links.forEach((link) => link.destroy());
        structured.destroy();
        page.destroy();
      }
    });
    const embeddedFiles = document.getEmbeddedFiles();
    const summary = {
      kind: "mupdfDocument",
      provider: "mupdf",
      providerVersion: MUPDF_VERSION,
      bytes: bytes.byteLength,
      sourceSha256: sha256(bytes),
      pages: document.countPages(),
      objects: document.countObjects(),
      repaired: document.wasRepaired(),
      versions: document.countVersions(),
      unsavedVersions: document.countUnsavedVersions(),
      canSaveIncrementally: document.canBeSavedIncrementally(),
      metadata: metadataFor(document),
      embeddedFiles: Object.keys(embeddedFiles).length,
    };
    Object.values(embeddedFiles).forEach((file) => file.destroy());
    const records = pageRecords.flat();
    const formFields = collectNativeFormFields(records.filter((record) => record.kind === "mupdfWidget"));
    return { summary, records: [summary, ...records, ...formFields] };
  } finally {
    document.destroy();
  }
}

export async function renderPdfWithMuPdf(input, options = {}) {
  const { document, limits } = await openPdfWithMuPdf(input, options);
  let page;
  let pixmap;
  try {
    const pageIndex = pageIndexFor(document, options, 0);
    const dpi = Number(options.dpi ?? 144);
    if (!Number.isFinite(dpi) || dpi <= 0 || dpi > 1200) throw new Error("dpi must be greater than 0 and no more than 1200.");
    const format = String(options.format || "png").toLowerCase();
    if (!new Set(["png", "jpeg", "jpg"]).has(format)) throw new Error(`MuPDF render format ${format} is unsupported; use png or jpeg.`);
    const scale = dpi / 72;
    page = document.loadPage(pageIndex);
    const bounds = page.getBounds();
    const targetWidth = Math.max(1, Math.ceil((bounds[2] - bounds[0]) * scale));
    const targetHeight = Math.max(1, Math.ceil((bounds[3] - bounds[1]) * scale));
    const targetPixels = targetWidth * targetHeight;
    if (targetPixels > limits.maxRenderPixels) throw new Error(`MuPDF render exceeds maxRenderPixels (${targetPixels} > ${limits.maxRenderPixels}).`);
    pixmap = page.toPixmap(mupdf.Matrix.scale(scale, scale), mupdf.ColorSpace.DeviceRGB, Boolean(options.alpha), options.showExtras !== false);
    const encoded = format === "png" ? pixmap.asPNG() : pixmap.asJPEG(Math.max(1, Math.min(100, Number(options.quality ?? 90))));
    const bytes = copyBytes(encoded);
    const type = format === "png" ? "image/png" : "image/jpeg";
    return new FileBlob(bytes, { type, metadata: { provider: "mupdf", providerVersion: MUPDF_VERSION, page: pageIndex + 1, dpi, width: pixmap.getWidth(), height: pixmap.getHeight() } });
  } finally {
    pixmap?.destroy();
    page?.destroy();
    document.destroy();
  }
}

function pdfObjectName(object) {
  try { return object.isName() ? object.asName() : undefined; } catch { return undefined; }
}

function isSignatureDictionary(object) {
  if (!object?.isDictionary()) return false;
  const type = object.get("Type");
  const fieldType = object.getInheritable("FT");
  const byteRange = object.get("ByteRange");
  const contents = object.get("Contents");
  try {
    return pdfObjectName(type) === "Sig"
      || pdfObjectName(fieldType) === "Sig"
      || (byteRange.isArray() && byteRange.length >= 2 && !contents.isNull());
  } finally {
    type.destroy();
    fieldType.destroy();
    byteRange.destroy();
    contents.destroy();
  }
}

function documentHasSignature(document) {
  for (let objectNumber = 1; objectNumber < document.countObjects(); objectNumber += 1) {
    let indirect;
    let object;
    try {
      indirect = document.newIndirect(objectNumber);
      object = indirect.resolve();
      if (isSignatureDictionary(object)) return true;
    } catch {
      // Sparse, deleted, or malformed xref entries are not signature evidence.
    } finally {
      object?.destroy();
      indirect?.destroy();
    }
  }
  return false;
}

function requireUnsignedPolicy(document, options) {
  const signed = documentHasSignature(document);
  if (signed && !options.allowSigned) throw new Error("PDF contains a signature; set allowSigned only after reviewing DocMDP/FieldMDP policy.");
  if (signed && options.savePolicy === "incremental") throw new Error("Signed-PDF incremental editing is unsupported because this API does not validate DocMDP/FieldMDP permissions or preserve a signature over the current revision.");
  if (signed && !options.invalidateSignatures) throw new Error("Editing a signed PDF requires invalidateSignatures: true and a rewrite that deliberately invalidates existing signatures.");
  return signed;
}

function applyTextRedaction(document, operation) {
  const targets = operation.page == null && operation.pageIndex == null
    ? Array.from({ length: document.countPages() }, (_, index) => index)
    : [pageIndexFor(document, operation)];
  const term = String(operation.term || "");
  if (!term) throw new Error("redact_text requires a non-empty term.");
  let matches = 0;
  const pages = [];
  for (const index of targets) {
    const page = document.loadPage(index);
    try {
      const hits = page.search(term);
      for (const hit of hits) {
        const annotation = page.createAnnotation("Redact");
        try {
          annotation.setRect(hitRect(hit));
          annotation.setColor(operation.fill || [0, 0, 0]);
          annotation.update();
        } finally {
          annotation.destroy();
        }
      }
      if (hits.length) {
        page.applyRedactions(Boolean(operation.blackBoxes ?? true), mupdf.PDFPage.REDACT_IMAGE_PIXELS, mupdf.PDFPage.REDACT_LINE_ART_REMOVE_IF_TOUCHED, mupdf.PDFPage.REDACT_TEXT_REMOVE);
        matches += hits.length;
        pages.push({ page: index + 1, matches: hits.length });
      }
    } finally {
      page.destroy();
    }
  }
  if (!matches) throw new Error(`redact_text term was not found: ${JSON.stringify(term)}.`);
  return { type: "redact_text", matches, pages };
}

function applyRectRedaction(document, operation) {
  const { index, page } = pageFor(document, operation);
  let annotation;
  try {
    annotation = page.createAnnotation("Redact");
    annotation.setRect(bboxToPdfRect(operation.bbox || operation.rect, "redaction rectangle"));
    annotation.setColor(operation.fill || [0, 0, 0]);
    annotation.update();
    page.applyRedactions(Boolean(operation.blackBoxes ?? true), mupdf.PDFPage.REDACT_IMAGE_PIXELS, mupdf.PDFPage.REDACT_LINE_ART_REMOVE_IF_TOUCHED, mupdf.PDFPage.REDACT_TEXT_REMOVE);
    return { type: "redact_rect", page: index + 1, bbox: operation.bbox || operation.rect };
  } finally {
    annotation?.destroy();
    page.destroy();
  }
}

function applyPageCrop(document, operation) {
  const { index, page } = pageFor(document, operation);
  try {
    const cropRect = bboxToPdfRect(operation.bbox || operation.rect, "crop rectangle");
    const mediaRect = rawPageBox(page, "MediaBox");
    if (!mediaRect) throw new Error("set_page_crop requires a finite four-number MediaBox.");
    const rotation = rawPageRotation(page);
    if (rotation !== 0) throw new Error("set_page_crop supports only unrotated pages; use an explicit specialist provider for rotated page boxes.");
    if (!pdfRectContains(mediaRect, cropRect)) {
      throw new Error("set_page_crop crop rectangle must fit fully inside the page MediaBox.");
    }
    const currentCropRect = rawPageBox(page, "CropBox") || mediaRect;
    const pageCoordinates = [
      cropRect[0] - currentCropRect[0],
      cropRect[1] - currentCropRect[1],
      cropRect[2] - currentCropRect[0],
      cropRect[3] - currentCropRect[1],
    ];
    page.setPageBox("CropBox", pageCoordinates);
    page.update();
    const writtenCropRect = rawPageBox(page, "CropBox");
    if (!writtenCropRect || !pdfRectsEqual(writtenCropRect, cropRect)) {
      throw new Error("MuPDF did not preserve the requested raw CropBox; refusing to save an ambiguous crop.");
    }
    return {
      type: "set_page_crop",
      page: index + 1,
      bbox: pdfRectToBbox(cropRect),
      mediaBox: pdfRectToBbox(mediaRect),
      contentRemoved: false,
    };
  } finally {
    page.destroy();
  }
}

function applyPageRotation(document, operation) {
  const { index, page } = pageFor(document, operation);
  let pageObject;
  try {
    const rotation = operation.rotation;
    if (!Number.isInteger(rotation) || !PAGE_ROTATIONS.has(rotation)) {
      throw new Error("rotate_page rotation must be 0, 90, 180, or 270.");
    }
    const previousRotation = rawPageRotation(page);
    if (previousRotation === undefined || !PAGE_ROTATIONS.has(previousRotation)) {
      throw new Error("rotate_page requires a finite right-angle inherited Rotate value.");
    }
    pageObject = page.getObject();
    pageObject.put("Rotate", rotation);
    page.update();
    const writtenRotation = rawPageRotation(page);
    if (writtenRotation !== rotation) {
      throw new Error("MuPDF did not preserve the requested page rotation; refusing to save an ambiguous rotation.");
    }
    return {
      type: "rotate_page",
      page: index + 1,
      rotation,
      previousRotation,
      contentRemoved: false,
    };
  } finally {
    pageObject?.destroy();
    page.destroy();
  }
}

function applyTextAnnotationAddition(document, operation, context = {}) {
  const { index, page } = pageFor(document, operation);
  let annotations = [];
  let retained = [];
  let created;
  try {
    if (!/^[a-f0-9]{64}$/u.test(String(operation.sourceSha256 || "")) || operation.sourceSha256 !== context.sourceSha256) {
      throw new Error("add_text_annotation sourceSha256 must exactly match PdfFile.inspectPdf(...).summary.sourceSha256 for the current input bytes.");
    }
    const expectedPage = pageExpectation(operation.expectedPage, "add_text_annotation");
    const request = textAnnotationRequest(operation);
    const actualPage = nativePageSnapshot(page, "add_text_annotation");
    const pageMismatch = pageExpectationMismatch(actualPage, expectedPage);
    if (pageMismatch) {
      throw new Error(`add_text_annotation precondition page ${pageMismatch} did not match page ${index + 1}; refusing stale coordinate evidence.`);
    }
    if (actualPage.rotation !== 0) {
      throw new Error(`add_text_annotation supports only unrotated pages; page ${index + 1} has /Rotate ${actualPage.rotation}. Use a specialist provider for rotated annotation placement.`);
    }
    const visibleRect = bboxToPdfRect(actualPage.bbox, "add_text_annotation inspected CropBox");
    const anchorRect = [
      request.point[0],
      request.point[1],
      request.point[0] + TEXT_ANNOTATION_ANCHOR_SIZE,
      request.point[1] + TEXT_ANNOTATION_ANCHOR_SIZE,
    ];
    if (!pdfRectContains(visibleRect, anchorRect)) {
      throw new Error(`add_text_annotation point plus its native Text-note icon footprint must fit fully inside the inspected visible CropBox on page ${index + 1}.`);
    }
    annotations = page.getAnnotations();
    // MuPDF's returned annotation list is a live view: creating a note can
    // mutate its length in place. Snapshot the count before the mutation.
    const beforeCount = annotations.length;
    created = page.createAnnotation("Text");
    created.setRect(anchorRect);
    created.setContents(request.contents);
    if (request.author !== undefined) created.setAuthor(request.author);
    if (request.subject !== undefined) created.setSubject(request.subject);
    created.update();
    page.update();
    const immediate = nativeAnnotation(created, index + 1);
    if (!immediate.id || !immediate.xref) {
      throw new Error("MuPDF did not create a uniquely addressable Text annotation; refusing to save an ambiguous addition.");
    }
    let mismatch = annotationAddedAtPointMismatch(immediate, request);
    if (mismatch) {
      throw new Error(`MuPDF did not preserve add_text_annotation ${mismatch} before save; refusing an ambiguous addition.`);
    }
    if (!immediate.rect || !pdfRectContains(visibleRect, bboxToPdfRect(immediate.rect))) {
      throw new Error("MuPDF placed the Text annotation outside the inspected visible CropBox; refusing an ambiguous addition.");
    }
    retained = page.getAnnotations();
    const added = retained.filter((annotation) => annotationXref(annotation) === immediate.xref);
    if (retained.length !== beforeCount + 1 || added.length !== 1) {
      throw new Error(`MuPDF did not retain exactly one uniquely addressable Text annotation on page ${index + 1}; refusing an ambiguous addition.`);
    }
    const addedRecord = nativeAnnotation(added[0], index + 1);
    mismatch = annotationAddedAtPointMismatch(addedRecord, request);
    if (mismatch) {
      throw new Error(`MuPDF did not retain add_text_annotation ${mismatch} on page ${index + 1}; refusing an ambiguous addition.`);
    }
    if (!addedRecord.rect || !pdfRectContains(visibleRect, bboxToPdfRect(addedRecord.rect))) {
      throw new Error("MuPDF retained the Text annotation outside the inspected visible CropBox; refusing an ambiguous addition.");
    }
    return {
      type: "add_text_annotation",
      page: index + 1,
      expectedPage,
      point: request.point,
      added: addedRecord,
      beforeCount,
      afterCount: retained.length,
    };
  } finally {
    for (const annotation of new Set([...annotations, ...retained, created].filter(Boolean))) annotation.destroy();
    page.destroy();
  }
}

function applyAnnotationDeletion(document, operation, context = {}) {
  const { index, page } = pageFor(document, operation);
  let annotations = [];
  let retained = [];
  let target;
  try {
    if (!/^[a-f0-9]{64}$/u.test(String(operation.sourceSha256 || "")) || operation.sourceSha256 !== context.sourceSha256) {
      throw new Error("delete_annotation sourceSha256 must exactly match PdfFile.inspectPdf(...).summary.sourceSha256 for the current input bytes.");
    }
    const locator = parseAnnotationId(operation.annotationId, "delete_annotation");
    if (locator.page !== index + 1) {
      throw new Error(`delete_annotation annotationId page ${locator.page} does not match operation page ${index + 1}.`);
    }
    const expected = annotationExpectation(operation.expected, "delete_annotation");
    annotations = page.getAnnotations();
    target = annotations.find((annotation) => annotationXref(annotation) === locator.xref);
    if (!target) {
      throw new Error(`delete_annotation could not find source-bound annotation ${operation.annotationId} on page ${index + 1}. Re-inspect the current source PDF before retrying.`);
    }
    const matched = nativeAnnotation(target, index + 1);
    if (matched.id !== operation.annotationId) {
      throw new Error(`delete_annotation locator ${operation.annotationId} did not resolve to the expected native annotation.`);
    }
    const mismatch = annotationExpectationMismatch(matched, expected);
    if (mismatch) {
      throw new Error(`delete_annotation precondition ${mismatch} did not match ${operation.annotationId}; refusing a stale or ambiguous mutation.`);
    }
    const beforeCount = annotations.length;
    page.deleteAnnotation(target);
    page.update();
    retained = page.getAnnotations();
    if (retained.some((annotation) => annotationXref(annotation) === locator.xref)) {
      throw new Error(`MuPDF did not remove annotation ${operation.annotationId}; refusing to save an ambiguous deletion.`);
    }
    return {
      type: "delete_annotation",
      page: index + 1,
      annotationId: matched.id,
      xref: locator.xref,
      matched,
      beforeCount,
      afterCount: retained.length,
    };
  } finally {
    for (const annotation of new Set([...annotations, ...retained, target].filter(Boolean))) annotation.destroy();
    page.destroy();
  }
}

function applyAnnotationUpdate(document, operation, context = {}) {
  const { index, page } = pageFor(document, operation);
  let annotations = [];
  let retained = [];
  let target;
  try {
    if (!/^[a-f0-9]{64}$/u.test(String(operation.sourceSha256 || "")) || operation.sourceSha256 !== context.sourceSha256) {
      throw new Error("update_annotation sourceSha256 must exactly match PdfFile.inspectPdf(...).summary.sourceSha256 for the current input bytes.");
    }
    const locator = parseAnnotationId(operation.annotationId, "update_annotation");
    if (locator.page !== index + 1) {
      throw new Error(`update_annotation annotationId page ${locator.page} does not match operation page ${index + 1}.`);
    }
    const expected = annotationExpectation(operation.expected, "update_annotation");
    const patch = annotationPatch(operation.patch);
    annotations = page.getAnnotations();
    target = annotations.find((annotation) => annotationXref(annotation) === locator.xref);
    if (!target) {
      throw new Error(`update_annotation could not find source-bound annotation ${operation.annotationId} on page ${index + 1}. Re-inspect the current source PDF before retrying.`);
    }
    const matched = nativeAnnotation(target, index + 1);
    if (matched.id !== operation.annotationId) {
      throw new Error(`update_annotation locator ${operation.annotationId} did not resolve to the expected native annotation.`);
    }
    const mismatch = annotationExpectationMismatch(matched, expected);
    if (mismatch) {
      throw new Error(`update_annotation precondition ${mismatch} did not match ${operation.annotationId}; refusing a stale or ambiguous mutation.`);
    }
    if (matched.type !== "Text") {
      throw new Error(`update_annotation supports only native Text annotations; ${operation.annotationId} resolves to ${matched.type}.`);
    }
    if (patch.contents !== undefined) target.setContents(patch.contents);
    if (patch.author !== undefined) target.setAuthor(patch.author);
    if (patch.subject !== undefined) target.setSubject(patch.subject);
    target.update();
    page.update();
    retained = page.getAnnotations();
    const updatedTargets = retained.filter((annotation) => annotationXref(annotation) === locator.xref);
    if (updatedTargets.length !== 1) {
      throw new Error(`MuPDF did not retain one uniquely addressable annotation ${operation.annotationId}; refusing to save an ambiguous update.`);
    }
    const updated = nativeAnnotation(updatedTargets[0], index + 1);
    const patchMismatch = annotationExpectationMismatch(updated, patch);
    if (patchMismatch) {
      throw new Error(`MuPDF did not preserve update_annotation patch ${patchMismatch} for ${operation.annotationId}; refusing to save an ambiguous update.`);
    }
    return {
      type: "update_annotation",
      page: index + 1,
      annotationId: matched.id,
      xref: locator.xref,
      matched,
      patch,
      updated,
    };
  } finally {
    for (const annotation of new Set([...annotations, ...retained, target].filter(Boolean))) annotation.destroy();
    page.destroy();
  }
}

function withSourceBoundLink(document, operation, context = {}, operationName, mutate) {
  const { index, page } = pageFor(document, operation);
  let links = [];
  let retained = [];
  let target;
  try {
    if (!/^[a-f0-9]{64}$/u.test(String(operation.sourceSha256 || "")) || operation.sourceSha256 !== context.sourceSha256) {
      throw new Error(`${operationName} sourceSha256 must exactly match PdfFile.inspectPdf(...).summary.sourceSha256 for the current input bytes.`);
    }
    const locator = parseLinkId(operation.linkId, operationName);
    if (locator.page !== index + 1) {
      throw new Error(`${operationName} linkId page ${locator.page} does not match operation page ${index + 1}.`);
    }
    const expected = linkExpectation(operation.expected, operationName);
    links = page.getLinks();
    const matches = links.filter((link) => nativeLink(link, index + 1).id === operation.linkId);
    if (matches.length !== 1) {
      throw new Error(`${operationName} could not uniquely find source-bound link ${operation.linkId} on page ${index + 1}. Re-inspect the current source PDF before retrying.`);
    }
    target = matches[0];
    const matched = nativeLink(target, index + 1);
    if (matched.id !== operation.linkId || !matched.id.endsWith(locator.fingerprint)) {
      throw new Error(`${operationName} locator ${operation.linkId} did not resolve to the expected native link.`);
    }
    const mismatch = linkExpectationMismatch(matched, expected);
    if (mismatch) {
      throw new Error(`${operationName} precondition ${mismatch} did not match ${operation.linkId}; refusing a stale or ambiguous mutation.`);
    }
    return mutate({
      index,
      page,
      links,
      target,
      matched,
      readRetained() {
        if (retained.length) throw new Error(`${operationName} read native links more than once while verifying one mutation.`);
        retained = page.getLinks();
        return retained;
      },
    });
  } finally {
    for (const link of new Set([...links, ...retained, target].filter(Boolean))) link.destroy();
    page.destroy();
  }
}

function applyLinkDeletion(document, operation, context = {}) {
  return withSourceBoundLink(document, operation, context, "delete_link", ({
    index, page, links, target, matched, readRetained,
  }) => {
    const beforeCount = links.length;
    page.deleteLink(target);
    const retained = readRetained();
    if (retained.some((link) => nativeLink(link, index + 1).id === operation.linkId)) {
      throw new Error(`MuPDF did not remove link ${operation.linkId}; refusing to save an ambiguous deletion.`);
    }
    return {
      type: "delete_link",
      page: index + 1,
      linkId: matched.id,
      matched,
      beforeCount,
      afterCount: retained.length,
    };
  });
}

function applyLinkUpdate(document, operation, context = {}) {
  const patch = linkPatch(operation.patch);
  return withSourceBoundLink(document, operation, context, "update_link", ({
    index, target, matched, readRetained,
  }) => {
    target.setURI(patch.url);
    const retained = readRetained();
    const updatedMatches = retained
      .map((link) => nativeLink(link, index + 1))
      .filter((link) => link.url === patch.url
        && pdfRectsEqual(bboxToPdfRect(link.bbox), bboxToPdfRect(matched.bbox)));
    if (updatedMatches.length !== 1) {
      throw new Error(`MuPDF did not retain one uniquely addressable link after update_link ${operation.linkId}; refusing to save an ambiguous update.`);
    }
    return {
      type: "update_link",
      page: index + 1,
      linkId: matched.id,
      matched,
      patch,
      updated: updatedMatches[0],
    };
  });
}

function applyLinkAddition(document, operation, context = {}) {
  const { index, page } = pageFor(document, operation);
  let links = [];
  let retained = [];
  let created;
  try {
    if (!/^[a-f0-9]{64}$/u.test(String(operation.sourceSha256 || "")) || operation.sourceSha256 !== context.sourceSha256) {
      throw new Error("add_link sourceSha256 must exactly match PdfFile.inspectPdf(...).summary.sourceSha256 for the current input bytes.");
    }
    const expectedPage = pageExpectation(operation.expectedPage, "add_link");
    const requestedRect = bboxToPdfRect(operation.bbox, "add_link bbox");
    const url = nativeLinkUri(operation.url, "add_link url");
    const actualPage = nativePageSnapshot(page, "add_link");
    const pageMismatch = pageExpectationMismatch(actualPage, expectedPage);
    if (pageMismatch) {
      throw new Error(`add_link precondition page ${pageMismatch} did not match page ${index + 1}; refusing stale coordinate evidence.`);
    }
    if (actualPage.rotation !== 0) {
      throw new Error(`add_link supports only unrotated pages; page ${index + 1} has /Rotate ${actualPage.rotation}. Use a specialist provider for rotated link placement.`);
    }
    if (!pdfRectContains(bboxToPdfRect(actualPage.bbox), requestedRect)) {
      throw new Error(`add_link bbox must be fully inside the inspected visible CropBox on page ${index + 1}.`);
    }
    links = page.getLinks();
    const beforeMatches = links
      .map((link) => nativeLink(link, index + 1))
      .filter((link) => link.url === url && pdfRectsEqual(bboxToPdfRect(link.bbox), requestedRect));
    if (beforeMatches.length) {
      throw new Error(`add_link would create a duplicate source-visible link on page ${index + 1}; re-inspect and choose a distinct URL or rectangle.`);
    }
    created = page.createLink(requestedRect, url);
    const immediate = nativeLink(created, index + 1);
    if (immediate.url !== url || !pdfRectsEqual(bboxToPdfRect(immediate.bbox), requestedRect)) {
      throw new Error(`MuPDF did not create the requested link URL and rectangle on page ${index + 1}; refusing to save an ambiguous addition.`);
    }
    retained = page.getLinks();
    const addedMatches = retained
      .map((link) => nativeLink(link, index + 1))
      .filter((link) => link.url === url && pdfRectsEqual(bboxToPdfRect(link.bbox), requestedRect));
    if (addedMatches.length !== 1 || retained.length !== links.length + 1) {
      throw new Error(`MuPDF did not retain exactly one newly addressable link on page ${index + 1}; refusing to save an ambiguous addition.`);
    }
    return {
      type: "add_link",
      page: index + 1,
      expectedPage,
      added: {
        url: addedMatches[0].url,
        bbox: addedMatches[0].bbox,
        external: addedMatches[0].external,
      },
      beforeCount: links.length,
      afterCount: retained.length,
    };
  } finally {
    for (const link of new Set([...links, ...retained, created].filter(Boolean))) link.destroy();
    page.destroy();
  }
}

function normalizeCheckboxValue(value, field, operationName = "fill_form") {
  if (typeof value === "boolean") return value;
  if (value === 0 || value === 1) return value === 1;
  const normalized = String(value ?? "").trim().toLowerCase();
  if (["true", "yes", "on", "1"].includes(normalized)) return true;
  if (["false", "no", "off", "0", ""].includes(normalized)) return false;
  throw new Error(`${operationName} checkbox field ${field} requires a boolean, on/off, yes/no, or 1/0 value.`);
}

function applyOperation(document, operation, context) {
  switch (operation.type) {
    case "add_text_annotation": return applyTextAnnotationAddition(document, operation, context);
    case "fill_form": {
      const field = String(operation.field || operation.name || "");
      if (!field) throw new Error("fill_form requires a field name.");
      const indexes = operation.page == null && operation.pageIndex == null ? Array.from({ length: document.countPages() }, (_, index) => index) : [pageIndexFor(document, operation)];
      const matches = [];
      for (const index of indexes) {
        const page = document.loadPage(index);
        const widgets = page.getWidgets();
        try {
          for (const widget of widgets) {
            if (widget.getName() !== field) continue;
            let accepted;
            if (widget.isText()) accepted = widget.setTextValue(String(operation.value ?? ""));
            else if (widget.isChoice()) accepted = widget.setChoiceValue(String(operation.value ?? ""));
            else if (widget.isCheckbox()) {
              const desired = normalizeCheckboxValue(operation.value, field);
              const current = !/^\s*(?:|off)\s*$/i.test(String(widget.getValue() ?? ""));
              accepted = desired === current ? 1 : widget.toggle();
            } else if (widget.isRadioButton()) {
              throw new Error(`fill_form does not yet expose a trustworthy radio export-value mapping for field ${field}; use the typed pypdf form workflow.`);
            }
            else throw new Error(`fill_form does not support widget type ${widget.getFieldType()} for field ${field}.`);
            if (!accepted) throw new Error(`MuPDF rejected value for form field ${field}.`);
            widget.update();
            matches.push({ page: index + 1, type: widget.getFieldType() });
          }
          page.update();
        } finally {
          widgets.forEach((widget) => widget.destroy());
          page.destroy();
        }
      }
      if (!matches.length) throw new Error(`Form field not found: ${field}.`);
      return { type: operation.type, field, widgets: matches };
    }
    case "update_form_field": return applySourceBoundFormFieldUpdate(document, operation, context);
    case "delete_page": {
      if (document.countPages() <= 1) throw new Error("delete_page cannot remove the final page.");
      const index = pageIndexFor(document, operation);
      document.deletePage(index);
      return { type: operation.type, page: index + 1 };
    }
    case "delete_annotation": return applyAnnotationDeletion(document, operation, context);
    case "update_annotation": return applyAnnotationUpdate(document, operation, context);
    case "set_page_crop": return applyPageCrop(document, operation);
    case "rotate_page": return applyPageRotation(document, operation);
    case "rearrange_pages": {
      const pages = operation.pages;
      if (!Array.isArray(pages) || pages.length !== document.countPages()) throw new Error("rearrange_pages requires a complete 1-based page permutation.");
      const indexes = pages.map((page) => Number(page) - 1);
      if (indexes.some((index) => !Number.isInteger(index) || index < 0 || index >= document.countPages()) || new Set(indexes).size !== indexes.length) throw new Error("rearrange_pages requires each current page exactly once.");
      document.rearrangePages(indexes);
      return { type: operation.type, pages };
    }
    case "set_metadata": {
      const values = operation.values || operation.metadata;
      if (!values || typeof values !== "object" || Array.isArray(values)) throw new Error("set_metadata requires a values object.");
      const keys = [];
      for (const [name, value] of Object.entries(values)) {
        const key = METADATA_KEYS[name];
        if (!key) throw new Error(`Unsupported metadata key: ${name}.`);
        document.setMetaData(key, String(value ?? ""));
        keys.push(name);
      }
      return { type: operation.type, keys };
    }
    case "delete_embedded_file": {
      const name = String(operation.name || operation.filename || "");
      const files = document.getEmbeddedFiles();
      try {
        if (!name || !files[name]) throw new Error(`Embedded file not found: ${name}.`);
        document.deleteEmbeddedFile(name);
        return { type: operation.type, name };
      } finally {
        Object.values(files).forEach((file) => file.destroy());
      }
    }
    case "add_link": return applyLinkAddition(document, operation, context);
    case "delete_link": return applyLinkDeletion(document, operation, context);
    case "update_link": return applyLinkUpdate(document, operation, context);
    case "redact_text": return applyTextRedaction(document, operation);
    case "redact_rect": return applyRectRedaction(document, operation);
    default: throw new Error(`Unsupported MuPDF edit operation: ${operation.type}.`);
  }
}

export async function editPdfWithMuPdf(input, options = {}) {
  const operations = options.operations;
  if (!Array.isArray(operations) || !operations.length) throw new Error("MuPDF editing requires a non-empty operations array.");
  const savePolicy = String(options.savePolicy || options.strategy || "rewrite").toLowerCase();
  if (!new Set(["rewrite", "incremental"]).has(savePolicy)) {
    throw new Error(`MuPDF savePolicy ${savePolicy} is unsupported; strict sanitize remains a separate audited workflow.`);
  }
  const destructiveIncremental = savePolicy === "incremental"
    ? operations.find((operation) => INCREMENTAL_DESTRUCTIVE_OPERATIONS.has(operation?.type))
    : undefined;
  if (destructiveIncremental) {
    const label = destructiveIncremental.type.startsWith("redact_")
      ? "redaction"
      : ["add_link", "add_text_annotation"].includes(destructiveIncremental.type)
        ? `source-bound operation ${destructiveIncremental.type}`
        : `destructive operation ${destructiveIncremental.type}`;
    throw new Error(`MuPDF ${label} cannot save incrementally because prior revisions retain the original content; use rewrite. Rewrite is still not a complete sanitize workflow.`);
  }
  const { document, bytes, limits } = await openPdfWithMuPdf(input, options);
  let saved;
  try {
    const signed = requireUnsignedPolicy(document, { ...options, savePolicy });
    const sourceSha256 = sha256(bytes);
    const applied = operations.map((operation) => applyOperation(document, operation || {}, { sourceSha256, limits }));
    if (savePolicy === "incremental" && !document.canBeSavedIncrementally()) throw new Error("MuPDF cannot save these changes incrementally; refusing a rewrite fallback.");
    saved = savePolicy === "incremental"
      ? document.saveToBuffer({ incremental: true })
      : document.saveToBuffer("garbage=2,compress=yes");
    const output = copyBytes(saved);
    if (savePolicy === "incremental") {
      const prefix = output.subarray(0, bytes.byteLength);
      if (prefix.byteLength !== bytes.byteLength || !prefix.every((value, index) => value === bytes[index])) throw new Error("MuPDF incremental output did not preserve the complete input byte prefix.");
    }
    return new FileBlob(output, {
      type: PDF_MIME,
      metadata: {
        provider: "mupdf",
        providerVersion: MUPDF_VERSION,
        savePolicy,
        signedInput: signed,
        signatureValidity: signed ? "unknown" : "not-applicable",
        signaturesInvalidated: signed && savePolicy === "rewrite",
        originalBytes: bytes.byteLength,
        outputBytes: output.byteLength,
        sourceSha256,
        outputSha256: sha256(output),
        operations: applied,
      },
    });
  } finally {
    saved?.destroy();
    document.destroy();
  }
}
