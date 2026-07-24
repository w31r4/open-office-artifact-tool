import { toUint8Array } from "../shared/binary.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { aid } from "../shared/ids.mjs";
import { attrEscape, xmlEscape } from "../shared/xml.mjs";

const MAX_EMBEDDED_WORKBOOK_BYTES = 16 * 1024 * 1024;
const MAX_EMBEDDED_OFFICE_PACKAGE_BYTES = 16 * 1024 * 1024;
const MAX_DIAGRAM_NODE_TEXT_LENGTH = 32_767;
const DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function normalizeOleOfficePackage(config) {
  if (!config) return undefined;
  const partPath = String(config.partPath || "");
  const contentType = String(config.contentType || "").toLowerCase();
  const sourceSha256 = String(config.sourceSha256 || "").toLowerCase();
  const relationshipId = String(config.relationshipId || "");
  const kind = String(config.kind || "").toLowerCase();
  if (!partPath || contentType !== DOCX_CONTENT_TYPE || !/^[0-9a-f]{64}$/i.test(sourceSha256) || !relationshipId || kind !== "docx") {
    throw new TypeError("Embedded Office package binding is incomplete or outside the bounded DOCX profile.");
  }
  return Object.freeze({ partPath, contentType, sourceSha256, relationshipId, kind });
}

function hasOnlyValidUnicodeScalars(value) {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
      index += 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      return false;
    }
  }
  return true;
}

function validDiagramNodeText(value) {
  return value.length <= MAX_DIAGRAM_NODE_TEXT_LENGTH &&
    !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(value) &&
    hasOnlyValidUnicodeScalars(value);
}

function validDiagramModelId(value) {
  if (value.length > 1_024 || /[\u0000-\u001f]/u.test(value) || !hasOnlyValidUnicodeScalars(value)) return false;
  if (/^[+-]?\d+$/u.test(value)) {
    const numeric = BigInt(value);
    return numeric >= -2_147_483_648n && numeric <= 2_147_483_647n;
  }
  return /^\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}$/iu.test(value);
}

function normalizeDiagramText(config) {
  if (!config) return undefined;
  const partPath = String(config.partPath || "");
  const contentType = String(config.contentType || "");
  const sourceSha256 = String(config.sourceSha256 || "").toLowerCase();
  const relationshipId = String(config.relationshipId || "");
  const nodes = config.nodes;
  if (!partPath || !contentType || !/^[0-9a-f]{64}$/i.test(sourceSha256) || !relationshipId || !Array.isArray(nodes) || !nodes.length) {
    throw new TypeError("SmartArt diagram text binding is incomplete.");
  }
  const seen = new Set();
  const normalizedNodes = nodes.map((node) => {
    const id = String(node?.id ?? node?.modelId ?? "");
    const text = String(node?.text ?? "");
    if (!id || !validDiagramModelId(id) || !validDiagramNodeText(text) || seen.has(id)) {
      throw new TypeError("SmartArt diagram text binding contains an invalid node.");
    }
    seen.add(id);
    return Object.freeze({ id, text });
  });
  return Object.freeze({
    partPath,
    contentType,
    sourceSha256,
    relationshipId,
    nodes: Object.freeze(normalizedNodes),
  });
}

function diagramTextRecord(binding, nodes) {
  if (!binding) return undefined;
  return Object.freeze({
    partPath: binding.partPath,
    contentType: binding.contentType,
    sourceSha256: binding.sourceSha256,
    relationshipId: binding.relationshipId,
    nodes: Object.freeze(nodes.map((node) => Object.freeze({ id: node.id, text: node.text }))),
  });
}

export function createNativePresentationObjectClass({ normalizeFrame }) {
  return class NativePresentationObject {
    constructor(slide, config = {}) {
      this.slide = slide;
      this.kind = "nativeObject";
      this.id = config.id || aid("no");
      this.nativeId = config.nativeId;
      this.creationId = config.creationId;
      this.name = config.name || "";
      this.nativeKind = config.nativeKind || "graphicFrame";
      this.position = normalizeFrame(config, { left: 0, top: 0, width: 1, height: 1 });
      this.rawXml = String(config.rawXml || "");
      this.sourcePart = config.sourcePart;
      Object.defineProperty(this, "editable", { enumerable: true, value: false, writable: false });
      this.relationshipReferences = (config.relationshipReferences || []).map((reference) => ({ ...reference }));
      this.rootRelationships = (config.rootRelationships || []).map((relationship) => ({ ...relationship }));
      this.parts = (config.parts || []).map((part) => ({ ...part, bytes: new Uint8Array(part.bytes), relationships: (part.relationships || []).map((relationship) => ({ ...relationship })) }));
      const oleWorkbook = config.oleWorkbook ? Object.freeze({
        partPath: String(config.oleWorkbook.partPath || ""),
        contentType: String(config.oleWorkbook.contentType || ""),
        sourceSha256: String(config.oleWorkbook.sourceSha256 || "").toLowerCase(),
        relationshipId: String(config.oleWorkbook.relationshipId || ""),
      }) : undefined;
      Object.defineProperty(this, "oleWorkbook", {
        configurable: false,
        enumerable: true,
        writable: false,
        value: oleWorkbook,
      });
      const oleOfficePackage = normalizeOleOfficePackage(config.oleOfficePackage);
      Object.defineProperty(this, "oleOfficePackage", {
        configurable: false,
        enumerable: true,
        writable: false,
        value: oleOfficePackage,
      });
      const diagramText = normalizeDiagramText(config.diagramText);
      Object.defineProperty(this, "_diagramTextBinding", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: diagramText,
      });
      Object.defineProperty(this, "_diagramTextNodes", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: diagramText ? diagramText.nodes.map((node) => ({ ...node })) : undefined,
      });
      Object.defineProperty(this, "diagramText", {
        configurable: false,
        enumerable: true,
        get: () => diagramTextRecord(this._diagramTextBinding, this._diagramTextNodes || []),
      });
      Object.defineProperty(this, "_embeddedWorkbookReplacement", {
        configurable: false,
        enumerable: false,
        writable: true,
        value: undefined,
      });
      Object.defineProperty(this, "_embeddedOfficePackageReplacement", {
        configurable: false,
        enumerable: false,
        writable: true,
        value: undefined,
      });
    }

    setName(value) {
      if (!this.editable) throw new Error(`Native ${this.nativeKind} object ${this.id} is read-only.`);
      const name = String(value ?? "");
      if (name.length > 1_024) throw new RangeError("Native presentation object names cannot exceed 1024 characters.");
      this.name = name;
      return this;
    }

    setPosition(value = {}) {
      if (!this.editable) throw new Error(`Native ${this.nativeKind} object ${this.id} is read-only.`);
      this.position = normalizeFrame({ position: { ...this.position, ...value } }, this.position);
      return this;
    }

    embeddedWorkbookPart() {
      if (!this.oleWorkbook) throw new Error(`Native ${this.nativeKind} object ${this.id} has no embedded XLSX workbook.`);
      const matches = this.parts.filter((part) => part.path === this.oleWorkbook.partPath && part.contentType === this.oleWorkbook.contentType);
      if (matches.length !== 1) throw new Error(`Native ${this.nativeKind} object ${this.id} no longer resolves to one embedded XLSX workbook part.`);
      return matches[0];
    }

    getEmbeddedWorkbook() {
      const part = this.embeddedWorkbookPart();
      const replacement = this._embeddedWorkbookReplacement;
      return new FileBlob(Uint8Array.from(replacement || part.bytes), {
        type: this.oleWorkbook.contentType,
        metadata: replacement
          ? { artifactKind: "workbook", source: "presentationOleObject", partPath: this.oleWorkbook.partPath, boundSourceSha256: this.oleWorkbook.sourceSha256, pendingReplacement: true }
          : { artifactKind: "workbook", source: "presentationOleObject", partPath: this.oleWorkbook.partPath, sourceSha256: this.oleWorkbook.sourceSha256 },
      });
    }

    replaceEmbeddedWorkbook(input) {
      this.embeddedWorkbookPart();
      if (input == null || typeof input === "string" || !(input instanceof FileBlob || input instanceof ArrayBuffer || input instanceof Uint8Array || ArrayBuffer.isView(input))) {
        throw new TypeError("Embedded workbook replacement must be a FileBlob, Uint8Array, ArrayBuffer, or ArrayBuffer view.");
      }
      const bytes = input instanceof FileBlob ? input.bytes : toUint8Array(input);
      if (!bytes.byteLength || bytes.byteLength > MAX_EMBEDDED_WORKBOOK_BYTES) {
        throw new RangeError(`Embedded workbook replacement must contain 1 through ${MAX_EMBEDDED_WORKBOOK_BYTES} bytes.`);
      }
      this._embeddedWorkbookReplacement = Uint8Array.from(bytes);
      return this;
    }

    _embeddedWorkbookReplacementBytes() {
      return this._embeddedWorkbookReplacement ? Uint8Array.from(this._embeddedWorkbookReplacement) : undefined;
    }

    embeddedOfficePackagePart() {
      if (!this.oleOfficePackage) throw new Error(`Native ${this.nativeKind} object ${this.id} has no bounded embedded Office package.`);
      const matches = this.parts.filter((part) => part.path === this.oleOfficePackage.partPath && part.contentType === this.oleOfficePackage.contentType);
      if (matches.length !== 1) throw new Error(`Native ${this.nativeKind} object ${this.id} no longer resolves to one embedded Office package part.`);
      return matches[0];
    }

    getEmbeddedOfficePackage() {
      if (this.oleWorkbook) {
        const workbook = this.getEmbeddedWorkbook();
        return new FileBlob(workbook.bytes, {
          type: workbook.type,
          metadata: { ...workbook.metadata, artifactKind: "officePackage", officePackageKind: "xlsx" },
        });
      }
      const part = this.embeddedOfficePackagePart();
      const replacement = this._embeddedOfficePackageReplacement;
      return new FileBlob(Uint8Array.from(replacement || part.bytes), {
        type: this.oleOfficePackage.contentType,
        metadata: replacement
          ? { artifactKind: "officePackage", officePackageKind: this.oleOfficePackage.kind, source: "presentationOleObject", partPath: this.oleOfficePackage.partPath, boundSourceSha256: this.oleOfficePackage.sourceSha256, pendingReplacement: true }
          : { artifactKind: "officePackage", officePackageKind: this.oleOfficePackage.kind, source: "presentationOleObject", partPath: this.oleOfficePackage.partPath, sourceSha256: this.oleOfficePackage.sourceSha256 },
      });
    }

    replaceEmbeddedOfficePackage(input) {
      if (this.oleWorkbook) return this.replaceEmbeddedWorkbook(input);
      this.embeddedOfficePackagePart();
      if (input == null || typeof input === "string" || !(input instanceof FileBlob || input instanceof ArrayBuffer || input instanceof Uint8Array || ArrayBuffer.isView(input))) {
        throw new TypeError("Embedded Office package replacement must be a FileBlob, Uint8Array, ArrayBuffer, or ArrayBuffer view.");
      }
      if (input instanceof FileBlob && String(input.type || "").toLowerCase() !== this.oleOfficePackage.contentType) {
        throw new TypeError(`Embedded Office package replacement must retain content type ${this.oleOfficePackage.contentType}.`);
      }
      const bytes = input instanceof FileBlob ? input.bytes : toUint8Array(input);
      if (!bytes.byteLength || bytes.byteLength > MAX_EMBEDDED_OFFICE_PACKAGE_BYTES) {
        throw new RangeError(`Embedded Office package replacement must contain 1 through ${MAX_EMBEDDED_OFFICE_PACKAGE_BYTES} bytes.`);
      }
      this._embeddedOfficePackageReplacement = Uint8Array.from(bytes);
      return this;
    }

    _embeddedOfficePackageReplacementBytes() {
      return this._embeddedOfficePackageReplacement ? Uint8Array.from(this._embeddedOfficePackageReplacement) : undefined;
    }

    setDiagramNodeText(nodeId, value) {
      if (!this._diagramTextBinding || !this._diagramTextNodes) {
        throw new Error(`Native ${this.nativeKind} object ${this.id} has no bounded SmartArt diagram-text capability.`);
      }
      const id = String(nodeId ?? "");
      const text = String(value ?? "");
      if (!validDiagramNodeText(text)) {
        throw new RangeError(`SmartArt node text must contain at most ${MAX_DIAGRAM_NODE_TEXT_LENGTH} XML-safe characters.`);
      }
      const node = this._diagramTextNodes.find((candidate) => candidate.id === id);
      if (!node) throw new Error(`SmartArt node ${id || "(empty)"} is not part of the source-bound diagram profile.`);
      node.text = text;
      return this;
    }

    _diagramTextSourceBinding() {
      return this._diagramTextBinding ? diagramTextRecord(this._diagramTextBinding, this._diagramTextBinding.nodes) : undefined;
    }

    _diagramTextReplacement() {
      if (!this._diagramTextBinding || !this._diagramTextNodes) return undefined;
      const changed = this._diagramTextNodes.some((node, index) => node.text !== this._diagramTextBinding.nodes[index].text);
      return changed ? diagramTextRecord(this._diagramTextBinding, this._diagramTextNodes) : undefined;
    }

    inspectRecord() {
      const frame = this.parentGroup ? this.parentGroup.absoluteChildFrame(this) : this.position;
      const editableFields = [
        ...(this.oleWorkbook ? ["embeddedWorkbook"] : []),
        ...(this.oleOfficePackage ? ["embeddedOfficePackage"] : []),
        ...(this._diagramTextBinding ? ["diagramText"] : []),
      ];
      return {
        kind: "nativeObject",
        id: this.id,
        slide: this.slide.index + 1,
        name: this.name || undefined,
        nativeKind: this.nativeKind,
        nativeId: this.nativeId,
        creationId: this.creationId,
        sourcePart: this.sourcePart,
        relationships: this.rootRelationships.length,
        preservedParts: this.parts.length,
        relationshipReferences: this.relationshipReferences.map(({ attribute, id, namespaceUri }) => ({ attribute, id, namespaceUri })),
        nativeRelationships: this.rootRelationships.map(({ id, type, target, targetMode }) => ({ id, type, target, targetMode })),
        nativeParts: this.parts.map((part) => ({ path: part.path, contentType: part.contentType, relationships: part.relationships.length })),
        embeddedWorkbook: this.oleWorkbook ? this._embeddedWorkbookRecord(true) : undefined,
        embeddedOfficePackage: this.oleOfficePackage ? this._embeddedOfficePackageRecord(true) : undefined,
        diagramText: this.diagramText,
        bbox: [frame.left, frame.top, frame.width, frame.height],
        bboxUnit: "px",
        editable: false,
        editableFields,
      };
    }

    _embeddedWorkbookRecord(includeSourceSha256 = false) {
      const replacement = this._embeddedWorkbookReplacement;
      const part = replacement ? undefined : this.embeddedWorkbookPart();
      return {
        partPath: this.oleWorkbook.partPath,
        contentType: this.oleWorkbook.contentType,
        bytes: (replacement || part.bytes).length,
        ...(includeSourceSha256 ? { sourceSha256: this.oleWorkbook.sourceSha256 } : {}),
        replacementPending: Boolean(replacement),
      };
    }

    _embeddedOfficePackageRecord(includeSourceSha256 = false) {
      const replacement = this._embeddedOfficePackageReplacement;
      const part = replacement ? undefined : this.embeddedOfficePackagePart();
      return {
        kind: this.oleOfficePackage.kind,
        partPath: this.oleOfficePackage.partPath,
        contentType: this.oleOfficePackage.contentType,
        bytes: (replacement || part.bytes).length,
        ...(includeSourceSha256 ? { sourceSha256: this.oleOfficePackage.sourceSha256 } : {}),
        replacementPending: Boolean(replacement),
      };
    }

    layoutJson() {
      return {
        kind: "nativeObject",
        id: this.id,
        name: this.name,
        nativeKind: this.nativeKind,
        frame: this.position,
        relationships: this.rootRelationships.length,
        preservedParts: this.parts.length,
        embeddedWorkbook: this.oleWorkbook ? this._embeddedWorkbookRecord() : undefined,
        embeddedOfficePackage: this.oleOfficePackage ? this._embeddedOfficePackageRecord() : undefined,
        diagramText: this.diagramText,
        editable: false,
        editableFields: [
          ...(this.oleWorkbook ? ["embeddedWorkbook"] : []),
          ...(this.oleOfficePackage ? ["embeddedOfficePackage"] : []),
          ...(this._diagramTextBinding ? ["diagramText"] : []),
        ],
      };
    }

    toSvg() {
      const p = this.position;
      if (!(p.width > 1 && p.height > 1)) return `<g data-native-object-id="${attrEscape(this.id)}" data-native-kind="${attrEscape(this.nativeKind)}"/>`;
      const label = this.name || this.nativeKind;
      return `<g data-native-object-id="${attrEscape(this.id)}" data-native-kind="${attrEscape(this.nativeKind)}"><rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#f8fafc" fill-opacity="0.72" stroke="#64748b" stroke-dasharray="6 4"/><text x="${p.left + 8}" y="${p.top + 20}" font-family="Arial" font-size="12" fill="#475569">${xmlEscape(label)}</text></g>`;
    }
  };
}
