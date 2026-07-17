import { toUint8Array } from "../shared/binary.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { aid } from "../shared/ids.mjs";
import { attrEscape, xmlEscape } from "../shared/xml.mjs";

const MAX_EMBEDDED_WORKBOOK_BYTES = 16 * 1024 * 1024;

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
      Object.defineProperty(this, "_embeddedWorkbookReplacement", {
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

    inspectRecord() {
      const frame = this.parentGroup ? this.parentGroup.absoluteChildFrame(this) : this.position;
      const editableFields = this.oleWorkbook ? ["embeddedWorkbook"] : [];
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
        editable: false,
        editableFields: this.oleWorkbook ? ["embeddedWorkbook"] : [],
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
