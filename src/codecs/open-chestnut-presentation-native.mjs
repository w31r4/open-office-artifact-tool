import { createHash } from "node:crypto";
import path from "node:path";
import JSZip from "jszip";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";

function fail(code, message) {
  throw new OpenChestnutCodecError(message, [], { code });
}

function safePartPath(value) {
  const raw = String(value || "");
  const segments = raw.split("/");
  const normalized = path.posix.normalize(raw);
  if (!raw || raw.startsWith("/") || raw.includes("\\") || [...raw].some((character) => character.charCodeAt(0) < 0x20) ||
      segments.some((segment) => !segment || segment === "." || segment === "..") || normalized !== raw) {
    fail("invalid_presentation_native_graph", `OpenChestnut returned an unsafe native-object part path: ${value}`);
  }
  return normalized;
}

function relationshipKey(sourcePath, id) {
  return `${String(sourcePath || "")}\0${String(id || "")}`;
}

function modelRelationship(relationship) {
  return {
    id: relationship.id,
    type: relationship.type,
    target: relationship.target,
    targetMode: relationship.targetMode,
  };
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

// Materialize only the bounded part closure selected by the C# codec. The
// complete source package remains canonical in opaque_opc; byte extraction is
// needed solely so the ordinary JS presentation model can retain the same
// read-only native object until its next canonical OpenChestnut export.
export async function materializePresentationNativeGraphs(envelope) {
  const opaqueOpc = envelope.opaqueOpc;
  const opaqueElements = envelope.payload?.case === "presentation"
    ? envelope.payload.value.slides.flatMap((slide) => slide.elements
      .filter((element) => element.content?.case === "opaque")
      .map((element) => element.content.value))
    : [];
  const requestedPaths = new Set(opaqueElements.flatMap((opaque) => opaque.preservedPartPaths || []).map(safePartPath));
  const partsByPath = new Map();
  for (const part of opaqueOpc?.parts || []) {
    const partPath = safePartPath(part.path);
    if (partsByPath.has(partPath)) fail("invalid_presentation_native_graph", `OpenChestnut returned duplicate opaque part metadata for ${partPath}.`);
    partsByPath.set(partPath, part);
  }
  for (const partPath of requestedPaths) {
    if (!partsByPath.has(partPath)) fail("missing_presentation_native_part", `OpenChestnut native-object graph references missing part metadata ${partPath}.`);
  }

  let zip;
  const sourceBytes = opaqueOpc?.sourcePackage?.data;
  if ([...requestedPaths].some((partPath) => !(partsByPath.get(partPath)?.data?.length))) {
    if (!sourceBytes?.length) fail("missing_source_package", "OpenChestnut native-object graph cannot be materialized because its source package snapshot is missing.");
    try {
      zip = await JSZip.loadAsync(sourceBytes, { createFolders: false });
    } catch (error) {
      fail("invalid_opc_package", `OpenChestnut source package snapshot is not a readable ZIP package: ${error.message}`);
    }
  }

  const materializedParts = new Map();
  await Promise.all([...requestedPaths].map(async (partPath) => {
    const metadata = partsByPath.get(partPath);
    let bytes = metadata.data?.length ? new Uint8Array(metadata.data) : undefined;
    if (!bytes) {
      const entry = zip.file(partPath);
      if (!entry) fail("missing_presentation_native_part", `OpenChestnut source package snapshot is missing native-object part ${partPath}.`);
      bytes = await entry.async("uint8array");
    }
    if (metadata.sha256 && sha256(bytes) !== metadata.sha256.toLowerCase()) {
      fail("presentation_native_part_hash_mismatch", `OpenChestnut native-object part ${partPath} does not match its opaque graph hash.`);
    }
    materializedParts.set(partPath, {
      path: partPath,
      contentType: metadata.contentType || "application/octet-stream",
      bytes,
      sourceSha256: (metadata.sha256 || sha256(bytes)).toLowerCase(),
      relationships: (metadata.relationships || []).map(modelRelationship),
    });
  }));

  const relationships = new Map();
  for (const relationship of opaqueOpc?.packageRelationships || []) {
    const key = relationshipKey(relationship.sourcePath, relationship.id);
    if (relationships.has(key)) fail("invalid_presentation_native_graph", `OpenChestnut returned duplicate relationship ${relationship.id} from ${relationship.sourcePath}.`);
    relationships.set(key, relationship);
  }

  return function nativeGraph(opaque, sourcePart) {
    const references = (opaque.relationshipReferences || []).map((reference) => ({
      attribute: reference.attribute,
      id: reference.relationshipId,
      namespaceUri: reference.namespaceUri,
    }));
    const rootRelationships = [];
    const seenIds = new Set();
    for (const reference of references) {
      if (seenIds.has(reference.id)) continue;
      seenIds.add(reference.id);
      const relationship = relationships.get(relationshipKey(sourcePart, reference.id));
      if (!relationship) fail("missing_presentation_native_relationship", `OpenChestnut native object in ${sourcePart} references missing relationship ${reference.id}.`);
      rootRelationships.push(modelRelationship(relationship));
    }
    const parts = (opaque.preservedPartPaths || []).map((partPath) => materializedParts.get(safePartPath(partPath)));
    if (parts.some((part) => !part)) fail("missing_presentation_native_part", `OpenChestnut native-object graph contains unresolved part metadata.`);
    return { relationshipReferences: references, rootRelationships, parts };
  };
}

export function presentationNativeGraphSnapshot(object, { ignoredPartPaths = [] } = {}) {
  const ignored = new Set(ignoredPartPaths);
  return {
    relationshipReferences: object.relationshipReferences,
    rootRelationships: object.rootRelationships,
    parts: (object.parts || []).map((part) => ({
      path: part.path,
      contentType: part.contentType,
      relationships: part.relationships,
      sourceSha256: part.sourceSha256,
      ...(ignored.has(part.path) ? {} : { bytes: part.bytes?.length || 0, sha256: sha256(part.bytes || []) }),
    })),
  };
}
