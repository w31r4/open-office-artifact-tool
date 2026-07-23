import { create, toBinary } from "@bufbuild/protobuf";
import { ChartElement, GroupShape, ImageElement, Presentation, Shape, Slide, TableElement } from "../presentation/index.mjs";
import {
  ArtifactFamily,
  PresentationDiagramTextNodeSchema,
  PresentationModernCommentAnchor_Kind,
  PresentationSlideSchema,
  PresentationSlideGuide_Orientation,
} from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { normalizePresentationRunLink } from "../presentation/ooxml-hyperlinks.mjs";
import { planPresentationCustomShows } from "../presentation/ooxml-custom-shows.mjs";
import { planPresentationSections } from "../presentation/ooxml-sections.mjs";
import { normalizePresentationTransition, PRESENTATION_TRANSITION_CAPABILITY } from "../presentation/ooxml-transitions.mjs";
import { deterministicPresentationGuid } from "../presentation/ooxml-modern-comments.mjs";
import { normalizePresentationThemeConfig } from "../presentation/ooxml-theme.mjs";
import { normalizePresentationTextBodyProperties } from "../presentation/text-body-properties.mjs";
import { effectivePresentationImageCrop, presentationImageCropFromWire, presentationImageCropToWire } from "../presentation/image-crop.mjs";
import { isPresentationAutoNumberType, normalizePresentationParagraphs, normalizePresentationParagraphStyles } from "../presentation/text-paragraphs.mjs";
import { resolveColorToken } from "../shared/colors.mjs";
import { createPresentationAssetCatalog, validatePictureBulletUri } from "./open-chestnut-assets.mjs";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";
import { modelPresentationChartFromWire, presentationChartToWire } from "./open-chestnut-presentation-charts.mjs";
import { materializePresentationNativeGraphs, presentationNativeGraphSnapshot } from "./open-chestnut-presentation-native.mjs";

const EMU_PER_PIXEL = 9525;
const EMU_PER_POINT = 12700;
const POINTS_PER_PIXEL = 0.75;
const MAX_FONT_SIZE_PIXELS = 1024;
const MAX_PARAGRAPH_COORDINATE_EMU = 51_206_400;
const MAX_TEXT_BODY_INSET_EMU = 2_147_483_647;
const ROTATION_UNITS_PER_DEGREE = 60_000;
const MAX_PARAGRAPH_SPACING_POINTS = 1584;
const MAX_PARAGRAPH_SPACING_MULTIPLIER = 132;
const PRESENTATION_STATE = Symbol.for("open-office-artifact-tool.open-chestnut-presentation-state");
const PRESENTATION_SLIDE_DUPLICATOR = Symbol.for("open-office-artifact-tool.open-chestnut-presentation-duplicate");
const PRESENTATION_SPEAKER_NOTES_CAPABILITY = Symbol.for("open-office-artifact-tool.open-chestnut-speaker-notes-capability");
const PRESENTATION_LEGACY_COMMENTS_CAPABILITY = Symbol.for("open-office-artifact-tool.open-chestnut-legacy-comments-capability");
const PRESENTATION_SCHEME_COLORS = new Set([
  "dk1", "lt1", "dk2", "lt2", "tx1", "bg1", "tx2", "bg2",
  "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink",
]);
const SOURCE_FREE_LAYOUT_TYPES = new Map([
  ["blank", "blank"],
  ["title", "title"],
  ["titleOnly", "titleOnly"],
  ["title-only", "titleOnly"],
  ["obj", "obj"],
  ["object", "obj"],
  ["content", "obj"],
  ["titleAndContent", "obj"],
  ["title-and-content", "obj"],
]);
const SOURCE_FREE_TEXT_PLACEHOLDER_TYPES = new Set(["title", "body", "ctrTitle", "subTitle"]);
const DEFAULT_PRESENTATION_THEME = JSON.stringify(normalizePresentationThemeConfig({}));
const RUN_STYLE_KEYS = new Set(["bold", "italic", "fontSize", "fontFamily", "color"]);
const TEXT_FRAME_PARAGRAPH_KEYS = new Set(["alignment", "tabStops", "marginLeft", "indent", "lineSpacing", "spaceBefore", "spaceBeforePercent", "spaceAfter", "spaceAfterPercent"]);
const PARAGRAPH_KEYS = new Set([
  "runs", "level", "alignment", "style", "bulletCharacter", "autoNumber", "bulletImage", "bulletNone",
  "bulletFont", "bulletFontFollowText", "bulletColor", "bulletColorFollowText",
  "bulletSize", "bulletSizePercent", "bulletSizeFollowText", "tabStops", "marginLeft", "indent",
  "lineSpacing", "spaceBefore", "spaceBeforePercent", "spaceAfter", "spaceAfterPercent",
]);
function modelPresentationSlideGuides(viewProperties) {
  return (viewProperties?.slideGuides || []).map((guide) => ({
    orientation: guide.orientation === PresentationSlideGuide_Orientation.VERTICAL ? "vertical" : "horizontal",
    position: Number(guide.position),
  }));
}

function modelPresentationView(viewProperties) {
  if (!viewProperties) return undefined;
  return {
    ...(viewProperties.gridSpacingCxEmu === undefined ? {} : { gridSpacingCxEmu: Number(viewProperties.gridSpacingCxEmu) }),
    ...(viewProperties.gridSpacingCyEmu === undefined ? {} : { gridSpacingCyEmu: Number(viewProperties.gridSpacingCyEmu) }),
    ...(viewProperties.slideViewSnapToGrid === undefined ? {} : { slideViewSnapToGrid: viewProperties.slideViewSnapToGrid }),
    ...(viewProperties.slideViewSnapToObjects === undefined ? {} : { slideViewSnapToObjects: viewProperties.slideViewSnapToObjects }),
    slideViewShowGuides: false,
    slideGuides: modelPresentationSlideGuides(viewProperties),
  };
}

function assertTrustedPresentationState(state) {
  if (!state) return;
  const sourceHash = String(state.source?.packageSha256 || "").toLowerCase();
  const snapshot = state.opaqueOpc?.sourcePackage;
  const snapshotHash = String(snapshot?.sha256 || "").toLowerCase();
  if (!sourceHash || !snapshotHash || sourceHash !== snapshotHash || !snapshot?.data?.length) {
    throw new OpenChestnutCodecError("PPTX source-bound export requires its validated source package snapshot.", [], { code: "missing_source_package" });
  }
}

// A source-bound slide stays attached to its imported SlidePart by object
// identity, not by whichever array index it happens to occupy now. Clone
// instances deliberately live in their own map: their source points at an
// origin Part, but they must never masquerade as a second binding to it.
function presentationSourceSlideStateMap(presentation, state) {
  if (!state) return undefined;
  const sourceBySlide = new Map();
  for (const sourceState of state.slides || []) {
    if (!(sourceState?.slide instanceof Slide) || sourceState.slide.presentation !== presentation || sourceBySlide.has(sourceState.slide)) {
      throw new OpenChestnutCodecError("Imported presentation source bindings are invalid or ambiguous.", [], { code: "presentation_topology_changed" });
    }
    sourceBySlide.set(sourceState.slide, sourceState);
  }
  const cloneBySlide = new Map();
  for (const cloneState of state.clones || []) {
    if (!(cloneState?.slide instanceof Slide) || cloneState.slide.presentation !== presentation ||
        !sourceBySlide.has(cloneState.source?.slide) || cloneBySlide.has(cloneState.slide) || sourceBySlide.has(cloneState.slide)) {
      throw new OpenChestnutCodecError("Imported presentation clone bindings are invalid or ambiguous.", [], { code: "presentation_topology_changed" });
    }
    cloneBySlide.set(cloneState.slide, cloneState);
  }
  if (presentation.slides.items.some((slide) => !sourceBySlide.has(slide) && !cloneBySlide.has(slide))) {
    throw new OpenChestnutCodecError("Source-preserving PPTX export does not accept newly added slides. Use a supported imported-slide clone operation or a source-free presentation.", [], { code: "presentation_topology_changed" });
  }
  return { sourceBySlide, cloneBySlide };
}

function clonedPresentationValue(value) {
  return value === undefined ? undefined : structuredClone(value);
}

function isPresentationConnectorElement(element) {
  return element?.kind === "connector" && typeof element.id === "string";
}

function createPresentationCloneContext() {
  return {
    cloneIdBySourceId: new Map(),
    sourceIdByCloneId: new Map(),
    pendingConnectors: [],
  };
}

// The pending clone has fresh public model IDs, whereas its first export must
// still prove equality with the origin's source-bound wire. Keep that identity
// translation private to the clone transaction instead of leaking native IDs
// through the Agent-facing model.
function registerPresentationCloneElement(context, source, clone) {
  const sourceId = String(source?.id || "");
  const cloneId = String(clone?.id || "");
  if (!sourceId || !cloneId || context.cloneIdBySourceId.has(sourceId) || context.sourceIdByCloneId.has(cloneId)) {
    throw new OpenChestnutCodecError("Imported presentation clone element identities are invalid or ambiguous.", [], { code: "unsupported_presentation_slide_clone" });
  }
  context.cloneIdBySourceId.set(sourceId, cloneId);
  context.sourceIdByCloneId.set(cloneId, sourceId);
  return clone;
}

function cloneImportedPresentationShape(container, source, context) {
  const clone = container.shapes.add({
    name: source.name,
    geometry: source.geometry,
    ...(source.customPaths?.length ? { customPaths: clonedPresentationValue(source.customPaths) } : {}),
    position: clonedPresentationValue(source.position),
    ...(source.transform ? { transform: clonedPresentationValue(source.transform) } : {}),
    fill: clonedPresentationValue(source.fill),
    line: clonedPresentationValue(source.line),
    ...(source.borderRadius === undefined ? {} : { borderRadius: source.borderRadius }),
    ...(source.shadow ? { shadow: clonedPresentationValue(source.shadow) } : {}),
    ...(source.placeholder ? { placeholder: clonedPresentationValue(source.placeholder) } : {}),
    ...(source.useBackgroundFill === undefined ? {} : { _openChestnutUseBackgroundFill: source.useBackgroundFill }),
    text: clonedPresentationValue(source.text.paragraphs),
    textBodyProperties: clonedPresentationValue(source.text.bodyProperties),
    textStyle: clonedPresentationValue(source.text.style),
  });
  clone.text.inheritedParagraphStyles = clonedPresentationValue(source.text.inheritedParagraphStyles);
  return registerPresentationCloneElement(context, source, clone);
}

// A clone needs a fresh model object so it cannot share mutable JavaScript
// identity with its origin. Its embedded asset stays content-addressed, and
// the native exporter deliberately shares that immutable ImagePart through a
// new relationship on the clone SlidePart.
function cloneImportedPresentationImage(container, source, context) {
  const clone = container.images.add({
    name: source.name,
    position: clonedPresentationValue(source.position),
    alt: source.alt,
    dataUrl: source.dataUrl,
    fit: source.fit,
    ...(source.crop ? { crop: clonedPresentationValue(source.crop) } : {}),
    geometry: source.geometry,
    ...(source.transform ? { transform: clonedPresentationValue(source.transform) } : {}),
  });
  return registerPresentationCloneElement(context, source, clone);
}

// Canonical tables are an accepted GraphicFrame leaf whose bounded DrawingML
// payload is inline in the SlidePart, so this creates a fresh model without
// copying an OPC part or relationship.
function cloneImportedPresentationTable(container, source, context) {
  const clone = container.tables.add({
    name: source.name,
    position: clonedPresentationValue(source.position),
    rows: source.rows,
    columns: source.columns,
    values: clonedPresentationValue(source.values),
    ...(source.style === undefined ? {} : { style: clonedPresentationValue(source.style) }),
    ...(source.styleOptions === undefined ? {} : { styleOptions: clonedPresentationValue(source.styleOptions) }),
    mergeRanges: clonedPresentationValue(source.mergeRanges),
  });
  if (source.border !== undefined) clone.border = clonedPresentationValue(source.border);
  return registerPresentationCloneElement(context, source, clone);
}

// A recognized literal-data chart is the one accepted relationship-owning
// GraphicFrame leaf. The JavaScript model must be independent immediately;
// OpenChestnut then copies the verified closed ChartPart into a distinct OPC
// part so origin and clone can be edited independently after export/reimport.
function cloneImportedPresentationChart(container, source, context) {
  if (source.externalData) {
    throw new OpenChestnutCodecError("The bounded imported-slide clone profile does not accept charts with embedded or external workbook data.", [], { code: "unsupported_presentation_slide_clone" });
  }
  const clone = container.charts.add(source.chartType, {
    name: source.name,
    position: clonedPresentationValue(source.position),
    title: source.title,
    categories: clonedPresentationValue(source.categories),
    series: clonedPresentationValue(source.series),
    axes: clonedPresentationValue(source.axes),
    legend: clonedPresentationValue(source.legend),
    dataLabels: clonedPresentationValue(source.dataLabels),
    ...(source.styleId === undefined ? {} : { styleId: source.styleId }),
    varyColors: source.varyColors,
    barOptions: clonedPresentationValue(source.barOptions),
    lineOptions: clonedPresentationValue(source.lineOptions),
  });
  return registerPresentationCloneElement(context, source, clone);
}

const CLONE_DIAGRAM_RELATIONSHIPS = new Map([
  ["dm", "/diagramData"],
  ["lo", "/diagramLayout"],
  ["qs", "/diagramQuickStyle"],
  ["cs", "/diagramColors"],
]);
const CLONE_DIAGRAM_CONTENT_TYPES = new Set([
  "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
  "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
  "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
  "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
]);
const CLONE_INK_CONTENT_TYPE = "application/inkml+xml";
const CLONE_MP4_CONTENT_TYPE = "video/mp4";
const CLONE_VIDEO_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
const CLONE_MEDIA_RELATIONSHIP = "http://schemas.microsoft.com/office/2007/relationships/media";
const CLONE_IMAGE_RELATIONSHIPS = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/image",
]);
const CLONE_CUSTOM_XML_RELATIONSHIPS = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml",
]);
const CLONE_RELATIONSHIP_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);

function isCloneDiagramGraphicFrame(source) {
  return /^<(?:[A-Za-z_][\w.-]*:)?graphicFrame(?:\s|>)/.test(String(source?.rawXml || "").trimStart());
}

function cloneDiagramReferenceIds(source) {
  const references = source?.relationshipReferences;
  if (!Array.isArray(references) || references.length !== 4) return undefined;
  const ids = new Map();
  for (const reference of references) {
    const localName = String(reference.attribute || "").split(":").at(-1);
    const id = String(reference.id ?? reference.relationshipId ?? "");
    if (!CLONE_DIAGRAM_RELATIONSHIPS.has(localName) || !CLONE_RELATIONSHIP_NAMESPACES.has(String(reference.namespaceUri || "")) ||
        !id || [...ids.values()].includes(id) || ids.has(localName)) return undefined;
    ids.set(localName, id);
  }
  return ids.size === 4 ? ids : undefined;
}

function cloneablePresentationDiagramWire(source) {
  const ids = source?.nativeKind === "diagram" ? cloneDiagramReferenceIds(source) : undefined;
  const paths = source?.preservedPartPaths;
  return Boolean(ids && isCloneDiagramGraphicFrame(source) && Array.isArray(paths) && paths.length === 4 && new Set(paths.map(String)).size === 4);
}

function cloneablePresentationDiagramModel(source) {
  const ids = source?.kind === "nativeObject" && source.nativeKind === "diagram" ? cloneDiagramReferenceIds(source) : undefined;
  if (!ids || !isCloneDiagramGraphicFrame(source) || source.oleWorkbook || !Array.isArray(source.rootRelationships) || source.rootRelationships.length !== 4 ||
      !Array.isArray(source.parts) || source.parts.length !== 4) return false;
  const roots = new Map(source.rootRelationships.map((relationship) => [relationship.id, relationship]));
  if (roots.size !== 4) return false;
  for (const [localName, id] of ids) {
    const relationship = roots.get(id);
    if (!relationship || String(relationship.targetMode || "").toLowerCase() === "external" ||
        !String(relationship.type || "").endsWith(CLONE_DIAGRAM_RELATIONSHIPS.get(localName))) return false;
  }
  const contentTypes = new Set();
  const paths = new Set();
  for (const part of source.parts) {
    if (!CLONE_DIAGRAM_CONTENT_TYPES.has(part.contentType) || contentTypes.has(part.contentType) ||
        !part.path || paths.has(part.path) || !part.bytes?.length || !/^[0-9a-f]{64}$/i.test(String(part.sourceSha256 || "")) ||
        !Array.isArray(part.relationships) || part.relationships.length !== 0) return false;
    contentTypes.add(part.contentType);
    paths.add(part.path);
  }
  return contentTypes.size === 4;
}

function isCloneInkContentPart(source) {
  return /^<(?:[A-Za-z_][\w.-]*:)?contentPart(?:\s|>)/.test(String(source?.rawXml || "").trimStart());
}

function cloneInkContentReference(source) {
  const references = source?.relationshipReferences;
  if (!Array.isArray(references) || references.length !== 1) return undefined;
  const reference = references[0];
  return String(reference.attribute || "").split(":").at(-1) === "id" &&
    CLONE_RELATIONSHIP_NAMESPACES.has(String(reference.namespaceUri || "")) &&
    String(reference.id ?? reference.relationshipId ?? "")
    ? String(reference.id ?? reference.relationshipId)
    : undefined;
}

function cloneablePresentationInkContentWire(source) {
  const relationshipId = source?.nativeKind === "contentPart" ? cloneInkContentReference(source) : undefined;
  return Boolean(relationshipId && isCloneInkContentPart(source) && Array.isArray(source.preservedPartPaths) &&
    source.preservedPartPaths.length === 1 && source.preservedPartPaths[0]);
}

function cloneablePresentationInkContentModel(source) {
  const relationshipId = source?.kind === "nativeObject" && source.nativeKind === "contentPart"
    ? cloneInkContentReference(source)
    : undefined;
  if (!relationshipId || !isCloneInkContentPart(source) || source.oleWorkbook ||
      !Array.isArray(source.rootRelationships) || source.rootRelationships.length !== 1 ||
      !Array.isArray(source.parts) || source.parts.length !== 1) return false;
  const relationship = source.rootRelationships[0];
  const part = source.parts[0];
  return relationship.id === relationshipId &&
    String(relationship.targetMode || "").toLowerCase() !== "external" &&
    CLONE_CUSTOM_XML_RELATIONSHIPS.has(String(relationship.type || "")) &&
    part.contentType === CLONE_INK_CONTENT_TYPE && Boolean(part.path) && Boolean(part.bytes?.length) &&
    /^[0-9a-f]{64}$/i.test(String(part.sourceSha256 || "")) &&
    Array.isArray(part.relationships) && part.relationships.length === 0;
}

function isCloneMediaPicture(source) {
  return /^<(?:[A-Za-z_][\w.-]*:)?pic(?:\s|>)/.test(String(source?.rawXml || "").trimStart());
}

function cloneMediaXmlTags(rawXml, localName) {
  return [...String(rawXml || "").matchAll(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>`, "gi"))]
    .map((match) => match[0]);
}

function cloneMediaXmlTag(rawXml, localName) {
  const matches = cloneMediaXmlTags(rawXml, localName);
  return matches.length === 1 ? matches[0] : undefined;
}

function cloneMediaXmlAttribute(tag, localName) {
  return new RegExp(`\\s(?:[A-Za-z_][\\w.-]*:)?${localName}="([^"]*)"`, "i").exec(String(tag || ""))?.[1];
}

function cloneablePresentationMediaMarkup(source, ids) {
  const rawXml = String(source?.rawXml || "");
  const click = cloneMediaXmlTag(rawXml, "hlinkClick");
  const video = cloneMediaXmlTag(rawXml, "videoFile");
  const media = cloneMediaXmlTag(rawXml, "media");
  const extensions = cloneMediaXmlTags(rawXml, "ext")
    .filter((tag) => cloneMediaXmlAttribute(tag, "uri") !== undefined);
  const extension = extensions.length === 1 &&
    cloneMediaXmlAttribute(extensions[0], "uri") === "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"
    ? extensions[0]
    : undefined;
  const blip = cloneMediaXmlTag(rawXml, "blip");
  if (!click || !video || !media || !extension || !blip || /<(?:[A-Za-z_][\w.-]*:)?audioFile\b/i.test(rawXml) ||
      cloneMediaXmlAttribute(click, "id") !== "" || cloneMediaXmlAttribute(click, "action") !== "ppaction://media" ||
      cloneMediaXmlAttribute(extension, "uri") !== "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" ||
      cloneMediaXmlAttribute(video, "link") !== ids.link) return false;
  const mediaId = cloneMediaXmlAttribute(media, "embed");
  const posterId = cloneMediaXmlAttribute(blip, "embed");
  return Boolean(mediaId && posterId && mediaId !== posterId && ids.embeds.includes(mediaId) && ids.embeds.includes(posterId));
}

function cloneMediaReferenceIds(source) {
  const references = source?.relationshipReferences;
  if (!Array.isArray(references) || references.length !== 3) return undefined;
  const byAttribute = new Map();
  const seenIds = new Set();
  for (const reference of references) {
    const attribute = String(reference.attribute || "").split(":").at(-1);
    const id = String(reference.id ?? reference.relationshipId ?? "");
    if (!CLONE_RELATIONSHIP_NAMESPACES.has(String(reference.namespaceUri || "")) || !id || seenIds.has(id) || !new Set(["link", "embed"]).has(attribute)) return undefined;
    seenIds.add(id);
    const values = byAttribute.get(attribute) || [];
    values.push(id);
    byAttribute.set(attribute, values);
  }
  const links = byAttribute.get("link") || [];
  const embeds = byAttribute.get("embed") || [];
  return links.length === 1 && embeds.length === 2 ? { link: links[0], embeds } : undefined;
}

function cloneablePresentationMediaWire(source) {
  const ids = source?.nativeKind === "media" ? cloneMediaReferenceIds(source) : undefined;
  return Boolean(ids && isCloneMediaPicture(source) && cloneablePresentationMediaMarkup(source, ids) && Array.isArray(source.preservedPartPaths) &&
    source.preservedPartPaths.length === 2 && new Set(source.preservedPartPaths.map(String)).size === 2);
}

function cloneablePresentationMediaModel(source) {
  const ids = source?.kind === "nativeObject" && source.nativeKind === "media" ? cloneMediaReferenceIds(source) : undefined;
  if (!ids || !isCloneMediaPicture(source) || !cloneablePresentationMediaMarkup(source, ids) || source.oleWorkbook || !Array.isArray(source.rootRelationships) ||
      source.rootRelationships.length !== 3 || !Array.isArray(source.parts) || source.parts.length !== 2) return false;
  const relationships = new Map(source.rootRelationships.map((relationship) => [relationship.id, relationship]));
  if (relationships.size !== 3) return false;
  const video = relationships.get(ids.link);
  const mediaId = ids.embeds.find((id) => relationships.get(id)?.type === CLONE_MEDIA_RELATIONSHIP);
  const imageId = ids.embeds.find((id) => CLONE_IMAGE_RELATIONSHIPS.has(relationships.get(id)?.type));
  const media = relationships.get(mediaId);
  const image = relationships.get(imageId);
  if (!video || video.type !== CLONE_VIDEO_RELATIONSHIP || !media || !image || mediaId === imageId ||
      String(video.targetMode || "").toLowerCase() === "external" || String(media.targetMode || "").toLowerCase() === "external" ||
      String(image.targetMode || "").toLowerCase() === "external" || !video.target || video.target !== media.target || video.target === image.target) return false;
  const mp4Parts = source.parts.filter((part) => part.contentType === CLONE_MP4_CONTENT_TYPE && /^(?:ppt\/)?media\/[^/]+\.mp4$/i.test(String(part.path || "")));
  const posterParts = source.parts.filter((part) => /^image\/(?:png|jpeg)$/i.test(String(part.contentType || "")) && /^ppt\/media\/[^/]+$/i.test(String(part.path || "")));
  return mp4Parts.length === 1 && posterParts.length === 1 && source.parts.every((part) =>
    Boolean(part.bytes?.length) && /^[0-9a-f]{64}$/i.test(String(part.sourceSha256 || "")) &&
    Array.isArray(part.relationships) && part.relationships.length === 0);
}

// Eligible OLE, SmartArt, InkML, and embedded MP4 objects remain opaque PresentationML, but their
// package graphs have already been proved closed. Give the pending slide clone
// a fresh JavaScript object while retaining the exact source graph snapshot;
// C# allocates independent mutable parts during the first export.
function cloneImportedPresentationNativeObject(container, source, context) {
  const cloneableOle = source?.kind === "nativeObject" && source.nativeKind === "oleObject" && source.oleWorkbook &&
    !source._embeddedWorkbookReplacementBytes?.();
  const cloneableDiagram = cloneablePresentationDiagramModel(source) && !source?._diagramTextReplacement?.();
  if (!cloneableOle && !cloneableDiagram && !cloneablePresentationInkContentModel(source) && !cloneablePresentationMediaModel(source)) {
    throw new OpenChestnutCodecError("The bounded imported-slide clone profile accepts only an unchanged eligible embedded-XLSX OLE object, canonical closed four-part SmartArt frame, canonical top-level closed InkML content part, or canonical top-level closed MP4 media picture.", [], { code: "unsupported_presentation_slide_clone" });
  }
  const clone = container.nativeObjects.add({
    name: source.name,
    nativeId: source.nativeId,
    creationId: source.creationId,
    nativeKind: source.nativeKind,
    position: clonedPresentationValue(source.position),
    rawXml: source.rawXml,
    sourcePart: source.sourcePart,
    relationshipReferences: clonedPresentationValue(source.relationshipReferences),
    rootRelationships: clonedPresentationValue(source.rootRelationships),
    parts: clonedPresentationValue(source.parts),
    oleWorkbook: clonedPresentationValue(source.oleWorkbook),
    diagramText: clonedPresentationValue(source._diagramTextSourceBinding?.()),
  });
  return registerPresentationCloneElement(context, source, clone);
}

function cloneImportedPresentationConnector(container, source, context) {
  const clone = container.connectors.add({
    name: source.name,
    connectorType: source.connectorType,
    start: clonedPresentationValue(source.start),
    end: clonedPresentationValue(source.end),
    line: clonedPresentationValue(source.line),
  });
  registerPresentationCloneElement(context, source, clone);
  context.pendingConnectors.push({ source, clone });
  return clone;
}

// A group is not automatically a clone leaf: it can contain connectors,
// native objects, or external edges. This recursive helper is called only
// after the source wire tree has proved every descendant is one of the same
// bounded clone-safe element kinds. A chart descendant is accepted only when
// the native preflight proves its ChartPart is closed. Each new model object
// has a fresh JS identity; the source bindings still make the pending clone
// immutable until its export/reimport boundary.
function cloneImportedPresentationGroup(container, source, context) {
  const clone = container.groups.add({
    name: source.name,
    position: clonedPresentationValue(source.position),
    childFrame: clonedPresentationValue(source.childFrame),
  });
  registerPresentationCloneElement(context, source, clone);
  for (const child of source.children) cloneImportedPresentationElement(clone, child, context);
  return clone;
}

function cloneImportedPresentationElement(container, source, context) {
  if (source instanceof Shape) return cloneImportedPresentationShape(container, source, context);
  if (source instanceof TableElement) return cloneImportedPresentationTable(container, source, context);
  if (source instanceof ChartElement) return cloneImportedPresentationChart(container, source, context);
  if (source instanceof ImageElement) return cloneImportedPresentationImage(container, source, context);
  if (isPresentationConnectorElement(source)) return cloneImportedPresentationConnector(container, source, context);
  if (source instanceof GroupShape) return cloneImportedPresentationGroup(container, source, context);
  if (source?.kind === "nativeObject") return cloneImportedPresentationNativeObject(container, source, context);
  throw new OpenChestnutCodecError("The bounded imported-slide clone profile encountered an unsupported group descendant.", [], { code: "unsupported_presentation_slide_clone" });
}

function cloneSupportedPresentationContent(content, allowNativeGraphLeaf = true) {
  if (content?.case === "shape" || content?.case === "table" || content?.case === "chart" || content?.case === "image" || content?.case === "connector") return true;
  if (content?.case === "opaque") {
    const opaque = content.value;
    const cloneableOle = opaque?.nativeKind === "oleObject" && Boolean(opaque.oleWorkbook?.partPath) &&
      Boolean(opaque.oleWorkbook?.sourceSha256) && !opaque.oleWorkbook?.replacementAssetId;
    return allowNativeGraphLeaf && (cloneableOle || cloneablePresentationDiagramWire(opaque) || cloneablePresentationInkContentWire(opaque) || cloneablePresentationMediaWire(opaque));
  }
  if (content?.case !== "group") return false;
  const children = content.value?.children;
  return Array.isArray(children) && children.length > 0 && children.every((child) => cloneSupportedPresentationContent(child?.content, false));
}

function collectPresentationCloneSourceIds(source, ids, allowNativeGraphLeaf = true) {
  const cloneableOle = allowNativeGraphLeaf && source?.kind === "nativeObject" && source.nativeKind === "oleObject" && Boolean(source.oleWorkbook) && !source._embeddedWorkbookReplacementBytes?.();
  const cloneableDiagram = allowNativeGraphLeaf && cloneablePresentationDiagramModel(source) && !source?._diagramTextReplacement?.();
  const cloneableInkContent = allowNativeGraphLeaf && cloneablePresentationInkContentModel(source);
  const cloneableMedia = allowNativeGraphLeaf && cloneablePresentationMediaModel(source);
  if (!(source instanceof Shape) && !(source instanceof TableElement) && !(source instanceof ChartElement) && !(source instanceof ImageElement) && !isPresentationConnectorElement(source) && !(source instanceof GroupShape) && !cloneableOle && !cloneableDiagram && !cloneableInkContent && !cloneableMedia) {
    throw new OpenChestnutCodecError("The bounded imported-slide clone profile encountered an unsupported source element.", [], { code: "unsupported_presentation_slide_clone" });
  }
  const id = String(source.id || "");
  if (!id || ids.has(id)) {
    throw new OpenChestnutCodecError("Imported presentation clone source element identities are invalid or ambiguous.", [], { code: "unsupported_presentation_slide_clone" });
  }
  ids.add(id);
  if (source instanceof GroupShape) {
    for (const child of source.children) collectPresentationCloneSourceIds(child, ids, false);
  }
}

function assertPresentationCloneConnectorTargets(source, sourceIds) {
  if (isPresentationConnectorElement(source)) {
    for (const targetId of [source.startTargetId, source.endTargetId]) {
      if (targetId && !sourceIds.has(targetId)) {
        throw new OpenChestnutCodecError("A bounded imported-slide connector may target only an element cloned in the same slide tree.", [], { code: "unsupported_presentation_slide_clone" });
      }
    }
  }
  if (source instanceof GroupShape) {
    for (const child of source.children) assertPresentationCloneConnectorTargets(child, sourceIds);
  }
}

function bindPresentationCloneConnectorTargets(context) {
  for (const { source, clone } of context.pendingConnectors) {
    const targetId = (value, side) => {
      if (!value) return undefined;
      const cloneTargetId = context.cloneIdBySourceId.get(value);
      if (!cloneTargetId) {
        throw new OpenChestnutCodecError(`Imported presentation clone connector ${source.id} has an unresolved ${side} target.`, [], { code: "unsupported_presentation_slide_clone" });
      }
      return cloneTargetId;
    };
    clone.startTargetId = targetId(source.startTargetId, "start");
    clone.endTargetId = targetId(source.endTargetId, "end");
  }
}

// A legacy comment has no JavaScript object identity that may be shared with
// its origin. Copy the imported thread record into a fresh slide model while
// retaining its native author/index evidence; the C# clone preflight then
// proves the clone-local comments XML and shared immutable author catalog are
// unchanged before writing any OPC graph.
function cloneImportedPresentationLegacyComments(slide, source) {
  for (const thread of source.comments.items) {
    const snapshot = clonedPresentationValue(thread.toJSON());
    slide.comments.addThread(undefined, snapshot.comments?.[0]?.text || "", snapshot);
  }
}

function duplicateImportedPresentationSlide(presentation, state, slide) {
  const source = (state.slides || []).find((entry) => entry.slide === slide);
  if (!source) {
    throw new OpenChestnutCodecError("Only an original imported PPTX slide can be duplicated in this bounded clone profile.", [], { code: "unsupported_presentation_slide_clone" });
  }
  if ((state.clones || []).some((entry) => entry.source === source)) {
    throw new OpenChestnutCodecError("The bounded imported-slide clone profile permits only one pending clone per origin; export and reimport it before cloning that source again.", [], { code: "unsupported_presentation_slide_clone" });
  }
  if (source.entries.some((entry) => !cloneSupportedPresentationContent(entry.wire.content))) {
    throw new OpenChestnutCodecError("The bounded imported-slide clone profile supports only canonical shapes, inline tables, closed literal-data charts, embedded images, eligible embedded-XLSX OLE frames, closed four-part SmartArt frames, top-level closed InkML content parts, top-level closed embedded-MP4 media pictures, bounded connectors, and recursively canonical groups; other native objects and graph edges require a broader OPC graph clone.", [], { code: "unsupported_presentation_slide_clone" });
  }
  const sourceIds = new Set();
  for (const entry of source.entries) collectPresentationCloneSourceIds(entry.model, sourceIds);
  for (const entry of source.entries) assertPresentationCloneConnectorTargets(entry.model, sourceIds);
  const clone = presentation.slides.insert({
    after: slide,
    name: slide.name,
    ...(slide.background?.fill ? { background: clonedPresentationValue(slide.background) } : {}),
    ...(slide.transition?.configured ? { transition: slide.transition.toJSON() } : {}),
    ...(source.wire.speakerNotes
      ? { notes: source.wire.speakerNotes.textBody ? slide.speakerNotes?.textFrame?.paragraphs || [] : slide.speakerNotes?.text || "" }
      : {}),
  });
  clone.layoutId = slide.layoutId;
  cloneImportedPresentationLegacyComments(clone, slide);
  const cloneContext = createPresentationCloneContext();
  const entries = source.entries.map((entry) => {
    const model = cloneImportedPresentationElement(clone, entry.model, cloneContext);
    return {
      wire: entry.wire,
      model,
      placeholderSnapshot: entry.wire.content.case === "shape" && entry.wire.content.value.placeholder
        ? slidePlaceholderState(model)
        : undefined,
      snapshot: entry.wire.content.case === "image"
        ? presentationImageReadOnlySnapshot(model)
        : entry.wire.content.case === "table"
          ? presentationTableReadOnlySnapshot(model)
          : entry.wire.content.case === "opaque"
            ? opaquePresentationSnapshot(model)
            : undefined,
    };
  });
  bindPresentationCloneConnectorTargets(cloneContext);
  const cloneState = {
    source,
    slide: clone,
    name: clone.name,
    commentSnapshot: presentationSlideCommentSnapshot(clone),
    entries,
    sourceIdByCloneId: cloneContext.sourceIdByCloneId,
  };
  state.clones.push(cloneState);
  return clone;
}

function presentationCloneBytes(slide) {
  const { id: _id, source: _source, cloneSource: _cloneSource, $typeName: _typeName, ...comparable } = slide;
  // Compare protobuf wire bytes instead of JavaScript object JSON. The
  // generated decoder intentionally materializes default nested messages;
  // the model serializer omits them. Canonical protobuf encoding treats those
  // forms alike while still detecting every meaningful clone mutation.
  return toBinary(PresentationSlideSchema, create(PresentationSlideSchema, comparable));
}

function presentationCloneMatches(requested, source) {
  const left = presentationCloneBytes(requested);
  const right = presentationCloneBytes(source);
  return left.byteLength === right.byteLength && left.every((value, index) => value === right[index]);
}

function emuFromPixels(value, name, { allowNegative = false } = {}) {
  const number = Number(value);
  if (!Number.isFinite(number) || (!allowNegative && number < 0)) {
    throw new OpenChestnutCodecError(`${name} must be a ${allowNegative ? "finite" : "non-negative finite"} number.`, [], { code: "invalid_presentation_frame" });
  }
  return BigInt(Math.round(number * EMU_PER_PIXEL));
}

// DrawingML permits a negative offset. The authoring profile deliberately
// does not, but an opaque source-bound element can already contain one. C#
// verifies that such an element remains semantically unchanged before it
// preserves it, so let the adapter carry exactly that source-bound value
// instead of inventing a separate per-shape escape hatch.
function sourceBoundFrameEmuFromPixels(value, name, original) {
  return emuFromPixels(value, name, { allowNegative: original?.source?.editable === false });
}

function signedEmuFromPixels(value, name) {
  const number = Number(value);
  if (!Number.isFinite(number) || Math.abs(number) > 10_000_000) throw new OpenChestnutCodecError(`${name} must be a bounded finite number.`, [], { code: "invalid_presentation_group" });
  return BigInt(Math.round(number * EMU_PER_PIXEL));
}

function paragraphEmuFromPixels(value, name, { allowNegative = false } = {}) {
  const number = Number(value);
  const emu = Math.round(number * EMU_PER_PIXEL);
  if (!Number.isFinite(number) || (!allowNegative && number < 0) || emu < (allowNegative ? -MAX_PARAGRAPH_COORDINATE_EMU : 0) || emu > MAX_PARAGRAPH_COORDINATE_EMU) {
    throw new OpenChestnutCodecError(`${name} is outside the supported DrawingML coordinate range.`, [], { code: "invalid_presentation_text" });
  }
  return BigInt(emu);
}

function presentationRgb(value, name) {
  const source = typeof value === "string" ? value : value?.color || value?.fill;
  const raw = resolveColorToken(source, source);
  if (raw == null || raw === "transparent" || raw === "none") return "";
  const match = /^#([0-9a-f]{6})$/i.exec(String(raw));
  if (!match) throw new OpenChestnutCodecError(`${name} must be transparent or a six-digit RGB color.`, [], { code: "unsupported_presentation_features" });
  return match[1].toUpperCase();
}

function presentationChart(chart, original) {
  return presentationChartToWire(chart, original, {
    emuFromPixels,
    rgb: presentationRgb,
    sourceBoundFrameEmuFromPixels,
  });
}

function modelPresentationChart(source) {
  return modelPresentationChartFromWire(source, EMU_PER_PIXEL);
}

function unsupportedStyleFields(style = {}) {
  return Object.keys(style).filter((key) => !RUN_STYLE_KEYS.has(key));
}

function wireTextStyle(style = {}, shapeId) {
  const unsupported = unsupportedStyleFields(style);
  if (unsupported.length) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported paragraph text style fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const fontSize = style.fontSize == null ? undefined : Number(style.fontSize);
  if (fontSize !== undefined && (!Number.isFinite(fontSize) || fontSize <= 0 || fontSize > MAX_FONT_SIZE_PIXELS)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a paragraph font size outside the supported 0-${MAX_FONT_SIZE_PIXELS} pixel range.`, [], { code: "invalid_presentation_text" });
  }
  const fontFamily = style.fontFamily == null ? undefined : String(style.fontFamily);
  if (fontFamily !== undefined && (!fontFamily.trim() || fontFamily.length > 255)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid paragraph font family.`, [], { code: "invalid_presentation_text" });
  }
  let color;
  if (style.color != null) {
    const token = String(style.color).trim();
    color = PRESENTATION_SCHEME_COLORS.has(token)
      ? { case: "colorScheme", value: token }
      : { case: "colorRgb", value: presentationRgb(style.color, `${shapeId}.text.paragraphStyle.color`) };
    if (!color.value) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a transparent paragraph color outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
  }
  return {
    ...(style.bold == null ? {} : { bold: Boolean(style.bold) }),
    ...(style.italic == null ? {} : { italic: Boolean(style.italic) }),
    ...(fontSize === undefined ? {} : { fontSizePoints: fontSize * POINTS_PER_PIXEL }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(color ? { color } : {}),
  };
}

function wireDefaultRunStyle(paragraph, original, shapeId) {
  if (Object.keys(paragraph.style || {}).length) {
    return { case: "defaultRunProperties", value: wireTextStyle(paragraph.style, shapeId) };
  }
  if (new Set(["defaultRunProperties", "noDefaultRunProperties"]).has(original?.defaultRunStyle?.case)) {
    return { case: "noDefaultRunProperties", value: true };
  }
  return undefined;
}

function wireHyperlink(value, original, shapeId, customShowLinks) {
  if (value == null) {
    if (new Set(["runHyperlink", "noHyperlink"]).has(original?.hyperlink?.case)) return { case: "noHyperlink", value: true };
    return undefined;
  }
  let link;
  try {
    link = normalizePresentationRunLink(value);
  } catch (error) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid run hyperlink: ${error.message}`, [], { code: "invalid_presentation_hyperlink" });
  }
  const target = link.uri
    ? { case: "uri", value: link.uri }
    : link.slideId
      ? { case: "slideId", value: link.slideId }
      : link.action
        ? { case: "action", value: link.action }
        : link.customShow
          ? { case: "customShowId", value: resolvePresentationCustomShowLinkId(link.customShow, original, shapeId, customShowLinks) }
        : undefined;
  if (!target) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an unsupported run hyperlink target.`, [], { code: "unsupported_presentation_features" });
  return {
    case: "runHyperlink",
    value: {
      target,
      ...(link.tooltip == null ? {} : { tooltip: link.tooltip }),
      ...(link.targetFrame == null ? {} : { targetFrame: link.targetFrame }),
      ...(link.history == null ? {} : { history: link.history }),
      ...(link.highlightClick == null ? {} : { highlightClick: link.highlightClick }),
      ...(link.returnToSlide == null ? {} : { returnToSlide: link.returnToSlide }),
    },
  };
}

function wireRun(run, inheritedStyle, shapeId, original, customShowLinks) {
  const unsupported = unsupportedStyleFields(run.style);
  if (unsupported.length) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported run style fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const style = { ...inheritedStyle, ...(run.style || {}) };
  const fontSize = style.fontSize == null ? undefined : Number(style.fontSize);
  if (fontSize !== undefined && (!Number.isFinite(fontSize) || fontSize <= 0 || fontSize > MAX_FONT_SIZE_PIXELS)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a font size outside the supported 0-${MAX_FONT_SIZE_PIXELS} pixel range.`, [], { code: "invalid_presentation_text" });
  }
  const fontFamily = style.fontFamily == null ? undefined : String(style.fontFamily);
  if (fontFamily !== undefined && (!fontFamily.trim() || fontFamily.length > 255)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid font family.`, [], { code: "invalid_presentation_text" });
  }
  const colorRgb = style.color == null ? undefined : presentationRgb(style.color, `${shapeId}.text.color`);
  if (colorRgb === "") {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a transparent run color outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
  }
  const hyperlink = wireHyperlink(run.link, original, shapeId, customShowLinks);
  return {
    content: run.break
      ? { case: "lineBreak", value: true }
      : run.field
        ? { case: "field", value: { id: run.field.id, type: run.field.type, text: run.field.text } }
        : { case: "text", value: String(run.text ?? "") },
    ...(style.bold == null ? {} : { bold: Boolean(style.bold) }),
    ...(style.italic == null ? {} : { italic: Boolean(style.italic) }),
    ...(fontSize === undefined ? {} : { fontSizePoints: fontSize * POINTS_PER_PIXEL }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(colorRgb === undefined ? {} : { colorRgb }),
    ...(hyperlink ? { hyperlink } : {}),
  };
}

function wireBullet(paragraph, original, shapeId, assetCatalog) {
  const choices = [paragraph.bulletCharacter != null, Boolean(paragraph.autoNumber), Boolean(paragraph.bulletImage), paragraph.bulletNone === true];
  if (choices.filter(Boolean).length > 1) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph selects more than one list marker.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletCharacter != null) {
    const character = String(paragraph.bulletCharacter);
    if ([...character].length !== 1) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} bullet character must contain one Unicode scalar value.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletCharacter", value: character };
  }
  if (paragraph.autoNumber) {
    const scheme = String(paragraph.autoNumber.type || paragraph.autoNumber.scheme || "");
    if (!isPresentationAutoNumberType(scheme)) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported auto-number scheme ${scheme || "(missing)"}.`, [], { code: "invalid_presentation_text" });
    const rawStart = paragraph.autoNumber.startAt ?? paragraph.autoNumber.start;
    const startAt = rawStart == null ? undefined : Number(rawStart);
    if (startAt !== undefined && (!Number.isInteger(startAt) || startAt < 1 || startAt > 32767)) {
      throw new OpenChestnutCodecError(`Presentation shape ${shapeId} auto-number start must be from 1 through 32767.`, [], { code: "invalid_presentation_text" });
    }
    return { case: "autoNumber", value: { scheme, ...(startAt === undefined ? {} : { startAt }) } };
  }
  if (paragraph.bulletImage) {
    if (paragraph.bulletImage.relationshipId) {
      throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an unresolved picture-bullet relationship ID.`, [], { code: "invalid_presentation_asset" });
    }
    const source = paragraph.bulletImage.dataUrl
      ? { case: "assetId", value: assetCatalog.addDataUrl(paragraph.bulletImage.dataUrl) }
      : { case: "uri", value: validatePictureBulletUri(paragraph.bulletImage.uri) };
    return { case: "pictureBullet", value: { source } };
  }
  if (paragraph.bulletNone === true || new Set(["noBullet", "bulletCharacter", "autoNumber", "pictureBullet"]).has(original?.bullet?.case)) {
    return { case: "noBullet", value: true };
  }
  return undefined;
}

function wireBulletFont(paragraph, original, shapeId) {
  if (paragraph.bulletFont != null && paragraph.bulletFontFollowText === true) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph selects both a bullet font and follow-text font.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletFont != null) {
    const family = String(paragraph.bulletFont).trim();
    if (!family || family.length > 255) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid bullet font family.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletFontFamily", value: family };
  }
  if (paragraph.bulletFontFollowText === true || new Set(["bulletFontFamily", "bulletFontFollowText"]).has(original?.bulletFont?.case)) {
    return { case: "bulletFontFollowText", value: true };
  }
  return undefined;
}

function wireBulletColor(paragraph, original, shapeId) {
  if (paragraph.bulletColor != null && paragraph.bulletColorFollowText === true) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph selects both a bullet color and follow-text color.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletColor != null) {
    const scheme = String(paragraph.bulletColor).trim();
    if (PRESENTATION_SCHEME_COLORS.has(scheme)) return { case: "bulletColorScheme", value: scheme };
    const rgb = presentationRgb(paragraph.bulletColor, `${shapeId}.text.bulletColor`);
    if (!rgb) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a transparent bullet color outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
    return { case: "bulletColorRgb", value: rgb };
  }
  if (paragraph.bulletColorFollowText === true || new Set(["bulletColorRgb", "bulletColorScheme", "bulletColorFollowText"]).has(original?.bulletColor?.case)) {
    return { case: "bulletColorFollowText", value: true };
  }
  return undefined;
}

function wireBulletSize(paragraph, original, shapeId) {
  const choices = [paragraph.bulletSize != null, paragraph.bulletSizePercent != null, paragraph.bulletSizeFollowText === true];
  if (choices.filter(Boolean).length > 1) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph selects more than one bullet size.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletSize != null) {
    const pixels = Number(paragraph.bulletSize);
    if (!Number.isFinite(pixels) || pixels < 4 / 3 || pixels > MAX_FONT_SIZE_PIXELS) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid fixed bullet size.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletSizePoints", value: pixels * POINTS_PER_PIXEL };
  }
  if (paragraph.bulletSizePercent != null) {
    const percent = Number(paragraph.bulletSizePercent);
    if (!Number.isFinite(percent) || percent < 0.25 || percent > 4) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid percentage bullet size.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletSizePercent", value: percent };
  }
  if (paragraph.bulletSizeFollowText === true || new Set(["bulletSizePoints", "bulletSizePercent", "bulletSizeFollowText"]).has(original?.bulletSize?.case)) {
    return { case: "bulletSizeFollowText", value: true };
  }
  return undefined;
}

function wireTabStops(paragraph, original, shapeId) {
  if (paragraph.tabStops?.length) {
    return { tabStops: paragraph.tabStops.map((tab) => ({ positionEmu: emuFromPixels(tab.position, `${shapeId}.text.tabStops.position`), alignment: tab.alignment })) };
  }
  if (original?.tabStops?.length || original?.noTabStops === true) return { noTabStops: true };
  return {};
}

function wireParagraphLayout(paragraph, original, shapeId) {
  const leftMargin = paragraph.marginLeft != null
    ? { case: "marginLeftEmu", value: paragraphEmuFromPixels(paragraph.marginLeft, `${shapeId}.text.marginLeft`) }
    : new Set(["marginLeftEmu", "noMarginLeft"]).has(original?.leftMargin?.case)
      ? { case: "noMarginLeft", value: true }
      : undefined;
  const indentation = paragraph.indent != null
    ? { case: "indentEmu", value: paragraphEmuFromPixels(paragraph.indent, `${shapeId}.text.indent`, { allowNegative: true }) }
    : new Set(["indentEmu", "noIndent"]).has(original?.indentation?.case)
      ? { case: "noIndent", value: true }
      : undefined;
  return {
    ...(leftMargin ? { leftMargin } : {}),
    ...(indentation ? { indentation } : {}),
  };
}

function paragraphSpacingValue(value, name, { allowZero = true, maximum }) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < (allowZero ? 0 : Number.EPSILON) || number > maximum) {
    throw new OpenChestnutCodecError(`${name} is outside the supported DrawingML spacing range.`, [], { code: "invalid_presentation_text" });
  }
  return number;
}

function paragraphSpacingPointsFromPixels(value, name, { allowZero = true } = {}) {
  const pixels = paragraphSpacingValue(value, name, { allowZero, maximum: MAX_PARAGRAPH_SPACING_POINTS / POINTS_PER_PIXEL });
  return pixels * POINTS_PER_PIXEL;
}

function wireParagraphSpacing(paragraph, original, shapeId) {
  let lineSpacing;
  if (paragraph.lineSpacing != null) {
    const value = Number(paragraph.lineSpacing);
    lineSpacing = value <= 10
      ? { case: "lineSpacingMultiplier", value: paragraphSpacingValue(value, `${shapeId}.text.lineSpacing`, { allowZero: false, maximum: MAX_PARAGRAPH_SPACING_MULTIPLIER }) }
      : { case: "lineSpacingPoints", value: paragraphSpacingPointsFromPixels(value, `${shapeId}.text.lineSpacing`, { allowZero: false }) };
  } else if (new Set(["lineSpacingPoints", "lineSpacingMultiplier", "noLineSpacing"]).has(original?.lineSpacing?.case)) {
    lineSpacing = { case: "noLineSpacing", value: true };
  }

  if (paragraph.spaceBefore != null && paragraph.spaceBeforePercent != null) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph must use either point or percentage space-before, not both.`, [], { code: "invalid_presentation_text" });
  }
  let spaceBefore;
  if (paragraph.spaceBefore != null) {
    spaceBefore = { case: "spaceBeforePoints", value: paragraphSpacingPointsFromPixels(paragraph.spaceBefore, `${shapeId}.text.spaceBefore`) };
  } else if (paragraph.spaceBeforePercent != null) {
    spaceBefore = { case: "spaceBeforeMultiplier", value: paragraphSpacingValue(paragraph.spaceBeforePercent, `${shapeId}.text.spaceBeforePercent`, { maximum: MAX_PARAGRAPH_SPACING_MULTIPLIER }) };
  } else if (new Set(["spaceBeforePoints", "spaceBeforeMultiplier", "noSpaceBefore"]).has(original?.spaceBefore?.case)) {
    spaceBefore = { case: "noSpaceBefore", value: true };
  }

  if (paragraph.spaceAfter != null && paragraph.spaceAfterPercent != null) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph must use either point or percentage space-after, not both.`, [], { code: "invalid_presentation_text" });
  }
  let spaceAfter;
  if (paragraph.spaceAfter != null) {
    spaceAfter = { case: "spaceAfterPoints", value: paragraphSpacingPointsFromPixels(paragraph.spaceAfter, `${shapeId}.text.spaceAfter`) };
  } else if (paragraph.spaceAfterPercent != null) {
    spaceAfter = { case: "spaceAfterMultiplier", value: paragraphSpacingValue(paragraph.spaceAfterPercent, `${shapeId}.text.spaceAfterPercent`, { maximum: MAX_PARAGRAPH_SPACING_MULTIPLIER }) };
  } else if (new Set(["spaceAfterPoints", "spaceAfterMultiplier", "noSpaceAfter"]).has(original?.spaceAfter?.case)) {
    spaceAfter = { case: "noSpaceAfter", value: true };
  }

  return {
    ...(lineSpacing ? { lineSpacing } : {}),
    ...(spaceBefore ? { spaceBefore } : {}),
    ...(spaceAfter ? { spaceAfter } : {}),
  };
}

function wireParagraph(paragraph, textStyle, original, shapeId, assetCatalog, { forceLevel = false, customShowLinks } = {}) {
  const unsupported = Object.keys(paragraph).filter((key) => !PARAGRAPH_KEYS.has(key));
  if (unsupported.length) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported paragraph fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const paragraphStyleUnsupported = unsupportedStyleFields(paragraph.style);
  if (paragraphStyleUnsupported.length) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported paragraph text style fields: ${paragraphStyleUnsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const level = Number(paragraph.level || 0);
  if (!Number.isInteger(level) || level < 0 || level > 8) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a paragraph level outside the supported 0-8 range.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.alignment && !new Set(["left", "center", "right", "justify"]).has(paragraph.alignment)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported paragraph alignment ${paragraph.alignment}.`, [], { code: "invalid_presentation_text" });
  }
  const originalLevel = original?.level;
  const includeLevel = forceLevel || level !== 0 || originalLevel !== undefined;
  const bullet = wireBullet(paragraph, original, shapeId, assetCatalog);
  const bulletFont = wireBulletFont(paragraph, original, shapeId);
  const bulletColor = wireBulletColor(paragraph, original, shapeId);
  const bulletSize = wireBulletSize(paragraph, original, shapeId);
  const tabs = wireTabStops(paragraph, original, shapeId);
  const layout = wireParagraphLayout(paragraph, original, shapeId);
  const spacing = wireParagraphSpacing(paragraph, original, shapeId);
  const defaultRunStyle = wireDefaultRunStyle(paragraph, original, shapeId);
  const directInheritedStyle = Object.fromEntries(Object.entries(textStyle).filter(([key]) => !Object.hasOwn(paragraph.style || {}, key)));
  return {
    ...(includeLevel ? { level } : {}),
    ...(paragraph.alignment ? { alignment: paragraph.alignment } : {}),
    runs: (paragraph.runs || []).map((run, index) => wireRun(run, directInheritedStyle, shapeId, original?.runs?.[index], customShowLinks)),
    ...(bullet ? { bullet } : {}),
    ...(bulletFont ? { bulletFont } : {}),
    ...(bulletColor ? { bulletColor } : {}),
    ...(bulletSize ? { bulletSize } : {}),
    ...tabs,
    ...layout,
    ...spacing,
    ...(defaultRunStyle ? { defaultRunStyle } : {}),
  };
}

function modelRunCase(run) {
  if (run.break) return "lineBreak";
  if (run.field) return "field";
  return "text";
}

function wireTextBodyProperties(value, original, shapeId) {
  let properties;
  try {
    properties = normalizePresentationTextBodyProperties(value);
  } catch (error) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses invalid text body properties: ${error.message}`, [], { code: "invalid_presentation_text" });
  }
  const originalProperties = original?.bodyProperties;
  const insetChoice = (key, wireName, noWireName) => {
    if (properties.insets?.[key] != null) {
      const emu = Math.round(properties.insets[key] * EMU_PER_PIXEL);
      if (emu < 0 || emu > MAX_TEXT_BODY_INSET_EMU) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an out-of-range ${key} text inset.`, [], { code: "invalid_presentation_text" });
      return { case: wireName, value: BigInt(emu) };
    }
    const originalCase = originalProperties?.[`${key}Inset`]?.case;
    return new Set([wireName, noWireName]).has(originalCase) ? { case: noWireName, value: true } : undefined;
  };
  const leftInset = insetChoice("left", "leftInsetEmu", "noLeftInset");
  const topInset = insetChoice("top", "topInsetEmu", "noTopInset");
  const rightInset = insetChoice("right", "rightInsetEmu", "noRightInset");
  const bottomInset = insetChoice("bottom", "bottomInsetEmu", "noBottomInset");
  const anchor = properties.anchor != null
    ? { case: "verticalAnchor", value: properties.anchor }
    : new Set(["verticalAnchor", "noVerticalAnchor"]).has(originalProperties?.anchor?.case)
      ? { case: "noVerticalAnchor", value: true }
      : undefined;
  const wrapping = properties.wrap != null
    ? { case: "wrap", value: properties.wrap }
    : new Set(["wrap", "noWrap"]).has(originalProperties?.wrapping?.case)
      ? { case: "noWrap", value: true }
      : undefined;
  const autoFit = properties.autoFit != null
    ? { case: "autoFitMode", value: properties.autoFit }
    : new Set(["autoFitMode", "noAutoFitMode"]).has(originalProperties?.autoFit?.case)
      ? { case: "noAutoFitMode", value: true }
      : undefined;
  const rotation = properties.rotation != null
    ? { case: "rotationAngle60000", value: Math.round(properties.rotation * ROTATION_UNITS_PER_DEGREE) }
    : new Set(["rotationAngle60000", "noRotation"]).has(originalProperties?.rotation?.case)
      ? { case: "noRotation", value: true }
      : undefined;
  const verticalText = properties.verticalText != null
    ? { case: "verticalTextMode", value: properties.verticalText }
    : new Set(["verticalTextMode", "noVerticalTextMode"]).has(originalProperties?.verticalText?.case)
      ? { case: "noVerticalTextMode", value: true }
      : undefined;
  const verticalOverflow = properties.verticalOverflow != null
    ? { case: "verticalOverflowMode", value: properties.verticalOverflow }
    : new Set(["verticalOverflowMode", "noVerticalOverflowMode"]).has(originalProperties?.verticalOverflow?.case)
      ? { case: "noVerticalOverflowMode", value: true }
      : undefined;
  const horizontalOverflow = properties.horizontalOverflow != null
    ? { case: "horizontalOverflowMode", value: properties.horizontalOverflow }
    : new Set(["horizontalOverflowMode", "noHorizontalOverflowMode"]).has(originalProperties?.horizontalOverflow?.case)
      ? { case: "noHorizontalOverflowMode", value: true }
      : undefined;
  const columnCount = properties.columns?.count != null
    ? { case: "columns", value: properties.columns.count }
    : new Set(["columns", "noColumns"]).has(originalProperties?.columnCount?.case)
      ? { case: "noColumns", value: true }
      : undefined;
  const columnSpacing = properties.columns?.spacing != null
    ? { case: "columnSpacingEmu", value: BigInt(Math.round(properties.columns.spacing * EMU_PER_PIXEL)) }
    : new Set(["columnSpacingEmu", "noColumnSpacing"]).has(originalProperties?.columnSpacing?.case)
      ? { case: "noColumnSpacing", value: true }
      : undefined;
  const columnDirection = properties.columns?.rightToLeft != null
    ? { case: "rightToLeftColumns", value: properties.columns.rightToLeft }
    : new Set(["rightToLeftColumns", "noColumnDirection"]).has(originalProperties?.columnDirection?.case)
      ? { case: "noColumnDirection", value: true }
      : undefined;
  const uprightText = properties.upright != null
    ? { case: "upright", value: properties.upright }
    : new Set(["upright", "noUpright"]).has(originalProperties?.uprightText?.case)
      ? { case: "noUpright", value: true }
      : undefined;
  if (![leftInset, topInset, rightInset, bottomInset, anchor, wrapping, autoFit, rotation, verticalText, verticalOverflow, horizontalOverflow, columnCount, columnSpacing, columnDirection, uprightText].some(Boolean)) return undefined;
  return {
    ...(leftInset ? { leftInset } : {}),
    ...(topInset ? { topInset } : {}),
    ...(rightInset ? { rightInset } : {}),
    ...(bottomInset ? { bottomInset } : {}),
    ...(anchor ? { anchor } : {}),
    ...(wrapping ? { wrapping } : {}),
    ...(autoFit ? { autoFit } : {}),
    ...(rotation ? { rotation } : {}),
    ...(verticalText ? { verticalText } : {}),
    ...(verticalOverflow ? { verticalOverflow } : {}),
    ...(horizontalOverflow ? { horizontalOverflow } : {}),
    ...(columnCount ? { columnCount } : {}),
    ...(columnSpacing ? { columnSpacing } : {}),
    ...(columnDirection ? { columnDirection } : {}),
    ...(uprightText ? { uprightText } : {}),
  };
}

function presentationTextBody(shape, original, assetCatalog, customShowLinks) {
  const textStyle = shape.text?.style || {};
  const textStyleUnsupported = Object.keys(textStyle).filter((key) => !RUN_STYLE_KEYS.has(key) && !TEXT_FRAME_PARAGRAPH_KEYS.has(key));
  if (textStyleUnsupported.length) throw new OpenChestnutCodecError(`Presentation shape ${shape.id} uses unsupported text-frame style fields: ${textStyleUnsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const inheritedRunStyle = Object.fromEntries(Object.entries(textStyle).filter(([key]) => RUN_STYLE_KEYS.has(key)));
  const inheritedParagraph = Object.fromEntries(Object.entries(textStyle).filter(([key]) => TEXT_FRAME_PARAGRAPH_KEYS.has(key)));
  const paragraphs = shape.text?.paragraphs || [];
  if (original?.textBody && (original.textBody.paragraphs.length !== paragraphs.length || original.textBody.paragraphs.some((paragraph, index) => paragraph.runs.length !== (paragraphs[index]?.runs || []).length || paragraph.runs.some((run, runIndex) => run.content?.case !== modelRunCase(paragraphs[index].runs[runIndex]))))) {
    throw new OpenChestnutCodecError(`Presentation shape ${shape.id} changed its source-bound paragraph/inline topology.`, [], { code: "presentation_text_topology_changed" });
  }
  const originalListStyles = new Map((original?.textBody?.listStyles || []).map((style) => [Number(style.level), style]));
  const inheritedParagraphStyles = Object.entries(shape.text?.inheritedParagraphStyles || {}).sort(([left], [right]) => Number(left) - Number(right));
  const listStyles = inheritedParagraphStyles.map(([level, style]) => wireParagraph(
    { ...style, level: Number(level), runs: [] },
    {},
    originalListStyles.get(Number(level)),
    shape.id,
    assetCatalog,
    { forceLevel: true, customShowLinks },
  ));
  const noListStyles = listStyles.length === 0 && (originalListStyles.size > 0 || original?.textBody?.noListStyles === true);
  const bodyProperties = wireTextBodyProperties(shape.text?.bodyProperties, original?.textBody, shape.id);
  return {
    paragraphs: paragraphs.map((paragraph, index) => wireParagraph({ ...inheritedParagraph, ...paragraph }, inheritedRunStyle, original?.textBody?.paragraphs?.[index], shape.id, assetCatalog, { customShowLinks })),
    ...(listStyles.length ? { listStyles } : {}),
    ...(noListStyles ? { noListStyles: true } : {}),
    ...(bodyProperties ? { bodyProperties } : {}),
  };
}

// Speaker notes reuse the public paragraph/run wire rather than creating a
// second rich-text format. Notes-local text is intentionally narrower than a
// slide shape: the native codec rejects relationships, fields, picture
// bullets, list styles, and body properties. An imported notes part without a
// projected textBody remains text-only, so an unchanged round trip can never
// silently turn an opaque source into a lossy rich edit request.
function presentationSpeakerNotes(slide, original, assetCatalog, customShowLinks) {
  const notes = slide.speakerNotes;
  if (original) {
    const result = { text: notes?.text || "", source: original.source };
    if (original.textBody) {
      result.textBody = presentationTextBody(
        { id: `${slide.id}/notes`, text: notes?.textFrame },
        original,
        assetCatalog,
        customShowLinks,
      );
    }
    return result;
  }
  if (!notes?.text) return undefined;
  return {
    text: notes.text,
    textBody: presentationTextBody(
      { id: `${slide.id}/notes`, text: notes.textFrame },
      undefined,
      assetCatalog,
      customShowLinks,
    ),
  };
}

const MASTER_STYLE_KINDS = [
  ["title", "titleLevels", "deletedTitleLevels"],
  ["body", "bodyLevels", "deletedBodyLevels"],
  ["other", "otherLevels", "deletedOtherLevels"],
];

function wireMasterTextStyles(master, original, assetCatalog) {
  const result = {};
  for (const [kind, levelsField, deletedField] of MASTER_STYLE_KINDS) {
    const originalLevels = new Map((original?.textStyles?.[levelsField] || []).map((style) => [Number(style.level), style]));
    const current = master.textParagraphStyles?.[kind] || {};
    const levels = Object.entries(current)
      .sort(([left], [right]) => Number(left) - Number(right))
      .map(([level, style]) => wireParagraph(
        { ...style, level: Number(level), runs: [] },
        {},
        originalLevels.get(Number(level)),
        `master ${master.id} ${kind} level ${Number(level) + 1}`,
        assetCatalog,
        { forceLevel: true },
      ));
    const currentLevels = new Set(Object.keys(current).map(Number));
    const deleted = [...originalLevels.keys()].filter((level) => !currentLevels.has(level)).sort((left, right) => left - right);
    result[levelsField] = levels;
    if (deleted.length) result[deletedField] = deleted;
  }
  return result;
}

function wireBackground(background, ownerId) {
  if (!background) return undefined;
  const fill = String(background.fill || "").trim();
  const color = PRESENTATION_SCHEME_COLORS.has(fill)
    ? { case: "colorScheme", value: fill }
    : { case: "colorRgb", value: presentationRgb(fill, `${ownerId}.background.fill`) };
  if (!color.value) throw new OpenChestnutCodecError(`Presentation ${ownerId} uses an unsupported transparent background.`, [], { code: "unsupported_presentation_features" });
  if (background.mode === "reference") {
    const index = Number(background.index);
    if (!Number.isInteger(index) || index < 0 || index > 4_294_967_295) {
      throw new OpenChestnutCodecError(`Presentation ${ownerId} background reference index must be an unsigned 32-bit integer.`, [], { code: "invalid_presentation_background" });
    }
    return { color, kind: { case: "styleReferenceIndex", value: index } };
  }
  if (background.mode !== "solid") throw new OpenChestnutCodecError(`Presentation ${ownerId} background mode must be solid or reference.`, [], { code: "invalid_presentation_background" });
  return { color, kind: { case: "solid", value: true } };
}

function modelBackground(background) {
  if (!background?.color?.case || !background?.kind?.case) return undefined;
  const fill = background.color.case === "colorScheme" ? background.color.value : `#${String(background.color.value).toLowerCase()}`;
  return background.kind.case === "styleReferenceIndex"
    ? { fill, mode: "reference", index: Number(background.kind.value) }
    : { fill, mode: "solid" };
}

function wirePresentationTransition(transition) {
  const value = transition?.toJSON?.();
  if (!value) return undefined;
  return {
    effect: value.effect,
    ...(value.direction ? { direction: value.direction } : {}),
    speed: value.speed,
    advanceOnClick: value.advanceOnClick,
    ...(value.advanceAfterMs === undefined ? {} : { advanceAfterMs: value.advanceAfterMs }),
  };
}

function modelPresentationTransition(source, slideIndex) {
  if (!source) return undefined;
  if (typeof source.advanceOnClick !== "boolean") {
    throw new OpenChestnutCodecError(`OpenChestnut returned slide ${slideIndex + 1} transition without an explicit advanceOnClick value.`, [], { code: "invalid_presentation_artifact" });
  }
  try {
    return normalizePresentationTransition({
      effect: source.effect,
      ...(source.direction ? { direction: source.direction } : {}),
      speed: source.speed,
      advanceOnClick: source.advanceOnClick,
      ...(source.advanceAfterMs === undefined ? {} : { advanceAfterMs: Number(source.advanceAfterMs) }),
    });
  } catch (error) {
    throw new OpenChestnutCodecError(`OpenChestnut returned invalid slide ${slideIndex + 1} transition semantics: ${error.message}`, [], { code: "invalid_presentation_artifact" });
  }
}

function wirePresentationTransform(transform, ownerLabel) {
  if (transform == null) return {};
  if (typeof transform !== "object" || Array.isArray(transform)) {
    throw new OpenChestnutCodecError(`Presentation ${ownerLabel} transform must be an object.`, [], { code: "invalid_presentation_transform" });
  }
  const output = {};
  if (Object.hasOwn(transform, "rotationDegrees") && transform.rotationDegrees != null) {
    const degrees = Number(transform.rotationDegrees);
    if (!Number.isFinite(degrees) || degrees < -360 || degrees > 360) {
      throw new OpenChestnutCodecError(`Presentation ${ownerLabel} rotation must be between -360 and 360 degrees.`, [], { code: "invalid_presentation_transform" });
    }
    output.rotationAngle60000 = Math.round(degrees * ROTATION_UNITS_PER_DEGREE);
  }
  for (const key of ["flipHorizontal", "flipVertical"]) {
    if (!Object.hasOwn(transform, key) || transform[key] == null) continue;
    if (typeof transform[key] !== "boolean") {
      throw new OpenChestnutCodecError(`Presentation ${ownerLabel} ${key} must be a boolean.`, [], { code: "invalid_presentation_transform" });
    }
    output[key] = transform[key];
  }
  if (Object.keys(output).length === 0) {
    throw new OpenChestnutCodecError(`Presentation ${ownerLabel} transform must define rotationDegrees, flipHorizontal, or flipVertical.`, [], { code: "invalid_presentation_transform" });
  }
  return output;
}

function masterReadOnlySnapshot(master) {
  return JSON.stringify(master.toJSON());
}

function layoutReadOnlySnapshot(layout) {
  return JSON.stringify(layout.toJSON());
}

function sourceFreeLayoutType(type, layoutId) {
  const requested = String(type || "blank").trim();
  const normalized = SOURCE_FREE_LAYOUT_TYPES.get(requested);
  if (!normalized) {
    throw new OpenChestnutCodecError(
      `Presentation layout ${layoutId} uses unsupported source-free type ${requested || "(empty)"}. Use blank, title, titleOnly, or obj/titleAndContent.`,
      [],
      { code: "unsupported_presentation_features" },
    );
  }
  return normalized;
}

function sourceFreePlaceholder(placeholder, ownerId, assetCatalog, customShowLinks) {
  const type = String(placeholder.type || "");
  if (!SOURCE_FREE_TEXT_PLACEHOLDER_TYPES.has(type)) {
    throw new OpenChestnutCodecError(
      `Presentation placeholder ${placeholder.id || ownerId} uses ${type || "(empty)"}; source-free layouts currently author only title, body, ctrTitle, and subTitle text placeholders.`,
      [],
      { code: "unsupported_presentation_features" },
    );
  }
  if (!placeholder.position) {
    throw new OpenChestnutCodecError(
      `Presentation placeholder ${placeholder.id || ownerId} requires a direct position for source-free PPTX export.`,
      [],
      { code: "invalid_presentation_placeholder" },
    );
  }
  const index = Number(placeholder.idx);
  if (!Number.isInteger(index) || index < 0 || index > 4_294_967_295) {
    throw new OpenChestnutCodecError(`Presentation placeholder ${placeholder.id || ownerId} has an invalid idx.`, [], { code: "invalid_presentation_placeholder" });
  }
  const shape = {
    id: placeholder.id || `${ownerId}/placeholder/${index}`,
    text: {
      style: { ...(placeholder.style || {}) },
      paragraphs: normalizePresentationParagraphs(placeholder.text ?? ""),
      inheritedParagraphStyles: { ...(placeholder.paragraphStyles || {}) },
      bodyProperties: placeholder.textBodyProperties,
    },
  };
  const position = placeholder.position;
  return {
    id: shape.id,
    name: String(placeholder.name || `${type} placeholder`),
    type,
    index,
    textBody: presentationTextBody(shape, undefined, assetCatalog, customShowLinks),
    directFrame: {
      leftEmu: emuFromPixels(position.left, `${shape.id}.position.left`),
      topEmu: emuFromPixels(position.top, `${shape.id}.position.top`),
      widthEmu: emuFromPixels(position.width, `${shape.id}.position.width`),
      heightEmu: emuFromPixels(position.height, `${shape.id}.position.height`),
      ...wirePresentationTransform(placeholder.transform, `placeholder ${shape.id}`),
    },
  };
}

function sourceFreeSlidePlaceholder(shape) {
  if (!shape.placeholder) return undefined;
  const type = String(shape.placeholder.type || "");
  if (!SOURCE_FREE_TEXT_PLACEHOLDER_TYPES.has(type)) {
    throw new OpenChestnutCodecError(
      `Presentation slide placeholder ${shape.id} uses ${type || "(empty)"}; source-free layouts currently author only title, body, ctrTitle, and subTitle text placeholders.`,
      [],
      { code: "unsupported_presentation_features" },
    );
  }
  const index = Number(shape.placeholder.idx ?? shape.placeholder.index);
  if (!Number.isInteger(index) || index < 0 || index > 4_294_967_295) {
    throw new OpenChestnutCodecError(`Presentation slide placeholder ${shape.id} has an invalid idx.`, [], { code: "invalid_presentation_placeholder" });
  }
  const position = shape.position || {};
  return {
    placeholder: { type, index, inheritsGeometry: false },
    directFrame: {
      leftEmu: emuFromPixels(position.left, `${shape.id}.position.left`),
      topEmu: emuFromPixels(position.top, `${shape.id}.position.top`),
      widthEmu: emuFromPixels(position.width, `${shape.id}.position.width`),
      heightEmu: emuFromPixels(position.height, `${shape.id}.position.height`),
      ...wirePresentationTransform(shape.transform, `placeholder ${shape.id}`),
    },
  };
}

function presentationMasters(presentation, state, assetCatalog, customShowLinks) {
  if (state) {
    if (presentation.masters.items.length !== state.masters.length || state.masters.some((entry, index) => presentation.masters.items[index] !== entry.model)) {
      throw new OpenChestnutCodecError(`Source-preserving PPTX export requires the original ${state.masters.length}-master topology.`, [], { code: "presentation_master_topology_changed" });
    }
    return state.masters.map((entry) => {
      if (masterReadOnlySnapshot(entry.model) !== entry.snapshot) {
        throw new OpenChestnutCodecError(`Presentation master ${entry.model.id} is source-bound and read-only in OpenChestnut 0.2.`, [], { code: "unsupported_presentation_edit" });
      }
      return entry.wire;
    });
  }
  const master = presentation.master;
  return master ? [{
    id: master.id,
    name: master.name,
    textStyles: wireMasterTextStyles(master, undefined, assetCatalog),
    background: wireBackground(master.background, `master ${master.id}`),
    placeholders: master.placeholders.map((placeholder) => sourceFreePlaceholder(placeholder, master.id, assetCatalog, customShowLinks)),
  }] : [];
}

function presentationLayouts(presentation, state, assetCatalog, customShowLinks) {
  if (!state) {
    return presentation.layouts.items.map((layout) => ({
      id: layout.id,
      name: layout.name,
      masterId: layout.masterId,
      type: sourceFreeLayoutType(layout.type, layout.id),
      ...(layout.background ? { background: wireBackground(layout.background, `layout ${layout.id}`) } : {}),
      placeholders: layout.placeholders.map((placeholder) => sourceFreePlaceholder(placeholder, layout.id, assetCatalog, customShowLinks)),
    }));
  }
  if (presentation.layouts.items.length !== state.layouts.length || state.layouts.some((entry, index) => presentation.layouts.items[index] !== entry.model)) {
    throw new OpenChestnutCodecError(`Source-preserving PPTX export requires the original ${state.layouts.length}-layout topology.`, [], { code: "presentation_layout_topology_changed" });
  }
  return state.layouts.map((entry) => {
    if (layoutReadOnlySnapshot(entry.model) !== entry.snapshot) {
      throw new OpenChestnutCodecError(`Presentation layout ${entry.model.id} is source-bound and read-only in OpenChestnut 0.2.`, [], { code: "unsupported_presentation_edit" });
    }
    return entry.wire;
  });
}

function presentationShadow(shadow, shapeId) {
  if (shadow == null || shadow === false || shadow === "shadow-none") return undefined;
  const presets = {
    "shadow-sm": { color: "#000000", blurRadius: 4, distance: 2, direction: 45, opacity: 0.15 },
    shadow: { color: "#000000", blurRadius: 6, distance: 3, direction: 45, opacity: 0.18 },
    "shadow-md": { color: "#000000", blurRadius: 10, distance: 4, direction: 45, opacity: 0.2 },
    "shadow-lg": { color: "#000000", blurRadius: 15, distance: 6, direction: 45, opacity: 0.22 },
    "shadow-xl": { color: "#000000", blurRadius: 22, distance: 9, direction: 45, opacity: 0.24 },
    "shadow-2xl": { color: "#000000", blurRadius: 32, distance: 14, direction: 45, opacity: 0.25 },
  };
  let source = typeof shadow === "string" ? presets[shadow] : shadow;
  if (!source && typeof shadow === "string") {
    const match = /^(-?\d+(?:\.\d+)?)px\s+(-?\d+(?:\.\d+)?)px\s+(\d+(?:\.\d+)?)px\s+(#[0-9a-f]{6})(?:\/(\d+(?:\.\d+)?))?$/i.exec(shadow.trim());
    if (match) {
      const offsetX = Number(match[1]);
      const offsetY = Number(match[2]);
      source = {
        color: match[4],
        blurRadius: Number(match[3]),
        distance: Math.hypot(offsetX, offsetY),
        direction: (Math.atan2(offsetY, offsetX) * 180 / Math.PI + 360) % 360,
        opacity: match[5] == null ? 1 : Number(match[5]) / 100,
      };
    }
  }
  if (!source || typeof source !== "object") {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an unsupported shadow.`, [], { code: "unsupported_presentation_features" });
  }
  const blurRadius = Number(source.blurRadius ?? source.blur ?? 0);
  const distance = Number(source.distance ?? 0);
  const direction = Number(source.direction ?? source.angle ?? 0);
  const opacity = Number(source.opacity ?? 0.2);
  if (![blurRadius, distance, direction, opacity].every(Number.isFinite) || blurRadius < 0 || distance < 0 || opacity < 0 || opacity > 1) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} has an invalid shadow.`, [], { code: "invalid_presentation_shadow" });
  }
  const normalizedDirection = ((direction % 360) + 360) % 360;
  return {
    colorRgb: presentationRgb(source.color || source.fill || "#000000", `${shapeId}.shadow.color`),
    blurRadiusEmu: emuFromPixels(blurRadius, `${shapeId}.shadow.blurRadius`),
    distanceEmu: emuFromPixels(distance, `${shapeId}.shadow.distance`),
    directionAngle60000: Math.round(normalizedDirection * ROTATION_UNITS_PER_DEGREE),
    opacityThousandthPercent: Math.round(opacity * 100_000),
  };
}

function modelPresentationShadow(shadow) {
  if (!shadow) return undefined;
  return {
    color: shadow.colorRgb ? `#${shadow.colorRgb}` : "#000000",
    blurRadius: Number(shadow.blurRadiusEmu) / EMU_PER_PIXEL,
    distance: Number(shadow.distanceEmu) / EMU_PER_PIXEL,
    direction: Number(shadow.directionAngle60000) / ROTATION_UNITS_PER_DEGREE,
    opacity: Number(shadow.opacityThousandthPercent) / 100_000,
  };
}

function sourceBoundCloneConnectorTargetId(value, sourceIdByCloneId, connector, side) {
  const targetId = String(value || "");
  if (!targetId || !sourceIdByCloneId) return targetId;
  const sourceId = sourceIdByCloneId.get(targetId);
  if (!sourceId) {
    throw new OpenChestnutCodecError(`Imported presentation clone connector ${connector.id} has an unresolved ${side} target.`, [], { code: "unsupported_presentation_slide_clone" });
  }
  return sourceId;
}

function presentationConnector(connector, original, sourceIdByCloneId) {
  const type = String(connector.connectorType || "straight");
  if (!new Set(["straight", "elbow"]).has(type)) {
    throw new OpenChestnutCodecError(`Presentation connector ${connector.id} uses unsupported type ${type}.`, [], { code: "unsupported_presentation_features" });
  }
  const line = connector.line || {};
  if (line.style != null && !new Set(["solid", "none"]).has(String(line.style))) {
    throw new OpenChestnutCodecError(`Presentation connector ${connector.id} uses an unsupported line style.`, [], { code: "unsupported_presentation_features" });
  }
  const width = Number(line.width ?? 2);
  if (!Number.isFinite(width) || width < 0) throw new OpenChestnutCodecError(`Presentation connector ${connector.id} has an invalid line width.`, [], { code: "invalid_presentation_connector" });
  const arrow = (value, name) => {
    if (value == null || value === false || value === "none") return "";
    if (value === true || value === "triangle") return "triangle";
    throw new OpenChestnutCodecError(`Presentation connector ${connector.id} uses unsupported ${name} ${value}.`, [], { code: "unsupported_presentation_features" });
  };
  return {
    id: original?.id || connector.id,
    name: connector.name || original?.name || "",
    source: original?.source,
    content: {
      case: "connector",
      value: {
        connectorType: type,
        startXEmu: sourceBoundFrameEmuFromPixels(connector.start?.x, `${connector.id}.start.x`, original),
        startYEmu: sourceBoundFrameEmuFromPixels(connector.start?.y, `${connector.id}.start.y`, original),
        endXEmu: sourceBoundFrameEmuFromPixels(connector.end?.x, `${connector.id}.end.x`, original),
        endYEmu: sourceBoundFrameEmuFromPixels(connector.end?.y, `${connector.id}.end.y`, original),
        lineRgb: presentationRgb(line.fill || line.color || (width > 0 ? "#334155" : "transparent"), `${connector.id}.line.fill`),
        lineWidthEmu: BigInt(Math.round(width * EMU_PER_POINT)),
        startArrow: arrow(line.startArrow ?? connector.startArrow, "start arrow"),
        endArrow: arrow(line.endArrow ?? connector.endArrow, "end arrow"),
        startTargetId: sourceBoundCloneConnectorTargetId(connector.startTargetId, sourceIdByCloneId, connector, "start"),
        endTargetId: sourceBoundCloneConnectorTargetId(connector.endTargetId, sourceIdByCloneId, connector, "end"),
      },
    },
  };
}

function presentationShape(shape, original, assetCatalog, customShowLinks) {
  const originalShape = original?.content?.case === "shape" ? original.content.value : original;
  if (!new Set(["rect", "ellipse", "roundRect", "textbox", "custom"]).has(shape.geometry)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shape.id} uses unsupported geometry ${shape.geometry}.`, [], { code: "unsupported_presentation_features" });
  }
  if (shape.geometry !== "custom" && shape.customPaths?.length) {
    throw new OpenChestnutCodecError(`Presentation shape ${shape.id} has custom paths without custom geometry.`, [], { code: "invalid_presentation_geometry" });
  }
  const customPaths = (shape.customPaths || []).map((path) => ({
    width: BigInt(path.width),
    height: BigInt(path.height),
    commands: path.commands.map((command) => {
      if (command.moveTo) return { command: { case: "moveTo", value: { x: BigInt(command.moveTo.x), y: BigInt(command.moveTo.y) } } };
      if (command.lineTo) return { command: { case: "lineTo", value: { x: BigInt(command.lineTo.x), y: BigInt(command.lineTo.y) } } };
      if (command.cubicBezTo) return {
        command: {
          case: "cubicBezierTo",
          value: {
            control1: { x: BigInt(command.cubicBezTo.x1), y: BigInt(command.cubicBezTo.y1) },
            control2: { x: BigInt(command.cubicBezTo.x2), y: BigInt(command.cubicBezTo.y2) },
            end: { x: BigInt(command.cubicBezTo.x), y: BigInt(command.cubicBezTo.y) },
          },
        },
      };
      return { command: { case: "close", value: true } };
    }),
  }));
  // The model deliberately withholds an unrecognized custom-path grammar.
  // An unchanged source-bound, non-editable shape can still be carried by the
  // C# codec, which rechecks its source binding and rejects every mutation.
  // Source-free or editable shapes must continue to provide the full grammar.
  const opaqueSourceBoundCustomGeometry = original?.source?.editable === false;
  if (shape.geometry === "custom" && customPaths.length === 0 && !opaqueSourceBoundCustomGeometry) {
    throw new OpenChestnutCodecError(`Presentation shape ${shape.id} requires custom paths.`, [], { code: "invalid_presentation_geometry" });
  }
  const position = shape.position || {};
  const lineWidth = Number(shape.line?.width ?? 1);
  if (!Number.isFinite(lineWidth) || lineWidth < 0) throw new OpenChestnutCodecError(`Presentation shape ${shape.id} has an invalid line width.`, [], { code: "invalid_presentation_frame" });
  const placeholder = !original && shape.placeholder ? sourceFreeSlidePlaceholder(shape) : undefined;
  const textBody = presentationTextBody(shape, originalShape, assetCatalog, customShowLinks);
  const shadow = presentationShadow(shape.shadow, shape.id);
  return {
    id: original?.id || shape.id,
    name: shape.name || original?.name || "",
    source: original?.source,
    content: {
      case: "shape",
      value: {
        geometry: shape.geometry,
        leftEmu: sourceBoundFrameEmuFromPixels(position.left, `${shape.id}.position.left`, original),
        topEmu: sourceBoundFrameEmuFromPixels(position.top, `${shape.id}.position.top`, original),
        widthEmu: emuFromPixels(position.width, `${shape.id}.position.width`),
        heightEmu: emuFromPixels(position.height, `${shape.id}.position.height`),
        text: shape.text?.value || "",
        textBody,
        fillRgb: presentationRgb(shape.fill, `${shape.id}.fill`),
        lineRgb: presentationRgb(shape.line?.fill || shape.line?.color || (lineWidth > 0 ? "#334155" : "transparent"), `${shape.id}.line.fill`),
        lineWidthEmu: BigInt(Math.round(lineWidth * EMU_PER_POINT)),
        ...(placeholder || {}),
        ...(placeholder || shape.transform == null ? {} : { transform: wirePresentationTransform(shape.transform, `shape ${shape.id}`) }),
        ...(shadow ? { shadow } : {}),
        ...(customPaths.length ? { customPaths } : {}),
        ...(shape.useBackgroundFill === undefined ? {} : { useBackgroundFill: shape.useBackgroundFill }),
      },
    },
  };
}

function presentationImage(image, original, assetCatalog) {
  const position = image.position || {};
  if (!image.dataUrl) {
    throw new OpenChestnutCodecError(`Presentation image ${image.id} requires an embedded dataUrl.`, [], { code: "invalid_presentation_image" });
  }
  if (image.uri || image.geometry !== "rect" || image.borderRadius != null) {
    throw new OpenChestnutCodecError(`Presentation image ${image.id} uses external, geometry, or mask semantics outside the bounded PPTX image slice.`, [], { code: "unsupported_presentation_features" });
  }
  const crop = effectivePresentationImageCrop({ crop: image.crop, fit: image.fit, dataUrl: image.dataUrl, frame: position });
  return {
    id: original?.id || image.id,
    name: image.name || original?.name || "",
    source: original?.source,
    content: {
      case: "image",
      value: {
        assetId: assetCatalog.addDataUrl(image.dataUrl),
        altText: image.alt || image.prompt || "",
        leftEmu: sourceBoundFrameEmuFromPixels(position.left, `${image.id}.position.left`, original),
        topEmu: sourceBoundFrameEmuFromPixels(position.top, `${image.id}.position.top`, original),
        widthEmu: emuFromPixels(position.width, `${image.id}.position.width`),
        heightEmu: emuFromPixels(position.height, `${image.id}.position.height`),
        ...(crop ? { crop: presentationImageCropToWire(crop) } : {}),
        ...(image.transform == null ? {} : { transform: wirePresentationTransform(image.transform, `image ${image.id}`) }),
      },
    },
  };
}

function presentationImageReadOnlySnapshot(image) {
  return JSON.stringify({
    uri: image.uri,
    contentType: image.contentType,
    geometry: image.geometry,
    borderRadius: image.borderRadius,
  });
}

function distributePresentationTableSize(total, count, ownerLabel) {
  const slots = Number(count);
  if (!Number.isInteger(slots) || slots < 1) {
    throw new OpenChestnutCodecError(`Presentation table ${ownerLabel} must contain at least one row and column.`, [], { code: "invalid_presentation_table" });
  }
  const base = total / BigInt(slots);
  const remainder = Number(total % BigInt(slots));
  if (base < 1n) throw new OpenChestnutCodecError(`Presentation table ${ownerLabel} is too small for its grid.`, [], { code: "invalid_presentation_table" });
  return Array.from({ length: slots }, (_, index) => base + (index < remainder ? 1n : 0n));
}

function scalePresentationTableSize(values, total, ownerLabel) {
  const source = values.map((value) => BigInt(value));
  const sourceTotal = source.reduce((sum, value) => sum + value, 0n);
  if (!source.length || sourceTotal < 1n) return distributePresentationTableSize(total, source.length, ownerLabel);
  const scaled = source.map((value) => ({ value: (value * total) / sourceTotal, remainder: (value * total) % sourceTotal }));
  let missing = total - scaled.reduce((sum, item) => sum + item.value, 0n);
  for (const index of scaled.map((item, index) => ({ index, remainder: item.remainder })).sort((left, right) => left.remainder === right.remainder ? left.index - right.index : left.remainder > right.remainder ? -1 : 1).map((item) => item.index)) {
    if (missing <= 0n) break;
    scaled[index].value += 1n;
    missing -= 1n;
  }
  if (scaled.some((item) => item.value < 1n)) {
    throw new OpenChestnutCodecError(`Presentation table ${ownerLabel} is too small for its imported grid.`, [], { code: "invalid_presentation_table" });
  }
  return scaled.map((item) => item.value);
}

function presentationTable(table, original) {
  const originalTable = original?.content?.case === "table" ? original.content.value : undefined;
  const rows = Number(table.rows);
  const columns = Number(table.columns);
  if (!Number.isInteger(rows) || rows < 1 || rows > 2048 || !Number.isInteger(columns) || columns < 1 || columns > 256 ||
      table.values.length !== rows || table.values.some((row) => !Array.isArray(row) || row.length !== columns)) {
    throw new OpenChestnutCodecError(`Presentation table ${table.id} requires a rectangular 1-2048 by 1-256 value matrix.`, [], { code: "invalid_presentation_table" });
  }
  const position = table.position || {};
  const leftEmu = sourceBoundFrameEmuFromPixels(position.left, `${table.id}.position.left`, original);
  const topEmu = sourceBoundFrameEmuFromPixels(position.top, `${table.id}.position.top`, original);
  const widthEmu = emuFromPixels(position.width, `${table.id}.position.width`);
  const heightEmu = emuFromPixels(position.height, `${table.id}.position.height`);
  if (widthEmu < 1n || heightEmu < 1n) throw new OpenChestnutCodecError(`Presentation table ${table.id} requires positive width and height.`, [], { code: "invalid_presentation_table" });
  if (originalTable && (originalTable.rows.length !== rows || originalTable.columnWidthsEmu.length !== columns)) {
    throw new OpenChestnutCodecError(`Source-preserving PPTX export requires presentation table ${table.id}'s original fixed topology.`, [], { code: "presentation_table_topology_changed" });
  }
  const columnWidthsEmu = originalTable
    ? scalePresentationTableSize(originalTable.columnWidthsEmu, widthEmu, `${table.id} columns`)
    : distributePresentationTableSize(widthEmu, columns, `${table.id} columns`);
  const rowHeightsEmu = originalTable
    ? scalePresentationTableSize(originalTable.rows.map((row) => row.heightEmu), heightEmu, `${table.id} rows`)
    : distributePresentationTableSize(heightEmu, rows, `${table.id} rows`);
  return {
    id: original?.id || table.id,
    name: String(table.name || original?.name || ""),
    source: original?.source,
    content: {
      case: "table",
      value: {
        leftEmu,
        topEmu,
        widthEmu,
        heightEmu,
        columnWidthsEmu,
        rows: table.values.map((row, rowIndex) => ({
          heightEmu: rowHeightsEmu[rowIndex],
          cells: row.map((value) => ({ text: String(value ?? "") })),
        })),
        mergeRanges: table.mergeRanges.map((range) => ({
          startRow: range.startRow,
          endRow: range.endRow,
          startColumn: range.startColumn,
          endColumn: range.endColumn,
        })),
        ...(originalTable?.firstRow === undefined ? { firstRow: Boolean(table.styleOptions?.headerRow) } : { firstRow: originalTable.firstRow }),
        ...(originalTable?.bandedRows === undefined ? { bandedRows: Boolean(table.styleOptions?.bandedRows) } : { bandedRows: originalTable.bandedRows }),
      },
    },
  };
}

function presentationTableReadOnlySnapshot(table) {
  return JSON.stringify({
    id: table.id,
    nativeId: table.nativeId,
    creationId: table.creationId,
    rows: table.rows,
    columns: table.columns,
    style: table.style,
    styleOptions: table.styleOptions,
    border: table.border,
    mergeRanges: table.mergeRanges,
  });
}

function presentationElement(element, original, assetCatalog, sourceIdByCloneId, customShowLinks) {
  if (element instanceof GroupShape) return presentationGroup(element, original, assetCatalog, sourceIdByCloneId, customShowLinks);
  if (element instanceof ImageElement) return presentationImage(element, original, assetCatalog);
  if (element instanceof TableElement) return presentationTable(element, original);
  if (element instanceof ChartElement) return presentationChart(element, original);
  if (element?.kind === "connector") return presentationConnector(element, original, sourceIdByCloneId);
  if (element instanceof Shape) return presentationShape(element, original, assetCatalog, customShowLinks);
  throw new OpenChestnutCodecError(`Presentation element ${element?.id || "<unknown>"} has no supported OpenChestnut wire projection.`, [], { code: "unsupported_presentation_element" });
}

function presentationGroup(group, original, assetCatalog, sourceIdByCloneId, customShowLinks) {
  const originalGroup = original?.content?.case === "group" ? original.content.value : undefined;
  if (!group.children.length) throw new OpenChestnutCodecError(`Presentation group ${group.id} requires at least one child.`, [], { code: "invalid_presentation_group" });
  if (originalGroup && originalGroup.children.length !== group.children.length) {
    throw new OpenChestnutCodecError(`Source-preserving PPTX export requires presentation group ${group.id}'s original ${originalGroup.children.length}-child topology.`, [], { code: "presentation_group_topology_changed" });
  }
  const frame = group.position || {};
  const childFrame = group.childFrame || {};
  const widthEmu = emuFromPixels(frame.width, `${group.id}.position.width`);
  const heightEmu = emuFromPixels(frame.height, `${group.id}.position.height`);
  const childWidthEmu = emuFromPixels(childFrame.width, `${group.id}.childFrame.width`);
  const childHeightEmu = emuFromPixels(childFrame.height, `${group.id}.childFrame.height`);
  if (widthEmu < 1n || heightEmu < 1n || childWidthEmu < 1n || childHeightEmu < 1n) {
    throw new OpenChestnutCodecError(`Presentation group ${group.id} requires positive outer and child extents.`, [], { code: "invalid_presentation_group" });
  }
  return {
    id: original?.id || group.id,
    name: String(group.name || original?.name || ""),
    source: original?.source,
    content: {
      case: "group",
      value: {
        leftEmu: sourceBoundFrameEmuFromPixels(frame.left, `${group.id}.position.left`, original),
        topEmu: sourceBoundFrameEmuFromPixels(frame.top, `${group.id}.position.top`, original),
        widthEmu,
        heightEmu,
        childLeftEmu: signedEmuFromPixels(childFrame.left, `${group.id}.childFrame.left`),
        childTopEmu: signedEmuFromPixels(childFrame.top, `${group.id}.childFrame.top`),
        childWidthEmu,
        childHeightEmu,
        children: group.children.map((child, index) => presentationElement(child, originalGroup?.children[index], assetCatalog, sourceIdByCloneId, customShowLinks)),
      },
    },
  };
}

function directSlideElements(slide) {
  return [
    ...slide.shapes.items,
    ...slide.tables.items,
    ...slide.charts.items,
    ...slide.images.items,
    ...slide.connectors.items,
    ...slide.groups.items,
    ...slide.nativeObjects.items,
  ];
}

function legacyCommentCoordinate(value, unit, name) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    throw new OpenChestnutCodecError(`${name} must be a finite coordinate.`, [], { code: "invalid_presentation_legacy_comment" });
  }
  if (unit === "emu") return Math.round(number);
  if (unit === undefined || unit === "px") return emuFromPixels(number, name);
  throw new OpenChestnutCodecError(`${name}.unit must be "px" or "emu".`, [], { code: "invalid_presentation_legacy_comment" });
}

function legacyCommentTimestamp(value, name) {
  const text = String(value ?? "");
  if (!text || Number.isNaN(Date.parse(text))) {
    throw new OpenChestnutCodecError(`${name} must be an ISO-8601 timestamp.`, [], { code: "invalid_presentation_legacy_comment" });
  }
  return text;
}

function legacyCommentInteger(value) {
  const number = Number(value);
  return Number.isSafeInteger(number) ? number : undefined;
}

// Keep native legacy-comment evidence only on the imported model. It is used
// to prove an unchanged source-bound export, never exposed as cross-file
// identity and never used to turn an element/thread API into a fake native
// anchor. New legacy comments are slide-level annotations at a fixed position.
function presentationLegacyComments(slide, slideIndex) {
  return slide.comments.items.map((thread, index) => {
    const label = `slide ${slideIndex + 1} legacy comment ${index + 1}`;
    if (thread.nativeFormat && thread.nativeFormat !== "legacy") {
      throw new OpenChestnutCodecError(`${label} uses ${thread.nativeFormat} comments, which are outside the legacy PPTX profile.`, [], { code: "unsupported_presentation_comment" });
    }
    if (thread.targetId) {
      throw new OpenChestnutCodecError(`${label} targets an element or text range. Legacy PPTX comments are slide-level only.`, [], { code: "unsupported_presentation_comment" });
    }
    if (thread.resolved) {
      throw new OpenChestnutCodecError(`${label} is resolved. Legacy PPTX comments do not encode thread state.`, [], { code: "unsupported_presentation_comment" });
    }
    if (!Array.isArray(thread.comments) || thread.comments.length !== 1) {
      throw new OpenChestnutCodecError(`${label} must contain exactly one root comment and no replies.`, [], { code: "unsupported_presentation_comment" });
    }
    const comment = thread.comments[0];
    const author = String(comment.author ?? thread.author ?? "").trim();
    if (!author) {
      throw new OpenChestnutCodecError(`${label} requires a non-empty author.`, [], { code: "invalid_presentation_legacy_comment" });
    }
    const position = thread.position;
    if (!position || typeof position !== "object") {
      throw new OpenChestnutCodecError(`${label} requires an explicit { x, y, unit? } position.`, [], { code: "invalid_presentation_legacy_comment" });
    }
    const anchor = thread.nativeFormat === "legacy" && thread.nativeAnchor && typeof thread.nativeAnchor === "object"
      ? thread.nativeAnchor
      : undefined;
    const anchorPositionXEmu = legacyCommentInteger(anchor?.positionXEmu);
    const anchorPositionYEmu = legacyCommentInteger(anchor?.positionYEmu);
    const positionXEmu = anchorPositionXEmu !== undefined
      ? anchorPositionXEmu
      : legacyCommentCoordinate(position.x, position.unit, `${label}.position.x`);
    const positionYEmu = anchorPositionYEmu !== undefined
      ? anchorPositionYEmu
      : legacyCommentCoordinate(position.y, position.unit, `${label}.position.y`);
    const result = {
      id: thread.id,
      author,
      text: String(comment.text ?? ""),
      createdAt: legacyCommentTimestamp(comment.created ?? thread.created, `${label}.created`),
      positionXEmu,
      positionYEmu,
    };
    const nativeAuthorId = legacyCommentInteger(anchor?.nativeAuthorId);
    const nativeIndex = legacyCommentInteger(anchor?.nativeIndex);
    if (nativeAuthorId !== undefined && nativeAuthorId >= 0) result.nativeAuthorId = nativeAuthorId;
    if (nativeIndex !== undefined && nativeIndex >= 0) result.nativeIndex = nativeIndex;
    return result;
  });
}

const PRESENTATION_MODERN_COMMENT_STATUSES = new Set(["active", "resolved", "closed"]);
const PRESENTATION_MODERN_COMMENT_GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;

function modernCommentGuid(value, seed, label) {
  const guid = String(value || deterministicPresentationGuid(seed)).toUpperCase();
  if (!PRESENTATION_MODERN_COMMENT_GUID.test(guid)) {
    throw new OpenChestnutCodecError(`${label} must be a brace-delimited GUID.`, [], { code: "invalid_presentation_modern_comment" });
  }
  return guid;
}

function modernCommentTimestamp(value, label) {
  const timestamp = String(value ?? "");
  if (!timestamp || Number.isNaN(Date.parse(timestamp))) {
    throw new OpenChestnutCodecError(`${label} must be an ISO-8601 timestamp.`, [], { code: "invalid_presentation_modern_comment" });
  }
  return timestamp;
}

function modernCommentInitials(name) {
  const words = String(name || "User").trim().split(/\s+/).filter(Boolean);
  return (words.length > 1 ? words.slice(0, 2).map((word) => [...word][0]) : [...(words[0] || "U")].slice(0, 2)).join("").toUpperCase();
}

function modernCommentAuthor(comment, thread) {
  const person = comment.person || {};
  const author = String(comment.author || person.name || person.displayName || thread.author || "").trim();
  if (!author) throw new OpenChestnutCodecError(`Modern comment ${thread.id} requires a non-empty author.`, [], { code: "invalid_presentation_modern_comment" });
  return {
    authorId: modernCommentGuid(comment.authorId || person.id, `author:${author}`, `Modern comment author ${author}`),
    author,
    initials: String(person.initials || comment.initials || modernCommentInitials(author)),
    userId: String(person.userId ?? comment.userId ?? author),
    providerId: String(person.providerId ?? comment.providerId ?? "None"),
  };
}

function flattenedPresentationWireElements(elements) {
  const output = [];
  const visit = (element) => {
    output.push(element);
    if (element.content?.case === "group") for (const child of element.content.value.children || []) visit(child);
  };
  for (const element of elements) visit(element);
  return output;
}

function modernCommentMoniker(wireElement) {
  return {
    shape: "spMk",
    image: "picMk",
    table: "graphicFrameMk",
    chart: "graphicFrameMk",
    connector: "cxnSpMk",
    group: "grpSpMk",
  }[wireElement?.content?.case];
}

function modernCommentCoordinate(value, unit, label) {
  const number = Number(value ?? 0);
  if (!Number.isFinite(number)) throw new OpenChestnutCodecError(`${label} must be finite.`, [], { code: "invalid_presentation_modern_comment" });
  if (unit === "px") return emuFromPixels(number, label);
  if (unit === undefined || unit === "emu") return Math.round(number);
  throw new OpenChestnutCodecError(`${label}.unit must be "emu" or "px".`, [], { code: "invalid_presentation_modern_comment" });
}

function presentationModernComments(slide, slideIndex, wireElements, originalThreads = []) {
  const flattened = flattenedPresentationWireElements(wireElements);
  const wireById = new Map(flattened.map((element, index) => [element.id, { element, nativeId: index + 2 }]));
  const directIds = new Set(wireElements.map((element) => element.id));
  return slide.comments.items.map((thread, threadIndex) => {
    const label = `slide ${slideIndex + 1} modern comment thread ${threadIndex + 1}`;
    if (thread.nativeFormat && thread.nativeFormat !== "modern") {
      throw new OpenChestnutCodecError(`${label} uses ${thread.nativeFormat} comments.`, [], { code: "unsupported_presentation_comment" });
    }
    if (!Array.isArray(thread.comments) || thread.comments.length === 0) {
      throw new OpenChestnutCodecError(`${label} requires one root comment.`, [], { code: "invalid_presentation_modern_comment" });
    }
    const target = slide.resolve(thread.targetId);
    const textRange = target?.kind === "textRange";
    const targetElementId = textRange ? target.parentId : target?.id;
    const targetWire = wireById.get(targetElementId);
    if (!targetWire || !directIds.has(targetElementId)) {
      throw new OpenChestnutCodecError(`${label} must target a supported top-level slide element or its text range.`, [], { code: "unsupported_presentation_comment" });
    }
    const sourceAnchor = thread.nativeAnchor?.format === "modern" || thread.nativeAnchor?.type ? thread.nativeAnchor : undefined;
    const monikerType = sourceAnchor?.moniker || modernCommentMoniker(targetWire.element);
    if (!monikerType || (textRange && monikerType !== "spMk")) {
      throw new OpenChestnutCodecError(`${label} has an unsupported target moniker.`, [], { code: "unsupported_presentation_comment" });
    }
    const nativeSlideId = Number(sourceAnchor?.nativeSlideId ?? sourceAnchor?.slideId ?? 256 + slideIndex);
    const nativeId = Number(sourceAnchor?.nativeId ?? targetWire.nativeId);
    const anchor = {
      kind: textRange ? PresentationModernCommentAnchor_Kind.TEXT_RANGE : PresentationModernCommentAnchor_Kind.ELEMENT,
      nativeSlideId,
      monikers: [{
        type: monikerType,
        nativeId,
        ...(sourceAnchor?.creationId ? { creationId: String(sourceAnchor.creationId).toUpperCase() } : {}),
      }],
      ...(textRange ? {
        textStart: Number(sourceAnchor?.textStart ?? sourceAnchor?.cp ?? 0),
        textLength: Number(sourceAnchor?.textLength ?? sourceAnchor?.length ?? String(target.text ?? "").length),
        ...(sourceAnchor?.contextLength === undefined ? {} : { contextLength: Number(sourceAnchor.contextLength) }),
        ...(sourceAnchor?.contextHash === undefined ? {} : { contextHash: Number(sourceAnchor.contextHash) }),
      } : {}),
    };
    const comments = thread.comments.map((comment, commentIndex) => {
      const author = modernCommentAuthor(comment, thread);
      const status = String(comment.status || (commentIndex === 0 && thread.resolved ? "resolved" : "active")).toLowerCase();
      if (!PRESENTATION_MODERN_COMMENT_STATUSES.has(status)) {
        throw new OpenChestnutCodecError(`${label} comment ${commentIndex + 1} has invalid status ${status}.`, [], { code: "invalid_presentation_modern_comment" });
      }
      return {
        id: modernCommentGuid(comment.nativeId || comment.id, `comment:${thread.id}:${commentIndex}`, `${label} comment ${commentIndex + 1}`),
        ...author,
        text: String(comment.text ?? ""),
        createdAt: modernCommentTimestamp(comment.created || thread.created, `${label} comment ${commentIndex + 1}.created`),
        status,
      };
    });
    const original = originalThreads[threadIndex];
    const position = thread.position;
    if (!position || typeof position !== "object") {
      throw new OpenChestnutCodecError(`${label} requires an explicit { x, y, unit? } position.`, [], { code: "invalid_presentation_modern_comment" });
    }
    return {
      id: comments[0].id,
      targetId: thread.targetId,
      anchor,
      positionXEmu: modernCommentCoordinate(position.x, position.unit, `${label}.position.x`),
      positionYEmu: modernCommentCoordinate(position.y, position.unit, `${label}.position.y`),
      root: comments[0],
      replies: comments.slice(1),
      ...(original?.source ? { source: original.source } : {}),
    };
  });
}

function presentationThemeSnapshot(theme) {
  return JSON.stringify({
    name: theme.name,
    colors: theme.colors,
    fonts: theme.fonts,
    textStyles: theme.textStyles,
    colorMap: theme.colorMap,
  });
}

function presentationAdvancedSnapshot(presentation) {
  return JSON.stringify({
    theme: JSON.parse(presentationThemeSnapshot(presentation.theme)),
    commentFormat: presentation.commentFormat,
  });
}

function presentationCustomShowLinkContext(currentShows, state) {
  const byId = new Map();
  const byName = new Map();
  for (const show of currentShows) {
    if (byId.has(show.id) || byName.has(show.name)) {
      throw new OpenChestnutCodecError("Presentation custom-show hyperlink targets require unique show IDs and names.", [], { code: "invalid_presentation_custom_show" });
    }
    byId.set(show.id, show);
    byName.set(show.name, show);
  }
  return {
    byId,
    byName,
    originalNameById: new Map((state?.customShows || []).map((entry) => [entry.wire.id, entry.wire.name])),
  };
}

function resolvePresentationCustomShowLinkId(name, originalRun, shapeId, context) {
  const originalTarget = originalRun?.hyperlink?.case === "runHyperlink"
    ? originalRun.hyperlink.value?.target
    : undefined;
  let show;
  if (originalTarget?.case === "customShowId" && context?.originalNameById.get(originalTarget.value) === name) {
    // A show rename does not implicitly retarget every referring run. The
    // public model still carries the imported display name, while this stable
    // wire identity follows the same native show across that rename.
    show = context.byId.get(originalTarget.value);
  }
  show ||= context?.byName.get(name);
  if (!show) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} references missing custom show ${name}.`, [], { code: "invalid_presentation_hyperlink" });
  }
  return show.id;
}

function presentationCustomShows(presentation, state) {
  const entries = planPresentationCustomShows(presentation).entries;
  if (!state) return entries.map((show) => ({
    id: show.id,
    name: show.name,
    nativeId: show.nativeId,
    slideIds: [...show.slideIds],
  }));
  if (state.customShowsOpaque) {
    if (entries.length) {
      throw new OpenChestnutCodecError("The imported PPTX contains an opaque custom-show graph; it can only be preserved unchanged.", [], { code: "unsupported_presentation_custom_show_edit" });
    }
    return [];
  }
  const sourceEntries = state.customShows || [];
  if (entries.length !== sourceEntries.length || entries.some((show, index) => show !== sourceEntries[index].model)) {
    throw new OpenChestnutCodecError("Imported PPTX custom shows keep their original count and order; adding, removing, or reordering shows is unsupported.", [], { code: "presentation_custom_show_topology_changed" });
  }
  return entries.map((show, index) => {
    const sourceEntry = sourceEntries[index];
    if (show.id !== sourceEntry.wire.id || show.nativeId !== sourceEntry.wire.nativeId) {
      throw new OpenChestnutCodecError(`Imported PPTX custom show ${index + 1} cannot change its facade or native identity.`, [], { code: "presentation_custom_show_topology_changed" });
    }
    return {
      id: sourceEntry.wire.id,
      name: show.name,
      nativeId: show.nativeId,
      slideIds: [...show.slideIds],
      source: sourceEntry.wire.source,
    };
  });
}

function presentationSections(presentation, state) {
  if (state?.sectionsOpaque) {
    if (presentation.sections.items.length) {
      throw new OpenChestnutCodecError("The imported PPTX contains an opaque PowerPoint section graph; it can only be preserved unchanged.", [], { code: "unsupported_presentation_section_edit" });
    }
    return [];
  }
  const sourceEntries = state?.sections || [];
  if (state && (presentation.sections.items.length !== sourceEntries.length || presentation.sections.items.some((section, index) => section !== sourceEntries[index].model))) {
    throw new OpenChestnutCodecError("Imported PPTX sections keep their original count and order; adding, removing, or reordering sections is unsupported.", [], { code: "presentation_section_topology_changed" });
  }
  const entries = planPresentationSections(presentation, { allowPendingClone: Boolean(state?.clones?.length) }).entries;
  if (!state) return entries.map((section) => ({
    id: section.id,
    name: section.name,
    nativeId: section.nativeId,
    slideIds: [...section.slideIds],
  }));
  return entries.map((section, index) => {
    const sourceEntry = sourceEntries[index];
    if (section.id !== sourceEntry.wire.id || section.nativeId !== sourceEntry.wire.nativeId) {
      throw new OpenChestnutCodecError(`Imported PPTX section ${index + 1} cannot change its facade or native GUID identity.`, [], { code: "presentation_section_topology_changed" });
    }
    return {
      id: sourceEntry.wire.id,
      name: section.name,
      nativeId: section.nativeId,
      slideIds: [...section.slideIds],
      source: sourceEntry.wire.source,
    };
  });
}

// Imported comment state belongs to its source SlidePart, not to its current
// display index. Keeping the snapshot per source-state lets a valid deletion
// omit that state while every surviving slide remains strictly read-only.
function presentationSlideCommentSnapshot(slide) {
  return JSON.stringify(slide.comments.items.map((comment) => comment.toJSON()));
}

function unsupportedPresentationFeatures(presentation) {
  const unsupported = [];
  if (presentationThemeSnapshot(presentation.theme) !== DEFAULT_PRESENTATION_THEME) unsupported.push("presentation theme customization");
  if (presentation.masters?.items?.length !== 1) unsupported.push("multiple slide masters");
  const master = presentation.master;
  if (master?.theme) unsupported.push("master theme override");
  if (!["legacy", "modern"].includes(presentation.commentFormat)) unsupported.push(`unknown comment format ${presentation.commentFormat}`);
  for (const slide of presentation.slides?.items || []) {
    const prefix = `slide ${slide.index + 1}`;
    if (presentation.commentFormat === "legacy" && slide.comments?.items?.length) {
      try { presentationLegacyComments(slide, slide.index); }
      catch (error) { unsupported.push(`${prefix} comments (${error.message})`); }
    }
    if (slide.nativeObjects?.items?.length) unsupported.push(`${prefix} native objects`);
  }
  return unsupported;
}

function opaquePresentationSnapshot(object) {
  const oleWorkbook = object.oleWorkbook ? {
    partPath: object.oleWorkbook.partPath,
    contentType: object.oleWorkbook.contentType,
    sourceSha256: object.oleWorkbook.sourceSha256,
    relationshipId: object.oleWorkbook.relationshipId,
  } : undefined;
  return JSON.stringify({
    id: object.id,
    name: object.name,
    position: object.position,
    nativeKind: object.nativeKind,
    rawXml: object.rawXml,
    oleWorkbook,
    diagramText: object._diagramTextSourceBinding?.(),
    ...presentationNativeGraphSnapshot(object),
  });
}

function presentationOpaque(object, original, snapshot, assetCatalog) {
  if (opaquePresentationSnapshot(object) !== snapshot) {
    const message = object.oleWorkbook
      ? `Presentation native element ${object.id} changed outside its bounded embedded-workbook replacement boundary.`
      : `Presentation native element ${object.id} is source-bound and read-only in OpenChestnut 0.2.`;
    throw new OpenChestnutCodecError(message, [], { code: "unsupported_presentation_edit" });
  }
  const replacement = object._embeddedWorkbookReplacementBytes?.();
  const diagramTextReplacement = object._diagramTextReplacement?.();
  if (!replacement && !diagramTextReplacement) return original;
  const originalOpaque = original?.content?.case === "opaque" ? original.content.value : undefined;
  if (!originalOpaque) {
    throw new OpenChestnutCodecError(`Presentation native element ${object.id} has no source-bound opaque payload.`, [], { code: "unsupported_presentation_edit" });
  }
  if (replacement && (!object.oleWorkbook || !originalOpaque.oleWorkbook)) {
    throw new OpenChestnutCodecError(`Presentation native element ${object.id} has no source-bound embedded XLSX workbook.`, [], { code: "unsupported_presentation_edit" });
  }
  if (diagramTextReplacement && !originalOpaque.diagramText) {
    throw new OpenChestnutCodecError(`Presentation native element ${object.id} has no source-bound SmartArt diagram-text binding.`, [], { code: "unsupported_presentation_edit" });
  }
  return {
    ...original,
    content: {
      case: "opaque",
      value: {
        ...originalOpaque,
        ...(replacement ? { oleWorkbook: {
          ...originalOpaque.oleWorkbook,
          replacementAssetId: assetCatalog.addOleWorkbook(replacement),
        } } : {}),
        ...(diagramTextReplacement ? { diagramText: {
          ...originalOpaque.diagramText,
          nodes: diagramTextReplacement.nodes.map((node) => create(PresentationDiagramTextNodeSchema, { modelId: node.id, text: node.text })),
        } } : {}),
      },
    },
  };
}

export function presentationEnvelope(presentation, protocolVersion) {
  if (!(presentation instanceof Presentation)) throw new TypeError("exportPptxWithOpenChestnut expects a Presentation instance.");
  if (!presentation.slides?.items?.length) throw new OpenChestnutCodecError("Presentation must contain at least one slide.", [], { code: "missing_slides" });
  const state = presentation[PRESENTATION_STATE];
  assertTrustedPresentationState(state);
  const sourceStates = presentationSourceSlideStateMap(presentation, state);
  if (!state) {
    const unsupported = unsupportedPresentationFeatures(presentation);
    if (unsupported.length) {
      throw new OpenChestnutCodecError(`OpenChestnut cannot author these source-free PPTX features: ${unsupported.slice(0, 8).join(", ")}${unsupported.length > 8 ? `, and ${unsupported.length - 8} more` : ""}. Export fails closed; use supported features or import a trustworthy source package for opaque preservation.`, [], { code: "unsupported_presentation_features" });
    }
  } else {
    if (presentationAdvancedSnapshot(presentation) !== state.advancedSnapshot) {
      throw new OpenChestnutCodecError("Imported presentation theme and comment wire family are source-bound and read-only in OpenChestnut 0.2.", [], { code: "unsupported_presentation_edit" });
    }
    // A source-bound canvas resize is explicit and intentionally narrow: the
    // native codec changes only p:presentation/p:sldSz. It never treats a
    // changed canvas as permission to rescale every slide/master coordinate.
  }

  const customShows = presentationCustomShows(presentation, state);
  const sections = presentationSections(presentation, state);
  const customShowLinks = presentationCustomShowLinkContext(customShows, state);
  const assetCatalog = createPresentationAssetCatalog();
  const masters = presentationMasters(presentation, state, assetCatalog, customShowLinks);
  const layouts = presentationLayouts(presentation, state, assetCatalog, customShowLinks);
  const slides = presentation.slides.items.map((slide, slideIndex) => {
    const sourceState = sourceStates?.sourceBySlide.get(slide);
    const cloneState = sourceStates?.cloneBySlide.get(slide);
    const bindingState = sourceState || cloneState?.source;
    if (bindingState) {
      if (cloneState && slide.name !== bindingState.name) throw new OpenChestnutCodecError(`Source-preserving PPTX export cannot rename pending clone slide ${slideIndex + 1}.`, [], { code: "unsupported_presentation_slide_clone" });
      if ((slide.layoutId || "") !== (bindingState.wire.layoutId || "")) throw new OpenChestnutCodecError(`Source-preserving PPTX export cannot change slide ${slideIndex + 1}'s layout binding.`, [], { code: cloneState ? "unsupported_presentation_slide_clone" : "presentation_slide_layout_binding_changed" });
      const commentsChanged = presentationSlideCommentSnapshot(slide) !== bindingState.commentSnapshot;
      const addingLegacyComments = !cloneState &&
        presentation.commentFormat === "legacy" &&
        !bindingState.wire.legacyComments?.length &&
        !bindingState.wire.modernComments?.length &&
        slide.comments.items.length > 0 &&
        bindingState.wire.source?.legacyCommentsAddable === true;
      if (commentsChanged && !addingLegacyComments && (!bindingState.wire.modernComments?.length || cloneState)) {
        throw new OpenChestnutCodecError(`Imported presentation slide ${slideIndex + 1} comments are source-bound outside the bounded modern text/status edit profile.`, [], { code: "unsupported_presentation_edit" });
      }
      if (commentsChanged && presentation.commentFormat === "legacy" && !addingLegacyComments) {
        throw new OpenChestnutCodecError(`Source-preserving PPTX export cannot add legacy comments to slide ${slideIndex + 1} because its presentation comment graph is not safely extensible.`, [], { code: "unsupported_presentation_edit" });
      }
      const current = directSlideElements(slide);
      const entries = cloneState?.entries || bindingState.entries;
      if (current.length !== entries.length || entries.some((entry) => !current.includes(entry.model))) {
        throw new OpenChestnutCodecError(`Source-preserving PPTX export requires slide ${slideIndex + 1}'s original ${entries.length}-element topology.`, [], { code: cloneState ? "unsupported_presentation_slide_clone" : "presentation_element_topology_changed" });
      }
      if (!bindingState.wire.speakerNotes && slide.speakerNotes?.text && !bindingState.wire.source?.speakerNotesAddable) {
        throw new OpenChestnutCodecError(`Source-preserving PPTX export cannot add speaker notes to slide ${slideIndex + 1} because its presentation notes graph is not safely extensible.`, [], { code: "unsupported_presentation_edit" });
      }
    }
    const legacyComments = presentation.commentFormat === "legacy"
      ? presentationLegacyComments(slide, Number(bindingState?.wire.source?.slideIndex ?? slideIndex))
      : [];
    const elements = bindingState
      ? (cloneState?.entries || bindingState.entries).map((entry) => {
        if (entry.wire.content.case === "shape") {
          if (entry.wire.content.value.placeholder) {
            return presentationSlidePlaceholder(entry.model, entry.wire, entry.placeholderSnapshot, assetCatalog, customShowLinks);
          }
          return presentationShape(entry.model, entry.wire, assetCatalog, customShowLinks);
        }
        if (entry.wire.content.case === "image") {
          if (presentationImageReadOnlySnapshot(entry.model) !== entry.snapshot) {
            throw new OpenChestnutCodecError(`Presentation image ${entry.model.id} changed outside its embedded rectangular image boundary.`, [], { code: "unsupported_presentation_edit" });
          }
          return presentationImage(entry.model, entry.wire, assetCatalog);
        }
        if (entry.wire.content.case === "table") {
          if (presentationTableReadOnlySnapshot(entry.model) !== entry.snapshot) {
            throw new OpenChestnutCodecError(`Presentation table ${entry.model.id} changed outside its name/frame/plain-text boundary.`, [], { code: "unsupported_presentation_edit" });
          }
          return presentationTable(entry.model, entry.wire);
        }
        if (entry.wire.content.case === "connector") return presentationConnector(entry.model, entry.wire, cloneState?.sourceIdByCloneId);
        if (entry.wire.content.case === "chart") return presentationChart(entry.model, entry.wire);
        if (entry.wire.content.case === "group") return presentationGroup(entry.model, entry.wire, assetCatalog, cloneState?.sourceIdByCloneId, customShowLinks);
        return presentationOpaque(entry.model, entry.wire, entry.snapshot, assetCatalog);
      })
      : directSlideElements(slide)
        .filter((element) => element instanceof Shape || element instanceof ImageElement || element instanceof TableElement || element instanceof ChartElement || element instanceof GroupShape || slide.connectors.items.includes(element))
        .map((element) => presentationElement(element, undefined, assetCatalog, undefined, customShowLinks));
    const modernComments = presentation.commentFormat === "modern"
      ? presentationModernComments(slide, slideIndex, elements, bindingState?.wire.modernComments || [])
      : [];
    const speakerNotes = presentationSpeakerNotes(
      slide,
      bindingState?.wire.speakerNotes,
      assetCatalog,
      customShowLinks,
    );
    const requested = {
      id: sourceState?.wire.id || slide.id,
      name: slide.name,
      source: sourceState?.wire.source,
      ...(slide.layoutId ? { layoutId: slide.layoutId } : {}),
      ...(slide.background?.fill ? { background: wireBackground(slide.background, `slide ${slideIndex + 1}`) } : {}),
      ...(slide.transition?.configured ? { transition: wirePresentationTransition(slide.transition) } : {}),
      ...(speakerNotes ? { speakerNotes } : {}),
      ...(legacyComments.length ? { legacyComments } : {}),
      ...(modernComments.length ? { modernComments } : {}),
      elements,
    };
    if (!cloneState) return requested;
    if (!presentationCloneMatches(requested, cloneState.source.wire)) {
      throw new OpenChestnutCodecError(`Imported presentation clone ${slideIndex + 1} must remain untouched until it has been exported and imported again.`, [], { code: "unsupported_presentation_slide_clone" });
    }
    delete requested.source;
    requested.cloneSource = cloneState.source.wire.source;
    return requested;
  });
  return {
    protocolVersion,
    family: ArtifactFamily.PRESENTATION,
    source: state?.source,
    assets: assetCatalog.assets(),
    opaqueOpc: state?.opaqueOpc,
    diagnostics: state?.diagnostics || [],
    payload: {
      case: "presentation",
      value: {
        id: presentation.id,
        name: state?.name || "",
        slideWidthEmu: emuFromPixels(presentation.slideSize.width, "slideSize.width"),
        slideHeightEmu: emuFromPixels(presentation.slideSize.height, "slideSize.height"),
        slides,
        masters,
        layouts,
        customShows,
        ...(state?.customShowsOpaque ? { customShowsOpaque: true } : {}),
        sections,
        ...(state?.sectionsOpaque ? { sectionsOpaque: true } : {}),
        ...(state?.viewProperties ? { viewProperties: state.viewProperties } : {}),
      },
    },
  };
}

function presentationNativeKind(elementName) {
  return ({ pic: "picture", graphicFrame: "graphicFrame", grpSp: "group", cxnSp: "connector", contentPart: "contentPart" })[elementName] || elementName || "nativeObject";
}

function modelRun(run, customShowLinks) {
  const hyperlink = run.hyperlink?.case === "runHyperlink" ? modelHyperlink(run.hyperlink.value, customShowLinks) : undefined;
  const content = run.content?.case === "lineBreak"
    ? { break: true }
    : run.content?.case === "field"
      ? { field: { id: run.content.value.id, type: run.content.value.type, text: run.content.value.text } }
      : { text: run.content?.case === "text" ? run.content.value : "" };
  return {
    ...content,
    style: {
      ...(run.bold === undefined ? {} : { bold: run.bold }),
      ...(run.italic === undefined ? {} : { italic: run.italic }),
      ...(run.fontSizePoints === undefined ? {} : { fontSize: run.fontSizePoints / POINTS_PER_PIXEL }),
      ...(run.fontFamily === undefined ? {} : { fontFamily: run.fontFamily }),
      ...(run.colorRgb === undefined ? {} : { color: `#${run.colorRgb}` }),
    },
    ...(hyperlink ? { link: hyperlink } : {}),
  };
}

function modelHyperlink(link, customShowLinks) {
  const customShowName = link.target?.case === "customShowId" ? customShowLinks?.get(link.target.value) : undefined;
  if (link.target?.case === "customShowId" && !customShowName) {
    throw new OpenChestnutCodecError(`Presentation run hyperlink references missing custom show ${link.target.value}.`, [], { code: "invalid_presentation_artifact" });
  }
  const target = link.target?.case === "uri"
    ? { uri: link.target.value }
    : link.target?.case === "slideId"
      ? { slideId: link.target.value }
      : link.target?.case === "action"
        ? { action: link.target.value }
        : link.target?.case === "customShowId"
          ? { customShow: customShowName }
        : {};
  return {
    ...target,
    ...(link.tooltip === undefined ? {} : { tooltip: link.tooltip }),
    ...(link.targetFrame === undefined ? {} : { targetFrame: link.targetFrame }),
    ...(link.history === undefined ? {} : { history: link.history }),
    ...(link.highlightClick === undefined ? {} : { highlightClick: link.highlightClick }),
    ...(link.returnToSlide === undefined ? {} : { returnToSlide: link.returnToSlide }),
  };
}

function modelBullet(bullet, assetCatalog) {
  if (bullet?.case === "noBullet") return { bulletNone: true };
  if (bullet?.case === "bulletCharacter") return { bulletCharacter: bullet.value };
  if (bullet?.case === "autoNumber") return { autoNumber: { type: bullet.value.scheme, ...(bullet.value.startAt === undefined ? {} : { startAt: bullet.value.startAt }) } };
  if (bullet?.case === "pictureBullet") {
    if (bullet.value.source?.case === "assetId") return { bulletImage: { dataUrl: assetCatalog.dataUrl(bullet.value.source.value), relationshipMode: "embed" } };
    if (bullet.value.source?.case === "uri") return { bulletImage: { uri: validatePictureBulletUri(bullet.value.source.value), relationshipMode: "link" } };
    throw new OpenChestnutCodecError("Presentation picture bullet has no source.", [], { code: "invalid_presentation_asset" });
  }
  return {};
}

function modelBulletStyle(paragraph) {
  return {
    ...(paragraph.bulletFont?.case === "bulletFontFamily" ? { bulletFont: paragraph.bulletFont.value } : {}),
    ...(paragraph.bulletFont?.case === "bulletFontFollowText" ? { bulletFontFollowText: true } : {}),
    ...(paragraph.bulletColor?.case === "bulletColorRgb" ? { bulletColor: `#${paragraph.bulletColor.value}` } : {}),
    ...(paragraph.bulletColor?.case === "bulletColorScheme" ? { bulletColor: paragraph.bulletColor.value } : {}),
    ...(paragraph.bulletColor?.case === "bulletColorFollowText" ? { bulletColorFollowText: true } : {}),
    ...(paragraph.bulletSize?.case === "bulletSizePoints" ? { bulletSize: paragraph.bulletSize.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.bulletSize?.case === "bulletSizePercent" ? { bulletSizePercent: paragraph.bulletSize.value } : {}),
    ...(paragraph.bulletSize?.case === "bulletSizeFollowText" ? { bulletSizeFollowText: true } : {}),
  };
}

function modelParagraphLayout(paragraph) {
  return {
    ...(paragraph.leftMargin?.case === "marginLeftEmu" ? { marginLeft: Number(paragraph.leftMargin.value) / EMU_PER_PIXEL } : {}),
    ...(paragraph.indentation?.case === "indentEmu" ? { indent: Number(paragraph.indentation.value) / EMU_PER_PIXEL } : {}),
  };
}

function modelParagraphSpacing(paragraph) {
  return {
    ...(paragraph.lineSpacing?.case === "lineSpacingPoints" ? { lineSpacing: paragraph.lineSpacing.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.lineSpacing?.case === "lineSpacingMultiplier" ? { lineSpacing: paragraph.lineSpacing.value } : {}),
    ...(paragraph.spaceBefore?.case === "spaceBeforePoints" ? { spaceBefore: paragraph.spaceBefore.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.spaceBefore?.case === "spaceBeforeMultiplier" ? { spaceBeforePercent: paragraph.spaceBefore.value } : {}),
    ...(paragraph.spaceAfter?.case === "spaceAfterPoints" ? { spaceAfter: paragraph.spaceAfter.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.spaceAfter?.case === "spaceAfterMultiplier" ? { spaceAfterPercent: paragraph.spaceAfter.value } : {}),
  };
}

function modelDefaultRunStyle(paragraph) {
  if (paragraph.defaultRunStyle?.case !== "defaultRunProperties") return {};
  const style = paragraph.defaultRunStyle.value;
  return {
    ...(style.bold === undefined ? {} : { bold: style.bold }),
    ...(style.italic === undefined ? {} : { italic: style.italic }),
    ...(style.fontSizePoints === undefined ? {} : { fontSize: style.fontSizePoints / POINTS_PER_PIXEL }),
    ...(style.fontFamily === undefined ? {} : { fontFamily: style.fontFamily }),
    ...(style.color?.case === "colorRgb" ? { color: `#${style.color.value}` } : {}),
    ...(style.color?.case === "colorScheme" ? { color: style.color.value } : {}),
  };
}

function modelParagraph(paragraph, assetCatalog, { includeRuns = true, customShowLinks } = {}) {
  return {
    ...(includeRuns ? { runs: paragraph.runs.map((run) => modelRun(run, customShowLinks)) } : {}),
    level: paragraph.level ?? 0,
    ...(paragraph.alignment ? { alignment: paragraph.alignment } : {}),
    ...modelBullet(paragraph.bullet, assetCatalog),
    ...modelBulletStyle(paragraph),
    ...modelParagraphLayout(paragraph),
    ...modelParagraphSpacing(paragraph),
    ...(paragraph.tabStops?.length ? { tabStops: paragraph.tabStops.map((tab) => ({ position: Number(tab.positionEmu) / EMU_PER_PIXEL, alignment: tab.alignment })) } : {}),
    style: modelDefaultRunStyle(paragraph),
  };
}

function modelText(shape, assetCatalog, customShowLinks) {
  if (!shape.textBody) return shape.text;
  return shape.textBody.paragraphs.map((paragraph) => modelParagraph(paragraph, assetCatalog, { customShowLinks }));
}

function modelListStyles(shape, assetCatalog) {
  if (!shape.textBody) return {};
  return Object.fromEntries(shape.textBody.listStyles.map((style) => [style.level, modelParagraph(style, assetCatalog, { includeRuns: false })]));
}

function modelTextBodyProperties(shape) {
  const source = shape.textBody?.bodyProperties;
  if (!source) return {};
  const properties = {};
  const insets = {};
  for (const [key, choice] of [["left", source.leftInset], ["top", source.topInset], ["right", source.rightInset], ["bottom", source.bottomInset]]) {
    if (choice?.case?.endsWith("InsetEmu")) insets[key] = Number(choice.value) / EMU_PER_PIXEL;
  }
  if (Object.keys(insets).length) properties.insets = insets;
  if (source.anchor?.case === "verticalAnchor") properties.anchor = source.anchor.value;
  if (source.wrapping?.case === "wrap") properties.wrap = source.wrapping.value;
  if (source.autoFit?.case === "autoFitMode") properties.autoFit = source.autoFit.value;
  if (source.rotation?.case === "rotationAngle60000") properties.rotation = source.rotation.value / ROTATION_UNITS_PER_DEGREE;
  if (source.verticalText?.case === "verticalTextMode") properties.verticalText = source.verticalText.value;
  if (source.verticalOverflow?.case === "verticalOverflowMode") properties.verticalOverflow = source.verticalOverflow.value;
  if (source.horizontalOverflow?.case === "horizontalOverflowMode") properties.horizontalOverflow = source.horizontalOverflow.value;
  const columns = {};
  if (source.columnCount?.case === "columns") columns.count = source.columnCount.value;
  if (source.columnSpacing?.case === "columnSpacingEmu") columns.spacing = Number(source.columnSpacing.value) / EMU_PER_PIXEL;
  if (source.columnDirection?.case === "rightToLeftColumns") columns.rightToLeft = source.columnDirection.value;
  if (Object.keys(columns).length) properties.columns = columns;
  if (source.uprightText?.case === "upright") properties.upright = source.uprightText.value;
  return properties;
}

function modelMasterTextStyles(source, assetCatalog) {
  return Object.fromEntries(MASTER_STYLE_KINDS.map(([kind, levelsField]) => [
    kind,
    Object.fromEntries((source?.textStyles?.[levelsField] || []).map((style) => [style.level, modelParagraph(style, assetCatalog, { includeRuns: false })])),
  ]));
}

function modelPlaceholder(source, assetCatalog, customShowLinks) {
  const shape = { textBody: source.textBody };
  const transform = modelPlaceholderTransform(source.directFrame);
  return {
    id: source.id,
    name: source.name,
    type: source.type,
    idx: source.index,
    ...(source.directFrame ? { position: modelPlaceholderFrame(source.directFrame) } : {}),
    ...(Object.keys(transform).length ? { transform } : {}),
    text: modelText(shape, assetCatalog, customShowLinks),
    paragraphStyles: modelListStyles(shape, assetCatalog),
    textBodyProperties: modelTextBodyProperties(shape),
  };
}

function modelPlaceholderFrame(frame) {
  return {
    left: Number(frame.leftEmu) / EMU_PER_PIXEL,
    top: Number(frame.topEmu) / EMU_PER_PIXEL,
    width: Number(frame.widthEmu) / EMU_PER_PIXEL,
    height: Number(frame.heightEmu) / EMU_PER_PIXEL,
  };
}

function modelPresentationTransform(frame) {
  const transform = {};
  if (frame?.rotationAngle60000 != null) transform.rotationDegrees = frame.rotationAngle60000 / ROTATION_UNITS_PER_DEGREE;
  if (frame?.flipHorizontal != null) transform.flipHorizontal = Boolean(frame.flipHorizontal);
  if (frame?.flipVertical != null) transform.flipVertical = Boolean(frame.flipVertical);
  return transform;
}

function modelCustomGeometryPaths(shape) {
  return (shape.customPaths || []).map((path) => ({
    width: Number(path.width),
    height: Number(path.height),
    commands: path.commands.map((command) => {
      if (command.command.case === "moveTo") return { moveTo: { x: Number(command.command.value.x), y: Number(command.command.value.y) } };
      if (command.command.case === "lineTo") return { lineTo: { x: Number(command.command.value.x), y: Number(command.command.value.y) } };
      if (command.command.case === "cubicBezierTo") return {
        cubicBezTo: {
          x1: Number(command.command.value.control1.x),
          y1: Number(command.command.value.control1.y),
          x2: Number(command.command.value.control2.x),
          y2: Number(command.command.value.control2.y),
          x: Number(command.command.value.end.x),
          y: Number(command.command.value.end.y),
        },
      };
      return { close: {} };
    }),
  }));
}

function modelPlaceholderTransform(frame) {
  return modelPresentationTransform(frame);
}

function slidePlaceholderSnapshot(shape) {
  return JSON.stringify(shape.layoutJson());
}

function slidePlaceholderTextStructureSnapshot(shape) {
  const paragraphs = clonedPresentationValue(shape.text?.paragraphs || []);
  for (const paragraph of paragraphs) {
    for (const run of paragraph.runs || []) {
      if (Object.hasOwn(run, "text")) run.text = "";
      if (run.field) run.field.text = "";
    }
  }
  return JSON.stringify(paragraphs);
}

function slidePlaceholderReadOnlySnapshot(shape) {
  const snapshot = clonedPresentationValue(shape.layoutJson());
  delete snapshot.text;
  delete snapshot.paragraphs;
  snapshot.inheritedParagraphStyles = clonedPresentationValue(shape.text?.inheritedParagraphStyles || {});
  return JSON.stringify(snapshot);
}

function slidePlaceholderState(shape) {
  return {
    full: slidePlaceholderSnapshot(shape),
    readOnly: slidePlaceholderReadOnlySnapshot(shape),
    textStructure: slidePlaceholderTextStructureSnapshot(shape),
  };
}

function isPlainPresentationTextRequest(shape) {
  return JSON.stringify(shape.text?.paragraphs || []) === JSON.stringify(normalizePresentationParagraphs(shape.text?.value || ""));
}

function sourceBoundSlidePlaceholderTextBody(shape, originalShape, originalState, assetCatalog, customShowLinks) {
  if (slidePlaceholderTextStructureSnapshot(shape) === originalState.textStructure) {
    return presentationTextBody(shape, originalShape, assetCatalog, customShowLinks);
  }

  if (!isPlainPresentationTextRequest(shape)) {
    throw new OpenChestnutCodecError(
      `Presentation slide placeholder ${shape.id} changed its source-bound paragraph, inline, or formatting topology. Use text.replace(...) for structured text, or text.set(...) with the source line-break topology intact.`,
      [],
      { code: "presentation_text_topology_changed" },
    );
  }

  // TextFrame.set(...) intentionally presents a plain request. Preserve the
  // imported run/paragraph formatting and map each newline-delimited segment
  // back to exactly one original text run. Ambiguous multi-run spans fail
  // closed; callers can use text.replace(...) for a precise structured edit.
  const textBody = clonedPresentationValue(originalShape.textBody);
  const spans = [[]];
  for (let paragraphIndex = 0; paragraphIndex < (textBody?.paragraphs || []).length; paragraphIndex += 1) {
    const paragraph = textBody.paragraphs[paragraphIndex];
    for (const run of paragraph.runs || []) {
      if (run.content?.case === "text") spans.at(-1).push(run);
      else if (run.content?.case === "lineBreak") spans.push([]);
      else {
        throw new OpenChestnutCodecError(
          `Presentation slide placeholder ${shape.id} contains a field or unsupported inline that cannot be replaced through text.set(...).`,
          [],
          { code: "presentation_text_topology_changed" },
        );
      }
    }
    if (paragraphIndex + 1 < textBody.paragraphs.length) spans.push([]);
  }
  const segments = shape.text.value.split("\n");
  if (segments.length !== spans.length || spans.some((runs) => runs.length !== 1)) {
    throw new OpenChestnutCodecError(
      `Presentation slide placeholder ${shape.id} cannot map text.set(...) onto its source-bound line-break and styled-run topology. Preserve the newline count or use text.replace(...).`,
      [],
      { code: "presentation_text_topology_changed" },
    );
  }
  for (let index = 0; index < segments.length; index += 1) {
    spans[index][0].content = { case: "text", value: segments[index] };
  }
  return textBody;
}

function presentationSlidePlaceholder(shape, original, originalState, assetCatalog, customShowLinks) {
  const currentState = slidePlaceholderState(shape);
  if (currentState.full === originalState.full) return original;
  if (original?.source?.textEditable !== true) {
    throw new OpenChestnutCodecError(
      `Presentation slide placeholder ${shape.id} is source-bound and has no safely editable owner-local text graph.`,
      [],
      { code: "unsupported_presentation_edit" },
    );
  }
  if (currentState.readOnly !== originalState.readOnly) {
    throw new OpenChestnutCodecError(
      `Presentation slide placeholder ${shape.id} may edit only its owner-local text; identity, geometry, formatting, and shape semantics remain source-bound.`,
      [],
      { code: "unsupported_presentation_edit" },
    );
  }
  const requested = clonedPresentationValue(original);
  const originalShape = original.content.value;
  requested.content.value.text = shape.text.value;
  requested.content.value.textBody = sourceBoundSlidePlaceholderTextBody(shape, originalShape, originalState, assetCatalog, customShowLinks);
  return requested;
}

function modelPresentationGroupChild(element, assetCatalog, customShowLinks) {
  const common = { id: element.id, name: element.name };
  if (element.content.case === "shape") {
    const shape = element.content.value;
    if (shape.placeholder) throw new OpenChestnutCodecError(`Presentation group ${element.id} contains an unsupported placeholder child.`, [], { code: "invalid_presentation_group" });
    return {
      kind: "shape",
      ...common,
      geometry: shape.geometry || "rect",
      ...(shape.customPaths?.length ? { customPaths: modelCustomGeometryPaths(shape) } : {}),
      position: {
        left: Number(shape.leftEmu) / EMU_PER_PIXEL,
        top: Number(shape.topEmu) / EMU_PER_PIXEL,
        width: Number(shape.widthEmu) / EMU_PER_PIXEL,
        height: Number(shape.heightEmu) / EMU_PER_PIXEL,
      },
      ...(shape.transform ? { transform: modelPresentationTransform(shape.transform) } : {}),
      fill: shape.fillRgb ? `#${shape.fillRgb}` : "transparent",
      line: { fill: shape.lineRgb ? `#${shape.lineRgb}` : "transparent", width: Number(shape.lineWidthEmu) / EMU_PER_POINT },
      ...(shape.shadow ? { shadow: modelPresentationShadow(shape.shadow) } : {}),
      ...(shape.useBackgroundFill === undefined ? {} : { _openChestnutUseBackgroundFill: shape.useBackgroundFill }),
      text: modelText(shape, assetCatalog, customShowLinks),
      textBodyProperties: modelTextBodyProperties(shape),
    };
  }
  if (element.content.case === "image") {
    const image = element.content.value;
    return {
      kind: "image",
      ...common,
      position: {
        left: Number(image.leftEmu) / EMU_PER_PIXEL,
        top: Number(image.topEmu) / EMU_PER_PIXEL,
        width: Number(image.widthEmu) / EMU_PER_PIXEL,
        height: Number(image.heightEmu) / EMU_PER_PIXEL,
      },
      alt: image.altText,
      dataUrl: assetCatalog.dataUrl(image.assetId),
      fit: "stretch",
      ...(image.crop ? { crop: presentationImageCropFromWire(image.crop) } : {}),
      geometry: "rect",
      ...(image.transform ? { transform: modelPresentationTransform(image.transform) } : {}),
    };
  }
  if (element.content.case === "table") {
    const table = element.content.value;
    return {
      kind: "table",
      ...common,
      position: {
        left: Number(table.leftEmu) / EMU_PER_PIXEL,
        top: Number(table.topEmu) / EMU_PER_PIXEL,
        width: Number(table.widthEmu) / EMU_PER_PIXEL,
        height: Number(table.heightEmu) / EMU_PER_PIXEL,
      },
      values: table.rows.map((row) => row.cells.map((cell) => cell.text)),
      rows: table.rows.length,
      columns: table.columnWidthsEmu.length,
      styleOptions: { headerRow: table.firstRow === true, bandedRows: table.bandedRows === true },
    };
  }
  if (element.content.case === "connector") {
    const connector = element.content.value;
    return {
      kind: "connector",
      ...common,
      connectorType: connector.connectorType || "straight",
      start: { x: Number(connector.startXEmu) / EMU_PER_PIXEL, y: Number(connector.startYEmu) / EMU_PER_PIXEL },
      end: { x: Number(connector.endXEmu) / EMU_PER_PIXEL, y: Number(connector.endYEmu) / EMU_PER_PIXEL },
      startTargetId: connector.startTargetId || undefined,
      endTargetId: connector.endTargetId || undefined,
      line: {
        fill: connector.lineRgb ? `#${connector.lineRgb}` : "transparent",
        width: Number(connector.lineWidthEmu) / EMU_PER_POINT,
        ...(connector.startArrow ? { startArrow: connector.startArrow } : {}),
        ...(connector.endArrow ? { endArrow: connector.endArrow } : {}),
      },
    };
  }
  if (element.content.case === "chart") return { kind: "chart", ...common, ...modelPresentationChart(element.content.value) };
  if (element.content.case === "group") return { kind: "groupShape", ...modelPresentationGroup(element, assetCatalog, customShowLinks) };
  throw new OpenChestnutCodecError(`Presentation group child ${element.id} has unsupported wire content ${element.content.case || "none"}.`, [], { code: "invalid_presentation_group" });
}

function modelPresentationGroup(element, assetCatalog, customShowLinks) {
  const group = element.content.value;
  return {
    id: element.id,
    name: element.name,
    position: {
      left: Number(group.leftEmu) / EMU_PER_PIXEL,
      top: Number(group.topEmu) / EMU_PER_PIXEL,
      width: Number(group.widthEmu) / EMU_PER_PIXEL,
      height: Number(group.heightEmu) / EMU_PER_PIXEL,
    },
    childFrame: {
      left: Number(group.childLeftEmu) / EMU_PER_PIXEL,
      top: Number(group.childTopEmu) / EMU_PER_PIXEL,
      width: Number(group.childWidthEmu) / EMU_PER_PIXEL,
      height: Number(group.childHeightEmu) / EMU_PER_PIXEL,
    },
    children: group.children.map((child) => modelPresentationGroupChild(child, assetCatalog, customShowLinks)),
  };
}

export async function presentationFromEnvelope(envelope) {
  if (envelope.family !== ArtifactFamily.PRESENTATION || envelope.payload.case !== "presentation") {
    throw new OpenChestnutCodecError("OpenChestnut response does not contain a presentation artifact.", [], { code: "invalid_presentation_artifact" });
  }
  const source = envelope.payload.value;
  if (source.customShowsOpaque && source.customShows?.length) {
    throw new OpenChestnutCodecError("OpenChestnut returned both opaque and semantic presentation custom shows.", [], { code: "invalid_presentation_artifact" });
  }
  if (source.sectionsOpaque && source.sections?.length) {
    throw new OpenChestnutCodecError("OpenChestnut returned both opaque and semantic PowerPoint sections.", [], { code: "invalid_presentation_artifact" });
  }
  const customShowLinks = new Map();
  for (const show of source.customShows || []) {
    if (!show.id || customShowLinks.has(show.id)) {
      throw new OpenChestnutCodecError("OpenChestnut returned an invalid or duplicate presentation custom-show ID.", [], { code: "invalid_presentation_artifact" });
    }
    customShowLinks.set(show.id, show.name);
  }
  const nativeGraph = await materializePresentationNativeGraphs(envelope);
  const assetCatalog = createPresentationAssetCatalog(envelope.assets || []);
  const presentation = Presentation.create({
    slideSize: { width: Number(source.slideWidthEmu) / EMU_PER_PIXEL, height: Number(source.slideHeightEmu) / EMU_PER_PIXEL },
  });
  presentation.id = source.id || presentation.id;
  const slideGuides = modelPresentationSlideGuides(source.viewProperties);
  presentation.view._setImportedProperties(modelPresentationView(source.viewProperties));
  const masterStates = [];
  if (source.masters?.length) {
    presentation.masters.items.length = 0;
    for (const sourceMaster of source.masters) {
      const model = presentation.masters.add({
        id: sourceMaster.id,
        name: sourceMaster.name,
        ...(sourceMaster.background ? { background: modelBackground(sourceMaster.background) } : {}),
        placeholders: (sourceMaster.placeholders || []).map((placeholder) => modelPlaceholder(placeholder, assetCatalog, customShowLinks)),
        textParagraphStyles: modelMasterTextStyles(sourceMaster, assetCatalog),
        slideGuides,
      });
      if (!sourceMaster.background) model.background = undefined;
      for (let index = 0; index < sourceMaster.placeholders.length; index += 1) {
        if (!sourceMaster.placeholders[index].directFrame) model.placeholders[index].position = undefined;
      }
      masterStates.push({
        wire: sourceMaster,
        model,
        snapshot: masterReadOnlySnapshot(model),
      });
    }
  }
  const layoutStates = [];
  for (const sourceLayout of source.layouts || []) {
    const model = presentation.layouts.add({
      id: sourceLayout.id,
      name: sourceLayout.name,
      type: sourceLayout.type,
      masterId: sourceLayout.masterId,
      ...(sourceLayout.background ? { background: modelBackground(sourceLayout.background) } : {}),
      placeholders: (sourceLayout.placeholders || []).map((placeholder) => modelPlaceholder(placeholder, assetCatalog, customShowLinks)),
      slideGuides,
    });
    layoutStates.push({
      wire: sourceLayout,
      model,
      snapshot: layoutReadOnlySnapshot(model),
    });
  }
  const slideStates = [];
  for (const sourceSlide of source.slides) {
    const slide = presentation.slides.add({
      name: sourceSlide.name,
      ...(sourceSlide.background ? { background: modelBackground(sourceSlide.background) } : {}),
      ...(sourceSlide.transition ? { transition: modelPresentationTransition(sourceSlide.transition, slideStates.length) } : {}),
    });
    slide.id = sourceSlide.id || slide.id;
    slide.layoutId = sourceSlide.layoutId || undefined;
    slide.addNotes(sourceSlide.speakerNotes?.textBody
      ? modelText(sourceSlide.speakerNotes, assetCatalog, customShowLinks)
      : sourceSlide.speakerNotes?.text || "");
    Object.defineProperty(slide.speakerNotes, PRESENTATION_SPEAKER_NOTES_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        partPresent: Boolean(sourceSlide.speakerNotes),
        editable: Boolean(sourceSlide.speakerNotes?.source?.editable),
        addable: Boolean(!sourceSlide.speakerNotes && sourceSlide.source?.speakerNotesAddable),
      }),
    });
    Object.defineProperty(slide.transition, PRESENTATION_TRANSITION_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        partPresent: Boolean(sourceSlide.source?.transitionPresent),
        editable: Boolean(sourceSlide.source?.transitionEditable),
        addable: false,
      }),
    });
    const entries = [];
    for (const element of sourceSlide.elements) {
      let model;
      if (element.content.case === "shape") {
        const shape = element.content.value;
        const placeholderIdentity = shape.placeholder;
        const layout = placeholderIdentity ? presentation.layouts.getItem(slide.layoutId) : undefined;
        // PowerPoint resolves slide placeholders against their linked layout by
        // idx. Type remains descriptive and deliberately does not participate
        // in this lookup.
        const inheritedPlaceholder = placeholderIdentity?.inheritsGeometry
          ? layout?.effectivePlaceholders().find((candidate) => candidate.idx === Number(placeholderIdentity.index))
          : undefined;
        const directFrame = shape.directFrame ? modelPlaceholderFrame(shape.directFrame) : undefined;
        const directTransform = shape.directFrame ? modelPlaceholderTransform(shape.directFrame) : undefined;
        const effectiveFrame = directFrame || inheritedPlaceholder?.position || {
          left: Number(shape.leftEmu) / EMU_PER_PIXEL,
          top: Number(shape.topEmu) / EMU_PER_PIXEL,
          width: Number(shape.widthEmu) / EMU_PER_PIXEL,
          height: Number(shape.heightEmu) / EMU_PER_PIXEL,
        };
        const geometrySource = directFrame
          ? "slide"
          : placeholderIdentity?.inheritsGeometry
            ? (inheritedPlaceholder?.geometrySource || "unresolved")
            : placeholderIdentity
              ? "slide-unrecognized"
              : undefined;
        const effectiveTransform = directFrame
          ? directTransform
          : placeholderIdentity
            ? inheritedPlaceholder?.transform
            : modelPresentationTransform(shape.transform);
        model = slide.shapes.add({
          id: element.id,
          name: element.name || inheritedPlaceholder?.name,
          geometry: shape.geometry || "rect",
          ...(shape.customPaths?.length ? { customPaths: modelCustomGeometryPaths(shape) } : {}),
          position: { ...effectiveFrame },
          ...(effectiveTransform && Object.keys(effectiveTransform).length ? { transform: effectiveTransform } : {}),
          ...(placeholderIdentity ? { placeholder: {
            layoutId: slide.layoutId,
            type: placeholderIdentity.type,
            idx: Number(placeholderIdentity.index),
            geometrySource,
            textEditable: element.source?.textEditable === true,
          } } : {}),
          fill: shape.fillRgb ? `#${shape.fillRgb}` : "transparent",
          line: { fill: shape.lineRgb ? `#${shape.lineRgb}` : "transparent", width: Number(shape.lineWidthEmu) / EMU_PER_POINT },
          ...(shape.shadow ? { shadow: modelPresentationShadow(shape.shadow) } : {}),
          ...(shape.useBackgroundFill === undefined ? {} : { _openChestnutUseBackgroundFill: shape.useBackgroundFill }),
          text: modelText(shape, assetCatalog, customShowLinks),
          textBodyProperties: modelTextBodyProperties(shape),
        });
        model.text.inheritedParagraphStyles = modelListStyles(shape, assetCatalog);
      } else if (element.content.case === "image") {
        const image = element.content.value;
        model = slide.images.add({
          id: element.id,
          name: element.name,
          position: {
            left: Number(image.leftEmu) / EMU_PER_PIXEL,
            top: Number(image.topEmu) / EMU_PER_PIXEL,
            width: Number(image.widthEmu) / EMU_PER_PIXEL,
            height: Number(image.heightEmu) / EMU_PER_PIXEL,
          },
          alt: image.altText,
          dataUrl: assetCatalog.dataUrl(image.assetId),
          fit: "stretch",
          ...(image.crop ? { crop: presentationImageCropFromWire(image.crop) } : {}),
          geometry: "rect",
          ...(image.transform ? { transform: modelPresentationTransform(image.transform) } : {}),
        });
      } else if (element.content.case === "table") {
        const table = element.content.value;
        model = slide.tables.add({
          id: element.id,
          name: element.name,
          position: {
            left: Number(table.leftEmu) / EMU_PER_PIXEL,
            top: Number(table.topEmu) / EMU_PER_PIXEL,
            width: Number(table.widthEmu) / EMU_PER_PIXEL,
            height: Number(table.heightEmu) / EMU_PER_PIXEL,
          },
          values: table.rows.map((row) => row.cells.map((cell) => cell.text)),
          rows: table.rows.length,
          columns: table.columnWidthsEmu.length,
          styleOptions: {
            headerRow: table.firstRow === true,
            bandedRows: table.bandedRows === true,
          },
          mergeRanges: table.mergeRanges.map((range) => ({
            startRow: Number(range.startRow),
            endRow: Number(range.endRow),
            startColumn: Number(range.startColumn),
            endColumn: Number(range.endColumn),
          })),
        });
      } else if (element.content.case === "connector") {
        const connector = element.content.value;
        model = slide.connectors.add({
          id: element.id,
          name: element.name,
          connectorType: connector.connectorType || "straight",
          start: { x: Number(connector.startXEmu) / EMU_PER_PIXEL, y: Number(connector.startYEmu) / EMU_PER_PIXEL },
          end: { x: Number(connector.endXEmu) / EMU_PER_PIXEL, y: Number(connector.endYEmu) / EMU_PER_PIXEL },
          startTargetId: connector.startTargetId || undefined,
          endTargetId: connector.endTargetId || undefined,
          line: {
            fill: connector.lineRgb ? `#${connector.lineRgb}` : "transparent",
            width: Number(connector.lineWidthEmu) / EMU_PER_POINT,
            ...(connector.startArrow ? { startArrow: connector.startArrow } : {}),
            ...(connector.endArrow ? { endArrow: connector.endArrow } : {}),
          },
        });
      } else if (element.content.case === "chart") {
        const chart = modelPresentationChart(element.content.value);
        model = slide.charts.add(chart.chartType, {
          id: element.id,
          name: element.name,
          ...chart,
        });
      } else if (element.content.case === "group") {
        model = slide.groups.add(modelPresentationGroup(element, assetCatalog, customShowLinks));
      } else if (element.content.case === "opaque") {
        const opaque = element.content.value;
        const sourcePart = sourceSlide.source?.partPath;
        model = slide.nativeObjects.add({
          id: element.id,
          name: element.name,
          nativeKind: opaque.nativeKind || presentationNativeKind(opaque.elementName),
          position: {
            left: Number(opaque.leftEmu) / EMU_PER_PIXEL,
            top: Number(opaque.topEmu) / EMU_PER_PIXEL,
            width: Number(opaque.widthEmu) / EMU_PER_PIXEL,
            height: Number(opaque.heightEmu) / EMU_PER_PIXEL,
          },
          rawXml: opaque.rawXml,
          sourcePart,
          editable: false,
          ...(opaque.oleWorkbook ? { oleWorkbook: {
            partPath: opaque.oleWorkbook.partPath,
            contentType: opaque.oleWorkbook.contentType,
            sourceSha256: opaque.oleWorkbook.sourceSha256,
            relationshipId: opaque.oleWorkbook.relationshipId,
          } } : {}),
          ...(opaque.diagramText ? { diagramText: {
            partPath: opaque.diagramText.partPath,
            contentType: opaque.diagramText.contentType,
            sourceSha256: opaque.diagramText.sourceSha256,
            relationshipId: opaque.diagramText.relationshipId,
            nodes: (opaque.diagramText.nodes || []).map((node) => ({ id: node.modelId, text: node.text })),
          } } : {}),
          ...nativeGraph(opaque, sourcePart),
        });
      } else {
        throw new OpenChestnutCodecError(`Presentation element ${element.id} has no supported wire content.`, [], { code: "invalid_presentation_artifact" });
      }
      entries.push({
        wire: element,
        model,
        placeholderSnapshot: element.content.case === "shape" && element.content.value.placeholder
          ? slidePlaceholderState(model)
          : undefined,
        snapshot: element.content.case === "opaque"
          ? opaquePresentationSnapshot(model)
          : element.content.case === "image"
            ? presentationImageReadOnlySnapshot(model)
            : element.content.case === "table"
              ? presentationTableReadOnlySnapshot(model)
            : undefined,
      });
    }
    for (const sourceThread of sourceSlide.modernComments || []) {
      presentation.commentFormat = "modern";
      const moniker = sourceThread.anchor?.monikers?.[0];
      const textRange = sourceThread.anchor?.kind === PresentationModernCommentAnchor_Kind.TEXT_RANGE;
      const target = slide.resolve(sourceThread.targetId);
      const targetElement = textRange ? slide.resolve(target?.parentId) : target;
      if (targetElement && moniker) {
        targetElement.nativeId = Number(moniker.nativeId);
        targetElement.creationId = moniker.creationId || undefined;
        targetElement.moniker = moniker.type;
      }
      slide.nativeSlideId = Number(sourceThread.anchor?.nativeSlideId || 0) || undefined;
      const sourceComments = [sourceThread.root, ...(sourceThread.replies || [])];
      const comments = sourceComments.map((comment) => ({
        nativeId: comment.id,
        authorId: comment.authorId,
        author: comment.author,
        initials: comment.initials || undefined,
        userId: comment.userId || undefined,
        providerId: comment.providerId || undefined,
        person: {
          id: comment.authorId,
          name: comment.author,
          initials: comment.initials || undefined,
          userId: comment.userId || undefined,
          providerId: comment.providerId || undefined,
        },
        text: comment.text,
        created: comment.createdAt,
        status: comment.status,
      }));
      const root = comments[0];
      slide.comments.addThread(sourceThread.targetId, root.text, {
        id: sourceThread.id,
        author: root.author,
        created: root.created,
        resolved: ["resolved", "closed"].includes(root.status),
        nativeFormat: "modern",
        nativeAnchor: {
          format: "modern",
          type: textRange ? "textRange" : "element",
          nativeId: Number(moniker?.nativeId || 0),
          creationId: moniker?.creationId || undefined,
          moniker: moniker?.type,
          nativeSlideId: Number(sourceThread.anchor?.nativeSlideId || 0),
          ...(textRange ? {
            textStart: Number(sourceThread.anchor?.textStart || 0),
            textLength: Number(sourceThread.anchor?.textLength || 0),
            cp: Number(sourceThread.anchor?.textStart || 0),
            length: Number(sourceThread.anchor?.textLength || 0),
            ...(sourceThread.anchor?.contextLength === undefined ? {} : { contextLength: Number(sourceThread.anchor.contextLength) }),
            ...(sourceThread.anchor?.contextHash === undefined ? {} : { contextHash: Number(sourceThread.anchor.contextHash) }),
          } : {}),
        },
        position: {
          x: Number(sourceThread.positionXEmu || 0),
          y: Number(sourceThread.positionYEmu || 0),
          unit: "emu",
        },
        comments,
      });
    }
    for (const sourceComment of sourceSlide.legacyComments || []) {
      const created = sourceComment.createdAt || new Date(0).toISOString();
      const nativeAuthorId = Number(sourceComment.nativeAuthorId || 0);
      const nativeIndex = Number(sourceComment.nativeIndex || 0);
      const positionXEmu = Number(sourceComment.positionXEmu || 0);
      const positionYEmu = Number(sourceComment.positionYEmu || 0);
      slide.comments.addThread(undefined, sourceComment.text, {
        id: sourceComment.id,
        author: sourceComment.author,
        created,
        nativeFormat: "legacy",
        nativeAnchor: {
          format: "legacy",
          nativeAuthorId,
          nativeIndex,
          positionXEmu,
          positionYEmu,
        },
        position: { x: positionXEmu / EMU_PER_PIXEL, y: positionYEmu / EMU_PER_PIXEL, unit: "px" },
        comments: [{
          nativeId: `legacy:${nativeAuthorId}:${nativeIndex}`,
          author: sourceComment.author,
          text: sourceComment.text,
          created,
        }],
      });
    }
    Object.defineProperty(slide.comments, PRESENTATION_LEGACY_COMMENTS_CAPABILITY, {
      value: Object.freeze({
        sourceBound: true,
        format: sourceSlide.source?.commentFamily || "legacy",
        partPresent: Boolean(sourceSlide.source?.commentPartPresent),
        addable: Boolean(
          !sourceSlide.legacyComments?.length &&
          !sourceSlide.modernComments?.length &&
          sourceSlide.source?.legacyCommentsAddable
        ),
      }),
    });
    slideStates.push({
      wire: sourceSlide,
      slide,
      name: slide.name,
      commentSnapshot: presentationSlideCommentSnapshot(slide),
      entries,
    });
  }
  const customShowStates = [];
  for (const sourceShow of source.customShows || []) {
    const model = presentation.customShows.add({
      id: sourceShow.id,
      name: sourceShow.name,
      nativeId: Number(sourceShow.nativeId),
      slideIds: [...sourceShow.slideIds],
    });
    customShowStates.push({ wire: sourceShow, model });
  }
  const sectionStates = [];
  for (const sourceSection of source.sections || []) {
    const model = presentation.sections.add({
      id: sourceSection.id,
      name: sourceSection.name,
      nativeId: sourceSection.nativeId,
      slideIds: [...sourceSection.slideIds],
    });
    sectionStates.push({ wire: sourceSection, model });
  }
  const presentationState = {
    source: envelope.source,
    opaqueOpc: envelope.opaqueOpc,
    diagnostics: envelope.diagnostics,
    name: source.name,
    slideWidthEmu: source.slideWidthEmu,
    slideHeightEmu: source.slideHeightEmu,
    viewProperties: source.viewProperties,
    customShowsOpaque: Boolean(source.customShowsOpaque),
    customShows: customShowStates,
    sectionsOpaque: Boolean(source.sectionsOpaque),
    sections: sectionStates,
    advancedSnapshot: presentationAdvancedSnapshot(presentation),
    masters: masterStates,
    layouts: layoutStates,
    slides: slideStates,
    clones: [],
  };
  Object.defineProperty(presentation, PRESENTATION_STATE, {
    configurable: true,
    value: presentationState,
    writable: true,
  });
  Object.defineProperty(presentation, PRESENTATION_SLIDE_DUPLICATOR, {
    configurable: true,
    value: (slide) => duplicateImportedPresentationSlide(presentation, presentationState, slide),
  });
  return presentation;
}
