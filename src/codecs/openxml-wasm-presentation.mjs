import { Presentation } from "../index.mjs";
import { ArtifactFamily } from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { normalizePresentationRunLink } from "../presentation/ooxml-hyperlinks.mjs";
import { isPresentationAutoNumberType } from "../presentation/text-paragraphs.mjs";
import { createPresentationAssetCatalog, validatePictureBulletUri } from "./openxml-wasm-assets.mjs";
import { OpenXmlWasmCodecError } from "./openxml-wasm-error.mjs";

const EMU_PER_PIXEL = 9525;
const EMU_PER_POINT = 12700;
const POINTS_PER_PIXEL = 0.75;
const MAX_FONT_SIZE_PIXELS = 1024;
const MAX_PARAGRAPH_COORDINATE_EMU = 51_206_400;
const PRESENTATION_STATE = Symbol.for("open-office-artifact-tool.openxml-wasm-presentation-state");
const PRESENTATION_SCHEME_COLORS = new Set([
  "dk1", "lt1", "dk2", "lt2", "tx1", "bg1", "tx2", "bg2",
  "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink",
]);
const RUN_STYLE_KEYS = new Set(["bold", "italic", "fontSize", "fontFamily", "color"]);
const PARAGRAPH_KEYS = new Set([
  "runs", "level", "alignment", "style", "bulletCharacter", "autoNumber", "bulletImage", "bulletNone",
  "bulletFont", "bulletFontFollowText", "bulletColor", "bulletColorFollowText",
  "bulletSize", "bulletSizePercent", "bulletSizeFollowText", "tabStops", "marginLeft", "indent",
]);

function emuFromPixels(value, name) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) throw new OpenXmlWasmCodecError(`${name} must be a non-negative finite number.`, [], { code: "invalid_presentation_frame" });
  return BigInt(Math.round(number * EMU_PER_PIXEL));
}

function paragraphEmuFromPixels(value, name, { allowNegative = false } = {}) {
  const number = Number(value);
  const emu = Math.round(number * EMU_PER_PIXEL);
  if (!Number.isFinite(number) || (!allowNegative && number < 0) || emu < (allowNegative ? -MAX_PARAGRAPH_COORDINATE_EMU : 0) || emu > MAX_PARAGRAPH_COORDINATE_EMU) {
    throw new OpenXmlWasmCodecError(`${name} is outside the supported DrawingML coordinate range.`, [], { code: "invalid_presentation_text" });
  }
  return BigInt(emu);
}

function presentationRgb(value, name) {
  const raw = typeof value === "string" ? value : value?.color || value?.fill;
  if (raw == null || raw === "transparent" || raw === "none") return "";
  const match = /^#([0-9a-f]{6})$/i.exec(String(raw));
  if (!match) throw new OpenXmlWasmCodecError(`${name} must be transparent or a six-digit RGB color.`, [], { code: "unsupported_presentation_features" });
  return match[1].toUpperCase();
}

function unsupportedStyleFields(style = {}) {
  return Object.keys(style).filter((key) => !RUN_STYLE_KEYS.has(key));
}

function wireHyperlink(value, original, shapeId) {
  if (value == null) {
    if (new Set(["runHyperlink", "noHyperlink"]).has(original?.hyperlink?.case)) return { case: "noHyperlink", value: true };
    return undefined;
  }
  let link;
  try {
    link = normalizePresentationRunLink(value);
  } catch (error) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses an invalid run hyperlink: ${error.message}`, [], { code: "invalid_presentation_hyperlink" });
  }
  if (link.customShow) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses a custom-show hyperlink outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
  const target = link.uri
    ? { case: "uri", value: link.uri }
    : link.slideId
      ? { case: "slideId", value: link.slideId }
      : link.action
        ? { case: "action", value: link.action }
        : undefined;
  if (!target) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses an unsupported run hyperlink target.`, [], { code: "unsupported_presentation_features" });
  return {
    case: "runHyperlink",
    value: {
      target,
      ...(link.tooltip == null ? {} : { tooltip: link.tooltip }),
      ...(link.targetFrame == null ? {} : { targetFrame: link.targetFrame }),
      ...(link.history == null ? {} : { history: link.history }),
      ...(link.highlightClick == null ? {} : { highlightClick: link.highlightClick }),
    },
  };
}

function wireRun(run, inheritedStyle, shapeId, original) {
  const unsupported = unsupportedStyleFields(run.style);
  if (unsupported.length) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses unsupported run style fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const style = { ...inheritedStyle, ...(run.style || {}) };
  const fontSize = style.fontSize == null ? undefined : Number(style.fontSize);
  if (fontSize !== undefined && (!Number.isFinite(fontSize) || fontSize <= 0 || fontSize > MAX_FONT_SIZE_PIXELS)) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses a font size outside the supported 0-${MAX_FONT_SIZE_PIXELS} pixel range.`, [], { code: "invalid_presentation_text" });
  }
  const fontFamily = style.fontFamily == null ? undefined : String(style.fontFamily);
  if (fontFamily !== undefined && (!fontFamily.trim() || fontFamily.length > 255)) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses an invalid font family.`, [], { code: "invalid_presentation_text" });
  }
  const colorRgb = style.color == null ? undefined : presentationRgb(style.color, `${shapeId}.text.color`);
  if (colorRgb === "") {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses a transparent run color outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
  }
  const hyperlink = wireHyperlink(run.link, original, shapeId);
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
    throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} paragraph selects more than one list marker.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletCharacter != null) {
    const character = String(paragraph.bulletCharacter);
    if ([...character].length !== 1) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} bullet character must contain one Unicode scalar value.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletCharacter", value: character };
  }
  if (paragraph.autoNumber) {
    const scheme = String(paragraph.autoNumber.type || paragraph.autoNumber.scheme || "");
    if (!isPresentationAutoNumberType(scheme)) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses unsupported auto-number scheme ${scheme || "(missing)"}.`, [], { code: "invalid_presentation_text" });
    const rawStart = paragraph.autoNumber.startAt ?? paragraph.autoNumber.start;
    const startAt = rawStart == null ? undefined : Number(rawStart);
    if (startAt !== undefined && (!Number.isInteger(startAt) || startAt < 1 || startAt > 32767)) {
      throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} auto-number start must be from 1 through 32767.`, [], { code: "invalid_presentation_text" });
    }
    return { case: "autoNumber", value: { scheme, ...(startAt === undefined ? {} : { startAt }) } };
  }
  if (paragraph.bulletImage) {
    if (paragraph.bulletImage.relationshipId) {
      throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses an unresolved picture-bullet relationship ID.`, [], { code: "invalid_presentation_asset" });
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
    throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} paragraph selects both a bullet font and follow-text font.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletFont != null) {
    const family = String(paragraph.bulletFont).trim();
    if (!family || family.length > 255) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses an invalid bullet font family.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletFontFamily", value: family };
  }
  if (paragraph.bulletFontFollowText === true || new Set(["bulletFontFamily", "bulletFontFollowText"]).has(original?.bulletFont?.case)) {
    return { case: "bulletFontFollowText", value: true };
  }
  return undefined;
}

function wireBulletColor(paragraph, original, shapeId) {
  if (paragraph.bulletColor != null && paragraph.bulletColorFollowText === true) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} paragraph selects both a bullet color and follow-text color.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletColor != null) {
    const scheme = String(paragraph.bulletColor).trim();
    if (PRESENTATION_SCHEME_COLORS.has(scheme)) return { case: "bulletColorScheme", value: scheme };
    const rgb = presentationRgb(paragraph.bulletColor, `${shapeId}.text.bulletColor`);
    if (!rgb) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses a transparent bullet color outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
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
    throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} paragraph selects more than one bullet size.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletSize != null) {
    const pixels = Number(paragraph.bulletSize);
    if (!Number.isFinite(pixels) || pixels < 4 / 3 || pixels > MAX_FONT_SIZE_PIXELS) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses an invalid fixed bullet size.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletSizePoints", value: pixels * POINTS_PER_PIXEL };
  }
  if (paragraph.bulletSizePercent != null) {
    const percent = Number(paragraph.bulletSizePercent);
    if (!Number.isFinite(percent) || percent < 0.25 || percent > 4) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses an invalid percentage bullet size.`, [], { code: "invalid_presentation_text" });
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

function wireParagraph(paragraph, textStyle, original, shapeId, assetCatalog) {
  const unsupported = Object.keys(paragraph).filter((key) => !PARAGRAPH_KEYS.has(key));
  if (unsupported.length) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses unsupported paragraph fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const paragraphStyleUnsupported = unsupportedStyleFields(paragraph.style);
  if (paragraphStyleUnsupported.length) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses unsupported paragraph text style fields: ${paragraphStyleUnsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const level = Number(paragraph.level || 0);
  if (!Number.isInteger(level) || level < 0 || level > 8) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses a paragraph level outside the supported 0-8 range.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.alignment && !new Set(["left", "center", "right", "justify"]).has(paragraph.alignment)) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses unsupported paragraph alignment ${paragraph.alignment}.`, [], { code: "invalid_presentation_text" });
  }
  const originalLevel = original?.level;
  const includeLevel = level !== 0 || originalLevel !== undefined;
  const bullet = wireBullet(paragraph, original, shapeId, assetCatalog);
  const bulletFont = wireBulletFont(paragraph, original, shapeId);
  const bulletColor = wireBulletColor(paragraph, original, shapeId);
  const bulletSize = wireBulletSize(paragraph, original, shapeId);
  const tabs = wireTabStops(paragraph, original, shapeId);
  const layout = wireParagraphLayout(paragraph, original, shapeId);
  return {
    ...(includeLevel ? { level } : {}),
    ...(paragraph.alignment ? { alignment: paragraph.alignment } : {}),
    runs: (paragraph.runs || []).map((run, index) => wireRun(run, { ...textStyle, ...(paragraph.style || {}) }, shapeId, original?.runs?.[index])),
    ...(bullet ? { bullet } : {}),
    ...(bulletFont ? { bulletFont } : {}),
    ...(bulletColor ? { bulletColor } : {}),
    ...(bulletSize ? { bulletSize } : {}),
    ...tabs,
    ...layout,
  };
}

function modelRunCase(run) {
  if (run.break) return "lineBreak";
  if (run.field) return "field";
  return "text";
}

function presentationTextBody(shape, original, assetCatalog) {
  if (Object.keys(shape.text?.inheritedParagraphStyles || {}).length) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shape.id} uses inherited paragraph styles outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
  }
  const textStyleUnsupported = unsupportedStyleFields(shape.text?.style);
  if (textStyleUnsupported.length) throw new OpenXmlWasmCodecError(`Presentation shape ${shape.id} uses unsupported text-frame style fields: ${textStyleUnsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const paragraphs = shape.text?.paragraphs || [];
  if (original?.textBody && (original.textBody.paragraphs.length !== paragraphs.length || original.textBody.paragraphs.some((paragraph, index) => paragraph.runs.length !== (paragraphs[index]?.runs || []).length || paragraph.runs.some((run, runIndex) => run.content?.case !== modelRunCase(paragraphs[index].runs[runIndex]))))) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shape.id} changed its source-bound paragraph/inline topology.`, [], { code: "presentation_text_topology_changed" });
  }
  return {
    paragraphs: paragraphs.map((paragraph, index) => wireParagraph(paragraph, shape.text?.style || {}, original?.textBody?.paragraphs?.[index], shape.id, assetCatalog)),
  };
}

function presentationShape(shape, original, assetCatalog) {
  const originalShape = original?.content?.case === "shape" ? original.content.value : original;
  if (!new Set(["rect", "ellipse"]).has(shape.geometry)) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shape.id} uses unsupported geometry ${shape.geometry}.`, [], { code: "unsupported_presentation_features" });
  }
  const position = shape.position || {};
  const lineWidth = Number(shape.line?.width ?? 1);
  if (!Number.isFinite(lineWidth) || lineWidth < 0) throw new OpenXmlWasmCodecError(`Presentation shape ${shape.id} has an invalid line width.`, [], { code: "invalid_presentation_frame" });
  const textBody = presentationTextBody(shape, originalShape, assetCatalog);
  return {
    id: original?.id || shape.id,
    name: shape.name || original?.name || "",
    source: original?.source,
    content: {
      case: "shape",
      value: {
        geometry: shape.geometry,
        leftEmu: emuFromPixels(position.left, `${shape.id}.position.left`),
        topEmu: emuFromPixels(position.top, `${shape.id}.position.top`),
        widthEmu: emuFromPixels(position.width, `${shape.id}.position.width`),
        heightEmu: emuFromPixels(position.height, `${shape.id}.position.height`),
        text: shape.text?.value || "",
        textBody,
        fillRgb: presentationRgb(shape.fill, `${shape.id}.fill`),
        lineRgb: presentationRgb(shape.line?.fill || shape.line?.color || (lineWidth > 0 ? "#334155" : "transparent"), `${shape.id}.line.fill`),
        lineWidthEmu: BigInt(Math.round(lineWidth * EMU_PER_POINT)),
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

function unsupportedPresentationFeatures(presentation) {
  const unsupported = [];
  if (presentation.layouts?.items?.length) unsupported.push("slide layouts");
  if (presentation.masters?.items?.length !== 1 || presentation.master?.configured) unsupported.push("configured slide masters");
  if (presentation.customShows?.items?.length) unsupported.push("custom shows");
  if (presentation.commentFormat !== "legacy") unsupported.push("modern comments");
  for (const slide of presentation.slides?.items || []) {
    const prefix = `slide ${slide.index + 1}`;
    if (slide.layoutId) unsupported.push(`${prefix} layout binding`);
    if (slide.speakerNotes?.text) unsupported.push(`${prefix} speaker notes`);
    if (slide.background?.fill) unsupported.push(`${prefix} background`);
    if (slide.comments?.items?.length) unsupported.push(`${prefix} comments`);
    if (slide.tables?.items?.length) unsupported.push(`${prefix} tables`);
    if (slide.charts?.items?.length) unsupported.push(`${prefix} charts`);
    if (slide.images?.items?.length) unsupported.push(`${prefix} images`);
    if (slide.connectors?.items?.length) unsupported.push(`${prefix} connectors`);
    if (slide.groups?.items?.length) unsupported.push(`${prefix} groups`);
    if (slide.nativeObjects?.items?.length) unsupported.push(`${prefix} native objects`);
  }
  return unsupported;
}

function opaquePresentationSnapshot(object) {
  return JSON.stringify({ id: object.id, name: object.name, nativeKind: object.nativeKind, position: object.position, rawXml: object.rawXml });
}

export function presentationEnvelope(presentation, protocolVersion) {
  if (!(presentation instanceof Presentation)) throw new TypeError("exportPptxWithOpenXmlWasm expects a Presentation instance.");
  if (!presentation.slides?.items?.length) throw new OpenXmlWasmCodecError("Presentation must contain at least one slide.", [], { code: "missing_slides" });
  const state = presentation[PRESENTATION_STATE];
  if (!state) {
    const unsupported = unsupportedPresentationFeatures(presentation);
    if (unsupported.length) {
      throw new OpenXmlWasmCodecError(`The PPTX WebAssembly vertical slice cannot author: ${unsupported.slice(0, 8).join(", ")}${unsupported.length > 8 ? `, and ${unsupported.length - 8} more` : ""}. Use PresentationFile.exportPptx until parity reaches these features.`, [], { code: "unsupported_presentation_features" });
    }
  } else {
    if (state.slides.length !== presentation.slides.items.length) throw new OpenXmlWasmCodecError(`Source-preserving PPTX export requires the original ${state.slides.length}-slide topology.`, [], { code: "presentation_topology_changed" });
    if (Number(state.slideWidthEmu) !== Math.round(Number(presentation.slideSize.width) * EMU_PER_PIXEL) || Number(state.slideHeightEmu) !== Math.round(Number(presentation.slideSize.height) * EMU_PER_PIXEL)) {
      throw new OpenXmlWasmCodecError("Source-preserving PPTX export does not yet support changing slide dimensions.", [], { code: "unsupported_presentation_edit" });
    }
  }

  const assetCatalog = createPresentationAssetCatalog();
  const slides = presentation.slides.items.map((slide, slideIndex) => {
    const sourceState = state?.slides[slideIndex];
    if (sourceState) {
      if (slide.name !== sourceState.name) throw new OpenXmlWasmCodecError(`Source-preserving PPTX export does not yet support renaming slide ${slideIndex + 1}.`, [], { code: "unsupported_presentation_edit" });
      const current = directSlideElements(slide);
      if (current.length !== sourceState.entries.length || sourceState.entries.some((entry) => !current.includes(entry.model))) {
        throw new OpenXmlWasmCodecError(`Source-preserving PPTX export requires slide ${slideIndex + 1}'s original ${sourceState.entries.length}-element topology.`, [], { code: "presentation_element_topology_changed" });
      }
    }
    return {
      id: sourceState?.wire.id || slide.id,
      name: slide.name,
      source: sourceState?.wire.source,
      elements: sourceState
        ? sourceState.entries.map((entry) => {
            if (entry.wire.content.case === "shape") return presentationShape(entry.model, entry.wire, assetCatalog);
            if (opaquePresentationSnapshot(entry.model) !== entry.snapshot) throw new OpenXmlWasmCodecError(`Presentation element ${entry.model.id} is preserved but not editable by this codec slice.`, [], { code: "unsupported_presentation_edit" });
            return entry.wire;
          })
        : slide.shapes.items.map((shape) => presentationShape(shape, undefined, assetCatalog)),
    };
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
      },
    },
  };
}

function presentationNativeKind(elementName) {
  return ({ pic: "picture", graphicFrame: "graphicFrame", grpSp: "group", cxnSp: "connector", contentPart: "contentPart" })[elementName] || elementName || "nativeObject";
}

function modelRun(run) {
  const hyperlink = run.hyperlink?.case === "runHyperlink" ? modelHyperlink(run.hyperlink.value) : undefined;
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

function modelHyperlink(link) {
  const target = link.target?.case === "uri"
    ? { uri: link.target.value }
    : link.target?.case === "slideId"
      ? { slideId: link.target.value }
      : link.target?.case === "action"
        ? { action: link.target.value }
        : {};
  return {
    ...target,
    ...(link.tooltip === undefined ? {} : { tooltip: link.tooltip }),
    ...(link.targetFrame === undefined ? {} : { targetFrame: link.targetFrame }),
    ...(link.history === undefined ? {} : { history: link.history }),
    ...(link.highlightClick === undefined ? {} : { highlightClick: link.highlightClick }),
  };
}

function modelBullet(bullet, assetCatalog) {
  if (bullet?.case === "noBullet") return { bulletNone: true };
  if (bullet?.case === "bulletCharacter") return { bulletCharacter: bullet.value };
  if (bullet?.case === "autoNumber") return { autoNumber: { type: bullet.value.scheme, ...(bullet.value.startAt === undefined ? {} : { startAt: bullet.value.startAt }) } };
  if (bullet?.case === "pictureBullet") {
    if (bullet.value.source?.case === "assetId") return { bulletImage: { dataUrl: assetCatalog.dataUrl(bullet.value.source.value), relationshipMode: "embed" } };
    if (bullet.value.source?.case === "uri") return { bulletImage: { uri: validatePictureBulletUri(bullet.value.source.value), relationshipMode: "link" } };
    throw new OpenXmlWasmCodecError("Presentation picture bullet has no source.", [], { code: "invalid_presentation_asset" });
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

function modelText(shape, assetCatalog) {
  if (!shape.textBody) return shape.text;
  return shape.textBody.paragraphs.map((paragraph) => ({
    runs: paragraph.runs.map(modelRun),
    level: paragraph.level ?? 0,
    ...(paragraph.alignment ? { alignment: paragraph.alignment } : {}),
    ...modelBullet(paragraph.bullet, assetCatalog),
    ...modelBulletStyle(paragraph),
    ...modelParagraphLayout(paragraph),
    ...(paragraph.tabStops?.length ? { tabStops: paragraph.tabStops.map((tab) => ({ position: Number(tab.positionEmu) / EMU_PER_PIXEL, alignment: tab.alignment })) } : {}),
    style: {},
  }));
}

export function presentationFromEnvelope(envelope) {
  if (envelope.family !== ArtifactFamily.PRESENTATION || envelope.payload.case !== "presentation") {
    throw new OpenXmlWasmCodecError("OpenXML WebAssembly response does not contain a presentation artifact.", [], { code: "invalid_presentation_artifact" });
  }
  const source = envelope.payload.value;
  const assetCatalog = createPresentationAssetCatalog(envelope.assets || []);
  const presentation = Presentation.create({
    slideSize: { width: Number(source.slideWidthEmu) / EMU_PER_PIXEL, height: Number(source.slideHeightEmu) / EMU_PER_PIXEL },
  });
  presentation.id = source.id || presentation.id;
  const slideStates = [];
  for (const sourceSlide of source.slides) {
    const slide = presentation.slides.add({ name: sourceSlide.name });
    slide.id = sourceSlide.id || slide.id;
    const entries = [];
    for (const element of sourceSlide.elements) {
      let model;
      if (element.content.case === "shape") {
        const shape = element.content.value;
        model = slide.shapes.add({
          id: element.id,
          name: element.name,
          geometry: shape.geometry || "rect",
          position: {
            left: Number(shape.leftEmu) / EMU_PER_PIXEL,
            top: Number(shape.topEmu) / EMU_PER_PIXEL,
            width: Number(shape.widthEmu) / EMU_PER_PIXEL,
            height: Number(shape.heightEmu) / EMU_PER_PIXEL,
          },
          fill: shape.fillRgb ? `#${shape.fillRgb}` : "transparent",
          line: { fill: shape.lineRgb ? `#${shape.lineRgb}` : "transparent", width: Number(shape.lineWidthEmu) / EMU_PER_POINT },
          text: modelText(shape, assetCatalog),
        });
      } else if (element.content.case === "opaque") {
        const opaque = element.content.value;
        model = slide.nativeObjects.add({
          id: element.id,
          name: element.name,
          nativeKind: presentationNativeKind(opaque.elementName),
          position: {
            left: Number(opaque.leftEmu) / EMU_PER_PIXEL,
            top: Number(opaque.topEmu) / EMU_PER_PIXEL,
            width: Number(opaque.widthEmu) / EMU_PER_PIXEL,
            height: Number(opaque.heightEmu) / EMU_PER_PIXEL,
          },
          rawXml: opaque.rawXml,
          sourcePart: sourceSlide.source?.partPath,
        });
      } else {
        throw new OpenXmlWasmCodecError(`Presentation element ${element.id} has no supported wire content.`, [], { code: "invalid_presentation_artifact" });
      }
      entries.push({ wire: element, model, snapshot: element.content.case === "opaque" ? opaquePresentationSnapshot(model) : undefined });
    }
    slideStates.push({ wire: sourceSlide, name: slide.name, entries });
  }
  Object.defineProperty(presentation, PRESENTATION_STATE, {
    configurable: true,
    value: {
      source: envelope.source,
      opaqueOpc: envelope.opaqueOpc,
      diagnostics: envelope.diagnostics,
      name: source.name,
      slideWidthEmu: source.slideWidthEmu,
      slideHeightEmu: source.slideHeightEmu,
      slides: slideStates,
    },
    writable: true,
  });
  return presentation;
}
