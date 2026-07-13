import { Presentation } from "../index.mjs";
import { ArtifactFamily } from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { OpenXmlWasmCodecError } from "./openxml-wasm-error.mjs";

const EMU_PER_PIXEL = 9525;
const EMU_PER_POINT = 12700;
const POINTS_PER_PIXEL = 0.75;
const MAX_FONT_SIZE_PIXELS = 1024;
const PRESENTATION_STATE = Symbol.for("open-office-artifact-tool.openxml-wasm-presentation-state");
const RUN_STYLE_KEYS = new Set(["bold", "italic", "fontSize", "fontFamily", "color"]);
const PARAGRAPH_KEYS = new Set(["runs", "level", "alignment", "style"]);

function emuFromPixels(value, name) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) throw new OpenXmlWasmCodecError(`${name} must be a non-negative finite number.`, [], { code: "invalid_presentation_frame" });
  return BigInt(Math.round(number * EMU_PER_PIXEL));
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

function wireRun(run, inheritedStyle, shapeId) {
  if (run.link) throw new OpenXmlWasmCodecError(`Presentation shape ${shapeId} uses a run hyperlink outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
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
  return {
    text: String(run.text ?? ""),
    ...(style.bold == null ? {} : { bold: Boolean(style.bold) }),
    ...(style.italic == null ? {} : { italic: Boolean(style.italic) }),
    ...(fontSize === undefined ? {} : { fontSizePoints: fontSize * POINTS_PER_PIXEL }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(colorRgb === undefined ? {} : { colorRgb }),
  };
}

function wireParagraph(paragraph, textStyle, original, shapeId) {
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
  return {
    ...(includeLevel ? { level } : {}),
    ...(paragraph.alignment ? { alignment: paragraph.alignment } : {}),
    runs: (paragraph.runs || []).map((run) => wireRun(run, { ...textStyle, ...(paragraph.style || {}) }, shapeId)),
  };
}

function presentationTextBody(shape, original) {
  if (Object.keys(shape.text?.inheritedParagraphStyles || {}).length) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shape.id} uses inherited paragraph styles outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
  }
  const textStyleUnsupported = unsupportedStyleFields(shape.text?.style);
  if (textStyleUnsupported.length) throw new OpenXmlWasmCodecError(`Presentation shape ${shape.id} uses unsupported text-frame style fields: ${textStyleUnsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const paragraphs = shape.text?.paragraphs || [];
  if (original?.textBody && (original.textBody.paragraphs.length !== paragraphs.length || original.textBody.paragraphs.some((paragraph, index) => paragraph.runs.length !== (paragraphs[index]?.runs || []).length))) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shape.id} changed its source-bound paragraph/run topology.`, [], { code: "presentation_text_topology_changed" });
  }
  return {
    paragraphs: paragraphs.map((paragraph, index) => wireParagraph(paragraph, shape.text?.style || {}, original?.textBody?.paragraphs?.[index], shape.id)),
  };
}

function presentationShape(shape, original) {
  const originalShape = original?.content?.case === "shape" ? original.content.value : original;
  if (!new Set(["rect", "ellipse"]).has(shape.geometry)) {
    throw new OpenXmlWasmCodecError(`Presentation shape ${shape.id} uses unsupported geometry ${shape.geometry}.`, [], { code: "unsupported_presentation_features" });
  }
  const position = shape.position || {};
  const lineWidth = Number(shape.line?.width ?? 1);
  if (!Number.isFinite(lineWidth) || lineWidth < 0) throw new OpenXmlWasmCodecError(`Presentation shape ${shape.id} has an invalid line width.`, [], { code: "invalid_presentation_frame" });
  const textBody = presentationTextBody(shape, originalShape);
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
            if (entry.wire.content.case === "shape") return presentationShape(entry.model, entry.wire);
            if (opaquePresentationSnapshot(entry.model) !== entry.snapshot) throw new OpenXmlWasmCodecError(`Presentation element ${entry.model.id} is preserved but not editable by this codec slice.`, [], { code: "unsupported_presentation_edit" });
            return entry.wire;
          })
        : slide.shapes.items.map((shape) => presentationShape(shape)),
    };
  });
  return {
    protocolVersion,
    family: ArtifactFamily.PRESENTATION,
    source: state?.source,
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
  return {
    text: run.text,
    style: {
      ...(run.bold === undefined ? {} : { bold: run.bold }),
      ...(run.italic === undefined ? {} : { italic: run.italic }),
      ...(run.fontSizePoints === undefined ? {} : { fontSize: run.fontSizePoints / POINTS_PER_PIXEL }),
      ...(run.fontFamily === undefined ? {} : { fontFamily: run.fontFamily }),
      ...(run.colorRgb === undefined ? {} : { color: `#${run.colorRgb}` }),
    },
  };
}

function modelText(shape) {
  if (!shape.textBody) return shape.text;
  return shape.textBody.paragraphs.map((paragraph) => ({
    runs: paragraph.runs.map(modelRun),
    level: paragraph.level ?? 0,
    ...(paragraph.alignment ? { alignment: paragraph.alignment } : {}),
    style: {},
  }));
}

export function presentationFromEnvelope(envelope) {
  if (envelope.family !== ArtifactFamily.PRESENTATION || envelope.payload.case !== "presentation") {
    throw new OpenXmlWasmCodecError("OpenXML WebAssembly response does not contain a presentation artifact.", [], { code: "invalid_presentation_artifact" });
  }
  const source = envelope.payload.value;
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
          text: modelText(shape),
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
