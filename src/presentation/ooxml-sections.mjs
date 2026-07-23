import { deterministicPresentationGuid } from "./ooxml-modern-comments.mjs";

export const PPTX_SECTION_EXTENSION_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";

const MAX_SECTIONS = 4096;
const MAX_SECTION_SLIDES = 16384;
const MAX_SECTION_NAME = 255;
const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;

function normalizeName(value) {
  const name = String(value || "").trim();
  if (!name || name.length > MAX_SECTION_NAME) throw new RangeError(`Presentation section name must contain 1 through ${MAX_SECTION_NAME} characters.`);
  if (/[\u0000-\u001f\u007f]/.test(name)) throw new TypeError("Presentation section name must not contain control characters.");
  return name;
}

function normalizeNativeId(value) {
  if (value == null) return undefined;
  const nativeId = String(value).trim().toUpperCase();
  if (!GUID.test(nativeId)) throw new TypeError("Presentation section nativeId must be a brace-delimited GUID.");
  return nativeId;
}

function normalizeSlideId(presentation, value) {
  const slide = typeof value === "number"
    ? presentation.slides.getItem(value)
    : value && typeof value === "object"
      ? value
      : presentation.slides.items.find((candidate) => candidate.id === String(value));
  if (!slide || slide.presentation !== presentation || !presentation.slides.items.includes(slide)) {
    throw new Error(`Presentation section references missing slide ${value?.id || value}.`);
  }
  return slide.id;
}

function normalizeSlideIds(presentation, value) {
  if (!Array.isArray(value) || value.length < 1) throw new RangeError("Presentation section requires at least one slide.");
  if (value.length > MAX_SECTION_SLIDES) throw new RangeError(`Presentation section exceeds ${MAX_SECTION_SLIDES} slides.`);
  return value.map((slide) => normalizeSlideId(presentation, slide));
}

export class PresentationSection {
  constructor(presentation, config = {}, modelId) {
    this.presentation = presentation;
    this.id = String(config.id || modelId);
    this.name = normalizeName(config.name);
    this.nativeId = normalizeNativeId(config.nativeId);
    this.slideIds = normalizeSlideIds(presentation, config.slideIds || config.slides);
  }

  setSlides(slides) {
    this.slideIds = normalizeSlideIds(this.presentation, slides);
    return this;
  }

  get slides() { return this.slideIds.map((id) => this.presentation.slides.items.find((slide) => slide.id === id)); }
  inspectRecord() { return { kind: "section", id: this.id, name: this.name, nativeId: this.nativeId, slideIds: [...this.slideIds], slides: this.slideIds.length }; }
  toJSON() { return { id: this.id, name: this.name, nativeId: this.nativeId, slideIds: [...this.slideIds] }; }
}

export class PresentationSectionCollection {
  constructor(presentation) {
    this.presentation = presentation;
    this.items = [];
  }

  add(nameOrConfig, slides) {
    if (this.items.length >= MAX_SECTIONS) throw new RangeError(`Presentation sections exceed ${MAX_SECTIONS} entries.`);
    const config = typeof nameOrConfig === "string" ? { name: nameOrConfig, slides } : nameOrConfig;
    if (!config || typeof config !== "object" || Array.isArray(config)) throw new TypeError("Presentation section must be a name/config object plus slides.");
    const name = normalizeName(config.name);
    if (this.items.some((section) => section.name.toLowerCase() === name.toLowerCase())) throw new Error(`Presentation section name ${name} already exists.`);
    const section = new PresentationSection(this.presentation, { ...config, name }, `section/${this.items.length + 1}`);
    if (this.items.some((candidate) => candidate.id === section.id)) throw new Error(`Presentation section ID ${section.id} already exists.`);
    if (section.nativeId && this.items.some((candidate) => candidate.nativeId === section.nativeId)) throw new Error(`Presentation section nativeId ${section.nativeId} already exists.`);
    this.items.push(section);
    return section;
  }

  getItem(idOrNameOrIndex) {
    if (typeof idOrNameOrIndex === "number") return this.items[idOrNameOrIndex];
    const value = String(idOrNameOrIndex || "");
    return this.items.find((section) => section.id === value || section.name === value);
  }

  get count() { return this.items.length; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

// A section is a partition, not an arbitrary slide set. Holding the invariant
// here means every caller (JS model, protobuf adapter, inspect, and export)
// gets one shared answer instead of each layer guessing where a newly inserted
// or moved slide belongs.
export function planPresentationSections(presentation, options = {}) {
  const entries = presentation.sections.items.map((section) => {
    section.name = normalizeName(section.name);
    section.slideIds = normalizeSlideIds(presentation, section.slideIds);
    section.nativeId = normalizeNativeId(section.nativeId || deterministicPresentationGuid(`section:${section.id}`));
    return section;
  });
  const ids = new Set();
  const names = new Set();
  const nativeIds = new Set();
  for (const section of entries) {
    if (!section.id || section.id.length > 1024 || ids.has(section.id)) throw new Error(`Presentation section ID ${section.id || "(missing)"} must be non-empty, bounded, and unique.`);
    ids.add(section.id);
    const name = section.name.toLowerCase();
    if (names.has(name)) throw new Error(`Presentation section name ${section.name} already exists.`);
    names.add(name);
    if (nativeIds.has(section.nativeId)) throw new Error(`Presentation section nativeId ${section.nativeId} already exists.`);
    nativeIds.add(section.nativeId);
  }
  if (entries.length && !options.allowPendingClone) {
    const expected = presentation.slides.items.map((slide) => slide.id);
    const actual = entries.flatMap((section) => section.slideIds);
    if (actual.length !== expected.length || actual.some((slideId, index) => slideId !== expected[index])) {
      throw new Error("Presentation sections must partition every slide exactly once and in presentation order. Update section membership explicitly after inserting, moving, or removing slides.");
    }
  }
  return { entries };
}
