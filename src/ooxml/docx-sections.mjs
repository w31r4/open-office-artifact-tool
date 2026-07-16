const REFERENCE_TYPES = ["default", "first", "even"];
const HEADER_FOOTER_KINDS = ["header", "footer"];

export function docxSectionCount(blocks = []) {
  return blocks.filter((block) => block?.kind === "section").length + 1;
}

export function normalizeDocxSectionSettings(value = [], blocksOrCount = 1) {
  const sectionCount = Array.isArray(blocksOrCount) ? docxSectionCount(blocksOrCount) : Math.max(1, Number(blocksOrCount) || 1);
  const source = Array.isArray(value) ? value : Object.entries(value || {}).map(([sectionIndex, settings]) => ({ sectionIndex: Number(sectionIndex), ...(settings || {}) }));
  const byIndex = new Map();
  for (const [position, entry] of source.entries()) {
    const sectionIndex = Number(entry?.sectionIndex ?? entry?.index ?? position);
    if (!Number.isInteger(sectionIndex) || sectionIndex < 0 || sectionIndex >= sectionCount) throw new RangeError(`DOCX section settings index ${sectionIndex} must be an integer from 0 through ${sectionCount - 1}.`);
    const rawDifferentFirstPage = entry?.differentFirstPage ?? entry?.titlePage;
    const settings = { sectionIndex };
    if (rawDifferentFirstPage !== undefined) settings.differentFirstPage = Boolean(rawDifferentFirstPage);
    byIndex.set(sectionIndex, settings);
  }
  return [...byIndex.values()].sort((left, right) => left.sectionIndex - right.sectionIndex);
}

function headerFooterGroups(document, kind) {
  const blocks = kind === "header" ? document.headers : document.footers;
  const finalSectionIndex = docxSectionCount(document.blocks) - 1;
  const groups = new Map();
  for (const block of blocks) {
    const referenceType = REFERENCE_TYPES.includes(block.referenceType) ? block.referenceType : "default";
    const sectionIndex = block.sectionIndex === undefined ? finalSectionIndex : Number(block.sectionIndex);
    if (!Number.isInteger(sectionIndex) || sectionIndex < 0 || sectionIndex > finalSectionIndex) throw new RangeError(`DOCX ${kind} ${block.id} sectionIndex must be an integer from 0 through ${finalSectionIndex}.`);
    const key = `${sectionIndex}\u0000${referenceType}`;
    if (!groups.has(key)) groups.set(key, { kind, sectionIndex, referenceType, blocks: [] });
    groups.get(key).blocks.push(block);
  }
  return groups;
}

function emptySlot() {
  return { blocks: [], ids: [], sourceSectionIndex: undefined, inherited: false };
}

function explicitSlot(groups, sectionIndex, referenceType) {
  const blocks = groups.get(`${sectionIndex}\u0000${referenceType}`)?.blocks || [];
  return blocks.length ? { blocks, ids: blocks.map((block) => block.id), sourceSectionIndex: sectionIndex, inherited: false } : emptySlot();
}

function inheritedSlot(explicit, previous, sectionIndex) {
  if (explicit.blocks.length) return explicit;
  if (!previous?.blocks.length) return emptySlot();
  return { blocks: previous.blocks, ids: previous.ids, sourceSectionIndex: previous.sourceSectionIndex, inherited: previous.sourceSectionIndex !== sectionIndex };
}

export function planDocxHeaderFooterSections(document) {
  const sectionCount = docxSectionCount(document.blocks);
  const groups = Object.fromEntries(HEADER_FOOTER_KINDS.map((kind) => [kind, headerFooterGroups(document, kind)]));
  const declaredSettings = new Map(normalizeDocxSectionSettings(document.sectionSettings || [], sectionCount).map((settings) => [settings.sectionIndex, settings]));
  const sections = [];
  for (let sectionIndex = 0; sectionIndex < sectionCount; sectionIndex += 1) {
    const explicit = {};
    const effective = {};
    for (const kind of HEADER_FOOTER_KINDS) {
      explicit[kind] = Object.fromEntries(REFERENCE_TYPES.map((referenceType) => [referenceType, explicitSlot(groups[kind], sectionIndex, referenceType)]));
      effective[kind] = Object.fromEntries(REFERENCE_TYPES.map((referenceType) => [referenceType, inheritedSlot(explicit[kind][referenceType], sections[sectionIndex - 1]?.effective?.[kind]?.[referenceType], sectionIndex)]));
    }
    const declared = declaredSettings.get(sectionIndex);
    const inferredFirstPage = [...explicit.header.first.blocks, ...explicit.footer.first.blocks].some((block) => block.variantActive !== false);
    sections.push({
      sectionIndex,
      differentFirstPage: declared?.differentFirstPage ?? inferredFirstPage,
      evenAndOddHeaders: Boolean(document.settings?.evenAndOddHeaders),
      explicit,
      effective,
    });
  }
  return { sectionCount, sections, groups };
}

function pageVariant(section, pageInSection) {
  if (section.differentFirstPage && pageInSection === 1) return "first";
  if (section.evenAndOddHeaders && pageInSection % 2 === 0) return "even";
  return "default";
}

function pageSlot(slot, referenceType) {
  return {
    referenceType,
    ids: slot.ids,
    sourceSectionIndex: slot.sourceSectionIndex,
    inherited: slot.inherited,
    blank: slot.ids.length === 0,
  };
}

export function resolveDocxPageHeaderFooter(plan, sectionIndex, pageInSection) {
  const section = plan.sections[sectionIndex];
  if (!section) throw new RangeError(`DOCX page section index ${sectionIndex} must be an integer from 0 through ${plan.sectionCount - 1}.`);
  const sectionPage = Math.max(1, Number(pageInSection) || 1);
  const referenceType = pageVariant(section, sectionPage);
  const header = pageSlot(section.effective.header[referenceType], referenceType);
  const footer = pageSlot(section.effective.footer[referenceType], referenceType);
  return {
    sectionIndex,
    pageInSection: sectionPage,
    differentFirstPage: section.differentFirstPage,
    evenAndOddHeaders: section.evenAndOddHeaders,
    headers: header.ids,
    footers: footer.ids,
    header,
    footer,
  };
}
