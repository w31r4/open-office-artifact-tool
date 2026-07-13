import { attributes, attrEscape, decodeXml } from "./source-reference-xml.mjs";

const MAX_PICTURE_BULLET_BYTES = 8 * 1024 * 1024;
const EMUS_PER_POINT = 12_700;
const IMAGE_RELATIONSHIP_SUFFIX = "/image";

function localBlock(xml, name) {
  const pattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?${name}>`);
  return pattern.exec(String(xml || ""))?.[0];
}

function localOpening(xml, name) {
  const pattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*\\/?>`);
  return pattern.exec(String(xml || ""))?.[0] || "";
}

function attributeByLocalName(tag, name) {
  const entry = Object.entries(attributes(tag)).find(([key]) => key === name || key.endsWith(`:${name}`));
  return entry?.[1];
}

function childValue(xml, name) {
  return attributeByLocalName(localOpening(xml, name), "val");
}

function nativeIdentifier(value) {
  if (value === undefined || value === null || value === "") return undefined;
  return Number.isFinite(Number(value)) ? Number(value) : String(value);
}

function pictureBulletSource(picture) {
  return picture?.dataUrl || picture?.uri;
}

function pictureBulletKey(picture) {
  if (!picture) return "";
  return JSON.stringify([
    pictureBulletSource(picture),
    picture.widthPt,
    picture.heightPt,
    picture.alt,
  ]);
}

function pictureBulletData(dataUrl) {
  const match = /^data:(image\/(?:png|jpeg|gif));base64,([A-Za-z0-9+/=]+)$/i.exec(String(dataUrl || ""));
  if (!match) return undefined;
  const contentType = match[1].toLowerCase();
  const extension = contentType === "image/jpeg" ? "jpg" : contentType.split("/")[1];
  const bytes = Buffer.from(match[2], "base64");
  if (!bytes.length || bytes.length > MAX_PICTURE_BULLET_BYTES) throw new RangeError(`DOCX picture bullet must contain 1 through ${MAX_PICTURE_BULLET_BYTES} decoded bytes.`);
  return { contentType, extension, bytes };
}

function contentTypeFromTarget(target = "") {
  const extension = /\.([A-Za-z0-9]+)(?:[?#].*)?$/.exec(String(target))?.[1]?.toLowerCase();
  if (extension === "png") return "image/png";
  if (extension === "jpg" || extension === "jpeg") return "image/jpeg";
  if (extension === "gif") return "image/gif";
  return undefined;
}

function boundedPoints(value, fallback, name) {
  const points = Number(value ?? fallback);
  if (!Number.isFinite(points) || points < 4 || points > 72) throw new RangeError(`DOCX picture bullet ${name} must be between 4 and 72 points.`);
  return points;
}

export function normalizeDocumentPictureBullet(value) {
  if (value == null || value === false) return undefined;
  const input = typeof value === "string"
    ? (String(value).startsWith("data:") ? { dataUrl: value } : { uri: value })
    : value;
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("DOCX picture bullet must be an embedded image data URL, absolute http(s) URI, or configuration object.");
  const dataUrl = input.dataUrl || input.data || input.src;
  const uri = input.uri || input.url;
  if (Boolean(dataUrl) === Boolean(uri)) throw new TypeError("DOCX picture bullet requires exactly one of dataUrl or uri.");
  if (dataUrl && !pictureBulletData(dataUrl)) throw new TypeError("DOCX picture bullet supports embedded PNG, JPEG, or GIF base64 data URLs.");
  if (uri && !/^https?:\/\//i.test(String(uri))) throw new TypeError("DOCX external picture bullet URI must be absolute http(s).");
  const widthPt = boundedPoints(input.widthPt ?? input.sizePt ?? input.size, 12, "width");
  const heightPt = boundedPoints(input.heightPt ?? input.sizePt ?? input.size, widthPt, "height");
  const alt = String(input.alt || input.description || "Picture bullet").trim().slice(0, 255) || "Picture bullet";
  return { dataUrl: dataUrl ? String(dataUrl) : undefined, uri: uri ? String(uri) : undefined, widthPt, heightPt, alt };
}

export function parseDocxStyleNumberingPropertiesXml(xml = "") {
  const paragraphProperties = localBlock(xml, "pPr");
  const numberingProperties = localBlock(paragraphProperties, "numPr");
  if (!numberingProperties) return {};
  const numberingId = nativeIdentifier(childValue(numberingProperties, "numId"));
  const numberingLevel = nativeIdentifier(childValue(numberingProperties, "ilvl"));
  return {
    ...(numberingId === undefined ? {} : { numberingId }),
    ...(Number.isInteger(numberingLevel) && numberingLevel >= 0 && numberingLevel <= 8 ? { numberingLevel } : {}),
  };
}

function sameLevel(left, right) {
  return left.listType === right.listType
    && left.numberFormat === right.numberFormat
    && left.start === right.start
    && left.levelText === right.levelText
    && pictureBulletKey(left.pictureBullet) === pictureBulletKey(right.pictureBullet);
}

export function collectDocxNumbering(document, options = {}) {
  const groups = new Map();
  for (const block of document.blocks.filter((item) => item.kind === "listItem")) {
    if (!Number.isInteger(block.level) || block.level < 0 || block.level > 8) throw new RangeError(`DOCX list item ${block.id} level must be an integer from 0 through 8.`);
    if (!Number.isInteger(block.start) || block.start < 1) throw new RangeError(`DOCX list item ${block.id} start must be a positive integer.`);
    if (!String(block.numberFormat || "").trim()) throw new TypeError(`DOCX list item ${block.id} numberFormat must be non-empty.`);
    const pictureBullet = normalizeDocumentPictureBullet(block.pictureBullet);
    const effectiveStyle = document.styles?.effective?.(block.styleId);
    const sourceNumberingId = block.numberingId ?? effectiveStyle?.numberingId;
    const key = sourceNumberingId === undefined || sourceNumberingId === null
      ? `default:${block.listType}:${pictureBulletKey(pictureBullet) || "text"}`
      : `native:${sourceNumberingId}`;
    if (!groups.has(key)) groups.set(key, { key, blocks: [], levels: new Map(), sourceNumberingIds: new Set(), numberingStyleIds: new Set(), paragraphStyleIdsByLevel: new Map() });
    const group = groups.get(key);
    group.blocks.push(block);
    const level = block.level;
    if (sourceNumberingId !== undefined && sourceNumberingId !== null) group.sourceNumberingIds.add(String(sourceNumberingId));
    if (block.numberingStyleId) group.numberingStyleIds.add(String(block.numberingStyleId));
    if (sourceNumberingId !== undefined && sourceNumberingId !== null && String(effectiveStyle?.numberingId) === String(sourceNumberingId) && block.styleId) {
      if (!group.paragraphStyleIdsByLevel.has(level)) group.paragraphStyleIdsByLevel.set(level, new Set());
      group.paragraphStyleIdsByLevel.get(level).add(String(block.styleId));
    }
    const config = {
      listType: block.listType,
      numberFormat: block.numberFormat || (block.listType === "number" ? "decimal" : "bullet"),
      start: Number.isInteger(block.start) && block.start > 0 ? block.start : 1,
      levelText: block.levelText || (block.listType === "number" ? `%${level + 1}.` : "•"),
      pictureBullet,
    };
    const existing = group.levels.get(level);
    if (existing && !sameLevel(existing, config)) throw new Error(`DOCX numbering ${block.numberingId ?? key} level ${level} has conflicting definitions.`);
    if (!existing) group.levels.set(level, config);
  }

  const numIdByBlock = new Map();
  const numIdBySource = new Map();
  const groupValues = [...groups.values()];
  const allSourceNumberingIds = new Set(groupValues.flatMap((group) => [...group.sourceNumberingIds]));
  const definitions = groupValues.map((group, index) => {
    const numId = index + 1;
    const abstractNumId = index + 1;
    const maxLevel = Math.max(0, ...group.levels.keys());
    const fallback = group.levels.get(0) || group.levels.values().next().value || { listType: "bullet", numberFormat: "bullet", start: 1, levelText: "•" };
    const levels = Array.from({ length: maxLevel + 1 }, (_, level) => {
      const config = group.levels.get(level) || { ...fallback, levelText: fallback.listType === "number" ? `%${level + 1}.` : "•" };
      const paragraphStyleIds = group.paragraphStyleIdsByLevel.get(level);
      return { level, ...config, ...(paragraphStyleIds?.size === 1 ? { paragraphStyleId: [...paragraphStyleIds][0] } : {}) };
    });
    for (const block of group.blocks) numIdByBlock.set(block.id, numId);
    for (const sourceNumberingId of group.sourceNumberingIds) numIdBySource.set(sourceNumberingId, numId);
    if (group.numberingStyleIds.size > 1) throw new Error(`DOCX numbering ${group.key} references conflicting numbering styles.`);
    const candidateStyleLink = [...group.numberingStyleIds][0];
    const styleNumberingId = candidateStyleLink ? document.styles?.get?.(candidateStyleLink)?.numberingId : undefined;
    const ownsStyleDefinition = styleNumberingId === undefined || styleNumberingId === null || group.sourceNumberingIds.has(String(styleNumberingId)) || !allSourceNumberingIds.has(String(styleNumberingId));
    const styleLink = ownsStyleDefinition ? candidateStyleLink : undefined;
    if (styleLink) {
      if (styleNumberingId !== undefined && styleNumberingId !== null) numIdBySource.set(String(styleNumberingId), numId);
    }
    return { numId, abstractNumId, levels, styleLink };
  });

  const pictureBullets = [];
  const pictureByKey = new Map();
  const assetBySource = new Map();
  const mediaParts = [];
  const relationships = [];
  let nextMediaPartId = Math.max(1, Number(options.startMediaPartId) || 1);
  for (const level of definitions.flatMap((definition) => definition.levels)) {
    const picture = level.pictureBullet;
    const source = pictureBulletSource(picture);
    if (!source) continue;
    const key = pictureBulletKey(picture);
    let planned = pictureByKey.get(key);
    if (!planned) {
      const pictureBulletId = pictureBullets.length;
      let asset = assetBySource.get(source);
      if (!asset) {
        const relId = `rIdPictureBullet${assetBySource.size + 1}`;
        const data = picture.dataUrl ? pictureBulletData(picture.dataUrl) : undefined;
        const mediaPart = data ? { outputPath: `word/media/image${nextMediaPartId++}.${data.extension}`, ...data } : undefined;
        asset = { relId, mediaPart };
        assetBySource.set(source, asset);
        if (mediaPart) mediaParts.push(mediaPart);
        relationships.push({ id: relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: mediaPart ? mediaPart.outputPath.replace(/^word\//, "") : picture.uri, targetMode: mediaPart ? undefined : "External" });
      }
      const { relId, mediaPart } = asset;
      planned = { pictureBulletId, relId, picture, mediaPart };
      pictureBullets.push(planned);
      pictureByKey.set(key, planned);
    }
    level.pictureBulletId = planned.pictureBulletId;
  }
  return { definitions, numIdByBlock, numIdBySource, pictureBullets, mediaParts, relationships };
}

function pictureBulletXml(entry) {
  const { picture, pictureBulletId, relId } = entry;
  const objectId = pictureBulletId + 1;
  const shapeTypeId = `_x0000_t${75 + pictureBulletId}`;
  return `<w:numPicBullet w:numPicBulletId="${pictureBulletId}"><w:pict><v:shapetype id="${shapeTypeId}" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="_x0000_i${1024 + objectId}" type="#${shapeTypeId}" style="width:${picture.widthPt}pt;height:${picture.heightPt}pt" o:bullet="t"><v:imagedata r:id="${attrEscape(relId)}" o:title="${attrEscape(picture.alt)}"/></v:shape></w:pict></w:numPicBullet>`;
}

export function docxNumberingXml(numbering) {
  const pictures = numbering.pictureBullets.map(pictureBulletXml).join("");
  const abstracts = numbering.definitions.map((definition) => `<w:abstractNum w:abstractNumId="${definition.abstractNumId}"><w:multiLevelType w:val="multilevel"/>${definition.styleLink ? `<w:styleLink w:val="${attrEscape(definition.styleLink)}"/>` : ""}${definition.levels.map((level) => `<w:lvl w:ilvl="${level.level}"><w:start w:val="${level.start}"/><w:numFmt w:val="${attrEscape(level.numberFormat)}"/>${level.paragraphStyleId ? `<w:pStyle w:val="${attrEscape(level.paragraphStyleId)}"/>` : ""}<w:lvlText w:val="${attrEscape(level.levelText)}"/>${level.pictureBulletId == null ? "" : `<w:lvlPicBulletId w:val="${level.pictureBulletId}"/>`}<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${(level.level + 1) * 720}" w:hanging="360"/></w:pPr>${level.numberFormat === "bullet" && level.pictureBulletId == null ? '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>' : ""}</w:lvl>`).join("")}</w:abstractNum>`).join("");
  const instances = numbering.definitions.map((definition) => `<w:num w:numId="${definition.numId}"><w:abstractNumId w:val="${definition.abstractNumId}"/></w:num>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">${pictures}${abstracts}${instances}</w:numbering>`;
}

function parseNumberingLevel(xml, pictureById, fallbackLevel = 0) {
  const opening = localOpening(xml, "lvl");
  const level = Number(attributeByLocalName(opening, "ilvl") ?? fallbackLevel);
  const numberFormat = childValue(xml, "numFmt") || "decimal";
  const pictureBulletId = childValue(xml, "lvlPicBulletId");
  const pictureBullet = pictureBulletId == null ? undefined : pictureById.get(String(pictureBulletId));
  if (pictureBulletId != null && !pictureBullet) throw new Error(`DOCX numbering level references missing picture bullet ${pictureBulletId}.`);
  return {
    level: Number.isInteger(level) ? level : fallbackLevel,
    listType: numberFormat === "bullet" ? "bullet" : "number",
    numberFormat,
    start: Math.max(1, Number(childValue(xml, "start") || 1)),
    levelText: childValue(xml, "lvlText") || (numberFormat === "bullet" ? "•" : `%${(Number.isInteger(level) ? level : fallbackLevel) + 1}.`),
    paragraphStyleId: decodeXml(childValue(xml, "pStyle") || "") || undefined,
    pictureBullet,
  };
}

async function parsePictureBullets(xml, context) {
  const result = new Map();
  const relationships = context.relationships || [];
  const pattern = /<(?:[A-Za-z_][\w.-]*:)?numPicBullet\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?numPicBullet>/g;
  for (const match of String(xml || "").matchAll(pattern)) {
    const opening = localOpening(match[0], "numPicBullet");
    const id = attributeByLocalName(opening, "numPicBulletId");
    if (id == null || !/^-?\d+$/.test(String(id))) throw new Error("DOCX picture bullet is missing a numeric numPicBulletId.");
    if (result.has(String(id))) throw new Error(`DOCX picture bullet ID ${id} is duplicated.`);
    const blip = localOpening(match[0], "blip");
    const embeddedId = attributeByLocalName(blip, "embed");
    const linkedId = attributeByLocalName(blip, "link");
    const imageData = localOpening(match[0], "imagedata");
    const vmlId = attributeByLocalName(imageData, "id");
    const relationshipIds = [embeddedId, linkedId, vmlId].filter(Boolean);
    if (relationshipIds.length !== 1) throw new Error(`DOCX picture bullet ${id} requires exactly one DrawingML or VML image relationship.`);
    const relationshipId = relationshipIds[0];
    const relationship = relationships.find((item) => item.id === relationshipId);
    if (!relationship || !String(relationship.type || "").endsWith(IMAGE_RELATIONSHIP_SUFFIX)) throw new Error(`DOCX picture bullet ${id} references missing image relationship ${relationshipId}.`);
    const isExternal = String(relationship.targetMode || "").toLowerCase() === "external";
    if ((linkedId && !isExternal) || (embeddedId && isExternal)) throw new Error(`DOCX picture bullet ${id} relationship ${relationshipId} has the wrong target mode.`);
    let dataUrl;
    let uri;
    if (isExternal) {
      if (!/^https?:\/\//i.test(String(relationship.target || ""))) throw new Error(`DOCX picture bullet ${id} external relationship must use an absolute http(s) URI.`);
      uri = relationship.target;
    } else {
      const target = context.resolveTarget?.(context.partPath || "word/numbering.xml", relationship.target);
      const contentType = contentTypeFromTarget(target || relationship.target);
      if (!target || !contentType) throw new Error(`DOCX picture bullet ${id} relationship ${relationshipId} has an unsupported image target.`);
      const bytes = await context.readPart?.(target);
      if (!bytes?.length) throw new Error(`DOCX picture bullet ${id} relationship ${relationshipId} targets missing part ${target}.`);
      if (bytes.length > (context.maxPictureBulletBytes || MAX_PICTURE_BULLET_BYTES)) throw new RangeError(`DOCX picture bullet ${id} exceeds the decoded byte budget.`);
      dataUrl = `data:${contentType};base64,${Buffer.from(bytes).toString("base64")}`;
    }
    const shape = localOpening(match[0], "shape");
    const style = attributeByLocalName(shape, "style") || "";
    const widthStyle = /(?:^|;)\s*width\s*:\s*([0-9.]+)pt(?:;|$)/i.exec(style)?.[1];
    const heightStyle = /(?:^|;)\s*height\s*:\s*([0-9.]+)pt(?:;|$)/i.exec(style)?.[1];
    const extent = attributes(localOpening(match[0], "extent"));
    const widthPt = boundedPoints(widthStyle ?? (Number(extent.cx || 12 * EMUS_PER_POINT) / EMUS_PER_POINT), 12, "width");
    const heightPt = boundedPoints(heightStyle ?? (Number(extent.cy || widthPt * EMUS_PER_POINT) / EMUS_PER_POINT), widthPt, "height");
    const docPr = localOpening(match[0], "docPr");
    const alt = decodeXml(attributeByLocalName(docPr, "descr") || attributeByLocalName(docPr, "title") || attributeByLocalName(imageData, "title") || "Picture bullet");
    result.set(String(id), normalizeDocumentPictureBullet({ dataUrl, uri, widthPt, heightPt, alt }));
  }
  return result;
}

export async function parseDocxNumberingXml(xml = "", context = {}) {
  const pictureById = await parsePictureBullets(xml, context);
  const abstracts = new Map();
  const abstractPattern = /<(?:[A-Za-z_][\w.-]*:)?abstractNum\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?abstractNum>/g;
  for (const match of String(xml || "").matchAll(abstractPattern)) {
    const abstractNumId = attributeByLocalName(localOpening(match[0], "abstractNum"), "abstractNumId");
    if (abstractNumId === undefined) continue;
    if (abstracts.has(String(abstractNumId))) throw new Error(`DOCX abstract numbering ID ${abstractNumId} is duplicated.`);
    const levels = new Map();
    const levelPattern = /<(?:[A-Za-z_][\w.-]*:)?lvl\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?lvl>/g;
    for (const levelMatch of match[0].matchAll(levelPattern)) {
      const level = parseNumberingLevel(levelMatch[0], pictureById);
      if (levels.has(level.level)) throw new Error(`DOCX abstract numbering ${abstractNumId} level ${level.level} is duplicated.`);
      levels.set(level.level, level);
    }
    const styleLink = decodeXml(childValue(match[0], "styleLink") || "") || undefined;
    const numberingStyleLink = decodeXml(childValue(match[0], "numStyleLink") || "") || undefined;
    if ((styleLink?.length || 0) > 253 || (numberingStyleLink?.length || 0) > 253) throw new RangeError(`DOCX abstract numbering ${abstractNumId} style link exceeds 253 characters.`);
    abstracts.set(String(abstractNumId), { levels, styleLink, numberingStyleLink });
  }

  const rawInstances = new Map();
  const instancePattern = /<(?:[A-Za-z_][\w.-]*:)?num\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?num>/g;
  for (const match of String(xml || "").matchAll(instancePattern)) {
    const numId = attributeByLocalName(localOpening(match[0], "num"), "numId");
    const abstractNumId = childValue(match[0], "abstractNumId");
    if (numId === undefined || abstractNumId === undefined) continue;
    if (rawInstances.has(String(numId))) throw new Error(`DOCX numbering instance ID ${numId} is duplicated.`);
    const overrides = new Map();
    const overridePattern = /<(?:[A-Za-z_][\w.-]*:)?lvlOverride\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?lvlOverride>/g;
    for (const overrideMatch of match[0].matchAll(overridePattern)) {
      const overrideLevel = Number(attributeByLocalName(localOpening(overrideMatch[0], "lvlOverride"), "ilvl") ?? 0);
      if (!Number.isInteger(overrideLevel) || overrideLevel < 0 || overrideLevel > 8) throw new RangeError(`DOCX numbering ${numId} override level must be from 0 through 8.`);
      if (overrides.has(overrideLevel)) throw new Error(`DOCX numbering ${numId} override level ${overrideLevel} is duplicated.`);
      const nestedLevelXml = localBlock(overrideMatch[0], "lvl");
      if (nestedLevelXml) overrides.set(overrideLevel, { level: parseNumberingLevel(nestedLevelXml, pictureById, overrideLevel) });
      else if (childValue(overrideMatch[0], "startOverride") !== undefined) {
        overrides.set(overrideLevel, { start: Math.max(1, Number(childValue(overrideMatch[0], "startOverride") || 1)) });
      } else overrides.set(overrideLevel, {});
    }
    rawInstances.set(String(numId), { numberingId: nativeIdentifier(numId), abstractNumberingId: nativeIdentifier(abstractNumId), overrides });
  }

  const styleById = (id) => context.styles?.get?.(id) || context.styles?.[id];
  const instances = new Map();
  const resolveInstance = (numId, trail = []) => {
    const key = String(numId);
    if (instances.has(key)) return instances.get(key);
    if (trail.includes(key)) throw new Error(`DOCX numbering style link cycle: ${[...trail, key].join(" -> ")}.`);
    const raw = rawInstances.get(key);
    if (!raw) throw new Error(`DOCX numbering style references missing numbering instance ${key}.`);
    const abstract = abstracts.get(String(raw.abstractNumberingId));
    let levels = new Map();
    let numberingStyleId = abstract?.styleLink;
    if (abstract?.numberingStyleLink) {
      const linkedStyle = styleById(abstract.numberingStyleLink);
      if (!linkedStyle || linkedStyle.type !== "numbering") throw new Error(`DOCX abstract numbering ${raw.abstractNumberingId} references missing numbering style ${abstract.numberingStyleLink}.`);
      if (linkedStyle.numberingId === undefined || linkedStyle.numberingId === null) throw new Error(`DOCX numbering style ${abstract.numberingStyleLink} is missing numPr/numId.`);
      const linked = resolveInstance(linkedStyle.numberingId, [...trail, key]);
      levels = new Map([...linked.levels].map(([level, config]) => [level, { ...config }]));
      numberingStyleId = abstract.numberingStyleLink;
    } else if (abstract) levels = new Map([...abstract.levels].map(([level, config]) => [level, { ...config }]));
    for (const [level, override] of raw.overrides) {
      if (override.level) levels.set(level, { ...override.level, level });
      else if (override.start !== undefined) {
        const current = levels.get(level) || { level, listType: "number", numberFormat: "decimal", start: 1, levelText: `%${level + 1}.` };
        levels.set(level, { ...current, start: override.start });
      }
    }
    const resolved = { numberingId: raw.numberingId, abstractNumberingId: raw.abstractNumberingId, levels, numberingStyleId };
    instances.set(key, resolved);
    return resolved;
  };
  for (const numId of rawInstances.keys()) resolveInstance(numId);
  return instances;
}

export function resolveDocxParagraphNumbering({ styleId = "Normal", directNumberingId, directLevel, styles, numberingById = new Map() } = {}) {
  const cascade = styles?.cascade?.(styleId) || [];
  const direct = directNumberingId !== undefined && directNumberingId !== null;
  const styleWithNumbering = direct ? undefined : [...cascade].reverse().find((style) => style.numberingId !== undefined && style.numberingId !== null);
  const numberingId = direct ? nativeIdentifier(directNumberingId) : nativeIdentifier(styleWithNumbering?.numberingId);
  if (numberingId === undefined || String(numberingId) === "0") return undefined;
  const numbering = numberingById.get(String(numberingId));
  let level = direct && Number.isInteger(Number(directLevel)) ? Number(directLevel) : 0;
  let paragraphStyleId;
  if (!direct && numbering) {
    for (const candidate of [...cascade].reverse()) {
      const linkedLevel = [...numbering.levels.values()].find((entry) => entry.paragraphStyleId === candidate.id);
      if (linkedLevel) {
        level = linkedLevel.level;
        paragraphStyleId = candidate.id;
        break;
      }
    }
  }
  if (!Number.isInteger(level) || level < 0 || level > 8) level = 0;
  return { numberingId, level, numbering, inheritedFromStyle: !direct, paragraphStyleId };
}
