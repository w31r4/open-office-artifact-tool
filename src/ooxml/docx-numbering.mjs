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
    const key = block.numberingId === undefined || block.numberingId === null
      ? `default:${block.listType}:${pictureBulletKey(pictureBullet) || "text"}`
      : `native:${block.numberingId}`;
    if (!groups.has(key)) groups.set(key, { key, blocks: [], levels: new Map() });
    const group = groups.get(key);
    group.blocks.push(block);
    const level = block.level;
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
  const definitions = [...groups.values()].map((group, index) => {
    const numId = index + 1;
    const abstractNumId = index + 1;
    const maxLevel = Math.max(0, ...group.levels.keys());
    const fallback = group.levels.get(0) || group.levels.values().next().value || { listType: "bullet", numberFormat: "bullet", start: 1, levelText: "•" };
    const levels = Array.from({ length: maxLevel + 1 }, (_, level) => {
      const config = group.levels.get(level) || { ...fallback, levelText: fallback.listType === "number" ? `%${level + 1}.` : "•" };
      return { level, ...config };
    });
    for (const block of group.blocks) numIdByBlock.set(block.id, numId);
    return { numId, abstractNumId, levels };
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
  return { definitions, numIdByBlock, pictureBullets, mediaParts, relationships };
}

function pictureBulletXml(entry) {
  const { picture, pictureBulletId, relId } = entry;
  const objectId = pictureBulletId + 1;
  const shapeTypeId = `_x0000_t${75 + pictureBulletId}`;
  return `<w:numPicBullet w:numPicBulletId="${pictureBulletId}"><w:pict><v:shapetype id="${shapeTypeId}" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="_x0000_i${1024 + objectId}" type="#${shapeTypeId}" style="width:${picture.widthPt}pt;height:${picture.heightPt}pt" o:bullet="t"><v:imagedata r:id="${attrEscape(relId)}" o:title="${attrEscape(picture.alt)}"/></v:shape></w:pict></w:numPicBullet>`;
}

export function docxNumberingXml(numbering) {
  const pictures = numbering.pictureBullets.map(pictureBulletXml).join("");
  const abstracts = numbering.definitions.map((definition) => `<w:abstractNum w:abstractNumId="${definition.abstractNumId}"><w:multiLevelType w:val="multilevel"/>${definition.levels.map((level) => `<w:lvl w:ilvl="${level.level}"><w:start w:val="${level.start}"/><w:numFmt w:val="${attrEscape(level.numberFormat)}"/><w:lvlText w:val="${attrEscape(level.levelText)}"/>${level.pictureBulletId == null ? "" : `<w:lvlPicBulletId w:val="${level.pictureBulletId}"/>`}<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${(level.level + 1) * 720}" w:hanging="360"/></w:pPr>${level.numberFormat === "bullet" && level.pictureBulletId == null ? '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>' : ""}</w:lvl>`).join("")}</w:abstractNum>`).join("");
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
    const levels = new Map();
    const levelPattern = /<(?:[A-Za-z_][\w.-]*:)?lvl\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?lvl>/g;
    for (const levelMatch of match[0].matchAll(levelPattern)) {
      const level = parseNumberingLevel(levelMatch[0], pictureById);
      levels.set(level.level, level);
    }
    abstracts.set(String(abstractNumId), levels);
  }

  const instances = new Map();
  const instancePattern = /<(?:[A-Za-z_][\w.-]*:)?num\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?num>/g;
  for (const match of String(xml || "").matchAll(instancePattern)) {
    const numId = attributeByLocalName(localOpening(match[0], "num"), "numId");
    const abstractNumId = childValue(match[0], "abstractNumId");
    if (numId === undefined || abstractNumId === undefined) continue;
    const levels = new Map([...(abstracts.get(String(abstractNumId)) || new Map()).entries()].map(([level, config]) => [level, { ...config }]));
    const overridePattern = /<(?:[A-Za-z_][\w.-]*:)?lvlOverride\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?lvlOverride>/g;
    for (const overrideMatch of match[0].matchAll(overridePattern)) {
      const overrideLevel = Number(attributeByLocalName(localOpening(overrideMatch[0], "lvlOverride"), "ilvl") ?? 0);
      const nestedLevelXml = localBlock(overrideMatch[0], "lvl");
      if (nestedLevelXml) levels.set(overrideLevel, parseNumberingLevel(nestedLevelXml, pictureById, overrideLevel));
      else if (childValue(overrideMatch[0], "startOverride") !== undefined) {
        const current = levels.get(overrideLevel) || { level: overrideLevel, listType: "number", numberFormat: "decimal", start: 1, levelText: `%${overrideLevel + 1}.` };
        levels.set(overrideLevel, { ...current, start: Math.max(1, Number(childValue(overrideMatch[0], "startOverride") || 1)) });
      }
    }
    instances.set(String(numId), { numberingId: Number.isFinite(Number(numId)) ? Number(numId) : numId, abstractNumberingId: Number.isFinite(Number(abstractNumId)) ? Number(abstractNumId) : abstractNumId, levels });
  }
  return instances;
}
