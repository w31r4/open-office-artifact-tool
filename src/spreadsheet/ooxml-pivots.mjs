function decodeXml(value) {
  return String(value ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function attributes(tag = "") {
  return Object.fromEntries([...String(tag).matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])(.*?)\2/g)].map((match) => [match[1], decodeXml(match[3])]));
}

function tag(xml, localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*\\/?>`).exec(String(xml))?.[0];
}

function body(xml, localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}>`).exec(String(xml))?.[1] || "";
}

function relationshipId(attrs = {}) {
  return Object.entries(attrs).find(([name]) => /:id$/.test(name))?.[1];
}

export function parseWorkbookPivotCaches(xml = "") {
  const block = body(xml, "pivotCaches");
  return [...block.matchAll(/<(?:[A-Za-z_][\w.-]*:)?pivotCache\b[^>]*\/?>/g)].map((match) => {
    const attrs = attributes(match[0]);
    return { cacheId: Number(attrs.cacheId), relationshipId: relationshipId(attrs) };
  }).filter((item) => Number.isFinite(item.cacheId) && item.relationshipId);
}

export function parsePivotCacheDefinition(xml = "") {
  const sourceAttrs = attributes(tag(body(xml, "cacheSource"), "worksheetSource"));
  const cacheFields = body(xml, "cacheFields");
  const fields = [...cacheFields.matchAll(/<(?:[A-Za-z_][\w.-]*:)?cacheField\b[^>]*>/g)].map((match, index) => attributes(match[0]).name || `Field${index + 1}`);
  return {
    source: {
      sheet: sourceAttrs.sheet,
      ref: sourceAttrs.ref,
      name: sourceAttrs.name,
      relationshipId: relationshipId(sourceAttrs),
    },
    fields,
  };
}

function indexedFields(xml, containerName, fieldNames) {
  const container = body(xml, containerName);
  return [...container.matchAll(/<(?:[A-Za-z_][\w.-]*:)?field\b[^>]*\/?>/g)]
    .map((match) => Number(attributes(match[0]).x))
    .filter((index) => Number.isInteger(index) && index >= 0 && index < fieldNames.length)
    .map((index) => fieldNames[index]);
}

export function parsePivotTableDefinition(xml = "", cache = {}) {
  const rootAttrs = attributes(tag(xml, "pivotTableDefinition"));
  const locationAttrs = attributes(tag(xml, "location"));
  const fields = cache.fields || [];
  const dataFields = body(xml, "dataFields");
  const valueFields = [...dataFields.matchAll(/<(?:[A-Za-z_][\w.-]*:)?dataField\b[^>]*\/?>/g)].map((match) => {
    const attrs = attributes(match[0]);
    const index = Number(attrs.fld);
    return {
      field: Number.isInteger(index) && index >= 0 && index < fields.length ? fields[index] : attrs.name,
      name: attrs.name,
      summarizeBy: attrs.subtotal || "sum",
    };
  }).filter((field) => field.field);
  return {
    name: rootAttrs.name,
    cacheId: Number(rootAttrs.cacheId),
    targetRange: locationAttrs.ref,
    rowFields: indexedFields(xml, "rowFields", fields),
    columnFields: indexedFields(xml, "colFields", fields),
    valueFields,
  };
}
