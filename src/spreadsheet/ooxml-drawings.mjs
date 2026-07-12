// Pure SpreadsheetDrawingML/chart parsing. Package relationships and workbook objects stay in the model layer.

function decodeXml(value) {
  return String(value ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function attributes(tag = "") {
  return Object.fromEntries([...String(tag).matchAll(/\b([A-Za-z_:][\w:.-]*)="([^"]*)"/g)].map((match) => [match[1], decodeXml(match[2])]));
}

function elementBody(xml, name) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${name}>`).exec(String(xml))?.[1] || "";
}

function elementValue(xml, name) {
  return decodeXml(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${name}>`).exec(String(xml))?.[1] || "");
}

function marker(anchorXml, name) {
  const body = elementBody(anchorXml, name);
  if (!body) return undefined;
  return {
    col: Number(elementValue(body, "col") || 0),
    row: Number(elementValue(body, "row") || 0),
    colOffsetPx: Number(elementValue(body, "colOff") || 0) / 9525,
    rowOffsetPx: Number(elementValue(body, "rowOff") || 0) / 9525,
  };
}

function drawingExtent(anchorXml) {
  const tag = /<(?:[A-Za-z_][\w.-]*:)?ext\b[^>]*\b(?:cx|cy)="[^"]+"[^>]*\/?\s*>/.exec(anchorXml)?.[0];
  const attrs = attributes(tag);
  return tag ? { widthPx: Number(attrs.cx || 0) / 9525, heightPx: Number(attrs.cy || 0) / 9525 } : undefined;
}

function absolutePosition(anchorXml) {
  const tag = /<(?:[A-Za-z_][\w.-]*:)?pos\b[^>]*\b(?:x|y)="[^"]+"[^>]*\/?\s*>/.exec(anchorXml)?.[0];
  const attrs = attributes(tag);
  return tag ? { leftPx: Number(attrs.x || 0) / 9525, topPx: Number(attrs.y || 0) / 9525 } : undefined;
}

export function parseSpreadsheetDrawing(xml = "") {
  const records = [];
  const anchorPattern = /<(?:[A-Za-z_][\w.-]*:)?(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?\1>/g;
  for (const match of String(xml).matchAll(anchorPattern)) {
    const body = match[2];
    const picture = /<(?:[A-Za-z_][\w.-]*:)?pic\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?pic>/.exec(body)?.[0];
    const graphicFrame = /<(?:[A-Za-z_][\w.-]*:)?graphicFrame\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?graphicFrame>/.exec(body)?.[0];
    const objectXml = picture || graphicFrame;
    if (!objectXml) continue;
    const propertiesTag = /<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*\/?\s*>/.exec(objectXml)?.[0];
    const properties = attributes(propertiesTag);
    const relationship = picture
      ? /<(?:[A-Za-z_][\w.-]*:)?blip\b[^>]*\b(?:[A-Za-z_][\w.-]*:)?embed="([^"]+)"/.exec(objectXml)?.[1]
      : /<(?:[A-Za-z_][\w.-]*:)?chart\b[^>]*\b(?:[A-Za-z_][\w.-]*:)?id="([^"]+)"/.exec(objectXml)?.[1];
    if (!relationship) continue;
    records.push({
      kind: picture ? "image" : "chart",
      relationshipId: relationship,
      name: properties.name,
      alt: properties.descr || properties.title,
      anchorType: match[1],
      from: marker(body, "from"),
      to: marker(body, "to"),
      position: absolutePosition(body),
      extent: drawingExtent(body),
    });
  }
  return records;
}

function chartPoints(xml = "", numeric = false) {
  return [...String(xml).matchAll(/<(?:[A-Za-z_][\w.-]*:)?pt\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?pt>/g)]
    .map((match, position) => ({ index: Number(attributes(match[1]).idx ?? position), value: elementValue(match[2], "v") }))
    .sort((left, right) => left.index - right.index)
    .map((point) => numeric ? Number(point.value) || 0 : point.value);
}

export function parseSpreadsheetChart(xml = "") {
  const text = String(xml || "");
  const type = /<(?:[A-Za-z_][\w.-]*:)?pieChart\b/.test(text) ? "pie" : /<(?:[A-Za-z_][\w.-]*:)?lineChart\b/.test(text) ? "line" : "bar";
  const titleBody = elementBody(text, "title");
  const title = [...titleBody.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g)].map((match) => decodeXml(match[1])).join("") || elementValue(titleBody, "v");
  const series = [...text.matchAll(/<(?:[A-Za-z_][\w.-]*:)?ser\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?ser>/g)].map((match, index) => {
    const body = match[1];
    const tx = elementBody(body, "tx");
    const categoryBody = elementBody(body, "cat") || elementBody(body, "xVal");
    const valueBody = elementBody(body, "val") || elementBody(body, "yVal");
    const color = /<(?:[A-Za-z_][\w.-]*:)?srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(body)?.[1];
    return {
      name: elementValue(tx, "v") || elementValue(tx, "f") || `Series ${index + 1}`,
      categoryFormula: elementValue(categoryBody, "f") || undefined,
      formula: elementValue(valueBody, "f") || undefined,
      categories: chartPoints(categoryBody),
      values: chartPoints(valueBody, true),
      fill: color ? `#${color.toUpperCase()}` : undefined,
    };
  });
  return { type, title, hasLegend: /<(?:[A-Za-z_][\w.-]*:)?legend\b/.test(text), categories: series[0]?.categories || [], series };
}
