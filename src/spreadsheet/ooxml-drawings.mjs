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

function directElementBody(xml, name) {
  const source = String(xml ?? "");
  const tags = /<\s*(\/)?(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)\b[^>]*>/g;
  let depth = 0;
  let start = -1;
  let match;
  while ((match = tags.exec(source))) {
    if (match[1]) {
      depth -= 1;
      if (start >= 0 && depth === 0 && match[2] === name) return source.slice(start, match.index);
      continue;
    }
    const selfClosing = /\/\s*>$/.test(match[0]);
    if (depth === 0 && match[2] === name) {
      if (selfClosing) return "";
      start = tags.lastIndex;
    }
    if (!selfClosing) depth += 1;
  }
  return "";
}

function directElement(xml, name) {
  const source = String(xml ?? "");
  const tags = /<\s*(\/)?(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)\b[^>]*>/g;
  let depth = 0;
  let start = -1;
  let bodyStart = -1;
  let startTag = "";
  let match;
  while ((match = tags.exec(source))) {
    if (match[1]) {
      depth -= 1;
      if (start >= 0 && depth === 0 && match[2] === name) return { startTag, body: source.slice(bodyStart, match.index), xml: source.slice(start, tags.lastIndex) };
      continue;
    }
    const selfClosing = /\/\s*>$/.test(match[0]);
    if (depth === 0 && match[2] === name) {
      if (selfClosing) return { startTag: match[0], body: "", xml: match[0] };
      start = match.index;
      bodyStart = tags.lastIndex;
      startTag = match[0];
    }
    if (!selfClosing) depth += 1;
  }
  return undefined;
}

function directChildNames(xml) {
  const source = String(xml ?? "");
  const tags = /<\s*(\/)?(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)\b[^>]*>/g;
  const names = [];
  let depth = 0;
  let match;
  while ((match = tags.exec(source))) {
    if (match[1]) { depth -= 1; continue; }
    const selfClosing = /\/\s*>$/.test(match[0]);
    if (depth === 0) names.push(match[2]);
    if (!selfClosing) depth += 1;
  }
  return names;
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

function pictureCrop(pictureXml) {
  const tag = /<(?:[A-Za-z_][\w.-]*:)?srcRect\b[^>]*\/?\s*>/.exec(pictureXml)?.[0];
  if (!tag) return undefined;
  const attrs = attributes(tag);
  const values = [attrs.l, attrs.t, attrs.r, attrs.b].map((value) => Number(value || 0));
  if (values.some((value) => !Number.isFinite(value) || !Number.isInteger(value) || value < -100_000 || value > 100_000) ||
      values[0] + values[2] >= 100_000 || values[1] + values[3] >= 100_000) return undefined;
  return {
    leftPercent: values[0] / 1000,
    topPercent: values[1] / 1000,
    rightPercent: values[2] / 1000,
    bottomPercent: values[3] / 1000,
  };
}

function pictureEffects(pictureXml) {
  const body = /<(?:[A-Za-z_][\w.-]*:)?blip\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?blip>/.exec(pictureXml)?.[1];
  if (body == null) return undefined;
  const supported = /<(?:[A-Za-z_][\w.-]*:)?(?:alphaModFix|grayscl|lum)\b[^>]*\/?\s*>/g;
  const tags = [...body.matchAll(supported)].map((match) => match[0]);
  if (body.replace(supported, "").trim() || tags.length === 0) return undefined;
  const hasLocalName = (tag, name) => new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b`).test(tag);
  const alpha = tags.filter((tag) => hasLocalName(tag, "alphaModFix"));
  const grayscale = tags.filter((tag) => hasLocalName(tag, "grayscl"));
  const luminance = tags.filter((tag) => hasLocalName(tag, "lum"));
  if (alpha.length > 1 || grayscale.length > 1 || luminance.length > 1) return undefined;
  const output = {};
  if (grayscale.length === 1) output.grayscale = true;
  if (luminance.length === 1) {
    const attrs = attributes(luminance[0]);
    const brightness = Number(attrs.bright || 0);
    const contrast = Number(attrs.contrast || 0);
    if (![brightness, contrast].every((value) => Number.isInteger(value) && value >= -100_000 && value <= 100_000)) return undefined;
    output.brightnessPercent = brightness / 1000;
    output.contrastPercent = contrast / 1000;
  }
  if (alpha.length === 1) {
    const amount = Number(attributes(alpha[0]).amt);
    if (!Number.isInteger(amount) || amount < 0 || amount > 100_000) return undefined;
    output.opacityPercent = amount / 1000;
  }
  return Object.keys(output).length > 0 ? output : undefined;
}

function pictureTransform(pictureXml) {
  const shapeProperties = elementBody(pictureXml, "spPr");
  if (!shapeProperties) return undefined;
  const tags = [...shapeProperties.matchAll(/<(?:[A-Za-z_][\w.-]*:)?xfrm\b[^>]*>/g)].map((match) => match[0]);
  if (tags.length !== 1) return undefined;
  const attrs = attributes(tags[0]);
  if (Object.keys(attrs).some((name) => !["rot", "flipH", "flipV"].includes(name) && !name.startsWith("xmlns"))) return undefined;
  const output = {};
  if (attrs.rot != null) {
    const angle = Number(attrs.rot);
    if (!Number.isInteger(angle) || angle < -21_600_000 || angle > 21_600_000) return undefined;
    output.rotationDegrees = angle / 60_000;
  }
  const booleanAttribute = (name, publicName) => {
    if (attrs[name] == null) return true;
    if (!["0", "1", "false", "true"].includes(attrs[name])) return false;
    output[publicName] = attrs[name] === "1" || attrs[name] === "true";
    return true;
  };
  if (!booleanAttribute("flipH", "flipHorizontal") || !booleanAttribute("flipV", "flipVertical")) return undefined;
  return Object.keys(output).length > 0 ? output : undefined;
}

export function parseSpreadsheetDrawing(xml = "") {
  const records = [];
  const anchorPattern = /<(?:[A-Za-z_][\w.-]*:)?(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?\1>/g;
  for (const match of String(xml).matchAll(anchorPattern)) {
    const body = match[2];
    const anchorAttributes = attributes(/^<[^>]+>/.exec(match[0])?.[0]);
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
      editAs: anchorAttributes.editAs,
      from: marker(body, "from"),
      to: marker(body, "to"),
      position: absolutePosition(body),
      extent: drawingExtent(body),
      crop: picture ? pictureCrop(picture) : undefined,
      effects: picture ? pictureEffects(picture) : undefined,
      transform: picture ? pictureTransform(picture) : undefined,
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

function drawingFontSize(xml, elementName) {
  const values = [...String(xml || "").matchAll(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${elementName}\\b[^>]*>`, "g"))]
    .map((match) => Number(attributes(match[0]).sz))
    .filter((value) => Number.isInteger(value) && value >= 100 && value <= 400_000);
  if (values.length === 0 || values.some((value) => value !== values[0])) return undefined;
  return values[0] / 100;
}

function chartSeriesLineStyle(shapeProperties) {
  const line = directElement(shapeProperties, "ln");
  if (!line) return undefined;
  const lineAttributes = attributes(line.startTag);
  if (Object.keys(lineAttributes).some((name) => name !== "w" && !name.startsWith("xmlns"))) return undefined;
  const children = directChildNames(line.body);
  if (children.some((name) => !["solidFill", "prstDash"].includes(name)) || children.filter((name) => name === "solidFill").length > 1 || children.filter((name) => name === "prstDash").length > 1) return undefined;

  const output = {};
  if (lineAttributes.w != null) {
    const width = Number(lineAttributes.w);
    if (!Number.isInteger(width) || width < 0 || width > 20_116_800) return undefined;
    output.width = width / 12_700;
  }
  const solidFill = directElement(line.body, "solidFill");
  if (solidFill) {
    if (Object.keys(attributes(solidFill.startTag)).some((name) => !name.startsWith("xmlns")) || directChildNames(solidFill.body).join("\0") !== "srgbClr") return undefined;
    const color = directElement(solidFill.body, "srgbClr");
    const colorAttributes = attributes(color?.startTag);
    if (!color || color.body.trim() || Object.keys(colorAttributes).some((name) => name !== "val" && !name.startsWith("xmlns")) || !/^[0-9a-f]{6}$/i.test(colorAttributes.val || "")) return undefined;
    output.fill = `#${colorAttributes.val.toUpperCase()}`;
  }
  const dash = directElement(line.body, "prstDash");
  if (dash) {
    const dashAttributes = attributes(dash.startTag);
    const style = new Map([["solid", "solid"], ["dash", "dashed"], ["dot", "dotted"], ["dashDot", "dash-dot"], ["lgDashDotDot", "dash-dot-dot"]]).get(dashAttributes.val);
    if (!style || dash.body.trim() || Object.keys(dashAttributes).some((name) => name !== "val" && !name.startsWith("xmlns"))) return undefined;
    output.style = style;
  }
  return output;
}

function chartDirectRgbFill(shapeProperties) {
  const solidFill = directElement(shapeProperties, "solidFill");
  if (!solidFill) return { present: false };
  if (Object.keys(attributes(solidFill.startTag)).some((name) => !name.startsWith("xmlns")) || directChildNames(solidFill.body).join("\0") !== "srgbClr") return { present: true };
  const color = directElement(solidFill.body, "srgbClr");
  const colorAttributes = attributes(color?.startTag);
  if (!color || color.body.trim() || Object.keys(colorAttributes).some((name) => name !== "val" && !name.startsWith("xmlns")) || !/^[0-9a-f]{6}$/i.test(colorAttributes.val || "")) return { present: true };
  return { present: true, value: `#${colorAttributes.val.toUpperCase()}` };
}

function chartSeriesMarker(seriesBody) {
  const marker = directElement(seriesBody, "marker");
  if (!marker) return undefined;
  if (Object.keys(attributes(marker.startTag)).some((name) => !name.startsWith("xmlns"))) return undefined;
  const children = directChildNames(marker.body);
  if (children.some((name) => !["symbol", "size", "spPr"].includes(name)) || children.filter((name) => name === "symbol").length > 1 || children.filter((name) => name === "size").length > 1 || children.filter((name) => name === "spPr").length > 1) return undefined;
  const output = {};
  const symbol = directElement(marker.body, "symbol");
  if (symbol) {
    const symbolAttributes = attributes(symbol.startTag);
    const symbols = new Set(["none", "dot", "circle", "square", "diamond", "triangle", "x", "star", "plus", "dash"]);
    if (symbol.body.trim() || Object.keys(symbolAttributes).some((name) => name !== "val" && !name.startsWith("xmlns")) || !symbols.has(symbolAttributes.val)) return undefined;
    output.symbol = symbolAttributes.val;
  }
  const size = directElement(marker.body, "size");
  if (size) {
    const sizeAttributes = attributes(size.startTag);
    const value = Number(sizeAttributes.val);
    if (size.body.trim() || Object.keys(sizeAttributes).some((name) => name !== "val" && !name.startsWith("xmlns")) || !Number.isInteger(value) || value < 2 || value > 72) return undefined;
    output.size = value;
  }
  const shapeProperties = directElement(marker.body, "spPr");
  if (shapeProperties) {
    if (Object.keys(attributes(shapeProperties.startTag)).some((name) => !name.startsWith("xmlns"))) return undefined;
    const shapeChildren = directChildNames(shapeProperties.body);
    if (shapeChildren.some((name) => !["solidFill", "ln"].includes(name)) || shapeChildren.filter((name) => name === "solidFill").length > 1 || shapeChildren.filter((name) => name === "ln").length > 1) return undefined;
    const fill = chartDirectRgbFill(shapeProperties.body);
    if (fill.present && fill.value == null) return undefined;
    if (fill.value != null) output.fill = fill.value;
    const nativeLine = directElement(shapeProperties.body, "ln");
    const line = chartSeriesLineStyle(shapeProperties.body);
    if (nativeLine && line == null) return undefined;
    if (line != null) output.line = line;
  }
  return output;
}

function chartLineOptions(xml) {
  const children = directChildNames(xml);
  if (children.filter((name) => name === "grouping").length > 1 || children.filter((name) => name === "varyColors").length > 1 || children.filter((name) => name === "smooth").length > 1) return undefined;
  const output = {};
  const grouping = directElement(xml, "grouping");
  if (grouping) {
    const groupingAttributes = attributes(grouping.startTag);
    if (grouping.body.trim() || Object.keys(groupingAttributes).some((name) => name !== "val" && !name.startsWith("xmlns")) || !["standard", "stacked", "percentStacked"].includes(groupingAttributes.val)) return undefined;
    output.grouping = groupingAttributes.val;
  }
  const varyColors = directElement(xml, "varyColors");
  if (varyColors) {
    const varyAttributes = attributes(varyColors.startTag);
    if (varyColors.body.trim() || Object.keys(varyAttributes).some((name) => name !== "val" && !name.startsWith("xmlns")) || !["0", "1", "false", "true"].includes(varyAttributes.val)) return undefined;
    if (["1", "true"].includes(varyAttributes.val)) output.varyColors = true;
  }
  const smooth = directElement(xml, "smooth");
  if (smooth) {
    const smoothAttributes = attributes(smooth.startTag);
    if (smooth.body.trim() || Object.keys(smoothAttributes).some((name) => name !== "val" && !name.startsWith("xmlns")) || !["0", "1", "false", "true"].includes(smoothAttributes.val)) return undefined;
    output.smooth = smoothAttributes.val === "1" || smoothAttributes.val === "true";
  }
  return Object.keys(output).length ? output : undefined;
}

function chartDataLabels(xml) {
  const plotChildren = directChildNames(xml);
  if (plotChildren.filter((name) => name === "dLbls").length !== 1) return undefined;
  const labels = directElement(xml, "dLbls");
  if (!labels || Object.keys(attributes(labels.startTag)).some((name) => !name.startsWith("xmlns"))) return undefined;
  const allowed = new Set(["dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize"]);
  const children = directChildNames(labels.body);
  if (children.some((name) => !allowed.has(name)) || [...allowed].some((name) => children.filter((child) => child === name).length > 1) || children.filter((name) => name === "showVal").length !== 1 || children.filter((name) => name === "showCatName").length !== 1) return undefined;
  const scalar = (name) => {
    const element = directElement(labels.body, name);
    if (!element) return undefined;
    const attrs = attributes(element.startTag);
    if (element.body.trim() || Object.keys(attrs).some((key) => key !== "val" && !key.startsWith("xmlns")) || !["0", "1", "false", "true"].includes(attrs.val)) return undefined;
    return attrs.val === "1" || attrs.val === "true";
  };
  const showValue = scalar("showVal");
  const showCategoryName = scalar("showCatName");
  if (showValue == null || showCategoryName == null) return undefined;
  for (const name of ["showLegendKey", "showSerName", "showPercent", "showBubbleSize"]) if (children.includes(name) && scalar(name) !== false) return undefined;
  const positionElement = directElement(labels.body, "dLblPos");
  let position;
  if (positionElement) {
    const attrs = attributes(positionElement.startTag);
    const positions = new Map([["bestFit", "bestFit"], ["b", "bottom"], ["ctr", "center"], ["inBase", "insideBase"], ["inEnd", "insideEnd"], ["l", "left"], ["outEnd", "outsideEnd"], ["r", "right"], ["t", "top"]]);
    if (positionElement.body.trim() || Object.keys(attrs).some((key) => key !== "val" && !key.startsWith("xmlns")) || !positions.has(attrs.val)) return undefined;
    position = positions.get(attrs.val);
  }
  return { showValue, showCategoryName, ...(position == null ? {} : { position }) };
}

function chartAxis(xml, name, kind) {
  const body = elementBody(xml, name);
  if (!body) return undefined;
  const titleBody = elementBody(body, "title");
  const title = [...titleBody.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g)].map((match) => decodeXml(match[1])).join("") || elementValue(titleBody, "v");
  const numberFormatTag = /<(?:[A-Za-z_][\w.-]*:)?numFmt\b[^>]*\/?\s*>/.exec(body)?.[0];
  const numberFormatCode = numberFormatTag ? attributes(numberFormatTag).formatCode : undefined;
  const tickTextProperties = directElementBody(body, "txPr");
  const tickFontSize = drawingFontSize(tickTextProperties, "defRPr");
  const valueAttribute = (source, element) => {
    const tag = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${element}\\b[^>]*\\/?\\s*>`).exec(source)?.[0];
    return tag ? Number(attributes(tag).val) : undefined;
  };
  const output = {
    axisType: kind === "x" ? "textAxis" : "valueAxis",
    title: { text: title || "" },
    ...(tickFontSize == null ? {} : { textStyle: { fontSize: tickFontSize } }),
    ...(numberFormatCode ? { numberFormatCode } : {}),
  };
  if (kind === "x") {
    const interval = valueAttribute(body, "tickLblSkip");
    if (Number.isInteger(interval) && interval > 0) output.tickLabelInterval = interval;
  } else {
    const scaling = elementBody(body, "scaling");
    const minimum = valueAttribute(scaling, "min");
    const maximum = valueAttribute(scaling, "max");
    const majorUnit = valueAttribute(body, "majorUnit");
    if (Number.isFinite(minimum)) output.min = minimum;
    if (Number.isFinite(maximum)) output.max = maximum;
    if (Number.isFinite(majorUnit) && majorUnit > 0) output.majorUnit = majorUnit;
  }
  return output;
}

export function parseSpreadsheetChart(xml = "") {
  const text = String(xml || "");
  const type = /<(?:[A-Za-z_][\w.-]*:)?pieChart\b/.test(text) ? "pie" : /<(?:[A-Za-z_][\w.-]*:)?lineChart\b/.test(text) ? "line" : "bar";
  const plotBody = elementBody(text, type === "pie" ? "pieChart" : type === "line" ? "lineChart" : "barChart");
  const titleBody = elementBody(text, "title");
  const title = [...titleBody.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g)].map((match) => decodeXml(match[1])).join("") || elementValue(titleBody, "v");
  const titleFontSize = drawingFontSize(titleBody, "rPr");
  const lineOptions = type === "line" ? chartLineOptions(plotBody) : undefined;
  const dataLabels = chartDataLabels(plotBody);
  const series = [...text.matchAll(/<(?:[A-Za-z_][\w.-]*:)?ser\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?ser>/g)].map((match, index) => {
    const body = match[1];
    const tx = elementBody(body, "tx");
    const categoryBody = elementBody(body, "cat") || elementBody(body, "xVal");
    const valueBody = elementBody(body, "val") || elementBody(body, "yVal");
    const shapeProperties = directElementBody(body, "spPr");
    const solidFill = directElement(shapeProperties, "solidFill")?.body || "";
    const color = /<(?:[A-Za-z_][\w.-]*:)?srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(solidFill)?.[1];
    const line = chartSeriesLineStyle(shapeProperties);
    const marker = chartSeriesMarker(body);
    return {
      name: elementValue(tx, "v") || elementValue(tx, "f") || `Series ${index + 1}`,
      categoryFormula: elementValue(categoryBody, "f") || undefined,
      formula: elementValue(valueBody, "f") || undefined,
      categories: chartPoints(categoryBody),
      values: chartPoints(valueBody, true),
      fill: color ? `#${color.toUpperCase()}` : undefined,
      line,
      marker,
    };
  });
  return {
    type,
    title,
    ...(titleFontSize == null ? {} : { titleTextStyle: { fontSize: titleFontSize } }),
    ...(lineOptions == null ? {} : { lineOptions }),
    ...(dataLabels == null ? {} : { dataLabels }),
    hasLegend: /<(?:[A-Za-z_][\w.-]*:)?legend\b/.test(text),
    categories: series[0]?.categories || [],
    series,
    xAxis: type === "pie" ? undefined : chartAxis(text, "catAx", "x"),
    yAxis: type === "pie" ? undefined : chartAxis(text, "valAx", "y"),
  };
}
