import { Buffer } from "node:buffer";
import { resolveColorToken } from "../shared/colors.mjs";
import { normalizePresentationTextBodyProperties } from "./text-body-properties.mjs";

export function node(type, props = {}, children = []) {
  const normalizedProps = props && !Array.isArray(props) && typeof props === "object" ? props : {};
  const rawChildren = Array.isArray(props) || typeof props === "string" || typeof props === "number" ? props : children;
  return { type, props: normalizedProps, children: normalizeComposeChildren(rawChildren) };
}

export const row = (props = {}, children = []) => node("row", props, children);
export const column = (props = {}, children = []) => node("column", props, children);
export const grid = (props = {}, children = []) => node("grid", props, children);
export const layers = (props = {}, children = []) => node("layers", props, children);
export const box = (props = {}, children = []) => node("box", props, children);
export const paragraph = (props = {}, children = []) => node("paragraph", props, children);
// Reference presentation templates use a children-first text helper. Keep it
// as a thin alias so the resulting node is identical to `paragraph(...)`.
export const text = (children = [], props = {}) => paragraph(props, children);
export const run = (props = {}, children = []) => node("run", props, children);
export const shape = (props = {}, children = []) => node("shape", props, children);
export const image = (props = {}, children = []) => node("image", props, children);
export const table = (props = {}, children = []) => node("table", props, children);
export const chart = (props = {}, children = []) => node("chart", props, children);
export const rule = (props = {}, children = []) => node("rule", props, children);

function normalizeComposeChildren(children) {
  if (children == null || children === false) return [];
  if (!Array.isArray(children)) return [children];
  return children.flatMap((child) => normalizeComposeChildren(child));
}

function isComposeNode(value) {
  return value && typeof value === "object" && typeof value.type === "string" && Array.isArray(value.children);
}

function textFromComposeChildren(children) {
  return normalizeComposeChildren(children).map((child) => {
    if (typeof child === "string" || typeof child === "number") return String(child);
    if (isComposeNode(child)) return textFromComposeChildren(child.children);
    if (child && typeof child === "object" && Array.isArray(child.runs)) return child.runs.map((run) => String(run?.run ?? run?.text ?? "")).join("");
    return "";
  }).join("");
}

function composeTokenRuns(runs = []) {
  return runs.flatMap((run) => {
    const text = String(run?.run ?? run?.text ?? "");
    const style = run?.textStyle || run?.style || {};
    const segments = text.split("\n");
    return segments.flatMap((segment, index) => [
      ...(segment ? [{ text: segment, style }] : []),
      ...(index < segments.length - 1 ? [{ break: true, style }] : []),
    ]);
  });
}

function composeRichText(children) {
  const items = normalizeComposeChildren(children);
  if (!items.some((item) => item && typeof item === "object" && !isComposeNode(item) && Array.isArray(item.runs))) return undefined;
  return items.map((item) => {
    if (!item || typeof item !== "object" || isComposeNode(item) || !Array.isArray(item.runs)) return { runs: [String(item ?? "")] };
    return {
      runs: composeTokenRuns(item.runs),
      ...(item.bulletCharacter != null ? { bulletCharacter: item.bulletCharacter } : {}),
      ...(item.marginLeft != null ? { marginLeft: Number(item.marginLeft) / 9525 } : {}),
      ...(item.indent != null ? { indent: Number(item.indent) / 9525 } : {}),
      ...(item.spaceBefore != null ? { spaceBefore: Number(item.spaceBefore) / 75 } : {}),
      ...(item.spaceAfter != null ? { spaceAfter: Number(item.spaceAfter) / 75 } : {}),
      ...(item.paragraphStyle?.lineSpacingPercent != null ? { lineSpacing: Number(item.paragraphStyle.lineSpacingPercent) / 100_000 } : {}),
    };
  });
}

function normalizePadding(padding = {}) {
  if (typeof padding === "number") return { top: padding, right: padding, bottom: padding, left: padding };
  return {
    top: padding.top ?? padding.y ?? 0,
    right: padding.right ?? padding.x ?? 0,
    bottom: padding.bottom ?? padding.y ?? 0,
    left: padding.left ?? padding.x ?? 0,
  };
}

function innerFrame(frame, padding) {
  return {
    left: frame.left + padding.left,
    top: frame.top + padding.top,
    width: Math.max(0, frame.width - padding.left - padding.right),
    height: Math.max(0, frame.height - padding.top - padding.bottom),
  };
}

function parseTextStyle(props = {}) {
  const style = {};
  const className = String(props.className || "");
  for (const token of className.split(/\s+/).filter(Boolean)) {
    if (token === "font-bold") style.bold = true;
    if (token === "font-semibold") style.bold = true;
    if (token.startsWith("text-")) {
      const sizeMap = { xs: 12, sm: 14, base: 16, lg: 18, xl: 20, "2xl": 24, "3xl": 30, "4xl": 36, "5xl": 48, "6xl": 60 };
      const key = token.slice(5);
      if (sizeMap[key]) style.fontSize = sizeMap[key];
      else if (/^\[\d+px\]$/.test(key)) style.fontSize = Number(key.slice(1, -3));
      else style.color = resolveColorToken(key, style.color);
    }
    if (token.startsWith("leading-")) {
      const key = token.slice(8);
      if (key === "tight") style.lineSpacing = 1.1;
      else if (key === "relaxed") style.lineSpacing = 1.35;
      else if (/^\[\d+(?:\.\d+)?\]$/.test(key)) style.lineSpacing = Number(key.slice(1, -1));
    }
  }
  if (typeof props.style === "string") {
    const font = /font:\s*(\d+)\s+(\d+)px/i.exec(props.style);
    if (font) { style.bold = Number(font[1]) >= 600; style.fontSize = Number(font[2]); }
    const color = /color:\s*([^;]+)/i.exec(props.style);
    if (color) style.color = resolveColorToken(color[1].trim(), style.color);
    const leading = /leading:\s*([\d.]+)/i.exec(props.style);
    if (leading) style.lineSpacing = Number(leading[1]);
  } else if (props.style && typeof props.style === "object") {
    Object.assign(style, props.style);
  }
  if (style.typeface != null && style.fontFamily == null) style.fontFamily = style.typeface;
  delete style.typeface;
  if (style.fontSize != null) {
    const raw = String(style.fontSize).trim();
    const number = Number(raw.replace(/(?:px|pt)$/i, ""));
    if (Number.isFinite(number)) style.fontSize = /pt$/i.test(raw) ? number * 4 / 3 : number;
  }
  delete style.verticalAlignment;
  delete style.autoFit;
  delete style.insets;
  delete style.wrap;
  return style;
}

function parseTextBodyProperties(props = {}) {
  const style = props.style && typeof props.style === "object" && !Array.isArray(props.style) ? props.style : {};
  const verticalAlignment = style.verticalAlignment ?? props.verticalAlignment;
  const autoFit = style.autoFit ?? props.autoFit;
  const bodyProperties = {
    ...(style.insets ?? props.insets ? { insets: style.insets ?? props.insets } : {}),
    ...(verticalAlignment != null ? { anchor: verticalAlignment === "middle" ? "center" : verticalAlignment } : {}),
    ...(autoFit != null ? { autoFit: autoFit === "resizeShapeToFitText" ? "resizeShape" : autoFit } : {}),
    ...(style.wrap ?? props.wrap ? { wrap: style.wrap ?? props.wrap } : {}),
  };
  return Object.keys(bodyProperties).length ? normalizePresentationTextBodyProperties(bodyProperties) : undefined;
}

function composeNodeFrame(frame, props = {}) {
  const position = props.position;
  if (!position || typeof position !== "object" || Array.isArray(position)) return frame;
  const left = Number(position.left ?? 0);
  const top = Number(position.top ?? 0);
  const width = Number(position.width ?? (typeof props.width === "number" ? props.width : frame.width));
  const height = Number(position.height ?? (typeof props.height === "number" ? props.height : frame.height));
  return {
    left: frame.left + (Number.isFinite(left) ? left : 0),
    top: frame.top + (Number.isFinite(top) ? top : 0),
    width: Number.isFinite(width) ? width : frame.width,
    height: Number.isFinite(height) ? height : frame.height,
  };
}

function composeImagePlaceholderDataUrl(props) {
  if (props.dataUrl || props.uri || !props.prompt) return undefined;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000"><rect width="1000" height="1000" fill="#eaf5fb"/><path d="M0 760 330 430 560 660 1000 210V1000H0Z" fill="#c7e4f2"/><circle cx="770" cy="245" r="105" fill="#9ecfe4"/><path d="M110 160h370v38H110zm0 78h250v24H110z" fill="#ffffff" opacity=".72"/></svg>`;
  return `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`;
}

function styleFromClassName(className = "") {
  const style = {};
  for (const token of String(className).split(/\s+/).filter(Boolean)) {
    if (token.startsWith("bg-")) style.fill = resolveColorToken(token.slice(3), style.fill);
    if (token.startsWith("rounded-")) style.borderRadius = token;
  }
  return style;
}

function composeIntrinsicSize(composeNode) {
  const props = composeNode.props || {};
  const text = textFromComposeChildren(composeNode.children);
  const textStyle = parseTextStyle(props);
  const fontSize = textStyle.fontSize || 20;
  const lineCount = Math.max(1, text.split(/\r?\n/).length);
  return {
    width: typeof props.width === "number" ? props.width : Math.max(160, Math.min(720, text.length * fontSize * 0.55 + 24)),
    height: typeof props.height === "number" ? props.height : Math.max(36, lineCount * fontSize * (textStyle.lineSpacing || 1.2) + 18),
  };
}

function composeChildFrames(children, frame, direction, gap) {
  const mainSize = direction === "row" ? "width" : "height";
  const crossSize = direction === "row" ? "height" : "width";
  const fixed = children.reduce((sum, child) => {
    const prop = child.props?.[mainSize];
    if (typeof prop === "number") return sum + prop;
    if (prop === "hug") return sum + composeIntrinsicSize(child)[mainSize];
    return sum;
  }, 0);
  const fillCount = children.filter((child) => child.props?.[mainSize] === "fill" || child.props?.[mainSize] == null).length || 1;
  const available = Math.max(0, frame[mainSize] - fixed - gap * Math.max(0, children.length - 1));
  const fillSize = available / fillCount;
  let cursor = direction === "row" ? frame.left : frame.top;
  return children.map((child) => {
    const intrinsic = composeIntrinsicSize(child);
    const main = typeof child.props?.[mainSize] === "number" ? child.props[mainSize] : child.props?.[mainSize] === "hug" ? intrinsic[mainSize] : fillSize;
    const cross = typeof child.props?.[crossSize] === "number" ? child.props[crossSize] : child.props?.[crossSize] === "hug" ? intrinsic[crossSize] : frame[crossSize];
    const childFrame = direction === "row"
      ? { left: cursor, top: frame.top, width: main, height: cross }
      : { left: frame.left, top: cursor, width: cross, height: main };
    cursor += main + gap;
    return childFrame;
  });
}

function normalizeTrack(track) {
  if (typeof track === "number") return { mode: "fixed", value: track };
  if (typeof track === "string") return track === "fixed" ? { mode: "fixed", value: 0 } : { mode: "fr", value: 1 };
  return { mode: track?.mode || "fr", value: Number(track?.value ?? 1) };
}

function resolveGridTracks(total, tracks, fallbackCount, gap) {
  const normalized = (tracks?.length ? tracks : Array.from({ length: fallbackCount }, () => ({ mode: "fr", value: 1 }))).map(normalizeTrack);
  const fixed = normalized.reduce((sum, track) => track.mode === "fixed" ? sum + track.value : sum, 0);
  const fr = normalized.reduce((sum, track) => track.mode === "fr" ? sum + Math.max(0, track.value) : sum, 0) || 1;
  const available = Math.max(0, total - fixed - gap * Math.max(0, normalized.length - 1));
  return normalized.map((track) => track.mode === "fixed" ? track.value : available * Math.max(0, track.value) / fr);
}

function gridChildFrame(frame, columns, rows, columnGap, rowGap, columnIndex, rowIndex, columnSpan = 1, rowSpan = 1) {
  const left = frame.left + columns.slice(0, columnIndex).reduce((sum, value) => sum + value, 0) + columnGap * columnIndex;
  const top = frame.top + rows.slice(0, rowIndex).reduce((sum, value) => sum + value, 0) + rowGap * rowIndex;
  const width = columns.slice(columnIndex, columnIndex + columnSpan).reduce((sum, value) => sum + value, 0) + columnGap * Math.max(0, columnSpan - 1);
  const height = rows.slice(rowIndex, rowIndex + rowSpan).reduce((sum, value) => sum + value, 0) + rowGap * Math.max(0, rowSpan - 1);
  return { left, top, width, height };
}

export function materializeComposeNode(slide, composeNode, frame) {
  if (typeof composeNode === "string" || typeof composeNode === "number") {
    return materializeComposeNode(slide, paragraph({}, [String(composeNode)]), frame);
  }
  if (!isComposeNode(composeNode)) return [];
  const props = composeNode.props || {};
  frame = composeNodeFrame(frame, props);
  const children = normalizeComposeChildren(composeNode.children).filter((child) => child !== null && child !== undefined && child !== false);
  const type = composeNode.type;
  if (type === "row" || type === "column") {
    const pad = normalizePadding(props.padding);
    const inner = innerFrame(frame, pad);
    const childFrames = composeChildFrames(children.filter(isComposeNode), inner, type, Number(props.gap || 0));
    return children.filter(isComposeNode).flatMap((child, index) => materializeComposeNode(slide, child, childFrames[index]));
  }
  if (type === "grid") {
    const gridChildren = children.filter(isComposeNode);
    const pad = normalizePadding(props.padding);
    const inner = innerFrame(frame, pad);
    const columnGap = Number(props.columnGap ?? props.gap ?? 0);
    const rowGap = Number(props.rowGap ?? props.gap ?? 0);
    const fallbackColumns = Math.max(1, props.columns?.length || Math.ceil(Math.sqrt(gridChildren.length || 1)));
    const columns = resolveGridTracks(inner.width, props.columns, fallbackColumns, columnGap);
    const fallbackRows = Math.max(1, props.rows?.length || Math.ceil((gridChildren.length || 1) / columns.length));
    const rows = resolveGridTracks(inner.height, props.rows, fallbackRows, rowGap);
    return gridChildren.flatMap((child, index) => {
      const columnIndex = Math.min(columns.length - 1, Number(child.props?.column ?? child.props?.col ?? (index % columns.length)));
      const rowIndex = Math.min(rows.length - 1, Number(child.props?.row ?? Math.floor(index / columns.length)));
      const columnSpan = Math.min(columns.length - columnIndex, Math.max(1, Number(child.props?.columnSpan ?? 1)));
      const rowSpan = Math.min(rows.length - rowIndex, Math.max(1, Number(child.props?.rowSpan ?? 1)));
      return materializeComposeNode(slide, child, gridChildFrame(inner, columns, rows, columnGap, rowGap, columnIndex, rowIndex, columnSpan, rowSpan));
    });
  }
  if (type === "layers") {
    const pad = normalizePadding(props.padding);
    const inner = innerFrame(frame, pad);
    return children.filter(isComposeNode).flatMap((child) => materializeComposeNode(slide, child, {
      left: inner.left,
      top: inner.top,
      width: typeof child.props?.width === "number" ? child.props.width : inner.width,
      height: typeof child.props?.height === "number" ? child.props.height : inner.height,
    }));
  }
  if (type === "box") {
    const classStyle = styleFromClassName(props.className);
    const surface = slide.shapes.add({
      id: props.id,
      name: props.name,
      geometry: props.geometry || "roundRect",
      position: frame,
      fill: props.fill || classStyle.fill || "transparent",
      line: props.line || { fill: "transparent", width: 0 },
      borderRadius: props.borderRadius || classStyle.borderRadius,
    });
    const pad = normalizePadding(props.padding ?? { x: 0, y: 0 });
    return [surface, ...children.filter(isComposeNode).flatMap((child) => materializeComposeNode(slide, child, innerFrame(frame, pad)))];
  }
  if (type === "paragraph") {
    const richText = composeRichText(children);
    const shape = slide.shapes.add({
      id: props.id,
      name: props.name,
      geometry: "textbox",
      position: frame,
      fill: "transparent",
      line: { fill: "transparent", width: 0 },
      text: richText || textFromComposeChildren(children),
    });
    shape.text.style = parseTextStyle(props);
    const bodyProperties = parseTextBodyProperties(props);
    if (bodyProperties) shape.text.bodyProperties = bodyProperties;
    return [shape];
  }
  if (type === "shape") {
    const classStyle = styleFromClassName(props.className);
    if (props.geometry === "straightConnector1") {
      return [slide.connectors.add({
        id: props.id,
        name: props.name,
        connectorType: "straight",
        start: { x: frame.left, y: frame.top },
        end: { x: frame.left + frame.width, y: frame.top + frame.height },
        line: props.line || { fill: props.fill || "#334155", width: 1 },
      })];
    }
    const richText = composeRichText(children);
    const shape = slide.shapes.add({
      ...props,
      position: frame,
      fill: props.fill || classStyle.fill || "transparent",
      text: richText || textFromComposeChildren(children) || props.text,
    });
    shape.text.style = parseTextStyle(props);
    const bodyProperties = parseTextBodyProperties(props);
    if (bodyProperties) shape.text.bodyProperties = bodyProperties;
    return [shape];
  }
  if (type === "table") {
    return [slide.tables.add({ ...props, position: frame })];
  }
  if (type === "chart") {
    return [slide.charts.add(props.chartType || props.type || "bar", { ...props, position: frame })];
  }
  if (type === "image") {
    const placeholderDataUrl = composeImagePlaceholderDataUrl(props);
    return [slide.images.add({
      ...props,
      position: frame,
      alt: props.alt || textFromComposeChildren(children) || props.name,
      ...(placeholderDataUrl ? { dataUrl: placeholderDataUrl, fit: "stretch", geometry: "rect", borderRadius: undefined } : {}),
    })];
  }
  if (type === "rule") {
    const horizontal = (props.width ?? frame.width) >= (props.height ?? props.weight ?? 2);
    const shape = slide.shapes.add({
      id: props.id,
      name: props.name,
      geometry: "rect",
      position: { left: frame.left, top: frame.top, width: horizontal ? frame.width : Number(props.weight || 2), height: horizontal ? Number(props.weight || 2) : frame.height },
      fill: props.stroke || "#0f172a",
      line: { fill: props.stroke || "#0f172a", width: 0 },
    });
    return [shape];
  }
  return children.filter(isComposeNode).flatMap((child) => materializeComposeNode(slide, child, frame));
}
