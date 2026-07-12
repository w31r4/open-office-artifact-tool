import { attrEscape, attributes, decodeXml } from "../ooxml/source-reference-xml.mjs";

const EMU_PER_PIXEL = 9525;

function localName(tag = "") {
  return /^<\/?(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)\b/.exec(tag)?.[1];
}

function openingTag(xml = "", name) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>`).exec(String(xml))?.[0] || "";
}

function pointTag(xml = "", name) {
  const tag = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*\\/\\s*>`).exec(String(xml))?.[0] || "";
  const attrs = attributes(tag);
  return { x: Number(attrs.x || 0) / EMU_PER_PIXEL, y: Number(attrs.y || 0) / EMU_PER_PIXEL };
}

function sizeTag(xml = "", name) {
  const tag = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}\\b[^>]*\\/\\s*>`).exec(String(xml))?.[0] || "";
  const attrs = attributes(tag);
  return { width: Number(attrs.cx || 0) / EMU_PER_PIXEL, height: Number(attrs.cy || 0) / EMU_PER_PIXEL };
}

export function directPresentationChildren(xml = "", parentLocalName) {
  const source = String(xml);
  const parent = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${parentLocalName}\\b[^>]*>`).exec(source);
  if (!parent) return [];
  const start = (parent.index || 0) + parent[0].length;
  const tokens = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\/?[A-Za-z_][\w:.-]*\b[^>]*>/g;
  tokens.lastIndex = start;
  const children = [];
  let depth = 0;
  let childStart;
  let childName;
  for (let token = tokens.exec(source); token; token = tokens.exec(source)) {
    const text = token[0];
    if (text.startsWith("<!--") || text.startsWith("<![CDATA[")) continue;
    const name = localName(text);
    const closing = text.startsWith("</");
    const selfClosing = /\/\s*>$/.test(text);
    if (closing && depth === 0 && name === parentLocalName) break;
    if (!closing && depth === 0) {
      childStart = token.index;
      childName = name;
      if (selfClosing) {
        children.push({ localName: childName, xml: source.slice(childStart, tokens.lastIndex) });
        childStart = undefined;
        childName = undefined;
      } else depth = 1;
      continue;
    }
    if (!closing && !selfClosing) depth += 1;
    else if (closing) depth -= 1;
    if (depth === 0 && childStart != null) {
      children.push({ localName: childName, xml: source.slice(childStart, tokens.lastIndex) });
      childStart = undefined;
      childName = undefined;
    }
  }
  return children;
}

export function normalizeGroupGeometry(position = {}, childFrame = {}) {
  const frame = {
    left: Number(position.left ?? 0),
    top: Number(position.top ?? 0),
    width: Math.max(1, Number(position.width ?? 320)),
    height: Math.max(1, Number(position.height ?? 180)),
  };
  const children = {
    left: Number(childFrame.left ?? 0),
    top: Number(childFrame.top ?? 0),
    width: Math.max(1, Number(childFrame.width ?? frame.width)),
    height: Math.max(1, Number(childFrame.height ?? frame.height)),
  };
  if (![...Object.values(frame), ...Object.values(children)].every(Number.isFinite)) throw new TypeError("Presentation group frames must contain finite coordinates and positive sizes.");
  return { frame, childFrame: children };
}

export function parsePresentationGroupGeometry(xml = "") {
  const transform = /<(?:[A-Za-z_][\w.-]*:)?xfrm\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?xfrm>/.exec(String(xml))?.[0] || "";
  const off = pointTag(transform, "off");
  const ext = sizeTag(transform, "ext");
  const chOff = pointTag(transform, "chOff");
  const chExt = sizeTag(transform, "chExt");
  return normalizeGroupGeometry(
    { left: off.x, top: off.y, width: ext.width || 1, height: ext.height || 1 },
    { left: chOff.x, top: chOff.y, width: chExt.width || ext.width || 1, height: chExt.height || ext.height || 1 },
  );
}

function emu(value) {
  return Math.round(Number(value) * EMU_PER_PIXEL);
}

export function presentationGroupShapeXml(group, childrenXml, creationIdExtension = "") {
  const { frame, childFrame } = normalizeGroupGeometry(group.position, group.childFrame);
  const nativeId = Number(group.nativeId);
  if (!Number.isInteger(nativeId) || nativeId < 1 || nativeId > 4_294_967_295) throw new RangeError(`Presentation group ${group.id} requires a valid native drawing ID.`);
  const transform = `<a:xfrm><a:off x="${emu(frame.left)}" y="${emu(frame.top)}"/><a:ext cx="${emu(frame.width)}" cy="${emu(frame.height)}"/><a:chOff x="${emu(childFrame.left)}" y="${emu(childFrame.top)}"/><a:chExt cx="${emu(childFrame.width)}" cy="${emu(childFrame.height)}"/></a:xfrm>`;
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${nativeId}" name="${attrEscape(group.name || group.id)}">${creationIdExtension}</p:cNvPr><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${transform}</p:grpSpPr>${childrenXml}</p:grpSp>`;
}

export function presentationGroupSvgTransform(position, childFrame) {
  const geometry = normalizeGroupGeometry(position, childFrame);
  const scaleX = geometry.frame.width / geometry.childFrame.width;
  const scaleY = geometry.frame.height / geometry.childFrame.height;
  return `translate(${geometry.frame.left} ${geometry.frame.top}) scale(${scaleX} ${scaleY}) translate(${-geometry.childFrame.left} ${-geometry.childFrame.top})`;
}

export function groupChildAbsoluteFrame(position, childFrame, childPosition) {
  const geometry = normalizeGroupGeometry(position, childFrame);
  const scaleX = geometry.frame.width / geometry.childFrame.width;
  const scaleY = geometry.frame.height / geometry.childFrame.height;
  return {
    left: geometry.frame.left + (Number(childPosition.left) - geometry.childFrame.left) * scaleX,
    top: geometry.frame.top + (Number(childPosition.top) - geometry.childFrame.top) * scaleY,
    width: Number(childPosition.width) * scaleX,
    height: Number(childPosition.height) * scaleY,
  };
}

export function presentationGroupName(xml = "") {
  const cNvPr = openingTag(xml, "cNvPr");
  return decodeXml(attributes(cNvPr).name || "");
}

export function parsePresentationGroupTree(owner, xml, context, adapters) {
  const geometry = parsePresentationGroupGeometry(xml);
  const group = owner.groups.add({ name: presentationGroupName(xml), position: geometry.frame, childFrame: geometry.childFrame });
  adapters.applyIdentity(group, xml, "grpSpMk");
  for (const child of directPresentationChildren(xml, "grpSp")) {
    if (child.localName === "sp") adapters.parseShape(group, child.xml, { ...context, layout: undefined });
    else if (child.localName === "cxnSp") adapters.parseConnector(group, child.xml);
    else if (child.localName === "grpSp") parsePresentationGroupTree(group, child.xml, context, adapters);
  }
  return group;
}

export function createPresentationGroupShapeClass(adapters) {
  return class GroupShape {
    constructor(slide, config = {}) {
      this.slide = slide;
      this.kind = "groupShape";
      this.id = config.id || adapters.createId("gr");
      this.nativeId = config.nativeId;
      this.creationId = config.creationId;
      this.name = config.name || "";
      const geometry = normalizeGroupGeometry(config.position || config.frame, config.childFrame || config.childrenFrame);
      this.position = geometry.frame;
      this.childFrame = geometry.childFrame;
      this.children = [];
      this.shapes = adapters.createShapeCollection(slide, this);
      this.connectors = adapters.createConnectorCollection(slide, this);
      this.groups = adapters.createGroupCollection(slide, this, this.constructor);
      for (const child of config.children || []) {
        if (child?.kind === "groupShape" || child?.kind === "group") this.groups.add(child);
        else if (child?.kind === "connector") this.connectors.add(child);
        else this.shapes.add(child);
      }
      for (const shape of config.shapes || []) this.shapes.add(shape);
      for (const connector of config.connectors || []) this.connectors.add(connector);
      for (const group of config.groups || []) this.groups.add(group);
    }

    _rememberChild(element) { this.children.push(element); }

    absoluteFrame() {
      return this.parentGroup ? groupChildAbsoluteFrame(this.parentGroup.absoluteFrame(), this.parentGroup.childFrame, this.position) : { ...this.position };
    }

    absoluteChildFrame(element) {
      return groupChildAbsoluteFrame(this.absoluteFrame(), this.childFrame, element.position);
    }

    resolve(id) {
      for (const child of this.children) {
        if (child.id === id) return child;
        if (String(id || "").endsWith("/text") && adapters.isShape(child) && `${child.id}/text` === id) return adapters.createTextRange(child, id);
        if (adapters.isGroup(child)) {
          const nested = child.resolve(id);
          if (nested) return nested;
        }
      }
      return undefined;
    }

    allElements() {
      return [this, ...this.children.flatMap((child) => adapters.isGroup(child) ? child.allElements() : [child])];
    }

    inspectRecord() {
      const frame = this.absoluteFrame();
      return { kind: "groupShape", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, children: this.children.length, childIds: this.children.map((child) => child.id), bbox: [frame.left, frame.top, frame.width, frame.height], bboxUnit: "px", childFrame: this.childFrame };
    }

    inspectRecords(kinds) {
      const records = [];
      if (kinds.has("groupShape") || kinds.has("group") || kinds.has("shape")) records.push(this.inspectRecord());
      for (const child of this.children) {
        if (adapters.isGroup(child)) records.push(...child.inspectRecords(kinds));
        else if (adapters.isShape(child)) {
          const frame = this.absoluteChildFrame(child);
          const bbox = [frame.left, frame.top, frame.width, frame.height];
          if (kinds.has("textbox") && child.text.value) records.push({ ...child.inspectRecord("textbox"), parentGroupId: this.id, bbox });
          else if (kinds.has("shape")) records.push({ ...child.inspectRecord("shape"), parentGroupId: this.id, bbox });
          if (kinds.has("textRange") && child.text.value) records.push(adapters.textRangeRecord(child, { parentKind: "shape", record: { slide: this.slide.index + 1, parentGroupId: this.id, bbox, bboxUnit: "px" } }));
        } else if (adapters.isConnector(child) && kinds.has("connector")) records.push({ ...child.inspectRecord(), parentGroupId: this.id });
      }
      return records;
    }

    layoutJson() {
      const children = this.children.map((child) => {
        if (adapters.isGroup(child)) return child.layoutJson();
        const record = child.layoutJson();
        return { ...record, localFrame: record.frame || child.position, frame: this.absoluteChildFrame(child), parentGroupId: this.id };
      });
      return { kind: "groupShape", id: this.id, name: this.name, frame: this.absoluteFrame(), localFrame: this.position, childFrame: this.childFrame, children };
    }

    toSvg() {
      return `<g data-group-id="${attrEscape(this.id)}" transform="${presentationGroupSvgTransform(this.position, this.childFrame)}">${this.children.map((child) => child.toSvg()).join("")}</g>`;
    }

    toPptxShape(index) {
      const childrenXml = this.children.map((child, childIndex) => child.toPptxShape(index + childIndex + 1)).join("");
      return presentationGroupShapeXml(this, childrenXml, adapters.creationIdExtensionXml(this.creationId));
    }

    validateLayout() {
      const issues = [];
      const bounds = this.childFrame;
      for (const child of this.children) {
        const frame = child.position;
        if (frame && (frame.left < bounds.left || frame.top < bounds.top || frame.left + frame.width > bounds.left + bounds.width || frame.top + frame.height > bounds.top + bounds.height)) issues.push({ kind: "layoutIssue", type: "groupChildOutOfBounds", severity: "error", slide: this.slide.index + 1, id: child.id, groupId: this.id, message: `${adapters.elementLabel(child)} extends outside group ${adapters.elementLabel(this)} child coordinates.` });
        if (adapters.isGroup(child)) issues.push(...child.validateLayout());
      }
      return issues;
    }

    toProto() {
      return { kind: "groupShape", id: this.id, name: this.name, position: this.position, childFrame: this.childFrame, children: this.children.map((child) => adapters.isGroup(child) ? child.toProto() : { ...child.layoutJson(), kind: adapters.isConnector(child) ? "connector" : "shape" }) };
    }
  };
}
