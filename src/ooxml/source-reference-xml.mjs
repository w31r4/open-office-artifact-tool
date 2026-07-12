const RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELATIONSHIP_NAMESPACES = new Set([
  RELATIONSHIP_NAMESPACE,
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);

export function decodeXml(value) {
  return String(value ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

export function attrEscape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

export function attributes(tag = "") {
  return Object.fromEntries([...String(tag).matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])(.*?)\2/g)].map((match) => [match[1], decodeXml(match[3])]));
}

export function qname(prefix, localName) {
  return prefix ? `${prefix}:${localName}` : localName;
}

export function rootTag(xml, localName) {
  return new RegExp(`<(?:([A-Za-z_][\\w.-]*):)?${localName}\\b[^>]*\\/?>`).exec(String(xml));
}

export function rootPrefix(xml, localName) {
  const root = rootTag(xml, localName);
  if (!root) throw new Error(`OOXML source reference could not find root element ${localName}.`);
  return root[1] || "";
}

export function ensureNamespacePrefix(xml, rootLocalName, namespace, preferred, allowDefault = true) {
  const root = rootTag(xml, rootLocalName);
  if (!root) throw new Error(`OOXML source reference could not find root element ${rootLocalName}.`);
  const attrs = attributes(root[0]);
  for (const [name, value] of Object.entries(attrs)) {
    if (value !== namespace) continue;
    if (name === "xmlns" && allowDefault) return { xml: String(xml), prefix: "" };
    if (name.startsWith("xmlns:")) return { xml: String(xml), prefix: name.slice(6) };
  }
  let prefix = preferred;
  let index = 1;
  while (attrs[`xmlns:${prefix}`]) prefix = `${preferred}${index++}`;
  const updatedRoot = root[0].replace(/\s*\/?>$/, (ending) => ` xmlns:${prefix}="${namespace}"${ending}`);
  return { xml: `${String(xml).slice(0, root.index)}${updatedRoot}${String(xml).slice(root.index + root[0].length)}`, prefix };
}

export function ensureRelationshipPrefix(xml, rootLocalName) {
  const root = rootTag(xml, rootLocalName);
  if (!root) throw new Error(`OOXML source reference could not find root element ${rootLocalName}.`);
  const attrs = attributes(root[0]);
  for (const [name, namespace] of Object.entries(attrs)) {
    if (name.startsWith("xmlns:") && RELATIONSHIP_NAMESPACES.has(namespace)) return { xml: String(xml), prefix: name.slice(6) };
  }
  return ensureNamespacePrefix(xml, rootLocalName, RELATIONSHIP_NAMESPACE, "r", false);
}

export function relationshipId(tag = "") {
  return Object.entries(attributes(tag)).find(([name]) => /:(?:id|embed|link)$/.test(name))?.[1];
}

export function removeReferenceTags(xml, localName, ids) {
  if (!ids.size) return String(xml);
  return String(xml).replace(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*\\/?>`, "g"), (tag) => ids.has(relationshipId(tag)) ? "" : tag);
}

export function setAttribute(tag, name, value) {
  const escapedName = String(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`(\\b${escapedName}\\s*=\\s*)(["'])(.*?)\\2`);
  if (pattern.test(tag)) return tag.replace(pattern, `$1"${attrEscape(value)}"`);
  return tag.replace(/\s*\/?>$/, (ending) => ` ${name}="${attrEscape(value)}"${ending}`);
}

export function appendToRoot(xml, rootLocalName, content) {
  const root = rootTag(xml, rootLocalName);
  if (!root) throw new Error(`OOXML source reference could not find root element ${rootLocalName}.`);
  const prefix = root[1] || "";
  if (/\/\s*>$/.test(root[0])) {
    const opening = root[0].replace(/\/\s*>$/, ">");
    return `${String(xml).slice(0, root.index)}${opening}${content}</${qname(prefix, rootLocalName)}>${String(xml).slice(root.index + root[0].length)}`;
  }
  const closing = new RegExp(`</${prefix ? `${prefix}:` : ""}${rootLocalName}>`);
  if (!closing.test(String(xml))) throw new Error(`OOXML source reference could not find closing root element ${rootLocalName}.`);
  return String(xml).replace(closing, `${content}</${qname(prefix, rootLocalName)}>`);
}

export function insertBeforeOrAppend(xml, rootLocalName, content, followingLocalNames = []) {
  if (followingLocalNames.length) {
    const alternatives = followingLocalNames.join("|");
    const following = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?(?:${alternatives})\\b`).exec(String(xml));
    if (following) return `${String(xml).slice(0, following.index)}${content}${String(xml).slice(following.index)}`;
  }
  return appendToRoot(xml, rootLocalName, content);
}

export function regexEscape(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
