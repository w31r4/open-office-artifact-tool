export function xmlEscape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

export function attrEscape(value) {
  return xmlEscape(value).replaceAll('"', "&quot;");
}
