const GENERIC_LINK_TEXT = /^(?:click here|here|link|learn more|more|read more|查看|点击这里|链接)$/i;
const SAFE_LINK_PROTOCOLS = new Set(["http:", "https:", "mailto:"]);

export function normalizePdfLink(config = {}, fallbackId) {
  return {
    id: String(config.id || fallbackId),
    text: String(config.text ?? config.label ?? config.title ?? "").trim(),
    url: String(config.url ?? config.uri ?? config.href ?? "").trim(),
    bbox: config.bbox || config.bounds || [72, 700, 240, 16],
  };
}

export function pdfLinkIssues(link = {}) {
  const issues = [];
  const text = String(link.text || "").trim();
  const url = String(link.url || "").trim();
  if (!text) issues.push({ code: "emptyLinkText", message: `PDF link ${link.id || "(unknown)"} requires visible meaningful text.` });
  else if (GENERIC_LINK_TEXT.test(text) || text === url) issues.push({ code: "genericLinkText", message: `PDF link ${link.id || "(unknown)"} requires text that explains its destination.` });
  try {
    const parsed = new URL(url);
    if (!SAFE_LINK_PROTOCOLS.has(parsed.protocol)) issues.push({ code: "unsafeLinkProtocol", message: `PDF link ${link.id || "(unknown)"} uses unsupported protocol ${parsed.protocol || "(none)"}.` });
  } catch {
    issues.push({ code: "invalidLinkUrl", message: `PDF link ${link.id || "(unknown)"} requires an absolute http, https, or mailto URL.` });
  }
  return issues;
}

export function pdfLinkAnnotationRect(page, link) {
  const [left, top, width, height] = (link.bbox || []).map(Number);
  return [left, Number(page.height) - top - height, left + width, Number(page.height) - top];
}

export function inspectPdfLinks(pdfText = "") {
  const text = String(pdfText);
  return {
    linkAnnotations: [...text.matchAll(/\/Subtype\s*\/Link\b/g)].length,
    uriActions: [...text.matchAll(/\/S\s*\/URI\b/g)].length,
    linkStructParents: [...text.matchAll(/\/Subtype\s*\/Link\b[\s\S]*?\/StructParent\s+\d+/g)].length,
  };
}
