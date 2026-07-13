const GENERIC_ALT_TEXT = /^(?:image|picture|photo|figure|graphic|chart|diagram|icon|decorative)(?:\s+\d+)?$/i;

export function normalizePdfFigureAccessibility(config = {}) {
  const altValue = config.alt ?? config.altText;
  return {
    alt: altValue == null ? undefined : String(altValue).trim(),
    decorative: Boolean(config.decorative ?? config.artifact),
  };
}

export function pdfFigureAccessibilityIssue(figure = {}, kind = "figure") {
  if (figure.decorative) return undefined;
  const alt = String(figure.alt || "").trim();
  if (!alt) return { code: "missingFigureAltText", message: `${kind} ${figure.id || "(unknown)"} requires alternative text or decorative=true.` };
  if (GENERIC_ALT_TEXT.test(alt)) return { code: "genericFigureAltText", message: `${kind} ${figure.id || "(unknown)"} uses generic alternative text ${JSON.stringify(alt)}.` };
  return undefined;
}

export function inspectPdfFigureAccessibility(pdfText = "") {
  const structureBodies = [...String(pdfText).matchAll(/(?:^|\n)\d+\s+0\s+obj\s*([\s\S]*?)\s*endobj/g)]
    .map((match) => match[1])
    .filter((body) => /\/Type\s*\/StructElem\b/.test(body));
  const figures = structureBodies.filter((body) => /\/S\s*\/Figure\b/.test(body));
  const figureAltTexts = figures.filter((body) => /\/Alt\s*(?:\((?:\\.|[^\\)])+\)|<[A-Fa-f0-9]+>)/.test(body)).length;
  const artifacts = [...String(pdfText).matchAll(/\/Artifact\s+BMC\b/g)].length;
  return { figures: figures.length, figureAltTexts, missingFigureAltTexts: Math.max(0, figures.length - figureAltTexts), artifacts };
}
