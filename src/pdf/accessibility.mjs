const GENERIC_ALT_TEXT = /^(?:image|picture|photo|figure|graphic|chart|diagram|icon|decorative)(?:\s+\d+)?$/i;

export function normalizePdfHeadingLevel(value) {
  if (value == null || value === "") return undefined;
  const level = Number(value);
  if (!Number.isInteger(level) || level < 1 || level > 6) throw new Error("PDF headingLevel must be an integer from 1 through 6.");
  return level;
}

export function pdfHeadingNestingIssues(sequence = []) {
  const issues = [];
  let previousLevel;
  for (const heading of sequence) {
    const level = normalizePdfHeadingLevel(heading.level);
    if (level == null) continue;
    if (previousLevel == null && level > 1) {
      issues.push({ code: "headingStartsBelowH1", id: heading.id, page: heading.page, level, message: `Heading ${heading.id || "(unknown)"} starts at H${level}; begin the logical heading sequence at H1.` });
    } else if (previousLevel != null && level > previousLevel + 1) {
      issues.push({ code: "headingLevelSkipped", id: heading.id, page: heading.page, level, previousLevel, message: `Heading ${heading.id || "(unknown)"} jumps from H${previousLevel} to H${level}.` });
    }
    previousLevel = level;
  }
  return issues;
}

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
