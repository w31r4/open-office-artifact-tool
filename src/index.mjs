import { DocumentFile, DocumentModel } from "./document/index.mjs";
import { PdfArtifact, PdfFile } from "./pdf/index.mjs";
import { ChartElement, GroupShape, ImageElement, Presentation, PresentationFile, Shape, Slide, TableElement } from "./presentation/index.mjs";
import { Range, SpreadsheetFile, Workbook, Worksheet } from "./spreadsheet/index.mjs";
import { queryHelpRecords } from "./help/index.mjs";
import { createArtifactVisualQaApi } from "./qa/artifact-visual.mjs";
import { ndjson, verificationIssue, verificationResult } from "./shared/inspection.mjs";

export { FileBlob } from "./shared/file-blob.mjs";
export { HELP_CATALOG } from "./help/index.mjs";
export { box, chart, column, grid, image, layers, node, paragraph, row, rule, run, shape, table, text } from "./presentation/compose.mjs";
export { ChartElement, DocumentFile, DocumentModel, GroupShape, ImageElement, PdfArtifact, PdfFile, Presentation, PresentationFile, Range, Shape, Slide, SpreadsheetFile, TableElement, Workbook, Worksheet };

function inferArtifactKind(artifact) {
  if (artifact instanceof Workbook) return "workbook";
  if (artifact instanceof Presentation) return "presentation";
  if (artifact instanceof DocumentModel) return "document";
  if (artifact instanceof PdfArtifact) return "pdf";
  return "unknown";
}

export function verifyArtifact(artifact, options = {}) {
  if (!artifact || typeof artifact.verify !== "function") {
    return verificationResult("unknown", [verificationIssue("unknown", "unsupportedArtifact", "Artifact does not expose a verify() method.")], options);
  }
  return artifact.verify(options);
}

const artifactVisualQaApi = createArtifactVisualQaApi({ inferArtifactKind });

export async function renderArtifact(artifact, options = {}) {
  return artifactVisualQaApi.renderArtifact(artifact, options);
}

export async function visualQaArtifact(artifact, options = {}) {
  return artifactVisualQaApi.visualQaArtifact(artifact, options);
}

export function helpArtifact(artifactOrKind = "*", query = "*", options = {}) {
  const artifactKind = typeof artifactOrKind === "string" ? artifactOrKind : inferArtifactKind(artifactOrKind);
  return ndjson(queryHelpRecords(artifactKind, query, options), options.maxChars ?? Infinity);
}
