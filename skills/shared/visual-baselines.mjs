import fs from "node:fs/promises";
import path from "node:path";

import { FileBlob, visualQaArtifact } from "open-office-artifact-tool";

function missingBaselineError(target) {
  return new Error(`Visual baseline is missing: ${target}. Run the workflow with writeBaseline=true first.`);
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export async function loadVisualBaseline(baselinePath, options = {}) {
  if (!baselinePath || options.writeBaseline === true) return undefined;
  try {
    return await FileBlob.load(baselinePath);
  } catch (error) {
    if (error?.code === "ENOENT") throw missingBaselineError(baselinePath);
    throw error;
  }
}

export async function prepareNumberedVisualBaselines(baselineDir, prefix, options = {}) {
  if (!baselineDir) return { files: [], expectedCount: undefined };
  const pattern = new RegExp(`^${escapeRegExp(prefix)}-(\\d+)\\.png$`);
  let entries = [];
  try {
    entries = await fs.readdir(baselineDir);
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
  const files = entries
    .map((name) => ({ name, match: pattern.exec(name) }))
    .filter((entry) => entry.match)
    .sort((left, right) => Number(left.match[1]) - Number(right.match[1]))
    .map((entry) => path.join(baselineDir, entry.name));
  if (options.writeBaseline === true) {
    await fs.mkdir(baselineDir, { recursive: true });
    await Promise.all(files.map((filePath) => fs.unlink(filePath)));
    return { files: [], expectedCount: undefined };
  }
  if (files.length === 0) throw missingBaselineError(path.join(baselineDir, `${prefix}-N.png`));
  return { files, expectedCount: files.length };
}

export function visualBaselineCountResult(baselineSet, actualCount, options = {}) {
  const baselinePageCount = baselineSet?.expectedCount;
  const pageCountMatches = baselinePageCount == null || baselinePageCount === actualCount;
  const issue = pageCountMatches ? undefined : JSON.stringify({
    kind: "visualPageCountDiff",
    artifactKind: options.artifactKind,
    severity: "warning",
    pageCount: actualCount,
    baselinePageCount,
    ...(options.baselineKind ? { baselineKind: options.baselineKind } : {}),
  });
  return { baselinePageCount, pageCountMatches, issue };
}

export async function runPngVisualQa(artifact, options = {}) {
  const baseline = await loadVisualBaseline(options.baselinePath, options);
  const qa = await visualQaArtifact(artifact, {
    ...options.renderOptions,
    format: "png",
    renderer: options.renderer,
    baseline,
    pixelDiff: Boolean(baseline),
    pixelThreshold: options.pixelThreshold ?? 0,
    diffAlignment: options.diffAlignment,
    diffPalette: options.diffPalette,
    pixelRegistration: options.pixelRegistration,
    minBytes: options.minBytes ?? 100,
    maxChars: options.maxChars ?? 20_000,
  });
  if (options.writeBaseline === true && options.baselinePath) {
    await fs.mkdir(path.dirname(options.baselinePath), { recursive: true });
    await qa.blob.save(options.baselinePath);
  }
  if (qa.diffBlob && options.diffPath) {
    await fs.mkdir(path.dirname(options.diffPath), { recursive: true });
    await qa.diffBlob.save(options.diffPath);
    qa.diffPath = options.diffPath;
  }
  return qa;
}
