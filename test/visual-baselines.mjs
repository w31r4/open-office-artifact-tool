import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { FileBlob } from "open-office-artifact-tool";
import {
  loadVisualBaseline,
  prepareNumberedVisualBaselines,
  runPngVisualQa,
  visualBaselineCountResult,
} from "../skills/shared/visual-baselines.mjs";

const whitePixelPng = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFgAI/ScL+YQAAAABJRU5ErkJggg==", "base64");
const root = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-visual-baselines-"));
const baselineDir = path.join(root, "baselines");

try {
  await assert.rejects(
    () => loadVisualBaseline(path.join(baselineDir, "model.png")),
    /Visual baseline is missing.*writeBaseline=true/,
  );
  await assert.rejects(
    () => prepareNumberedVisualBaselines(baselineDir, "native-page"),
    /Visual baseline is missing.*native-page-N\.png/,
  );

  const artifact = { render: () => new FileBlob(whitePixelPng, { type: "image/png" }) };
  const modelPath = path.join(baselineDir, "model.png");
  const written = await runPngVisualQa(artifact, { baselinePath: modelPath, writeBaseline: true, minBytes: 1 });
  assert.equal(written.ok, true);
  assert.ok((await fs.stat(modelPath)).size > 1);
  const compared = await runPngVisualQa(artifact, { baselinePath: modelPath, minBytes: 1 });
  assert.equal(compared.ok, true);
  assert.equal(compared.summary.baselineHash, compared.summary.hash);

  await Promise.all([
    fs.writeFile(path.join(baselineDir, "native-page-2.png"), whitePixelPng),
    fs.writeFile(path.join(baselineDir, "native-page-1.png"), whitePixelPng),
    fs.writeFile(path.join(baselineDir, "keep.txt"), "keep", "utf8"),
  ]);
  const baselineSet = await prepareNumberedVisualBaselines(baselineDir, "native-page");
  assert.deepEqual(baselineSet.files.map((filePath) => path.basename(filePath)), ["native-page-1.png", "native-page-2.png"]);
  assert.deepEqual(visualBaselineCountResult(baselineSet, 2, { artifactKind: "document" }), {
    baselinePageCount: 2,
    pageCountMatches: true,
    issue: undefined,
  });
  const mismatch = visualBaselineCountResult(baselineSet, 1, { artifactKind: "document", baselineKind: "native" });
  assert.equal(mismatch.pageCountMatches, false);
  assert.match(mismatch.issue, /"baselinePageCount":2/);
  await prepareNumberedVisualBaselines(baselineDir, "native-page", { writeBaseline: true });
  assert.equal((await fs.readFile(path.join(baselineDir, "keep.txt"), "utf8")), "keep");
  await assert.rejects(() => fs.stat(path.join(baselineDir, "native-page-1.png")), { code: "ENOENT" });
  await Promise.all([
    fs.writeFile(path.join(baselineDir, "native-page-1.png"), whitePixelPng),
    fs.writeFile(path.join(baselineDir, "native-page-3.png"), whitePixelPng),
  ]);
  await assert.rejects(
    () => prepareNumberedVisualBaselines(baselineDir, "native-page"),
    /numbered continuously from 1/,
  );
} finally {
  await fs.rm(root, { recursive: true, force: true });
}

console.log("visual baselines smoke ok");
