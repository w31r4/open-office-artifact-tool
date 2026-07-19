import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";

import {
  listSkillPngFiles,
  optimizePngBytes,
  pngIdentity,
  REPO_ROOT,
} from "../scripts/optimize-skill-pngs.mjs";

const referenceRoot = path.join(REPO_ROOT, "reference", "office-artifact-tool");
const projectFiles = await listSkillPngFiles(REPO_ROOT);
const referenceFiles = await listSkillPngFiles(referenceRoot);
const relative = (root, filename) => path.relative(root, filename).split(path.sep).join("/");
const projectPaths = projectFiles.map((filename) => relative(REPO_ROOT, filename));
const referencePaths = referenceFiles.map((filename) => relative(referenceRoot, filename));
const extraProjectPath = "skills/spreadsheets/skills/excel-live-control/assets/file-spreadsheet.png";

assert.equal(projectFiles.length, 40, "the four npm Skill bundles must retain their 40 published PNG assets");
assert.equal(referenceFiles.length, 39, "the pinned reference must retain its 39 source PNG assets");
assert.deepEqual(
  projectPaths.filter((filename) => !referencePaths.includes(filename)),
  [extraProjectPath],
  "only the compatibility Excel live-control icon may extend the reference PNG path set",
);

let projectBytes = 0;
let referenceBytes = 0;
for (const referenceFilename of referenceFiles) {
  const assetPath = relative(referenceRoot, referenceFilename);
  const projectFilename = path.join(REPO_ROOT, assetPath);
  const [referencePng, projectPng] = await Promise.all([
    fs.readFile(referenceFilename),
    fs.readFile(projectFilename),
  ]);
  referenceBytes += referencePng.length;
  projectBytes += projectPng.length;
  const referenceIdentity = pngIdentity(referencePng);
  const projectIdentity = pngIdentity(projectPng);
  assert.ok(projectIdentity.inflated.equals(referenceIdentity.inflated), `${assetPath} changed its inflated PNG scanline stream`);
  assert.ok(projectIdentity.nonIdat.equals(referenceIdentity.nonIdat), `${assetPath} changed PNG metadata or another non-IDAT chunk`);
  assert.equal(optimizePngBytes(projectPng).changed, false, `${assetPath} has remaining deterministic lossless savings`);
}

const extraPng = await fs.readFile(path.join(REPO_ROOT, extraProjectPath));
projectBytes += extraPng.length;
const canonicalSpreadsheetIcon = await fs.readFile(path.join(
  referenceRoot,
  "skills/spreadsheets/skills/spreadsheets/assets/file-spreadsheet.png",
));
const extraIdentity = pngIdentity(extraPng);
const canonicalIdentity = pngIdentity(canonicalSpreadsheetIcon);
assert.ok(extraIdentity.inflated.equals(canonicalIdentity.inflated), "Excel live-control icon changed its source pixel stream");
assert.ok(extraIdentity.nonIdat.equals(canonicalIdentity.nonIdat), "Excel live-control icon changed its source metadata");
assert.equal(optimizePngBytes(extraPng).changed, false, "Excel live-control icon has remaining deterministic lossless savings");
const canonicalOptimization = optimizePngBytes(canonicalSpreadsheetIcon);
assert.equal(canonicalOptimization.changed, true, "the pinned reference icon must exercise the write path");
const canonicalOptimizedIdentity = pngIdentity(canonicalOptimization.bytes);
assert.ok(canonicalOptimizedIdentity.inflated.equals(canonicalIdentity.inflated), "write path changed the inflated scanline stream");
assert.ok(canonicalOptimizedIdentity.nonIdat.equals(canonicalIdentity.nonIdat), "write path changed a non-IDAT chunk");
assert.equal(optimizePngBytes(canonicalOptimization.bytes).changed, false, "write path must be idempotent");

const badCrc = Buffer.from(canonicalSpreadsheetIcon);
badCrc[32] ^= 1;
assert.throws(() => pngIdentity(badCrc), /IHDR CRC mismatch/, "corrupt PNG chunks must fail closed");
assert.throws(
  () => pngIdentity(Buffer.concat([canonicalSpreadsheetIcon, Buffer.from([0])])),
  /IEND must be the final chunk with no trailing bytes/,
  "trailing data must fail closed",
);

assert.equal(referenceBytes, 4_385_190, "pinned reference PNG byte inventory changed without review");
assert.equal(projectBytes, 3_548_674, "optimized public Skill PNG byte inventory changed without review");
assert.equal(4_397_178 - projectBytes, 848_504, "expected lossless package recovery changed");

console.log("Skill PNG asset integrity ok: 40 files, 848504 bytes recovered without semantic or metadata drift");
