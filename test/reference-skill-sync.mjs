import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";

import {
  checkReferenceSkillSync,
  createReferenceSkillSnapshot,
  REFERENCE_SKILL_BUNDLES,
} from "../scripts/reference-skill-sync.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const checked = await checkReferenceSkillSync();
const rebuilt = await createReferenceSkillSnapshot();
const recorded = JSON.parse(await fs.readFile(path.join(repoRoot, "skills", "reference-sync.json"), "utf8"));

assert.deepEqual(checked, rebuilt);
assert.deepEqual(rebuilt, recorded);
assert.equal(rebuilt.schemaVersion, 1);
assert.equal(rebuilt.source.commit, "256cb31bfe0a07b3cef0051b6b159342be381378");
assert.equal(rebuilt.totalFiles, 343);
assert.deepEqual(Object.keys(rebuilt.bundles), REFERENCE_SKILL_BUNDLES);
assert.equal(Object.values(rebuilt.bundles).reduce((sum, bundle) => sum + bundle.files, 0), rebuilt.totalFiles);
assert.equal(Object.values(rebuilt.bundles).reduce((sum, bundle) => sum + bundle.bytes, 0), rebuilt.totalBytes);

const referenceChecklist = await fs.readFile(path.join(
  repoRoot,
  "reference",
  "office-artifact-tool",
  "skills",
  "documents",
  "skills",
  "documents",
  "examples",
  "end_to_end_smoke_test.md",
));
const projectChecklist = await fs.readFile(path.join(
  repoRoot,
  "skills",
  "documents",
  "skills",
  "documents",
  "examples",
  "end_to_end_smoke_test.md",
));
assert.deepEqual(projectChecklist, referenceChecklist, "the newly synchronized reference checklist must remain byte-identical");

console.log(`reference Skill source sync ok: ${rebuilt.totalFiles} files at ${rebuilt.source.commit}`);
