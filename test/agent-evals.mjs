import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  fingerprintPath,
  loadSuite,
  makeReadOnly,
  oracleFingerprint,
  removePreparedTree,
  scorePrepared,
  validateSuite,
  visibleCase,
} from "../scripts/run-agent-evals.mjs";

const { suite, cases } = await loadSuite();
assert.deepEqual(validateSuite(suite, cases), { cases: 26, pdfCases: 19, ready: 7 });
assert.equal(cases.filter((item) => item.family === "pdf" && item.status === "ready").length, 7);

const visible = visibleCase(suite, cases.find((item) => item.id === "pdf-bounded-contract-id-replace"));
assert.match(visible.prompt, /outputs\/contract-updated\.pdf/);
assert.match(visible.prompt, /outputs\/audit\.json/);
assert.doesNotMatch(visible.prompt, /expectedOutcome|oracleSha256|pymupdf\.readthedocs|"grade"/i);

const badNetwork = structuredClone(cases);
badNetwork[0].policy.network = true;
assert.throws(() => validateSuite(suite, badNetwork), /network must be false/);
const badPath = structuredClone(cases);
badPath[0].inputs[0].to = "../source.pdf";
assert.throws(() => validateSuite(suite, badPath), /escapes the workspace|under inputs/);
const normalizedEscape = structuredClone(cases);
normalizedEscape[0].deliverables[0].path = "outputs/../inputs/source.pdf";
assert.throws(() => validateSuite(suite, normalizedEscape), /deliverable must stay under outputs/);

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-agent-eval-test-"));
try {
  const missingPython = spawnSync(process.execPath, ["scripts/run-agent-evals.mjs", "prepare", "pdf-bounded-contract-id-replace", "--run-root", path.join(temporary, "missing-python")], {
    cwd: path.resolve(import.meta.dirname, ".."),
    encoding: "utf8",
    env: { ...process.env, OPEN_OFFICE_AGENT_EVAL_PYTHON: process.execPath, OPEN_OFFICE_PDF_PROVIDER_PYTHON: "" },
  });
  assert.notEqual(missingPython.status, 0);
  assert.match(missingPython.stderr, /Generated Agent eval fixtures require a Python environment with reportlab and pypdf/);
  assert.doesNotMatch(missingPython.stderr, /def base_page|contract-five-page/);

  const removable = path.join(temporary, "removable", "nested");
  await fs.mkdir(removable, { recursive: true });
  await fs.writeFile(path.join(removable, "locked.txt"), "locked");
  await makeReadOnly(path.join(temporary, "removable"));
  await removePreparedTree(path.join(temporary, "removable"));
  await assert.rejects(() => fs.access(path.join(temporary, "removable")), /ENOENT/);

  const item = cases.find((candidate) => candidate.id === "pdf-richmedia-opaque-preservation");
  const workspace = path.join(temporary, "workspace");
  const evaluator = path.join(temporary, "evaluator");
  const credentials = path.join(workspace, "inputs", "credentials");
  await fs.mkdir(path.join(workspace, "outputs"), { recursive: true });
  await fs.mkdir(credentials, { recursive: true });
  await fs.mkdir(path.join(workspace, ".agents", "skills", "pdf"), { recursive: true });
  await fs.mkdir(path.join(workspace, "node_modules", "open-office-artifact-tool"), { recursive: true });
  await fs.mkdir(evaluator, { recursive: true });
  await fs.writeFile(path.join(credentials, "signer.pem"), "test-only-key-material");
  await fs.writeFile(path.join(workspace, "PROMPT.md"), "test prompt");
  await fs.writeFile(path.join(workspace, "package.json"), "{}");
  await fs.writeFile(path.join(workspace, "package-lock.json"), "{}");
  await fs.writeFile(path.join(workspace, ".agents", "skills", "pdf", "SKILL.md"), "test skill");
  const installedModule = path.join(workspace, "node_modules", "open-office-artifact-tool", "index.mjs");
  await fs.writeFile(installedModule, "export const version = 1;\n");
  const before = await fingerprintPath(credentials);
  assert.match(before, /^tree:[0-9a-f]{64}$/);

  const workspaceHashes = {};
  for (const relative of ["PROMPT.md", "package.json", "package-lock.json", ".agents", "node_modules"]) {
    workspaceHashes[relative] = await fingerprintPath(path.join(workspace, relative));
  }

  await fs.writeFile(path.join(workspace, "outputs", "audit.json"), JSON.stringify({ status: "failed_closed", delivered_modified_pdf: false }));
  await fs.writeFile(path.join(evaluator, "run.json"), JSON.stringify({ oracleSha256: oracleFingerprint(item), workspaceHashes, inputHashes: { "inputs/credentials": before } }));
  await fs.writeFile(path.join(evaluator, "exit.json"), JSON.stringify({ status: 0, signal: null }));
  await fs.writeFile(path.join(evaluator, "final.txt"), "Opaque RichMedia runtime behavior cannot be proven; failed closed.");
  await fs.writeFile(path.join(evaluator, "trace.jsonl"), "{\"type\":\"result\"}\n");

  const prepared = { workspace, evaluator };
  const refusal = await scorePrepared(item, prepared);
  assert.equal(refusal.observedOutcome, "safe-refusal");
  assert.equal(refusal.hardGatesPassed, true);

  await fs.writeFile(installedModule, "export const version = 2;\n");
  const alteredRuntime = await scorePrepared(item, prepared);
  assert.equal(alteredRuntime.checks.find((check) => check.id === "workspace-immutable:node_modules")?.passed, false);
  assert.equal(alteredRuntime.hardGatesPassed, false);
  await fs.writeFile(installedModule, "export const version = 1;\n");

  await fs.writeFile(path.join(credentials, "signer.pem"), "mutated-key-material");
  const mutated = await scorePrepared(item, prepared);
  assert.equal(mutated.checks.find((check) => check.id === "source-immutable:inputs/credentials")?.passed, false);
  assert.equal(mutated.hardGatesPassed, false);
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}

console.log("agent eval suite smoke ok");
