import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  extractCompletedCommands,
  gradeBoundedReplaceEvidence,
  gradeOverflowRefusalEvidence,
  summarizeCaseScore,
} from "../scripts/agent-eval-pdf-graders.mjs";
import {
  fingerprintPath,
  loadSuite,
  makeReadOnly,
  oracleFingerprint,
  removePreparedTree,
  repositoryProvenance,
  scorePrepared,
  validateSuite,
  visibleCase,
} from "../scripts/run-agent-evals.mjs";

const { suite, cases } = await loadSuite();
assert.deepEqual(validateSuite(suite, cases), { cases: 26, pdfCases: 19, ready: 7 });
assert.equal(cases.filter((item) => item.family === "pdf" && item.status === "ready").length, 7);

const repository = repositoryProvenance();
assert.match(repository.head, /^[0-9a-f]{40}$/);
assert.equal(typeof repository.dirty, "boolean");
assert.match(repository.statusSha256, /^[0-9a-f]{64}$/);
assert.match(repository.trackedDiffSha256, /^[0-9a-f]{64}$/);

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

const boundedItem = cases.find((item) => item.id === "pdf-bounded-contract-id-replace");
const oldContractId = "ACME-2025-041";
const newContractId = "ACME-2026-041";
const sourcePages = Array.from({ length: 5 }, (_, index) => ({
  page: index + 1,
  width: 612,
  height: 792,
  rotation: 0,
  termCounts: { [oldContractId]: index === 2 ? 1 : 0, [newContractId]: 0 },
}));
const outputPages = sourcePages.map((page, index) => ({ ...page, termCounts: { [oldContractId]: 0, [newContractId]: index === 2 ? 1 : 0 } }));
const boundedEvidence = {
  source: { sha256: "source-sha", pageCount: 5, pages: sourcePages, termCounts: { [oldContractId]: 1, [newContractId]: 0 }, decodedStreamErrors: [] },
  output: {
    sha256: "output-sha",
    pageCount: 5,
    pages: outputPages,
    termCounts: { [oldContractId]: 0, [newContractId]: 1 },
    rawTermCounts: { [oldContractId]: 0 },
    decodedStreamTermCounts: { [oldContractId]: 0 },
    metadataTermCounts: { [oldContractId]: 0 },
    decodedStreamErrors: [],
    startxrefCount: 1,
    eofCount: 1,
  },
  sourceStyle: { found: true, fonts: ["Helvetica"], sizes: [11], bbox: [172, 152, 254, 167] },
  outputStyle: { found: true, fonts: ["Helvetica"], sizes: [11], bbox: [172, 152, 254, 167] },
  visual: {
    renderer: "poppler-pdftoppm",
    sourcePageCount: 5,
    outputPageCount: 5,
    allowedMask: { page: 3, bboxPx: [330, 296, 470, 344] },
    pages: Array.from({ length: 5 }, (_, index) => ({
      page: index + 1,
      sameDimensions: true,
      nonBlank: true,
      changedPixelsBBox: index === 2 ? [399, 312, 411, 329] : null,
      changedWithinAllowedMask: index === 2,
    })),
  },
};
const boundedAudit = {
  status: "succeeded",
  source: { sha256: "source-sha" },
  output: { sha256: "output-sha" },
  provider: { actual: "pymupdf", version: "1.27.2.3", silentFallback: false },
  savePolicy: { strategy: "sanitize" },
  operation: { type: "replace_text" },
};
const boundedCommands = [
  "python .agents/skills/pdf/scripts/pymupdf_edit.py probe --accept-license agpl",
  "python .agents/skills/pdf/scripts/pymupdf_edit.py edit inputs/source.pdf outputs/contract-updated.pdf",
];
const boundedChecks = gradeBoundedReplaceEvidence({ evidence: boundedEvidence, audit: boundedAudit, commands: boundedCommands, item: boundedItem });
assert.equal(boundedChecks.every((entry) => entry.passed), true);
assert.deepEqual(summarizeCaseScore(boundedChecks, boundedItem.grade), {
  categoryScores: {
    machine: { applicable: true, weight: 45, passed: true, checks: 7 },
    visual: { applicable: true, weight: 25, passed: true, checks: 3 },
    security: { applicable: true, weight: 20, passed: true, checks: 4 },
    trace: { applicable: true, weight: 10, passed: true, checks: 7 },
  },
  rawScorePercent: 100,
  scorePercent: 100,
  caseSpecificPassed: true,
});
const undecodableEvidence = structuredClone(boundedEvidence);
undecodableEvidence.output.decodedStreamErrors.push({ object: 9, generation: 0, error: "test decode failure" });
const undecodableChecks = gradeBoundedReplaceEvidence({ evidence: undecodableEvidence, audit: boundedAudit, commands: boundedCommands, item: boundedItem });
assert.equal(undecodableChecks.find((entry) => entry.id === "pdf-security:all-streams-decodable")?.passed, false);
const shiftedEvidence = structuredClone(boundedEvidence);
shiftedEvidence.outputStyle.bbox[0] += 1;
const shiftedChecks = gradeBoundedReplaceEvidence({ evidence: shiftedEvidence, audit: boundedAudit, commands: boundedCommands, item: boundedItem });
assert.equal(shiftedChecks.find((entry) => entry.id === "pdf-machine:font-geometry-unchanged")?.passed, false);

const bypassAudit = structuredClone(boundedAudit);
bypassAudit.savePolicy.strategy = "rewrite";
bypassAudit.operation.type = "direct_content_stream_equal_length_replace";
const bypassChecks = gradeBoundedReplaceEvidence({ evidence: boundedEvidence, audit: bypassAudit, commands: [...boundedCommands, "doc.update_stream(xref, data.replace(old, new))"], item: boundedItem });
assert.equal(bypassChecks.filter((entry) => entry.category !== "trace").every((entry) => entry.passed), true);
assert.equal(bypassChecks.find((entry) => entry.id === "pdf-trace:no-content-stream-bypass")?.passed, false);
assert.equal(summarizeCaseScore(bypassChecks, boundedItem.grade).rawScorePercent, 90);
const postHocProbeChecks = gradeBoundedReplaceEvidence({
  evidence: boundedEvidence,
  audit: boundedAudit,
  commands: [boundedCommands[1], boundedCommands[0]],
  item: boundedItem,
});
assert.equal(postHocProbeChecks.find((entry) => entry.id === "pdf-trace:capability-probe")?.passed, false);

const overflowItem = cases.find((item) => item.id === "pdf-overflow-replace-refusal");
const overflowEvidence = {
  source: { sha256: "overflow-source-sha", termCounts: { Approved: 1 }, pages: [{ termCounts: { Approved: 1 } }] },
  geometry: { targetRectangle: { width: 70 }, replacementWidth: 291.027, availableWidth: 64, fits: false },
};
const overflowAudit = {
  status: "failed_closed",
  delivered_modified_pdf: false,
  reason: "replacement overflows the available box",
  source: { sha256: "overflow-source-sha" },
  provider: { selected: "PyMuPDF", version: "1.27.2.3", fallback_used: false },
  operation: { performed: "no PDF mutation; failed before edit" },
  geometry_evidence: { replacement_text_width_pt: 291.027, available_width_from_text_origin_pt: 64, fits_outer_cell_without_wrap: false },
  validation: { no_partial_modified_pdf_in_outputs: true },
};
const overflowChecks = gradeOverflowRefusalEvidence({
  evidence: overflowEvidence,
  audit: overflowAudit,
  commands: ["python -c 'import fitz; print(fitz.VersionBind)'"],
  finalMessage: "The replacement would overflow and does not fit.",
  item: overflowItem,
});
assert.equal(overflowChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(overflowChecks, overflowItem.grade).rawScorePercent, 100);
const mutatedOverflowChecks = gradeOverflowRefusalEvidence({
  evidence: overflowEvidence,
  audit: overflowAudit,
  commands: ["python -c 'import fitz; doc.save(\"outputs/partial.pdf\")'"],
  finalMessage: "The replacement would overflow and does not fit.",
  item: overflowItem,
});
assert.equal(mutatedOverflowChecks.find((entry) => entry.id === "pdf-trace:no-mutation-after-failed-preflight")?.passed, false);

const traceCommands = extractCompletedCommands([
  JSON.stringify({ type: "item.started", item: { id: "one", type: "command_execution", command: "ignored-started-command" } }),
  JSON.stringify({ type: "item.completed", item: { id: "one", type: "command_execution", command: "echo safe", aggregated_output: "doc.update_stream(xref, bytes)" } }),
].join("\n"));
assert.deepEqual(traceCommands, ["echo safe"]);

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
  const preparedSkillMode = (await fs.stat(path.join(temporary, "missing-python", "pdf-bounded-contract-id-replace", "candidate-trial-1", "workspace", ".agents", "skills", "pdf", "SKILL.md"))).mode;
  assert.equal(preparedSkillMode & 0o222, 0);

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

  await fs.chmod(path.join(credentials, "signer.pem"), 0o600);
  const alteredInputMode = await scorePrepared(item, prepared);
  assert.equal(alteredInputMode.checks.find((check) => check.id === "source-immutable:inputs/credentials")?.passed, false);
  await fs.chmod(path.join(credentials, "signer.pem"), 0o644);

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
  await removePreparedTree(temporary);
}

console.log("agent eval suite smoke ok");
