import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  extractCompletedCommands,
  gradeAcroFormEvidence,
  gradeAccessibleReportEvidence,
  gradeActiveContentSanitizeEvidence,
  gradeAttachmentQuarantineEvidence,
  gradeBoundedReplaceEvidence,
  gradeMergeStampEvidence,
  gradeOverflowRefusalEvidence,
  summarizeCaseScore,
} from "../scripts/agent-eval-pdf-graders.mjs";
import {
  fingerprintPath,
  loadSuite,
  makeReadOnly,
  oracleFingerprint,
  providerRuntimeInstruction,
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

const accessibleItem = cases.find((item) => item.id === "pdf-greenfield-accessible-report");
const accessiblePages = Array.from({ length: 6 }, (_, index) => ({ page: index + 1, width: 1224, height: 1584, nonBlank: true, inkBBox: [50, 50, 1100, 1500], touchesEdge: false, bytes: 20_000 }));
const accessibleEvidence = {
  source: { sha256: "accessible-source-sha" },
  output: { sha256: "accessible-output-sha", pageCount: 6 },
  structure: {
    tagged: true,
    language: "zh-CN",
    title: "Agent Artifact Readiness",
    roles: { H1: 1, H2: 4, H3: 7, Table: 1, TR: 5, TH: 6, TD: 9, Figure: 1, Link: 1 },
    tables: [{ id: "risk-register", pages: [3, 4], rows: 5, headers: 6, dataCells: 9 }],
    figuresWithAlt: 1,
    links: [{ page: 6, uri: "https://www.w3.org/WAI/", structParent: 6 }],
    linkObjrAssociations: 1,
    artifactMarkers: 12,
    rootIds: ["board-page-1/text", "summary-h2", "summary-h3", "format-pass-rate-chart", "risks-h2", "risks-h3", "risk-register", "mitigation-h3", "validation-h2", "modeled-h3", "machine-h3", "human-h3", "conclusion-h2", "conclusion-h3", "wai-guidance-link"],
    pageText: ["封面", "摘要", "风险 级别 缓解措施", "风险 级别 缓解措施", "验证", "结论"],
  },
  visual: { renderer: "poppler-pdftoppm", pageCount: 6, pages: accessiblePages },
};
const accessibleAudit = {
  status: "succeeded",
  source: { sha256: "accessible-source-sha" },
  output: { sha256: "accessible-output-sha" },
  provider: { actual: "artifact-tool", version: "0.2.0", silentFallback: false },
  savePolicy: { strategy: "rewrite" },
  operation: { type: "create-accessible-report" },
  validation: {
    modeledVerify: { status: "passed", scope: "PdfArtifact modeled invariants" },
    poppler: { status: "passed", pages: accessiblePages },
    veraPdfMachine: { available: false, status: "not-run", claim: "No veraPDF machine validation was performed." },
    humanPdfUa: { status: "required", claim: "No complete PDF/UA certification is claimed." },
  },
};
const accessibleChecks = gradeAccessibleReportEvidence({
  evidence: accessibleEvidence,
  audit: accessibleAudit,
  commands: ["node .agents/skills/pdf/examples/accessible-board-report.mjs inputs/report-data.json outputs/readiness-report.pdf outputs/audit.json"],
  finalMessage: "已创建；不声明完整 PDF/UA 认证。",
  item: accessibleItem,
});
assert.equal(accessibleChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(accessibleChecks, accessibleItem.grade).rawScorePercent, 100);
const overclaimChecks = gradeAccessibleReportEvidence({
  evidence: accessibleEvidence,
  audit: accessibleAudit,
  commands: ["node .agents/skills/pdf/examples/accessible-board-report.mjs inputs/report-data.json outputs/readiness-report.pdf outputs/audit.json"],
  finalMessage: "The report is PDF/UA certified.",
  item: accessibleItem,
});
assert.equal(overclaimChecks.find((entry) => entry.id === "pdf-security:no-pdfua-overclaim")?.passed, false);

const mergeItem = cases.find((item) => item.id === "pdf-merge-reorder-stamp-links");
const mergeSequence = ["cover:1", "appendix:3", "report:1", "report:2", "appendix:1", "appendix:2"];
const geometryBySource = {
  cover: { boxes: { mediabox: [0, 0, 612, 792] }, rotation: 0 },
  report: { boxes: { mediabox: [0, 0, 792, 612] }, rotation: 0 },
  appendix: { boxes: { mediabox: [0, 0, 595.2756, 841.8898] }, rotation: 0 },
};
const mergePageMap = mergeSequence.map((identity, index) => {
  const [source, sourcePage] = identity.split(":");
  return {
    outputPage: index + 1,
    source,
    sourcePage: Number(sourcePage),
    sourceGeometry: geometryBySource[source],
    outputGeometry: structuredClone(geometryBySource[source]),
    watermarkCount: [3, 4].includes(index + 1) ? 1 : 0,
    opacities: [3, 4].includes(index + 1) ? [0.2] : [],
  };
});
const mergeOutlines = mergeSequence.map((identity, index) => {
  const [source, sourcePage] = identity.split(":");
  return { title: `${source}-${sourcePage}`, page: index + 1, parentPath: [] };
});
const mergeNamed = Object.fromEntries(mergeOutlines.map((entry) => [`${entry.title}-named`, entry.page]));
const mergeLinks = mergeOutlines.map((entry, index) => ({ page: entry.page, targetPage: mergeOutlines[(index + 1) % mergeOutlines.length].page, rect: [68, 100, 280, 124] }));
const mergePathRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-eval-merge-paths-"));
const mergePathAlias = `${mergePathRoot}-alias`;
await fs.symlink(mergePathRoot, mergePathAlias);
const mergeSources = Object.fromEntries([
  ["cover", 1], ["report", 2], ["appendix", 3],
].map(([id, pageCount]) => [id, { path: path.join(mergePathRoot, `${id}.pdf`), bytes: 100 + pageCount, sha256: `${id}-sha`, pageCount, termCounts: { CONFIDENTIAL: 0 } }]));
await Promise.all(Object.values(mergeSources).map((source) => fs.writeFile(source.path, "fixture", "utf8")));
const mergeEvidence = {
  manifest: {
    path: "/tmp/merge-stamp.json",
    bytes: 400,
    sha256: "manifest-sha",
    value: {
      schema: "open-office-artifact-tool.pdf-merge-stamp.v1",
      sources: [{ id: "cover" }, { id: "report" }, { id: "appendix" }],
      sequence: [
        { source: "cover", pages: "all" },
        { source: "appendix", pages: [3] },
        { source: "report", pages: "all" },
        { source: "appendix", pages: [1, 2] },
      ],
      watermarks: [{ source: "report", text: "CONFIDENTIAL", opacity: 0.2 }],
    },
  },
  sources: mergeSources,
  output: { sha256: "merge-output-sha", pageCount: 6, startxrefCount: 1, eofCount: 1, decodedStreamErrors: [] },
  pageMap: mergePageMap,
  navigation: { expected: { outlines: mergeOutlines, namedDestinations: mergeNamed, internalLinks: mergeLinks }, actual: { outlines: structuredClone(mergeOutlines), namedDestinations: Object.fromEntries(Object.entries(mergeNamed).reverse()), internalLinks: structuredClone(mergeLinks) } },
  visual: { renderer: "poppler-pdftoppm", pageCount: 6, pages: mergePageMap.map((entry) => ({ page: entry.outputPage, source: entry.source, sourcePage: entry.sourcePage, sameDimensions: true, nonBlank: true, pixelStable: ![3, 4].includes(entry.outputPage), changedPixelsBBox: [3, 4].includes(entry.outputPage) ? [400, 300, 800, 700] : null, watermarkExpected: [3, 4].includes(entry.outputPage) })) },
};
const mergeAudit = {
  status: "succeeded",
  source: { sha256: "manifest-sha" },
  inputs: Object.values(mergeSources).map(({ path: sourcePath, bytes, sha256 }) => ({ path: path.join(mergePathAlias, path.basename(sourcePath)), bytes, sha256 })),
  output: { sha256: "merge-output-sha" },
  provider: { actual: "pypdf", version: "6.10.0", silentFallback: false },
  savePolicy: { strategy: "rewrite" },
  preflight: { probeCompleted: true, planCompleted: true },
  operation: { type: "merge-stamp" },
};
const mergeCommands = [
  "python -c 'from reportlab.pdfgen import canvas; print(canvas)'",
  `"$PYTHON_BIN" "$S/pdf_provider.py" check --provider pypdf --require`,
  `"$PYTHON_BIN" "$S/pdf_provider.py" plan --task merge-stamp --provider pypdf --strategy rewrite --input merge-stamp.json --output outputs/merged.pdf --require-provider`,
  `"$PYTHON_BIN" "$S/pypdf_edit.py" merge-stamp merge-stamp.json outputs/merged.pdf --strategy rewrite`,
  `"$PYTHON_BIN" "$S/poppler_compare.py" merge-stamp merge-stamp.json outputs/merged.pdf --report merge-visual.json --render-dir merge-rendered`,
  `"$PYTHON_BIN" "$S/pdf_audit.py" validate outputs/audit.json --source merge-stamp.json --input inputs/cover.pdf --input inputs/report.pdf --input inputs/appendix.pdf --artifact outputs/merged.pdf --require-operation merge-stamp`,
];
const mergeChecks = gradeMergeStampEvidence({ evidence: mergeEvidence, audit: mergeAudit, commands: mergeCommands, item: mergeItem });
assert.equal(mergeChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(mergeChecks, mergeItem.grade).rawScorePercent, 100);
const danglingMergeEvidence = structuredClone(mergeEvidence);
danglingMergeEvidence.navigation.actual.internalLinks[0].targetPage = 99;
const danglingMergeChecks = gradeMergeStampEvidence({ evidence: danglingMergeEvidence, audit: mergeAudit, commands: mergeCommands, item: mergeItem });
assert.equal(danglingMergeChecks.find((entry) => entry.id === "pdf-security:navigation-resolved")?.passed, false);
const adHocMergeChecks = gradeMergeStampEvidence({ evidence: mergeEvidence, audit: mergeAudit, commands: [...mergeCommands, "python -c 'from pypdf import PdfWriter; PdfWriter()'"], item: mergeItem });
assert.equal(adHocMergeChecks.find((entry) => entry.id === "pdf-trace:no-ad-hoc-pdf-writer")?.passed, false);
await fs.rm(mergePathAlias);
await fs.rm(mergePathRoot, { recursive: true, force: true });

const providerInstruction = providerRuntimeInstruction(mergeItem, { OPEN_OFFICE_AGENT_EVAL_PYTHON: "/opt/eval python/bin/python3" });
assert.match(providerInstruction, /OPEN_OFFICE_PDF_PROVIDER_PYTHON="\/opt\/eval python\/bin\/python3"/);
assert.match(providerInstruction, /Do not replace it/);
assert.equal(providerRuntimeInstruction({ family: "xlsx" }, { OPEN_OFFICE_AGENT_EVAL_PYTHON: "/opt/python" }), "");

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
  preflight: { probeCompleted: true, planCompleted: true },
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

const formItem = cases.find((item) => item.id === "pdf-acroform-visible-preserved");
const formFieldNames = Object.keys(formItem.grade.machine.fields);
const formSourceFields = Object.fromEntries(formFieldNames.map((name) => [name, {
  fieldType: name === "company_type" ? "/Btn" : "/Tx",
  value: "",
  defaultValue: "",
  readOnly: false,
  states: name === "company_type" ? ["/LLC", "/Corporation"] : [],
}]));
formSourceFields.terms_ack = { fieldType: "/Btn", value: "/Yes", defaultValue: "", readOnly: false, states: ["/Off", "/Yes"] };
const formOutputFields = structuredClone(formSourceFields);
for (const [name, value] of Object.entries(formItem.grade.machine.fields)) {
  formOutputFields[name].value = name === "company_type" ? `/${value}` : value;
}
const formWidgets = [
  ["full_name", "/Tx", [190, 88, 450, 112], []],
  ["address", "/Tx", [190, 133, 450, 157], []],
  ["effective_date", "/Tx", [190, 178, 450, 202], []],
  ["tin", "/Tx", [190, 223, 450, 247], []],
  ["signature", "/Tx", [190, 268, 450, 292], []],
  ["company_type", "/Btn", [190, 317, 210, 337], ["/LLC", "/Off"]],
  ["company_type", "/Btn", [260, 317, 280, 337], ["/Corporation", "/Off"]],
  ["terms_ack", "/Btn", [260, 362, 280, 382], ["/Off", "/Yes"]],
].map(([name, fieldType, rect, appearanceStates]) => ({
  page: 1,
  name,
  fieldType,
  rect,
  appearancePresent: true,
  appearanceStates,
  selectedState: "/Off",
  readOnly: false,
}));
formWidgets[7].selectedState = "/Yes";
const outputFormWidgets = structuredClone(formWidgets);
outputFormWidgets[5].selectedState = "/LLC";
const formWidgetChanges = formWidgets.map((widget, index) => ({
  name: widget.name,
  page: 1,
  fieldType: widget.fieldType,
  appearanceStates: widget.appearanceStates,
  expectedChange: [0, 1, 2, 5].includes(index),
  changedPixelsBBox: [0, 1, 2, 5].includes(index) ? [4, 4, 20, 18] : null,
  changedInteriorPixelsBBox: [0, 1, 2, 5].includes(index) ? [1, 1, 16, 14] : null,
}));
const formEvidence = {
  source: { sha256: "form-source-sha", pageCount: 1, pages: [{ page: 1, width: 612, height: 792, rotation: 0 }], decodedStreamErrors: [], startxrefCount: 1, eofCount: 1 },
  output: { sha256: "form-output-sha", pageCount: 1, pages: [{ page: 1, width: 612, height: 792, rotation: 0 }], decodedStreamErrors: [], startxrefCount: 2, eofCount: 2 },
  sourceForm: { acroFormPresent: true, needAppearances: false, fieldTreeRoots: 7, fields: formSourceFields, widgets: formWidgets },
  outputForm: { acroFormPresent: true, needAppearances: false, fieldTreeRoots: 7, fields: formOutputFields, widgets: outputFormWidgets },
  originalPrefixPreserved: true,
  visual: {
    renderer: "poppler-pdftoppm",
    sourcePageCount: 1,
    outputPageCount: 1,
    pages: [{ page: 1, sameDimensions: true, nonBlank: true, changedOnlyWithinAllowedMasks: true, changedOutsideAllowedMasksBBox: null }],
    widgetChanges: formWidgetChanges,
  },
};
const formAudit = {
  status: "succeeded",
  source: { sha256: "form-source-sha" },
  output: { sha256: "form-output-sha" },
  provider: { actual: "pypdf", version: "6.10.0", silentFallback: false },
  savePolicy: { strategy: "incremental" },
  preflight: { probeCompleted: true, planCompleted: true },
  operation: { type: "fill-form" },
};
const formCommands = [
  "python pypdf_edit.py inspect inputs/source.pdf --output tmp/inspect.json",
  "python pdf_provider.py check --provider pypdf --require",
  "python pdf_provider.py plan --task fill-form --provider pypdf --strategy incremental",
  "python pypdf_edit.py fill-form inputs/source.pdf outputs/form-filled.pdf --strategy incremental",
  "pdftoppm -png outputs/form-filled.pdf tmp/form-page",
  "python pdf_audit.py validate outputs/audit.json --artifact outputs/form-filled.pdf",
];
const formChecks = gradeAcroFormEvidence({ evidence: formEvidence, audit: formAudit, commands: formCommands, item: formItem });
assert.equal(formChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(formChecks, formItem.grade).rawScorePercent, 100);
const brokenRadioEvidence = structuredClone(formEvidence);
brokenRadioEvidence.outputForm.widgets[5].selectedState = "/Off";
const brokenRadioChecks = gradeAcroFormEvidence({ evidence: brokenRadioEvidence, audit: formAudit, commands: formCommands, item: formItem });
assert.equal(brokenRadioChecks.find((entry) => entry.id === "pdf-machine:radio-appearance-state")?.passed, false);
const formBypassChecks = gradeAcroFormEvidence({
  evidence: formEvidence,
  audit: formAudit,
  commands: [...formCommands, "writer = PdfWriter(reader, incremental=True); writer.update_page_form_field_values(None, values)"],
  item: formItem,
});
assert.equal(formBypassChecks.find((entry) => entry.id === "pdf-trace:no-ad-hoc-pypdf-writer")?.passed, false);

const attachmentItem = cases.find((item) => item.id === "pdf-attachment-quarantine-inventory");
const attachmentPayloads = [
  ["../escape.exe", "document", null, null, "application/vnd.microsoft.portable-executable", 17, "hash-escape", "escape.exe"],
  ["archive.zip", "document", null, null, "application/zip", 13, "hash-archive", "archive.zip"],
  ["report.txt", "document", null, null, "text/plain", 5, "hash-first", "report.txt"],
  ["report.txt", "document", null, null, "text/plain", 6, "hash-second", "report__2.txt"],
  ["unicode-测试.txt", "document", null, null, "text/plain", 7, "hash-unicode", "unicode-测试.txt"],
  ["report.txt", "page", 1, 0, "text/plain", 17, "hash-page", "report__3.txt"],
];
const expectedAttachments = attachmentPayloads.map(([displayName, scope, page, annotationIndex, mime, bytes, sha256], index) => ({
  scope,
  page,
  annotationIndex,
  internalKey: displayName,
  displayName,
  mime,
  bytes,
  sha256,
  ordinal: index + 1,
}));
const manifestAttachments = attachmentPayloads.map(([displayName, scope, page, annotationIndex, mime, bytes, sha256, savedName], index) => ({
  index: index + 1,
  scope,
  page,
  annotationIndex,
  internalKey: displayName,
  displayName,
  mime,
  bytes,
  sha256,
  savedName,
  savedPath: `quarantine/${savedName}`,
  nameSanitized: savedName !== displayName,
}));
const attachmentEvidence = {
  source: { sha256: "attachment-source-sha", pageCount: 1 },
  expectedAttachments,
  unsafeRawPaths: [{ displayName: "../escape.exe", resolved: "/workspace/outputs/escape.exe" }],
  manifest: {
    schema: "open-office-artifact-tool.pdf-attachments.v1",
    source: { sha256: "attachment-source-sha" },
    attachments: manifestAttachments,
    validation: { sourceUnchanged: true, attachmentsOpenedOrExecuted: false },
  },
  manifestFile: { sha256: "attachment-manifest-sha", bytes: 2000 },
  quarantine: {
    invalid: [],
    files: manifestAttachments.map((entry) => ({ path: entry.savedName, bytes: entry.bytes, sha256: entry.sha256, flat: true })),
  },
};
const attachmentAudit = {
  status: "succeeded",
  source: { sha256: "attachment-source-sha" },
  output: { sha256: "attachment-manifest-sha" },
  provider: { actual: "pypdf", version: "6.10.0", silentFallback: false },
  savePolicy: { strategy: "read-only" },
  preflight: { probeCompleted: true, planCompleted: true },
  operation: { type: "extract-attachments" },
  validation: {
    sourceUnchanged: true,
    allHashesVerified: true,
    allPathsContained: true,
    duplicateNamesSeparated: true,
    attachmentsOpenedOrExecuted: false,
  },
};
const attachmentCommands = [
  "python pypdf_edit.py inspect inputs/source.pdf --output tmp/inspect.json",
  "python pdf_provider.py check --provider pypdf --require",
  "python pdf_provider.py plan --task extract-attachments --provider pypdf --strategy read-only",
  "python pypdf_edit.py extract-attachments inputs/source.pdf outputs/quarantine --manifest outputs/attachments.json",
  "python pdf_audit.py validate outputs/audit.json --source inputs/source.pdf --artifact outputs/attachments.json --require-operation extract-attachments",
];
const attachmentChecks = gradeAttachmentQuarantineEvidence({ evidence: attachmentEvidence, audit: attachmentAudit, commands: attachmentCommands, item: attachmentItem });
assert.equal(attachmentChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(attachmentChecks, attachmentItem.grade).rawScorePercent, 100);
const quotedAttachmentAudit = structuredClone(attachmentAudit);
delete quotedAttachmentAudit.validation.allHashesVerified;
quotedAttachmentAudit.validation.allAttachmentHashesVerified = true;
const quotedAttachmentCommands = attachmentCommands.map((command) => command
  .replace("pypdf_edit.py ", 'pypdf_edit.py" ')
  .replace("pdf_provider.py ", 'pdf_provider.py" ')
  .replace("pdf_audit.py ", 'pdf_audit.py" '));
quotedAttachmentCommands.push('/bin/zsh -lc "shasum -a 256 outputs/quarantine/*"');
const quotedAttachmentChecks = gradeAttachmentQuarantineEvidence({
  evidence: attachmentEvidence,
  audit: quotedAttachmentAudit,
  commands: quotedAttachmentCommands,
  item: attachmentItem,
});
assert.equal(quotedAttachmentChecks.every((entry) => entry.passed), true);
const escapedAttachmentEvidence = structuredClone(attachmentEvidence);
escapedAttachmentEvidence.manifest.attachments[0].savedPath = "../escape.exe";
const escapedAttachmentChecks = gradeAttachmentQuarantineEvidence({ evidence: escapedAttachmentEvidence, audit: attachmentAudit, commands: attachmentCommands, item: attachmentItem });
assert.equal(escapedAttachmentChecks.find((entry) => entry.id === "pdf-security:path-traversal-contained")?.passed, false);
const manualAttachmentChecks = gradeAttachmentQuarantineEvidence({
  evidence: attachmentEvidence,
  audit: attachmentAudit,
  commands: [...attachmentCommands, "reader = PdfReader('inputs/source.pdf'); list(reader.attachment_list)"],
  item: attachmentItem,
});
assert.equal(manualAttachmentChecks.find((entry) => entry.id === "pdf-trace:no-ad-hoc-pypdf-extraction")?.passed, false);
const executedAttachmentChecks = gradeAttachmentQuarantineEvidence({
  evidence: attachmentEvidence,
  audit: attachmentAudit,
  commands: [...attachmentCommands, "unzip outputs/quarantine/archive.zip"],
  item: attachmentItem,
});
assert.equal(executedAttachmentChecks.find((entry) => entry.id === "pdf-trace:no-payload-open-or-execution")?.passed, false);
const interpretedAttachmentChecks = gradeAttachmentQuarantineEvidence({
  evidence: attachmentEvidence,
  audit: attachmentAudit,
  commands: [...attachmentCommands, '; "$PYTHON_BIN" outputs/quarantine/escape.exe'],
  item: attachmentItem,
});
assert.equal(interpretedAttachmentChecks.find((entry) => entry.id === "pdf-trace:no-payload-open-or-execution")?.passed, false);

const activeItem = cases.find((item) => item.id === "pdf-active-content-public-sanitize");
const activeTerms = activeItem.grade.machine.residueTerms;
const emptyTermCounts = Object.fromEntries(activeTerms.map((term) => [term, 0]));
const presentTermCounts = Object.fromEntries(activeTerms.map((term) => [term, 1]));
const activeEvidence = {
  source: {
    sha256: "active-source-sha",
    pageCount: 1,
    pages: [{ page: 1, width: 612, height: 792, rotation: 0 }],
    termCounts: emptyTermCounts,
    rawTermCounts: emptyTermCounts,
    decodedStreamTermCounts: emptyTermCounts,
    metadataTermCounts: emptyTermCounts,
    decodedStreamErrors: [],
  },
  output: {
    sha256: "active-output-sha",
    pageCount: 1,
    pages: [{ page: 1, width: 612, height: 792, rotation: 0 }],
    termCounts: emptyTermCounts,
    rawTermCounts: emptyTermCounts,
    decodedStreamTermCounts: emptyTermCounts,
    metadataTermCounts: emptyTermCounts,
    decodedStreamErrors: [],
    startxrefCount: 1,
    eofCount: 1,
  },
  sourceStructure: {
    structuralNameCounts: { "/AA": 1, "/EmbeddedFiles": 1, "/JS": 1, "/JavaScript": 1, "/Launch": 0, "/OpenAction": 1 },
    actionTypeCounts: { "/JavaScript": 2, "/Launch": 1, "/SubmitForm": 1 },
    attachments: [{ name: "internal.txt" }],
    attachmentTermCounts: emptyTermCounts,
    structureTermCounts: presentTermCounts,
    commentAnnotations: [{ subtype: "/Text" }],
    populatedWidgets: [{ name: "reviewer", values: { "/V": "Private Person" } }],
    personalMetadata: { "/Author": "Private Person" },
  },
  outputStructure: {
    structuralNameCounts: { "/AA": 0, "/EmbeddedFiles": 0, "/JS": 0, "/JavaScript": 0, "/Launch": 0, "/OpenAction": 0 },
    actionTypeCounts: { "/JavaScript": 0, "/Launch": 0, "/SubmitForm": 0 },
    attachments: [],
    attachmentTermCounts: emptyTermCounts,
    structureTermCounts: emptyTermCounts,
    commentAnnotations: [],
    populatedWidgets: [],
    personalMetadata: {},
  },
  originalPrefixPreserved: false,
  visual: {
    renderer: "poppler-pdftoppm",
    sourcePageCount: 1,
    outputPageCount: 1,
    allowedMasks: [{ page: 1, bboxPx: [136, 288, 512, 352] }],
    pages: [{ page: 1, sameDimensions: true, nonBlank: true, changedPixelsBBox: [154, 300, 220, 325], changedOutsideAllowedMasksBBox: null, changedOnlyWithinAllowedMasks: true }],
  },
};
const activeAudit = {
  status: "succeeded",
  source: { sha256: "active-source-sha" },
  output: { sha256: "active-output-sha" },
  provider: { actual: "pymupdf", version: "1.27.2.3", silentFallback: false },
  savePolicy: { strategy: "sanitize" },
  preflight: { probeCompleted: true, planCompleted: true },
  operation: [{ type: "scrub" }, { type: "active_content_cleanup" }],
  validation: { residue: { ok: true }, render: { pages: 1 } },
};
const activeCommands = [
  "python pymupdf_edit.py probe --accept-license agpl",
  "python pdf_provider.py plan --task sanitize --provider pymupdf --strategy sanitize",
  "python pymupdf_edit.py edit input.pdf output.pdf --strategy sanitize",
  "python residue_scan.py output.pdf --require-inert",
  "pdftoppm -png output.pdf page",
  "python pdf_audit.py validate audit.json",
];
const activeChecks = gradeActiveContentSanitizeEvidence({ evidence: activeEvidence, audit: activeAudit, commands: activeCommands, item: activeItem });
assert.equal(activeChecks.every((entry) => entry.passed), true);
assert.equal(summarizeCaseScore(activeChecks, activeItem.grade).rawScorePercent, 100);
const helpBeforeMutationChecks = gradeActiveContentSanitizeEvidence({
  evidence: activeEvidence,
  audit: activeAudit,
  commands: [
    "python pymupdf_edit.py probe --help && python pymupdf_edit.py inspect --help && python pymupdf_edit.py edit --help",
    "python pdf_provider.py plan --help",
    ...activeCommands,
  ],
  item: activeItem,
});
assert.equal(helpBeforeMutationChecks.find((entry) => entry.id === "pdf-trace:probe-plan-before-mutation")?.passed, true);
const editBeforePlanChecks = gradeActiveContentSanitizeEvidence({
  evidence: activeEvidence,
  audit: activeAudit,
  commands: [activeCommands[2], activeCommands[0], activeCommands[1], ...activeCommands.slice(3)],
  item: activeItem,
});
assert.equal(editBeforePlanChecks.find((entry) => entry.id === "pdf-trace:probe-plan-before-mutation")?.passed, false);
const residualActionEvidence = structuredClone(activeEvidence);
residualActionEvidence.outputStructure.actionTypeCounts["/SubmitForm"] = 1;
const residualActionChecks = gradeActiveContentSanitizeEvidence({ evidence: residualActionEvidence, audit: activeAudit, commands: activeCommands, item: activeItem });
assert.equal(residualActionChecks.find((entry) => entry.id === "pdf-security:active-names-absent")?.passed, false);
const visualDriftEvidence = structuredClone(activeEvidence);
visualDriftEvidence.visual.pages[0].changedOutsideAllowedMasksBBox = [20, 20, 40, 40];
visualDriftEvidence.visual.pages[0].changedOnlyWithinAllowedMasks = false;
const visualDriftChecks = gradeActiveContentSanitizeEvidence({ evidence: visualDriftEvidence, audit: activeAudit, commands: activeCommands, item: activeItem });
assert.equal(visualDriftChecks.find((entry) => entry.id === "pdf-visual:ordinary-content-stable")?.passed, false);
const activeBypassChecks = gradeActiveContentSanitizeEvidence({ evidence: activeEvidence, audit: activeAudit, commands: [...activeCommands, "doc.xref_set_key(1, 'AA', 'null')"], item: activeItem });
assert.equal(activeBypassChecks.find((entry) => entry.id === "pdf-trace:no-content-stream-bypass")?.passed, false);
const preMutationQaOnly = [
  activeCommands[0],
  activeCommands[1],
  "python residue_scan.py input.pdf --require-inert",
  "pdftoppm -png input.pdf source-page",
  "python pdf_audit.py validate preflight.json",
  activeCommands[2],
];
const preMutationQaChecks = gradeActiveContentSanitizeEvidence({ evidence: activeEvidence, audit: activeAudit, commands: preMutationQaOnly, item: activeItem });
for (const id of ["pdf-trace:post-mutation-residue-scan", "pdf-trace:post-mutation-poppler-render", "pdf-trace:audit-byte-validation"]) {
  assert.equal(preMutationQaChecks.find((entry) => entry.id === id)?.passed, false, `${id} must require evidence after mutation`);
}

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
