import assert from "node:assert/strict";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { PdfArtifact, PdfFile } from "../src/index.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const skillRoot = path.join(repoRoot, "skills", "pdf", "skills", "pdf");
const scriptsRoot = path.join(skillRoot, "scripts");
const python = "python3";

function run(executable, args, options = {}) {
  const result = spawnSync(executable, args, {
    cwd: options.cwd || repoRoot,
    encoding: "utf8",
    env: { ...process.env, PYTHONDONTWRITEBYTECODE: "1", ...options.env },
    maxBuffer: 20 * 1024 * 1024,
  });
  if (options.status !== undefined) {
    assert.equal(result.status, options.status, `${executable} ${args.join(" ")}\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
  }
  return result;
}

function parseResult(result, stream = "stdout") {
  const source = result[stream]?.trim();
  assert.ok(source, `expected JSON on ${stream}\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
  return JSON.parse(source);
}

async function walk(root) {
  const files = [];
  for (const entry of await fs.readdir(root, { withFileTypes: true })) {
    if (entry.name === "__pycache__") continue;
    const target = path.join(root, entry.name);
    if (entry.isDirectory()) files.push(...await walk(target));
    else if (entry.isFile()) files.push(target);
  }
  return files;
}

const manifest = (await fs.readFile(path.join(skillRoot, "manifest.txt"), "utf8"))
  .split(/\r?\n/)
  .map((entry) => entry.trim())
  .filter(Boolean);
assert.equal(new Set(manifest).size, manifest.length, "PDF manifest must not contain duplicates");
for (const entry of manifest) {
  assert.equal(path.isAbsolute(entry), false, `PDF manifest entry must be relative: ${entry}`);
  assert.ok(!entry.split("/").includes(".."), `PDF manifest entry must stay inside the Skill: ${entry}`);
  await fs.access(path.join(skillRoot, entry));
}
const actualFiles = (await walk(skillRoot)).map((file) => path.relative(skillRoot, file)).sort();
assert.deepEqual([...manifest].sort(), actualFiles, "PDF manifest must cover the complete shipped Skill tree");

const requiredFiles = [
  "references/PROVIDER_MATRIX.md",
  "references/SAVE_POLICIES.md",
  "references/SECURITY_CHECKLIST.md",
  "references/PRODUCT_BOUNDARIES.md",
  "references/AUDIT_SCHEMA.md",
  "references/pdf-audit-v1.schema.json",
  "tasks/create.md",
  "tasks/read_review.md",
  "tasks/edit_existing.md",
  "tasks/transform.md",
  "tasks/forms_annotations.md",
  "tasks/sign_verify.md",
  "tasks/redact.md",
  "tasks/accessibility.md",
  "tasks/render_review.md",
  "tasks/provider_setup.md",
  "scripts/mupdf.mjs",
  "scripts/pdf_provider.py",
  "scripts/reportlab_create.py",
  "scripts/pdfplumber_extract.py",
  "scripts/pypdf_edit.py",
  "scripts/pymupdf_edit.py",
  "scripts/residue_scan.py",
  "scripts/pdf_audit.py",
  "scripts/python_runtime.py",
  "examples/provider-workflows.md",
  "examples/accessible-board-report.mjs",
  "examples/reportlab-report-spec.json",
  "examples/pymupdf-edit-operations.json",
  "examples/pymupdf-redaction-operations.json",
  "examples/merge-stamp-manifest.json",
  "scripts/poppler_compare.py",
];
for (const file of requiredFiles) assert.ok(manifest.includes(file), `PDF manifest is missing ${file}`);

const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
assert.match(skillText, /scripts\/mupdf\.mjs/);
assert.match(skillText, /MuPDF\.js/);
assert.match(skillText, /set_page_crop/);
assert.match(skillText, /rotate_page/);
assert.match(skillText, /delete_annotation/);
assert.match(skillText, /sourceSha256/);
assert.match(skillText, /not redaction/i);
for (const pattern of [
  /ReportLab/,
  /pdfplumber/,
  /pypdf/,
  /PyMuPDF/,
  /Poppler/,
  /pyHanko/,
  /veraPDF/,
  /rewrite/,
  /incremental/,
  /sanitize/,
  /silent fallback/i,
  /original bytes/i,
  /Word-style reflow/,
  /Dynamic XFA/,
  /accessible board report example/i,
]) assert.match(skillText, pattern);

const pythonScripts = (await fs.readdir(scriptsRoot))
  .filter((file) => file.endsWith(".py"))
  .map((file) => path.join(scriptsRoot, file));
const mupdfCliText = await fs.readFile(path.join(scriptsRoot, "mupdf.mjs"), "utf8");
assert.match(mupdfCliText, /open-office-artifact-tool\/pdf\/mupdf/);
assert.match(mupdfCliText, /PdfFile\.editPdf/);
assert.doesNotMatch(mupdfCliText, /postinstall/);
const compile = run(python, [
  "-c",
  "import pathlib,sys; [compile(pathlib.Path(p).read_text('utf-8'), p, 'exec') for p in sys.argv[1:]]",
  ...pythonScripts,
], { status: 0 });
assert.equal(compile.stderr, "");
const auditSchema = JSON.parse(await fs.readFile(path.join(skillRoot, "references", "pdf-audit-v1.schema.json"), "utf8"));
assert.equal(auditSchema.properties.schema.const, "open-office-artifact-tool.pdf-audit.v1");
assert.equal(auditSchema.properties.inputs.type, "array");
const fitUnit = run(python, ["-c", [
  "import importlib.util,sys",
  "spec=importlib.util.spec_from_file_location('pymupdf_edit',sys.argv[1])",
  "module=importlib.util.module_from_spec(spec)",
  "spec.loader.exec_module(module)",
  "quantized=module.text_width_fit(81.9169996380806,81.91697692871094)",
  "assert quantized['fits'] and 0 < quantized['overflow'] <= quantized['tolerance']",
  "overflow=module.text_width_fit(81.9185,81.91697692871094)",
  "assert not overflow['fits'] and overflow['overflow'] > overflow['tolerance']",
  "wide=module.text_width_fit(20000.0004,20000)",
  "assert wide['fits'] and wide['tolerance'] <= 0.0005",
  "wide_overflow=module.text_width_fit(20000.0006,20000)",
  "assert not wide_overflow['fits']",
].join(";"), path.join(scriptsRoot, "pymupdf_edit.py")], { status: 0 });
assert.equal(fitUnit.stderr, "");

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-pdf-provider-skill-"));
try {
  const mupdfCli = path.join(scriptsRoot, "mupdf.mjs");
  const mupdfInput = path.join(tempRoot, "mupdf-input.pdf");
  const mupdfRender = path.join(tempRoot, "mupdf-render.png");
  const mupdfOperations = path.join(tempRoot, "mupdf-operations.json");
  const mupdfOutput = path.join(tempRoot, "mupdf-output.pdf");
  const mupdfCropOperations = path.join(tempRoot, "mupdf-crop-operations.json");
  const mupdfCropOutput = path.join(tempRoot, "mupdf-crop-output.pdf");
  const mupdfRotationOperations = path.join(tempRoot, "mupdf-rotation-operations.json");
  const mupdfRotationOutput = path.join(tempRoot, "mupdf-rotation-output.pdf");
  const mupdfAnnotationDeleteOperations = path.join(tempRoot, "mupdf-annotation-delete-operations.json");
  const mupdfAnnotationDeleteOutput = path.join(tempRoot, "mupdf-annotation-delete-output.pdf");
  const mupdfLinkDeleteOperations = path.join(tempRoot, "mupdf-link-delete-operations.json");
  const mupdfLinkDeleteOutput = path.join(tempRoot, "mupdf-link-delete-output.pdf");
  const mupdfArtifact = PdfArtifact.create({ text: "MuPDF Skill CLI fixture" });
  mupdfArtifact.addLink({ text: "CLI link", url: "https://example.com/cli-link", bbox: [72, 120, 80, 16] });
  const mupdfFixture = await PdfFile.exportPdf(mupdfArtifact);
  await fs.writeFile(mupdfInput, mupdfFixture.bytes);
  const mupdfProbe = parseResult(run(process.execPath, [mupdfCli, "probe"], { status: 0 }));
  assert.deepEqual({ provider: mupdfProbe.provider, version: mupdfProbe.version, license: mupdfProbe.license }, { provider: "mupdf", version: "1.28.0", license: "AGPL-3.0-or-later" });
  const mupdfInspection = parseResult(run(process.execPath, [mupdfCli, "inspect", mupdfInput], { status: 0 }));
  assert.equal(mupdfInspection.summary.nativeProvider, "mupdf");
  assert.equal(mupdfInspection.summary.pages, 1);
  const mupdfLink = mupdfInspection.records.find((record) => record.kind === "mupdfLink");
  assert.match(mupdfLink.id, /^mupdf-link-1-[a-f0-9]{64}$/);
  await fs.writeFile(mupdfLinkDeleteOperations, JSON.stringify({
    savePolicy: "rewrite",
    operations: [{
      type: "delete_link",
      page: mupdfLink.page,
      linkId: mupdfLink.id,
      sourceSha256: mupdfInspection.summary.sourceSha256,
      expected: { url: mupdfLink.url, bbox: mupdfLink.bbox, external: mupdfLink.external },
    }],
  }), "utf8");
  const mupdfLinkDeleted = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfInput, mupdfLinkDeleteOperations, mupdfLinkDeleteOutput], { status: 0 }));
  assert.equal(mupdfLinkDeleted.savePolicy, "rewrite");
  assert.equal(mupdfLinkDeleted.operations[0].type, "delete_link");
  assert.equal((await PdfFile.inspectPdf(await fs.readFile(mupdfLinkDeleteOutput))).records.some((record) => record.kind === "mupdfLink"), false);
  const mupdfRendered = parseResult(run(process.execPath, [mupdfCli, "render", mupdfInput, mupdfRender, "--page", "1", "--dpi", "72"], { status: 0 }));
  assert.equal(mupdfRendered.provider, "mupdf");
  assert.deepEqual([...new Uint8Array(await fs.readFile(mupdfRender)).subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
  const mupdfNestedRender = path.join(tempRoot, "nested", "qa", "mupdf-render.png");
  run(process.execPath, [mupdfCli, "render", mupdfInput, mupdfNestedRender, "--page", "1", "--dpi", "72"], { status: 0 });
  assert.deepEqual([...new Uint8Array(await fs.readFile(mupdfNestedRender)).subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
  await fs.writeFile(mupdfOperations, JSON.stringify({ operations: [{ type: "add_text_annotation", page: 1, bbox: [40, 40, 24, 24], text: "CLI review" }], savePolicy: "incremental" }), "utf8");
  const mupdfEdited = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfInput, mupdfOperations, mupdfOutput], { status: 0 }));
  assert.equal(mupdfEdited.savePolicy, "incremental");
  const mupdfAnnotationInspection = await PdfFile.inspectPdf(await fs.readFile(mupdfOutput));
  assert.equal(mupdfAnnotationInspection.records.find((record) => record.kind === "mupdfPage").annotations, 1);
  const mupdfAnnotation = mupdfAnnotationInspection.records.find((record) => record.kind === "mupdfAnnotation");
  assert.match(mupdfAnnotation.id, /^mupdf-annotation-1-\d+$/);
  await fs.writeFile(mupdfAnnotationDeleteOperations, JSON.stringify({
    savePolicy: "rewrite",
    operations: [{
      type: "delete_annotation",
      page: mupdfAnnotation.page,
      annotationId: mupdfAnnotation.id,
      sourceSha256: mupdfAnnotationInspection.summary.sourceSha256,
      expected: { type: mupdfAnnotation.type, contents: mupdfAnnotation.contents, rect: mupdfAnnotation.rect },
    }],
  }), "utf8");
  const mupdfAnnotationDeleted = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfOutput, mupdfAnnotationDeleteOperations, mupdfAnnotationDeleteOutput], { status: 0 }));
  assert.equal(mupdfAnnotationDeleted.savePolicy, "rewrite");
  assert.equal(mupdfAnnotationDeleted.operations[0].type, "delete_annotation");
  assert.equal((await PdfFile.inspectPdf(await fs.readFile(mupdfAnnotationDeleteOutput))).records.find((record) => record.kind === "mupdfPage").annotations, 0);
  await fs.writeFile(mupdfCropOperations, JSON.stringify({ operations: [{ type: "set_page_crop", page: 1, bbox: [72, 72, 468, 648] }], savePolicy: "incremental" }), "utf8");
  const mupdfCropped = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfInput, mupdfCropOperations, mupdfCropOutput], { status: 0 }));
  assert.equal(mupdfCropped.savePolicy, "incremental");
  assert.equal(mupdfCropped.operations[0].contentRemoved, false);
  const cropPage = (await PdfFile.inspectPdf(await fs.readFile(mupdfCropOutput))).records.find((record) => record.kind === "mupdfPage");
  assert.deepEqual(cropPage.mediaBox, [0, 0, 612, 792]);
  assert.deepEqual(cropPage.cropBox, [72, 72, 468, 648]);
  const cropRender = await PdfFile.renderPdf(await fs.readFile(mupdfCropOutput), { page: 1, dpi: 72 });
  assert.deepEqual([cropRender.metadata.width, cropRender.metadata.height], [468, 648]);
  await fs.writeFile(mupdfRotationOperations, JSON.stringify({ operations: [{ type: "rotate_page", page: 1, rotation: 90 }], savePolicy: "incremental" }), "utf8");
  const mupdfRotated = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfInput, mupdfRotationOperations, mupdfRotationOutput], { status: 0 }));
  assert.equal(mupdfRotated.savePolicy, "incremental");
  assert.deepEqual(mupdfRotated.operations[0], {
    type: "rotate_page",
    page: 1,
    rotation: 90,
    previousRotation: 0,
    contentRemoved: false,
  });
  const rotationPage = (await PdfFile.inspectPdf(await fs.readFile(mupdfRotationOutput))).records.find((record) => record.kind === "mupdfPage");
  assert.equal(rotationPage.rotation, 90);
  const rotationRender = await PdfFile.renderPdf(await fs.readFile(mupdfRotationOutput), { page: 1, dpi: 72 });
  assert.deepEqual([rotationRender.metadata.width, rotationRender.metadata.height], [792, 612]);
  const mupdfInputAlias = path.join(tempRoot, "mupdf-input-alias.pdf");
  await fs.symlink(mupdfInput, mupdfInputAlias);
  const sourceOverwrite = run(process.execPath, [mupdfCli, "edit", mupdfInputAlias, mupdfOperations, mupdfInput], { status: 2 });
  assert.match(sourceOverwrite.stderr, /Refusing to overwrite the source PDF.*symlink alias/);

  const dummyInput = path.join(tempRoot, "input.pdf");
  const dummyOutput = path.join(tempRoot, "output.pdf");
  await fs.writeFile(dummyInput, "%PDF-1.4\n%%EOF\n", "ascii");
  const auditArtifact = path.join(tempRoot, "audit-artifact.pdf");
  await fs.writeFile(auditArtifact, "%PDF-1.4\naudit artifact\n%%EOF\n", "ascii");
  const evidence = async (target) => {
    const bytes = await fs.readFile(target);
    return { path: target, bytes: bytes.length, sha256: crypto.createHash("sha256").update(bytes).digest("hex") };
  };
  const auditPath = path.join(tempRoot, "audit.json");
  await fs.writeFile(auditPath, JSON.stringify({
    schema: "open-office-artifact-tool.pdf-audit.v1",
    status: "succeeded",
    source: await evidence(dummyInput),
    output: await evidence(auditArtifact),
    provider: { actual: "pymupdf", version: "test", silentFallback: false },
    savePolicy: { strategy: "sanitize" },
    preflight: { probeCompleted: true, planCompleted: true },
    operation: { type: "replace_text" },
    validation: {},
  }), "utf8");
  const auditValidation = parseResult(run(python, [path.join(scriptsRoot, "pdf_audit.py"), "validate", auditPath, "--source", dummyInput, "--artifact", auditArtifact, "--require-operation", "replace_text"], { status: 0 }));
  assert.equal(auditValidation.ok, true);
  await fs.appendFile(auditArtifact, "mutated");
  const staleAudit = run(python, [path.join(scriptsRoot, "pdf_audit.py"), "validate", auditPath, "--source", dummyInput, "--artifact", auditArtifact], { status: 2 });
  assert.match(staleAudit.stderr, /bytes\/hash do not match/);
  const refusalAuditPath = path.join(tempRoot, "refusal-audit.json");
  await fs.writeFile(refusalAuditPath, JSON.stringify({
    schema: "open-office-artifact-tool.pdf-audit.v1",
    status: "failed_closed",
    source: await evidence(dummyInput),
    output: null,
    provider: { actual: "pymupdf", version: "test", silentFallback: false },
    savePolicy: { strategy: "sanitize" },
    preflight: { probeCompleted: true, planCompleted: false },
    operation: { type: "replace_text" },
    validation: {},
    reason: "replacement does not fit",
  }), "utf8");
  const refusalValidation = parseResult(run(python, [path.join(scriptsRoot, "pdf_audit.py"), "validate", refusalAuditPath, "--source", dummyInput, "--require-operation", "replace_text"], { status: 0 }));
  assert.equal(refusalValidation.status, "failed_closed");
  const readOnlyManifest = path.join(tempRoot, "read-only-manifest.json");
  await fs.writeFile(readOnlyManifest, JSON.stringify({ attachments: [] }), "utf8");
  const readOnlyAuditPath = path.join(tempRoot, "read-only-audit.json");
  await fs.writeFile(readOnlyAuditPath, JSON.stringify({
    schema: "open-office-artifact-tool.pdf-audit.v1",
    status: "succeeded",
    source: await evidence(dummyInput),
    output: await evidence(readOnlyManifest),
    provider: { actual: "pypdf", version: "test", silentFallback: false },
    savePolicy: { strategy: "read-only" },
    preflight: { probeCompleted: true, planCompleted: true },
    operation: { type: "extract-attachments" },
    validation: { sourceUnchanged: true, attachmentsOpenedOrExecuted: false },
  }), "utf8");
  const readOnlyAuditValidation = parseResult(run(python, [
    path.join(scriptsRoot, "pdf_audit.py"), "validate", readOnlyAuditPath,
    "--source", dummyInput, "--artifact", readOnlyManifest, "--require-operation", "extract-attachments",
  ], { status: 0 }));
  assert.equal(readOnlyAuditValidation.savePolicy, "read-only");
  const secondInput = path.join(tempRoot, "second-input.pdf");
  const mergeManifest = path.join(tempRoot, "merge-manifest.json");
  const mergeArtifact = path.join(tempRoot, "merge-artifact.pdf");
  await fs.writeFile(secondInput, "%PDF-1.4\nsecond input\n%%EOF\n", "ascii");
  await fs.writeFile(mergeManifest, JSON.stringify({ schema: "open-office-artifact-tool.pdf-merge-stamp.v1" }), "utf8");
  await fs.writeFile(mergeArtifact, "%PDF-1.4\nmerged artifact\n%%EOF\n", "ascii");
  const multiAuditPath = path.join(tempRoot, "multi-audit.json");
  await fs.writeFile(multiAuditPath, JSON.stringify({
    schema: "open-office-artifact-tool.pdf-audit.v1",
    status: "succeeded",
    source: await evidence(mergeManifest),
    inputs: [await evidence(dummyInput), await evidence(secondInput)],
    output: await evidence(mergeArtifact),
    provider: { actual: "pypdf", version: "test", silentFallback: false },
    savePolicy: { strategy: "rewrite" },
    preflight: { probeCompleted: true, planCompleted: true },
    operation: { type: "merge-stamp" },
    validation: {},
  }), "utf8");
  const multiAuditValidation = parseResult(run(python, [
    path.join(scriptsRoot, "pdf_audit.py"), "validate", multiAuditPath,
    "--source", mergeManifest, "--input", secondInput, "--input", dummyInput,
    "--artifact", mergeArtifact, "--require-operation", "merge-stamp",
  ], { status: 0 }));
  assert.equal(multiAuditValidation.inputs, 2);
  const missingMultiInput = run(python, [
    path.join(scriptsRoot, "pdf_audit.py"), "validate", multiAuditPath,
    "--source", mergeManifest, "--input", dummyInput, "--artifact", mergeArtifact,
  ], { status: 2 });
  assert.match(missingMultiInput.stderr, /2 records but 1 --input/);

  const check = parseResult(run(python, [path.join(scriptsRoot, "pdf_provider.py"), "check", "--provider", "all"], { status: 0 }));
  assert.ok(check.providers.length >= 12);
  assert.ok(check.providers.every((provider) => typeof provider.integration === "string"));
  assert.equal(check.providers.find((provider) => provider.provider === "pymupdf")?.license, "agpl-or-commercial");

  const noLicense = run(python, [
    path.join(scriptsRoot, "pdf_provider.py"), "plan",
    "--task", "edit-content", "--provider", "pymupdf", "--strategy", "rewrite",
    "--input", dummyInput, "--output", dummyOutput,
  ], { status: 2 });
  assert.equal(parseResult(noLicense, "stderr").silentFallback, false);
  assert.match(noLicense.stderr, /accept-license/);

  const plan = parseResult(run(python, [
    path.join(scriptsRoot, "pdf_provider.py"), "plan",
    "--task", "edit-content", "--provider", "pymupdf", "--strategy", "rewrite",
    "--input", dummyInput, "--output", dummyOutput, "--accept-license", "agpl",
  ], { status: 0 }));
  assert.equal(plan.provider, "pymupdf");
  assert.equal(plan.strategy, "rewrite");
  assert.equal(plan.silentFallback, false);
  assert.equal(plan.input.sha256.length, 64);
  const attachmentPlan = parseResult(run(python, [
    path.join(scriptsRoot, "pdf_provider.py"), "plan",
    "--task", "extract-attachments", "--provider", "pypdf", "--strategy", "read-only",
    "--input", dummyInput,
  ], { status: 0 }));
  assert.equal(attachmentPlan.task, "extract-attachments");
  assert.equal(attachmentPlan.provider, "pypdf");
  assert.equal(attachmentPlan.strategy, "read-only");
  assert.equal(attachmentPlan.mutation, false);

  const samePath = run(python, [
    path.join(scriptsRoot, "pdf_provider.py"), "plan",
    "--task", "annotate", "--provider", "pypdf", "--strategy", "incremental",
    "--input", dummyInput, "--output", dummyInput,
  ], { status: 2 });
  assert.match(samePath.stderr, /must be different/);

  const invalidProvider = run(python, [
    path.join(scriptsRoot, "pdf_provider.py"), "plan",
    "--task", "edit-content", "--provider", "pypdf", "--strategy", "rewrite",
    "--input", dummyInput, "--output", dummyOutput,
  ], { status: 2 });
  assert.match(invalidProvider.stderr, /cannot perform task/);

  const unsafeSanitize = run(python, [
    path.join(scriptsRoot, "pdf_provider.py"), "plan",
    "--task", "redact", "--provider", "pymupdf", "--strategy", "sanitize",
    "--input", dummyInput, "--output", dummyOutput, "--accept-license", "agpl",
  ], { status: 2 });
  assert.match(unsafeSanitize.stderr, /invalidate-signatures/);

  const invalidConfiguredPython = run(python, [path.join(scriptsRoot, "pdf_provider.py"), "check", "--provider", "all"], {
    env: { OPEN_OFFICE_PDF_PROVIDER_PYTHON: path.join(tempRoot, "missing-python") },
    status: 2,
  });
  assert.match(invalidConfiguredPython.stderr, /OPEN_OFFICE_PDF_PROVIDER_PYTHON is not an executable file/);

  const directNoLicense = run(python, [path.join(scriptsRoot, "pymupdf_edit.py"), "probe"], { status: 2 });
  assert.equal(parseResult(directNoLicense, "stderr").silentFallback, false);
  assert.match(directNoLicense.stderr, /accept-license/);

  const integrationPython = process.env.OPEN_OFFICE_PDF_PROVIDER_PYTHON;
  if (integrationPython) {
    const sourcePath = path.join(tempRoot, "source.pdf");
    const specPath = path.join(tempRoot, "report.json");
    await fs.writeFile(specPath, JSON.stringify({
      title: "Provider contract test",
      subtitle: "Raw-byte editing and sanitization",
      sections: [{
        heading: "Decision",
        paragraphs: ["Customer Secret must be removed before delivery."],
        table: [["Gate", "Status"], ["Provider", "Ready"]],
      }],
    }), "utf8");
    const routedProbe = parseResult(run(python, [path.join(scriptsRoot, "pymupdf_edit.py"), "probe", "--accept-license", "agpl"], {
      env: { OPEN_OFFICE_PDF_PROVIDER_PYTHON: integrationPython },
      status: 0,
    }));
    assert.equal(routedProbe.available, true);
    assert.notEqual(routedProbe.providerVersion, "unavailable");
    const routedPlan = parseResult(run(python, [
      path.join(scriptsRoot, "pdf_provider.py"), "plan",
      "--task", "sanitize", "--provider", "pymupdf", "--strategy", "sanitize",
      "--input", dummyInput, "--output", dummyOutput, "--accept-license", "agpl",
      "--invalidate-signatures", "--require-provider",
    ], { env: { OPEN_OFFICE_PDF_PROVIDER_PYTHON: integrationPython }, status: 0 }));
    assert.equal(routedPlan.providerProbe.available, true);
    parseResult(run(integrationPython, [path.join(scriptsRoot, "reportlab_create.py"), "--spec", specPath, "--output", sourcePath], { status: 0 }));
    const sourceBytes = await fs.readFile(sourcePath);

    const extractionPath = path.join(tempRoot, "extraction.json");
    run(integrationPython, [path.join(scriptsRoot, "pdfplumber_extract.py"), sourcePath, "--output", extractionPath], { status: 0 });
    const extraction = JSON.parse(await fs.readFile(extractionPath, "utf8"));
    assert.match(extraction.pages[0].text, /Customer Secret/);

    const pypdfRewrite = path.join(tempRoot, "pypdf-rewrite.pdf");
    const pypdfIncremental = path.join(tempRoot, "pypdf-incremental.pdf");
    parseResult(run(integrationPython, [path.join(scriptsRoot, "pypdf_edit.py"), "add-note", sourcePath, pypdfRewrite, "--strategy", "rewrite", "--page", "1", "--rect", "500,700,524,724", "--text", "Rewrite note"], { status: 0 }));
    parseResult(run(integrationPython, [path.join(scriptsRoot, "pypdf_edit.py"), "add-note", sourcePath, pypdfIncremental, "--strategy", "incremental", "--page", "1", "--rect", "500,670,524,694", "--text", "Incremental note"], { status: 0 }));
    assert.equal((await fs.readFile(pypdfRewrite)).subarray(0, sourceBytes.length).equals(sourceBytes), false);
    assert.equal((await fs.readFile(pypdfIncremental)).subarray(0, sourceBytes.length).equals(sourceBytes), true);

    const createMergeSource = (target, label, pages, width, height) => run(integrationPython, ["-c", [
      "import pathlib, sys",
      "from reportlab.pdfgen import canvas",
      "from pypdf import PdfReader, PdfWriter",
      "out = pathlib.Path(sys.argv[1])",
      "label = sys.argv[2]",
      "slug = label.lower()",
      "count = int(sys.argv[3])",
      "width, height = float(sys.argv[4]), float(sys.argv[5])",
      "document = canvas.Canvas(str(out), pagesize=(width, height), invariant=1)",
      "for page in range(1, count + 1):",
      "    document.setFont('Helvetica-Bold', 14)",
      "    document.drawString(72, height - 48, f'{label} source page {page}')",
      "    document.bookmarkPage(f'{slug}-{page}')",
      "    document.addOutlineEntry(f'{label} {page}', f'{slug}-{page}', 0)",
      "    target_page = page % count + 1",
      "    document.drawString(72, height - 96, f'Internal link to {label} page {target_page}')",
      "    document.linkRect('', f'{slug}-{target_page}', (68, height - 104, 280, height - 80), relative=0, thickness=1)",
      "    document.showPage()",
      "document.save()",
      "reader = PdfReader(str(out), strict=True)",
      "writer = PdfWriter(clone_from=reader)",
      "for page in range(1, count + 1):",
      "    writer.add_named_destination(f'{slug}-named-{page}', page - 1)",
      "temporary = out.with_suffix('.named.pdf')",
      "writer.write(str(temporary))",
      "temporary.replace(out)",
    ].join("\n"), target, label, String(pages), String(width), String(height)], { status: 0 });
    const mergeCover = path.join(tempRoot, "merge-cover.pdf");
    const mergeReport = path.join(tempRoot, "merge-report.pdf");
    const mergeAppendix = path.join(tempRoot, "merge-appendix.pdf");
    createMergeSource(mergeCover, "Cover", 1, 612, 792);
    createMergeSource(mergeReport, "Report", 2, 792, 612);
    createMergeSource(mergeAppendix, "Appendix", 3, 595.2756, 841.8898);
    const mergeSourceHashes = await Promise.all([mergeCover, mergeReport, mergeAppendix].map(async (target) => (await evidence(target)).sha256));
    const mergeSpec = path.join(tempRoot, "merge-stamp.json");
    await fs.writeFile(mergeSpec, JSON.stringify({
      schema: "open-office-artifact-tool.pdf-merge-stamp.v1",
      sources: [
        { id: "cover", path: mergeCover },
        { id: "report", path: mergeReport },
        { id: "appendix", path: mergeAppendix },
      ],
      sequence: [
        { source: "cover", pages: "all" },
        { source: "appendix", pages: [3] },
        { source: "report", pages: "all" },
        { source: "appendix", pages: [1, 2] },
      ],
      watermarks: [{ source: "report", text: "CONFIDENTIAL", opacity: 0.2, angle: 45, fontSize: 48 }],
    }), "utf8");
    const mergeOutput = path.join(tempRoot, "merge-output.pdf");
    const mergeResult = parseResult(run(integrationPython, [
      path.join(scriptsRoot, "pypdf_edit.py"), "merge-stamp", mergeSpec, mergeOutput, "--strategy", "rewrite",
    ], { status: 0 }));
    assert.equal(mergeResult.schema, "open-office-artifact-tool.pdf-merge-stamp-result.v1");
    assert.deepEqual(mergeResult.operation.watermarks[0].outputPages, [3, 4]);
    assert.deepEqual(mergeResult.validation.pageMap.map((entry) => `${entry.source}:${entry.sourcePage}`), ["cover:1", "appendix:3", "report:1", "report:2", "appendix:1", "appendix:2"]);
    assert.equal(mergeResult.validation.navigation.outlines.length, 6);
    assert.equal(Object.keys(mergeResult.validation.navigation.namedDestinations).length, 6);
    assert.equal(mergeResult.validation.navigation.internalLinks.length, 6);
    assert.deepEqual(await Promise.all([mergeCover, mergeReport, mergeAppendix].map(async (target) => (await evidence(target)).sha256)), mergeSourceHashes);
    const mergeEvidence = parseResult(run(integrationPython, ["-c", [
      "import json, sys",
      "from pypdf import PdfReader",
      "reader = PdfReader(sys.argv[1], strict=True)",
      "print(json.dumps({",
      "  'pages': len(reader.pages),",
      "  'named': {name: reader.get_destination_page_number(destination) + 1 for name, destination in reader.named_destinations.items()},",
      "  'watermarks': [(page.extract_text() or '').count('CONFIDENTIAL') for page in reader.pages],",
      "  'sizes': [[float(page.mediabox.width), float(page.mediabox.height)] for page in reader.pages],",
      "}))",
    ].join("\n"), mergeOutput], { status: 0 }));
    assert.equal(mergeEvidence.pages, 6);
    assert.equal(Object.keys(mergeEvidence.named).length, 6);
    assert.deepEqual(mergeEvidence.watermarks, [0, 0, 1, 1, 0, 0]);
    assert.deepEqual(mergeEvidence.sizes, [[612, 792], [595.2756, 841.8898], [792, 612], [792, 612], [595.2756, 841.8898], [595.2756, 841.8898]]);
    const mergeVisualReport = path.join(tempRoot, "merge-visual-report.json");
    const mergeVisual = parseResult(run(integrationPython, [
      path.join(scriptsRoot, "poppler_compare.py"), "merge-stamp", mergeSpec, mergeOutput,
      "--report", mergeVisualReport, "--render-dir", path.join(tempRoot, "merge-rendered"),
    ], { status: 0 }));
    assert.equal(mergeVisual.schema, "open-office-artifact-tool.pdf-poppler-compare.v1");
    assert.equal(mergeVisual.status, "passed");
    assert.deepEqual(mergeVisual.pages.map((entry) => entry.pixelStable), [true, true, false, false, true, true]);
    assert.equal(mergeVisual.pages.every((entry) => entry.passed), true);
    const incompleteMergeSpec = path.join(tempRoot, "merge-stamp-incomplete.json");
    await fs.writeFile(incompleteMergeSpec, JSON.stringify({
      schema: "open-office-artifact-tool.pdf-merge-stamp.v1",
      sources: [{ id: "cover", path: mergeCover }, { id: "report", path: mergeReport }],
      sequence: [{ source: "cover", pages: "all" }, { source: "report", pages: [1] }],
      watermarks: [{ source: "report", text: "CONFIDENTIAL", opacity: 0.2 }],
    }), "utf8");
    const incompleteMergeOutput = path.join(tempRoot, "merge-incomplete.pdf");
    const incompleteMerge = run(integrationPython, [
      path.join(scriptsRoot, "pypdf_edit.py"), "merge-stamp", incompleteMergeSpec, incompleteMergeOutput, "--strategy", "rewrite",
    ], { status: 2 });
    assert.match(incompleteMerge.stderr, /select every source page exactly once/);
    await assert.rejects(fs.access(incompleteMergeOutput));

    const formSource = path.join(tempRoot, "form-source.pdf");
    run(integrationPython, ["-c", [
      "from reportlab.pdfgen import canvas",
      "import sys",
      "c=canvas.Canvas(sys.argv[1])",
      "c.drawString(72,720,'Approval')",
      "form=c.acroForm",
      "form.textfield(name='sender.city',tooltip='City',x=72,y=670,width=180,height=24,value='')",
      "form.radio(name='company_type',value='LLC',selected=False,x=72,y=620,buttonStyle='circle')",
      "form.radio(name='company_type',value='Corporation',selected=False,x=140,y=620,buttonStyle='circle')",
      "form.checkbox(name='terms_ack',checked=False,x=72,y=570,buttonStyle='check')",
      "c.save()",
    ].join(";"), formSource], { status: 0 });
    const formFilled = path.join(tempRoot, "form-filled.pdf");
    const formMutation = parseResult(run(integrationPython, [path.join(scriptsRoot, "pypdf_edit.py"), "fill-form", formSource, formFilled, "--strategy", "incremental", "--field", "sender.city=Shanghai", "--field", "company_type=LLC", "--field", "terms_ack=Yes"], { status: 0 }));
    assert.equal(formMutation.operation.fieldEvidence["sender.city"].fieldType, "/Tx");
    assert.equal(formMutation.operation.fieldEvidence.company_type.appearanceState, "/LLC");
    assert.equal(formMutation.operation.fieldEvidence.terms_ack.appearanceState, "/Yes");
    assert.equal(formMutation.originalPrefixPreserved, true);
    const formEvidence = parseResult(run(integrationPython, ["-c", [
      "import json, pypdf, sys",
      "reader = pypdf.PdfReader(sys.argv[1], strict=True)",
      "fields = reader.get_fields() or {}",
      "widgets = []",
      "for page in reader.pages:",
      "    for reference in page.get('/Annots', []) or []:",
      "        widget = reference.get_object()",
      "        if str(widget.get('/Subtype', '')) != '/Widget':",
      "            continue",
      "        parent = widget.get('/Parent')",
      "        parent = parent.get_object() if parent else widget",
      "        name = str(widget.get('/T') or parent.get('/T') or '')",
      "        widgets.append({'name': name, 'state': str(widget.get('/AS', '')), 'appearance': widget.get('/AP') is not None})",
      "result = {",
      "    'values': {name: str(field.get('/V', '')) for name, field in fields.items()},",
      "    'readOnly': {name: bool(int(field.get('/Ff', 0) or 0) & 1) for name, field in fields.items()},",
      "    'needAppearances': reader.trailer['/Root']['/AcroForm'].get('/NeedAppearances', False) == True,",
      "    'widgets': widgets,",
      "}",
      "print(json.dumps(result))",
    ].join("\n"), formFilled], { status: 0 }));
    assert.equal(formEvidence.values["sender.city"], "Shanghai");
    assert.equal(formEvidence.values.company_type, "/LLC");
    assert.equal(formEvidence.values.terms_ack, "/Yes");
    assert.equal(formEvidence.readOnly["sender.city"], false);
    assert.equal(formEvidence.readOnly.company_type, false);
    assert.equal(formEvidence.readOnly.terms_ack, false);
    assert.equal(formEvidence.needAppearances, false);
    assert.equal(formEvidence.widgets.every((widget) => widget.appearance), true);
    assert.deepEqual(formEvidence.widgets.filter((widget) => widget.name === "company_type").map((widget) => widget.state), ["/LLC", "/Off"]);
    assert.deepEqual(formEvidence.widgets.filter((widget) => widget.name === "terms_ack").map((widget) => widget.state), ["/Yes"]);

    const invalidRadioOutput = path.join(tempRoot, "form-invalid-radio.pdf");
    const invalidRadio = run(integrationPython, [path.join(scriptsRoot, "pypdf_edit.py"), "fill-form", formSource, invalidRadioOutput, "--strategy", "incremental", "--field", "company_type=Partnership"], { status: 2 });
    assert.match(invalidRadio.stderr, /no appearance state 'Partnership'/);
    assert.equal(await fs.stat(invalidRadioOutput).then(() => true, () => false), false);

    const attachmentFixtureScript = path.join(tempRoot, "attachment-fixture.py");
    await fs.writeFile(attachmentFixtureScript, [
      "import pathlib, sys",
      "from reportlab.lib.pagesizes import letter",
      "from reportlab.pdfgen import canvas",
      "from pypdf import PdfReader, PdfWriter",
      "from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject, NumberObject, RectangleObject, TextStringObject",
      "out = pathlib.Path(sys.argv[1])",
      "base = out.with_suffix('.base.pdf')",
      "c = canvas.Canvas(str(base), pagesize=letter, invariant=1)",
      "c.drawString(72, 680, 'Attachment quarantine fixture')",
      "c.save()",
      "writer = PdfWriter(clone_from=PdfReader(str(base)))",
      "for name, payload, mime in [('../escape.exe', b'MZ-safe-fixture', 'application#2Fvnd.microsoft.portable-executable'), ('report.txt', b'first', 'text#2Fplain'), ('report.txt', b'second', 'text#2Fplain')]:",
      "    embedded = writer.add_attachment(name, payload)",
      "    embedded.subtype = NameObject('/' + mime)",
      "stream = DecodedStreamObject()",
      "stream.set_data(b'page-level-review')",
      "stream[NameObject('/Type')] = NameObject('/EmbeddedFile')",
      "stream[NameObject('/Subtype')] = NameObject('/text#2Fplain')",
      "stream_ref = writer._add_object(stream)",
      "filespec = DictionaryObject({NameObject('/Type'): NameObject('/Filespec'), NameObject('/F'): TextStringObject('report.txt'), NameObject('/UF'): TextStringObject('report.txt'), NameObject('/EF'): DictionaryObject({NameObject('/F'): stream_ref, NameObject('/UF'): stream_ref})})",
      "annotation = DictionaryObject({NameObject('/Type'): NameObject('/Annot'), NameObject('/Subtype'): NameObject('/FileAttachment'), NameObject('/Rect'): RectangleObject((500, 680, 520, 700)), NameObject('/FS'): writer._add_object(filespec), NameObject('/Contents'): TextStringObject('Page attachment'), NameObject('/F'): NumberObject(4)})",
      "writer.add_annotation(0, annotation)",
      "with out.open('wb') as handle: writer.write(handle)",
      "base.unlink()",
    ].join("\n"), "utf8");
    const attachmentSource = path.join(tempRoot, "attachment-source.pdf");
    run(integrationPython, [attachmentFixtureScript, attachmentSource], { status: 0 });
    const attachmentSourceHash = crypto.createHash("sha256").update(await fs.readFile(attachmentSource)).digest("hex");
    const attachmentInspect = parseResult(run(integrationPython, [path.join(scriptsRoot, "pypdf_edit.py"), "inspect", attachmentSource], { status: 0 }));
    assert.equal(attachmentInspect.summary.attachments, 4);
    assert.equal(attachmentInspect.summary.documentAttachments, 3);
    assert.equal(attachmentInspect.summary.pageAttachments, 1);
    const quarantineDirectory = path.join(tempRoot, "quarantine");
    const attachmentManifest = path.join(tempRoot, "attachments.json");
    const attachmentResult = parseResult(run(integrationPython, [
      path.join(scriptsRoot, "pypdf_edit.py"), "extract-attachments", attachmentSource, quarantineDirectory,
      "--manifest", attachmentManifest,
    ], { status: 0 }));
    assert.equal(attachmentResult.schema, "open-office-artifact-tool.pdf-attachments.v1");
    assert.equal(attachmentResult.strategy, "read-only");
    assert.equal(attachmentResult.silentFallback, false);
    assert.equal(attachmentResult.source.sha256, attachmentSourceHash);
    assert.equal(attachmentResult.validation.sourceUnchanged, true);
    assert.equal(attachmentResult.validation.documentAttachments, 3);
    assert.equal(attachmentResult.validation.pageAttachments, 1);
    assert.equal(attachmentResult.validation.attachmentsOpenedOrExecuted, false);
    assert.equal(attachmentResult.attachments.length, 4);
    assert.equal(new Set(attachmentResult.attachments.map((entry) => entry.savedName.toLowerCase())).size, 4);
    assert.ok(attachmentResult.attachments.some((entry) => entry.displayName === "../escape.exe" && entry.savedName === "escape.exe" && entry.nameSanitized));
    assert.equal(attachmentResult.attachments.filter((entry) => entry.displayName === "report.txt").length, 3);
    assert.equal(attachmentResult.attachments.filter((entry) => entry.scope === "page").length, 1);
    for (const entry of attachmentResult.attachments) {
      const saved = path.join(tempRoot, entry.savedPath);
      const bytes = await fs.readFile(saved);
      assert.equal(bytes.length, entry.bytes);
      assert.equal(crypto.createHash("sha256").update(bytes).digest("hex"), entry.sha256);
      assert.equal(path.relative(quarantineDirectory, saved).startsWith(".."), false);
    }
    assert.equal(crypto.createHash("sha256").update(await fs.readFile(attachmentSource)).digest("hex"), attachmentSourceHash);
    await assert.rejects(fs.access(path.join(tempRoot, "escape.exe")));
    const limitedDirectory = path.join(tempRoot, "quarantine-limited");
    const limitedManifest = path.join(tempRoot, "attachments-limited.json");
    const limited = run(integrationPython, [
      path.join(scriptsRoot, "pypdf_edit.py"), "extract-attachments", attachmentSource, limitedDirectory,
      "--manifest", limitedManifest, "--max-attachments", "3",
    ], { status: 2 });
    assert.match(limited.stderr, /4 attachments; max-attachments is 3/);
    await assert.rejects(fs.access(limitedDirectory));
    await assert.rejects(fs.access(limitedManifest));

    const rewriteOps = path.join(tempRoot, "rewrite-ops.json");
    await fs.writeFile(rewriteOps, JSON.stringify([
      { type: "insert_textbox", page: 1, rect: [72, 650, 420, 690], text: "Reviewed through PyMuPDF", font_size: 11, color: [0.05, 0.35, 0.42] },
      { type: "add_text_annotation", page: 1, point: [510, 72], text: "Provider contract verified.", title: "QA" },
    ]), "utf8");
    const pymupdfRewrite = path.join(tempRoot, "pymupdf-rewrite.pdf");
    const rewrite = parseResult(run(integrationPython, [path.join(scriptsRoot, "pymupdf_edit.py"), "edit", sourcePath, pymupdfRewrite, "--strategy", "rewrite", "--operations", rewriteOps, "--accept-license", "agpl"], { status: 0 }));
    assert.equal(rewrite.silentFallback, false);
    assert.equal(rewrite.originalPrefixPreserved, false);

    const incrementalOps = path.join(tempRoot, "incremental-ops.json");
    await fs.writeFile(incrementalOps, JSON.stringify([{ type: "add_text_annotation", page: 1, point: [500, 100], text: "Incremental revision" }]), "utf8");
    const pymupdfIncremental = path.join(tempRoot, "pymupdf-incremental.pdf");
    const incremental = parseResult(run(integrationPython, [path.join(scriptsRoot, "pymupdf_edit.py"), "edit", sourcePath, pymupdfIncremental, "--strategy", "incremental", "--operations", incrementalOps, "--accept-license", "agpl"], { status: 0 }));
    assert.equal(incremental.originalPrefixPreserved, true);
    assert.equal((await fs.readFile(pymupdfIncremental)).subarray(0, sourceBytes.length).equals(sourceBytes), true);

    const redactOps = path.join(tempRoot, "redact-ops.json");
    await fs.writeFile(redactOps, JSON.stringify([
      { type: "replace_text", term: "Customer Secret", replacement: "Public Record", font_name: "helv", font_size: 9, fill: [1, 1, 1], color: [0, 0, 0] },
      { type: "scrub" },
    ]), "utf8");
    const sanitizedPath = path.join(tempRoot, "sanitized.pdf");
    const sanitized = parseResult(run(integrationPython, [path.join(scriptsRoot, "pymupdf_edit.py"), "edit", sourcePath, sanitizedPath, "--strategy", "sanitize", "--operations", redactOps, "--sensitive-term", "Customer Secret", "--accept-license", "agpl", "--invalidate-signatures"], { status: 0 }));
    assert.equal(sanitized.originalPrefixPreserved, false);
    assert.equal(sanitized.residueScan.ok, true);
    assert.equal(sanitized.residueScan.revisions.eofMarkers, 1);
    assert.equal(sanitized.residueScan.revisions.prevPointers, 0);

    const sanitizedExtractionPath = path.join(tempRoot, "sanitized-extraction.json");
    run(integrationPython, [path.join(scriptsRoot, "pdfplumber_extract.py"), sanitizedPath, "--output", sanitizedExtractionPath], { status: 0 });
    const sanitizedExtraction = JSON.parse(await fs.readFile(sanitizedExtractionPath, "utf8"));
    assert.doesNotMatch(sanitizedExtraction.pages[0].text, /Customer Secret/);
    assert.match(sanitizedExtraction.pages[0].text, /Public Record/);

    const activeFixtureScript = path.join(tempRoot, "active-fixture.py");
    await fs.writeFile(activeFixtureScript, [
      "import pathlib, sys",
      "from reportlab.lib.pagesizes import letter",
      "from reportlab.pdfgen import canvas",
      "from pypdf import PdfReader, PdfWriter",
      "from pypdf.annotations import Text",
      "from pypdf.generic import DictionaryObject, NameObject, TextStringObject",
      "out = pathlib.Path(sys.argv[1])",
      "base = out.with_suffix('.base.pdf')",
      "c = canvas.Canvas(str(base), pagesize=letter, invariant=1)",
      "c.drawString(72, 680, 'Stable public content')",
      "c.acroForm.textfield(name='reviewer', tooltip='Reviewer', x=72, y=620, width=180, height=24, value='Private Person')",
      "hidden = c.beginText(72, 580)",
      "hidden.setTextRenderMode(3)",
      "hidden.textLine('HIDDEN-ACTIVE-CANARY')",
      "c.drawText(hidden)",
      "c.save()",
      "writer = PdfWriter(clone_from=PdfReader(str(base)))",
      "writer.add_metadata({'/Author': 'Private Person', '/Subject': 'Internal only'})",
      "writer.add_attachment('internal.txt', b'ATTACHMENT-ACTIVE-CANARY')",
      "writer.add_js(\"app.alert('JS-ACTIVE-CANARY');\")",
      "writer._root_object[NameObject('/OpenAction')] = DictionaryObject({NameObject('/S'): NameObject('/JavaScript'), NameObject('/JS'): TextStringObject(\"app.alert('OPEN-ACTIVE-CANARY');\")})",
      "writer._root_object[NameObject('/AA')] = DictionaryObject({NameObject('/WC'): DictionaryObject({NameObject('/S'): NameObject('/Launch'), NameObject('/F'): TextStringObject('LAUNCH-ACTIVE-CANARY.exe')}), NameObject('/WP'): DictionaryObject({NameObject('/S'): NameObject('/SubmitForm'), NameObject('/F'): TextStringObject('https://invalid.example/SUBMIT-ACTIVE-CANARY')})})",
      "writer.add_annotation(0, Text(rect=(500, 680, 520, 700), text='COMMENT-ACTIVE-CANARY'))",
      "with out.open('wb') as handle: writer.write(handle)",
      "base.unlink()",
    ].join("\n"), "utf8");
    const activeSource = path.join(tempRoot, "active-source.pdf");
    run(integrationPython, [activeFixtureScript, activeSource], { status: 0 });
    const activeSourceScan = parseResult(run(integrationPython, [path.join(scriptsRoot, "residue_scan.py"), activeSource, "--require-inert", "--require-single-revision"], { status: 2 }));
    const activeSourceCategories = new Set(activeSourceScan.activeContent.violations.map((entry) => entry.category));
    for (const category of ["attachments", "annotations", "form-values", "personal-metadata", "hidden-text", "active-structure"]) assert.ok(activeSourceCategories.has(category), `missing source active-content category ${category}`);
    const scrubOnlyOps = path.join(tempRoot, "scrub-only.json");
    await fs.writeFile(scrubOnlyOps, JSON.stringify([{ type: "scrub" }]), "utf8");
    const activeOutput = path.join(tempRoot, "active-public.pdf");
    const activeResult = parseResult(run(integrationPython, [path.join(scriptsRoot, "pymupdf_edit.py"), "edit", activeSource, activeOutput, "--strategy", "sanitize", "--operations", scrubOnlyOps, "--accept-license", "agpl", "--invalidate-signatures"], { status: 0 }));
    const activeCleanup = activeResult.operations.find((operation) => operation.type === "active_content_cleanup");
    assert.equal(activeResult.residueScan.ok, true);
    assert.equal(activeResult.originalPrefixPreserved, false);
    assert.equal(activeCleanup.annotationsRemoved, 1);
    assert.equal(activeCleanup.widgetsCleared, 1);
    assert.equal(activeCleanup.hiddenTextSpansRemoved.length, 1);
    assert.ok(activeCleanup.actionKeysRemoved.length >= 2);
    assert.ok(activeCleanup.nameTreeKeysRemoved.length >= 2);
    assert.ok(activeCleanup.nullActiveContentKeysRemoved.some((entry) => entry.keys.includes("OpenAction")));
    assert.ok(activeCleanup.nullActiveContentKeysRemoved.some((entry) => entry.keys.includes("AA")));
    assert.ok(activeCleanup.nullActiveContentKeysRemoved.some((entry) => entry.keys.includes("JavaScript")));
    assert.ok(activeCleanup.nullActiveContentKeysRemoved.some((entry) => entry.keys.includes("EmbeddedFiles")));
    const activeOutputScan = parseResult(run(integrationPython, [path.join(scriptsRoot, "residue_scan.py"), activeOutput, "--require-inert", "--require-single-revision"], { status: 0 }));
    assert.equal(activeOutputScan.ok, true);
    assert.equal(activeOutputScan.summary.attachments, 0);
    assert.equal(activeOutputScan.summary.commentAnnotations, 0);
    assert.equal(activeOutputScan.summary.populatedFormValues, 0);
    assert.equal(activeOutputScan.summary.hiddenTextSpans, 0);
    assert.deepEqual(activeOutputScan.activeContent.violations, []);

    const overlapSource = path.join(tempRoot, "hidden-overlap.pdf");
    run(integrationPython, ["-c", [
      "from reportlab.lib.pagesizes import letter",
      "from reportlab.pdfgen import canvas",
      "import sys",
      "c=canvas.Canvas(sys.argv[1],pagesize=letter,invariant=1)",
      "c.drawString(72,680,'VISIBLE')",
      "t=c.beginText(72,680)",
      "t.setTextRenderMode(3)",
      "t.textLine('HIDDEN')",
      "c.drawText(t)",
      "c.save()",
    ].join(";"), overlapSource], { status: 0 });
    const overlapOutput = path.join(tempRoot, "hidden-overlap-output.pdf");
    const overlap = run(integrationPython, [path.join(scriptsRoot, "pymupdf_edit.py"), "edit", overlapSource, overlapOutput, "--strategy", "sanitize", "--operations", scrubOnlyOps, "--accept-license", "agpl", "--invalidate-signatures"], { status: 2 });
    assert.match(overlap.stderr, /hidden text .* overlaps visible text/);
    await assert.rejects(fs.access(overlapOutput));

    const sourceResidue = run(integrationPython, [path.join(scriptsRoot, "residue_scan.py"), sourcePath, "--term", "Customer Secret", "--require-single-revision"], { status: 2 });
    const sourceResidueReport = parseResult(sourceResidue, "stdout");
    assert.equal(sourceResidueReport.ok, false);
    assert.ok(sourceResidueReport.summary.matches > 0);

    const overflowOps = path.join(tempRoot, "overflow-ops.json");
    await fs.writeFile(overflowOps, JSON.stringify([
      { type: "replace_text", term: "Customer Secret", replacement: "This replacement deliberately cannot fit inside the original short text box and must not trigger general reflow", font_name: "helv", font_size: 12 },
      { type: "scrub" },
    ]), "utf8");
    const overflowOutput = path.join(tempRoot, "overflow-output.pdf");
    const overflow = run(integrationPython, [path.join(scriptsRoot, "pymupdf_edit.py"), "edit", sourcePath, overflowOutput, "--strategy", "sanitize", "--operations", overflowOps, "--sensitive-term", "Customer Secret", "--accept-license", "agpl", "--invalidate-signatures"], { status: 2 });
    assert.match(overflow.stderr, /general PDF reflow/);
    await assert.rejects(fs.access(overflowOutput));

    const toleranceSource = path.join(tempRoot, "tolerance-source.pdf");
    run(integrationPython, ["-c", [
      "from reportlab.lib.pagesizes import letter",
      "from reportlab.pdfgen import canvas",
      "import sys",
      "c=canvas.Canvas(sys.argv[1],pagesize=letter,invariant=1)",
      "c.setFont('Helvetica',11)",
      "c.drawString(84,628,'Contract ID: ACME-2025-041')",
      "c.save()",
    ].join(";"), toleranceSource], { status: 0 });
    const toleranceOps = path.join(tempRoot, "tolerance-ops.json");
    await fs.writeFile(toleranceOps, JSON.stringify([
      { type: "replace_text", page: 1, term: "ACME-2025-041", replacement: "ACME-2026-041" },
    ]), "utf8");
    const toleranceOutput = path.join(tempRoot, "tolerance-output.pdf");
    const toleranceResult = parseResult(run(integrationPython, [path.join(scriptsRoot, "pymupdf_edit.py"), "edit", toleranceSource, toleranceOutput, "--strategy", "sanitize", "--operations", toleranceOps, "--sensitive-term", "ACME-2025-041", "--accept-license", "agpl", "--invalidate-signatures"], { status: 0 }));
    const fitEvidence = toleranceResult.operations.find((operation) => operation.type === "replacement_overlays")?.fitChecks?.[0];
    assert.equal(fitEvidence?.fits, true);
    assert.ok(fitEvidence.overflow > 0 && fitEvidence.overflow <= fitEvidence.tolerance);
    assert.equal(fitEvidence.sourceFont, "Helvetica");
    assert.equal(fitEvidence.sourceFontSize, 11);
    assert.equal(fitEvidence.outputFont, "helv");
    assert.equal(fitEvidence.outputFontSize, 11);
    const toleranceExtraction = path.join(tempRoot, "tolerance-extraction.json");
    run(integrationPython, [path.join(scriptsRoot, "pdfplumber_extract.py"), toleranceOutput, "--output", toleranceExtraction], { status: 0 });
    const toleranceText = JSON.parse(await fs.readFile(toleranceExtraction, "utf8")).pages[0].text;
    assert.doesNotMatch(toleranceText, /ACME-2025-041/);
    assert.match(toleranceText, /ACME-2026-041/);

    const beyondToleranceOps = path.join(tempRoot, "beyond-tolerance-ops.json");
    await fs.writeFile(beyondToleranceOps, JSON.stringify([
      { type: "replace_text", page: 1, term: "ACME-2025-041", replacement: "ACME-2026-041", font_name: "helv", font_size: 11.0002 },
    ]), "utf8");
    const beyondToleranceOutput = path.join(tempRoot, "beyond-tolerance-output.pdf");
    const beyondTolerance = run(integrationPython, [path.join(scriptsRoot, "pymupdf_edit.py"), "edit", toleranceSource, beyondToleranceOutput, "--strategy", "sanitize", "--operations", beyondToleranceOps, "--sensitive-term", "ACME-2025-041", "--accept-license", "agpl", "--invalidate-signatures"], { status: 2 });
    assert.match(beyondTolerance.stderr, /numeric tolerance .*general PDF reflow/);
    await assert.rejects(fs.access(beyondToleranceOutput));

    const rotatedSource = path.join(tempRoot, "rotated-source.pdf");
    run(integrationPython, ["-c", [
      "from reportlab.lib.pagesizes import letter",
      "from reportlab.pdfgen import canvas",
      "import sys",
      "c=canvas.Canvas(sys.argv[1],pagesize=letter,invariant=1)",
      "c.saveState()",
      "c.translate(160,420)",
      "c.rotate(90)",
      "c.setFont('Helvetica',11)",
      "c.drawString(0,0,'ROTATED-ID')",
      "c.restoreState()",
      "c.save()",
    ].join(";"), rotatedSource], { status: 0 });
    const rotatedOps = path.join(tempRoot, "rotated-ops.json");
    await fs.writeFile(rotatedOps, JSON.stringify([
      { type: "replace_text", page: 1, term: "ROTATED-ID", replacement: "UPDATED-ID" },
    ]), "utf8");
    const rotatedOutput = path.join(tempRoot, "rotated-output.pdf");
    const rotated = run(integrationPython, [path.join(scriptsRoot, "pymupdf_edit.py"), "edit", rotatedSource, rotatedOutput, "--strategy", "sanitize", "--operations", rotatedOps, "--sensitive-term", "ROTATED-ID", "--accept-license", "agpl", "--invalidate-signatures"], { status: 2 });
    assert.match(rotated.stderr, /only horizontal source text|rotated or skewed text/);
    await assert.rejects(fs.access(rotatedOutput));

    const imageOps = path.join(tempRoot, "image-ops.json");
    await fs.writeFile(imageOps, JSON.stringify([{ type: "insert_image", page: 1, rect: [470, 630, 540, 700], path: path.join(repoRoot, "skills", "pdf", "assets", "icon.png") }]), "utf8");
    const imagePath = path.join(tempRoot, "image-source.pdf");
    parseResult(run(integrationPython, [path.join(scriptsRoot, "pymupdf_edit.py"), "edit", sourcePath, imagePath, "--strategy", "rewrite", "--operations", imageOps, "--accept-license", "agpl"], { status: 0 }));
    const ocrScan = run(integrationPython, [path.join(scriptsRoot, "residue_scan.py"), imagePath, "--term", "__NO_SUCH_SECRET__", "--require-ocr", "--require-single-revision"]);
    assert.ok([0, 2].includes(ocrScan.status));
    const ocrReport = parseResult(ocrScan, "stdout");
    if (ocrScan.status === 2) {
      assert.equal(ocrReport.ok, false);
      assert.ok(ocrReport.incomplete.some((item) => item.category === "image-ocr"));
    } else {
      assert.equal(ocrReport.ok, true);
    }

    const poppler = run("pdftoppm", ["-v"]);
    if (poppler.status === 0) {
      const renderedPrefix = path.join(tempRoot, "sanitized-render");
      run("pdftoppm", ["-singlefile", "-png", "-r", "144", sanitizedPath, renderedPrefix], { status: 0 });
      assert.ok((await fs.stat(`${renderedPrefix}.png`)).size > 1_000);
    }
  }

  console.log(`pdf provider skill smoke ok${process.env.OPEN_OFFICE_PDF_PROVIDER_PYTHON ? " (real providers)" : " (contract-only; set OPEN_OFFICE_PDF_PROVIDER_PYTHON for real providers)"}`);
} finally {
  await fs.rm(tempRoot, { recursive: true, force: true });
}
