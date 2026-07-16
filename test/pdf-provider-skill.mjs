import assert from "node:assert/strict";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

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
  "tasks/forms_annotations.md",
  "tasks/sign_verify.md",
  "tasks/redact.md",
  "tasks/accessibility.md",
  "tasks/render_review.md",
  "tasks/provider_setup.md",
  "scripts/pdf_provider.py",
  "scripts/reportlab_create.py",
  "scripts/pdfplumber_extract.py",
  "scripts/pypdf_edit.py",
  "scripts/pymupdf_edit.py",
  "scripts/residue_scan.py",
  "scripts/pdf_audit.py",
  "examples/provider-workflows.md",
  "examples/reportlab-report-spec.json",
  "examples/pymupdf-edit-operations.json",
  "examples/pymupdf-redaction-operations.json",
];
for (const file of requiredFiles) assert.ok(manifest.includes(file), `PDF manifest is missing ${file}`);

const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
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
]) assert.match(skillText, pattern);

const pythonScripts = (await fs.readdir(scriptsRoot))
  .filter((file) => file.endsWith(".py"))
  .map((file) => path.join(scriptsRoot, file));
const compile = run(python, [
  "-c",
  "import pathlib,sys; [compile(pathlib.Path(p).read_text('utf-8'), p, 'exec') for p in sys.argv[1:]]",
  ...pythonScripts,
], { status: 0 });
assert.equal(compile.stderr, "");
const auditSchema = JSON.parse(await fs.readFile(path.join(skillRoot, "references", "pdf-audit-v1.schema.json"), "utf8"));
assert.equal(auditSchema.properties.schema.const, "open-office-artifact-tool.pdf-audit.v1");
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

    const formSource = path.join(tempRoot, "form-source.pdf");
    run(integrationPython, ["-c", [
      "from reportlab.pdfgen import canvas",
      "import sys",
      "c=canvas.Canvas(sys.argv[1])",
      "c.drawString(72,720,'Approval')",
      "c.acroForm.textfield(name='sender.city',tooltip='City',x=72,y=670,width=180,height=24,value='')",
      "c.save()",
    ].join(";"), formSource], { status: 0 });
    const formFilled = path.join(tempRoot, "form-filled.pdf");
    parseResult(run(integrationPython, [path.join(scriptsRoot, "pypdf_edit.py"), "fill-form", formSource, formFilled, "--strategy", "incremental", "--field", "sender.city=Shanghai"], { status: 0 }));

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
