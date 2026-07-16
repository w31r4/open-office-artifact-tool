import assert from "node:assert/strict";
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
    env: { ...process.env, ...options.env },
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

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-pdf-provider-skill-"));
try {
  const dummyInput = path.join(tempRoot, "input.pdf");
  const dummyOutput = path.join(tempRoot, "output.pdf");
  await fs.writeFile(dummyInput, "%PDF-1.4\n%%EOF\n", "ascii");

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
