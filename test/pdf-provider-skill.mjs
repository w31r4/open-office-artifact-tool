import assert from "node:assert/strict";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import sharp from "sharp";

import { PdfArtifact, PdfFile } from "../src/index.mjs";
import { plainPdfBytes } from "./fixtures/plain-pdf.mjs";

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

async function assertOcrRedactionPixelScope(source, output, operation, tempRoot, label) {
  const sourceRender = path.join(tempRoot, `${label}-source-render`);
  const outputRender = path.join(tempRoot, `${label}-output-render`);
  run("pdftoppm", ["-singlefile", "-png", "-r", "144", source, sourceRender], { status: 0 });
  run("pdftoppm", ["-singlefile", "-png", "-r", "144", output, outputRender], { status: 0 });
  const sourcePixels = await sharp(`${sourceRender}.png`).removeAlpha().raw().toBuffer({ resolveWithObject: true });
  const outputPixels = await sharp(`${outputRender}.png`).removeAlpha().raw().toBuffer({ resolveWithObject: true });
  assert.deepEqual(outputPixels.info, sourcePixels.info);

  const [imageRect] = operation.displayImageRects;
  const [redactionRect] = operation.displayRects;
  const pageRect = operation.displayPageRect;
  assert.ok(imageRect && redactionRect && pageRect, `${label} must report display-space QA geometry`);
  const xScale = sourcePixels.info.width / (pageRect[2] - pageRect[0]);
  const yScale = sourcePixels.info.height / (pageRect[3] - pageRect[1]);
  const scaleRect = (bounds) => [
    (bounds[0] - pageRect[0]) * xScale,
    (bounds[1] - pageRect[1]) * yScale,
    (bounds[2] - pageRect[0]) * xScale,
    (bounds[3] - pageRect[1]) * yScale,
  ];
  const imageBounds = scaleRect(imageRect);
  const redactBounds = scaleRect(redactionRect);
  let changedPixels = 0;
  let changedOutsideImage = 0;
  let changedInsideRedaction = 0;
  let darkInsideRedaction = 0;
  for (let y = 0; y < sourcePixels.info.height; y += 1) {
    for (let x = 0; x < sourcePixels.info.width; x += 1) {
      const offset = (y * sourcePixels.info.width + x) * sourcePixels.info.channels;
      const changed = Array.from(
        { length: sourcePixels.info.channels },
        (_, channel) => Math.abs(sourcePixels.data[offset + channel] - outputPixels.data[offset + channel]),
      ).some((delta) => delta > 5);
      const insideImage = x >= imageBounds[0] - 2 && x <= imageBounds[2] + 2 && y >= imageBounds[1] - 2 && y <= imageBounds[3] + 2;
      const insideRedaction = x >= redactBounds[0] && x <= redactBounds[2] && y >= redactBounds[1] && y <= redactBounds[3];
      if (changed) {
        changedPixels += 1;
        if (!insideImage) changedOutsideImage += 1;
        if (insideRedaction) changedInsideRedaction += 1;
      }
      if (
        insideRedaction
        && outputPixels.data[offset] < 32
        && outputPixels.data[offset + 1] < 32
        && outputPixels.data[offset + 2] < 32
      ) darkInsideRedaction += 1;
    }
  }
  assert.ok(changedPixels > 500, `${label} expected visible OCR redaction, changed ${changedPixels} pixels`);
  assert.equal(changedOutsideImage, 0, `${label} OCR redaction must not alter content outside the source image placement`);
  assert.ok(changedInsideRedaction > 500, `${label} expected changed pixels inside OCR match, found ${changedInsideRedaction}`);
  assert.ok(darkInsideRedaction > 500, `${label} expected opaque redaction fill, found ${darkInsideRedaction} dark pixels`);
}

async function assertNativePlacementPixelScope(source, output, page, operations, tempRoot, label) {
  const sourceRender = path.join(tempRoot, `${label}-source-render`);
  const outputRender = path.join(tempRoot, `${label}-output-render`);
  run("pdftoppm", ["-singlefile", "-png", "-r", "144", source, sourceRender], { status: 0 });
  run("pdftoppm", ["-singlefile", "-png", "-r", "144", output, outputRender], { status: 0 });
  const sourcePixels = await sharp(`${sourceRender}.png`).removeAlpha().raw().toBuffer({ resolveWithObject: true });
  const outputPixels = await sharp(`${outputRender}.png`).removeAlpha().raw().toBuffer({ resolveWithObject: true });
  assert.deepEqual(outputPixels.info, sourcePixels.info);

  const annotation = operations.find((operation) => operation.type === "add_text_annotation");
  const highlight = operations.find((operation) => operation.type === "add_text_highlight");
  const link = operations.find((operation) => operation.type === "add_link");
  assert.ok(
    annotation?.added?.rect
      && annotation.added.appearanceBbox
      && highlight?.added?.quadPoints?.length
      && highlight.added.appearanceBbox
      && link?.added?.bbox,
    `${label} must report provider placement geometry`,
  );
  const pageRect = [page.bbox[0], page.bbox[1], page.bbox[0] + page.bbox[2], page.bbox[1] + page.bbox[3]];
  const noteRect = [
    annotation.added.appearanceBbox[0],
    annotation.added.appearanceBbox[1],
    annotation.added.appearanceBbox[0] + annotation.added.appearanceBbox[2],
    annotation.added.appearanceBbox[1] + annotation.added.appearanceBbox[3],
  ];
  const highlightRect = [
    highlight.added.appearanceBbox[0],
    highlight.added.appearanceBbox[1],
    highlight.added.appearanceBbox[0] + highlight.added.appearanceBbox[2],
    highlight.added.appearanceBbox[1] + highlight.added.appearanceBbox[3],
  ];
  const linkRect = [
    link.added.bbox[0],
    link.added.bbox[1],
    link.added.bbox[0] + link.added.bbox[2],
    link.added.bbox[1] + link.added.bbox[3],
  ];
  const xScale = sourcePixels.info.width / page.bbox[2];
  const yScale = sourcePixels.info.height / page.bbox[3];
  const scaleRect = (bounds) => [
    (bounds[0] - pageRect[0]) * xScale,
    (bounds[1] - pageRect[1]) * yScale,
    (bounds[2] - pageRect[0]) * xScale,
    (bounds[3] - pageRect[1]) * yScale,
  ];
  const allowed = [scaleRect(noteRect), scaleRect(highlightRect), scaleRect(linkRect)];
  let changedPixels = 0;
  let changedOutsidePlacement = 0;
  const changedBounds = [Infinity, Infinity, -Infinity, -Infinity];
  const outsideBounds = [Infinity, Infinity, -Infinity, -Infinity];
  for (let y = 0; y < sourcePixels.info.height; y += 1) {
    for (let x = 0; x < sourcePixels.info.width; x += 1) {
      const offset = (y * sourcePixels.info.width + x) * sourcePixels.info.channels;
      const changed = Array.from(
        { length: sourcePixels.info.channels },
        (_, channel) => Math.abs(sourcePixels.data[offset + channel] - outputPixels.data[offset + channel]),
      ).some((delta) => delta > 5);
      if (!changed) continue;
      changedPixels += 1;
      changedBounds[0] = Math.min(changedBounds[0], x);
      changedBounds[1] = Math.min(changedBounds[1], y);
      changedBounds[2] = Math.max(changedBounds[2], x);
      changedBounds[3] = Math.max(changedBounds[3], y);
      if (!allowed.some((bounds) => x >= bounds[0] - 4 && x <= bounds[2] + 4 && y >= bounds[1] - 4 && y <= bounds[3] + 4)) {
        changedOutsidePlacement += 1;
        outsideBounds[0] = Math.min(outsideBounds[0], x);
        outsideBounds[1] = Math.min(outsideBounds[1], y);
        outsideBounds[2] = Math.max(outsideBounds[2], x);
        outsideBounds[3] = Math.max(outsideBounds[3], y);
      }
    }
  }
  assert.ok(changedPixels > 500, `${label} expected visible note/highlight changes, found ${changedPixels}`);
  assert.equal(changedOutsidePlacement, 0, `${label} changed pixels outside provider geometry: ${JSON.stringify({ changedBounds, outsideBounds, allowed })}`);
}

async function assertDuplicatePagePixelIdentity(source, output, mappings, tempRoot, label) {
  for (const [index, mapping] of mappings.entries()) {
    const sourceRender = path.join(tempRoot, `${label}-${index + 1}-source`);
    const outputRender = path.join(tempRoot, `${label}-${index + 1}-output`);
    run("pdftoppm", ["-f", String(mapping.sourcePage), "-l", String(mapping.sourcePage), "-singlefile", "-png", "-r", "144", source, sourceRender], { status: 0 });
    run("pdftoppm", ["-f", String(mapping.outputPage), "-l", String(mapping.outputPage), "-singlefile", "-png", "-r", "144", output, outputRender], { status: 0 });
    const sourcePixels = await sharp(`${sourceRender}.png`).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const outputPixels = await sharp(`${outputRender}.png`).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    assert.deepEqual(outputPixels.info, sourcePixels.info, `${label} page geometry changed for ${JSON.stringify(mapping)}`);
    assert.equal(Buffer.compare(outputPixels.data, sourcePixels.data), 0, `${label} rendered pixels changed for ${JSON.stringify(mapping)}`);
  }
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
  "tasks/repair_linearize.md",
  "tasks/ocr.md",
  "tasks/structure_clean.md",
  "scripts/mupdf.mjs",
  "scripts/pdf_provider.py",
  "scripts/qpdf_provider.py",
  "scripts/pyhanko_provider.py",
  "scripts/pyhanko_sign_provider.py",
  "scripts/verapdf_provider.py",
  "scripts/ocrmypdf_provider.py",
  "scripts/pikepdf_provider.py",
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
  "examples/pymupdf-ocr-redaction-operations.json",
  "examples/merge-stamp-manifest.json",
  "scripts/poppler_compare.py",
];
for (const file of requiredFiles) assert.ok(manifest.includes(file), `PDF manifest is missing ${file}`);

const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
assert.ok(skillText.split(/\r?\n/).length <= 210, "PDF Skill overview should route, not duplicate task references");
assert.match(skillText, /scripts\/mupdf\.mjs/);
assert.match(skillText, /MuPDF\.js/);
assert.match(skillText, /PdfProviders\.resolve/);
assert.match(skillText, /PdfProviders\.ensure/);
assert.match(skillText, /PdfProviders\.probe/);
assert.match(skillText, /PdfFile\.inspectPdf\("input\.pdf"\).*PdfProviders\.resolve/is);
assert.match(skillText, /status === "installable"[\s\S]*PdfProviders\.ensure[\s\S]*status !== "ready"/);
assert.match(skillText, /open-office-artifact-tool\/pdf\/providers/);
assert.match(skillText, /installPolicy: "disabled"/);
assert.match(skillText, /installPolicy": "managed"/);
assert.match(skillText, /system-only/);
assert.match(skillText, /non-MuPDF packs have no published immutable assets yet/);
assert.match(skillText, /hash-pinned.*versioned.*release assets/is);
assert.match(skillText, /silent fallback/i);
assert.match(skillText, /set_page_crop/);
assert.match(skillText, /rotate_page/);
assert.match(skillText, /duplicate_page.*source SHA-256.*only operation.*full\s+rewrite.*Poppler.*pixel identity/is);
assert.match(skillText, /delete_annotation.*update_annotation.*delete_link.*update_link.*update_form_field/is);
assert.match(skillText, /add_text_annotation.*visible pin.*rewrite/is);
assert.match(skillText, /add_text_highlight.*unique native text selection.*rewrite/is);
assert.match(skillText, /mupdf-page-space.*0\/90\/180\/270.*appearanceBbox/is);
assert.match(skillText, /raw `mediaBox`\/`cropBox`.*unrotated PDF-space/is);
assert.match(skillText, /sourceSha256/);
assert.match(skillText, /mupdf-link/);
assert.match(skillText, /virtual\s+environment\s+executable.*pyvenv\.cfg/s);
assert.match(skillText, /verapdf_provider\.py/);
assert.match(skillText, /ocrmypdf_provider\.py/);
assert.match(skillText, /redact_ocr_text/);
assert.match(skillText, /redact_ocr_text.*expected_rotation.*0.*90.*180.*270.*unrotated PyMuPDF page space/is);
assert.match(skillText, /pikepdf_provider\.py/);
assert.match(skillText, /pyhanko_sign_provider\.py/);
assert.match(skillText, /passphrase.*stdin/is);
assert.match(skillText, /local PKCS#12/i);
assert.match(skillText, /timestamp.*LTV.*external/is);
assert.match(skillText, /active-content.*active-and-auxiliary/is);
assert.match(skillText, /not.*redaction.*metadata.*XFA/is);
assert.match(skillText, /complete imported PDF.*not a sanitizer/is);
assert.match(skillText, /machine-rule gate.*human review.*PDF\/UA/is);
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
]) assert.match(skillText, pattern);
assert.doesNotMatch(skillText, /brew install|apt-get|uv pip install/i);
const providerSetupText = await fs.readFile(path.join(skillRoot, "tasks", "provider_setup.md"), "utf8");
assert.match(providerSetupText, /PdfProviders\.resolve/);
assert.match(providerSetupText, /PdfProviders\.ensure/);
assert.match(providerSetupText, /PdfProviders\.probe/);
assert.match(providerSetupText, /symlink\/hardlink/);
assert.match(providerSetupText, /enterprise mirror.*identical hash-pinned bytes/is);
assert.match(providerSetupText, /Current catalog state.*not yet\s+published/is);
assert.doesNotMatch(providerSetupText, /brew install|apt-get|uv pip install/i);
const providerMatrixText = await fs.readFile(path.join(skillRoot, "references", "PROVIDER_MATRIX.md"), "utf8");
assert.match(providerMatrixText, /does \*\*not\*\* duplicate.*versions.*hashes.*URLs/is);
assert.doesNotMatch(providerMatrixText, /1\.28\.0|1\.27\.2|10\.10\.x|17\.8\.x/i);
const pdfPluginReadme = await fs.readFile(path.join(repoRoot, "skills", "pdf", "README.md"), "utf8");
assert.match(pdfPluginReadme, /open-office-artifact-tool\/pdf\/providers/);
assert.match(pdfPluginReadme, /system-only.*hash-pinned managed pack/is);
assert.match(pdfPluginReadme, /managed release assets are not published yet.*blocked/is);
assert.doesNotMatch(pdfPluginReadme, /remain separately installed|brew install|apt-get|uv pip install/i);
const nativePlacementDocs = [
  await fs.readFile(path.join(skillRoot, "artifact_tool", "API_QUICK_START.md"), "utf8"),
  await fs.readFile(path.join(skillRoot, "tasks", "forms_annotations.md"), "utf8"),
  await fs.readFile(path.join(skillRoot, "tasks", "edit_existing.md"), "utf8"),
  await fs.readFile(path.join(skillRoot, "references", "PROVIDER_MATRIX.md"), "utf8"),
].join("\n");
assert.match(nativePlacementDocs, /mupdf-page-space/);
assert.match(nativePlacementDocs, /0\/90\/180\/270/);
assert.match(nativePlacementDocs, /appearanceBbox/);
assert.match(nativePlacementDocs, /raw (?:unrotated )?`?mediaBox`?\/`?cropBox`?.*not.*placement/is);
assert.doesNotMatch(nativePlacementDocs, /`add_link` accepts only an unrotated/);
assert.doesNotMatch(nativePlacementDocs, /rotated page, stale hash\/page snapshot/);
const redactTaskText = await fs.readFile(path.join(skillRoot, "tasks", "redact.md"), "utf8");
assert.match(redactTaskText, /expected_rotation/);
assert.match(redactTaskText, /temporarily clears `\/Rotate`.*restores `\/Rotate`/s);
assert.doesNotMatch(redactTaskText, /Rotated pages must first become a separately reviewed normalized version/);

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
  "assert module.normalize_ocr_language('eng+snum') == 'eng+snum'",
  "assert module.validate_ocr_dpi(150) == 150",
].join(";"), path.join(scriptsRoot, "pymupdf_edit.py")], { status: 0 });
assert.equal(fitUnit.stderr, "");

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-pdf-provider-skill-"));
try {
  // A configured provider virtual environment must survive re-exec. Its
  // bin/python is commonly a symlink to a base interpreter, so resolving the
  // target before exec would silently lose the venv site-packages.
  const runtimeVenv = path.join(tempRoot, "provider-runtime-venv");
  run(python, ["-m", "venv", runtimeVenv], { status: 0 });
  const runtimePython = process.platform === "win32"
    ? path.join(runtimeVenv, "Scripts", "python.exe")
    : path.join(runtimeVenv, "bin", "python");
  const runtimeProbe = path.join(tempRoot, "provider-runtime-probe.py");
  await fs.writeFile(runtimeProbe, [
    "import json, sys",
    "sys.path.insert(0, sys.argv[1])",
    "from python_runtime import reexec_configured_provider_python",
    "reexec_configured_provider_python()",
    "print(json.dumps({'executable': sys.executable, 'prefix': sys.prefix}))",
  ].join("\n"), "utf8");
  const runtimeResult = parseResult(run(python, [runtimeProbe, scriptsRoot], {
    env: { OPEN_OFFICE_PDF_PROVIDER_PYTHON: runtimePython },
    status: 0,
  }));
  assert.equal(await fs.realpath(runtimeResult.prefix), await fs.realpath(runtimeVenv));

  const mupdfCli = path.join(scriptsRoot, "mupdf.mjs");
  const mupdfInput = path.join(tempRoot, "mupdf-input.pdf");
  const mupdfRender = path.join(tempRoot, "mupdf-render.png");
  const mupdfOperations = path.join(tempRoot, "mupdf-operations.json");
  const mupdfOutput = path.join(tempRoot, "mupdf-output.pdf");
  const mupdfHighlightOperations = path.join(tempRoot, "mupdf-highlight-operations.json");
  const mupdfHighlightOutput = path.join(tempRoot, "mupdf-highlight-output.pdf");
  const mupdfCropOperations = path.join(tempRoot, "mupdf-crop-operations.json");
  const mupdfCropOutput = path.join(tempRoot, "mupdf-crop-output.pdf");
  const mupdfRotationOperations = path.join(tempRoot, "mupdf-rotation-operations.json");
  const mupdfRotationOutput = path.join(tempRoot, "mupdf-rotation-output.pdf");
  const mupdfDuplicateInput = path.join(tempRoot, "mupdf-duplicate-input.pdf");
  const mupdfDuplicateOperations = path.join(tempRoot, "mupdf-duplicate-operations.json");
  const mupdfDuplicateOutput = path.join(tempRoot, "mupdf-duplicate-output.pdf");
  const mupdfAnnotationUpdateOperations = path.join(tempRoot, "mupdf-annotation-update-operations.json");
  const mupdfAnnotationUpdateOutput = path.join(tempRoot, "mupdf-annotation-update-output.pdf");
  const mupdfAnnotationDeleteOperations = path.join(tempRoot, "mupdf-annotation-delete-operations.json");
  const mupdfAnnotationDeleteOutput = path.join(tempRoot, "mupdf-annotation-delete-output.pdf");
  const mupdfLinkUpdateOperations = path.join(tempRoot, "mupdf-link-update-operations.json");
  const mupdfLinkUpdateOutput = path.join(tempRoot, "mupdf-link-update-output.pdf");
  const mupdfLinkMoveOperations = path.join(tempRoot, "mupdf-link-move-operations.json");
  const mupdfLinkMoveOutput = path.join(tempRoot, "mupdf-link-move-output.pdf");
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
  const mupdfLinkPage = mupdfInspection.records.find((record) => record.kind === "mupdfPage" && record.page === mupdfLink.page);
  assert.match(mupdfLink.id, /^mupdf-link-1-[a-f0-9]{64}$/);
  assert.ok(mupdfLinkPage);
  await fs.writeFile(mupdfLinkMoveOperations, JSON.stringify({
    savePolicy: "rewrite",
    operations: [
      {
        type: "delete_link",
        page: mupdfLink.page,
        linkId: mupdfLink.id,
        sourceSha256: mupdfInspection.summary.sourceSha256,
        expected: { url: mupdfLink.url, bbox: mupdfLink.bbox, external: mupdfLink.external },
      },
      {
        type: "add_link",
        page: mupdfLinkPage.page,
        sourceSha256: mupdfInspection.summary.sourceSha256,
        expectedPage: { bbox: mupdfLinkPage.bbox, rotation: mupdfLinkPage.rotation },
        bbox: [156, 180, 96, 18],
        url: mupdfLink.url,
      },
    ],
  }), "utf8");
  const mupdfLinkMoved = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfInput, mupdfLinkMoveOperations, mupdfLinkMoveOutput], { status: 0 }));
  assert.deepEqual(mupdfLinkMoved.operations.map((operation) => operation.type), ["delete_link", "add_link"]);
  const mupdfLinkMovedInspection = await PdfFile.inspectPdf(await fs.readFile(mupdfLinkMoveOutput));
  const mupdfMovedLink = mupdfLinkMovedInspection.records.find((record) => record.kind === "mupdfLink");
  assert.ok(mupdfMovedLink);
  assert.equal(mupdfMovedLink.url, mupdfLink.url);
  assert.deepEqual(mupdfMovedLink.bbox, [156, 180, 96, 18]);
  await fs.writeFile(mupdfLinkUpdateOperations, JSON.stringify({
    savePolicy: "rewrite",
    operations: [{
      type: "update_link",
      page: mupdfLink.page,
      linkId: mupdfLink.id,
      sourceSha256: mupdfInspection.summary.sourceSha256,
      expected: { url: mupdfLink.url, bbox: mupdfLink.bbox, external: mupdfLink.external },
      patch: { url: "https://example.com/cli-link-updated" },
    }],
  }), "utf8");
  const mupdfLinkUpdated = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfInput, mupdfLinkUpdateOperations, mupdfLinkUpdateOutput], { status: 0 }));
  assert.equal(mupdfLinkUpdated.savePolicy, "rewrite");
  assert.equal(mupdfLinkUpdated.operations[0].type, "update_link");
  const mupdfLinkUpdatedInspection = await PdfFile.inspectPdf(await fs.readFile(mupdfLinkUpdateOutput));
  const mupdfUpdatedLink = mupdfLinkUpdatedInspection.records.find((record) => record.kind === "mupdfLink");
  assert.ok(mupdfUpdatedLink);
  assert.equal(mupdfUpdatedLink.url, "https://example.com/cli-link-updated");
  assert.deepEqual(mupdfUpdatedLink.bbox, mupdfLink.bbox);
  assert.notEqual(mupdfLinkUpdatedInspection.summary.sourceSha256, mupdfInspection.summary.sourceSha256);
  await fs.writeFile(mupdfLinkDeleteOperations, JSON.stringify({
    savePolicy: "rewrite",
    operations: [{
      type: "delete_link",
      page: mupdfUpdatedLink.page,
      linkId: mupdfUpdatedLink.id,
      sourceSha256: mupdfLinkUpdatedInspection.summary.sourceSha256,
      expected: { url: mupdfUpdatedLink.url, bbox: mupdfUpdatedLink.bbox, external: mupdfUpdatedLink.external },
    }],
  }), "utf8");
  const mupdfLinkDeleted = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfLinkUpdateOutput, mupdfLinkDeleteOperations, mupdfLinkDeleteOutput], { status: 0 }));
  assert.equal(mupdfLinkDeleted.savePolicy, "rewrite");
  assert.equal(mupdfLinkDeleted.operations[0].type, "delete_link");
  assert.equal((await PdfFile.inspectPdf(await fs.readFile(mupdfLinkDeleteOutput))).records.some((record) => record.kind === "mupdfLink"), false);
  const mupdfRendered = parseResult(run(process.execPath, [mupdfCli, "render", mupdfInput, mupdfRender, "--page", "1", "--dpi", "72"], { status: 0 }));
  assert.equal(mupdfRendered.provider, "mupdf");
  assert.deepEqual([...new Uint8Array(await fs.readFile(mupdfRender)).subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
  const mupdfNestedRender = path.join(tempRoot, "nested", "qa", "mupdf-render.png");
  run(process.execPath, [mupdfCli, "render", mupdfInput, mupdfNestedRender, "--page", "1", "--dpi", "72"], { status: 0 });
  assert.deepEqual([...new Uint8Array(await fs.readFile(mupdfNestedRender)).subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
  const mupdfAnnotationPage = mupdfInspection.records.find((record) => record.kind === "mupdfPage" && record.page === 1);
  await fs.writeFile(mupdfHighlightOperations, JSON.stringify({
    savePolicy: "rewrite",
    operations: [{
      type: "add_text_highlight",
      page: mupdfAnnotationPage.page,
      sourceSha256: mupdfInspection.summary.sourceSha256,
      expectedPage: { bbox: mupdfAnnotationPage.bbox, rotation: mupdfAnnotationPage.rotation },
      text: "MuPDF Skill CLI fixture",
      color: [0.2, 0.8, 0.3],
      contents: "CLI highlight",
      author: "CLI reviewer",
    }],
  }), "utf8");
  const mupdfHighlighted = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfInput, mupdfHighlightOperations, mupdfHighlightOutput], { status: 0 }));
  assert.equal(mupdfHighlighted.savePolicy, "rewrite");
  assert.equal(mupdfHighlighted.operations[0].type, "add_text_highlight");
  assert.equal(mupdfHighlighted.operations[0].added.contents, "CLI highlight");
  const mupdfHighlightInspection = await PdfFile.inspectPdf(await fs.readFile(mupdfHighlightOutput));
  const mupdfHighlight = mupdfHighlightInspection.records.find((record) => record.kind === "mupdfAnnotation" && record.type === "Highlight");
  assert.ok(mupdfHighlight);
  assert.ok(mupdfHighlight.color.every((component, index) => Math.abs(component - [0.2, 0.8, 0.3][index]) < 0.001));
  assert.equal(mupdfHighlight.quadPoints.length, 1);
  await fs.writeFile(mupdfOperations, JSON.stringify({
    savePolicy: "rewrite",
    operations: [{
      type: "add_text_annotation",
      page: 1,
      sourceSha256: mupdfInspection.summary.sourceSha256,
      expectedPage: { bbox: mupdfAnnotationPage.bbox, rotation: mupdfAnnotationPage.rotation },
      point: [40, 40],
      contents: "CLI review",
    }],
  }), "utf8");
  const mupdfEdited = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfInput, mupdfOperations, mupdfOutput], { status: 0 }));
  assert.equal(mupdfEdited.savePolicy, "rewrite");
  assert.equal(mupdfEdited.operations[0].added.contents, "CLI review");
  const mupdfAnnotationInspection = await PdfFile.inspectPdf(await fs.readFile(mupdfOutput));
  assert.equal(mupdfAnnotationInspection.records.find((record) => record.kind === "mupdfPage").annotations, 1);
  const mupdfAnnotation = mupdfAnnotationInspection.records.find((record) => record.kind === "mupdfAnnotation");
  assert.match(mupdfAnnotation.id, /^mupdf-annotation-1-\d+$/);
  await fs.writeFile(mupdfAnnotationUpdateOperations, JSON.stringify({
    savePolicy: "rewrite",
    operations: [{
      type: "update_annotation",
      page: mupdfAnnotation.page,
      annotationId: mupdfAnnotation.id,
      sourceSha256: mupdfAnnotationInspection.summary.sourceSha256,
      expected: { type: mupdfAnnotation.type, contents: mupdfAnnotation.contents, rect: mupdfAnnotation.rect },
      patch: { contents: "CLI review updated", author: "CLI reviewer", subject: "Resolved" },
    }],
  }), "utf8");
  const mupdfAnnotationUpdated = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfOutput, mupdfAnnotationUpdateOperations, mupdfAnnotationUpdateOutput], { status: 0 }));
  assert.equal(mupdfAnnotationUpdated.savePolicy, "rewrite");
  assert.equal(mupdfAnnotationUpdated.operations[0].type, "update_annotation");
  const mupdfAnnotationUpdatedInspection = await PdfFile.inspectPdf(await fs.readFile(mupdfAnnotationUpdateOutput));
  const mupdfUpdatedAnnotation = mupdfAnnotationUpdatedInspection.records.find((record) => record.kind === "mupdfAnnotation");
  assert.equal(mupdfUpdatedAnnotation.contents, "CLI review updated");
  assert.equal(mupdfUpdatedAnnotation.author, "CLI reviewer");
  assert.equal(mupdfUpdatedAnnotation.subject, "Resolved");
  assert.notEqual(mupdfAnnotationUpdatedInspection.summary.sourceSha256, mupdfAnnotationInspection.summary.sourceSha256);
  await fs.writeFile(mupdfAnnotationDeleteOperations, JSON.stringify({
    savePolicy: "rewrite",
    operations: [{
      type: "delete_annotation",
      page: mupdfUpdatedAnnotation.page,
      annotationId: mupdfUpdatedAnnotation.id,
      sourceSha256: mupdfAnnotationUpdatedInspection.summary.sourceSha256,
      expected: { type: mupdfUpdatedAnnotation.type, contents: mupdfUpdatedAnnotation.contents, rect: mupdfUpdatedAnnotation.rect },
    }],
  }), "utf8");
  const mupdfAnnotationDeleted = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfAnnotationUpdateOutput, mupdfAnnotationDeleteOperations, mupdfAnnotationDeleteOutput], { status: 0 }));
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
  const hasQpdf = run("qpdf", ["--version"]).status === 0;
  const hasPoppler = run("pdftoppm", ["-v"]).status === 0;
  const hasPdfInfo = run("pdfinfo", ["-v"]).status === 0;
  await fs.writeFile(mupdfDuplicateInput, plainPdfBytes([
    { text: "CLI PAGE ONE", width: 612, height: 792 },
    { text: "CLI ROTATED PAGE TWO", width: 540, height: 720, rotation: 90 },
    { text: "CLI PAGE THREE", width: 420, height: 600 },
  ]));
  const mupdfDuplicateSourceHash = crypto.createHash("sha256").update(await fs.readFile(mupdfDuplicateInput)).digest("hex");
  const mupdfDuplicateInspection = parseResult(run(process.execPath, [mupdfCli, "inspect", mupdfDuplicateInput], { status: 0 }));
  const mupdfDuplicatePage = mupdfDuplicateInspection.records.find((record) => record.kind === "mupdfPage" && record.page === 2);
  await fs.writeFile(mupdfDuplicateOperations, JSON.stringify({
    savePolicy: "rewrite",
    operations: [{
      type: "duplicate_page",
      page: 2,
      sourceSha256: mupdfDuplicateInspection.summary.sourceSha256,
      expectedPage: { bbox: mupdfDuplicatePage.bbox, rotation: mupdfDuplicatePage.rotation },
    }],
  }), "utf8");
  const mupdfDuplicated = parseResult(run(process.execPath, [mupdfCli, "edit", mupdfDuplicateInput, mupdfDuplicateOperations, mupdfDuplicateOutput], { status: 0 }));
  const { objectCountBefore: cliObjectCountBefore, objectCountAfter: cliObjectCountAfter, ...cliDuplicateAudit } = mupdfDuplicated.operations[0];
  assert.deepEqual(cliDuplicateAudit, {
    type: "duplicate_page",
    sourcePage: 2,
    sourcePageAfterInsertion: 2,
    insertedPage: 3,
    insertAt: 3,
    expectedPage: { bbox: [0, 0, 720, 540], rotation: 90 },
    pageCountBefore: 3,
    pageCountAfter: 4,
    interactiveObjectsCopied: 0,
    taggedInput: false,
    navigationSynthesized: false,
  });
  assert.equal(cliObjectCountBefore, mupdfDuplicateInspection.summary.nativeObjects);
  assert.ok(cliObjectCountAfter > cliObjectCountBefore);
  assert.equal(crypto.createHash("sha256").update(await fs.readFile(mupdfDuplicateInput)).digest("hex"), mupdfDuplicateSourceHash);
  const mupdfDuplicateOutputInspection = parseResult(run(process.execPath, [mupdfCli, "inspect", mupdfDuplicateOutput], { status: 0 }));
  assert.equal(mupdfDuplicateOutputInspection.summary.pages, 4);
  const mupdfDuplicateOutputPages = mupdfDuplicateOutputInspection.records.filter((record) => record.kind === "mupdfPage");
  assert.deepEqual(mupdfDuplicateOutputPages.map((record) => record.rotation), [0, 90, 90, 0]);
  assert.ok(mupdfDuplicateOutputPages.every((record) => record.annotations === 0 && record.widgets === 0 && record.links === 0));
  if (hasQpdf) run("qpdf", ["--check", mupdfDuplicateOutput], { status: 0 });
  if (hasPdfInfo) assert.match(run("pdfinfo", [mupdfDuplicateOutput], { status: 0 }).stdout, /Pages:\s+4/);
  if (hasPoppler) await assertDuplicatePagePixelIdentity(mupdfDuplicateInput, mupdfDuplicateOutput, [
    { sourcePage: 1, outputPage: 1 },
    { sourcePage: 2, outputPage: 2 },
    { sourcePage: 2, outputPage: 3 },
    { sourcePage: 3, outputPage: 4 },
  ], tempRoot, "mupdf-duplicate-page");
  for (const rotation of [90, 180, 270]) {
    const rotatedSource = rotation === 90
      ? mupdfRotationOutput
      : path.join(tempRoot, `mupdf-native-placement-${rotation}-source.pdf`);
    if (rotation !== 90) {
      const rotateOperations = path.join(tempRoot, `mupdf-native-placement-${rotation}-rotate.json`);
      await fs.writeFile(rotateOperations, JSON.stringify({
        savePolicy: "rewrite",
        operations: [{ type: "rotate_page", page: 1, rotation }],
      }), "utf8");
      run(process.execPath, [mupdfCli, "edit", mupdfInput, rotateOperations, rotatedSource], { status: 0 });
    }
    const sourceBytes = await fs.readFile(rotatedSource);
    const sourceHash = crypto.createHash("sha256").update(sourceBytes).digest("hex");
    const sourceInspection = parseResult(run(process.execPath, [mupdfCli, "inspect", rotatedSource], { status: 0 }));
    const sourcePage = sourceInspection.records.find((record) => record.kind === "mupdfPage" && record.page === 1);
    assert.equal(sourcePage.rotation, rotation);
    assert.equal(sourcePage.coordinateSpace, "mupdf-page-space");
    assert.deepEqual(sourcePage.bbox, rotation === 180 ? [0, 0, 612, 792] : [0, 0, 792, 612]);

    const placementOperations = path.join(tempRoot, `mupdf-native-placement-${rotation}-operations.json`);
    const placementOutput = path.join(tempRoot, `mupdf-native-placement-${rotation}-output.pdf`);
    await fs.writeFile(placementOperations, JSON.stringify({
      savePolicy: "rewrite",
      operations: [
        {
          type: "add_text_annotation",
          page: 1,
          sourceSha256: sourceInspection.summary.sourceSha256,
          expectedPage: { bbox: sourcePage.bbox, rotation: sourcePage.rotation },
          point: [40, 40],
          contents: `Rotated CLI review ${rotation}`,
        },
        {
          type: "add_text_highlight",
          page: 1,
          sourceSha256: sourceInspection.summary.sourceSha256,
          expectedPage: { bbox: sourcePage.bbox, rotation: sourcePage.rotation },
          text: "MuPDF Skill CLI fixture",
          color: [0.9, 0.7, 0.1],
        },
        {
          type: "add_link",
          page: 1,
          sourceSha256: sourceInspection.summary.sourceSha256,
          expectedPage: { bbox: sourcePage.bbox, rotation: sourcePage.rotation },
          bbox: [300, 400, 100, 18],
          url: `https://example.com/native-placement-${rotation}`,
        },
      ],
    }), "utf8");
    const placed = parseResult(run(process.execPath, [mupdfCli, "edit", rotatedSource, placementOperations, placementOutput], { status: 0 }));
    assert.deepEqual(
      placed.operations.map((operation) => ({ type: operation.type, coordinateSpace: operation.coordinateSpace, pageRotation: operation.pageRotation })),
      ["add_text_annotation", "add_text_highlight", "add_link"].map((type) => ({ type, coordinateSpace: "mupdf-page-space", pageRotation: rotation })),
    );
    const expectedNoteAppearance = {
      90: [60, 40, 20, 20],
      180: [60, 60, 20, 20],
      270: [40, 60, 20, 20],
    }[rotation];
    assert.deepEqual(placed.operations[0].added.appearanceBbox, expectedNoteAppearance);
    assert.equal(crypto.createHash("sha256").update(await fs.readFile(rotatedSource)).digest("hex"), sourceHash);
    const outputInspection = parseResult(run(process.execPath, [mupdfCli, "inspect", placementOutput], { status: 0 }));
    const outputPage = outputInspection.records.find((record) => record.kind === "mupdfPage" && record.page === 1);
    const outputNote = outputInspection.records.find((record) => record.kind === "mupdfAnnotation" && record.contents === `Rotated CLI review ${rotation}`);
    const outputHighlight = outputInspection.records.find((record) => record.kind === "mupdfAnnotation" && record.type === "Highlight");
    const outputLink = outputInspection.records.find((record) => record.kind === "mupdfLink" && record.url === `https://example.com/native-placement-${rotation}`);
    assert.equal(outputPage.rotation, rotation);
    assert.deepEqual(outputPage.bbox, sourcePage.bbox);
    assert.deepEqual(outputNote.rect, [40, 40, 20, 20]);
    assert.deepEqual(outputNote.appearanceBbox, expectedNoteAppearance);
    assert.deepEqual(outputNote.appearanceBbox, placed.operations[0].added.appearanceBbox);
    assert.equal(outputHighlight.quadPoints.length, 1);
    assert.deepEqual(outputHighlight.appearanceBbox, placed.operations[1].added.appearanceBbox);
    assert.deepEqual(outputLink.bbox, [300, 400, 100, 18]);
    if (hasQpdf) run("qpdf", ["--check", placementOutput], { status: 0 });
    if (hasPdfInfo) assert.match(run("pdfinfo", [placementOutput], { status: 0 }).stdout, new RegExp(`Page rot:\\s+${rotation}`));
    if (hasPoppler) await assertNativePlacementPixelScope(rotatedSource, placementOutput, sourcePage, placed.operations, tempRoot, `mupdf-native-placement-${rotation}`);
  }
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
    assert.ok(routedProbe.operations.includes("redact_ocr_text"));
    assert.equal(routedProbe.ocr.available, true);
    assert.equal(routedProbe.ocr.language, "eng");
    assert.deepEqual(routedProbe.ocr.supportedPageRotations, [0, 90, 180, 270]);
    assert.equal(routedProbe.ocr.rotationPrecondition, "expected_rotation");
    assert.equal(routedProbe.ocr.coordinateSpace, "unrotated-pymupdf-page-space");
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

    const rasterSecret = "IMAGE SECRET 8842";
    const rasterImage = path.join(tempRoot, "ocr-secret.png");
    await sharp(Buffer.from([
      '<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="300">',
      '<rect width="1200" height="300" fill="white"/>',
      '<text x="60" y="190" font-family="sans-serif" font-size="72" font-weight="700" fill="black">IMAGE SECRET 8842</text>',
      '</svg>',
    ].join(""))).png().toFile(rasterImage);
    const rasterSource = path.join(tempRoot, "ocr-redaction-source.pdf");
    run(integrationPython, ["-c", [
      "from reportlab.lib.pagesizes import letter",
      "from reportlab.pdfgen import canvas",
      "import sys",
      "c=canvas.Canvas(sys.argv[1],pagesize=letter,invariant=1)",
      "c.setFont('Helvetica',12)",
      "c.drawString(72,720,'Public heading remains unchanged')",
      "c.drawImage(sys.argv[2],72,500,width=468,height=117,mask='auto')",
      "c.save()",
    ].join(";"), rasterSource, rasterImage], { status: 0 });
    const rasterSourceBytes = await fs.readFile(rasterSource);
    const rasterSourceHash = crypto.createHash("sha256").update(rasterSourceBytes).digest("hex");
    const rasterSourceScan = parseResult(run(integrationPython, [
      path.join(scriptsRoot, "residue_scan.py"), rasterSource,
      "--term", rasterSecret, "--require-ocr", "--require-single-revision", "--ocr-dpi", "200",
    ], { status: 2 }));
    assert.equal(rasterSourceScan.ok, false);
    assert.ok(rasterSourceScan.terms[rasterSecret].evidence.some((entry) => entry.category === "image-ocr"));

    const rasterOperations = path.join(tempRoot, "ocr-redaction-operations.json");
    await fs.writeFile(rasterOperations, JSON.stringify([{
      type: "redact_ocr_text",
      page: 1,
      term: rasterSecret,
      expected_matches: 1,
      fill: [0, 0, 0],
    }]), "utf8");
    const rasterOutput = path.join(tempRoot, "ocr-redaction-output.pdf");
    const rasterResult = parseResult(run(integrationPython, [
      path.join(scriptsRoot, "pymupdf_edit.py"), "edit", rasterSource, rasterOutput,
      "--strategy", "sanitize", "--operations", rasterOperations,
      "--sensitive-term", rasterSecret, "--ocr-language", "eng", "--ocr-dpi", "200",
      "--accept-license", "agpl", "--invalidate-signatures",
    ], { status: 0 }));
    const rasterOperation = rasterResult.operations.find((operation) => operation.type === "redact_ocr_text");
    assert.equal(rasterOperation.matches, 1);
    assert.equal(rasterOperation.expectedMatches, 1);
    assert.equal(rasterOperation.ocr.language, "eng");
    assert.equal(rasterOperation.ocr.dpi, 200);
    assert.equal(rasterOperation.pageRotation, 0);
    assert.equal(rasterOperation.expectedRotation, 0);
    assert.equal(rasterOperation.coordinateSpace, "unrotated-pymupdf-page-space");
    assert.ok(rasterOperation.minimumImageCoverage >= 0.9);
    assert.equal(rasterOperation.rects.length, 1);
    assert.equal(rasterOperation.displayRects.length, 1);
    assert.equal(rasterOperation.imageRects.length, 1);
    assert.equal(rasterOperation.displayImageRects.length, 1);
    assert.equal(rasterResult.originalPrefixPreserved, false);
    assert.equal(rasterResult.residueScan.ok, true);
    assert.deepEqual(rasterResult.ocrRotationChecks, [{ page: 1, expectedRotation: 0, actualRotation: 0, preserved: true }]);
    assert.equal(crypto.createHash("sha256").update(await fs.readFile(rasterSource)).digest("hex"), rasterSourceHash);
    const rasterOutputScan = parseResult(run(integrationPython, [
      path.join(scriptsRoot, "residue_scan.py"), rasterOutput,
      "--term", rasterSecret, "--require-ocr", "--require-inert", "--require-single-revision", "--ocr-dpi", "200",
    ], { status: 0 }));
    assert.equal(rasterOutputScan.ok, true);
    assert.equal(rasterOutputScan.terms[rasterSecret].matches, 0);

    const rasterMismatchOperations = path.join(tempRoot, "ocr-redaction-mismatch.json");
    await fs.writeFile(rasterMismatchOperations, JSON.stringify([{
      type: "redact_ocr_text",
      page: 1,
      term: rasterSecret,
      expected_matches: 2,
    }]), "utf8");
    const rasterMismatchOutput = path.join(tempRoot, "ocr-redaction-mismatch.pdf");
    const rasterMismatch = run(integrationPython, [
      path.join(scriptsRoot, "pymupdf_edit.py"), "edit", rasterSource, rasterMismatchOutput,
      "--strategy", "sanitize", "--operations", rasterMismatchOperations,
      "--sensitive-term", rasterSecret, "--ocr-dpi", "200",
      "--accept-license", "agpl", "--invalidate-signatures",
    ], { status: 2 });
    assert.match(rasterMismatch.stderr, /expected 2 image-backed match\(es\).*found 1/);
    await assert.rejects(fs.access(rasterMismatchOutput));

    const rasterMissingLanguageOutput = path.join(tempRoot, "ocr-redaction-missing-language.pdf");
    const rasterMissingLanguage = run(integrationPython, [
      path.join(scriptsRoot, "pymupdf_edit.py"), "edit", rasterSource, rasterMissingLanguageOutput,
      "--strategy", "sanitize", "--operations", rasterOperations,
      "--sensitive-term", rasterSecret, "--ocr-language", "definitely_missing",
      "--accept-license", "agpl", "--invalidate-signatures",
    ], { status: 2 });
    assert.match(rasterMissingLanguage.stderr, /language data is unavailable/);
    await assert.rejects(fs.access(rasterMissingLanguageOutput));

    const rasterUnsafeLanguageOutput = path.join(tempRoot, "ocr-redaction-unsafe-language.pdf");
    const rasterUnsafeLanguage = run(integrationPython, [
      path.join(scriptsRoot, "pymupdf_edit.py"), "edit", rasterSource, rasterUnsafeLanguageOutput,
      "--strategy", "sanitize", "--operations", rasterOperations,
      "--sensitive-term", rasterSecret, "--ocr-language", "../eng",
      "--accept-license", "agpl", "--invalidate-signatures",
    ], { status: 2 });
    assert.match(rasterUnsafeLanguage.stderr, /safe Tesseract language codes/);
    await assert.rejects(fs.access(rasterUnsafeLanguageOutput));

    const rasterExcessiveDpiOutput = path.join(tempRoot, "ocr-redaction-excessive-dpi.pdf");
    const rasterExcessiveDpi = run(integrationPython, [
      path.join(scriptsRoot, "pymupdf_edit.py"), "edit", rasterSource, rasterExcessiveDpiOutput,
      "--strategy", "sanitize", "--operations", rasterOperations,
      "--sensitive-term", rasterSecret, "--ocr-dpi", "301",
      "--accept-license", "agpl", "--invalidate-signatures",
    ], { status: 2 });
    assert.match(rasterExcessiveDpi.stderr, /OCR dpi must be an integer between 72 and 300/);
    await assert.rejects(fs.access(rasterExcessiveDpiOutput));

    const rotatedRasterCases = [];
    for (const rotation of [90, 180, 270]) {
      const rotatedSource = path.join(tempRoot, `ocr-redaction-rotated-${rotation}-source.pdf`);
      run(integrationPython, ["-c", [
        "import pymupdf,sys",
        "doc=pymupdf.open(sys.argv[1])",
        "doc[0].set_rotation(int(sys.argv[3]))",
        "doc.save(sys.argv[2])",
        "doc.close()",
      ].join(";"), rasterSource, rotatedSource, String(rotation)], { status: 0 });
      const rotatedSourceBytes = await fs.readFile(rotatedSource);
      const rotatedSourceHash = crypto.createHash("sha256").update(rotatedSourceBytes).digest("hex");
      const rotatedSourceScan = parseResult(run(integrationPython, [
        path.join(scriptsRoot, "residue_scan.py"), rotatedSource,
        "--term", rasterSecret, "--require-ocr", "--require-single-revision", "--ocr-dpi", "200",
      ], { status: 2 }));
      assert.equal(rotatedSourceScan.pages[0].rotation, rotation);
      assert.equal(rotatedSourceScan.pages[0].ocr.pageRotation, rotation);
      assert.ok(rotatedSourceScan.terms[rasterSecret].evidence.some((entry) => entry.category === "image-ocr"));

      const rotatedOperations = path.join(tempRoot, `ocr-redaction-rotated-${rotation}-operations.json`);
      await fs.writeFile(rotatedOperations, JSON.stringify([{
        type: "redact_ocr_text",
        page: 1,
        expected_rotation: rotation,
        term: rasterSecret,
        expected_matches: 1,
        fill: [0, 0, 0],
      }]), "utf8");
      const rotatedOutput = path.join(tempRoot, `ocr-redaction-rotated-${rotation}-output.pdf`);
      const rotatedResult = parseResult(run(integrationPython, [
        path.join(scriptsRoot, "pymupdf_edit.py"), "edit", rotatedSource, rotatedOutput,
        "--strategy", "sanitize", "--operations", rotatedOperations,
        "--sensitive-term", rasterSecret, "--ocr-language", "eng", "--ocr-dpi", "200",
        "--accept-license", "agpl", "--invalidate-signatures",
      ], { status: 0 }));
      const rotatedOperation = rotatedResult.operations.find((operation) => operation.type === "redact_ocr_text");
      assert.equal(rotatedOperation.pageRotation, rotation);
      assert.equal(rotatedOperation.expectedRotation, rotation);
      assert.equal(rotatedOperation.matches, 1);
      assert.equal(rotatedOperation.coordinateSpace, "unrotated-pymupdf-page-space");
      assert.notDeepEqual(rotatedOperation.displayRects, rotatedOperation.rects);
      assert.equal(rotatedResult.residueScan.ok, true);
      assert.deepEqual(rotatedResult.ocrRotationChecks, [{
        page: 1,
        expectedRotation: rotation,
        actualRotation: rotation,
        preserved: true,
      }]);
      assert.equal(crypto.createHash("sha256").update(await fs.readFile(rotatedSource)).digest("hex"), rotatedSourceHash);
      const rotatedOutputScan = parseResult(run(integrationPython, [
        path.join(scriptsRoot, "residue_scan.py"), rotatedOutput,
        "--term", rasterSecret, "--require-ocr", "--require-inert", "--require-single-revision", "--ocr-dpi", "200",
      ], { status: 0 }));
      assert.equal(rotatedOutputScan.pages[0].rotation, rotation);
      assert.equal(rotatedOutputScan.terms[rasterSecret].matches, 0);
      rotatedRasterCases.push({ rotation, source: rotatedSource, output: rotatedOutput, operation: rotatedOperation });
    }

    const rasterStaleRotationOutput = path.join(tempRoot, "ocr-redaction-stale-rotation.pdf");
    const rasterStaleRotation = run(integrationPython, [
      path.join(scriptsRoot, "pymupdf_edit.py"), "edit", rotatedRasterCases[0].source, rasterStaleRotationOutput,
      "--strategy", "sanitize", "--operations", rasterOperations,
      "--sensitive-term", rasterSecret, "--ocr-dpi", "200",
      "--accept-license", "agpl", "--invalidate-signatures",
    ], { status: 2 });
    assert.match(rasterStaleRotation.stderr, /expected_rotation 0 does not match page 1 rotation 90/);
    await assert.rejects(fs.access(rasterStaleRotationOutput));

    if (run("qpdf", ["--version"]).status === 0) run("qpdf", ["--check", rasterOutput], { status: 0 });
    if (run("pdftoppm", ["-v"]).status === 0) {
      await assertOcrRedactionPixelScope(rasterSource, rasterOutput, rasterOperation, tempRoot, "ocr-rotation-0");
      for (const rotated of rotatedRasterCases) {
        await assertOcrRedactionPixelScope(rotated.source, rotated.output, rotated.operation, tempRoot, `ocr-rotation-${rotated.rotation}`);
        if (run("qpdf", ["--version"]).status === 0) run("qpdf", ["--check", rotated.output], { status: 0 });
      }
    }

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
