import assert from "node:assert/strict";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import sharp from "sharp";

import { PdfArtifact, PdfFile } from "../src/index.mjs";


const repoRoot = path.resolve(import.meta.dirname, "..");
const skillRoot = path.join(repoRoot, "skills", "pdf", "skills", "pdf");
const provider = path.join(skillRoot, "scripts", "ocrmypdf_provider.py");
const registry = path.join(skillRoot, "scripts", "pdf_provider.py");
const python = "python3";


function run(executable, args, options = {}) {
  const result = spawnSync(executable, args, {
    cwd: options.cwd || repoRoot,
    encoding: "utf8",
    env: { ...process.env, PYTHONDONTWRITEBYTECODE: "1", ...options.env },
    maxBuffer: 32 * 1024 * 1024,
  });
  if (options.status !== undefined) {
    assert.equal(
      result.status,
      options.status,
      `${executable} ${args.join(" ")}\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`,
    );
  }
  return result;
}


function jsonResult(result, stream = "stdout") {
  const value = result[stream]?.trim();
  assert.ok(value, `expected JSON on ${stream}\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
  return JSON.parse(value);
}


function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}


function commandPath(command) {
  const result = spawnSync("/usr/bin/env", ["which", command], { encoding: "utf8" });
  return result.status === 0 ? result.stdout.trim() : undefined;
}


async function writeControl(file, value = {}) {
  await fs.writeFile(file, JSON.stringify({
    ocrmypdfVersion: "17.8.1",
    tesseractVersion: "5.5.2",
    qpdfVersion: "12.3.2",
    popplerVersion: "26.05.0",
    languages: ["eng", "osd"],
    sidecarText: "OPEN CHESTNUT OCR TEST\nInvoice 2026 Amount 12345\n",
    extractedText: "OPEN CHESTNUT OCR TEST\nInvoice 2026 Amount 12345\n\f",
    ...value,
  }), "utf8");
}


const manifest = (await fs.readFile(path.join(skillRoot, "manifest.txt"), "utf8")).split(/\r?\n/).filter(Boolean);
assert.ok(manifest.includes("scripts/ocrmypdf_provider.py"));
assert.ok(manifest.includes("tasks/ocr.md"));
const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
assert.match(skillText, /ocrmypdf_provider\.py/);
assert.match(skillText, /not.*saniti[sz]er/is);
const taskText = await fs.readFile(path.join(skillRoot, "tasks", "ocr.md"), "utf8");
assert.match(taskText, /--expected-sha256/);
assert.match(taskText, /--input-trust/);
assert.match(taskText, /Poppler/);

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-ocrmypdf-provider-"));
try {
  const control = path.join(tempRoot, "control.json");
  const lastArgs = path.join(tempRoot, "last-ocr-args.json");
  const fakeOcr = path.join(tempRoot, "fake-ocrmypdf.mjs");
  const fakeTesseract = path.join(tempRoot, "fake-tesseract.mjs");
  const fakeQpdf = path.join(tempRoot, "fake-qpdf.mjs");
  const fakePdftotext = path.join(tempRoot, "fake-pdftotext.mjs");
  const nodeShebang = `#!${process.execPath}`;

  await fs.writeFile(fakeOcr, `${nodeShebang}
import fs from "node:fs";
import path from "node:path";
const root = path.dirname(process.argv[1]);
const control = JSON.parse(fs.readFileSync(path.join(root, "control.json"), "utf8"));
const args = process.argv.slice(2);
if (args.includes("--version")) { console.log(control.ocrmypdfVersion); process.exit(0); }
if (args.length === 1 && args[0] === "-v") process.exit(2);
fs.writeFileSync(path.join(root, "last-ocr-args.json"), JSON.stringify(args));
if (control.hang) Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0);
if (control.bigStdout) { process.stdout.write("x".repeat(1024 * 1024)); process.exit(0); }
if (control.exitCode) { console.error("fake OCR failure"); process.exit(control.exitCode); }
const sidecarIndex = args.indexOf("--sidecar");
const sidecar = args[sidecarIndex + 1];
const input = args.at(-2);
const output = args.at(-1);
const source = fs.readFileSync(input);
if (control.mutateSnapshot) fs.appendFileSync(input, "snapshot mutation");
const rewritten = control.retainPrefix
  ? Buffer.concat([source, Buffer.from("\\n% appended fake OCR revision\\n")])
  : Buffer.concat([Buffer.from("%PDF-1.7\\n% fake full rewrite\\n"), source.subarray(Math.min(9, source.length))]);
fs.writeFileSync(output, rewritten);
fs.writeFileSync(sidecar, control.sidecarText || "");
`, "utf8");

  await fs.writeFile(fakeTesseract, `${nodeShebang}
import fs from "node:fs";
import path from "node:path";
const root = path.dirname(process.argv[1]);
const control = JSON.parse(fs.readFileSync(path.join(root, "control.json"), "utf8"));
if (process.argv.includes("--version")) { console.log("tesseract " + control.tesseractVersion); process.exit(0); }
if (process.argv.includes("--list-langs")) {
  console.log("List of available languages (" + control.languages.length + "):");
  for (const language of control.languages) console.log(language);
  process.exit(0);
}
process.exit(2);
`, "utf8");

  await fs.writeFile(fakeQpdf, `${nodeShebang}
import fs from "node:fs";
import path from "node:path";
const root = path.dirname(process.argv[1]);
const control = JSON.parse(fs.readFileSync(path.join(root, "control.json"), "utf8"));
const args = process.argv.slice(2);
if (args.includes("--version")) { console.log("qpdf version " + control.qpdfVersion); process.exit(0); }
const input = args.at(-1);
if (args.includes("--check")) {
  console.log("checking " + input);
  console.log("PDF Version: 1.7");
  console.log(control.encrypted ? "File is encrypted" : "File is not encrypted");
  console.log("File is not linearized");
  console.log("No syntax or stream encoding errors found; the file may still contain");
  console.log("errors that qpdf cannot detect");
  process.exit(0);
}
if (args.includes("--json=2")) {
  const fields = [];
  const objects = {};
  if (control.tagged) objects["1 0 R"] = { value: { "/StructTreeRoot": "2 0 R", "/MarkInfo": { "/Marked": true } } };
  if (control.signed) {
    fields.push({ object: "7 0 R", fieldtype: "/Sig", fullname: "Signature1", value: "6 0 R" });
    objects["6 0 R"] = { value: { "/Type": "/Sig", "/ByteRange": [0, 1, 2, 3], "/TransformMethod": "/DocMDP" } };
  }
  if (control.formField) fields.push({ object: "8 0 R", fieldtype: "/Tx", fullname: "Name", value: "" });
  console.log(JSON.stringify({
    version: 2,
    parameters: {},
    pages: [{ object: "3 0 R", pageposfrom1: 1, contents: [], images: [], outlines: [], label: null }],
    outlines: [],
    acroform: { fields, hasacroform: fields.length > 0, needappearances: false },
    attachments: {},
    encrypt: {
      encrypted: Boolean(control.encrypted), ownerpasswordmatched: false, userpasswordmatched: false,
      parameters: { method: control.encrypted ? "AESv3" : "none", bits: control.encrypted ? 256 : 0 },
    },
    qpdf: [{ jsonversion: 2, pdfversion: "1.7" }, objects],
  }));
  process.exit(0);
}
process.exit(2);
`, "utf8");

  await fs.writeFile(fakePdftotext, `${nodeShebang}
import fs from "node:fs";
import path from "node:path";
const root = path.dirname(process.argv[1]);
const control = JSON.parse(fs.readFileSync(path.join(root, "control.json"), "utf8"));
if (process.argv.includes("-v")) { console.error("pdftotext version " + control.popplerVersion); process.exit(0); }
if (control.bigExtractedText) { process.stdout.write("x".repeat(1024 * 1024)); process.exit(0); }
process.stdout.write(control.extractedText || "");
`, "utf8");

  for (const executable of [fakeOcr, fakeTesseract, fakeQpdf, fakePdftotext]) await fs.chmod(executable, 0o755);
  const fakeEnv = {
    OPEN_OFFICE_PDF_OCRMYPDF: fakeOcr,
    OPEN_OFFICE_PDF_TESSERACT: fakeTesseract,
    OPEN_OFFICE_PDF_QPDF: fakeQpdf,
    OPEN_OFFICE_PDF_PDFTOTEXT: fakePdftotext,
  };
  await writeControl(control);

  const source = path.join(tempRoot, "source.pdf");
  const sourceBytes = Buffer.from("%PDF-1.7\n% bounded OCR provider fixture\n%%EOF\n", "ascii");
  const sourceHash = sha256(sourceBytes);
  await fs.writeFile(source, sourceBytes);

  const probe = jsonResult(run(python, [provider, "probe"], { env: fakeEnv, status: 0 }));
  assert.equal(probe.providerVersion, "17.8.1");
  assert.deepEqual(probe.modes, ["skip", "redo", "force"]);
  assert.equal(probe.providerIsSanitizer, false);
  assert.equal(probe.adapterSandboxEnforced, false);
  assert.ok(probe.languages.includes("eng"));

  const registryProbe = jsonResult(run(python, [registry, "check", "--provider", "ocrmypdf", "--require"], {
    env: fakeEnv,
    status: 0,
  }));
  assert.equal(registryProbe.providers[0].available, true);
  assert.equal(registryProbe.providers[0].integration, "shipped-thin-script-external-cli");
  assert.equal(registryProbe.providers[0].evidence.minimumVersion, "17.8.0");
  assert.equal(registryProbe.providers[0].evidence.maximumVersionExclusive, "17.9.0");

  const plannedOutput = path.join(tempRoot, "planned.pdf");
  const plan = jsonResult(run(python, [
    registry, "plan", "--task", "ocr", "--provider", "ocrmypdf", "--strategy", "rewrite",
    "--input", source, "--output", plannedOutput, "--require-provider",
  ], { env: fakeEnv, status: 0 }));
  assert.equal(plan.silentFallback, false);
  assert.equal(plan.providerProbe.available, true);

  const output = path.join(tempRoot, "ocr-output.pdf");
  const success = jsonResult(run(python, [
    provider, "ocr", source, output,
    "--expected-sha256", sourceHash,
    "--mode", "skip",
    "--language", "eng",
    "--input-trust", "trusted",
    "--require-text", "open chestnut   ocr test",
  ], { env: fakeEnv, status: 0 }));
  assert.equal(success.schema, "open-office-artifact-tool.ocrmypdf-ocr.v1");
  assert.equal(success.savePolicy, "rewrite");
  assert.equal(success.sourceProtected, true);
  assert.equal(success.sourcePrefixRetained, false);
  assert.equal(success.transaction.atomicDistinctOutput, true);
  assert.equal(success.transaction.privateSidecarRetained, false);
  assert.equal(success.provider.components.rasterizer, "pypdfium");
  assert.equal(success.provider.components.pdfRenderer, "fpdf2");
  assert.equal(success.textValidation.requiredTextMatched, true);
  assert.deepEqual(await fs.readFile(source), sourceBytes);
  const invoked = JSON.parse(await fs.readFile(lastArgs, "utf8"));
  for (const fixed of ["--output-type", "pdf", "--optimize", "0", "--jobs", "1", "--ocr-engine", "tesseract", "--rasterizer", "pypdfium", "--pdf-renderer", "fpdf2", "--no-overwrite"]) {
    assert.ok(invoked.includes(fixed), `missing fixed OCRmyPDF argument ${fixed}`);
  }
  assert.ok(!invoked.includes("--plugin"));
  assert.ok(!invoked.includes("--pages"), "the first shipped slice must remain complete-document only");

  const outputBytesBeforeCollision = await fs.readFile(output);
  const collision = run(python, [
    provider, "ocr", source, output, "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(collision, "stderr").error, /output already exists/);
  assert.deepEqual(await fs.readFile(output), outputBytesBeforeCollision, "existing output must not be replaced");

  const staleOutput = path.join(tempRoot, "stale.pdf");
  const stale = run(python, [
    provider, "ocr", source, staleOutput, "--expected-sha256", "0".repeat(64),
    "--mode", "skip", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(stale, "stderr").error, /source SHA-256 mismatch/);
  await assert.rejects(fs.access(staleOutput));

  const missingTrust = run(python, [
    provider, "ocr", source, path.join(tempRoot, "missing-trust.pdf"), "--expected-sha256", sourceHash, "--mode", "skip",
  ], { env: fakeEnv, status: 2 });
  assert.match(missingTrust.stderr, /--input-trust/);

  await writeControl(control, { ocrmypdfVersion: "17.7.0" });
  const oldOcr = run(python, [provider, "probe"], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(oldOcr, "stderr").error, />= 17\.8\.0 and < 17\.9\.0/);
  await writeControl(control, { tesseractVersion: "4.1.3" });
  const oldTesseract = run(python, [provider, "probe"], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(oldTesseract, "stderr").error, /Tesseract >= 5\.0\.0/);
  await writeControl(control);

  const unavailableLanguage = run(python, [
    provider, "ocr", source, path.join(tempRoot, "missing-language.pdf"), "--expected-sha256", sourceHash,
    "--mode", "skip", "--language", "deu", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(unavailableLanguage, "stderr").error, /language 'deu' is unavailable/);

  const redoWithoutAcknowledgement = run(python, [
    provider, "ocr", source, path.join(tempRoot, "redo.pdf"), "--expected-sha256", sourceHash,
    "--mode", "redo", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(redoWithoutAcknowledgement, "stderr").error, /allow-structure-loss/);
  const forceWithoutAcknowledgement = run(python, [
    provider, "ocr", source, path.join(tempRoot, "force.pdf"), "--expected-sha256", sourceHash,
    "--mode", "force", "--allow-structure-loss", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(forceWithoutAcknowledgement, "stderr").error, /allow-rasterize-all/);

  await writeControl(control, { tagged: true });
  const taggedRejected = run(python, [
    provider, "ocr", source, path.join(tempRoot, "tagged-rejected.pdf"), "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(taggedRejected, "stderr").error, /Tagged PDF input requires --allow-structure-loss/);
  const taggedOutput = path.join(tempRoot, "tagged-accepted.pdf");
  const taggedAccepted = jsonResult(run(python, [
    provider, "ocr", source, taggedOutput, "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted", "--allow-structure-loss",
  ], { env: fakeEnv, status: 0 }));
  assert.equal(taggedAccepted.structureBefore.tagged, true);
  assert.equal(taggedAccepted.fidelityPolicy.structureLossAcknowledged, true);

  await writeControl(control, { encrypted: true });
  const encrypted = run(python, [
    provider, "ocr", source, path.join(tempRoot, "encrypted.pdf"), "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(encrypted, "stderr").error, /encrypted input is unsupported/);
  await writeControl(control, { signed: true });
  const signed = run(python, [
    provider, "ocr", source, path.join(tempRoot, "signed.pdf"), "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(signed, "stderr").error, /invalidate-signatures/);

  await writeControl(control, { formField: true });
  const interactiveForce = run(python, [
    provider, "ocr", source, path.join(tempRoot, "interactive-force.pdf"), "--expected-sha256", sourceHash,
    "--mode", "force", "--allow-structure-loss", "--allow-rasterize-all", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(interactiveForce, "stderr").error, /allow-interactive-flattening/);

  await writeControl(control, { retainPrefix: true });
  const prefix = run(python, [
    provider, "ocr", source, path.join(tempRoot, "prefix.pdf"), "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(prefix, "stderr").error, /complete source byte prefix/);

  await writeControl(control, { mutateSnapshot: true });
  const mutatedSnapshot = run(python, [
    provider, "ocr", source, path.join(tempRoot, "mutated-snapshot.pdf"), "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(mutatedSnapshot, "stderr").error, /permission denied|changed the private read-only source snapshot/i);
  assert.deepEqual(await fs.readFile(source), sourceBytes);

  await writeControl(control, { extractedText: "", sidecarText: "" });
  const empty = run(python, [
    provider, "ocr", source, path.join(tempRoot, "empty.pdf"), "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(empty, "stderr").error, /no extractable non-whitespace text/);
  const allowedEmptyOutput = path.join(tempRoot, "allowed-empty.pdf");
  const allowedEmpty = jsonResult(run(python, [
    provider, "ocr", source, allowedEmptyOutput, "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted", "--allow-empty-text",
  ], { env: fakeEnv, status: 0 }));
  assert.equal(allowedEmpty.textValidation.finalExtraction.nonWhitespaceCharacters, 0);

  await writeControl(control, { extractedText: "different text\f" });
  const missingRequiredText = run(python, [
    provider, "ocr", source, path.join(tempRoot, "missing-required.pdf"), "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted", "--require-text", "must exist",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(missingRequiredText, "stderr").error, /failed required text gates/);

  await writeControl(control, { bigStdout: true });
  const oversizedDiagnostics = run(python, [
    provider, "ocr", source, path.join(tempRoot, "big-stdout.pdf"), "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted", "--max-stdout-bytes", "1024",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(oversizedDiagnostics, "stderr").error, /stdout exceeded the 1024 byte budget/);

  await writeControl(control, { hang: true });
  const timedOut = run(python, [
    provider, "ocr", source, path.join(tempRoot, "timeout.pdf"), "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted", "--timeout-seconds", "1",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(timedOut, "stderr").error, /timed out after 1 seconds/);

  await writeControl(control);
  const raisedBudget = run(python, [
    provider, "ocr", source, path.join(tempRoot, "raised-budget.pdf"), "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted", "--max-input-bytes", String(512 * 1024 * 1024 + 1),
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(raisedBudget, "stderr").error, /cannot exceed the hard maximum/);

  const symlinkOutput = path.join(tempRoot, "symlink.pdf");
  const symlinkTarget = path.join(tempRoot, "symlink-target.pdf");
  await fs.symlink(symlinkTarget, symlinkOutput);
  const symlinkRejected = run(python, [
    provider, "ocr", source, symlinkOutput, "--expected-sha256", sourceHash,
    "--mode", "skip", "--input-trust", "trusted",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(symlinkRejected, "stderr").error, /symbolic link.*will not be followed/);
  await assert.rejects(fs.access(symlinkTarget));

  const realProvider = process.env.OPEN_OFFICE_PDF_OCRMYPDF_TEST;
  if (realProvider) {
    const realPdftotext = process.env.OPEN_OFFICE_PDF_PDFTOTEXT_TEST || commandPath("pdftotext");
    const pdftoppm = process.env.OPEN_OFFICE_PDF_PDFTOPPM_TEST || commandPath("pdftoppm");
    assert.ok(realPdftotext, "real OCR test requires Poppler pdftotext");
    assert.ok(pdftoppm, "real OCR test requires Poppler pdftoppm");
    const realEnv = {
      OPEN_OFFICE_PDF_OCRMYPDF: realProvider,
      OPEN_OFFICE_PDF_PDFTOTEXT: realPdftotext,
      ...(process.env.OPEN_OFFICE_PDF_TESSERACT_TEST ? { OPEN_OFFICE_PDF_TESSERACT: process.env.OPEN_OFFICE_PDF_TESSERACT_TEST } : {}),
    };
    const realProbe = jsonResult(run(python, [provider, "probe"], { env: realEnv, status: 0 }));
    assert.equal(realProbe.providerVersion, "17.8.1");
    assert.ok(realProbe.languages.includes("eng"));

    const svg = Buffer.from(
      '<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="600">' +
      '<rect width="1200" height="600" fill="white"/>' +
      '<text x="80" y="250" font-family="Arial,sans-serif" font-size="72" font-weight="700">OPEN CHESTNUT OCR TEST</text>' +
      '<text x="80" y="360" font-family="Arial,sans-serif" font-size="56">Invoice 2026 Amount 12345</text>' +
      "</svg>",
    );
    const png = await sharp(svg).png().toBuffer();
    const scan = PdfArtifact.create({ pages: [{ text: "", width: 612, height: 792 }] });
    scan.addImage({
      name: "scanned-page",
      alt: "Synthetic OCR test scan",
      dataUrl: `data:image/png;base64,${png.toString("base64")}`,
      bbox: [36, 200, 540, 270],
    });
    const exported = await PdfFile.exportPdf(scan, { tagged: false });
    const realSource = path.join(tempRoot, "real-scan.pdf");
    const realOutput = path.join(tempRoot, "real-searchable.pdf");
    const realSourceBytes = Buffer.from(exported.bytes);
    await fs.writeFile(realSource, realSourceBytes);
    const sourceText = run(realPdftotext, [realSource, "-"], { status: 0 }).stdout;
    assert.equal(sourceText.trim(), "", "synthetic scan must not already contain extractable text");

    const realReport = jsonResult(run(python, [
      provider, "ocr", realSource, realOutput,
      "--expected-sha256", sha256(realSourceBytes),
      "--mode", "skip",
      "--language", "eng",
      "--input-trust", "trusted",
      "--require-text", "OPEN CHESTNUT OCR TEST",
    ], { env: realEnv, status: 0 }));
    assert.equal(realReport.operation.scope, "complete-document");
    assert.equal(realReport.structureBefore.pageCount, 1);
    assert.equal(realReport.structureAfter.pageCount, 1);
    assert.equal(realReport.sourceProtected, true);
    assert.deepEqual(await fs.readFile(realSource), realSourceBytes);
    const finalText = run(realPdftotext, [realOutput, "-"], { status: 0 }).stdout;
    assert.match(finalText, /OPEN CHESTNUT OCR TEST/);
    assert.match(finalText, /Invoice 2026 Amount 12345/);
    const imported = await PdfFile.importPdf(await fs.readFile(realOutput), { preferParser: true });
    assert.match(imported.extractText(), /OPEN CHESTNUT OCR TEST/);
    const nativeInspect = await PdfFile.inspectPdf(await fs.readFile(realOutput));
    assert.equal(nativeInspect.summary.pages, 1);
    assert.equal(nativeInspect.summary.sourceSha256, realReport.output.sha256);

    const sourceRender = path.join(tempRoot, "real-source-render");
    const outputRender = path.join(tempRoot, "real-output-render");
    run(pdftoppm, ["-png", "-singlefile", "-r", "96", realSource, sourceRender], { status: 0 });
    run(pdftoppm, ["-png", "-singlefile", "-r", "96", realOutput, outputRender], { status: 0 });
    const sourcePixels = await sharp(`${sourceRender}.png`).raw().toBuffer({ resolveWithObject: true });
    const outputPixels = await sharp(`${outputRender}.png`).raw().toBuffer({ resolveWithObject: true });
    assert.deepEqual(
      { width: outputPixels.info.width, height: outputPixels.info.height, channels: outputPixels.info.channels },
      { width: sourcePixels.info.width, height: sourcePixels.info.height, channels: sourcePixels.info.channels },
    );
    assert.deepEqual(outputPixels.data, sourcePixels.data, "skip/pdf/O0 OCR layer must preserve this scan's Poppler rendering exactly");
  }
} finally {
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log(
  process.env.OPEN_OFFICE_PDF_OCRMYPDF_TEST
    ? "OCRmyPDF provider smoke ok"
    : "OCRmyPDF provider smoke ok (real provider skipped: set OPEN_OFFICE_PDF_OCRMYPDF_TEST)",
);
