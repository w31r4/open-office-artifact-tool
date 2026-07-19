import assert from "node:assert/strict";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { PdfArtifact, PdfFile } from "../src/index.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const skillRoot = path.join(repoRoot, "skills", "pdf", "skills", "pdf");
const qpdfProvider = path.join(skillRoot, "scripts", "qpdf_provider.py");
const providerRegistry = path.join(skillRoot, "scripts", "pdf_provider.py");
const python = "python3";

function run(executable, args, options = {}) {
  const result = spawnSync(executable, args, {
    cwd: options.cwd || repoRoot,
    encoding: "utf8",
    env: { ...process.env, PYTHONDONTWRITEBYTECODE: "1", ...options.env },
    maxBuffer: 24 * 1024 * 1024,
  });
  if (options.status !== undefined) {
    assert.equal(result.status, options.status, `${executable} ${args.join(" ")}\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
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

function buildPdfFixture(objects) {
  let pdf = "%PDF-1.7\n%\xE2\xE3\xCF\xD3\n";
  const offsets = [0];
  for (let index = 0; index < objects.length; index += 1) {
    offsets.push(Buffer.byteLength(pdf, "latin1"));
    pdf += `${index + 1} 0 obj\n${objects[index]}\nendobj\n`;
  }
  const xref = Buffer.byteLength(pdf, "latin1");
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets.slice(1)) pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(pdf, "latin1");
}

function buildSignedFixture() {
  return buildPdfFixture([
    "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R /Perms << /DocMDP 6 0 R >> >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents 4 0 R /Annots [7 0 R] >>",
    "<< /Length 0 >>\nstream\n\nendstream",
    "<< /Fields [7 0 R] /SigFlags 3 >>",
    "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 0 0 0] /Contents <00> /Reference [<< /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /P 2 /V /1.2 >> >>] >>",
    "<< /Type /Annot /Subtype /Widget /FT /Sig /T (Signature1) /V 6 0 R /Rect [0 0 0 0] /P 3 0 R >>",
  ]);
}

function buildAttachmentFixture() {
  return buildPdfFixture([
    "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles 5 0 R >> >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents 4 0 R >>",
    "<< /Length 0 >>\nstream\n\nendstream",
    "<< /Names [(note.txt) 6 0 R] >>",
    "<< /Type /Filespec /F (note.txt) /UF (note.txt) /EF << /F 7 0 R /UF 7 0 R >> >>",
    "<< /Type /EmbeddedFile /Subtype /text#2Fplain /Length 5 >>\nstream\nhello\nendstream",
  ]);
}

function commandAvailable(command) {
  return spawnSync(command, ["--version"], { stdio: "ignore" }).status === 0;
}

const manifest = (await fs.readFile(path.join(skillRoot, "manifest.txt"), "utf8")).split(/\r?\n/).filter(Boolean);
assert.ok(manifest.includes("scripts/qpdf_provider.py"));
assert.ok(manifest.includes("tasks/repair_linearize.md"));
const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
assert.match(skillText, /qpdf_provider\.py/);
assert.match(skillText, /source SHA-256.*repair.*linearize/is);

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-qpdf-provider-"));
try {
  const dummyInput = path.join(tempRoot, "dummy.pdf");
  const dummyBytes = Buffer.from("%PDF-1.7\nfake provider fixture\n%%EOF\n", "ascii");
  await fs.writeFile(dummyInput, dummyBytes);
  const fakeQpdf = path.join(tempRoot, "fake-qpdf.mjs");
  await fs.writeFile(fakeQpdf, [
    "#!/usr/bin/env node",
    "import fs from 'node:fs';",
    "const args = process.argv.slice(2);",
    "if (args.includes('--version')) { console.log(`qpdf version ${process.env.FAKE_QPDF_VERSION || '12.3.2'}`); process.exit(0); }",
    "const input = args.at(-1);",
    "if (args.includes('--check')) {",
    "  if (process.env.FAKE_QPDF_HANG_CHECK === '1') { setInterval(() => {}, 1000); } else {",
    "  console.log(`checking ${input}`);",
    "  console.log('PDF Version: 1.7');",
    "  console.log('File is not encrypted');",
    "  console.log('File is not linearized');",
    "  console.log('No syntax or stream encoding errors found; the file may still contain');",
    "  console.log('errors that qpdf cannot detect');",
    "  process.exit(0);",
    "  }",
    "} else if (args.includes('--json=2')) {",
    "  if (process.env.FAKE_QPDF_BAD_JSON === '1') { console.log('{}'); process.exit(0); }",
    "  console.log(JSON.stringify({version:2,parameters:{},pages:[{object:'1 0 R',pageposfrom1:1,contents:[],images:[],outlines:[],label:null}],outlines:[],acroform:{fields:[],hasacroform:false,needappearances:false},attachments:{},encrypt:{encrypted:false,ownerpasswordmatched:false,userpasswordmatched:false,parameters:{method:'none',bits:0}},qpdf:[{jsonversion:2,pdfversion:'1.7'},{}]}));",
    "  process.exit(0);",
    "} else {",
    "const positional = args.filter((arg) => !arg.startsWith('--'));",
    "if (positional.length === 2) { fs.copyFileSync(positional[0], positional[1]); process.exit(0); }",
    "console.error('unsupported fake qpdf invocation', args);",
    "process.exit(2);",
    "}",
  ].join("\n"), "utf8");
  await fs.chmod(fakeQpdf, 0o755);
  const fakeEnv = { OPEN_OFFICE_PDF_QPDF: fakeQpdf };

  const fakeProbe = jsonResult(run(python, [qpdfProvider, "probe"], { env: fakeEnv, status: 0 }));
  assert.equal(fakeProbe.provider, "qpdf");
  assert.equal(fakeProbe.integration, "shipped-thin-script-external-cli");
  assert.equal(fakeProbe.silentFallback, false);
  const oldRegistry = run(python, [providerRegistry, "check", "--provider", "qpdf", "--require"], {
    env: { ...fakeEnv, FAKE_QPDF_VERSION: "10.6.3" },
    status: 2,
  });
  assert.equal(jsonResult(oldRegistry).providers[0].evidence.minimumMajor, 11);
  const oldAdapter = run(python, [qpdfProvider, "probe"], {
    env: { ...fakeEnv, FAKE_QPDF_VERSION: "10.6.3" },
    status: 2,
  });
  assert.match(oldAdapter.stderr, /qpdf 11 or newer/);
  const fakeInspect = jsonResult(run(python, [qpdfProvider, "inspect", dummyInput], { env: fakeEnv, status: 0 }));
  assert.equal(fakeInspect.schema, "open-office-artifact-tool.qpdf-inspect.v1");
  assert.equal(fakeInspect.structure.pageCount, 1);
  assert.equal(fakeInspect.source.sha256, sha256(dummyBytes));
  const fakeOutput = path.join(tempRoot, "fake-repaired.pdf");
  const fakeRewrite = jsonResult(run(python, [
    qpdfProvider, "rewrite", dummyInput, fakeOutput,
    "--mode", "repair", "--expected-sha256", fakeInspect.source.sha256,
  ], { env: fakeEnv, status: 0 }));
  assert.equal(fakeRewrite.savePolicy, "rewrite");
  assert.equal(fakeRewrite.sourceProtected, true);
  assert.equal(fakeRewrite.transaction.atomicDistinctOutput, true);
  assert.deepEqual(await fs.readFile(fakeOutput), dummyBytes);

  const providerProbe = jsonResult(run(python, [providerRegistry, "check", "--provider", "qpdf", "--require"], { env: fakeEnv, status: 0 }));
  assert.equal(providerProbe.providers[0].available, true);
  assert.equal(providerProbe.providers[0].integration, "shipped-thin-script-external-cli");
  const plannedOutput = path.join(tempRoot, "planned-output.pdf");
  const repairPlan = jsonResult(run(python, [
    providerRegistry, "plan", "--task", "repair", "--provider", "qpdf", "--strategy", "rewrite",
    "--input", dummyInput, "--output", plannedOutput, "--require-provider",
  ], { env: fakeEnv, status: 0 }));
  assert.equal(repairPlan.providerProbe.available, true);
  const falseClean = run(python, [
    providerRegistry, "plan", "--task", "structure-clean", "--provider", "qpdf", "--strategy", "rewrite",
    "--input", dummyInput, "--output", plannedOutput, "--invalidate-signatures",
  ], { env: fakeEnv, status: 2 });
  assert.match(falseClean.stderr, /cannot perform task/);

  const staleOutput = path.join(tempRoot, "stale-output.pdf");
  const stale = run(python, [
    qpdfProvider, "rewrite", dummyInput, staleOutput,
    "--mode", "repair", "--expected-sha256", "0".repeat(64),
  ], { env: fakeEnv, status: 2 });
  assert.match(stale.stderr, /source SHA-256 mismatch/);
  await assert.rejects(fs.access(staleOutput));
  const overBudget = run(python, [
    qpdfProvider, "inspect", dummyInput, "--max-json-bytes", "64",
  ], { env: fakeEnv, status: 2 });
  assert.match(overBudget.stderr, /stdout exceeded the 64 byte budget/);
  const malformedJson = run(python, [qpdfProvider, "inspect", dummyInput], {
    env: { ...fakeEnv, FAKE_QPDF_BAD_JSON: "1" },
    status: 2,
  });
  assert.match(malformedJson.stderr, /JSON v2 object table is missing or malformed/);
  const timedOut = run(python, [qpdfProvider, "inspect", dummyInput, "--timeout-seconds", "1"], {
    env: { ...fakeEnv, FAKE_QPDF_HANG_CHECK: "1" },
    status: 2,
  });
  assert.match(timedOut.stderr, /timed out after 1 seconds/);
  const danglingTarget = path.join(tempRoot, "dangling-target.pdf");
  const symlinkOutput = path.join(tempRoot, "symlink-output.pdf");
  await fs.symlink(danglingTarget, symlinkOutput);
  const symlinkRefused = run(python, [
    qpdfProvider, "rewrite", dummyInput, symlinkOutput,
    "--mode", "repair", "--expected-sha256", fakeInspect.source.sha256,
  ], { env: fakeEnv, status: 2 });
  assert.match(symlinkRefused.stderr, /symbolic link.*will not be followed/);
  await assert.rejects(fs.access(danglingTarget));
  const missingProvider = run(python, [qpdfProvider, "probe"], {
    env: { OPEN_OFFICE_PDF_QPDF: path.join(tempRoot, "missing-qpdf") },
    status: 2,
  });
  assert.equal(jsonResult(missingProvider, "stderr").silentFallback, false);

  if (commandAvailable("qpdf")) {
    const artifact = PdfArtifact.create({ text: "qpdf source-bound provider fixture" });
    artifact.addLink({ text: "Structural review", url: "https://example.com/qpdf", bbox: [72, 120, 120, 16] });
    const exported = await PdfFile.exportPdf(artifact);
    const source = path.join(tempRoot, "source.pdf");
    const sourceBytes = Buffer.from(exported.bytes);
    await fs.writeFile(source, sourceBytes);
    const realInspect = jsonResult(run(python, [qpdfProvider, "inspect", source], { env: { OPEN_OFFICE_PDF_QPDF: "" }, status: 0 }));
    assert.equal(realInspect.check.status, "clean");
    assert.equal(realInspect.structure.pageCount, 1);
    assert.equal(realInspect.signaturePolicy.hasSignatureEvidence, false);

    const repaired = path.join(tempRoot, "repaired.pdf");
    const repairedResult = jsonResult(run(python, [
      qpdfProvider, "rewrite", source, repaired,
      "--mode", "repair", "--expected-sha256", realInspect.source.sha256,
    ], { env: { OPEN_OFFICE_PDF_QPDF: "" }, status: 0 }));
    assert.equal(repairedResult.checkAfter.status, "clean");
    assert.equal(repairedResult.structureAfter.pageCount, 1);
    assert.deepEqual(await fs.readFile(source), sourceBytes, "qpdf rewrite must not mutate source bytes");
    assert.equal((await PdfFile.inspectPdf(await fs.readFile(repaired))).summary.pages, 1);

    const linearized = path.join(tempRoot, "linearized.pdf");
    const linearizedResult = jsonResult(run(python, [
      qpdfProvider, "rewrite", source, linearized,
      "--mode", "linearize", "--expected-sha256", realInspect.source.sha256,
    ], { env: { OPEN_OFFICE_PDF_QPDF: "" }, status: 0 }));
    assert.equal(linearizedResult.structureAfter.linearized, true);
    assert.equal(linearizedResult.structureAfter.pageCount, 1);

    const broken = path.join(tempRoot, "broken-xref.pdf");
    const brokenBytes = Buffer.from(sourceBytes.toString("latin1").replace(/startxref\n\d+\n%%EOF\s*$/, "startxref\n0\n%%EOF\n"), "latin1");
    assert.notDeepEqual(brokenBytes, sourceBytes);
    await fs.writeFile(broken, brokenBytes);
    const brokenInspect = jsonResult(run(python, [qpdfProvider, "inspect", broken], { env: { OPEN_OFFICE_PDF_QPDF: "" }, status: 0 }));
    assert.equal(brokenInspect.check.status, "warnings");
    assert.ok(brokenInspect.check.lines.some((line) => /reconstruct cross-reference/i.test(line)));
    const recovered = path.join(tempRoot, "recovered.pdf");
    const recoveredResult = jsonResult(run(python, [
      qpdfProvider, "rewrite", broken, recovered,
      "--mode", "repair", "--expected-sha256", brokenInspect.source.sha256,
    ], { env: { OPEN_OFFICE_PDF_QPDF: "" }, status: 0 }));
    assert.equal(recoveredResult.checkBefore.status, "warnings");
    assert.equal(recoveredResult.checkAfter.status, "clean");
    assert.equal(recoveredResult.structureAfter.pageCount, 1);

    const attachmentSource = path.join(tempRoot, "attachment.pdf");
    const attachmentBytes = buildAttachmentFixture();
    await fs.writeFile(attachmentSource, attachmentBytes);
    const attachmentInspect = jsonResult(run(python, [qpdfProvider, "inspect", attachmentSource], { env: { OPEN_OFFICE_PDF_QPDF: "" }, status: 0 }));
    assert.equal(attachmentInspect.structure.attachmentCount, 1);
    const attachmentOutput = path.join(tempRoot, "attachment-rewritten.pdf");
    const attachmentRewrite = jsonResult(run(python, [
      qpdfProvider, "rewrite", attachmentSource, attachmentOutput,
      "--mode", "repair", "--expected-sha256", attachmentInspect.source.sha256,
    ], { env: { OPEN_OFFICE_PDF_QPDF: "" }, status: 0 }));
    assert.equal(attachmentRewrite.structureBefore.attachmentCount, 1);
    assert.equal(attachmentRewrite.structureAfter.attachmentCount, 1);
    assert.deepEqual(await fs.readFile(attachmentSource), attachmentBytes);

    const signed = path.join(tempRoot, "signed.pdf");
    const signedBytes = buildSignedFixture();
    await fs.writeFile(signed, signedBytes);
    const signedInspect = jsonResult(run(python, [qpdfProvider, "inspect", signed], { env: { OPEN_OFFICE_PDF_QPDF: "" }, status: 0 }));
    assert.equal(signedInspect.signaturePolicy.hasSignatureFields, true);
    assert.equal(signedInspect.signaturePolicy.hasByteRange, true);
    assert.equal(signedInspect.signaturePolicy.hasDocMDP, true);
    const signedRefused = path.join(tempRoot, "signed-refused.pdf");
    const refused = run(python, [
      qpdfProvider, "rewrite", signed, signedRefused,
      "--mode", "repair", "--expected-sha256", signedInspect.source.sha256,
    ], { env: { OPEN_OFFICE_PDF_QPDF: "" }, status: 2 });
    assert.match(refused.stderr, /invalidate-signatures.*pyHanko\/DocMDP/);
    await assert.rejects(fs.access(signedRefused));
    const invalidated = path.join(tempRoot, "signed-invalidated.pdf");
    const invalidatedResult = jsonResult(run(python, [
      qpdfProvider, "rewrite", signed, invalidated,
      "--mode", "repair", "--expected-sha256", signedInspect.source.sha256,
      "--invalidate-signatures",
    ], { env: { OPEN_OFFICE_PDF_QPDF: "" }, status: 0 }));
    assert.equal(invalidatedResult.signatureInvalidated, true);
    assert.equal(invalidatedResult.signaturePolicyAfter.hasSignatureEvidence, true);
    assert.equal(invalidatedResult.signaturePolicyAfter.trust, "unknown");
    assert.deepEqual(await fs.readFile(signed), signedBytes);

    if (commandAvailable("pdftoppm")) {
      for (const [input, prefix] of [[source, "source"], [repaired, "repaired"], [linearized, "linearized"]]) {
        run("pdftoppm", ["-png", "-singlefile", "-r", "96", input, path.join(tempRoot, prefix)], { status: 0 });
      }
      const sourcePng = await fs.readFile(path.join(tempRoot, "source.png"));
      assert.deepEqual(await fs.readFile(path.join(tempRoot, "repaired.png")), sourcePng);
      assert.deepEqual(await fs.readFile(path.join(tempRoot, "linearized.png")), sourcePng);
    }
  }
} finally {
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log("qpdf provider smoke ok");
