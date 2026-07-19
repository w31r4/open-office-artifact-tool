import assert from "node:assert/strict";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { PdfFile } from "../src/index.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const skillRoot = path.join(repoRoot, "skills", "pdf", "skills", "pdf");
const provider = path.join(skillRoot, "scripts", "pikepdf_provider.py");
const providerRegistry = path.join(skillRoot, "scripts", "pdf_provider.py");
const qpdfProvider = path.join(skillRoot, "scripts", "qpdf_provider.py");
const configuredPython = process.env.OPEN_OFFICE_PDF_PIKEPDF_TEST_PYTHON || "python3";

function run(executable, args, options = {}) {
  const result = spawnSync(executable, args, {
    cwd: options.cwd || repoRoot,
    encoding: "utf8",
    env: {
      ...process.env,
      PYTHONDONTWRITEBYTECODE: "1",
      ...options.env,
    },
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

function streamObject(dictionary, data) {
  const bytes = Buffer.isBuffer(data) ? data : Buffer.from(data, "latin1");
  return Buffer.concat([
    Buffer.from(`<< ${dictionary} /Length ${bytes.length} >>\nstream\n`, "latin1"),
    bytes,
    Buffer.from("\nendstream", "latin1"),
  ]);
}

function buildPdfFixture(objects, trailerExtras = "") {
  const chunks = [Buffer.from("%PDF-1.7\n%\xE2\xE3\xCF\xD3\n", "latin1")];
  const offsets = [0];
  let length = chunks[0].length;
  for (let index = 0; index < objects.length; index += 1) {
    offsets.push(length);
    const body = Buffer.isBuffer(objects[index]) ? objects[index] : Buffer.from(objects[index], "latin1");
    const chunk = Buffer.concat([
      Buffer.from(`${index + 1} 0 obj\n`, "latin1"),
      body,
      Buffer.from("\nendobj\n", "latin1"),
    ]);
    chunks.push(chunk);
    length += chunk.length;
  }
  const xref = length;
  const lines = [
    `xref\n0 ${objects.length + 1}\n`,
    "0000000000 65535 f \n",
    ...offsets.slice(1).map((offset) => `${String(offset).padStart(10, "0")} 00000 n \n`),
    `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R ${trailerExtras} >>\nstartxref\n${xref}\n%%EOF\n`,
  ];
  chunks.push(Buffer.from(lines.join(""), "latin1"));
  return Buffer.concat(chunks);
}

function buildActiveContentFixture() {
  const content = "BT /F1 24 Tf 72 700 Td (Visible pikepdf fixture) Tj ET";
  const metadata = "<?xpacket begin=''?><x:xmpmeta xmlns:x='adobe:ns:meta/'><rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'><rdf:Description rdf:about='' xmlns:dc='http://purl.org/dc/elements/1.1/'><dc:title><rdf:Alt><rdf:li xml:lang='x-default'>Keep Metadata Canary</rdf:li></rdf:Alt></dc:title></rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end='w'?>";
  return buildPdfFixture([
    "<< /Type /Catalog /Pages 2 0 R /OpenAction 6 0 R /AA << /WC 7 0 R >> /Names << /JavaScript 8 0 R /EmbeddedFiles 9 0 R /Renditions 10 0 R >> /Collection << /Type /Collection >> /SpiderInfo << /V 1 >> /PieceInfo << /SearchIndex << /Canary (stale-index-canary) >> /VendorPrivate << /Canary (private-data-canary) >> >> /AcroForm 11 0 R /Metadata 16 0 R /Outlines 17 0 R /AF [18 0 R] /MarkInfo << /Marked true >> /StructTreeRoot << /Type /StructTreeRoot /K [] >> >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /Annots [12 0 R 13 0 R 14 0 R 24 0 R] /Thumb 15 0 R /PieceInfo << /PagePrivate << /Canary (page-private-canary) >> >> /AA << /O 6 0 R >> >>",
    streamObject("", content),
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    "<< /S /JavaScript /JS (app.alert('javascript-canary')) >>",
    "<< /S /Launch /F (launch-canary.exe) >>",
    "<< /Names [(DocScript) 6 0 R] >>",
    "<< /Names [(payload.txt) 18 0 R] >>",
    "<< /Names [(MovieOne) 19 0 R] >>",
    "<< /Fields [12 0 R] /XFA 20 0 R /NeedAppearances true >>",
    "<< /Type /Annot /Subtype /Widget /FT /Tx /T (KeepForm) /V (keep-form-value-canary) /Rect [0 0 0 0] /AA << /K 6 0 R >> >>",
    "<< /Type /Annot /Subtype /Link /Rect [0 0 0 0] /A << /S /URI /URI (https://external-canary.invalid/) >> >>",
    "<< /Type /Annot /Subtype /RichMedia /Rect [0 0 0 0] /RichMediaContent 21 0 R /RichMediaSettings << /Activation << /Condition /PO >> >> >>",
    streamObject("/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8", Buffer.from([0x80])),
    streamObject("/Type /Metadata /Subtype /XML", metadata),
    "<< /Type /Outlines /First 22 0 R /Last 22 0 R /Count 1 >>",
    "<< /Type /Filespec /F (payload.txt) /UF (payload.txt) /EF << /F 23 0 R /UF 23 0 R >> >>",
    "<< /S /Rendition /R << /C 21 0 R >> >>",
    streamObject("", "<xfa>keep-xfa-canary</xfa>"),
    "<< /Type /RichMediaContent /Assets << /Names [(media.bin) 18 0 R] >> /Canary (rich-media-canary) >>",
    "<< /Title (Remote outline) /Parent 17 0 R /A << /S /GoToR /F (remote-canary.pdf) /D [0 /Fit] >> >>",
    streamObject("/Type /EmbeddedFile /Subtype /text#2Fplain", "payload-secret-canary"),
    "<< /Type /Annot /Subtype /FileAttachment /Rect [0 0 0 0] /FS 18 0 R >>",
    "<< /Title (Keep Info Title) /Author (Keep Author Canary) /Subject (metadata-must-remain) >>",
  ], "/Info 25 0 R");
}

function available(executable, args = ["--version"]) {
  return spawnSync(executable, args, { stdio: "ignore" }).status === 0;
}

function pikepdfOracle(input, env) {
  const code = [
    "import json,pikepdf,sys",
    "canaries=['javascript-canary','launch-canary.exe','remote-canary.pdf','rich-media-canary','payload-secret-canary','private-data-canary','page-private-canary','stale-index-canary','Visible pikepdf fixture','Keep Metadata Canary','keep-form-value-canary','keep-xfa-canary']",
    "with pikepdf.open(sys.argv[1], attempt_recovery=False, inherit_page_attributes=False) as pdf:",
    "  corpus=[]",
    "  for obj in pdf.objects:",
    "    corpus.append(str(obj))",
    "    if getattr(obj,'_type_code',None)==pikepdf.ObjectType.stream:",
    "      try: corpus.append(obj.read_bytes().decode('latin1','replace'))",
    "      except Exception: pass",
    "  text='\\n'.join(corpus)",
    "  root=pdf.Root",
    "  page=pdf.pages[0].obj",
    "  form=root.get('/AcroForm')",
    "  fields=list(form.get('/Fields',[])) if form is not None else []",
    "  xfa=form.get('/XFA') if form is not None else None",
    "  attachments={name:pdf.attachments[name].get_file().read_bytes().decode('latin1','replace') for name in pdf.attachments.keys()}",
    "  root_piece=root.get('/PieceInfo')",
    "  page_piece=page.get('/PieceInfo')",
    "  print(json.dumps({'attachmentPayloads':attachments,'rootKeys':sorted(str(k) for k in root.keys()),'pageKeys':sorted(str(k) for k in page.keys()),'annotationCount':len(page.get('/Annots',[])),'formValue':str(fields[0].get('/V')) if fields else None,'xfa':xfa.read_bytes().decode('latin1','replace') if xfa is not None else None,'infoAuthor':str(pdf.docinfo.get('/Author')) if '/Author' in pdf.docinfo else None,'rootPrivateCanary':str(root_piece.get('/VendorPrivate').get('/Canary')) if root_piece is not None and root_piece.get('/VendorPrivate') is not None else None,'rootSearchCanary':str(root_piece.get('/SearchIndex').get('/Canary')) if root_piece is not None and root_piece.get('/SearchIndex') is not None else None,'pagePrivateCanary':str(page_piece.get('/PagePrivate').get('/Canary')) if page_piece is not None and page_piece.get('/PagePrivate') is not None else None,'canaries':{value:value in text for value in canaries}},sort_keys=True))",
  ].join("\n");
  return jsonResult(run(configuredPython, ["-c", code, input], { env, status: 0 }));
}

const manifest = (await fs.readFile(path.join(skillRoot, "manifest.txt"), "utf8"))
  .split(/\r?\n/)
  .filter(Boolean);
assert.ok(manifest.includes("scripts/pikepdf_provider.py"));
assert.ok(manifest.includes("tasks/structure_clean.md"));
const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
assert.match(skillText, /pikepdf_provider\.py/);
assert.match(skillText, /active-content.*active-and-auxiliary/is);
assert.match(skillText, /not.*redaction.*metadata.*XFA/is);
run("python3", ["-m", "py_compile", provider], { status: 0 });

const moduleProbe = run(configuredPython, [
  "-c",
  "from importlib.metadata import version; import pikepdf, pikepdf.sanitize; assert version('pikepdf') == '10.10.0'",
]);
if (moduleProbe.status !== 0) {
  console.log("pikepdf provider smoke skipped (set OPEN_OFFICE_PDF_PIKEPDF_TEST_PYTHON to a pikepdf 10.10.0 environment)");
  process.exit(0);
}

const providerEnv = {
  OPEN_OFFICE_PDF_PROVIDER_PYTHON: configuredPython,
  PYTHONNOUSERSITE: "1",
};
const probe = jsonResult(run(configuredPython, [provider, "probe"], { env: providerEnv, status: 0 }));
assert.equal(probe.provider, "pikepdf");
assert.equal(probe.providerVersion, "10.10.0");
assert.equal(probe.integration, "shipped-thin-script-external-python");
assert.equal(probe.providerIsRedactor, false);
assert.equal(probe.providerIsCompleteSanitizer, false);
assert.equal(probe.providerIsSandbox, false);
assert.equal(probe.silentFallback, false);
assert.deepEqual(Object.keys(probe.profiles).sort(), ["active-and-auxiliary", "active-content"]);

const registryProbe = jsonResult(run(configuredPython, [
  providerRegistry,
  "check",
  "--provider",
  "pikepdf",
  "--require",
], { env: providerEnv, status: 0 }));
assert.equal(registryProbe.providers[0].available, true);
assert.equal(registryProbe.providers[0].integration, "shipped-thin-script-external-python");
assert.equal(registryProbe.providers[0].evidence.minimumVersion, "10.10.0");
assert.equal(registryProbe.providers[0].evidence.maximumVersionExclusive, "10.11.0");

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-pikepdf-provider-"));
try {
  const source = path.join(tempRoot, "active-source.pdf");
  const sourceBytes = buildActiveContentFixture();
  const sourceHash = sha256(sourceBytes);
  await fs.writeFile(source, sourceBytes);

  const inspect = jsonResult(run(configuredPython, [
    provider,
    "inspect",
    source,
    "--expected-sha256",
    sourceHash,
    "--trusted-input",
  ], { env: providerEnv, status: 0 }));
  assert.equal(inspect.schema, "open-office-artifact-tool.pikepdf-inspect.v1");
  assert.equal(inspect.source.sha256, sourceHash);
  assert.equal(inspect.sourceProtected, true);
  assert.equal(inspect.sourceSnapshot.readOnly, true);
  assert.equal(inspect.structure.pageCount, 1);
  assert.equal(inspect.structure.annotationCount, 4);
  assert.equal(inspect.structure.formFieldCount, 1);
  assert.equal(inspect.structure.attachmentCount, 1);
  assert.equal(inspect.structure.hasXfa, true);
  assert.equal(inspect.structure.hasMetadata, true);
  assert.equal(inspect.structure.hasStructTreeRoot, true);
  assert.equal(inspect.structure.hasOutlines, true);
  assert.ok(inspect.structure.featureCounts.javascriptActions > 0);
  assert.ok(inspect.structure.featureCounts.externalAccessActions > 0);
  assert.ok(inspect.structure.featureCounts.multimediaActions > 0);
  assert.ok(inspect.structure.featureCounts.embeddedFileNameTrees > 0);
  assert.ok(inspect.structure.featureCounts.thumbnails > 0);
  assert.ok(inspect.structure.featureCounts.privateApplicationData > 0);
  assert.deepEqual(await fs.readFile(source), sourceBytes);

  const missingTrust = run(configuredPython, [
    provider,
    "inspect",
    source,
    "--expected-sha256",
    sourceHash,
  ], { env: providerEnv, status: 2 });
  assert.match(missingTrust.stderr, /one of the arguments --trusted-input --caller-isolated is required/);

  const inspectionPlan = jsonResult(run(configuredPython, [
    providerRegistry,
    "plan",
    "--task",
    "inspect",
    "--provider",
    "pikepdf",
    "--strategy",
    "read-only",
    "--input",
    source,
    "--require-provider",
  ], { env: providerEnv, status: 0 }));
  assert.equal(inspectionPlan.provider, "pikepdf");
  assert.equal(inspectionPlan.silentFallback, false);

  const activeOutput = path.join(tempRoot, "active-only.pdf");
  const cleanPlan = jsonResult(run(configuredPython, [
    providerRegistry,
    "plan",
    "--task",
    "structure-clean",
    "--provider",
    "pikepdf",
    "--strategy",
    "rewrite",
    "--input",
    source,
    "--output",
    activeOutput,
    "--invalidate-signatures",
    "--require-provider",
  ], { env: providerEnv, status: 0 }));
  assert.equal(cleanPlan.mutation, true);
  assert.equal(cleanPlan.invalidateSignatures, true);

  const noInvalidation = run(configuredPython, [
    provider,
    "clean",
    source,
    activeOutput,
    "--profile",
    "active-content",
    "--expected-sha256",
    sourceHash,
    "--trusted-input",
  ], { env: providerEnv, status: 2 });
  assert.match(jsonResult(noInvalidation, "stderr").error, /requires --invalidate-signatures/);
  await assert.rejects(fs.access(activeOutput));

  const activeResult = jsonResult(run(configuredPython, [
    provider,
    "clean",
    source,
    activeOutput,
    "--profile",
    "active-content",
    "--expected-sha256",
    sourceHash,
    "--trusted-input",
    "--invalidate-signatures",
  ], { env: providerEnv, status: 0 }));
  assert.equal(activeResult.schema, "open-office-artifact-tool.pikepdf-structure-clean.v1");
  assert.equal(activeResult.profile, "active-content");
  assert.equal(activeResult.savePolicy, "rewrite");
  assert.equal(activeResult.providerIsCompleteSanitizer, false);
  assert.equal(activeResult.validation.profilePostconditionsPassed, true);
  assert.equal(activeResult.validation.fullRewrite.sourcePrefixPreserved, false);
  assert.equal(activeResult.validation.fullRewrite.startxrefCount, 1);
  assert.equal(activeResult.validation.fullRewrite.eofCount, 1);
  assert.equal(activeResult.structureAfter.featureCounts.javascriptActions, 0);
  assert.equal(activeResult.structureAfter.featureCounts.javascriptNameTrees, 0);
  assert.equal(activeResult.structureAfter.featureCounts.externalAccessActions, 0);
  assert.equal(activeResult.structureAfter.featureCounts.multimediaActions, 0);
  assert.equal(activeResult.structureAfter.featureCounts.multimediaPayloadReferences, 0);
  assert.equal(activeResult.structureAfter.attachmentCount, 1);
  assert.ok(activeResult.structureAfter.featureCounts.privateApplicationData > 0);
  assert.ok(activeResult.structureAfter.featureCounts.thumbnails > 0);
  assert.equal(activeResult.structureAfter.formFieldCount, 1);
  assert.equal(activeResult.structureAfter.hasXfa, true);
  assert.equal(activeResult.structureAfter.hasMetadata, true);
  assert.equal(activeResult.structureAfter.annotationCount, 4);
  const activeBytes = await fs.readFile(activeOutput);
  assert.equal(activeResult.output.sha256, sha256(activeBytes));
  const activeOracle = pikepdfOracle(activeOutput, providerEnv);
  assert.equal(activeOracle.attachmentPayloads["payload.txt"], "payload-secret-canary");
  assert.ok(activeOracle.rootKeys.includes("/Collection"));
  assert.ok(activeOracle.rootKeys.includes("/PieceInfo"));
  assert.ok(activeOracle.pageKeys.includes("/PieceInfo"));
  assert.ok(activeOracle.pageKeys.includes("/Thumb"));
  assert.equal(activeOracle.canaries["payload-secret-canary"], true);
  assert.equal(activeOracle.rootPrivateCanary, "private-data-canary");
  assert.equal(activeOracle.rootSearchCanary, "stale-index-canary");
  assert.equal(activeOracle.pagePrivateCanary, "page-private-canary");
  assert.equal(activeOracle.canaries["Keep Metadata Canary"], true);
  assert.equal(activeOracle.canaries["keep-form-value-canary"], true);
  assert.equal(activeOracle.canaries["javascript-canary"], false);
  assert.equal(activeOracle.canaries["launch-canary.exe"], false);
  assert.equal(activeOracle.canaries["remote-canary.pdf"], false);
  assert.equal(activeOracle.canaries["rich-media-canary"], false);

  const allOutput = path.join(tempRoot, "active-and-auxiliary.pdf");
  const allResult = jsonResult(run(configuredPython, [
    provider,
    "clean",
    source,
    allOutput,
    "--profile",
    "active-and-auxiliary",
    "--expected-sha256",
    sourceHash,
    "--caller-isolated",
    "--invalidate-signatures",
  ], { env: providerEnv, status: 0 }));
  assert.equal(allResult.inputTrust, "caller-isolated");
  assert.equal(allResult.structureAfter.attachmentCount, 0);
  for (const count of Object.values(allResult.structureAfter.featureCounts)) assert.equal(count, 0);
  assert.deepEqual(allResult.validation.stableTopology, {
    annotationCount: 4,
    formFieldCount: 1,
    hasMetadata: true,
    hasOutlines: true,
    hasStructTreeRoot: true,
    hasXfa: true,
    pageCount: 1,
  });
  const allBytes = await fs.readFile(allOutput);
  assert.equal(allResult.output.sha256, sha256(allBytes));
  assert.ok(!allBytes.subarray(0, sourceBytes.length).equals(sourceBytes));
  assert.equal((allBytes.toString("latin1").match(/startxref/g) || []).length, 1);
  assert.equal((allBytes.toString("latin1").match(/%%EOF/g) || []).length, 1);
  const allOracle = pikepdfOracle(allOutput, providerEnv);
  for (const removed of [
    "javascript-canary",
    "launch-canary.exe",
    "remote-canary.pdf",
    "rich-media-canary",
    "payload-secret-canary",
  ]) assert.equal(allOracle.canaries[removed], false, `${removed} must be unreachable and absent after rewrite`);
  for (const retained of ["Visible pikepdf fixture", "Keep Metadata Canary", "keep-form-value-canary", "keep-xfa-canary"]) {
    assert.equal(allOracle.canaries[retained], true, `${retained} must remain outside the bounded structure-clean scope`);
  }
  assert.deepEqual(allOracle.attachmentPayloads, {});
  assert.ok(!allOracle.rootKeys.includes("/Collection"));
  assert.ok(!allOracle.rootKeys.includes("/PieceInfo"));
  assert.ok(!allOracle.pageKeys.includes("/PieceInfo"));
  assert.ok(!allOracle.pageKeys.includes("/Thumb"));
  assert.equal(allOracle.rootPrivateCanary, null);
  assert.equal(allOracle.rootSearchCanary, null);
  assert.equal(allOracle.pagePrivateCanary, null);
  assert.equal(allOracle.annotationCount, 4);
  assert.equal(allOracle.formValue, "keep-form-value-canary");
  assert.equal(allOracle.xfa, "<xfa>keep-xfa-canary</xfa>");
  assert.equal(allOracle.infoAuthor, "Keep Author Canary");
  assert.deepEqual(await fs.readFile(source), sourceBytes, "pikepdf must not mutate source bytes");

  const outputInspect = jsonResult(run(configuredPython, [
    provider,
    "inspect",
    allOutput,
    "--expected-sha256",
    allResult.output.sha256,
    "--trusted-input",
  ], { env: providerEnv, status: 0 }));
  assert.equal(outputInspect.structure.attachmentCount, 0);
  assert.ok(Object.values(outputInspect.structure.featureCounts).every((count) => count === 0));
  assert.equal((await PdfFile.inspectPdf(allBytes)).summary.pages, 1);

  const wrongHashOutput = path.join(tempRoot, "wrong-hash.pdf");
  const wrongHash = run(configuredPython, [
    provider,
    "clean",
    source,
    wrongHashOutput,
    "--profile",
    "active-content",
    "--expected-sha256",
    "0".repeat(64),
    "--trusted-input",
    "--invalidate-signatures",
  ], { env: providerEnv, status: 2 });
  assert.match(jsonResult(wrongHash, "stderr").error, /source SHA-256 mismatch/);
  await assert.rejects(fs.access(wrongHashOutput));

  const collision = run(configuredPython, [
    provider,
    "clean",
    source,
    allOutput,
    "--profile",
    "active-content",
    "--expected-sha256",
    sourceHash,
    "--trusted-input",
    "--invalidate-signatures",
  ], { env: providerEnv, status: 2 });
  assert.match(jsonResult(collision, "stderr").error, /already exists.*not be replaced/);

  const danglingTarget = path.join(tempRoot, "dangling-target.pdf");
  const symlinkOutput = path.join(tempRoot, "symlink-output.pdf");
  await fs.symlink(danglingTarget, symlinkOutput);
  const symlink = run(configuredPython, [
    provider,
    "clean",
    source,
    symlinkOutput,
    "--profile",
    "active-content",
    "--expected-sha256",
    sourceHash,
    "--trusted-input",
    "--invalidate-signatures",
  ], { env: providerEnv, status: 2 });
  assert.match(jsonResult(symlink, "stderr").error, /symbolic link.*not be followed/);
  await assert.rejects(fs.access(danglingTarget));

  const overObjectBudget = run(configuredPython, [
    provider,
    "inspect",
    source,
    "--expected-sha256",
    sourceHash,
    "--trusted-input",
    "--max-objects",
    "5",
  ], { env: providerEnv, status: 2 });
  assert.match(jsonResult(overObjectBudget, "stderr").error, /object count.*exceeds.*object budget/);

  const overWorkerOutput = run(configuredPython, [
    provider,
    "inspect",
    source,
    "--expected-sha256",
    sourceHash,
    "--trusted-input",
    "--max-stdout-bytes",
    "64",
  ], { env: providerEnv, status: 2 });
  assert.match(jsonResult(overWorkerOutput, "stderr").error, /worker stdout exceeded the 64 byte budget/);

  const outputBudget = path.join(tempRoot, "output-budget.pdf");
  const overOutputBudget = run(configuredPython, [
    provider,
    "clean",
    source,
    outputBudget,
    "--profile",
    "active-content",
    "--expected-sha256",
    sourceHash,
    "--trusted-input",
    "--invalidate-signatures",
    "--max-output-bytes",
    "64",
  ], { env: providerEnv, status: 2 });
  assert.match(jsonResult(overOutputBudget, "stderr").error, /output PDF size.*outside/);
  await assert.rejects(fs.access(outputBudget));

  const encrypted = path.join(tempRoot, "encrypted.pdf");
  run(configuredPython, [
    "-c",
    "import pikepdf,sys; p=pikepdf.open(sys.argv[1]); p.save(sys.argv[2], encryption=pikepdf.Encryption(owner='owner-secret', user='', R=6)); p.close()",
    source,
    encrypted,
  ], { env: providerEnv, status: 0 });
  const encryptedHash = sha256(await fs.readFile(encrypted));
  const encryptedOutput = path.join(tempRoot, "encrypted-output.pdf");
  const encryptedRefused = run(configuredPython, [
    provider,
    "clean",
    encrypted,
    encryptedOutput,
    "--profile",
    "active-content",
    "--expected-sha256",
    encryptedHash,
    "--trusted-input",
    "--invalidate-signatures",
  ], { env: providerEnv, status: 2 });
  assert.match(jsonResult(encryptedRefused, "stderr").error, /encrypted PDFs are unsupported/);
  await assert.rejects(fs.access(encryptedOutput));

  if (available("qpdf")) {
    const qpdfInspect = jsonResult(run("python3", [qpdfProvider, "inspect", allOutput], {
      env: { OPEN_OFFICE_PDF_QPDF: "" },
      status: 0,
    }));
    assert.equal(qpdfInspect.check.status, "clean");
    assert.equal(qpdfInspect.structure.pageCount, 1);
    assert.equal(qpdfInspect.structure.attachmentCount, 0);
    assert.equal(qpdfInspect.structure.annotationCount, 4);
  }

  if (available("pdftoppm", ["-v"])) {
    for (const [input, name] of [[source, "source"], [activeOutput, "active"], [allOutput, "all"]]) {
      run("pdftoppm", ["-png", "-singlefile", "-r", "96", input, path.join(tempRoot, name)], { status: 0 });
    }
    const sourcePng = await fs.readFile(path.join(tempRoot, "source.png"));
    assert.deepEqual(await fs.readFile(path.join(tempRoot, "active.png")), sourcePng);
    assert.deepEqual(await fs.readFile(path.join(tempRoot, "all.png")), sourcePng);
  }
} finally {
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log("pikepdf provider smoke ok");
