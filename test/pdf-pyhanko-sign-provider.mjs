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
const provider = path.join(skillRoot, "scripts", "pyhanko_sign_provider.py");
const validator = path.join(skillRoot, "scripts", "pyhanko_provider.py");
const registry = path.join(skillRoot, "scripts", "pdf_provider.py");
const configuredPython = process.env.OPEN_OFFICE_PYHANKO_TEST_PYTHON;
const passphrase = Buffer.from("open-office-signing-fixture-passphrase\n", "utf8");

function run(executable, args, options = {}) {
  const result = spawnSync(executable, args, {
    cwd: options.cwd || repoRoot,
    encoding: "utf8",
    env: { ...process.env, PYTHONDONTWRITEBYTECODE: "1", ...options.env },
    input: options.input,
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

function supportedPyHanko(executable) {
  if (!executable) return false;
  return run(executable, [
    "-c",
    "from importlib.metadata import version; h=tuple(int(x) for x in version('pyHanko').split('.')[:3]); c=tuple(int(x) for x in version('pyhanko-certvalidator').split('.')[:3]); assert (0,35,0) <= h < (0,36,0); assert (0,31,0) <= c < (0,32,0)",
  ]).status === 0;
}

function commandAvailable(command, args = ["--version"]) {
  return spawnSync(command, args, { stdio: "ignore" }).status === 0;
}

async function assertPixelStable(source, output, tempRoot, label) {
  const sourcePrefix = path.join(tempRoot, `${label}-source`);
  const outputPrefix = path.join(tempRoot, `${label}-output`);
  run("pdftoppm", ["-png", "-singlefile", "-r", "72", source, sourcePrefix], { status: 0 });
  run("pdftoppm", ["-png", "-singlefile", "-r", "72", output, outputPrefix], { status: 0 });
  const sourcePixels = await sharp(`${sourcePrefix}.png`).raw().toBuffer({ resolveWithObject: true });
  const outputPixels = await sharp(`${outputPrefix}.png`).raw().toBuffer({ resolveWithObject: true });
  assert.deepEqual(outputPixels.info, sourcePixels.info);
  assert.deepEqual(outputPixels.data, sourcePixels.data);
}

async function changedPixelBounds(source, output, tempRoot, label) {
  const sourcePrefix = path.join(tempRoot, `${label}-source`);
  const outputPrefix = path.join(tempRoot, `${label}-output`);
  run("pdftoppm", ["-png", "-singlefile", "-r", "72", source, sourcePrefix], { status: 0 });
  run("pdftoppm", ["-png", "-singlefile", "-r", "72", output, outputPrefix], { status: 0 });
  const sourcePixels = await sharp(`${sourcePrefix}.png`).raw().toBuffer({ resolveWithObject: true });
  const outputPixels = await sharp(`${outputPrefix}.png`).raw().toBuffer({ resolveWithObject: true });
  assert.deepEqual(outputPixels.info, sourcePixels.info);
  const { width, height, channels } = sourcePixels.info;
  let minX = width;
  let minY = height;
  let maxX = -1;
  let maxY = -1;
  let count = 0;
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const offset = (y * width + x) * channels;
      let different = false;
      for (let channel = 0; channel < channels; channel += 1) {
        if (sourcePixels.data[offset + channel] !== outputPixels.data[offset + channel]) {
          different = true;
          break;
        }
      }
      if (different) {
        count += 1;
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x);
        maxY = Math.max(maxY, y);
      }
    }
  }
  return { count, bounds: [minX, minY, maxX, maxY], width, height };
}

const manifest = (await fs.readFile(path.join(skillRoot, "manifest.txt"), "utf8"))
  .split(/\r?\n/)
  .filter(Boolean);
assert.ok(manifest.includes("scripts/pyhanko_sign_provider.py"));
const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
assert.match(skillText, /pyhanko_sign_provider\.py/);
assert.match(skillText, /passphrase.*stdin/is);
const providerText = await fs.readFile(provider, "utf8");
assert.doesNotMatch(providerText, /candidate\.read_bytes\(\)|source_snapshot\.read_bytes\(\)/);
assert.match(providerText, /def has_exact_prefix/);
run("python3", ["-m", "py_compile", provider], { status: 0 });

if (configuredPython) {
  assert.ok(supportedPyHanko(configuredPython), `OPEN_OFFICE_PYHANKO_TEST_PYTHON must provide the supported pyHanko runtime: ${configuredPython}`);
}
const python = configuredPython || (supportedPyHanko("python3") ? "python3" : null);
if (!python) {
  const unavailable = run("python3", [provider, "probe"], { status: 2 });
  assert.equal(jsonResult(unavailable, "stderr").silentFallback, false);
  console.log("pyHanko signing provider smoke ok (real provider skipped: set OPEN_OFFICE_PYHANKO_TEST_PYTHON)");
  process.exit(0);
}

const providerEnv = {
  OPEN_OFFICE_PDF_PROVIDER_PYTHON: python,
  PYTHONNOUSERSITE: "1",
};
run(python, ["-c", [
  "import importlib.util,pathlib,sys,types",
  "path=pathlib.Path(sys.argv[1])",
  "sys.path.insert(0,str(path.parent))",
  "spec=importlib.util.spec_from_file_location('pyhanko_sign_provider',path)",
  "module=importlib.util.module_from_spec(spec)",
  "spec.loader.exec_module(module)",
  "module.sys.stdin=type('TTY',(),{'isatty':lambda self: True})()",
  "module.getpass.getpass=lambda prompt: 'hidden-tty-fixture'",
  "secret=module.read_passphrase(types.SimpleNamespace(no_passphrase=False))",
  "assert bytes(secret)==b'hidden-tty-fixture'",
  "assert 'hidden' in module.probe()['passphraseChannels'][0]",
].join(";"), provider], { env: providerEnv, status: 0 });
const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-pyhanko-sign-provider-"));
try {
  const artifact = PdfArtifact.create({ text: "Source-bound pyHanko signing fixture" });
  const exported = await PdfFile.exportPdf(artifact);
  const source = path.join(tempRoot, "source.pdf");
  const sourceBytes = Buffer.from(exported.bytes);
  const sourceHash = sha256(sourceBytes);
  await fs.writeFile(source, sourceBytes);

  const credentialBuilder = path.join(tempRoot, "build_credential.py");
  await fs.writeFile(credentialBuilder, String.raw`
from datetime import datetime, timedelta, timezone
from pathlib import Path
import sys

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives.serialization import pkcs12
from cryptography.x509.oid import NameOID
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from pyhanko.sign.fields import SigFieldSpec, append_signature_field

root = Path(sys.argv[1])
password = sys.argv[2].encode("utf-8")
now = datetime.now(timezone.utc)

root_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
root_name = x509.Name([
    x509.NameAttribute(NameOID.COMMON_NAME, "Open Office Signing Test Root"),
    x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Open Office Artifact Tool"),
    x509.NameAttribute(NameOID.COUNTRY_NAME, "US"),
])
root_cert = (
    x509.CertificateBuilder()
    .subject_name(root_name).issuer_name(root_name).public_key(root_key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(now - timedelta(days=1)).not_valid_after(now + timedelta(days=3650))
    .add_extension(x509.BasicConstraints(ca=True, path_length=1), critical=True)
    .add_extension(x509.KeyUsage(
        digital_signature=True, content_commitment=True, key_encipherment=False,
        data_encipherment=False, key_agreement=False, key_cert_sign=True,
        crl_sign=True, encipher_only=None, decipher_only=None,
    ), critical=True)
    .sign(root_key, hashes.SHA256())
)

signer_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
signer_name = x509.Name([
    x509.NameAttribute(NameOID.COMMON_NAME, "Open Office Signing Fixture"),
    x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Open Office Artifact Tool"),
    x509.NameAttribute(NameOID.COUNTRY_NAME, "US"),
])
signer_cert = (
    x509.CertificateBuilder()
    .subject_name(signer_name).issuer_name(root_name).public_key(signer_key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(now - timedelta(days=1)).not_valid_after(now + timedelta(days=365))
    .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
    .add_extension(x509.KeyUsage(
        digital_signature=True, content_commitment=True, key_encipherment=False,
        data_encipherment=False, key_agreement=False, key_cert_sign=False,
        crl_sign=False, encipher_only=None, decipher_only=None,
    ), critical=True)
    .sign(root_key, hashes.SHA256())
)

(root / "credential.p12").write_bytes(pkcs12.serialize_key_and_certificates(
    b"open-office-signer", signer_key, signer_cert, [root_cert],
    serialization.BestAvailableEncryption(password),
))
(root / "credential-unencrypted.p12").write_bytes(pkcs12.serialize_key_and_certificates(
    b"open-office-signer", signer_key, signer_cert, [root_cert],
    serialization.NoEncryption(),
))
(root / "root.pem").write_bytes(root_cert.public_bytes(serialization.Encoding.PEM))
(root / "signer.pem").write_bytes(signer_cert.public_bytes(serialization.Encoding.PEM))

with (root / "source.pdf").open("rb") as source, (root / "existing-field.pdf").open("wb") as output:
    writer = IncrementalPdfFileWriter(source, strict=True)
    append_signature_field(writer, SigFieldSpec(
        sig_field_name="ExistingApproval", on_page=0, box=(320, 72, 520, 140)
    ))
    writer.write(output)
`, "utf8");
  run(python, [credentialBuilder, tempRoot, passphrase.toString("utf8").trimEnd()], {
    env: providerEnv,
    status: 0,
  });

  const credential = path.join(tempRoot, "credential.p12");
  const rootCertificate = path.join(tempRoot, "root.pem");
  const credentialBytes = await fs.readFile(credential);
  const credentialHash = sha256(credentialBytes);
  await fs.chmod(credential, 0o600);

  const probe = jsonResult(run(python, [provider, "probe"], { env: providerEnv, status: 0 }));
  assert.match(probe.providerVersion, /^0\.35\./);
  assert.match(probe.certvalidatorVersion, /^0\.31\./);
  assert.deepEqual(probe.fieldModes, ["existing", "create-invisible", "create-visible"]);
  assert.equal(probe.networkAllowed, false);
  assert.equal(probe.timestampAuthoritySupported, false);
  assert.equal(probe.ltvEmbeddingSupported, false);
  assert.equal(probe.silentFallback, false);

  const inspect = jsonResult(run(python, [
    provider,
    "inspect",
    source,
    "--expected-sha256",
    sourceHash,
    "--page-index",
    "0",
    "--trusted-input",
  ], { env: providerEnv, status: 0 }));
  assert.equal(inspect.schema, "open-office-artifact-tool.pyhanko-signing-inspect.v1");
  assert.equal(inspect.summary.signatureCount, 0);
  assert.equal(inspect.summary.fieldCount, 0);
  assert.equal(inspect.pageCount, 1);
  assert.equal(inspect.selectedPage.rotation, 0);
  assert.deepEqual(inspect.selectedPage.cropBox, [0, 0, 612, 792]);
  assert.deepEqual(await fs.readFile(source), sourceBytes);

  const registryProbe = jsonResult(run(python, [
    registry,
    "check",
    "--provider",
    "pyhanko",
    "--require",
  ], { env: providerEnv, status: 0 }));
  assert.equal(registryProbe.providers[0].available, true);
  assert.match(registryProbe.providers[0].role, /signing plus read-only signature validation/);
  const plan = jsonResult(run(python, [
    registry,
    "plan",
    "--task",
    "sign",
    "--provider",
    "pyhanko",
    "--strategy",
    "incremental",
    "--input",
    source,
    "--output",
    path.join(tempRoot, "planned.pdf"),
    "--require-provider",
  ], { env: providerEnv, status: 0 }));
  assert.equal(plan.integration, "shipped-thin-script");
  assert.equal(plan.silentFallback, false);

  const visible = path.join(tempRoot, "certified-visible.pdf");
  const visibleResult = jsonResult(run(python, [
    provider,
    "sign",
    source,
    visible,
    "--expected-sha256",
    sourceHash,
    "--credential",
    credential,
    "--credential-sha256",
    credentialHash,
    "--passphrase-stdin",
    "--field-name",
    "Certification",
    "--field-mode",
    "create-visible",
    "--page-index",
    "0",
    "--box",
    "72,72,300,150",
    "--signature-kind",
    "certification",
    "--docmdp-permission",
    "fill-forms",
    "--subfilter",
    "pades",
    "--expected-signature-count",
    "0",
    "--reason",
    "Approve bounded signing fixture",
    "--location",
    "Test environment",
    "--trusted-input",
  ], { env: providerEnv, input: passphrase, status: 0 }));
  assert.equal(visibleResult.schema, "open-office-artifact-tool.pyhanko-sign.v1");
  assert.equal(visibleResult.savePolicy, "incremental");
  assert.equal(visibleResult.source.sha256, sourceHash);
  assert.equal(visibleResult.credential.sha256, credentialHash);
  assert.equal(visibleResult.credential.passphraseChannel, "stdin");
  assert.equal(visibleResult.credential.secretLogged, false);
  assert.equal(visibleResult.credential.certificateTrustValidated, false);
  assert.match(visibleResult.credential.certificate.subject, /Open Office Signing Fixture/);
  assert.equal(visibleResult.signature.fieldName, "Certification");
  assert.equal(visibleResult.signature.signatureKind, "certification");
  assert.equal(visibleResult.signature.docMDPPermission, "fill-forms");
  assert.equal(visibleResult.signature.timestampAuthorityUsed, false);
  assert.equal(visibleResult.signature.ltvEnabled, false);
  assert.equal(visibleResult.validation.sourcePrefixPreserved, true);
  assert.equal(visibleResult.validation.signatureCountDelta, 1);
  assert.equal(visibleResult.validation.newSignature.coverage, "entire-file");
  assert.equal(visibleResult.validation.newSignature.subFilter, "/ETSI.CAdES.detached");
  assert.equal(visibleResult.validation.allIntegrityValid, true);
  assert.equal(visibleResult.validation.allDocMDPCompliant, true);
  assert.ok(!JSON.stringify(visibleResult).includes(passphrase.toString("utf8").trim()));
  const visibleBytes = await fs.readFile(visible);
  assert.ok(visibleBytes.subarray(0, sourceBytes.length).equals(sourceBytes));
  assert.deepEqual(await fs.readFile(source), sourceBytes);
  assert.deepEqual(await fs.readFile(credential), credentialBytes);

  const trusted = jsonResult(run(python, [
    validator,
    "verify",
    visible,
    "--expected-sha256",
    sha256(visibleBytes),
    "--trust-policy",
    "explicit-roots",
    "--trust-root",
    rootCertificate,
    "--revocation-policy",
    "none",
    "--require-signature",
    "--require-all-integrity-valid",
    "--require-all-trusted",
    "--require-docmdp-compliant",
    "--require-all-bottom-line",
  ], { env: providerEnv, status: 0 }));
  assert.equal(trusted.conclusion, "valid-under-selected-policy");
  assert.equal(trusted.signatures[0].fieldName, "Certification");
  assert.equal(trusted.signatures[0].docMDP.permission, "fill-forms");
  assert.equal(trusted.signatures[0].trusted, true);

  assert.equal((await PdfFile.inspectPdf(visibleBytes)).summary.pages, 1);
  if (commandAvailable("qpdf")) {
    run("qpdf", ["--check", visible], { status: 0 });
  }
  if (commandAvailable("pdftoppm", ["-v"])) {
    const changed = await changedPixelBounds(source, visible, tempRoot, "visible-signature");
    assert.ok(changed.count > 0, "visible signature appearance must change rendered pixels");
    const [minX, minY, maxX, maxY] = changed.bounds;
    assert.ok(minX >= 68 && maxX <= 304, `visible signature changed pixels outside horizontal field bounds: ${changed.bounds}`);
    assert.ok(minY >= changed.height - 154 && maxY <= changed.height - 68, `visible signature changed pixels outside vertical field bounds: ${changed.bounds}`);
  }

  const countersigned = path.join(tempRoot, "countersigned.pdf");
  const countersignedResult = jsonResult(run(python, [
    provider,
    "sign",
    visible,
    countersigned,
    "--expected-sha256",
    sha256(visibleBytes),
    "--credential",
    credential,
    "--credential-sha256",
    credentialHash,
    "--passphrase-stdin",
    "--field-name",
    "Reviewer",
    "--field-mode",
    "create-invisible",
    "--signature-kind",
    "approval",
    "--subfilter",
    "adobe-pkcs7-detached",
    "--expected-signature-count",
    "1",
    "--allow-existing-signatures",
    "--trusted-input",
  ], { env: providerEnv, input: passphrase, status: 0 }));
  assert.equal(countersignedResult.existingSignatures.before.signatureCount, 1);
  assert.equal(countersignedResult.existingSignatures.after.signatureCount, 2);
  assert.equal(countersignedResult.existingSignatures.preflightValidated, true);
  assert.equal(countersignedResult.existingSignatures.preflightAllIntegrityValid, true);
  assert.equal(countersignedResult.existingSignatures.preflightAllDocMDPCompliant, true);
  assert.equal(countersignedResult.existingSignatures.oldSignerApprovalOfNewRevisionClaimed, false);
  const countersignedBytes = await fs.readFile(countersigned);
  assert.ok(countersignedBytes.subarray(0, visibleBytes.length).equals(visibleBytes));
  const countersignedVerify = jsonResult(run(python, [
    validator,
    "verify",
    countersigned,
    "--expected-sha256",
    sha256(countersignedBytes),
    "--trust-policy",
    "explicit-roots",
    "--trust-root",
    rootCertificate,
    "--revocation-policy",
    "none",
    "--require-signature",
    "--require-all-integrity-valid",
    "--require-all-trusted",
    "--require-docmdp-compliant",
    "--require-all-bottom-line",
  ], { env: providerEnv, status: 0 }));
  assert.deepEqual(countersignedVerify.signatures.map((record) => record.fieldName), ["Certification", "Reviewer"]);
  assert.equal(countersignedVerify.signatures[0].modificationLevel, "form-filling");
  assert.equal(countersignedVerify.signatures[1].coverage, "entire-file");
  if (commandAvailable("pdftoppm", ["-v"])) {
    await assertPixelStable(visible, countersigned, tempRoot, "invisible-countersignature");
  }

  const existingFieldSource = path.join(tempRoot, "existing-field.pdf");
  const existingFieldBytes = await fs.readFile(existingFieldSource);
  const existingFieldOutput = path.join(tempRoot, "existing-field-signed.pdf");
  const existingResult = jsonResult(run(python, [
    provider,
    "sign",
    existingFieldSource,
    existingFieldOutput,
    "--expected-sha256",
    sha256(existingFieldBytes),
    "--credential",
    credential,
    "--credential-sha256",
    credentialHash,
    "--passphrase-stdin",
    "--field-name",
    "ExistingApproval",
    "--field-mode",
    "existing",
    "--signature-kind",
    "approval",
    "--expected-signature-count",
    "0",
    "--trusted-input",
  ], { env: providerEnv, input: passphrase, status: 0 }));
  assert.equal(existingResult.signature.fieldMode, "existing");
  assert.equal(existingResult.existingSignatures.before.emptyFieldCount, 1);
  assert.ok((await fs.readFile(existingFieldOutput)).subarray(0, existingFieldBytes.length).equals(existingFieldBytes));

  const unencryptedCredential = path.join(tempRoot, "credential-unencrypted.p12");
  const unencryptedCredentialBytes = await fs.readFile(unencryptedCredential);
  const noPassphraseOutput = path.join(tempRoot, "no-passphrase.pdf");
  const noPassphraseResult = jsonResult(run(python, [
    provider,
    "sign",
    source,
    noPassphraseOutput,
    "--expected-sha256",
    sourceHash,
    "--credential",
    unencryptedCredential,
    "--credential-sha256",
    sha256(unencryptedCredentialBytes),
    "--no-passphrase",
    "--field-name",
    "UnencryptedCredentialApproval",
    "--field-mode",
    "create-invisible",
    "--signature-kind",
    "approval",
    "--expected-signature-count",
    "0",
    "--trusted-input",
  ], { env: providerEnv, status: 0 }));
  assert.equal(noPassphraseResult.credential.passphraseChannel, "none");
  assert.equal(noPassphraseResult.validation.allIntegrityValid, true);
  assert.ok((await fs.readFile(noPassphraseOutput)).subarray(0, sourceBytes.length).equals(sourceBytes));

  const missingTrust = run(python, [
    provider,
    "inspect",
    source,
    "--expected-sha256",
    sourceHash,
  ], { env: providerEnv, status: 2 });
  assert.match(missingTrust.stderr, /trusted-input --caller-isolated/);

  const staleSourceOutput = path.join(tempRoot, "stale-source.pdf");
  const staleSource = run(python, [
    provider,
    "sign",
    source,
    staleSourceOutput,
    "--credential",
    credential,
    "--credential-sha256",
    credentialHash,
    "--passphrase-stdin",
    "--field-name",
    "StaleSource",
    "--field-mode",
    "create-invisible",
    "--signature-kind",
    "approval",
    "--expected-signature-count",
    "0",
    "--expected-sha256",
    "0".repeat(64),
    "--trusted-input",
  ], { env: providerEnv, input: passphrase, status: 2 });
  assert.match(jsonResult(staleSource, "stderr").error, /source SHA-256 mismatch/);
  await assert.rejects(fs.access(staleSourceOutput));

  const baseSignArgs = [
    provider,
    "sign",
    source,
    path.join(tempRoot, "negative.pdf"),
    "--credential",
    credential,
    "--credential-sha256",
    credentialHash,
    "--passphrase-stdin",
    "--field-name",
    "Negative",
    "--field-mode",
    "create-invisible",
    "--signature-kind",
    "approval",
    "--expected-signature-count",
    "0",
    "--expected-sha256",
    sourceHash,
    "--trusted-input",
  ];
  const missingVisibleGeometryArgs = baseSignArgs.map((value, index) => (
    value === "create-invisible" && baseSignArgs[index - 1] === "--field-mode" ? "create-visible" : value
  ));
  const missingVisibleGeometry = run(python, missingVisibleGeometryArgs, {
    env: providerEnv,
    input: passphrase,
    status: 2,
  });
  assert.match(jsonResult(missingVisibleGeometry, "stderr").error, /requires both --page-index and --box/);

  const outsideVisibleBox = run(python, [
    ...missingVisibleGeometryArgs,
    "--page-index", "0",
    "--box", "72,72,700,150",
  ], { env: providerEnv, input: passphrase, status: 2 });
  assert.match(jsonResult(outsideVisibleBox, "stderr").error, /fit wholly inside.*CropBox/);

  const approvalDocMDP = run(python, [
    ...baseSignArgs,
    "--docmdp-permission", "fill-forms",
  ], { env: providerEnv, input: passphrase, status: 2 });
  assert.match(jsonResult(approvalDocMDP, "stderr").error, /accepted only for certification/);

  const certificationWithoutDocMDPArgs = baseSignArgs.map((value, index) => (
    value === "approval" && baseSignArgs[index - 1] === "--signature-kind" ? "certification" : value
  ));
  const certificationWithoutDocMDP = run(python, certificationWithoutDocMDPArgs, {
    env: providerEnv,
    input: passphrase,
    status: 2,
  });
  assert.match(jsonResult(certificationWithoutDocMDP, "stderr").error, /require an explicit --docmdp-permission/);

  const staleCredential = run(python, baseSignArgs.map((value, index) => (
    value === credentialHash && baseSignArgs[index - 1] === "--credential-sha256" ? "0".repeat(64) : value
  )), { env: providerEnv, input: passphrase, status: 2 });
  assert.match(jsonResult(staleCredential, "stderr").error, /credential SHA-256 mismatch/);

  const wrongPasswordOutput = path.join(tempRoot, "wrong-password.pdf");
  const wrongPasswordArgs = baseSignArgs.map((value) => value === path.join(tempRoot, "negative.pdf") ? wrongPasswordOutput : value);
  const wrongPassword = run(python, wrongPasswordArgs, {
    env: providerEnv,
    input: Buffer.from("do-not-log-this-secret\n"),
    status: 2,
  });
  assert.ok(!wrongPassword.stderr.includes("do-not-log-this-secret"));
  assert.match(jsonResult(wrongPassword, "stderr").error, /signing worker failed/);
  await assert.rejects(fs.access(wrongPasswordOutput));

  const oversizedSecret = run(python, baseSignArgs, {
    env: providerEnv,
    input: Buffer.alloc(4_097, 0x61),
    status: 2,
  });
  assert.match(jsonResult(oversizedSecret, "stderr").error, /passphrase exceeds/);

  const emptySecret = run(python, baseSignArgs, {
    env: providerEnv,
    input: Buffer.from("\n"),
    status: 2,
  });
  assert.match(jsonResult(emptySecret, "stderr").error, /empty PKCS#12 passphrase.*--no-passphrase/);

  const missingExistingOutput = path.join(tempRoot, "missing-existing.pdf");
  const missingExisting = run(python, [
    ...baseSignArgs.slice(0, 3), missingExistingOutput,
    ...baseSignArgs.slice(4).map((value, index) => {
      const previous = baseSignArgs[index + 3];
      if (previous === "--field-name") return "MissingExisting";
      if (previous === "--field-mode") return "existing";
      return value;
    }),
  ], { env: providerEnv, input: passphrase, status: 2 });
  assert.match(jsonResult(missingExisting, "stderr").error, /exactly one empty signature field/);
  await assert.rejects(fs.access(missingExistingOutput));

  const noCounterAck = run(python, [
    provider,
    "sign",
    visible,
    path.join(tempRoot, "no-counter-ack.pdf"),
    "--credential",
    credential,
    "--credential-sha256",
    credentialHash,
    "--passphrase-stdin",
    "--field-name",
    "UnacknowledgedReviewer",
    "--field-mode",
    "create-invisible",
    "--signature-kind",
    "approval",
    "--expected-signature-count",
    "1",
    "--expected-sha256",
    sha256(visibleBytes),
    "--trusted-input",
  ], { env: providerEnv, input: passphrase, status: 2 });
  assert.match(jsonResult(noCounterAck, "stderr").error, /existing signatures require/);

  const recertify = run(python, [
    provider,
    "sign",
    visible,
    path.join(tempRoot, "recertify.pdf"),
    "--credential",
    credential,
    "--credential-sha256",
    credentialHash,
    "--passphrase-stdin",
    "--field-name",
    "SecondCertification",
    "--field-mode",
    "create-invisible",
    "--signature-kind",
    "certification",
    "--docmdp-permission",
    "fill-forms",
    "--expected-signature-count",
    "1",
    "--allow-existing-signatures",
    "--expected-sha256",
    sha256(visibleBytes),
    "--trusted-input",
  ], { env: providerEnv, input: passphrase, status: 2 });
  assert.match(jsonResult(recertify, "stderr").error, /must be the first signature/);

  const outputBudget = path.join(tempRoot, "output-budget.pdf");
  const outputBudgetResult = run(python, [
    ...baseSignArgs.slice(0, 3), outputBudget, ...baseSignArgs.slice(4),
    "--max-output-bytes", String(sourceBytes.length + 1),
  ], { env: providerEnv, input: passphrase, status: 2 });
  assert.match(jsonResult(outputBudgetResult, "stderr").error, /signed PDF size.*outside/);
  await assert.rejects(fs.access(outputBudget));

  const collision = path.join(tempRoot, "collision.pdf");
  await fs.writeFile(collision, "owned-by-caller");
  const collisionResult = run(python, [
    ...baseSignArgs.slice(0, 3), collision, ...baseSignArgs.slice(4),
  ], { env: providerEnv, input: passphrase, status: 2 });
  assert.match(jsonResult(collisionResult, "stderr").error, /already exists.*not be replaced/);
  assert.equal(await fs.readFile(collision, "utf8"), "owned-by-caller");

  const credentialLink = path.join(tempRoot, "credential-link.p12");
  await fs.symlink(credential, credentialLink);
  const symlinkArgs = baseSignArgs.map((value, index) => (
    value === credential && baseSignArgs[index - 1] === "--credential" ? credentialLink : value
  ));
  const symlinkCredential = run(python, symlinkArgs, { env: providerEnv, input: passphrase, status: 2 });
  assert.match(jsonResult(symlinkCredential, "stderr").error, /credential is a symbolic link/);

  if (commandAvailable("qpdf")) {
    const encrypted = path.join(tempRoot, "encrypted.pdf");
    run("qpdf", ["--encrypt", "user", "owner", "256", "--", source, encrypted], { status: 0 });
    const encryptedBytes = await fs.readFile(encrypted);
    const encryptedResult = run(python, [
      provider,
      "inspect",
      encrypted,
      "--expected-sha256",
      sha256(encryptedBytes),
      "--trusted-input",
    ], { env: providerEnv, status: 2 });
    assert.match(jsonResult(encryptedResult, "stderr").error, /encrypted PDFs are unsupported/);
  }

  assert.deepEqual(await fs.readFile(source), sourceBytes);
  assert.deepEqual(await fs.readFile(credential), credentialBytes);
  assert.deepEqual(await fs.readFile(unencryptedCredential), unencryptedCredentialBytes);
} finally {
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log("pyHanko signing provider smoke ok");
