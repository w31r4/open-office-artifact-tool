import assert from "node:assert/strict";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { PdfArtifact, PdfFile } from "../src/index.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const skillRoot = path.join(repoRoot, "skills", "pdf", "skills", "pdf");
const provider = path.join(skillRoot, "scripts", "pyhanko_provider.py");
const registry = path.join(skillRoot, "scripts", "pdf_provider.py");

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

function supportedPyHanko(executable) {
  if (!executable) return false;
  const result = run(executable, [
    "-c",
    "from importlib.metadata import version; v=tuple(int(x) for x in version('pyHanko').split('.')[:3]); raise SystemExit(0 if (0,35,0) <= v < (0,36,0) else 1)",
  ]);
  return result.status === 0;
}

const manifest = (await fs.readFile(path.join(skillRoot, "manifest.txt"), "utf8")).split(/\r?\n/).filter(Boolean);
assert.ok(manifest.includes("scripts/pyhanko_provider.py"));
const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
assert.match(skillText, /pyhanko_provider\.py/);
assert.match(skillText, /explicit trust root/i);
const signVerifyText = await fs.readFile(path.join(skillRoot, "tasks", "sign_verify.md"), "utf8");
assert.match(signVerifyText, /require-all-integrity-valid/);
assert.match(signVerifyText, /complete PAdES profile conformance/i);

const configuredPython = process.env.OPEN_OFFICE_PYHANKO_TEST_PYTHON;
if (configuredPython) {
  assert.ok(supportedPyHanko(configuredPython), `OPEN_OFFICE_PYHANKO_TEST_PYTHON must provide pyHanko 0.35.x: ${configuredPython}`);
}
const python = configuredPython || (supportedPyHanko("python3") ? "python3" : null);

if (!python) {
  const unavailable = run("python3", [provider, "probe"], { status: 2 });
  assert.equal(jsonResult(unavailable, "stderr").silentFallback, false);
  console.log("pyHanko provider smoke ok (real provider skipped: set OPEN_OFFICE_PYHANKO_TEST_PYTHON)");
  process.exit(0);
}

const providerEnv = {
  OPEN_OFFICE_PDF_PROVIDER_PYTHON: python,
  PYTHONNOUSERSITE: "1",
};

function runProvider(args, options = {}) {
  return run(python, args, {
    ...options,
    env: { ...providerEnv, ...options.env },
  });
}

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-pyhanko-provider-"));
try {
  const artifact = PdfArtifact.create({ text: "pyHanko source-bound signature fixture" });
  const exported = await PdfFile.exportPdf(artifact);
  const source = path.join(tempRoot, "source.pdf");
  const sourceBytes = Buffer.from(exported.bytes);
  await fs.writeFile(source, sourceBytes);
  const sourceHash = sha256(sourceBytes);

  const probe = jsonResult(runProvider([provider, "probe"], { status: 0 }));
  assert.equal(probe.provider, "pyhanko");
  assert.match(probe.providerVersion, /^0\.35\./);
  assert.match(probe.certvalidatorVersion, /^0\.31\./);
  assert.equal(probe.integration, "shipped-thin-script-external-python");
  assert.equal(probe.networkAllowed, false);
  assert.equal(probe.silentFallback, false);
  assert.equal(probe.padesProfileConformanceClaimed, false);

  const registryProbe = jsonResult(runProvider([registry, "check", "--provider", "pyhanko", "--require"], { status: 0 }));
  assert.equal(registryProbe.providers[0].available, true);
  assert.equal(registryProbe.providers[0].evidence.minimumVersion, "0.35.0");
  assert.equal(registryProbe.providers[0].evidence.maximumVersionExclusive, "0.36.0");
  assert.match(registryProbe.providers[0].evidence.companionVersion, /^0\.31\./);
  assert.equal(registryProbe.providers[0].evidence.companionMaximumVersionExclusive, "0.32.0");
  const plan = jsonResult(runProvider([
    registry, "plan", "--task", "verify-signature", "--provider", "pyhanko", "--strategy", "read-only",
    "--input", source, "--require-provider",
  ], { status: 0 }));
  assert.equal(plan.integration, "shipped-thin-script-external-python");
  assert.equal(plan.silentFallback, false);

  const unsigned = jsonResult(runProvider([provider, "verify", source, "--expected-sha256", sourceHash], { status: 0 }));
  assert.equal(unsigned.schema, "open-office-artifact-tool.pyhanko-verify.v1");
  assert.equal(unsigned.conclusion, "unsigned");
  assert.equal(unsigned.summary.signatureCount, 0);
  assert.equal(unsigned.sourceProtected, true);
  assert.deepEqual(await fs.readFile(source), sourceBytes);
  const unsignedRequired = runProvider([
    provider, "verify", source, "--expected-sha256", sourceHash, "--require-signature",
  ], { status: 2 });
  assert.match(jsonResult(unsignedRequired, "stderr").policyGates.failures[0], /no embedded signatures/);

  const stale = runProvider([
    provider, "verify", source, "--expected-sha256", "0".repeat(64),
  ], { status: 2 });
  assert.match(jsonResult(stale, "stderr").error, /source SHA-256 mismatch/);
  const tooSmall = runProvider([
    provider, "verify", source, "--expected-sha256", sourceHash, "--max-input-bytes", "64",
  ], { status: 2 });
  assert.match(jsonResult(tooSmall, "stderr").error, /outside the 5\.\.64 byte budget/);
  const raisedHardLimit = runProvider([
    provider, "verify", source, "--expected-sha256", sourceHash,
    "--max-input-bytes", String(512 * 1024 * 1024 + 1),
  ], { status: 2 });
  assert.match(jsonResult(raisedHardLimit, "stderr").error, /cannot exceed the hard maximum/);
  const naiveMoment = runProvider([
    provider, "verify", source, "--expected-sha256", sourceHash, "--moment", "2026-07-19T12:00:00",
  ], { status: 2 });
  assert.match(jsonResult(naiveMoment, "stderr").error, /explicit UTC offset/);
  const missingRoot = runProvider([
    provider, "verify", source, "--expected-sha256", sourceHash, "--trust-policy", "explicit-roots",
  ], { status: 2 });
  assert.match(jsonResult(missingRoot, "stderr").error, /requires at least one --trust-root/);
  const implicitTrustGate = runProvider([
    provider, "verify", source, "--expected-sha256", sourceHash, "--require-all-trusted",
  ], { status: 2 });
  assert.match(jsonResult(implicitTrustGate, "stderr").error, /requires --trust-policy explicit-roots/);

  const fixtureBuilder = path.join(tempRoot, "build_signed_fixtures.py");
  await fs.writeFile(fixtureBuilder, String.raw`
from datetime import datetime, timedelta, timezone
from pathlib import Path
import sys

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID
from pyhanko.pdf_utils import generic
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from pyhanko.sign import signers
from pyhanko.sign.fields import MDPPerm

root = Path(sys.argv[1])
key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
name = x509.Name([
    x509.NameAttribute(NameOID.COMMON_NAME, "Open Office Test Root"),
    x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Open Office Artifact Tool"),
    x509.NameAttribute(NameOID.COUNTRY_NAME, "US"),
])
now = datetime.now(timezone.utc)
cert = (
    x509.CertificateBuilder()
    .subject_name(name)
    .issuer_name(name)
    .public_key(key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(now - timedelta(days=1))
    .not_valid_after(now + timedelta(days=3650))
    .add_extension(x509.BasicConstraints(ca=True, path_length=None), critical=True)
    .add_extension(x509.KeyUsage(
        digital_signature=True, content_commitment=True, key_encipherment=False,
        data_encipherment=False, key_agreement=False, key_cert_sign=True,
        crl_sign=True, encipher_only=None, decipher_only=None,
    ), critical=True)
    .sign(key, hashes.SHA256())
)
key_path = root / "key.pem"
cert_path = root / "cert.pem"
key_path.write_bytes(key.private_bytes(
    serialization.Encoding.PEM,
    serialization.PrivateFormat.PKCS8,
    serialization.NoEncryption(),
))
cert_path.write_bytes(cert.public_bytes(serialization.Encoding.PEM))
signer = signers.SimpleSigner.load(key_path, cert_path)

with (root / "source.pdf").open("rb") as source, (root / "signed.pdf").open("wb") as output:
    writer = IncrementalPdfFileWriter(source)
    signers.sign_pdf(
        writer,
        signers.PdfSignatureMetadata(
            field_name="Approval", certify=True, docmdp_permissions=MDPPerm.FILL_FORMS
        ),
        signer=signer,
        output=output,
    )

with (root / "signed.pdf").open("rb") as source, (root / "double-signed.pdf").open("wb") as output:
    writer = IncrementalPdfFileWriter(source)
    signers.sign_pdf(
        writer,
        signers.PdfSignatureMetadata(field_name="Reviewer"),
        signer=signer,
        output=output,
    )

with (root / "signed.pdf").open("rb") as source, (root / "modified.pdf").open("wb") as output:
    writer = IncrementalPdfFileWriter(source)
    writer.set_info(generic.DictionaryObject({
        generic.pdf_name("/Producer"): generic.pdf_string("post-signature validation fixture")
    }))
    writer.write(output)
`, "utf8");
  runProvider([fixtureBuilder, tempRoot], { status: 0 });

  const cert = path.join(tempRoot, "cert.pem");
  const signed = path.join(tempRoot, "signed.pdf");
  const signedBytes = await fs.readFile(signed);
  const signedHash = sha256(signedBytes);
  const cryptographicOnly = jsonResult(runProvider([
    provider, "verify", signed, "--expected-sha256", signedHash,
    "--require-signature", "--require-all-integrity-valid",
  ], { status: 0 }));
  assert.equal(cryptographicOnly.conclusion, "cryptographically-valid-untrusted");
  assert.equal(cryptographicOnly.signatures[0].intact, true);
  assert.equal(cryptographicOnly.signatures[0].cryptographicallyValid, true);
  assert.equal(cryptographicOnly.signatures[0].trusted, false);
  assert.equal(cryptographicOnly.validationPolicy.trustRoots.length, 0);
  assert.equal(cryptographicOnly.networkAllowed, false);

  const trustedArgs = [
    provider, "verify", signed, "--expected-sha256", signedHash,
    "--trust-policy", "explicit-roots", "--trust-root", cert,
    "--revocation-policy", "none", "--require-signature",
    "--require-all-integrity-valid", "--require-all-trusted",
    "--require-docmdp-compliant", "--require-all-bottom-line",
  ];
  const trusted = jsonResult(runProvider(trustedArgs, { status: 0 }));
  assert.equal(trusted.ok, true);
  assert.equal(trusted.conclusion, "valid-under-selected-policy");
  assert.equal(trusted.summary.allBottomLine, true);
  assert.equal(trusted.summary.allDocMDPCompliant, true);
  assert.equal(trusted.signatures[0].coverage, "entire-file");
  assert.equal(trusted.signatures[0].modificationLevel, "none");
  assert.equal(trusted.signatures[0].docMDP.permission, "fill-forms");
  assert.equal(trusted.validationPolicy.trustPolicy, "explicit-roots");
  assert.equal(trusted.validationPolicy.trustRoots.length, 1);
  assert.equal(trusted.validationPolicy.trustRoots[0].certificate.sha256Fingerprint.length, 64);
  assert.equal(trusted.padesProfileConformanceClaimed, false);
  assert.deepEqual(await fs.readFile(signed), signedBytes);

  const doubleSigned = path.join(tempRoot, "double-signed.pdf");
  const doubleBytes = await fs.readFile(doubleSigned);
  const doubleReport = jsonResult(runProvider([
    provider, "verify", doubleSigned, "--expected-sha256", sha256(doubleBytes),
    "--trust-policy", "explicit-roots", "--trust-root", cert,
    "--require-signature", "--require-all-integrity-valid", "--require-all-trusted",
    "--require-docmdp-compliant", "--require-all-bottom-line",
  ], { status: 0 }));
  assert.equal(doubleReport.revisionCount, 3);
  assert.deepEqual(doubleReport.signatures.map((signature) => signature.fieldName), ["Approval", "Reviewer"]);
  assert.deepEqual(doubleReport.signatures.map((signature) => signature.signedRevision), [1, 2]);
  assert.equal(doubleReport.signatures[0].modificationLevel, "form-filling");
  assert.equal(doubleReport.signatures[1].coverage, "entire-file");

  const modified = path.join(tempRoot, "modified.pdf");
  const modifiedBytes = await fs.readFile(modified);
  const modifiedReport = jsonResult(runProvider([
    provider, "verify", modified, "--expected-sha256", sha256(modifiedBytes),
    "--trust-policy", "explicit-roots", "--trust-root", cert,
    "--require-signature", "--require-all-integrity-valid", "--require-all-trusted",
    "--require-docmdp-compliant", "--require-all-bottom-line",
  ], { status: 0 }));
  assert.equal(modifiedReport.summary.hasPostSigningChanges, true);
  assert.equal(modifiedReport.signatures[0].coverage, "entire-revision");
  assert.equal(modifiedReport.signatures[0].modificationLevel, "lta-updates");
  assert.equal(modifiedReport.signatures[0].docMDPCompliant, true);

  const tampered = path.join(tempRoot, "tampered.pdf");
  const tamperedBytes = Buffer.from(signedBytes);
  assert.equal(tamperedBytes[9], 0x25, "expected the second-line PDF comment marker");
  tamperedBytes[10] ^= 1;
  await fs.writeFile(tampered, tamperedBytes);
  const tamperedResult = runProvider([
    provider, "verify", tampered, "--expected-sha256", sha256(tamperedBytes),
    "--trust-policy", "explicit-roots", "--trust-root", cert,
    "--require-signature", "--require-all-integrity-valid",
  ], { status: 2 });
  const tamperedReport = jsonResult(tamperedResult, "stderr");
  assert.equal(tamperedReport.conclusion, "integrity-failure");
  assert.equal(tamperedReport.signatures[0].intact, false);
  assert.match(tamperedReport.policyGates.failures.join("\n"), /not intact and cryptographically valid/);
  assert.deepEqual(await fs.readFile(tampered), tamperedBytes);

  const tamperedWithoutRequestedGate = runProvider([
    provider, "verify", tampered, "--expected-sha256", sha256(tamperedBytes),
    "--trust-policy", "explicit-roots", "--trust-root", cert,
  ], { status: 2 });
  assert.equal(jsonResult(tamperedWithoutRequestedGate, "stderr").conclusion, "integrity-failure");

  const implicitRoot = runProvider([
    provider, "verify", signed, "--expected-sha256", signedHash, "--trust-root", cert,
  ], { status: 2 });
  assert.match(jsonResult(implicitRoot, "stderr").error, /trust is never inferred silently/);
} finally {
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log("pyHanko provider smoke ok");
