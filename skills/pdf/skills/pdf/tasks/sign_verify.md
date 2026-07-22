# Sign and verify PDFs

Use pyHanko for PDF signatures. The shipped `scripts/pyhanko_sign_provider.py`
adapter inventories signature fields and adds exactly one local-PKCS#12
approval or certification signature as a bounded incremental revision. The
separate read-only `scripts/pyhanko_provider.py` validates exact bytes and emits
typed integrity, trust, revision, and DocMDP evidence for an Agent. PyMuPDF,
MuPDF.js, pypdf, and qpdf are not signature-trust authorities.

## Resolve and probe the signing runtime

The adapter requires the pyHanko core library, not the separately packaged
`pyhanko` command. First resolve the exact `sign` task through the public
capability API. A signing runtime is installable only through an authorized,
hash-pinned managed pack; otherwise select an already-provisioned
`system-only` Python runtime in [provider setup](provider_setup.md). Do not
repair a missing runtime with `pip`, `uv`, a package manager, or a global
installation command.

```js
import { PdfFile } from "open-office-artifact-tool";
import { PdfProviders } from "open-office-artifact-tool/pdf/providers";

const inspection = await PdfFile.inspectPdf("input.pdf");

let resolution = await PdfProviders.resolve({
  task: "sign",
  provider: "pyhanko",
  inspection,
  savePolicy: "incremental",
  mutationAuthorized: true,
  credentials: ["local-pkcs12"],
  policyPath: ".open-office-artifact-tool/pdf-providers.json",
});
if (resolution.status === "installable") {
  resolution = await PdfProviders.ensure({ resolution, policyPath: ".open-office-artifact-tool/pdf-providers.json" });
}
if (resolution.status !== "ready") throw new Error(resolution.reason.message);
await PdfProviders.probe({ provider: "pyhanko", task: "sign", policyPath: ".open-office-artifact-tool/pdf-providers.json" });
```

After the selected route is ready, run the task-specific probes through the
same configured Python executable:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:?select a ready pyHanko runtime first}"
"$PYTHON_BIN" scripts/pyhanko_sign_provider.py probe
"$PYTHON_BIN" scripts/pyhanko_provider.py probe
"$PYTHON_BIN" scripts/pdf_provider.py check --provider pyhanko --require
```

Neither adapter uses a system trust store, fetches certificates, CRLs, or OCSP
responses, invokes a CLI, or routes to another provider. The signer supports
local PKCS#12 credentials only; TSA, LTV/DSS, PKCS#11, remote signing, and
complete PAdES conformance remain external.

## Inspect and sign one exact source

Hash the source and credential immediately before use. Inspect current fields,
signature count, certification state, revision count, and the selected page's
unrotated CropBox before choosing a field mode:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
SOURCE_SHA256="$(shasum -a 256 input.pdf | awk '{print $1}')"
CREDENTIAL_SHA256="$(shasum -a 256 /secure/signer.p12 | awk '{print $1}')"

"$PYTHON_BIN" scripts/pyhanko_sign_provider.py inspect input.pdf \
  --expected-sha256 "$SOURCE_SHA256" --page-index 0 --trusted-input \
  > tmp/pdfs/signature-inventory.json

# Invisible approval signature. A terminal gets a hidden prompt.
"$PYTHON_BIN" scripts/pyhanko_sign_provider.py sign \
  input.pdf tmp/pdfs/signed.pdf \
  --expected-sha256 "$SOURCE_SHA256" --trusted-input \
  --credential /secure/signer.p12 \
  --credential-sha256 "$CREDENTIAL_SHA256" --passphrase-stdin \
  --field-name Approval --field-mode create-invisible \
  --signature-kind approval --subfilter pades \
  --expected-signature-count 0 \
  > tmp/pdfs/signing-report.json
# Automation pipes stdin directly from its secret manager without staging the
# value in argv, env, or a file.
```

Use `--no-passphrase` only for a deliberately unencrypted PKCS#12. Secrets are
never accepted on argv or through an environment option and are omitted from
the versioned report. The provider rejects symlink credentials, stale hashes,
encryption, source overwrite, output collisions, oversized inputs/outputs,
unsupported runtime versions, and missing trust/isolation declarations.

Field modes are explicit:

- `existing` fills exactly one named empty signature field;
- `create-invisible` creates an invisible field on page 0;
- `create-visible` also requires `--page-index` and an integer
  `--box x1,y1,x2,y2` wholly inside an unrotated inspected CropBox.

A certification signature must be first and requires
`--docmdp-permission no-changes|fill-forms|annotate`. Finalize content and
Poppler visual QA before applying restrictive certification. A later approval
signature requires both the exact `--expected-signature-count` and
`--allow-existing-signatures`. This acknowledges a new revision; it never means
an earlier signer approved it.

The output is a distinct transaction: it preserves the complete source byte
prefix, appends one revision, adds one signature, passes internal integrity and
DocMDP validation, and is promoted without replacement. Re-run qpdf structure
inspection, explicit-root validation, and Poppler rendering before delivery.

## Validate one exact source

Record a fresh SHA-256, then choose one of two trust policies:

```bash
SOURCE_SHA256="$(shasum -a 256 input.pdf | awk '{print $1}')"
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"

# Integrity/difference evidence only. Trust is deliberately reported false.
"$PYTHON_BIN" scripts/pyhanko_provider.py verify input.pdf \
  --expected-sha256 "$SOURCE_SHA256" \
  --trust-policy cryptographic-only \
  --require-signature \
  --require-all-integrity-valid \
  > tmp/pdfs/signature-cryptographic.json

# Delivery gate against caller-supplied trust roots.
"$PYTHON_BIN" scripts/pyhanko_provider.py verify input.pdf \
  --expected-sha256 "$SOURCE_SHA256" \
  --trust-policy explicit-roots \
  --trust-root /trusted/root-ca.pem \
  --other-cert /trusted/intermediate.pem \
  --revocation-policy hard-fail \
  --require-signature \
  --require-all-integrity-valid \
  --require-all-trusted \
  --require-docmdp-compliant \
  --require-all-bottom-line \
  > tmp/pdfs/signature-validation.json
```

`--moment` accepts an ISO 8601 timestamp with an explicit UTC offset. The
revocation policies are `none`, `soft-fail`, `hard-fail`, and `require`; choose
one deliberately. Network fetching is always disabled, so strict revocation
policies can succeed only when pyHanko already has adequate embedded/local
evidence. Do not weaken the policy merely to turn a report green.

The versioned `open-office-artifact-tool.pyhanko-verify.v1` report keeps these
facts separate for every signature:

- signed revision and ByteRange coverage;
- byte integrity and cryptographic validity;
- certificate path trust under the exact supplied roots, moment, and
  revocation policy;
- signer certificate identity, digest, and signature mechanism;
- timestamp evidence;
- difference-analysis modification level and changed form fields;
- DocMDP/FieldMDP constraints and DocMDP compliance;
- seed-value status and a policy-specific bottom line.

The adapter validates a private source snapshot under hard input, signature,
certificate, subprocess-time, stdout, and stderr budgets, then proves the
source and trust inputs did not change. An unsigned file can be inventoried, but
`--require-signature` fails it. A stale source hash, implicit/system trust,
encrypted input, unsupported pyHanko version, incomplete signature validation,
cryptographic/DocMDP failure, or unmet required gate fails closed with
structured JSON and no fallback. The configurable byte/time limits can be
lowered for a task but never raised above the adapter's hard maxima.

The report does **not** claim complete PAdES profile conformance. `trusted` means
only that pyHanko accepted the certificate path under the recorded validation
policy. An intact older ByteRange proves the signed revision is unchanged; it
does not mean the signer approved arbitrary later revisions. Review
`coverage`, `modificationLevel`, `docMDPCompliant`, timestamps, and every
signature in revision order.

## Capabilities outside the shipped signer

Use an explicit external pyHanko workflow for PKCS#11/HSM credentials, remote
signing services, timestamp authorities, revocation-material embedding, LTV/DSS
updates, or a claimed PAdES profile. Keep private keys, tokens, PINs, and
passphrases outside scripts, logs, shell history, reports, and repository files.
Review pyHanko's official
[signing](https://docs.pyhanko.eu/en/latest/cli-guide/signing.html) and
[validation](https://docs.pyhanko.eu/en/latest/lib-guide/validation.html)
documentation.

After any later incremental form, annotation, DSS, or timestamp update, run the
typed validator again against the new exact bytes. A material content rewrite
normally belongs in a new version that is signed again.
