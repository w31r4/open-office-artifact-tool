# Sign and verify PDFs

Use pyHanko for PDF signatures. The shipped `scripts/pyhanko_provider.py` adapter
is deliberately read-only: it validates the exact inspected bytes and emits
typed evidence for an Agent. Signature creation, private-key access, timestamp
authorities, and LTV updates remain explicit pyHanko workflows outside that
adapter. PyMuPDF, MuPDF.js, pypdf, and qpdf are not signature-trust authorities.

## Install and probe the validation runtime

The adapter requires the pyHanko core library, not the separately packaged
`pyhanko` command. Install the validated version range into an explicit Python
environment:

```bash
uv venv .venv-pdf
uv pip install --python .venv-pdf/bin/python \
  'pyHanko>=0.35.0,<0.36.0' 'pyhanko-certvalidator>=0.31.0,<0.32.0'
export OPEN_OFFICE_PDF_PROVIDER_PYTHON="$PWD/.venv-pdf/bin/python"
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pyhanko_provider.py probe
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pdf_provider.py check \
  --provider pyhanko --require
```

No lifecycle hook installs this dependency. The adapter does not use a system
trust store, fetch certificates, CRLs, or OCSP responses, invoke a CLI, mutate
the PDF, or route to another provider.

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

## Create or add signatures

Install the separate CLI package only when the task truly needs command-line
signing:

```bash
uv pip install --python .venv-pdf/bin/python 'pyhanko-cli>=0.4.0,<0.5.0'

pyhanko sign addfields \
  --field '1/72,72,260,120/Sig1' \
  input.pdf tmp/pdfs/with-signature-field.pdf

pyhanko sign addsig --field Sig1 pemder \
  --key /secure/key.pem \
  --cert /secure/cert.pem \
  input.pdf output.pdf
```

Keep private keys, tokens, PINs, and passphrases outside scripts, logs, shell
history, and repository files. Inspect existing signatures, `/Perms`, DocMDP,
FieldMDP, and field locks before signing. Finalize content and visual QA before
applying a restrictive certification signature. Use PKCS#12, PKCS#11,
timestamp, revocation, and PAdES/LTV options only after reviewing pyHanko's
official [signing](https://docs.pyhanko.eu/en/latest/cli-guide/signing.html) and
[validation](https://docs.pyhanko.eu/en/latest/lib-guide/validation.html)
documentation.

After any later incremental form, annotation, DSS, or timestamp update, run the
typed validator again against the new exact bytes. A material content rewrite
normally belongs in a new version that is signed again.
