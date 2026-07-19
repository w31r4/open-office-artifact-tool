# Inspect, repair, and linearize with qpdf

Use the shipped `scripts/qpdf_provider.py` wrapper when the task needs qpdf's
structural diagnosis, recoverable cross-reference repair, or linearization. It
opens the original PDF directly and never reconstructs it through `PdfArtifact`.
qpdf remains a separately installed Apache-2.0 command-line provider; absence is
an explicit capability error, not a reason to choose another backend.

## Probe and inspect

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/qpdf_provider.py probe
"$PYTHON_BIN" scripts/qpdf_provider.py inspect input.pdf \
  > tmp/pdfs/qpdf-inspect.json
```

Inspection is read-only and bounded. The JSON report binds the absolute source
path, byte count, and SHA-256; qpdf version and executable; clean/warning exit
status; PDF version; page/object/form/attachment/outline counts; encryption and
linearization state; and object-level signature, ByteRange, `/Perms`, DocMDP,
and FieldMDP evidence. Exit code 3 is retained as recoverable warning evidence.
Exit code 2, malformed/over-budget JSON, source drift, an unsupported qpdf
version, or a missing provider fails closed.

qpdf is a structural checker, not a strict PDF-spec, trust, text, accessibility,
or rendering validator. A clean result means only that qpdf found no errors or
warnings. Use pyHanko for signature trust/policy, veraPDF for requested PDF/A or
PDF/UA machine rules, and Poppler for visual truth.

## Repair a recoverable structure

Review the inspect report and copy its `source.sha256` exactly:

```bash
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task repair --provider qpdf --strategy rewrite \
  --input input.pdf --output outputs/repaired.pdf --require-provider
"$PYTHON_BIN" scripts/qpdf_provider.py rewrite \
  input.pdf outputs/repaired.pdf \
  --mode repair \
  --expected-sha256 '<sha256-from-inspect>' \
  > tmp/pdfs/qpdf-repair.json
```

The wrapper copies the exact inspected source into a private transaction,
invokes qpdf only on that snapshot, re-inspects the candidate, requires a clean
output, preserves page/form/attachment/outline counts, re-proves the source
hash, and publishes a distinct output without replacing an existing path. A
damaged input
may have warning status before repair; the promoted output may not.

## Linearize

Linearization is the same full-rewrite transaction with an additional validated
postcondition:

```bash
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task linearize --provider qpdf --strategy rewrite \
  --input input.pdf --output outputs/linearized.pdf --require-provider
"$PYTHON_BIN" scripts/qpdf_provider.py rewrite \
  input.pdf outputs/linearized.pdf \
  --mode linearize \
  --expected-sha256 '<sha256-from-inspect>' \
  > tmp/pdfs/qpdf-linearize.json
```

Object numbers, byte offsets, compression, document IDs, and serialized bytes
may change during either rewrite. They are not stable identities. Linearization
is not optimization of visible content and does not authorize content edits.

## Signatures, encryption, and cleanup boundary

- Any signature field, ByteRange, `/Perms`, DocMDP, or FieldMDP evidence blocks
  rewrite until `--invalidate-signatures` is explicitly supplied after pyHanko
  and policy review. The output report says `signatureInvalidated: true`; it
  never claims that the old signer approved the new bytes. Invalid signature
  fields or appearances may remain in the rewritten PDF; inspect the reported
  output-side signature evidence and use pyHanko to explain or remove them.
- This bounded adapter rejects encrypted rewrite. Password handling, decrypt,
  re-encrypt, and permission changes need a separate explicit policy workflow.
- qpdf repair/linearize does not remove JavaScript, attachments, metadata,
  hidden text, OCR layers, signatures, or sensitive content. Do not call it
  sanitize or redaction. Use the strict PyMuPDF sanitize/residue route.
- `structure-clean` is not a qpdf capability in this Skill. The planned pikepdf
  route remains unavailable until a separately tested active-content adapter is
  shipped.

## Delivery gates

After rewrite, inspect the output afresh, run `pdfinfo`, render every page with
Poppler, compare page count and pixels/geometry against the source, review qpdf
warnings, and bind the exact source/output bytes into the canonical PDF audit.
