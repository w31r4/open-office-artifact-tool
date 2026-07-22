# Remove active and auxiliary PDF structures with pikepdf

Use the shipped `scripts/pikepdf_provider.py` adapter when an Agent must create
a structurally quieter copy of an imported PDF without rebuilding its visible
pages. The adapter opens the exact original bytes through a separately selected
pikepdf runtime, applies only pikepdf's curated sanitizer operations, saves one
new full-rewrite revision, and validates the selected postconditions.

This is the `structure-clean` route. It is deliberately narrower than strict
`sanitize`: it does not redact page content, erase metadata, clear form values,
flatten XFA, remove comments, scan OCR/image residue, or certify an untrusted
file as safe.

## Resolve and probe

Resolve `task: "structure-clean"` with `provider: "pikepdf"`,
`savePolicy: "rewrite"`, mutation authorization, and signature-invalidation
authorization before invoking the adapter. Select a ready managed pack when one
is published or an explicit `system-only` Python runtime; [provider setup](provider_setup.md)
owns policy and runtime preparation. The npm package never installs Python
packages and has no lifecycle hook, global-pip path, or fallback.

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pdf_provider.py check --provider pikepdf --require
"$PYTHON_BIN" scripts/pikepdf_provider.py probe
```

A missing curated API, invalid interpreter, or incompatible provider is an
explicit capability error. There is no qpdf or PyMuPDF fallback.

## Choose a fixed profile

Only two profiles exist:

- `active-content` removes JavaScript, network/filesystem actions, and
  multimedia/RichMedia/3D execution or payload references. Link and media
  annotations remain in place after their selected actions or payloads are
  defanged, so page geometry stays stable.
- `active-and-auxiliary` additionally removes embedded/associated files,
  thumbnails, embedded search indexes, Web Capture data, private page-piece
  dictionaries, and the PDF portfolio/collection presentation.

The broader profile is still not a public-release guarantee. Both profiles
retain DocumentInfo/XMP metadata, AcroForm fields and values, XFA, annotations,
hidden text/OCR layers, and signature fields or appearances. Choose the strict
PyMuPDF sanitize/redaction workflow when those channels are in scope.

## Inspect the exact source

Compute and retain the source SHA-256 independently, then use the adapter's
bounded read-only inspection. For attacker-chosen inputs, run the whole process
inside a caller-managed OS/container sandbox and select `--caller-isolated`;
the adapter process is not itself a malware boundary.

```bash
SOURCE_SHA="$($PYTHON_BIN -c 'import hashlib,sys; print(hashlib.sha256(open(sys.argv[1], "rb").read()).hexdigest())' input.pdf)"

"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task inspect --provider pikepdf --strategy read-only \
  --input input.pdf --require-provider

"$PYTHON_BIN" scripts/pikepdf_provider.py inspect input.pdf \
  --expected-sha256 "$SOURCE_SHA" \
  --trusted-input \
  > tmp/pdfs/pikepdf-inspect.json
```

Use exactly one of `--trusted-input` or `--caller-isolated`. Inspection works
on a private read-only snapshot, rejects encrypted or parser-warning input,
enforces input/page/object/time/output budgets, and reports the bounded feature
counts plus signature, metadata, XFA, form, annotation, outline, and tagging
evidence. Damaged files route to qpdf repair before a new inspection.

## Create a structure-clean copy

Review signature/DocMDP evidence first. The provider always rewrites and
coalesces prior revisions, so `--invalidate-signatures` is required even when
the inspection currently reports no signature. This is a deliberate policy
acknowledgement, not a statement that the old signer approved the new bytes.

```bash
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task structure-clean --provider pikepdf --strategy rewrite \
  --input input.pdf --output outputs/structure-clean.pdf \
  --invalidate-signatures --require-provider

"$PYTHON_BIN" scripts/pikepdf_provider.py clean \
  input.pdf outputs/structure-clean.pdf \
  --profile active-and-auxiliary \
  --expected-sha256 "$SOURCE_SHA" \
  --trusted-input \
  --invalidate-signatures \
  > tmp/pdfs/pikepdf-structure-clean.json
```

The adapter copies the exact source into a private read-only transaction,
runs pikepdf in a budgeted child process, reopens the candidate, requires every
selected feature category to be absent, and checks that page, annotation, form,
XFA, metadata, tagging, and outline topology did not drift. It then proves the
result is one non-incremental revision, re-proves the source hash, and promotes
the output with a no-replace hard-link transaction. Existing paths, symlinks,
stale hashes, encrypted PDFs, warnings, topology drift, budget failures, or a
late source/output collision leave no promoted output.

## Delivery gates

The provider report is mutation evidence, not the whole delivery audit. Run
independent checks over the exact final bytes:

```bash
"$PYTHON_BIN" scripts/qpdf_provider.py inspect outputs/structure-clean.pdf \
  > tmp/pdfs/qpdf-after-structure-clean.json

pdftoppm -png -r 144 input.pdf tmp/pdfs/source-page
pdftoppm -png -r 144 outputs/structure-clean.pdf tmp/pdfs/output-page

"$PYTHON_BIN" scripts/residue_scan.py outputs/structure-clean.pdf \
  --report tmp/pdfs/structure-clean-residue.json
```

- Compare every Poppler page and explain any visual change. JavaScript can
  legitimately affect viewer rendering, so equality is a gate for the tested
  file, not a universal promise.
- Use `residue_scan.py --require-inert` only when its stricter policy is truly
  intended. It will correctly fail if retained metadata, form values, comments,
  links, hidden text, or other channels remain.
- Run pyHanko when the source had signature evidence. Run veraPDF again when
  PDF/A or PDF/UA matters; a structural rewrite may affect conformance even
  when the visible pages are unchanged.
- Bind the exact source/output hashes, selected profile, provider report,
  qpdf evidence, residue policy, and Poppler review into the canonical PDF
  audit before delivery.

## Product and security boundary

- No password, decryption, re-encryption, incremental save, arbitrary pikepdf
  call, custom plugin, or low-level object mutation is exposed.
- `structure-clean` is not redaction, strict sanitize, malware analysis,
  signature validation, PDF/A or PDF/UA validation, or visual QA.
- Removing JavaScript can break form validation. Removing external access also
  disables ordinary URI links. Removing attachments can discard files that are
  semantically integral to a portfolio or signing workflow.
- XFA, AcroForm values, metadata, comments, hidden text, OCR text, and visible
  content remain deliberately outside this primitive. Route those needs to an
  explicit specialist workflow or fail closed.
