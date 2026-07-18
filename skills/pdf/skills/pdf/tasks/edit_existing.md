# Edit an existing PDF

Do not route an existing PDF through `PdfArtifact` for mutation. Pass the original file directly to the chosen provider.

## Mandatory preflight

The default MuPDF.js path probes and inspects before mutation:

```bash
node scripts/mupdf.mjs probe
node scripts/mupdf.mjs inspect input.pdf
```

Its typed operations are `add_text_annotation`, text/choice/checkbox `fill_form`, `delete_page`, source-bound `delete_annotation`, complete `rearrange_pages`, visible-only `set_page_crop`, absolute-quarter-turn `rotate_page`, `set_metadata`, `delete_embedded_file`, source-bound `delete_link`, `redact_text`, and `redact_rect`. Run with one explicit save policy:

```bash
node scripts/mupdf.mjs edit input.pdf tmp/pdfs/edit-operations.json tmp/pdfs/edited.pdf \
  --save-policy rewrite
```

The CLI refuses source overwrite, writes atomically, and rejects incremental redaction/deletion and signed-PDF incremental edits. Unsupported operations do not route elsewhere.

## Visible page crop

Use `set_page_crop` only when the task is to change the visible page window without deleting underlying content. Inspect first and use the raw unrotated `MediaBox`/`CropBox` coordinates returned for the target page:

```json
[
  { "type": "set_page_crop", "page": 1, "bbox": [72, 72, 468, 648] }
]
```

The box must be fully inside the inspected `MediaBox`; rotated pages fail closed and need an explicitly selected specialist route. This operation writes only `CropBox`, retains content outside the crop, and may use unsigned `incremental` save. It is never a redaction, deletion, or sanitize substitute.

## Page rotation

Use `rotate_page` when the task is only to change the viewer orientation of one
existing page. It writes an absolute normalized `/Rotate` value and does not
transform, reflow, or remove content:

```json
[
  { "type": "rotate_page", "page": 2, "rotation": 90 }
]
```

`rotation` must be exactly `0`, `90`, `180`, or `270`; inspect before and after
to retain the prior value and prove the requested orientation. This bounded
unsigned operation may use `incremental` save, subject to the same source-prefix
and signature refusal rules. It is not a substitute for rotated-coordinate text
or image editing; route those tasks explicitly to the specialist provider.

## Delete one imported annotation

First run `inspect` on the exact input. Select one `mupdfAnnotation` record by
semantic facts such as page, type, contents, author, and rectangle; never use
its array position. Copy its `id`, the inspection `summary.sourceSha256`, and a
snapshot into a rewrite operation:

```json
[
  {
    "type": "delete_annotation",
    "page": 2,
    "annotationId": "mupdf-annotation-2-42",
    "sourceSha256": "<inspect summary sourceSha256>",
    "expected": {
      "type": "Text",
      "contents": "Resolved in board review",
      "rect": [72, 128, 20, 20]
    }
  }
]
```

The source SHA-256, page encoded in the locator, xref, and every supplied
snapshot field must match before mutation. `delete_annotation` is a destructive
rewrite-only operation; incremental output is refused. Its locator is bound to
the inspected source bytes rather than a persistent document identity, so
re-inspect the output before any later annotation operation. This prevents a
rewritten PDF's xref reuse from silently targeting a different annotation.

## Delete one imported link

Select one `mupdfLink` record from the exact source inspection by page, URL,
rectangle, and externality. Never pass a mutable link-array index or URL by
itself. Copy its locator, source hash, and snapshot into a rewrite operation:

```json
[
  {
    "type": "delete_link",
    "page": 2,
    "linkId": "mupdf-link-2-<inspect fingerprint>",
    "sourceSha256": "<inspect summary sourceSha256>",
    "expected": {
      "url": "https://example.com/obsolete-policy",
      "bbox": [72, 128, 160, 18],
      "external": true
    }
  }
]
```

`delete_link` verifies every supplied fact and deletes only one unique source
link. It is rewrite-only. The link fingerprint is derived from source-visible
page/URL/rectangle/external facts, not a durable PDF object ID; a duplicate
fingerprint or any later output requires a fresh inspection and fails closed
instead of selecting by order.

## Optional PyMuPDF specialist path

For a capability outside the JavaScript contract, run the exact specialist adapter probe and route plan before any mutation. Do not defer either command until audit generation.

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pymupdf_edit.py probe --accept-license agpl
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task edit-content \
  --provider pymupdf \
  --strategy rewrite \
  --input input.pdf \
  --output tmp/pdfs/edited.pdf \
  --accept-license agpl \
  --require-provider
```

For `replace_text` with the required `sanitize` policy, plan `--task redact --strategy sanitize --invalidate-signatures` instead. The adapter probe proves that `replace_text` is in the installed operation surface; the plan binds provider, save policy, source, destination, license, availability, and signature-invalidating acknowledgement. If either fails, stop before `pymupdf_edit.py edit`.

## Specialist operations

Prepare an operation list:

```json
[
  {
    "type": "insert_textbox",
    "page": 1,
    "rect": [72, 640, 540, 700],
    "text": "Reviewed and approved",
    "font_size": 12,
    "font_name": "helv",
    "color": [0.06, 0.36, 0.42]
  },
  {
    "type": "insert_image",
    "page": 1,
    "rect": [450, 620, 540, 710],
    "path": "tmp/pdfs/approved-mark.png",
    "keep_proportion": true
  }
]
```

Then run one explicit save policy:

```bash
"$PYTHON_BIN" scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/edited.pdf \
  --strategy rewrite \
  --operations tmp/pdfs/edit-operations.json \
  --accept-license agpl
```

Use `--strategy incremental` only for bounded changes where retaining old revisions is intended. The script copies the original bytes to the destination, requires `Document.can_save_incrementally()`, appends the update, and verifies that the original prefix is byte-identical.

`replace_image` requires an xref observed on the selected page and replaces every use of that image object. Confirm shared-object effects before delivery.

Text replacement in an ordinary PDF is not Word-style reflow. For a short replacement that fits the original geometry, use `replace_text` under `sanitize`; it requires each match to resolve to one horizontal source span, preserves its baseline and default font/size/color, performs real redaction and a same-box overlay, then runs the full residue gate. The fit check allows only a fixed sub-millipoint numerical tolerance for provider/search-box float quantization and reports the source/output style, measured width, overflow, baseline, and tolerance in `operations[].fitChecks`; it is not user-configurable layout slack. Cross-span/rotated text or replacement beyond that bound fails closed. For paragraph/page reflow, use a trusted source model or explicitly create a reconstructed new document.

Signed input requires prior signature/DocMDP inspection. Use `--allow-signed` only after the requested operation has been reviewed against the signature policy; validate before and after with pyHanko. Rewrite requires explicit `--invalidate-signatures`.

After editing, compare intended deltas, reopen independently, render every page, and preserve the source file and operation manifest.

Write the canonical [`open-office-artifact-tool.pdf-audit.v1`](../references/AUDIT_SCHEMA.md) envelope and validate it against the exact delivered bytes:

```bash
python3 scripts/pdf_audit.py validate outputs/audit.json \
  --source input.pdf --artifact outputs/edited.pdf \
  --require-operation replace_text
```
