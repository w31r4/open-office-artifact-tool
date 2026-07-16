# Edit an existing PDF

Do not route an existing PDF through `PdfArtifact` for mutation. Pass the original file directly to the chosen provider.

## Plan first

```bash
python3 scripts/pdf_provider.py plan \
  --task edit-content \
  --provider pymupdf \
  --strategy rewrite \
  --input input.pdf \
  --output tmp/pdfs/edited.pdf \
  --accept-license agpl
```

## PyMuPDF operations

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
    "type": "rotate_page",
    "page": 2,
    "rotation": 90
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
python3 scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/edited.pdf \
  --strategy rewrite \
  --operations tmp/pdfs/edit-operations.json \
  --accept-license agpl
```

Use `--strategy incremental` only for bounded changes where retaining old revisions is intended. The script copies the original bytes to the destination, requires `Document.can_save_incrementally()`, appends the update, and verifies that the original prefix is byte-identical.

`replace_image` requires an xref observed on the selected page and replaces every use of that image object. Confirm shared-object effects before delivery.

Text replacement in an ordinary PDF is not Word-style reflow. For a short replacement that fits the original geometry, use `replace_text` under `sanitize`; it performs real redaction and a same-box overlay, then runs the full residue gate. If the replacement does not fit, the script fails. For paragraph/page reflow, use a trusted source model or explicitly create a reconstructed new document.

Signed input requires prior signature/DocMDP inspection. Use `--allow-signed` only after the requested operation has been reviewed against the signature policy; validate before and after with pyHanko. Rewrite requires explicit `--invalidate-signatures`.

After editing, compare intended deltas, reopen independently, render every page, and preserve the source file and operation manifest.
