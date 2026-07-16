# Imported-PDF security checklist

Treat PDFs as untrusted structured programs.

## Before opening

- Work on a copy in `tmp/pdfs/`; keep the original immutable.
- Record SHA-256, bytes, and file name.
- Set page/object/stream/time limits where the provider exposes them.
- Do not execute embedded JavaScript, launch actions, attachments, media, or external links.
- Extract attachments only through a path-confined quarantine primitive with duplicate-name separation, count/byte budgets, and per-file SHA-256; never trust a FileSpec filename or recursively open payloads by default.
- Detect encryption and request the authorized password rather than bypassing controls.

## Before mutation

- Select one provider and one save policy explicitly.
- Probe the requested capability; do not infer support from the package being installed.
- Inspect signatures, signature fields, `/Perms`, and DocMDP before changing bytes.
- Inventory forms, annotations, attachments, metadata/XMP, JavaScript/actions, optional content, images, and OCR/hidden text.
- Keep source and destination paths different.

## After mutation

- Reopen with an independent parser where practical.
- Compare page count, page boxes, forms/annotations, attachment count, metadata, and signatures against the intended delta.
- Run `pdfinfo`, `qpdf --check` where installed, and applicable pyHanko/veraPDF validation.
- Render every page from final bytes with Poppler; model/SVG preview is not enough.
- For sanitize/redaction, run the strict residue scan including image OCR and reject any incomplete evidence. For an inert public copy, also pass `--require-inert`; zero sensitive terms are valid only for a scrub-only operation.

## Fail-closed conditions

- Provider or exact capability is missing.
- Requested save strategy is incompatible with the operation.
- Encrypted input cannot be opened with authorized credentials.
- Signed-document policy is unknown or not explicitly accepted.
- Output overwrites the input.
- Incremental output does not retain the exact original prefix.
- Sanitized output retains the original prefix, sensitive residues, active actions, attachments, comments, populated form values, personal metadata, links, invisible text, unscanned images, or old revisions.
- Invisible text overlaps visible text, so safe rectangle removal cannot preserve ordinary page content.
- Final page render has clipping, overlap, missing glyphs, blank pages, or unexplained visual changes.
