# Imported-PDF security checklist

Treat PDFs as untrusted structured programs.

## Before opening

- Work on a copy in `tmp/pdfs/`; keep the original immutable.
- Record SHA-256, bytes, and file name.
- Set page/object/stream/time limits where the provider exposes them.
- Do not send attacker-chosen PDFs to OCRmyPDF in the host process. The shipped
  adapter requires either `--input-trust trusted` or the caller's explicit
  assertion that the entire process is already isolated; it does not enforce a
  VM/container sandbox and OCR output is not sanitize evidence.
- The shipped pikepdf adapter likewise requires either `--trusted-input` or
  `--caller-isolated`. Its budgeted child process limits time and output but is
  not an OS/container sandbox; do not parse attacker-chosen files in the host
  trust domain merely because the selected operation removes active content.
- Do not execute embedded JavaScript, launch actions, attachments, media, or external links.
- Extract attachments only through a path-confined quarantine primitive with duplicate-name separation, count/byte budgets, and per-file SHA-256; never trust a FileSpec filename or recursively open payloads by default.
- Detect encryption and request the authorized password rather than bypassing controls.

## Before mutation

- Select one provider and one save policy explicitly.
- Probe the requested capability; do not infer support from the package being installed.
- Inspect signatures, signature fields, `/Perms`, and DocMDP before changing bytes.
- Before signing, bind both PDF and PKCS#12 to fresh SHA-256 values; reject symlinks, encryption, unexpected signature counts, an existing certification signature, or a requested visible box outside the inspected unrotated CropBox. Supply the PKCS#12 passphrase only through stdin, never argv, environment, logs, reports, or repository files.
- Inventory forms, annotations, attachments, metadata/XMP, JavaScript/actions, optional content, images, and OCR/hidden text.
- For raster-only term redaction, bind one unrotated page, an exact term, the
  requested Tesseract language/DPI, and an expected image-backed match count.
  Do not convert OCR uncertainty into a broad deletion rectangle.
- Keep source and destination paths different.

## After mutation

- Reopen with an independent parser where practical.
- Compare page count, page boxes, forms/annotations, attachment count, metadata, and signatures against the intended delta.
- Run `pdfinfo`, the shipped `qpdf_provider.py inspect` when qpdf 11+ is installed,
  and `pyhanko_provider.py verify` against the exact source/output SHA-256 when
  signatures are present; provide explicit trust roots and retain integrity,
  trust, revision coverage, difference level, and DocMDP evidence separately.
  Run `verapdf_provider.py validate` against the exact final SHA-256 and one
  explicit built-in profile when conformance is requested; retain the typed
  machine-rule result and separate PDF/UA human-review requirement. Retain qpdf
  warning status rather than suppressing it.
- Render every page from final bytes with Poppler; model/SVG preview is not enough.
- For sanitize/redaction, run the strict residue scan including image OCR and reject any incomplete evidence. For an inert public copy, also pass `--require-inert`; zero sensitive terms are valid only for a scrub-only operation.

## Fail-closed conditions

- Provider or exact capability is missing.
- Requested save strategy is incompatible with the operation.
- Encrypted input cannot be opened with authorized credentials.
- Signed-document policy is unknown or not explicitly accepted.
- Signature validation relies on an implicit system trust store, network fetch,
  stale source hash, unreviewed revocation mode, or a collapsed one-boolean
  interpretation of pyHanko's integrity/trust/difference evidence.
- Signing relies on a stale source/credential hash, an unsupported pyHanko
  runtime, a passphrase outside stdin, an implicit field/DocMDP choice, a
  mismatched signature count, source overwrite, output replacement, a
  non-prefix result, or absent post-sign integrity/DocMDP validation.
- Conformance validation relies on veraPDF automatic profile selection, a
  custom/unbounded profile, stale source bytes, or treats PDF/UA machine rules
  as proof of author intent and real assistive-technology usability.
- OCR relies on an unpinned/unsupported OCRmyPDF or Tesseract version, a missing
  language pack, implicit mode, stale source hash, unacknowledged tagged/
  rasterization/form/signature loss, empty required text, or non-isolated
  attacker-chosen input.
- OCR redaction relies on PyMuPDF outside `>=1.27.2,<1.28`, unsafe/missing
  Tesseract language data, a rotated page, more than 300 dpi or 100 million
  raster pixels, absent/mismatched expected matches, or a hit with less than
  90% native image-placement coverage.
- Structure cleanup relies on pikepdf outside `>=10.10,<10.11`, a custom or
  partial operation list, stale source hash, encrypted/parser-warning input,
  missing trust/isolation declaration, missing signature invalidation, an
  incremental prefix, or a selected feature category remaining after rewrite.
- Output overwrites the input.
- Incremental output does not retain the exact original prefix.
- Sanitized output retains the original prefix, sensitive residues, active actions, attachments, comments, populated form values, personal metadata, links, invisible text, unscanned images, or old revisions.
- Invisible text overlaps visible text, so safe rectangle removal cannot preserve ordinary page content.
- Final page render has clipping, overlap, missing glyphs, blank pages, or unexplained visual changes.
