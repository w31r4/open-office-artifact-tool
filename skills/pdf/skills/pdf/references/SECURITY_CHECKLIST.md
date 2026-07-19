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
- Conformance validation relies on veraPDF automatic profile selection, a
  custom/unbounded profile, stale source bytes, or treats PDF/UA machine rules
  as proof of author intent and real assistive-technology usability.
- Output overwrites the input.
- Incremental output does not retain the exact original prefix.
- Sanitized output retains the original prefix, sensitive residues, active actions, attachments, comments, populated form values, personal metadata, links, invisible text, unscanned images, or old revisions.
- Invisible text overlaps visible text, so safe rectangle removal cannot preserve ordinary page content.
- Final page render has clipping, overlap, missing glyphs, blank pages, or unexplained visual changes.
