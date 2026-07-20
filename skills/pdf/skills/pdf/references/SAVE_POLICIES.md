# PDF save policies

Choose the policy before mutation. The policy is part of the requested operation, not an implementation detail.

## `rewrite`

The provider reads the original file and writes a new complete PDF.

Use for:

- greenfield output;
- page insertion/removal/reordering;
- merge, stamp, structural repair, or broad content changes;
- normalization where preserving previous revisions is not required.

Consequences:

- Original incremental revisions are not intentionally retained.
- Existing signatures will normally become invalid or be removed. Require explicit acknowledgement before rewriting a signed file.
- A rewrite is not automatically a security scrub: unreferenced objects, attachments, XMP, JavaScript, hidden text, and image pixels must be addressed explicitly when sensitive data is involved.
- The shipped qpdf adapter permits only source-hash-bound `repair` or `linearize` full rewrites, requires a clean re-inspection and stable page/form/attachment/outline counts, rejects encrypted inputs, and requires explicit signature invalidation when signature-policy evidence exists.
- The shipped pikepdf adapter permits only source-hash-bound `active-content` or `active-and-auxiliary` structure cleanup. It works on a private read-only snapshot, requires caller trust/isolation plus explicit signature invalidation, preserves page/annotation/form/XFA/metadata/tag/outline topology, and proves one non-incremental output revision. This is a `rewrite`, not the stricter `sanitize` policy: metadata, form values, XFA, comments, hidden/OCR text, and visible content remain outside its scope.
- The shipped OCRmyPDF adapter is rewrite-only: it requires the exact source
  hash, proves the output does not retain the complete source prefix, fixes
  standard-PDF/O0/one-job settings, and atomically publishes only after qpdf
  topology and Poppler text gates. This is searchable-layer generation, not
  sanitize; `redo`/`force`, Tagged PDFs, forms, and signatures have separate
  explicit loss acknowledgements.

## `incremental`

The provider copies or opens the original bytes and appends a new PDF revision. The original byte prefix must remain byte-identical.

Use only for bounded operations supported by the selected provider, such as a source-bound single-widget form-field update, visible-only CropBox change, rotation, or signing workflow.

Consequences:

- Old revisions remain in the file by design. Never use this mode for deletion, redaction, privacy scrubbing, or claims that prior content is gone.
- A prior cryptographic signature may remain mathematically verifiable while the new revision is still forbidden by DocMDP or the signer's modification policy. Validate before and after with pyHanko and review the policy.
- The shipped pyHanko signer is incremental-only. It binds exact source and PKCS#12 hashes, expects the current signature count, preserves the complete source prefix, adds exactly one approval or first-document certification signature, validates all resulting signatures for integrity/DocMDP compliance, and publishes to a distinct path without replacement. A countersignature requires explicit acknowledgement and never implies an earlier signer approved the new revision.
- The default MuPDF.js primitive permits byte-prefix-verified incremental save only for its bounded unsigned, non-destructive operations. `update_form_field` binds exact source bytes plus one inspected `mupdfFormField` snapshot and is limited to one non-password text/compatible combo/checkbox widget; `set_page_crop` changes only visible `CropBox` and retains hidden content, while `rotate_page` writes only an absolute right-angle `/Rotate` value. These may be incremental but are never redaction claims. It rejects redaction, page duplication, delete operations, source-bound annotation/link creation or mutation (including `add_text_highlight`), shared or complex form fields, and every signed input in incremental mode; provider support never proves that a signature policy authorizes the change. The bounded `duplicate_page` primitive is a source-bound, single-operation full rewrite because a page graft must publish one coherent page/object graph and invalidates current page locators.
- Optional specialist providers may expose other incremental workflows, but they must prove the identical source prefix and satisfy the same signature-policy gate independently.

## `sanitize`

The provider applies destructive privacy/security changes, scrubs active and hidden structures, and performs a full rewrite with old revisions removed.

Required for:

- true redaction;
- metadata/XMP removal;
- attachment, JavaScript, hidden/OCR text, annotation-response, or thumbnail removal;
- high-trust privacy delivery.

Mandatory sequence:

1. Record source SHA-256 and inspect encryption, signatures, DocMDP, forms, annotations, attachments, metadata/XMP, images, and page count.
2. Add and apply real redaction annotations; drawing opaque rectangles is not redaction. Raster-only term selection must use the typed `redact_ocr_text` preflight with an exact page/expected-rotation/term/expected-match contract and native image-placement evidence. Any temporary OCR orientation normalization must restore the original `/Rotate` before mutation continues.
3. Run PyMuPDF scrub with the requested removal policy.
4. Save a full non-incremental rewrite with garbage collection and stream cleanup.
5. Prove the output does not retain the original byte prefix.
6. Run residue scans over raw bytes, decoded objects/streams, extracted text, metadata/XMP, attachments, annotations, widgets, and OCR of image-bearing pages.
7. Run `qpdf_provider.py inspect` when qpdf 11+ is available, signature/conformance checks when relevant, then Poppler-render and inspect every page. Do not call qpdf repair a sanitize step.
8. If any scanner/provider is unavailable, any sensitive token remains, or OCR cannot be performed on an image-bearing page, fail closed and do not deliver.

Sanitization invalidates existing signatures. Require explicit `--invalidate-signatures` acknowledgement.

## Source/output manifest

For every mutation record:

- source and output absolute paths;
- SHA-256 and byte counts;
- selected provider and version;
- save policy;
- capability probe result;
- signature/DocMDP findings;
- operation summary;
- semantic, structure, residue, and render QA results.

Do not overwrite the original during work. Promote the output only after every required gate passes.
