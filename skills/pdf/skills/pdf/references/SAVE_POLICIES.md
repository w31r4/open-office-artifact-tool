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

## `incremental`

The provider copies or opens the original bytes and appends a new PDF revision. The original byte prefix must remain byte-identical.

Use only for bounded operations supported by the selected provider, such as a source-bound single-widget form-field update, visible-only CropBox change, rotation, or signing workflow.

Consequences:

- Old revisions remain in the file by design. Never use this mode for deletion, redaction, privacy scrubbing, or claims that prior content is gone.
- A prior cryptographic signature may remain mathematically verifiable while the new revision is still forbidden by DocMDP or the signer's modification policy. Validate before and after with pyHanko and review the policy.
- The default MuPDF.js primitive permits byte-prefix-verified incremental save only for its bounded unsigned, non-destructive operations. `update_form_field` binds exact source bytes plus one inspected `mupdfFormField` snapshot and is limited to one non-password text/compatible combo/checkbox widget; `set_page_crop` changes only visible `CropBox` and retains hidden content, while `rotate_page` writes only an absolute right-angle `/Rotate` value. These may be incremental but are never redaction claims. It rejects redaction, delete operations, source-bound annotation/link creation or mutation (including `add_text_highlight`), shared or complex form fields, and every signed input in incremental mode; provider support never proves that a signature policy authorizes the change.
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
2. Add and apply real redaction annotations; drawing opaque rectangles is not redaction.
3. Run PyMuPDF scrub with the requested removal policy.
4. Save a full non-incremental rewrite with garbage collection and stream cleanup.
5. Prove the output does not retain the original byte prefix.
6. Run residue scans over raw bytes, decoded objects/streams, extracted text, metadata/XMP, attachments, annotations, widgets, and OCR of image-bearing pages.
7. Run `qpdf --check` when available, signature/conformance checks when relevant, then Poppler-render and inspect every page.
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
