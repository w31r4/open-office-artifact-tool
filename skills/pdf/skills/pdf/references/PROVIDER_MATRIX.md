# PDF provider matrix

This Skill is a capability router, not a single PDF backend. Select one provider before touching a file. Never catch a provider failure and silently retry through another provider.

## Routing matrix

| Provider | Primary role | Input rule | Save policy | Important boundary |
| --- | --- | --- | --- | --- |
| `PdfArtifact` | Greenfield semantic authoring, tagged structure, reading order, inspect/verify, modeled preview | Create a new artifact or reopen a package-generated model envelope | `rewrite` | Never use arbitrary-PDF model reconstruction as a fidelity-preserving edit path. |
| `PdfFile` + required MuPDF.js | Default arbitrary-PDF parsing, native inspect including raw page boxes/rotation, structured text/image/link evidence plus source-bound widget/form-field snapshots, PNG/JPEG render, bounded annotation/form/page/metadata/link edits, visible CropBox changes, absolute page rotation, and rewrite redaction | Open original bytes/path directly through the package or `scripts/mupdf.mjs` | `read-only`, `rewrite`, or explicit `incremental` | Runtime-lazy but required npm dependency. `update_form_field` binds exact source bytes plus one `mupdfFormField` snapshot and permits only a single non-password text widget, compatible single combo, or checkbox; shared/radio/list/multi-select/password/mismatched-export fields route explicitly to pypdf. It may be incremental only on unsigned input. `set_page_crop` is unrotated-only, retains hidden content, and is never redaction; `rotate_page` writes only an absolute right-angle `/Rotate` value. Incremental redaction/deletion and signed-PDF incremental edits are rejected. Rewrite redaction is not full sanitize. General reflow, complex image replacement, signature trust, and strict scrub remain outside this contract. |
| ReportLab | Greenfield visual/layout-oriented PDF generation | New document only | `rewrite` | Does not inherit the `PdfArtifact` tagged/reading-order contract. Verify accessibility separately. |
| pdfplumber | Read-only text, word geometry, table, line, and rectangle extraction | Open original PDF directly | `read-only` | Extraction is evidence, not layout fidelity or an edit representation. |
| pypdf | Read/inspect, path-safe read-only attachment quarantine, complete-source merge/reorder/selective stamp, basic AcroForm and annotation operations | Open original PDF directly | `read-only`, `rewrite`, or explicit `incremental` | The shipped merge manifest selects every source page exactly once and preserves resolvable navigation; ambiguous collisions or unsupported geometry fail closed. Inspect signatures/DocMDP before mutation; incremental preserves prior bytes but does not prove a permitted signed-document change. |
| PyMuPDF | Specialist strict scrub/residue/OCR path and retained high-level operations not yet covered by MuPDF.js | Open the original file/bytes directly | explicit `rewrite`, `incremental`, or `sanitize` | Optional external Python provider. It is not the digital-signature authority and must fail closed when a requested capability is absent. |
| Poppler (`pdfinfo`, `pdftoppm`) | Independent page-count/file evidence and final native raster QA | Read final PDF bytes | `read-only` | A renderer, not an editor or conformance validator. |
| qpdf | Structural checks, JSON/QDF inspection, recovery, and content-preserving low-level rewrites | Open original PDF directly | `read-only` or `rewrite` | Not a renderer, text extractor, or full PDF-spec conformance checker. Review recovery warnings. |
| pikepdf | Python access to qpdf for structure, encryption, attachments, and active-content cleanup | Open original PDF directly | `rewrite` | Planned provider only: the registry can probe it, but this release ships no pikepdf mutation adapter. |
| pyHanko | PDF signature fields, signing, timestamps/LTV workflows, and signature validation | Open original PDF directly | normally `incremental` or `read-only` | Trust roots, revocation policy, DocMDP, and post-signature changes require explicit configuration. |
| veraPDF | PDF/A and PDF/UA machine-verifiable validation | Read final PDF bytes | `read-only` | PDF/UA also has human checkpoints; a green report is not full accessibility certification. |
| OCRmyPDF / Tesseract | Scanned-page OCR and strict image residue evidence | Open original or final PDF according to the task | `rewrite` or `read-only` | Planned OCR adapter; Tesseract is currently consumed only through PyMuPDF's strict OCR gate when separately installed. |

## Mandatory routing rules

1. For an existing PDF, preserve the original bytes and pass them directly to the selected provider. Do not import through PDF.js or `PdfArtifact`, export a reconstructed model, and call that a faithful edit.
2. Declare `read-only`, `rewrite`, `incremental`, or `sanitize` before the operation. Read [save policies](SAVE_POLICIES.md).
3. Probe the exact provider and capability before editing. Missing provider, unsupported operation, encrypted input, signature restriction, or unsafe save mode is an error.
4. Keep input and output paths distinct. Never overwrite the source until the output passes semantic/file checks and final rendering review.
5. Inspect signatures and DocMDP constraints before any mutation. `incremental` describes byte layout, not authorization under a signature policy.
6. Any redaction or delete operation rejects `incremental` because prior revisions retain the original content. A bounded MuPDF.js rewrite removes matched page content from the rewritten revision but is not a full sanitize claim. High-trust redaction uses `sanitize`: apply redactions, scrub, fully rewrite, scan residue, then render every page.

## Dependency and license record

MuPDF.js is a required direct dependency resolved by a normal npm installation and loaded only on the first PDF operation. It remains in its own dependency tarball; there is no lifecycle hook or standalone downloader. All other providers in this matrix are optional external tools installed and licensed separately.

- ReportLab: official [PDF generation documentation](https://docs.reportlab.com/reportlab/userguide/).
- pdfplumber: project [repository and MIT license](https://github.com/jsvine/pdfplumber).
- pypdf: official [user/API documentation](https://pypdf.readthedocs.io/en/latest/) and repository license notices.
- MuPDF.js: official `mupdf@1.28.0`, required by the package under GNU AGPL-3.0-or-later.
- PyMuPDF: official [documentation](https://pymupdf.readthedocs.io/en/latest/) states GNU AGPL or a commercial license. Install and redistribute it only under terms applicable to the deployment.
- qpdf: official [manual](https://qpdf.readthedocs.io/en/stable/) and Apache-2.0 project repository.
- pikepdf: optional qpdf-based Python provider; this release probes availability but does not ship a mutation wrapper.
- pyHanko: official [signing and validation documentation](https://docs.pyhanko.eu/en/stable/).
- veraPDF: official [PDF/A and PDF/UA validation documentation](https://docs.verapdf.org/validation/).
- OCRmyPDF and Tesseract: optional OCR tools; no OCRmyPDF adapter is shipped in this release.
- Poppler: separately installed command-line renderer; retain its applicable license notices.

The `scripts/pdf_provider.py check` command reports availability without selecting a substitute.
