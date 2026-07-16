# PDF provider matrix

This Skill is a capability router, not a single PDF backend. Select one provider before touching a file. Never catch a provider failure and silently retry through another provider.

## Routing matrix

| Provider | Primary role | Input rule | Save policy | Important boundary |
| --- | --- | --- | --- | --- |
| `open-office-artifact-tool` `PdfArtifact` / `PdfFile` | Greenfield semantic authoring, tagged structure, reading order, inspect/verify, modeled preview | Create a new artifact or reopen a package-generated model envelope | `rewrite` | Never use arbitrary-PDF model reconstruction as a fidelity-preserving edit path. |
| ReportLab | Greenfield visual/layout-oriented PDF generation | New document only | `rewrite` | Does not inherit the `PdfArtifact` tagged/reading-order contract. Verify accessibility separately. |
| pdfplumber | Read-only text, word geometry, table, line, and rectangle extraction | Open original PDF directly | `read-only` | Extraction is evidence, not layout fidelity or an edit representation. |
| pypdf | Read/inspect, merge/split/stamp, basic AcroForm and annotation operations | Open original PDF directly | `rewrite` or explicit `incremental` | Inspect signatures/DocMDP first. Incremental preserves prior bytes but does not prove a permitted signed-document change. |
| PyMuPDF | Advanced imported-PDF page/content/font edits, annotations/widgets, redaction, scrub, and native object inspection | Open the original file/bytes directly | explicit `rewrite`, `incremental`, or `sanitize` | Project-approved optional AGPL provider. It is not the digital-signature authority and must fail closed when a requested capability is absent. |
| Poppler (`pdfinfo`, `pdftoppm`) | Independent page-count/file evidence and final native raster QA | Read final PDF bytes | `read-only` | A renderer, not an editor or conformance validator. |
| qpdf | Structural checks, JSON/QDF inspection, recovery, and content-preserving low-level rewrites | Open original PDF directly | `read-only` or `rewrite` | Not a renderer, text extractor, or full PDF-spec conformance checker. Review recovery warnings. |
| pikepdf | Python access to qpdf for structure, encryption, attachments, and active-content cleanup | Open original PDF directly | `rewrite` | Planned provider only: the registry can probe it, but this release ships no pikepdf mutation adapter. |
| pyHanko | PDF signature fields, signing, timestamps/LTV workflows, and signature validation | Open original PDF directly | normally `incremental` or `read-only` | Trust roots, revocation policy, DocMDP, and post-signature changes require explicit configuration. |
| veraPDF | PDF/A and PDF/UA machine-verifiable validation | Read final PDF bytes | `read-only` | PDF/UA also has human checkpoints; a green report is not full accessibility certification. |
| OCRmyPDF / Tesseract | Scanned-page OCR and strict image residue evidence | Open original or final PDF according to the task | `rewrite` or `read-only` | Planned OCR adapter; Tesseract is currently consumed only through PyMuPDF's strict OCR gate when separately installed. |

## Mandatory routing rules

1. For an existing PDF, preserve the original bytes and pass them directly to the selected provider. Do not import through PDF.js or `PdfArtifact`, export a reconstructed model, and call that a faithful edit.
2. Declare `rewrite`, `incremental`, or `sanitize` before mutation. Read [save policies](SAVE_POLICIES.md).
3. Probe the exact provider and capability before editing. Missing provider, unsupported operation, encrypted input, signature restriction, or unsafe save mode is an error.
4. Keep input and output paths distinct. Never overwrite the source until the output passes semantic/file checks and final rendering review.
5. Inspect signatures and DocMDP constraints before any mutation. `incremental` describes byte layout, not authorization under a signature policy.
6. Redaction is always `sanitize`: apply redactions, scrub, fully rewrite, scan residue, then render every page. Incremental redaction is forbidden.

## Dependency and license record

These providers are optional external tools for the Skill and are not bundled in the npm package.

- ReportLab: official [PDF generation documentation](https://docs.reportlab.com/reportlab/userguide/).
- pdfplumber: project [repository and MIT license](https://github.com/jsvine/pdfplumber).
- pypdf: official [user/API documentation](https://pypdf.readthedocs.io/en/latest/) and repository license notices.
- PyMuPDF: official [documentation](https://pymupdf.readthedocs.io/en/latest/) states GNU AGPL or a commercial license. This project records the user's explicit acceptance of the AGPL provider path. Install and redistribute it only under terms applicable to the deployment.
- qpdf: official [manual](https://qpdf.readthedocs.io/en/stable/) and Apache-2.0 project repository.
- pikepdf: optional qpdf-based Python provider; this release probes availability but does not ship a mutation wrapper.
- pyHanko: official [signing and validation documentation](https://docs.pyhanko.eu/en/stable/).
- veraPDF: official [PDF/A and PDF/UA validation documentation](https://docs.verapdf.org/validation/).
- OCRmyPDF and Tesseract: optional OCR tools; no OCRmyPDF adapter is shipped in this release.
- Poppler: separately installed command-line renderer; retain its applicable license notices.

The `scripts/pdf_provider.py check` command reports availability without selecting a substitute.
