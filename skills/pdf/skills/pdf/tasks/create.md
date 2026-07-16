# Create a new PDF

Choose the authoring contract first.

## Tagged semantic document

Use `PdfArtifact` / `PdfFile` for greenfield content that needs explicit headings, Table/TR/TH/TD structure, figure alternative text, reading order, inspect/resolve, and semantic verification.

Follow [the API quick start](../artifact_tool/API_QUICK_START.md) and the [public API example](../examples/public-api-end-to-end.mjs).
For a six-page CJK board report with running artifacts, meaningful links, a constrained cross-page logical table, Poppler review, and separated modeled/veraPDF/human evidence, use [the accessible board report example](../examples/accessible-board-report.mjs).

Important model behavior:

- Non-empty `page.text` is painted as body/title text and contributes an implicit H1 in the current writer. Do not add another H1 unless two top-level headings are intended.
- `${page.id}/text` is a reading-order target only when `page.text` is non-empty. Omit it for an empty page body.
- `addText(...)` is positioned single-line text and does not wrap. Use `addFlowText(...)` for paragraphs and automatic pagination.
- Set `artifact: true` on repeating headers/footers so they remain visible but do not enter logical reading order. Artifact text cannot also be a heading.
- Use `addLink(...)` with meaningful visible text, an absolute `http`, `https`, or `mailto` URL, and an explicit reading-order position. The writer emits both a URI annotation and tagged Link/OBJR association.
- A table may share `semanticId` across consecutive pages. To avoid ambiguous interleaving, each continuation must be the first semantic item on its page and each non-final segment must be the last.
- PDF.js table reconstruction for arbitrary files is heuristic. Treat extracted tables as candidates requiring geometry/text review.

## Visual/layout-oriented document

Use ReportLab for a greenfield PDF when precise programmatic layout is the primary requirement and the `PdfArtifact` tagged structure contract is not required.

```bash
python3 scripts/reportlab_create.py \
  --spec tmp/pdfs/report-spec.json \
  --output output/pdf/report.pdf
```

ReportLab output still requires `pdfinfo`, Poppler page rendering, extraction checks, and any separately required PDF/A/UA accessibility validation. Do not claim `PdfArtifact` tagging or reading-order guarantees for ReportLab output.

## Required final gates

1. Reopen with the originating provider.
2. Extract expected text and tables with an independent provider where available.
3. Run structural/tag/conformance checks relevant to the requested deliverable.
4. Render every final page through Poppler and inspect it.
