# Merge, reorder, and stamp PDFs

Use the typed pypdf `merge-stamp` primitive when complete source PDFs must be merged, their pages reordered, and a watermark applied only to pages selected by source identity. Do not reconstruct pages through `PdfArtifact`, write a temporary PDF engine, or fall back to another provider.

## Bounded contract

The manifest must use `open-office-artifact-tool.pdf-merge-stamp.v1` and declare:

- two or more unique source IDs and PDF paths;
- a sequence that selects every source page exactly once, in final output order;
- one or more watermark rules selected by source ID, with explicit text and opacity.

Selecting every page exactly once is intentional. It makes outlines, named destinations, and internal links unambiguous even when one source appears in discontiguous sequence segments. Use another reviewed workflow if pages must be duplicated or omitted.

The primitive refuses encrypted input because it has no output-encryption policy. It refuses signed or signature-constrained input unless `--invalidate-signatures` explicitly acknowledges the rewrite. Watermark placement currently fails closed for a source with non-zero page `/Rotate`; inherent portrait/landscape page sizes are supported and preserved.

Start from [`examples/merge-stamp-manifest.json`](../examples/merge-stamp-manifest.json), using paths relative to the manifest or absolute paths.

## Execute

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pdf_provider.py check --provider pypdf --require
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task merge-stamp --provider pypdf --strategy rewrite \
  --input tmp/pdfs/merge-stamp.json --output outputs/merged.pdf \
  --require-provider
"$PYTHON_BIN" scripts/pypdf_edit.py merge-stamp \
  tmp/pdfs/merge-stamp.json outputs/merged.pdf --strategy rewrite
"$PYTHON_BIN" scripts/poppler_compare.py merge-stamp \
  tmp/pdfs/merge-stamp.json outputs/merged.pdf \
  --report tmp/pdfs/merge-visual-qa.json \
  --render-dir tmp/pdfs/merge-rendered
```

`pypdf_edit.py` imports every source once so pypdf can translate its navigation graph, reorders the resulting page-tree references, then applies ReportLab-generated transparent overlays through pypdf. It transactionally reopens the result and verifies page boxes/rotation, page order, watermark text/opacity, outlines, named destinations, internal links, and source immutability before promoting the output.

## Audit and render

For a multi-source operation, the canonical audit `source` is the exact merge manifest and `inputs` lists the source PDFs with absolute path, bytes, and SHA-256. Validate every binding:

```bash
"$PYTHON_BIN" scripts/pdf_audit.py validate outputs/audit.json \
  --source tmp/pdfs/merge-stamp.json \
  --input inputs/cover.pdf --input inputs/report.pdf --input inputs/appendix.pdf \
  --artifact outputs/merged.pdf --require-operation merge-stamp
```

The typed comparison renders every source and output page with Poppler, checks the manifest page mapping, requires non-watermarked pages to be pixel-identical, requires watermarked pages to change, and fails on dimension, blank-state, or excessive dark-pixel drift. Use its JSON status as the visual delivery gate. Inspect its rendered pages for watermark legibility and navigation, but do not delete a passing artifact based only on a subjective thumbnail impression.
