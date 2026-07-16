# Render and review

Visual QA is mandatory after every meaningful mutation and before delivery.

## File evidence

```bash
pdfinfo output.pdf > tmp/pdfs/pdfinfo.txt
```

Confirm page count, page sizes, encryption, metadata, and other expected properties.

## Render all pages

```bash
mkdir -p tmp/pdfs/pages
pdftoppm -png -r 144 output.pdf tmp/pdfs/pages/page
```

Or use `createPopplerRenderer()` from `open-office-artifact-tool/renderers/poppler` when working in JavaScript.

Inspect every page, not only the first. Check:

- clipping, overlaps, unexpected blank pages, and changed page boxes;
- font substitution, missing glyphs, text wrapping, line spacing, and hierarchy;
- table borders, merged cells, repeated headers, and overflow;
- images, masks, transparency, charts, labels, and raster sharpness;
- form/widget appearances, annotations, signature appearances, links, and footers;
- redaction rectangles and surrounding reflow without revealing hidden content.

Semantic or structural verification can pass while text is visually clipped. The independent greenfield forward test reproduced this exact failure, so Poppler page review cannot be skipped.

When a baseline exists, compare page count and pixel images with declared thresholds, but still inspect material differences. Record renderer/version/DPI and keep diff evidence for release-critical outputs.

For a manifest-driven merge/reorder/stamp, use `scripts/poppler_compare.py merge-stamp` instead of relying on thumbnail judgment. It provides page-to-source mapping, pixel stability, changed bounding boxes, blank-state evidence, and dark-pixel ratios. A passing typed comparison is authoritative for those declared invariants; subjective inspection can add findings but must not invent a blocking defect without reproducible contrary evidence.
