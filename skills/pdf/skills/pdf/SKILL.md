---
name: "pdf"
description: "Read, create, inspect, render, and verify PDF files where semantic structure and visual layout matter. Use open-office-artifact-tool as the primary PDF model/API, PDF.js for explicit arbitrary-file parsing, and Poppler for final native rendering QA."
---

# PDF Skill

## When To Use

- Read or review PDF content where layout, tables, figures, or reading order matter.
- Create an accessible, tagged PDF programmatically.
- Make bounded semantic edits to a modeled PDF artifact.
- Extract text, tables, images, and page geometry from an arbitrary PDF.
- Validate final rendering before delivery.

## Primary Workflow

1. Use `open-office-artifact-tool` for creation, semantic editing, inspect, resolve, verification, export, and modeled rendering.
2. For a PDF created by this package, import it with `PdfFile.importPdf(...)`, make bounded model edits, verify, and export again.
3. For an arbitrary PDF, explicitly inject `createPdfjsParser()` from `open-office-artifact-tool/pdf/pdfjs`. Treat the result as a reconstructed model, not a lossless native object graph.
4. Run both `pdf.verify()` (model semantics) and `PdfFile.inspectPdf(...)` (file/tag structure) before delivery.
5. Render every exported page through `createPopplerRenderer()` or `pdftoppm`, inspect the PNGs, fix defects, and repeat.

Start with [PDF API quick start](artifact_tool/API_QUICK_START.md). The shipped [public API end-to-end example](examples/public-api-end-to-end.mjs) creates a tagged scorecard, round-trips it, edits a table cell, inspects its structure, and optionally renders every page through Poppler.

```bash
node skills/pdf/skills/pdf/examples/public-api-end-to-end.mjs \
  output/pdf/release-readiness-scorecard.pdf \
  tmp/pdfs/release-readiness-scorecard-pages
```

## Create And Edit

Use the public package surface:

```js
import {
  FileBlob,
  PdfArtifact,
  PdfFile,
  verifyArtifact,
} from "open-office-artifact-tool";
```

- Build pages with `PdfArtifact.create(...)` and `pdf.addPage(...)`.
- Add positioned text and headings with `pdf.addText(...)`.
- Use `pdf.addFlowText(...)` for wrapped, automatically paginated body text.
- Add semantic tables, figures, and charts with `pdf.addTable(...)`, `pdf.addImage(...)`, and `pdf.addChart(...)`.
- Set a complete logical sequence with `page.setReadingOrder(...)`; visual paint order and reading order are independent.
- Locate edits with bounded `inspect(...)`, then `resolve(...)` or a semantic kind/name/text/table-position lookup.
- Export with `PdfFile.exportPdf(...)`; re-import and verify the final bytes before delivery.

Provide meaningful `alt` text for images and charts, or mark genuinely decorative content with `decorative: true`. Use valid H1-H6 nesting and TH/TD table semantics. Do not fake tables with aligned text or rely on visual position as reading order.

Model IDs and names are locators in the current object graph, not a persistent identity protocol across unrelated imports. Resolve targets again after importing another file.

## Arbitrary PDF Import

Use PDF.js explicitly when the source is not a package-generated modeled PDF:

```js
import { createPdfjsParser } from "open-office-artifact-tool/pdf/pdfjs";

const input = await FileBlob.load("input.pdf");
const pdf = await PdfFile.importPdf(input, {
  parser: createPdfjsParser(),
  preferParser: true,
  parserName: "pdfjs",
});
```

Parser-backed import supports extraction, inspect, page geometry, modeled rendering, and bounded reconstruction. Arbitrary native object-stream/content-stream editing, signatures, forms, annotations, and incremental updates are outside this model. If a requested change crosses that boundary, explain it rather than silently flattening or rebuilding the file.

## Inspect And Verify

Use all three evidence layers:

1. `pdf.inspect(...)` and `pdf.layoutJson(...)` for semantic objects, geometry, and targeted context.
2. `pdf.verify()` or `verifyArtifact(pdf)` for accessibility, reading order, heading nesting, table semantics, figure metadata, bounds, and extraction sanity.
3. `PdfFile.inspectPdf(...)` for byte-level evidence: PDF version, pages/objects, EOF, tagged status, language, reading-order IDs, H1-H6 roles, Figure `/Alt`, Artifact content, Table/TR/TH/TD structure, spans, and font evidence.

Verification success is necessary but not sufficient. Always inspect final rendered pages when visual layout matters.

## Render And Visual Review

During authoring, use `pdf.render({ pageIndex })` for an SVG model preview. For final QA, render the actual exported PDF with the public Poppler adapter:

```js
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";

const renderer = createPopplerRenderer({ dpi: 144, timeoutMs: 60_000 });
for (let pageIndex = 0; pageIndex < pdf.pages.length; pageIndex += 1) {
  const png = await pdf.render({
    source: "pdf",
    format: "png",
    pageIndex,
    renderer,
  });
  await png.save(`tmp/pdfs/page-${pageIndex + 1}.png`);
}
```

Direct Poppler commands remain valid for explicit native QA:

```bash
pdfinfo "$INPUT_PDF"
pdftoppm -png -r 144 "$INPUT_PDF" "$OUTPUT_PREFIX"
```

Inspect every page, including page-count changes after edits. Check clipping, overlap, whitespace balance, typography, table legibility, chart labels, image sharpness, footer consistency, and missing glyphs.

## Temp And Output Conventions

- Use `tmp/pdfs/` for intermediate PDFs, SVGs, PNGs, layouts, and diffs; delete them when done.
- Write final artifacts under `output/pdf/` when working in this repo.
- Keep filenames stable and descriptive.

## Dependencies

The primary dependency is the installed `open-office-artifact-tool` package. Resolve it through the Codex workspace dependency loader when necessary.

Optional capabilities:

- `pdfjs-dist` for explicit arbitrary-PDF parsing through `open-office-artifact-tool/pdf/pdfjs`.
- Poppler `pdftoppm` and `pdfinfo` for native render/file QA.
- Playwright or sharp adapters for additional model/raster QA when explicitly selected.

Do not default to reportlab, pdfplumber, pypdf, or another helper stack for ordinary creation/editing. They may be used only for a user-selected comparison or unsupported low-level investigation, with that boundary stated clearly.

## Quality Expectations

- Maintain polished visual design: consistent typography, spacing, margins, and hierarchy.
- Avoid clipped text, overlaps, broken tables, black squares, unreadable glyphs, or unexplained blank pages.
- Keep charts, tables, and images sharp, aligned, labeled, and accessible.
- Use ASCII hyphens only; avoid U+2011 and other Unicode dashes unless an embedded Unicode font is deliberately supplied.
- Keep citations and references human-readable; never leave tool tokens or placeholder strings.

## Final Checks

- Final semantic verification passes.
- Final byte-level inspection reports the expected pages, EOF, tagged structure, headings, tables, figures, and reading order.
- PDF.js extraction is checked when arbitrary-file interoperability matters.
- Every final page has been rendered from the exported PDF and visually inspected.
- Final output and QA evidence are in the requested locations; intermediate files are removed or organized.
