---
name: open-office-pdf
description: Create, import, inspect, extract, render, baseline, and verify PDF artifacts with open-office-artifact-tool, PDF.js, Playwright, and Poppler.
---

# Open Office PDF

Use this project skill for standalone `.pdf` artifact work. It is the clean-room PDF workflow for `open-office-artifact-tool`; it calls the package's public facade and legally usable parser/render adapters only.

## Contract

- Never import or copy the reference package, runtime artifact, runtime module, or runtime bindings.
- Preserve imported content and page geometry unless the user requests a redesign.
- Use real positioned text, tables, PNG/JPEG images, charts, page regions, and multi-page PDF objects rather than flattening semantic objects into labels.
- Use ASCII hyphens. Avoid U+2010 through U+2015 compatibility dashes.
- Keep text, tables, charts, images, page numbers, and section transitions inside the page frame with consistent margins and legible sizing.
- Treat extraction and visual rendering as complementary evidence: neither replaces the other.
- Do not deliver a PDF until semantic/file checks, PDF.js extraction, and full-size Poppler review of every page pass.

## Authoring workflow

1. Create a `PdfArtifact` or import an existing PDF with `PdfFile.importPdf` and an explicit parser when arbitrary-PDF extraction is required.
2. Inspect pages, positioned text, regions, tables, images, and charts before editing.
3. Add content with page-space bounding boxes and stable names.
4. Run `pdf.verify()`; fix page bounds, malformed data, empty objects, Unicode dashes, and non-numeric chart issues.
5. Export with `PdfFile.exportPdf()` and import the exported file again.
6. Inspect the binary structure with `PdfFile.inspectPdf()`.
7. Parse the real PDF through PDF.js and check extracted text, geometry, tables, and page count.
8. Render every modeled page with Playwright and every real PDF page with Poppler.
9. Inspect every page PNG at full size. When a baseline is approved, compare both modeled and native PNG pixels on later runs.

```js
import { PdfArtifact, PdfFile } from "open-office-artifact-tool";

const pdf = PdfArtifact.create({
  pages: [{
    text: "QA report\nSemantic and visual evidence agree",
    textItems: [{ text: "Agent-ready", bbox: [72, 142, 180, 20], fontSize: 18, color: "#16a34a", bold: true }],
    tables: [{
      name: "qa-gates",
      bbox: [72, 190, 468, 128],
      values: [["Gate", "Result"], ["Semantic", "Pass"], ["Visual", "Required"]],
    }],
    charts: [{
      name: "evidence-chart",
      title: "Evidence by gate",
      chartType: "bar",
      bbox: [72, 380, 468, 220],
      categories: ["Model", "Native"],
      series: [{ name: "Checks", values: [4, 6], color: "#3D8DFF" }],
    }],
  }],
});

await (await PdfFile.exportPdf(pdf)).save("qa-report.pdf");
```

## Verification commands

Verify any PDF and write model/file/parser/layout/native evidence:

```sh
node skills/pdf/scripts/verify-pdf.mjs \
  --input qa-report.pdf \
  --output-dir tmp/pdf-qa \
  --pdfjs required \
  --native-render required
```

Create an approved visual baseline:

```sh
node skills/pdf/scripts/verify-pdf.mjs \
  --input qa-report.pdf \
  --output-dir tmp/pdf-baseline-run \
  --baseline-dir tmp/pdf-baselines \
  --write-baseline true \
  --pdfjs required \
  --native-render required
```

Compare a later render:

```sh
node skills/pdf/scripts/verify-pdf.mjs \
  --input qa-report.pdf \
  --output-dir tmp/pdf-compare-run \
  --baseline-dir tmp/pdf-baselines \
  --pdfjs required \
  --native-render required
```

Run the checked-in multi-page fixture:

```sh
node skills/pdf/scripts/run-fixture.mjs \
  --fixture skills/pdf/fixtures/qa-report.json \
  --output-dir tmp/pdf-skill-fixture \
  --pdfjs required \
  --native-render required
```

`auto` runs an optional parser/renderer when available and records a skip otherwise. Use `required` for final local delivery when PDF.js and Poppler are installed, and `off` only for an explicitly documented structural-only check.

## QA gates

- `PdfFile.inspectPdf(...)` checks the PDF header/version, page and object counts, embedded clean-room model, and EOF marker.
- `pdf.inspect(...)`, `pdf.extractText()`, `pdf.extractTables()`, and `pdf.verify()` prove modeled agent-facing semantics.
- PDF.js independently parses the real exported bytes into page text geometry, regions, inferred tables, and image operators.
- Per-page Playwright PNGs catch modeled preview regressions.
- Poppler PNGs are the real PDF render gate; inspect every page at full size.
- Optional PNG baselines use `visualQaArtifact(..., { pixelDiff: true })`; approve changes only after full-page review.
- Baseline approval replaces stale model/native page files; later page-count changes fail QA even if all remaining pages match.
- Changed same-size pages write PNG diff heatmaps into the QA output directory.
- Deliver only the requested PDF; previews, baselines, extraction dumps, and QA reports are internal unless requested.

## References

- Generated public API catalog: `../../docs/api.md`
- Current implementation coverage: `../../docs/coverage.md`
- Fixture runner: `scripts/run-fixture.mjs`
- Generic PDF verifier: `scripts/verify-pdf.mjs`
