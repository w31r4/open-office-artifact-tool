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
3. Use `pdf.addFlowText(...)` for long prose so wrapping, margins, long tokens, and page creation remain deterministic; use explicit bounding boxes for intentionally positioned labels, tables, images, and charts.
4. Run `pdf.verify()`; fix page bounds, malformed data, empty objects, Unicode dashes, and non-numeric chart issues.
5. Export with `PdfFile.exportPdf()` (tagged by default) and import the exported file again.
6. Inspect the binary structure with `PdfFile.inspectPdf()`; require tagged status, language, structure roles, and cell-level marked content for agent-authored delivery PDFs. Modeled tables must have matching `Table` → `TR` → `TH`/`TD` hierarchy.
7. Parse the real PDF through PDF.js and check extracted text, geometry, tables, embedded-image data URLs/placement boxes, and page count.
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

await (await PdfFile.exportPdf(pdf, {
  // Required when content is not ASCII. Supply a legally usable standalone .ttf.
  font: process.env.PDF_UNICODE_FONT,
})).save("qa-report.pdf");
```

The built-in Helvetica path is intentionally ASCII-only. For Chinese, Cyrillic, Greek, accented Latin, or other Unicode text, provide a standalone glyf-based TrueType `.ttf` through `font`; export fails instead of silently replacing unsupported text. The writer embeds a Type0/CIDFontType2 font, glyph widths, and a `ToUnicode` map. Fonts are subset by default to used glyphs and recursive composite dependencies; use `subsetFont: false` or CLI `--subset-font false` only when comparing full-font interoperability. It does not yet perform complex-script shaping or accept `.ttc` collections.

## Verification commands

Verify any PDF and write model/file/parser/layout/native evidence:

```sh
node skills/pdf/scripts/verify-pdf.mjs \
  --input qa-report.pdf \
  --output-dir tmp/pdf-qa \
  --pdfjs required \
  --require-tagged true \
  --native-render required
```

Fixture creation accepts the same font input with `--font /path/to/unicode.ttf`, an optional `--max-font-bytes` safety bound, and `--subset-font false` for diagnostic full-font output.

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

- `PdfFile.inspectPdf(...)` checks the PDF header/version, page/object counts, embedded clean-room model, EOF marker, tagged status, language, structure-element/role counts, and MCID count.
- Agent-authored fixture PDFs require `--require-tagged true`. The verifier also derives the expected table, row, header-cell, and data-cell counts from the modeled artifact and rejects flattened or missing table roles. This is stronger structural evidence, but not proof of full PDF/UA conformance; review spans/header associations, reading order, alt text, Unicode fonts, and contrast separately.
- `pdf.inspect(...)`, `pdf.extractText()`, `pdf.extractTables()`, and `pdf.verify()` prove modeled agent-facing semantics.
- PDF.js independently parses the real exported bytes into page text geometry, regions, inferred tables, and bounded PNG image data when XObject pixels are available; placeholders must be reported when masks or unsupported color spaces prevent extraction.
- Per-page Playwright PNGs catch modeled preview regressions.
- Poppler PNGs are the real PDF render gate; inspect every page at full size.
- Optional PNG baselines use `visualQaArtifact(..., { pixelDiff: true })`; approve changes only after full-page review.
- Baseline approval replaces stale model/native page files; later page-count changes fail QA even if all remaining pages match.
- Changed pages write PNG diff heatmaps into the QA output directory; dimension mismatches require a non-strict alignment mode.
- Use `--diff-alignment center|top-left|strict`, `--diff-color '#ff1848'`, and `--diff-unchanged-color '#334155'` to make dimension changes and review palettes explicit.
- For same-size renders with a known platform jitter, opt in to bounded registration with `--registration-offset 2`; use `--registration-improvement 0.1` to require at least 10% sampled mismatch improvement. QA records the chosen baseline translation and ignored edge pixels.
- Deliver only the requested PDF; previews, baselines, extraction dumps, and QA reports are internal unless requested.

## References

- Generated public API catalog: `../../docs/api.md`
- Current implementation coverage: `../../docs/coverage.md`
- Fixture runner: `scripts/run-fixture.mjs`
- Generic PDF verifier: `scripts/verify-pdf.mjs`
