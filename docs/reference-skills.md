# Reference Skill compatibility

The 0.2 source tree publishes the reference file-type layout as four native plugin bundles, not as the earlier flat project-specific Skills:

```text
skills/
  documents/{.codex-plugin,README.md,assets,skills/documents}
  spreadsheets/{.codex-plugin,.app.json,README.md,assets,skills/{spreadsheets,excel-live-control}}
  presentations/{.codex-plugin,README.md,assets,skills/presentations}
  pdf/{.codex-plugin,README.md,assets,skills/pdf}
```

There are therefore four plugin packages and five Skills. The old fixture runners remain under `test/skill-harness`; they are development tests and are excluded from the npm package.

## Verified compatibility

| Surface | Status | Evidence and remaining boundary |
| --- | --- | --- |
| Plugin manifests and discovery | done | All four manifests pass the plugin validator. Every declared Skill, `agents/openai.yaml`, plugin icon, and Skill icon resolves inside its plugin bundle. |
| Private-package imports | done | Published JavaScript examples and runners import `open-office-artifact-tool`; the package test rejects imports from `office-artifact-tool`. |
| Presentations built-in template | done | The unflattened 26-slide `codex-grid-layout-library` runs through its shipped `create-presentation.mjs`, canonical OpenChestnut export, and second import. Text runs, body properties, absolute frames, 11 literal custom geometries, one prompt placeholder image, and two connectors survive. |
| Presentation workspace helper | done | The shipped setup helper resolves the public package root or bundled runtime, creates a module workspace, and links/imports `open-office-artifact-tool` without a private runtime path. |
| Spreadsheet core example | done | The reference-style create/value/formula/fill/style/chart/SVG-render/export/import example, `Workbook.fromCSV`, and the shipped `openchestnut-range-workflow.mjs` and `openchestnut-sparkline-workflow.mjs` run against the public package. They cover R1C1 formulas, block writes, formula evidence, current-region/navigation, formatting, charting, standard line/column sparklines, verification, property edits, and two canonical OpenChestnut round trips. An independent forward test also authored a polished three-sheet operating forecast with formula-driven financials, model checks, zero spreadsheet errors, a formula-bound line chart, OpenChestnut roundtrip, and visual review of every sheet. |
| Full Spreadsheet Quick API | partial | The high-value Range slice is compatible: `formulasR1C1`, `displayFormulas`, `formulaInfos`, `Range.write/writeValues`, clear/copy modes, relative formula translation, even tiling, used/current regions, navigation aliases, and scalar/matrix number formats are documented and tested. `containsText` derives its required native formula, direct text references retain labels, formula-only internal chart series resolve live caches for inspect/render/export, and standard Office 2010 line/column/stacked sparkline groups support reversible source-free authoring plus fixed-topology imported edits. Remaining reference breadth includes source-free data tables, non-reversible native sparkline graphs, chart families beyond the bounded bar/line/pie profile, reply-capable threaded comments, advanced conditional-format graphs, and formulas outside the public Help catalog. |
| Excel live control | partial | Routing content, Skill metadata, icon, and `.app.json` connector declaration are present. Execution requires the host-provided connected-document app plus an active Excel add-in session; the npm package does not implement that service. |
| Documents | partial | The ordinary native workflow runs a shipped `DocumentModel` create → OpenChestnut export → import/edit → export → second-import example, with semantic verification and render-backed QA. Bounded whole-block bookmarks/internal hyperlinks and whole-paragraph tracked insertions/deletions use the public API and native OpenChestnut markup; Python/OOXML helpers remain explicit for complex bookmark graphs, in-paragraph replacement, accept/reject batches, complex revision graphs, and audits. Full status remains partial because content controls, complex fields, and other source-bound features still exceed the public semantic model. |
| Full Presentation API guide | partial | The built-in source-free template is compatible. Direct solid/style-reference slide backgrounds support source-free authoring, inspect evidence, and semantic-hash-bound imported add/edit/remove without flattening Layout/Master inheritance. Plain-text speaker notes support source-free authoring, stable inspect/resolve, and bounded hash-bound imported edits. Embedded rectangular images support source-free and imported `stretch`/`cover`/`contain`, signed explicit crop, crop-aware SVG preview, and native DrawingML `a:srcRect` add/edit/remove; import exposes the native rectangle as `fit: "stretch"` plus explicit crop because PPTX has no fit keyword. An imported top-level OLE object with exactly one internal, uniquely bound, unshared XLSX package supports defensive workbook extraction and payload-only replacement through `getEmbeddedWorkbook()` / `replaceEmbeddedWorkbook()`; OpenChestnut validates the new workbook while preserving the OLE shell, preview, relationship topology, and unrelated package parts. Shared, external, ambiguous, non-XLSX, and irregular graphs remain opaque and read-only. Complex backgrounds, rich notes, external pictures, non-rectangular masks, blip effects, and broader reference instructions for advanced Master/Layout, comments, custom shows, notes styling/visibility, and other package graphs still exceed the current canonical export boundary. |
| PDF | partial | The native guide is a reference-compatible capability router. A normal npm install now supplies required, runtime-lazy MuPDF.js; the shipped `scripts/mupdf.mjs` CLI drives default arbitrary-file inspect, PNG/JPEG render, and bounded direct-original rewrite/incremental edits with budgets, atomic distinct-path output, object-level signature detection, and fail-closed redaction/deletion rules. The guide also retains greenfield `PdfArtifact`, the six-page CJK accessibility example, ReportLab, pdfplumber, typed pypdf attachment/forms/merge/stamp, specialist PyMuPDF sanitize, canonical audits, residue scans, and independent Poppler comparison. Existing fixed matrices retain their recorded results. qpdf, pyHanko, and veraPDF remain documented/probed external workflows; pikepdf and OCRmyPDF have no shipped adapter; image OCR requires separately installed Tesseract. |

## Compatibility discipline

Plugin packaging and workflow compatibility are separate gates. A plugin is not marked fully compatible merely because its manifest validates or its files appear in the tarball.

For each remaining Skill, convergence requires:

1. run the published instructions from a clean npm install;
2. exercise the public package rather than repository-only test harnesses;
3. keep DOCX/XLSX/PPTX ordinary I/O on the single OpenChestnut facade;
4. retain direct OOXML patching only as an explicit, user-selected low-level operation;
5. render and inspect representative output where the environment supports it;
6. record unsupported reference calls and either implement them or narrow the published instruction honestly;
7. add the workflow to `test/reference-skills.mjs` before promoting its coverage status.

For PDF, “compatible” means preserving the reference foundation while routing advanced work to explicit mature providers. It does not mean forcing every operation through `PdfArtifact`: arbitrary-file extraction models are never exported as fidelity-preserving edits, and provider absence or unsupported capability fails closed.

`test/reference-skills.mjs` is the publication-contract smoke test. The older format-specific tests under `test/*-skill.mjs` continue to exercise the deeper development fixtures under `test/skill-harness`.
