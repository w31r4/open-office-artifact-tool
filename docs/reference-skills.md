# Reference Skill compatibility

The 0.2 source tree publishes the reference file-type layout as four native Codex plugin bundles, not as the earlier flat project-specific Skills:

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
| Plugin manifests and discovery | done | All four manifests pass the Codex plugin validator. Every declared Skill, `agents/openai.yaml`, plugin icon, and Skill icon resolves inside its plugin bundle. |
| Private-package imports | done | Published JavaScript examples and runners import `open-office-artifact-tool`; the package test rejects imports from `office-artifact-tool`. |
| Presentations built-in template | done | The unflattened 26-slide `codex-grid-layout-library` runs through its shipped `create-presentation.mjs`, canonical OpenChestnut export, and second import. Text runs, body properties, absolute frames, 11 literal custom geometries, one prompt placeholder image, and two connectors survive. |
| Presentation workspace helper | done | The shipped setup helper resolves the public package root or bundled runtime, creates a module workspace, and links/imports `open-office-artifact-tool` without a private runtime path. |
| Spreadsheet core example | done | The reference-style create/value/formula/fill/style/chart/SVG-render/export/import example, `Workbook.fromCSV`, and the shipped `openchestnut-range-workflow.mjs` and `openchestnut-sparkline-workflow.mjs` run against the public package. They cover R1C1 formulas, block writes, formula evidence, current-region/navigation, formatting, charting, standard line/column sparklines, verification, property edits, and two canonical OpenChestnut round trips. An independent forward test also authored a polished three-sheet operating forecast with formula-driven financials, model checks, zero spreadsheet errors, a formula-bound line chart, OpenChestnut roundtrip, and visual review of every sheet. |
| Full Spreadsheet Quick API | partial | The high-value Range slice is compatible: `formulasR1C1`, `displayFormulas`, `formulaInfos`, `Range.write/writeValues`, clear/copy modes, relative formula translation, even tiling, used/current regions, navigation aliases, and scalar/matrix number formats are documented and tested. `containsText` derives its required native formula, direct text references retain labels, formula-only internal chart series resolve live caches for inspect/render/export, and standard Office 2010 line/column/stacked sparkline groups support reversible source-free authoring plus fixed-topology imported edits. Remaining reference breadth includes source-free data tables, non-reversible native sparkline graphs, chart families beyond the bounded bar/line/pie profile, reply-capable threaded comments, advanced conditional-format graphs, and formulas outside the public Help catalog. |
| Excel live control | partial | Routing content, Skill metadata, icon, and `.app.json` connector declaration are present. Execution requires the host-provided connected-document app plus an active Excel add-in session; the npm package does not implement that service. |
| Documents | partial | The ordinary native workflow now runs a shipped `DocumentModel` create → OpenChestnut export → import/edit → export → second-import example, with semantic verification and render-backed QA. Python/OOXML helpers are explicitly limited to advanced package patches and audits. Full status remains partial because the broader reference task set includes tracked revisions, content controls, complex fields, and other source-bound features outside the public semantic model. |
| Full Presentation API guide | partial | The built-in source-free template is compatible. Broader reference instructions for advanced Master/Layout, notes/comments, custom shows, groups, and other package graphs exceed the current canonical export boundary and remain fail-closed or preservation-only. |
| PDF | partial | The native guide is a reference-compatible capability router rather than a flattened single-backend Skill. Its shipped paths cover greenfield `PdfArtifact` tagged authoring, ReportLab creation, pdfplumber extraction, type-aware pypdf forms/annotations, PyMuPDF direct-original page/text/image/form/annotation edits, mandatory probe/route preflight, explicit rewrite/incremental/sanitize policy, a canonical mutation-audit schema/byte validator, real redaction, bounded active-content cleanup, inert/residue scans, and Poppler QA. The active-content same-prompt fixed matrices pass candidate/reference `3/3 + 3/3`; the AcroForm comparison passes candidate `3/3` and reference `2/3`, with the remaining reference run correctly rejected for bypassing the typed provider/audit workflow despite correct final appearance. This demonstrates that the richer native Skill reduces workflow variance while retaining the reference foundation. The public model example still round-trips and renders a tagged scorecard. qpdf, pyHanko, and veraPDF remain documented/probed external workflows; pikepdf and OCRmyPDF have no shipped adapter; image OCR requires separately installed Tesseract. |

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
