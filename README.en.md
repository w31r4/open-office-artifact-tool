# open-office-artifact-tool

[简体中文](README.md) | **English**

An Office and PDF toolkit for agents to create, read, edit, inspect, render, and verify artifacts.

`open-office-artifact-tool` provides a unified JavaScript object model. DOCX, XLSX, and PPTX files are read and written by **OpenChestnut**, implemented in C# with the Open XML SDK and compiled to .NET WebAssembly. PDF uses a separate semantic model and a runtime-lazy **MuPDF.js** native pipeline.

> **Current status:** `0.2.0` release candidate. The source tree, reproducible WASM build, and npm tarball have verification gates, but the package has not yet been formally published to npm.

## Quick start

Until the first npm release, run the project from source:

```sh
git clone https://github.com/w31r4/open-office-artifact-tool.git
cd open-office-artifact-tool
npm install
node examples/create-xlsx-dashboard.mjs
```

The local release gates have passed on Node.js 26.5.0, while hosted CI uses Node.js 22. These are verified environments, not a frozen minimum version. Normal consumers load the WASM bundled with the repository or npm package and do not need a local .NET installation. Rebuilding OpenChestnut, or building and testing the optional OfficeBridge, requires .NET SDK 8.

You can also use the public API directly:

```js
import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Summary");

sheet.getRange("A1:B2").values = [
  ["Metric", "Value"],
  ["Revenue", 42.5],
];

const xlsx = await SpreadsheetFile.exportXlsx(workbook, { recalculate: true });
const reopened = await SpreadsheetFile.importXlsx(xlsx);
console.log(reopened.inspect({ kind: "worksheet,table,chart" }).ndjson);
```

More runnable examples:

- [Create a DOCX report](examples/create-docx-report.mjs)
- [Create an XLSX dashboard](examples/create-xlsx-dashboard.mjs)
- [Create a PPTX deck with Compose](examples/create-pptx-compose.mjs)
- [Parse and render a PDF](examples/parse-render-pdf.mjs)

### PDF runtime

The official `mupdf@1.28.0` package is a required npm dependency and is resolved by a normal `npm install`. Its WASM runtime initializes only on the first PDF read, inspect, render, or edit operation. There is no `postinstall`, standalone downloader, or global environment mutation. ReportLab, pdfplumber, pypdf, Poppler, pikepdf, pyHanko, veraPDF, and OCRmyPDF remain separately installed, task-specific external tools.

## Why it exists

- **Designed for agents:** artifact models expose `inspect`, `resolve`, `verify`, render, and visual QA primitives.
- **Fidelity first:** Office content that cannot be modeled safely stays bound to its source package and is preserved unchanged; unsupported edits fail explicitly.
- **Native Skills included:** npm ships five plugin bundles for Documents, Spreadsheets, Presentations, PDF, and Template Creator (six Skills in total). The repository additionally retains a MIT-licensed, repository-only Office Template Library with 20 templates; workflows that require a host session or an external provider state their prerequisites explicitly.

## Supported surface

| Format | File pipeline | Current core capabilities |
| --- | --- | --- |
| XLSX | OpenChestnut C# WASM | Cells and formulas, styles and layout, tables, images, validation, conditional formatting including standard data bars and icon sets, comments, charts, sparklines, bounded What-If data tables, and bounded native PivotTables. |
| DOCX | OpenChestnut C# WASM | Structured text and styles, sections, headers and footers, lists, tables, links, fields, images, classic comments, bounded modern comment threads, and block/inline plain-text, canonical checkbox, canonical drop-down, canonical combo-box with bounded custom values, or strict `YYYY-MM-DD` date-picker content controls. |
| PPTX | OpenChestnut C# WASM | Shapes and rich text, images with reversible cropping, tables, connectors, charts, direct backgrounds, plain-text speaker notes, legacy comments, and bounded Office 2021 modern threads; slide masters and layouts are preserved but cannot be edited. |
| PDF | Independent model + MuPDF.js | Tagged PDF authoring; native read/inspect/render for arbitrary PDFs; bounded annotation, form, page, metadata, link, rewrite, and incremental edits; real rewrite redaction; bounded local-PKCS#12 signing with independent validation. Specialist tools verify strict sanitization, PDF/UA, OCR, and advanced signing. |

See the [coverage matrix](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/coverage.md) for complete, continuously updated support boundaries.

## How it works

```text
Agent / Skill
├─ Office → JavaScript model → OpenChestnut C# WASM → DOCX / XLSX / PPTX
├─ PDF    → PdfArtifact (new files) or MuPDF.js (import/edit) → PDF
└─ QA     → inspect / resolve → render → verify / visual QA
```

OpenChestnut is the only parser/writer used by normal Office import and export. Explicit OOXML inspect/patch functions are advanced operations that must be invoked manually, never an automatic fallback.

## Native Skills

The repository contains six plugin bundles and twenty-six Skills. The first five bundles provide six npm-distributed Skills; the last bundle provides twenty repository-only template Skills:

- [Documents](skills/documents/skills/documents/SKILL.md)
- [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md)
- [Excel Live Control](skills/spreadsheets/skills/excel-live-control/SKILL.md) — requires a live Excel session supplied by the host
- [Presentations](skills/presentations/skills/presentations/SKILL.md)
- [PDF](skills/pdf/skills/pdf/SKILL.md)
- [Template Creator](skills/template-creator/skills/template-creator/SKILL.md) — creates or explicitly updates reusable local templates from DOCX, PPTX, or XLSX references
- [Office Template Library](skills/default-template-library/README.md) — 20 retained MIT-licensed templates: 7 DOCX, 7 PPTX, and 6 XLSX; repository-only and excluded from the npm tarball

The first five `skills/<name>` directories are shipped in the package; loading is handled by the Agent host. Normal Office Skill workflows use OpenChestnut. The PDF Skill defaults to a thin MuPDF.js CLI that calls the same package APIs installed by npm. Template Creator writes only below `${OFFICE_ARTIFACT_HOME:-~/.office-artifact-tool}/skills`, transactionally retains the explicitly supplied local reference and PNG preview, performs no network fetch, and never overwrites an unnamed template. The Default Template Library retains original Office and PNG files from a MIT reference repository and records their provenance and hashes. An Agent must materialize a named template to a new output file and must never mutate the checked-in reference. All 20 templates are verified for import, unchanged export, second import, and native rendering; verified mutations are deliberately bounded to PPTX slide names, the DOCX update-fields setting, and ordinary XLSX text cells. Rich source-bound content still fails explicitly rather than being silently rebuilt or replaced with an approximate layout. See the [PDF Provider Matrix](skills/pdf/skills/pdf/references/PROVIDER_MATRIX.md) and [template provenance boundary](docs/template-library-provenance.md).

## Important boundaries

- To preserve unmodeled objects from an imported Office file, keep using the model returned by import and leave those objects structurally unchanged. Discarding the source snapshot or changing unsupported topology causes export to fail.
- An arbitrary existing PDF cannot be reflowed reliably like a Word document. Original-file editing must stay within explicit, verifiable bounded operations.
- Shipped pyHanko adapters provide source-bound local-PKCS#12 approval/certification signing and independent validation; the pyHanko runtime remains separately installed. TSA/LTV, PKCS#11, remote signing, and complete PAdES claims remain external workflows. PDF/A/PDF/UA validation and scanned-PDF OCR use the shipped bounded veraPDF and OCRmyPDF adapters.
- Active/auxiliary structure cleanup uses the shipped bounded pikepdf 10.10.x adapter, while pikepdf remains separately installed. The operation retains metadata, form values, XFA, annotations, and hidden text, so it is not complete sanitize or redaction evidence.
- MuPDF.js supports bounded original-file operations, not Word-style arbitrary reflow. Rewrite redaction is also not full sanitization; signatures, residue, OCR, and PDF/UA still require independent evidence.
- LibreOffice, Poppler, Playwright, and the native Office Bridge are rendering and validation tools, not hidden Office codec fallbacks.

## Development and verification

```sh
npm test
npm run test:pack
npm run docs:api
npm run release:check
```

Continue with the [API reference](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/api.md), [runtime architecture](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/reference-runtime-architecture.md), [Skill compatibility](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/reference-skills.md), [Agent PromptBench](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/agent-evals.md), and [release gates](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/release.md).

These documentation links follow the current development branch. They will be pinned to the corresponding version tag when the release is published.

## License

[GNU AGPL v3 or later](LICENSE). Network deployment, modification, and redistribution must satisfy the applicable AGPL obligations. Third-party runtime licenses and provenance are recorded in `THIRD_PARTY_NOTICES.md` and the OpenChestnut runtime notices.
