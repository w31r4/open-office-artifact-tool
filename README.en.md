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

The official `mupdf@1.28.0` package is a required npm dependency and is resolved by a normal `npm install`. Its WASM runtime initializes only on the first PDF read, inspect, render, or edit operation. There is no `postinstall`, standalone downloader, or global environment mutation. ReportLab, pdfplumber, pypdf, Poppler, pyHanko, and veraPDF remain task-specific external tools.

## Why it exists

- **Designed for agents:** artifact models expose `inspect`, `resolve`, `verify`, render, and visual QA primitives.
- **Fidelity first:** Office content that cannot be modeled safely stays bound to its source package and is preserved unchanged; unsupported edits fail explicitly.
- **Native Skills included:** the package ships four installable Skill bundles for Documents, Spreadsheets, Presentations, and PDF; workflows that require a host session or an external provider state their prerequisites explicitly.

## Supported surface

| Format | File pipeline | Current core capabilities |
| --- | --- | --- |
| XLSX | OpenChestnut C# WASM | Cells and formulas, styles and layout, tables, images, basic validation and conditional formatting, comments, charts, and sparklines. |
| DOCX | OpenChestnut C# WASM | Structured text and styles, sections, headers and footers, lists, tables, links, fields, images, classic comments, and inline plain-text content controls. |
| PPTX | OpenChestnut C# WASM | Shapes and rich text, images with reversible cropping, tables, connectors, charts, direct backgrounds, and plain-text speaker notes; slide masters and layouts are preserved but cannot be edited. |
| PDF | Independent model + MuPDF.js | Tagged PDF authoring; native read/inspect/render for arbitrary PDFs; bounded annotation, form, page, metadata, link, rewrite, and incremental edits; real rewrite redaction. Specialist tools verify strict sanitization, signatures, PDF/UA, and OCR. |

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

The repository contains four plugin bundles and five Skills:

- [Documents](skills/documents/skills/documents/SKILL.md)
- [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md)
- [Excel Live Control](skills/spreadsheets/skills/excel-live-control/SKILL.md) — requires a live Excel session supplied by the host
- [Presentations](skills/presentations/skills/presentations/SKILL.md)
- [PDF](skills/pdf/skills/pdf/SKILL.md)

Each `skills/<name>` directory contains the Skill definitions and resources shipped with the package; loading is handled by the Agent host. Normal Office Skill workflows use OpenChestnut. The PDF Skill defaults to a thin MuPDF.js CLI that calls the same package APIs installed by npm; Python and system tools are reserved for specialist capabilities without an equivalent implementation. See the [PDF Provider Matrix](skills/pdf/skills/pdf/references/PROVIDER_MATRIX.md).

## Important boundaries

- To preserve unmodeled objects from an imported Office file, keep using the model returned by import and leave those objects structurally unchanged. Discarding the source snapshot or changing unsupported topology causes export to fail.
- An arbitrary existing PDF cannot be reflowed reliably like a Word document. Original-file editing must stay within explicit, verifiable bounded operations.
- PDF signing, timestamps, and LTV rely on external pyHanko workflows. PDF/A and PDF/UA machine validation relies on external veraPDF. The package does not bundle complete adapters for either.
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
