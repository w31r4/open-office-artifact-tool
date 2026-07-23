# OfficeKit

**Turn agent output into deliverable Office and PDF files.**

[简体中文](README.md) | **English**

Create reports, build spreadsheet models, prepare slide decks, and handle PDFs.
OfficeKit puts creation, import, editing, rendering, and verification into one
traceable workflow. The result is a file people can open, edit, and inspect—not
just chat text or a preview image.

OfficeKit is for developers and teams building automated workflows with agents.
The entry points are installable Skills and a JavaScript API. OpenChestnut
handles Office through C#/.NET WASM; MuPDF.js handles PDF and loads only when
needed.

> The repository package is still named `open-office-artifact-tool`; `OfficeKit` is the product name. The current version is the `0.3.0` release candidate and has not been formally published to npm.

## Built for real work

- **Reports and documents**: turn an outline, source material, and data into a styled DOCX with tables, images, fields, and comments; re-import it and check the rendered pages.
- **Finance and operations models**: turn CSVs, assumptions, and business data into an XLSX with formulas, validation, conditional formats, charts, and calculated results.
- **Briefings and proposals**: start from a template or Compose, create a PPTX, preserve crops, charts, notes, and layout, then render pages for review.
- **PDF operations**: read text, tables, images, links, and forms from arbitrary PDFs; enable the matching provider when a task needs repair, OCR, strict cleanup, signatures, or conformance checks.
- **Template-driven production**: turn a local DOCX, XLSX, or PPTX reference into a reusable template and write new files without changing the reference.

## Get started

### Install Skills only

No repository clone, .NET, Python, or Office installation is required:

```sh
npx skills add w31r4/open-office-artifact-tool \
  --skill documents \
  --skill Spreadsheets \
  --skill Presentations \
  --skill pdf \
  --skill template-creator \
  --yes
```

To install every native Skill and the open-source template collection:

```sh
npx skills add w31r4/open-office-artifact-tool --skill '*' --yes
```

The Skill files are placed where the agent can discover them; there is no need
to copy the `skills/` tree by hand. Add `--global` for a user-level install or
`--agent` to select a specific host.

### Start a complete project from an empty directory

```sh
mkdir officekit-agent && cd officekit-agent
npm init -y
npx skills add w31r4/open-office-artifact-tool --skill documents --skill Spreadsheets --skill Presentations --skill pdf --yes
npm install github:w31r4/open-office-artifact-tool
```

After the npm release, replace the last line with
`npm install open-office-artifact-tool`.

Consumers only need Node.js; the package includes the OpenChestnut WASM runtime,
so a local .NET SDK is not required. Node.js 22 or newer is recommended. MuPDF
loads on the first PDF operation; installation and ordinary Office work do not
download additional runtimes.

### Try one Skill

```sh
npx skills use w31r4/open-office-artifact-tool --skill pdf
```

## Call the JavaScript API

```js
import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Summary");
sheet.getRange("A1:B2").values = [
  ["Metric", "Value"],
  ["Revenue", 42.5],
];

const file = await SpreadsheetFile.exportXlsx(workbook, { recalculate: true });
await file.save("summary.xlsx");
```

Runnable examples:

- [Create a DOCX report](examples/create-docx-report.mjs)
- [Create an XLSX dashboard](examples/create-xlsx-dashboard.mjs)
- [Create a PPTX deck with Compose](examples/create-pptx-compose.mjs)
- [Parse and render a PDF](examples/parse-render-pdf.mjs)

## Four formats, one delivery standard

| Format | Default engine | Main use |
| --- | --- | --- |
| DOCX | OpenChestnut C# WASM | Structured documents, styles, sections, tables, images, fields, comments, and bounded content controls. |
| XLSX | OpenChestnut C# WASM | Cells, formulas, styles, layout, tables, images, validation, conditional formats, comments, charts, sparklines, What-If data tables, and bounded PivotTables. |
| PPTX | OpenChestnut C# WASM | Shapes, rich text, reversible image crops, tables, connectors, charts, notes, comments, master/layout fidelity, and bounded source-bound edits. |
| PDF | PdfArtifact + MuPDF.js | Tagged authoring; reading, inspection, rendering, forms, links, annotations, page edits, rewrite redaction, and bounded signing. |

Ordinary Office import and export use one OpenChestnut path. Content that cannot
be modeled safely stays bound to its source package; unsupported edits fail
explicitly instead of producing a file that only looks successful.

## PDF specialist capabilities on demand

`mupdf@1.28.0` is the required PDF npm dependency. A normal `npm install` resolves
it and the runtime loads lazily. Specialist tools such as qpdf, Python, OCR, and
veraPDF/JRE are not placed in the npm tarball and are never silently installed by
a lifecycle hook or global package manager.

Let the provider API select a route from the task and inspection evidence:

```js
import { PdfFile } from "open-office-artifact-tool";
import { PdfProviders } from "open-office-artifact-tool/pdf/providers";

const inspection = await PdfFile.inspectPdf("input.pdf");
const resolution = await PdfProviders.resolve({
  task: "repair",
  provider: "qpdf",
  inspection,
});

console.log(resolution.status); // ready | installable | blocked
```

The default policy is `disabled`. Only an explicit project `managed` policy with
matching platform, hash, size, license, and language constraints can authorize
`ensure`; a deployment may instead select `system-only`. There is no implicit
fallback. See [PDF Provider Setup](skills/pdf/skills/pdf/tasks/provider_setup.md).

## Skills and templates

- [Documents](skills/documents/skills/documents/SKILL.md)
- [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md)
- [Presentations](skills/presentations/skills/presentations/SKILL.md)
- [PDF](skills/pdf/skills/pdf/SKILL.md)
- [Template Creator](skills/template-creator/skills/template-creator/SKILL.md)
- [Office Template Library](skills/default-template-library/README.md) — 20 MIT-licensed templates, repository-distributed and excluded from the npm runtime tarball.

Templates and the Office engine share one delivery chain: an agent materializes
a selected template into a new output file, then uses the same APIs to inspect,
edit, render, and verify it. The retained reference is never overwritten.

## Verify before delivery

```text
intent → inspect → resolve → edit/create → export → re-import → render → verify
```

Missing source snapshots, untrusted topology, signature invalidation, possible
PDF revision residue, or missing external credentials produce a fail-closed
result. See [coverage](docs/coverage.md) for the complete boundary.

## Development and verification

```sh
npm test
npm run test:pack
npm run docs:api
npm run release:check
```

See the [API reference](docs/api.md) and [reference Skill compatibility](docs/reference-skills.md).

## License

[GNU AGPL v3 or later](LICENSE). Network deployment, modification, and
redistribution must satisfy the applicable AGPL obligations. Third-party
runtime, MuPDF, and specialist-provider licenses and provenance are recorded
in `THIRD_PARTY_NOTICES.md` and the relevant runtime notices.
