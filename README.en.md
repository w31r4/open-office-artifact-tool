# OfficeKit

**The multi-tool for agent-operated Office and PDF files.**

[简体中文](README.md) | **English**

OfficeKit gives agents one workflow for creating, reading, editing, inspecting,
rendering, and verifying DOCX, XLSX, PPTX, and PDF artifacts.

It is not a traditional CLI with a large command tree. It is three layers that
work together:

- **Skills** tell an agent when to use a capability, how to verify the result,
  and when to refuse an unsafe edit.
- **JavaScript APIs** provide the object models, calculation, Compose, inspect,
  render, verify, and explicit package-patch primitives.
- **Native engines** provide OpenChestnut C#/.NET WASM for Office and lazy
  MuPDF.js for PDF.

> The repository package is still named `open-office-artifact-tool`; `OfficeKit` is the user-facing product name. The current version is the `0.3.0` release candidate and has not been formally published to npm.

## Deploy in 30 seconds

### Install Skills only

You do not need to clone the repository or install .NET, Python, or Office. Add
the Skills you need to the current agent project:

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

`npx skills` places the Skill files where the agent can discover them; you do
not need to copy the `skills/` tree by hand. Add `--global` for a user-level
installation, or `--agent` to select a specific host.

### Install the JavaScript runtime

Until the first npm release, install the release candidate from GitHub. After
publication, replace the source with the package name:

```sh
# Current release candidate
npm install github:w31r4/open-office-artifact-tool

# After the npm release
npm install open-office-artifact-tool
```

Consumers only need Node.js. The package includes the OpenChestnut WASM
runtime; a local .NET SDK is not required. Node.js 22 or newer is recommended.
MuPDF WASM initializes only on the first PDF operation. Root import, ordinary
Office work, and `npm install` do not download additional runtimes.

Run a complete example after installation:

```sh
node node_modules/open-office-artifact-tool/examples/create-xlsx-dashboard.mjs
```

### Try one Skill without installing it

For a one-off prompt, use:

```sh
npx skills use w31r4/open-office-artifact-tool --skill pdf
```

## First API call

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

More runnable examples:

- [Create a DOCX report](examples/create-docx-report.mjs)
- [Create an XLSX dashboard](examples/create-xlsx-dashboard.mjs)
- [Create a PPTX deck with Compose](examples/create-pptx-compose.mjs)
- [Parse and render a PDF](examples/parse-render-pdf.mjs)

## Four formats, one agent workflow

| Format | Default engine | What it is for |
| --- | --- | --- |
| DOCX | OpenChestnut C# WASM | Structured documents, styles, sections, tables, images, fields, comments, and bounded content controls. |
| XLSX | OpenChestnut C# WASM | Cells, formulas, styles, layout, tables, images, validation, conditional formats, comments, charts, sparklines, What-If data tables, and bounded PivotTables. |
| PPTX | OpenChestnut C# WASM | Shapes, rich text, reversible image crops, tables, connectors, charts, notes, comments, master/layout fidelity, and bounded source-bound edits. |
| PDF | PdfArtifact + MuPDF.js | Tagged authoring; reading, inspection, rendering, forms, links, annotations, page edits, rewrite redaction, and bounded signing. |

Ordinary Office import and export use one OpenChestnut path. Content that cannot
be modeled safely remains bound to its source package; unsupported edits fail
explicitly instead of silently switching to a second codec.

## PDF specialist capabilities are on demand

`mupdf@1.28.0` is the required PDF npm dependency. It is resolved by a normal
`npm install` and loaded lazily at runtime. Large specialist tools such as qpdf,
Python, OCR, and veraPDF/JRE are not placed in the npm tarball and are never
silently installed by a lifecycle hook or global package manager.

Ask the public provider API to decide whether a specialist is ready:

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

The default policy is `disabled`. Only an explicit project `managed` policy,
matching platform/hash/size/license/language constraints, can authorize
`ensure`; a deployment may instead select `system-only`. There is no implicit
fallback. See [PDF Provider Setup](skills/pdf/skills/pdf/tasks/provider_setup.md)
for the complete policy and security boundary.

## Skills and templates

The repository provides the four reference file Skills plus template tooling:

- [Documents](skills/documents/skills/documents/SKILL.md)
- [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md)
- [Presentations](skills/presentations/skills/presentations/SKILL.md)
- [PDF](skills/pdf/skills/pdf/SKILL.md)
- [Template Creator](skills/template-creator/skills/template-creator/SKILL.md)
- [Office Template Library](skills/default-template-library/README.md) — 20 MIT-licensed templates, repository-distributed and excluded from the npm runtime tarball.

Templates are not a second codec. An agent materializes a named template into a
new output file, then uses the same Office APIs to inspect, edit, render, and
verify it. The retained reference is never overwritten.

## Verify before delivery

Agents should use this loop before handing off an artifact:

```text
intent → inspect → resolve → edit/create → export → re-import → render → verify
```

The APIs expose structured evidence wherever possible. Missing source snapshots,
untrusted topology, signature invalidation, possible PDF revision residue, or
missing external credentials cause a fail-closed result.

## Development and verification

```sh
npm test
npm run test:pack
npm run docs:api
npm run release:check
```

See [coverage](docs/coverage.md) for the support boundary, [API reference](docs/api.md),
and [reference Skill compatibility](docs/reference-skills.md).

## License

[GNU AGPL v3 or later](LICENSE). Network deployment, modification, and
redistribution must satisfy the applicable AGPL obligations. Third-party
runtime, MuPDF, and specialist-provider licenses and provenance are recorded
in `THIRD_PARTY_NOTICES.md` and the relevant runtime notices.
