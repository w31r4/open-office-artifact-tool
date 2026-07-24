# OfficeKit

## Office files your agent can actually hand off

[简体中文](README.md) | **English**

Drafting text is easy. Turning it into a file someone will present tomorrow, send to a customer, keep as a record, and revise next month is harder.

OfficeKit gives agents Skills and JavaScript APIs for Word, Excel, PowerPoint, and PDF work: create a file, read an existing one, make supported changes, then reopen, render, and check the result. When an edit cannot preserve complex material reliably, it stops and tells you why instead of silently damaging the file.

## Ask for the file you need

Give an agent with OfficeKit tasks such as:

> “Turn these CSVs into an operating model for next week’s review. Keep the key metrics as formulas and make the charts ready for the deck.”

> “Use our presentation template for a QBR. Replace the data and images, then check every slide for overflow or misplaced elements.”

> “Update the dates, owners, and clauses in this Word document. Do not disturb the header, table of contents, citations, or existing comments.”

> “Make these scanned PDFs searchable, identify sensitive information, redact it, and leave me a result I can verify.”

OfficeKit can create from scratch or work from an existing file. It is built for reports, financial models, customer proposals, training material, contract drafts, batch template work, and PDF operations.

## Add it to a project in two steps

From the Node.js project where the agent will work:

```sh
npm install github:w31r4/open-office-artifact-tool
npx skills add w31r4/open-office-artifact-tool --skill '*' --yes
```

The first command installs the runtime; the second installs every Skill and the open-source templates. The initial setup needs no repository clone, Office installation, .NET SDK, or Python setup. The formal npm package has not shipped yet, so use the GitHub source for now; after release, replace the first command with `npm install open-office-artifact-tool`.

For a smaller install:

```sh
npx skills add w31r4/open-office-artifact-tool \
  --skill documents \
  --skill spreadsheets \
  --skill presentations \
  --skill pdf \
  --yes
```

Node.js 22 or newer is recommended. The Office runtime is included in the package. MuPDF.js loads only when the first PDF operation needs it.

## Skills tell the agent how to work

The Skills are operating instructions, not a feature checklist. They tell an agent what to inspect first, how to check its result, and when it must stop:

| Skill | Good for |
| --- | --- |
| [Documents](skills/documents/skills/documents/SKILL.md) | Word reports, letters, contract drafts, and formal documents with tables and images. |
| [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md) | Excel models, data preparation, formulas, charts, validation, and visualization. |
| [Presentations](skills/presentations/skills/presentations/SKILL.md) | PowerPoint decks, template work, charts, images, notes, and layout review. |
| [PDF](skills/pdf/skills/pdf/SKILL.md) | PDF reading, authoring, forms, annotations, page work, rendering, and specialist operations. |
| [Template Creator](skills/template-creator/skills/template-creator/SKILL.md) | Turning your own DOCX, XLSX, and PPTX reference files into reusable templates. |

Skills and application code use the same package. An agent can follow a Skill to complete a task, while an application can call the API directly and bring the same file operations into its product or automation.

## Call it from code

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

## Check the file before handing it over

The usual path is straightforward:

```text
read the source → create or change → export → reopen → render pages → check the result
```

- DOCX, XLSX, and PPTX all use OpenChestnut C#/.NET WASM. There is no second JavaScript Office writer hidden behind it, and no local .NET SDK is required.
- For Office material that cannot be modeled safely, OfficeKit retains the relevant source-package content where possible. If a requested edit would damage it, the edit fails rather than producing a file that merely looks correct.
- PDF uses MuPDF.js for normal reading, writing, inspection, and rendering. Repair, OCR, strict cleanup, signing, PDF/A, and PDF/UA work require an explicitly selected provider.

| File | Common deliverables |
| --- | --- |
| DOCX | Styles, paragraphs, sections, headers and footers, tables, images, fields, comments, and bounded edits to existing documents. |
| XLSX | Cells, formulas, styles, merges, dimensions, freeze panes, tables, images, validation, conditional formats, charts, sparklines, and bounded PivotTables. |
| PPTX | Shapes, rich text, images with reversible crops, tables, connectors, charts, notes, comments, and master/layout fidelity. |
| PDF | Authoring; text, table, image, and link extraction; forms and annotations; page operations; rewrite redaction; and bounded signing. |

See [coverage](docs/coverage.md) for the complete supported boundary.

## Heavy PDF capabilities need project approval

MuPDF is a normal dependency. qpdf, Python, OCR, veraPDF/JRE, and similar tools are not packed into npm and are never silently fetched by an install script or a global package manager.

The agent resolves a provider from the task and inspection evidence. Downloads are disabled by default. A managed capability pack can be installed only when `.open-office-artifact-tool/pdf-providers.json` explicitly enables `managed` and its platform, hash, size, license, and OCR-language constraints match. An existing runtime can instead be selected through `system-only`.

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

[PDF Provider Setup](skills/pdf/skills/pdf/tasks/provider_setup.md) covers policy, installation, and operational limits.

## Start with a good template

The [Office Template Library](skills/default-template-library/README.md) has 20 MIT-licensed Office templates. The template files stay in the repository rather than the npm runtime package. An agent creates a new output from a selected template, then uses the same APIs to inspect, edit, and render it; the reference file is not used as the output.

Use [Template Creator](skills/template-creator/skills/template-creator/SKILL.md) to turn your team’s DOCX, XLSX, or PPTX reference files into your own templates.

## For users and contributors

- [API reference](docs/api.md)
- [Reference Skill compatibility](docs/reference-skills.md)
- [Complete capability boundary](docs/coverage.md)

For development:

```sh
npm test
npm run test:pack
npm run docs:api
npm run release:check
```

> `OfficeKit` is the product name; the current package name remains `open-office-artifact-tool`. Version `0.3.0` is a release candidate and is not formally published to npm yet.

## License

[GNU AGPL v3 or later](LICENSE). Network deployment, modification, and redistribution must meet the applicable AGPL obligations. Third-party runtime, MuPDF, and specialist-provider licenses and provenance are recorded in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) and the relevant runtime notices.
