# OpenChestnut Documents API quick start

Use `open-office-artifact-tool` for ordinary DOCX creation, import, semantic editing, export, inspect, and verification. `DocumentFile.importDocx` and `DocumentFile.exportDocx` always use the bundled OpenChestnut C# WebAssembly codec; do not pass codec selectors or lossy-fallback options.

## Startup

Resolve Node.js and the package directory through the Codex workspace dependency loader. Work in a writable task directory, link that loader-provided package directory into the task workspace when necessary, and use ES modules.

```js
import {
  DocumentFile,
  DocumentModel,
  FileBlob,
} from "open-office-artifact-tool";
```

## Create and export

```js
const document = DocumentModel.create({
  name: "Decision brief",
  defaultRunStyle: { fontFamily: "Aptos", fontSize: 11, color: "#172033" },
  blocks: [],
});

document.styles.add("BriefTitle", {
  name: "Brief Title",
  type: "paragraph",
  basedOn: "Normal",
  fontFamily: "Aptos Display",
  fontSize: 26,
  bold: true,
  color: "#123B5D",
  spaceAfterTwips: 120,
});

document.addParagraph("Decision brief", { styleId: "BriefTitle" });
const recommendation = document.addParagraph("Approve the proposed rollout.", {
  styleId: "Normal",
  paragraphFormat: { spaceAfterTwips: 180 },
});
document.addListItem("Confirm the owner and launch date.", {
  listType: "bullet",
  numberFormat: "bullet",
  levelText: "•",
});
document.addTable({
  name: "decision-evidence",
  widthDxa: 9000,
  indentDxa: 120,
  columnWidthsDxa: [2400, 1800, 4800],
  cellMarginsDxa: { top: 80, right: 120, bottom: 80, left: 120 },
  borderColor: "B8C4CE",
  borderSize: 6,
  headerFill: "DCEAF3",
  values: [
    ["Metric", "Value", "Interpretation"],
    ["Readiness", "92%", "Core gates passed"],
  ],
});
document.addComment(recommendation, "Confirm wording before publication.", {
  author: "Reviewer",
  initials: "RV",
});
document.addHeader("Decision brief | Internal");
document.addFooter("1", { fieldInstruction: "PAGE" });

const output = await DocumentFile.exportDocx(document);
await output.save("output.docx");
```

Use real `addListItem` calls for lists and exact `widthDxa`, `indentDxa`, `columnWidthsDxa`, and `cellMarginsDxa` values for tables. Do not fake lists with text markers or use tables as prose layout containers.

`document.addSection(...)` inserts a section break before the blocks that follow it. Use it only when the document actually changes section geometry or header/footer behavior; never append an otherwise unused section block at the end of a document.

## Import and edit

```js
const input = await FileBlob.load("input.docx");
const document = await DocumentFile.importDocx(input);

const target = document.blocks.find(
  (block) => block.kind === "paragraph" && block.text.includes("Approve"),
);
if (!target) throw new Error("Target paragraph was not found.");

const text = document.resolve(`${target.id}/text`);
text.text = "Approve the revised rollout after final QA.";

const table = document.blocks.find((block) => block.kind === "table");
if (table) table.getCell(1, 1).value = "95%";

const output = await DocumentFile.exportDocx(document);
await output.save("edited.docx");
```

Imported unsupported package graphs are source-bound. Keep edits within recognized editable blocks; if OpenChestnut rejects an edit, narrow the edit or report the unsupported boundary instead of flattening or silently rebuilding the document.

## Inspect, resolve, verify, and render

```js
const inspection = document.inspect({
  kind: "document,paragraph,listItem,table,comment,header,footer,section,layout",
  maxChars: 20_000,
});
console.log(inspection.ndjson);

const report = document.verify({ visualQa: true });
if (!report.ok) throw new Error(report.ndjson || JSON.stringify(report.issues));

const preview = await document.render({ format: "svg" });
await preview.save("preview.svg");
```

For final visual QA, export the DOCX and use the packaged `render_docx.py` workflow to inspect every rendered page. SVG/model verification complements native rendering; it does not replace it.

## Supported ordinary authoring boundary

- Paragraphs and formatted runs
- Named paragraph/character/table styles with `basedOn`
- Numbered and character-bulleted lists
- Fixed-geometry tables
- Sections, headers, footers, and PAGE/simple fields
- External/internal hyperlinks
- PNG/JPEG inline images
- Classic whole-paragraph comments

Bookmarks, bibliography-backed citations, tracked revisions, modern comment replies, content controls, complex fields, floating drawings, and other advanced graphs are not source-free authoring features. Imported versions are preserved only while their source evidence remains valid.

Use `DocumentFile.inspectDocx` or `DocumentFile.patchDocx` only when the user explicitly requests package-level inspection or patching. These are deliberate low-level operations, never an automatic fallback for ordinary authoring.
