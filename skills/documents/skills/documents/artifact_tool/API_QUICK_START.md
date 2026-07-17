# OpenChestnut Documents API quick start

Use `open-office-artifact-tool` for ordinary DOCX creation, import, semantic editing, export, inspect, and verification. `DocumentFile.importDocx` and `DocumentFile.exportDocx` always use the bundled OpenChestnut C# WebAssembly codec; do not pass codec selectors or lossy-fallback options.

## Startup

Resolve Node.js and the package directory through the host workspace dependency loader. Work in a writable task directory, link that loader-provided package directory into the task workspace when necessary, and use ES modules.

```js
import {
  DocumentFile,
  DocumentModel,
  FileBlob,
} from "open-office-artifact-tool";
```

Use `document.fontFamilies` when you need a fresh sorted inventory of theme and
explicit run/style fonts before render or handoff QA.

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

const title = document.addParagraph("Decision brief", { styleId: "BriefTitle" });
const titleBookmark = document.addBookmark(title, "DecisionBrief");
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
document.addInsertion("Added release condition.", {
  author: "Reviewer",
  date: "2026-07-17T08:00:00Z",
});
document.addDeletion("Superseded release condition.", {
  author: "Reviewer",
  date: "2026-07-17T08:05:00Z",
});
document.addHyperlink("Back to decision brief", titleBookmark, {
  tooltip: "Jump to the decision brief heading",
});
const approvalCondition = document.addParagraph("The rollout has one final quality condition.");
document.addFootnote(approvalCondition, "Approval remains conditional on final QA.");
const provenance = document.addParagraph("Evidence was collected from the release candidate.");
document.addEndnote(provenance, "Evidence snapshot: 2026-07-17.");

const output = await DocumentFile.exportDocx(document);
await output.save("output.docx");
```

## Inline plain-text content controls

Use a paragraph run-level content control when an Agent must fill a bounded
plain-text template field by tag:

```js
const customer = document.addParagraph("Customer: ");
customer.addTextContentControl("{{CUSTOMER_NAME}}", {
  id: "customer-name",
  tag: "CUSTOMER_NAME",
  alias: "Customer name",
  style: { bold: true },
});

const template = await DocumentFile.exportDocx(document);
await template.save("template.docx");

const importedTemplate = await DocumentFile.importDocx(template);
const fill = importedTemplate.fillContentControls({
  CUSTOMER_NAME: "Ada Lovelace",
});
if (fill.missingTags.length) throw new Error("Required template field missing");

const filled = await DocumentFile.exportDocx(importedTemplate);
await filled.save("filled.docx");
```

`document.contentControls` returns fresh handles with `id`, `targetId`,
`runIndex`, `tag`, `alias`, read-only `nativeId`, and mutable `text`.
`fillContentControls()` fills every duplicate tag and rejects all unknown tags
before mutation unless `{ strict: false }` is explicit. Re-resolve controls
after each independent import because model IDs are object-lifetime locators.

OpenChestnut authors and imports only the bounded run-level plain-text profile.
Rich, block, cell, nested, data-bound, dropdown, date, checkbox,
placeholder-document, locked, or extension-bearing SDTs remain opaque and
source-bound. Do not flatten them; follow `tasks/forms_content_controls.md` for
explicit advanced routing and render-backed QA.

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

const footnote = document.notes.find((note) => note.kind === "footnote");
if (footnote) footnote.text = "Approval remains conditional on final render QA.";

const output = await DocumentFile.exportDocx(document);
await output.save("edited.docx");
```

Imported unsupported package graphs are source-bound. Keep edits within recognized editable blocks; if OpenChestnut rejects an edit, narrow the edit or report the unsupported boundary instead of flattening or silently rebuilding the document.

Model IDs and `name` values are locators for the current object graph, not a persistent identity protocol across independent DOCX imports. After a file round-trip, resolve targets again by bounded semantic text, style, block kind, or table position before editing.

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
- Named paragraph and character styles with `basedOn`; source-free tables may use the bounded `TableGrid` style plus explicit direct geometry/formatting
- Numbered and character-bulleted lists
- Fixed-geometry tables
- Sections, headers, footers, and PAGE/simple fields
- External/internal hyperlinks and source-free bookmarks around one paragraph-like block
- Plain-text footnotes/endnotes anchored at the end of one paragraph or list item; recognized imported note bodies allow text-only edits
- PNG/JPEG inline images
- Classic whole-paragraph comments
- Standalone whole-paragraph tracked insertions/deletions with one text run, author, and optional ISO timestamp
- Inline plain-text content-control runs with tag/alias identity, transactional fill-by-tag, and fixed-topology imported edits

In-paragraph tracked replacements, mixed accepted/revision runs, nested revisions, moves, property changes, and automatic future-change tracking are advanced package workflows, not ordinary public-model authoring. Bookmarks spanning multiple blocks or table cells, nested/crossing ranges, multi-paragraph or reused note graphs, bibliography-backed citations, modern comment replies, rich/block/cell/data-bound/dropdown/date/checkbox content controls, complex fields, floating drawings, and other advanced graphs are likewise not source-free authoring features. Recognized imported whole-block bookmarks are inspectable/resolvable but fixed-topology and read-only. Canonical imported footnote/endnote text and bounded inline plain-text control text/tag/alias may change, but their anchors, native IDs, and topology remain source-bound; other imported advanced graphs are preserved only while their source evidence remains valid.

Use `DocumentFile.inspectDocx` or `DocumentFile.patchDocx` only when the user explicitly requests package-level inspection or patching. These are deliberate low-level operations, never an automatic fallback for ordinary authoring.

OpenChestnut preserves imported table-style catalogs, but direct source-free table authoring cannot materialize an arbitrary custom table style graph. Use `styleId: "TableGrid"` with explicit `widthDxa`, columns, margins, borders, and header fill; an unsupported custom table style fails closed.
