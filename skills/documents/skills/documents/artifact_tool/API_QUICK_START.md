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

## Inline SEQ, REF, and PAGEREF fields

Use `paragraph.addField(...)` when a field must be mixed with ordinary text in
one paragraph. The visible argument is the cached result used before a
compatible host refreshes fields:

```js
const caption = document.addParagraph("", { styleId: "Caption" });
caption.addRun("Figure ");
caption.addField("SEQ Figure \\* ARABIC", "0", {
  bookmarkName: "fig1",
  style: { bold: true },
});
caption.addRun(": Revenue. See figure ");
caption.addField("REF fig1 \\h", "0");
caption.addRun(" on page ");
caption.addField("PAGEREF fig1 \\h", "0");
caption.addRun(".");
```

OpenChestnut writes each logical field run as the canonical five-run native
`begin` / `instrText` / `separate` / cached-result / `end` graph. For a `SEQ`
run, `bookmarkName` inserts a paired Word bookmark around only the cached
result. OpenChestnut imports that exact profile back into one `run.inlineField`
object. Source-free authoring
accepts only `SEQ <label> \\* ARABIC`, `REF <bookmark> \\h`, and `PAGEREF
<bookmark> \\h`, with bounded Word-compatible names. On an imported paragraph,
ordinary text and cached field results may change, but field positions,
instructions, bookmark names, and native bookmark IDs are source-bound.

Use `tasks/captions_crossrefs.md` for the full workflow. Its explicit package
helper remains useful for bulk caption discovery/insertion and deterministic
SEQ/REF materialization. Never claim a cached result is current without host
refresh or materialization plus render review.

## Bibliography-backed citations

Use a canonical Word bibliography catalog plus a whole-paragraph citation field
when the source must remain machine-addressable in Word:

```js
document.addBibliographySource({
  id: "bibliography/AgentSource",
  tag: "AgentSource",
  sourceType: "Book",
  title: "Sketch of the Analytical Engine",
  year: "1843",
  publisher: "Scientific Memoirs",
  authors: [{ first: "Ada", last: "Lovelace" }],
});
document.addCitation("(Lovelace, 1843)", { tag: "AgentSource" });

const imported = await DocumentFile.importDocx(await DocumentFile.exportDocx(document));
imported.bibliographySources[0].title = "Notes on the Analytical Engine";
imported.blocks.find((block) => block.kind === "citation").text = "(Lovelace, 1843, revised)";
```

OpenChestnut authors one `b:Sources` Custom XML part and canonical
`w:fldSimple` `CITATION <tag>` blocks. Source-free catalogs may contain one or
more supported source types, scalar fields, ordinary personal authors, or one
corporate author. Recognized imports permit settings, source-type, author, and
scalar-field edits plus citation display-text edits. Imported source count,
order, IDs, tags, and citation tags remain source-bound.

Contributor roles other than ordinary Author, complex field switches, multiple
or irregular bibliography parts, nested/mixed result runs, and bibliography
output fields remain opaque/read-only. Never rebuild those graphs to claim a
lossless edit; report the boundary or use an explicit narrow package workflow.

## Native table of contents field

For a new document, author the bounded native TOC placeholder through the
public model instead of patching the package:

```js
document.addParagraph("Executive summary", { styleId: "Heading1" });
document.addParagraph("Decision", { styleId: "Heading2" });

const toc = document.addTableOfContents({
  levels: "1-3",
  display: "Refresh fields before delivery",
});

const docx = await DocumentFile.exportDocx(document);
const imported = await DocumentFile.importDocx(docx);
if (!imported.settings.updateFields) throw new Error("TOC refresh hint missing");
const importedToc = imported.blocks.find(
  (block) => block.kind === "field" && block.complex,
);
importedToc.display = "Update this table of contents in Word";
```

`addTableOfContents()` writes one canonical complex field with `TOC \\o
"1-3" \\h \\z \\u` and enables `w:updateFields` by default. The setting is
only a request to refresh fields when a compatible host opens the document; it
does not prove that cached headings or page numbers are current. Open the final
DOCX in Word, update fields, save, and render every page before delivery.

OpenChestnut can re-import and edit the unrefreshed one-paragraph canonical
placeholder. Once Word expands a TOC across multiple paragraphs, the whole
field span is deliberately opaque/source-bound and read-only. Do not rebuild
that result graph to claim a lossless edit. Use `tasks/toc_workflow.md` for the
native-refresh and deterministic static-TOC alternatives.

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
- Sections, headers, footers, PAGE/simple fields, and one canonical complex TOC placeholder with an explicit `updateFields` refresh hint
- External/internal hyperlinks and source-free bookmarks around one paragraph-like block
- Plain-text footnotes/endnotes anchored at the end of one paragraph or list item; recognized imported note bodies allow text-only edits
- PNG/JPEG inline images
- Classic whole-paragraph comments
- Standalone whole-paragraph tracked insertions/deletions with one text run, author, and optional ISO timestamp
- Inline plain-text content-control runs with tag/alias identity, transactional fill-by-tag, and fixed-topology imported edits
- Canonical bibliography source catalogs and whole-paragraph `CITATION` fields with fixed imported source/tag topology

In-paragraph tracked replacements, mixed accepted/revision runs, nested revisions, moves, property changes, and automatic future-change tracking are advanced package workflows, not ordinary public-model authoring. Bookmarks spanning multiple blocks or table cells, nested/crossing ranges, multi-paragraph or reused note graphs, complex bibliography contributor roles/field switches/output fields, modern comment replies, rich/block/cell/data-bound/dropdown/date/checkbox content controls, complex fields other than the canonical one-paragraph TOC placeholder, floating drawings, and other advanced graphs are likewise outside source-free authoring. Recognized imported whole-block bookmarks are inspectable/resolvable but fixed-topology and read-only. Canonical imported footnote/endnote text, bounded citation/source content, bounded inline plain-text control text/tag/alias, and canonical unrefreshed TOC instruction/display may change, but their anchors, native IDs, tags, and topology remain source-bound; refreshed cross-paragraph TOC graphs and other imported advanced graphs are preserved only while their source evidence remains valid.

Use `DocumentFile.inspectDocx` or `DocumentFile.patchDocx` only when the user explicitly requests package-level inspection or patching. These are deliberate low-level operations, never an automatic fallback for ordinary authoring.

OpenChestnut preserves imported table-style catalogs, but direct source-free table authoring cannot materialize an arbitrary custom table style graph. Use `styleId: "TableGrid"` with explicit `widthDxa`, columns, margins, borders, and header fill; an unsupported custom table style fails closed.
