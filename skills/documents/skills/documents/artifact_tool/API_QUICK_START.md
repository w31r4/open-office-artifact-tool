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

## Imported text: capability-routed local edits

Imported paragraphs and table cells advertise separate capabilities. Use
`textEditable` for modeled whole-text assignment and `textPatchable` for a
source-bound, unique literal replacement that preserves the surrounding native
OOXML graph. Never infer either capability from visible text alone.

```js
const source = await FileBlob.load("input.docx");
const document = await DocumentFile.importDocx(source);
const candidates = document.blocks.filter(
  (block) => block.kind === "paragraph" && block.text.includes("old wording"),
);
if (candidates.length !== 1) {
  throw new Error(`Expected one target paragraph, found ${candidates.length}.`);
}

const target = candidates[0];
if (!target.textEditable && !target.textPatchable) {
  throw new Error("Target text is source-bound and has no safe public edit capability.");
}
const range = document.resolve(`${target.id}/text`);
if (!range) throw new Error("Advertised text range did not resolve.");
range.replace("old wording", "replacement wording");

const output = await DocumentFile.exportDocx(document);
await output.save("edited.docx");
const reimported = await DocumentFile.importDocx(await FileBlob.load("edited.docx"));
if (!reimported.blocks.some((block) => block.id === target.id && block.text.includes("replacement wording"))) {
  throw new Error("Edited text did not survive the OpenChestnut round-trip.");
}
```

When only `textPatchable` is true, assignment to `range.text` is rejected. The
literal search must be non-empty and resolve to exactly one ordinary native
`w:r/w:t` node or adjacent non-empty ordinary runs with byte-identical `w:rPr`.
Mixed-format spans, empty-run gaps, paragraph boundaries, hyperlinks, fields,
content controls, tracked revisions, or duplicate visible matches fail closed.
Complex table cells use the same contract through
`document.resolve(cell.id + "/text")` or
`table.getCell(row, column).replaceText(old, next)`; `cell.value = ...` remains
unavailable when `cell.editable` is false. Export, re-import, and native-render
every result before delivery.

Use `examples/openchestnut-source-text-patch-workflow.mjs` when editing a real
input file. It takes an explicit paragraph or physical table-cell selector,
keeps the input immutable, requires `textPatchable` rather than whole-text
editing, checks that only `word/document.xml` changed, publishes without
overwrite, reimports, verifies, and writes an audit JSON.

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

const statusMarker = {
  dataUrl: "data:image/png;base64,...",
  widthPt: 14,
  heightPt: 14,
  alt: "Ready status",
};
for (const itemText of ["Validate the package", "Review the native render"]) {
  document.addListItem(itemText, {
    listType: "bullet",
    numberFormat: "bullet",
    levelText: "•",
    numberingId: 41,
    abstractNumberingId: 4,
    pictureBullet: statusMarker,
  });
}
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

## Inline and bounded floating images

`document.addImage(...)` creates an inline image when `placement` is absent.
For intentional wrap-around, pass the bounded floating profile explicitly:

```js
const figure = document.addImage({
  dataUrl: "data:image/png;base64,...",
  alt: "Architecture overview",
  widthPx: 240,
  heightPx: 120,
  placement: {
    type: "floating",
    horizontal: { relativeTo: "margin", offsetPx: 260 },
    vertical: { relativeTo: "paragraph", offsetPx: 0 },
    wrap: "square",
    wrapSide: "bothSides",
    distanceFromTextPx: { top: 4, right: 12, bottom: 4, left: 12 },
  },
});
document.addParagraph("Figure 1. Architecture overview.");
```

Horizontal `relativeTo` accepts `margin`, `page`, or `column`; vertical accepts
`margin`, `page`, or `paragraph`. `topAndBottom` wrap omits `wrapSide`. All four
text distances default to zero. Imported recognized anchors expose the same
placement object, but only fixed-topology placement edits are allowed: an
imported inline image cannot become floating and an imported floating image
cannot become inline. Behind-text, overlapping, tight/through, aligned or
percentage-positioned, effect-bearing, external, and irregular drawings remain
source-owned and fail closed on semantic replacement.

Read `../tasks/images_figures.md` before using floating placement. Re-import the
result, inspect the image record, then run native DOCX render QA; the model SVG
is useful for planning but is not authoritative for cross-renderer anchoring.

## Canonical text watermarks

For a new document, add one text watermark per section/header-reference scope:

```js
document.addWatermark("DRAFT", {
  sectionIndex: 0,
  referenceType: "default",
});
```

After importing an existing DOCX, inspect `document.watermarks`; never infer
editability from a visible background object. A recognized object exposes
`sourceBound === true` and `editable === true`. Locate exactly one record, edit
only `.text`, or call `.remove()`, then export and import again. OpenChestnut
protects the exact VML paragraph and the residual header part. It rejects
scope changes, reordering, new watermark relationships in an imported package,
shared headers, multiple objects, DrawingML, images, and irregular VML rather
than rebuilding or heuristically deleting them.

Use `examples/openchestnut-watermark-workflow.mjs` for imported files. It keeps
the input immutable, allows exactly one changed `word/headerN.xml` part, writes
an audit, and requires native page render review. Read
`../tasks/watermarks_background.md` for the full workflow and the explicitly
separate heuristic package-tool boundary.

## Imported classic comment: bounded text-only edit

For one ordinary imported classic comment, locate both its paragraph and
comment uniquely, preserve its source-bound identity/anchor metadata, change
only `text`, then export, re-import, verify, and render. The shipped
`examples/openchestnut-classic-comment-edit-workflow.mjs` performs the full
transactional workflow and writes a byte-bound audit:

```js
import { DocumentFile, FileBlob } from "open-office-artifact-tool";

const source = await FileBlob.load("input.docx");
const document = await DocumentFile.importDocx(source);
const anchor = document.blocks.filter(
  (block) => block.kind === "paragraph" && block.text.includes("Decision: proceed with controlled rollout."),
);
if (anchor.length !== 1) throw new Error("Expected one target paragraph.");
const comments = document.comments.filter((comment) => comment.targetId === anchor[0].id);
if (comments.length !== 1 || comments[0].parentId || comments[0].resolved || comments[0].person) {
  throw new Error("Only one unresolved classic comment without modern metadata is editable here.");
}
comments[0].text = "Approved after legal review.";
const output = await DocumentFile.exportDocx(document);
await output.save("reviewed.docx");

const reimported = await DocumentFile.importDocx(await FileBlob.load("reviewed.docx"));
if (reimported.comments.length !== 1 || reimported.comments[0].text !== "Approved after legal review.") {
  throw new Error("Classic comment text did not survive the round-trip.");
}
if (!reimported.verify({ visualQa: true }).ok) throw new Error("Verification failed.");
await reimported.render({ format: "svg" });
```

The imported classic comment's ID, `targetId`, author, initials, date, and
anchor topology are source-bound. Do not use this slice to add/delete comments,
reply, resolve/reopen, or manipulate `commentsExtended.xml`/`people.xml`;
modern comment and reply graphs must be preserved or explicitly refused.

## Bounded modern comment thread

For source-free authoring, a root may have direct replies. OpenChestnut writes
the native commentsExtended graph and optional durable/UTC/person parts:

```js
const root = document.addComment(target, "Please confirm the evidence.", {
  author: "Lead reviewer",
  resolved: false,
  dateUtc: "2026-07-19T08:00:00Z",
  person: { providerId: "directory", userId: "lead@example.test" },
});
document.replyToComment(root, "Evidence confirmed.", {
  author: "Release reviewer",
  dateUtc: "2026-07-19T08:05:00Z",
  person: { providerId: "directory", userId: "release@example.test" },
});
```

For an imported recognized thread, locate the root and reply unambiguously,
change only `.text`, and call `root.resolve()` or `root.reopen()`. The shipped
`examples/openchestnut-modern-comment-thread-workflow.mjs` performs the full
source-hash-bound transaction, second import, identity/topology checks, model
render, and audit. Imported parentage, paragraph/durable IDs, UTC/person data,
anchors, and comment count cannot change. Nested replies and irregular support
parts fail closed.

## Block/inline/table-cell text, checkbox, drop-down, combo-box, and date content controls

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

When the entire paragraph is the template field, author a real body-level
`w:sdt` instead of an inline wrapper:

```js
document.addBlockTextContentControl("{{EXECUTIVE_SUMMARY}}", {
  blockId: "executive-summary-paragraph",
  id: "executive-summary",
  tag: "EXECUTIVE_SUMMARY",
  alias: "Executive summary",
  paragraphFormat: { keepNext: true },
  runStyle: { bold: true },
});
```

This bounded body-level profile contains exactly one paragraph and one ordinary
run. Multi-run, nested, locked, placeholder, repeating-section, and data-bound
block SDTs remain opaque/source-bound.

When one source-free rectangular table cell is the template field, wrap its
existing text or typed value with the matching cell primitive:

```js
const approvalTable = document.addTable({
  id: "approval-matrix",
  values: [
    ["Field", "Value"],
    ["Owner", "{{TABLE_OWNER}}"],
    ["Approved", "pending"],
    ["Priority", "pending"],
    ["Contact", "pending"],
    ["Review date", "pending"],
  ],
});
approvalTable.getCell(1, 1).addTextContentControl({
  id: "table-owner",
  tag: "TABLE_OWNER",
  alias: "Table owner",
});

document.fillContentControls({ TABLE_OWNER: "Katherine Johnson" });

approvalTable.getCell(2, 1).addCheckboxContentControl(false, {
  tag: "TABLE_APPROVED",
  alias: "Table approved",
});
approvalTable.getCell(3, 1).addDropdownContentControl([
  { displayText: "Low", value: "low" },
  { displayText: "High", value: "high" },
], {
  tag: "TABLE_PRIORITY",
  alias: "Table priority",
  selectedValue: "low",
});
approvalTable.getCell(4, 1).addComboBoxContentControl(["Email", "Phone"], {
  tag: "TABLE_CONTACT",
  alias: "Table contact",
  value: "Email",
});
approvalTable.getCell(5, 1).addDateContentControl("2026-07-22", {
  tag: "TABLE_REVIEW_DATE",
  alias: "Table review date",
});
```

OpenChestnut emits one direct cell-level `w:sdt` containing exactly one
`w:p/w:r/w:t`. Recognized imports expose the same type-specific state/tag/alias
handle, but native ID, type, table coordinates, wrapper placement, list choices,
checkbox symbols, date profile, and topology stay fixed; adding or removing a
control in an imported ordinary table fails closed.

`document.contentControls` returns fresh handles with `id`, `targetId`,
`placement`, optional inline-only `runIndex`, optional table-cell-only `row` and
`column`, `tag`, `alias`, read-only `nativeId`, and `controlType`. Plain-text
handles have mutable `text`; checkbox handles have mutable boolean `checked`;
drop-down handles expose defensive `choices` and mutable `selectedValue`;
combo-box handles expose defensive `choices` and mutable `value`; date handles
expose a mutable canonical `dateValue`:

```js
const approval = document.addParagraph("Approved: ");
approval.addCheckboxContentControl(false, {
  id: "approved",
  tag: "APPROVED",
  alias: "Approved",
});

const checkboxUpdate = document.setCheckboxContentControls({ APPROVED: true });
if (checkboxUpdate.missingTags.length) throw new Error("Required checkbox missing");

const priority = document.addParagraph("Priority: ");
priority.addDropdownContentControl([
  { displayText: "Low", value: "low" },
  { displayText: "High", value: "high" },
], {
  id: "priority",
  tag: "PRIORITY",
  alias: "Priority",
  selectedValue: "low",
});
const dropdownUpdate = document.setDropdownContentControls({ PRIORITY: "high" });
if (dropdownUpdate.missingTags.length) throw new Error("Required drop-down missing");

const contact = document.addParagraph("Contact method: ");
contact.addComboBoxContentControl([
  { displayText: "Email", value: "email" },
  { displayText: "Phone call", value: "phone" },
], {
  id: "contact-method",
  tag: "CONTACT_METHOD",
  alias: "Contact method",
  value: "email",
});
const comboUpdate = document.setComboBoxContentControls({
  CONTACT_METHOD: "Pager duty",
});
if (comboUpdate.missingTags.length) throw new Error("Required combo box missing");

const review = document.addParagraph("Review date: ");
review.addDateContentControl("2026-07-21", {
  id: "review-date",
  tag: "REVIEW_DATE",
  alias: "Review date",
});
const dateUpdate = document.setDateContentControls({
  REVIEW_DATE: "2028-02-29",
});
if (dateUpdate.missingTags.length) throw new Error("Required date missing");
```

The visible checkbox glyph and canonical Word 2010+ `w14` symbol declarations
are codec-owned; edit `checked`, never run text.
`fillContentControls()` fills every duplicate tag and rejects all unknown tags
before mutation unless `{ strict: false }` is explicit. It matches text controls
only. `setCheckboxContentControls()` applies the same transaction rules to
tag-to-boolean checkbox state. `setDropdownContentControls()` validates every
tag-to-choice-value selection before mutation. Drop-down visible text is
derived from the selected choice. `setComboBoxContentControls()` accepts either
a declared choice value or 1–255 characters of XML-safe custom text; matching
choices use their display text and custom values render verbatim. Imported
drop-down/combo-box choices and order are source-bound.
`setDateContentControls()` accepts only real Gregorian dates in exact
`YYYY-MM-DD` form. OpenChestnut owns the matching visible text and the native
`w:date` projection: UTC midnight `w:fullDate`, `yyyy-MM-dd`, `en-US`, `date`
mapping, and Gregorian calendar. JavaScript `Date` objects, locale-formatted
strings, and direct visible-text edits fail closed instead of invoking timezone
or machine-locale behavior.
Re-resolve controls after each independent import because model IDs are
object-lifetime locators.

OpenChestnut authors and imports the bounded run-level and whole-table-cell
plain-text, Word 2010+ checkbox, `w:dropDownList`, `w:comboBox`, and
ISO/Gregorian `w:date` profiles, plus the separate one-paragraph/one-run
body-block plain-text profile.
Rich, multi-paragraph, inline-within-cell, nested, data-bound, locked,
placeholder, repeating-section, merge-continuation, or irregular cell,
and other irregular block,
irregular drop-down/combo-box, localized-date, legacy/custom-symbol checkbox,
or unrelated extension-bearing SDTs remain
opaque and source-bound. Do not flatten them; follow
`tasks/forms_content_controls.md` for explicit advanced routing and
render-backed QA.

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
work on arbitrary packages. For the canonical model, dry-run and materialize
bounded caches directly:

```js
const plan = document.materializeFields({ dryRun: true });
const result = document.materializeFields();
```

Both results report `seqFields`, `refFields`, `skippedPageReferences`, missing
targets, and exact cache changes. Strict mode fails before mutation. `PAGEREF`
is never fabricated because trustworthy page numbers require a real pagination
host. Never claim a cached result is current without host refresh or
materialization plus render review.

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

Use real `addListItem` calls for lists and exact `widthDxa`, `indentDxa`, `columnWidthsDxa`, and `cellMarginsDxa` values for tables. Do not fake lists with text markers or use tables as prose layout containers. A `pictureBullet` may be an embedded PNG/JPEG/GIF data URL, an absolute HTTP(S) URI, or a configuration object with 4–72 point dimensions and non-empty alt text. The marker belongs to the numbering level, so all items with the same `numberingId` and `level` must match. Imported edits must update that complete group and must not switch between embedded and external sources; external URIs are preserved but never fetched. Render the result through the native QA path before delivery.

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

Enable native future-change tracking independently from authored redlines:

```js
document.setSettings({ trackRevisions: true });
```

Use facing-page inside/outside margins without changing the section margin
numbers themselves:

```js
document.setSettings({ mirrorMargins: true });
// Disable/remove the canonical leaf setting with false.
```

OpenChestnut owns exactly one leaf `w:mirrorMargins`. Duplicate, child-bearing,
extension, or otherwise irregular imported markup stays source-owned: unrelated
body edits preserve it, while semantic replacement fails closed. Structurally
irregular markup also blocks sibling settings edits when exact reserialization
cannot be proved.

Reserve binding space in each modeled section and choose whether Word places
that allowance at the top edge or the binding side:

```js
document.setSettings({ gutterAtTop: true });
document.addSection({
  breakType: "continuous",
  margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, gutter: 720 },
});
```

All margin values are twentieths of a point. `margins.gutter` maps to
`w:pgMar/@w:gutter`; `gutterAtTop: true` owns one canonical leaf
`w:gutterAtTop`, while false/omission uses the binding side. OpenChestnut
requires the ordinary margins plus the gutter to leave a positive content
area. If imported `mirrorMargins` or `gutterAtTop` markup is duplicate,
child-bearing, extension-bearing, or otherwise irregular, it remains
source-owned and the affected section geometry is reported read-only.

Create a canonical equal-width multi-column section with an optional rule:

```js
document.addSection({
  breakType: "continuous",
  pageSize: { widthTwips: 12240, heightTwips: 15840 },
  margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
  columns: { count: 2, spacing: 720, separator: true },
});
```

Or define each asymmetric text column explicitly:

```js
document.addSection({
  breakType: "continuous",
  pageSize: { widthTwips: 12240, heightTwips: 15840 },
  margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
  columns: {
    definitions: [
      { width: 3000, spacing: 720 },
      { width: 5640, spacing: 0 },
    ],
    separator: true,
  },
});
```

Both profiles accept 1 through 45 columns and use twentieths of a point.
Equal-width columns own `count` and one shared `spacing`. Explicit-width
columns instead own ordered `definitions`; every definition has a positive
`width`, and its `spacing` is the space after that column. Do not combine
`definitions` with `count` or root `spacing`. The
occupied widths/gaps plus horizontal margins and any binding-side gutter must
fit the page content width. OpenChestnut authors, imports, edits, and removes
both canonical profiles (`section.columns = undefined`). Duplicate containers,
ignored equal-width root attributes on a custom graph, unknown children,
extension attributes, and other ambiguous graphs stay source-owned and make
the imported section `editable: false` rather than being flattened.

Control the numbering sequence and display format used by PAGE fields in one
modeled section:

```js
document.addSection({
  breakType: "nextPage",
  lineNumbering: { countBy: 5, start: 0, distance: 360, restart: "newPage" },
  pageNumbering: { start: 1, format: "lowerRoman" },
});
```

`lineNumbering` owns one canonical `w:lnNumType` leaf. `{}` defaults to
`{ countBy: 1 }`; otherwise `countBy` accepts 1 through 32767. Optional
`start` accepts the native zero-based range 0 through 32767 (so `start: 0`
first displays 1), optional `distance` accepts 0 through 31680 twentieths of
a point, and optional `restart` is `newPage`, `newSection`, or `continuous`.
Set `section.lineNumbering = undefined` to remove it. Numbers appear before
each text column. Paragraph-level suppression, duplicate/child-bearing/
extension-bearing leaves, invalid values, and unknown restart modes remain
source-owned and make section geometry read-only. Use a native pagination
host for final visual QA.

`start` is optional and accepts integers from 0 through 2147483647. Omit it to
continue the preceding section's sequence. `format` is optional and accepts
`decimal`, `upperRoman`, `lowerRoman`, `upperLetter`, or `lowerLetter`; at
least one property is required. Set `section.pageNumbering = undefined` to
remove the canonical leaf. This setting does not insert a PAGE field or update
its cached display, so add the header/footer field separately and use a native
pagination host for final QA. Chapter style/separator settings, unsupported
formats, duplicate/empty leaves, extensions, and other irregular
`w:pgNumType` graphs remain source-owned and make section geometry read-only.

Set one bounded passwordless Word editing restriction through the same settings state:

```js
document.setSettings({ documentProtection: "readOnly" });
// Also supported: "comments", "trackedChanges", "forms", or explicit "none".
// Remove the element with false, null, or "off".
```

The canonical object form is `{ edit, enforcement, formatting }`; string modes default to `enforcement: true` and `formatting: false`. This is a Word editing restriction, not encryption or access control. Password verifiers, cryptographic attributes, IRM, and permission exceptions remain source-owned. OpenChestnut preserves such imported markup when the semantic setting is left untouched and fails closed if the public model tries to replace it.

Add one exact in-paragraph tracked replacement directly to original source bytes:

```js
import crypto from "node:crypto";
import fs from "node:fs/promises";
import { DocumentFile, FileBlob } from "open-office-artifact-tool";

const bytes = await fs.readFile("input.docx");
const source = new FileBlob(bytes);
const document = await DocumentFile.importDocx(source);
const targetBlockIndex = document.blocks.findIndex(
  (block) => block.kind === "paragraph" && block.text === "The term is 30 days.",
);
const expectedSourceSha256 = crypto.createHash("sha256").update(bytes).digest("hex");
const reviewed = await DocumentFile.addTrackedReplacement(source, {
  target: { kind: "paragraph", blockIndex: targetBlockIndex },
  expectedText: "The term is 30 days.",
  search: "30 days",
  replacement: "45 days",
  author: "Reviewer",
  date: "2026-07-21T09:30:00Z",
  expectedSourceSha256,
});
console.log(reviewed.metadata.trackedReplacement);
await reviewed.save("reviewed.docx");
```

For a bounded table cell, retain the table block index plus the physical row and cell indexes returned by the exact import:

```js
const reviewedCell = await DocumentFile.addTrackedReplacement(source, {
  target: { kind: "tableCell", blockIndex: tableBlockIndex, row: 1, column: 2 },
  expectedText: "Payment is due in 30 days.",
  search: "30 days",
  replacement: "45 days",
  author: "Reviewer",
  expectedSourceSha256,
});
```

The structured selector and full paragraph/cell snapshot bind the semantic target; `targetBlockIndex` remains a paragraph-only compatibility option and is mutually exclusive with `target`. A table target must be a direct body table with a valid physical grid, a non-continuation cell, and exactly one direct paragraph. In both profiles the literal must occur exactly once inside one direct ordinary native text node or adjacent non-empty ordinary runs with byte-identical `w:rPr`. OpenChestnut retains every matched fragment in one `w:del`, writes one `w:ins` with the shared formatting, allocates collision-free native IDs, changes only `word/document.xml`, and reports `matchedSourceRunCount` with the structured target plus byte, element, text, index, native-ID, and changed-part evidence. Stale or duplicate text, empty-run gaps, mixed-format spans, multi-paragraph/nested/continuation table cells, hyperlinks, fields, content controls, drawings, or already-revised targets fail closed. Use `examples/openchestnut-tracked-replacement-workflow.mjs` for immutable-source, unique paragraph/table-cell discovery, no-overwrite publication, reimport, render, and audit checks.

Finalize either bounded revision profile from the original bytes, not from a reconstructed model:

```js
import crypto from "node:crypto";
import fs from "node:fs/promises";
import { DocumentFile, FileBlob } from "open-office-artifact-tool";

const bytes = await fs.readFile("reviewed.docx");
const expectedSourceSha256 = crypto.createHash("sha256").update(bytes).digest("hex");
const clean = await DocumentFile.finalizeRevisions(new FileBlob(bytes), {
  mode: "accept", // or "reject"
  expectedSourceSha256,
  keepTracking: false,
});
console.log(clean.metadata.revisionFinalization);
await clean.save("accepted.docx");
```

This primitive accepts direct body whole-paragraph `w:ins`/`w:del` wrappers with one recognized run and one exact adjacent `w:del` + single-run `w:ins` pair in either a direct body paragraph or the bounded direct table-cell profile above. The deletion may retain multiple source fragments only when all runs share the insertion's exact `w:rPr`. It fails closed for other mixed-format, nested, moved, property-level, irregular-table, or non-body-story revision graphs. Use `examples/openchestnut-revision-finalization-workflow.mjs` for the richer whole-block projection workflow; the tracked-replacement example and API metadata provide the source-bound evidence for inline pairs.

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
- PNG/JPEG inline images plus an explicit bounded foreground floating profile with absolute margin/page/column and margin/page/paragraph offsets, square or top-and-bottom wrap, and fixed-topology imported placement edits
- Classic whole-paragraph comments and bounded modern root/direct-reply threads
- Standalone whole-paragraph tracked insertions/deletions plus one exact source-bound in-paragraph replacement as adjacent native deletion/insertion runs; native `trackRevisions` intent; and source-hash-bound accept/reject finalization for both bounded profiles
- Canonical facing-page `mirrorMargins`, per-section binding `margins.gutter`, document-wide `gutterAtTop`, canonical equal-/explicit-width section columns, bounded section line-number increment/start/distance/restart and page-number restart/format settings, plus passwordless `documentProtection` settings for `none`, `readOnly`, `comments`, `trackedChanges`, and `forms`, with explicit enforcement/formatting flags and source-bound preservation of irregular/cryptographic variants
- Block plain-text plus inline/table-cell plain-text, canonical Word 2010+ checkbox, canonical Word drop-down, canonical Word combo-box, and canonical ISO/Gregorian date content controls with explicit placement, typed values, tag/alias identity, transactional tag updates, and fixed-topology imported edits
- Canonical bibliography source catalogs and whole-paragraph `CITATION` fields with fixed imported source/tag topology

In-paragraph revision graphs beyond the exact single-format deletion/insertion pair, other mixed accepted/revision runs, mixed-format or nested revisions, moves, property changes, multi-paragraph/nested/continuation/irregular table targets, and non-body revision stories are advanced package workflows, not ordinary public-model authoring or bounded finalization. Bookmarks spanning multiple blocks or table cells, nested/crossing ranges, multi-paragraph or reused note graphs, complex bibliography contributor roles/field switches/output fields, nested/irregular modern comment graphs, rich/multi-paragraph/inline-within-cell/nested/data-bound/locked/placeholder/repeating-section cell controls and other irregular content controls, irregular lists, localized dates, custom checkbox symbols, complex fields other than the canonical one-paragraph TOC placeholder, and floating drawings outside the bounded foreground absolute-offset square/top-and-bottom profile are likewise outside source-free authoring. Recognized imported whole-block bookmarks are inspectable/resolvable but fixed-topology and read-only. Canonical imported footnote/endnote text, bounded citation/source content, bounded block/inline/table-cell text-control state, inline/table-cell checkbox checked state, inline/table-cell drop-down selectedValue, inline/table-cell combo-box value, inline/table-cell ISO-date dateValue, and every recognized control's tag/alias may change, as may canonical modern-comment text/resolved state, canonical unrefreshed TOC instruction/display, and bounded floating-image placement. Their anchors, native IDs, control types, table coordinates, list choices/order, symbols, native date profile, and topology remain source-bound; refreshed cross-paragraph TOC graphs and other imported advanced graphs are preserved only while their source evidence remains valid.

Use `DocumentFile.inspectDocx` or `DocumentFile.patchDocx` only when the user explicitly requests package-level inspection or patching. These are deliberate low-level operations, never an automatic fallback for ordinary authoring.

OpenChestnut preserves imported table-style catalogs, but direct source-free table authoring cannot materialize an arbitrary custom table style graph. Use `styleId: "TableGrid"` with explicit `widthDxa`, columns, margins, borders, and header fill; an unsupported custom table style fails closed.
