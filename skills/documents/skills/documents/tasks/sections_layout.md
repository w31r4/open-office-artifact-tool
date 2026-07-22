# Task: Section breaks and mixed page layout

## Goal

Use public `DocumentModel` section blocks without breaking page geometry or
section-scoped headers and footers.

## Key concept

A section block inserts a break before the content that follows it. It defines
page size, orientation, and margins for the following section. Do not append an
unused section block at the end of a document; that can produce a blank page.

## Audit an existing package

```bash
python scripts/section_audit.py input.docx
```

The script is an audit, not an authoring engine. Use it to inventory section
count, geometry, and header/footer relationships before a semantic edit.

## Public API pattern

```js
import { DocumentFile, DocumentModel } from "open-office-artifact-tool";

const document = DocumentModel.create({ blocks: [] });
document.addParagraph("Portrait-section evidence.");

document.addSection({
  name: "landscape-evidence",
  breakType: "nextPage",
  orientation: "landscape",
  pageSize: { widthTwips: 15840, heightTwips: 12240 },
  margins: { top: 720, right: 900, bottom: 720, left: 900 },
  lineNumbering: { countBy: 5, start: 0, distance: 360, restart: "newPage" },
  pageNumbering: { start: 1, format: "lowerRoman" },
  columns: { count: 2, spacing: 720, separator: true },
});
document.addParagraph("Landscape-section evidence.");

document.addHeader("Landscape appendix", {
  referenceType: "default",
  sectionIndex: 1,
});
document.addFooter("1", {
  referenceType: "default",
  sectionIndex: 1,
  fieldInstruction: "PAGE",
});

await (await DocumentFile.exportDocx(document)).save("out.docx");
```

Imported section relationship/linkage graphs beyond the modeled section
boundary are source-bound. If changing one would invalidate source evidence,
OpenChestnut fails closed; do not flatten the document to force the edit.

`columns` has two mutually exclusive bounded profiles. Equal-width layout uses
`{ count, spacing, separator }`. Asymmetric layout uses ordered native column
definitions:

```js
columns: {
  definitions: [
    { width: 3000, spacing: 720 },
    { width: 5640, spacing: 0 },
  ],
  separator: true,
}
```

Both profiles allow 1–45 columns and use twentieths of a point. In the custom
profile, `spacing` means space after that definition; never combine
`definitions` with equal-width `count` or root `spacing`. Ordinary margins,
binding gutter, widths, and inter-column gaps
must fit the page content width. Duplicate containers, ignored equal-width
root attributes, unknown children, extension-bearing definitions, and other
ambiguous graphs remain source-owned; inspect reports that section as
`editable: false`.

`lineNumbering` owns one canonical native `w:lnNumType` leaf and places line
numbers before each text column. An empty object enables every-line numbering
with `countBy: 1`. The bounded fields are:

```js
lineNumbering: {
  countBy: 5,         // 1..32767
  start: 0,           // optional native zero-based value; first display is 1
  distance: 360,      // optional twentieths of a point
  restart: "newPage", // optional: newPage | newSection | continuous
}
```

Set `section.lineNumbering = undefined` to remove the canonical leaf. This
paragraph can opt out of the displayed sequence and its calculation, or
explicitly override suppression inherited from a named style:

```js
document.addParagraph("Unnumbered heading", {
  styleId: "Heading1",
  paragraphFormat: { suppressLineNumbers: true },
});
document.addParagraph("Numbered evidence.");

// In a style definition, the same property suppresses every inheriting paragraph.
// An explicit false on a paragraph overrides that inherited suppression.
```

`true` emits direct/style `w:suppressLineNumbers`; `false` retains an explicit
direct override; omission inherits the style/default behavior. Recognized
canonical direct and style leaves are editable. Duplicate leaves, children,
extension attributes, and invalid lexical values remain source-owned and fail
closed on semantic replacement. Duplicate or irregular `w:lnNumType` leaves,
invalid numeric values, and unknown restart values likewise make the section
read-only. Use native Word/LibreOffice pagination to check the displayed
numbers and column placement; model preview alone is not authoritative.

`pageNumbering` owns one canonical native `w:pgNumType` leaf. Use `start` to
restart a section at an integer from 0 through 2147483647; omit it to continue
the previous section's sequence. The optional `format` is one of `decimal`,
`upperRoman`, `lowerRoman`, `upperLetter`, or `lowerLetter`. At least one of
`start` or `format` is required:

```js
pageNumbering: { start: 1, format: "lowerRoman" }
```

This setting controls page numbers displayed by PAGE fields; it does not add a
footer, insert a field, paginate the document, or refresh cached field text.
Add a PAGE footer/header separately and use native Word/LibreOffice rendering
for final QA. Chapter style/separator attributes, unsupported number formats,
duplicate leaves, extensions, and empty `w:pgNumType` markup remain
source-owned and make the section read-only.

## Render review

- Only the intended pages change orientation.
- Page size and margins match the explicit twip values.
- Text flows through the requested equal-width or asymmetric columns and separator rules.
- Line numbers use the requested increment, offset, restart behavior, and distance in every text column.
- PAGE fields restart/continue and display in the requested section format.
- Headers and footers appear in the intended section.
- First/even variants behave as requested.
- No empty trailing page was introduced.

When a document mixes page sizes or orientations, pass an explicit `--dpi` to
`render_docx.py` if exact output pixel dimensions matter.
