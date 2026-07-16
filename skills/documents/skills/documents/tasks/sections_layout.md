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

## Render review

- Only the intended pages change orientation.
- Page size and margins match the explicit twip values.
- Headers and footers appear in the intended section.
- First/even variants behave as requested.
- No empty trailing page was introduced.

When a document mixes page sizes or orientations, pass an explicit `--dpi` to
`render_docx.py` if exact output pixel dimensions matter.
