# Task: Heading hierarchy and numbering

## Goal

Produce a consistent, readable hierarchy using public `DocumentModel` styles
and real list blocks.

## Rules

1. Use named paragraph styles, not repeated direct formatting.
2. Keep the hierarchy consistent; do not jump from level 1 to level 3 without a
   real structural reason.
3. Do not type manual heading numbers or bullet characters into paragraph text.
   Use `addListItem(...)` for ordinary numbered/bulleted content.
4. Imported complex multilevel heading-numbering graphs are source-bound. Do
   not rebuild or flatten them when OpenChestnut rejects an edit.

## Public API pattern

```js
import { DocumentFile, DocumentModel } from "open-office-artifact-tool";

const document = DocumentModel.create({ blocks: [] });
for (const [id, fontSize, spaceBeforeTwips] of [
  ["Heading1", 18, 280],
  ["Heading2", 14, 220],
  ["Heading3", 12, 160],
]) {
  document.styles.add(id, {
    name: id.replace(/(\d)/, " $1"),
    type: "paragraph",
    basedOn: "Normal",
    fontSize,
    bold: true,
    keepNext: true,
    spaceBeforeTwips,
    spaceAfterTwips: 80,
  });
}

document.addParagraph("Executive Summary", { styleId: "Heading1" });
document.addParagraph("Background", { styleId: "Heading2" });
document.addParagraph("Prior Work", { styleId: "Heading3" });
document.addListItem("First action", {
  listType: "number",
  numberFormat: "decimal",
  levelText: "%1.",
  numberingId: 41,
  abstractNumberingId: 4,
});

await (await DocumentFile.exportDocx(document)).save("out.docx");
```

Complex automatic TOCs and outline-linked multilevel heading numbering are not
source-free authoring features in the current model. Preserve imported versions
unchanged or report that boundary.

## Validate and render

```bash
python scripts/heading_audit.py input.docx
python render_docx.py input.docx --output_dir out
```

Inspect every page for consistent size, weight, spacing, indentation, and no
fake headings or manual markers.
