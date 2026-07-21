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
4. Picture markers belong to a numbering level, not an individual paragraph.
   Every item sharing `numberingId` and `level` must use the same embedded
   PNG/JPEG/GIF or external HTTP(S) marker. Edit the complete imported group
   together and retain its embedded-versus-external source kind.
5. Imported complex multilevel heading-numbering or irregular picture-bullet
   graphs are source-bound. Do
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

const pictureMarker = {
  dataUrl: "data:image/png;base64,...",
  sizePt: 12,
  alt: "Action marker",
};
for (const text of ["Validate the output", "Review every native page"]) {
  document.addListItem(text, {
    listType: "bullet",
    numberFormat: "bullet",
    levelText: "•",
    numberingId: 42,
    abstractNumberingId: 5,
    pictureBullet: pictureMarker,
  });
}

await (await DocumentFile.exportDocx(document)).save("out.docx");
```

Picture bullets accept bounded embedded PNG/JPEG/GIF data URLs or absolute
HTTP(S) references; external bytes are not fetched. Imported canonical markers
can be changed only as a complete same-`numberingId`/same-level group without
switching source kind. Complex automatic TOCs, outline-linked multilevel heading
numbering, inherited picture bullets, and irregular VML are not source-free
authoring features in the current model. Preserve imported versions unchanged
or report that boundary.

## Validate and render

```bash
python scripts/heading_audit.py input.docx
python render_docx.py input.docx --output_dir out
```

Inspect every page for consistent size, weight, spacing, indentation, and no
fake headings or manual markers.
