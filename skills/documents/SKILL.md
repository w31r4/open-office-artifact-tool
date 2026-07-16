---
name: documents
description: Create, import, edit, inspect, render, and verify DOCX artifacts through the canonical OpenChestnut Office path.
---

# Documents

Use this project skill for standalone `.docx` work. `DocumentFile.exportDocx` and `DocumentFile.importDocx` are the only high-level file path; both load the bundled OpenChestnut C# WebAssembly engine. JavaScript remains responsible for the public object model, Compose/JSX, calculation, inspection, rendering, QA, and explicit package patching.

Do not select an alternate Office writer. The high-level methods accept no routing flags. Pass only `limits` when a tighter resource budget is needed.

## Supported 0.2 workflow

The canonical path covers the common Documents workflow:

- paragraphs and runs with bounded direct font, size, RGB color, character spacing, bold, italic, and underline formatting;
- paragraph alignment, indents, spacing, line rules, keep-with-next, and page-break-before;
- paragraph, character, and table styles with bounded `basedOn` chains plus document run defaults;
- direct numbered and bulleted lists;
- fixed-layout tables, including supported merge geometry and bounded direct table formatting;
- external/internal hyperlinks and simple safe fields such as `PAGE` and `NUMPAGES`;
- classic whole-paragraph comments with fixed-topology source-bound edits;
- PNG/JPEG inline images with alt text and explicit dimensions;
- page geometry, margins, orientation, section breaks, first/even/default headers and footers, and header/footer fields.

Imported constructs outside this profile remain attached to their validated source package. Leave them unchanged to preserve them. Attempts to author or edit unsupported constructs must fail explicitly; never flatten them into approximate content.

## Author or import

Create a document with the public model:

```js
import { DocumentFile, DocumentModel } from "open-office-artifact-tool";

const document = DocumentModel.create({
  name: "Readiness brief",
  defaultRunStyle: { fontFamily: "Aptos", fontSize: 22 },
});

document.addParagraph("Readiness brief", { styleId: "Title" });
document.addParagraph("All required gates passed.", {
  styleId: "Normal",
  paragraphFormat: { alignment: "left", spaceAfterTwips: 240 },
  runs: [{
    text: "All required gates passed.",
    style: { bold: true, color: "#315A83" },
  }],
});

const output = await DocumentFile.exportDocx(document);
await output.save("out/readiness-brief.docx");
```

Import and make a source-bound edit:

```js
import { DocumentFile, FileBlob } from "open-office-artifact-tool";

const document = await DocumentFile.importDocx(
  await FileBlob.load("in/existing.docx"),
);

const paragraph = document.blocks.find(
  (block) => block.kind === "paragraph" && block.text === "Draft",
);
paragraph.text = "Approved";
if (paragraph.runs.length === 1) paragraph.runs[0].text = "Approved";

const output = await DocumentFile.exportDocx(document);
await output.save("out/approved.docx");
```

Keep block count, order, comment anchors, table merge geometry, and unsupported source-owned objects unchanged during preservation-sensitive edits. A rejected edit is a safety result, not a prompt to rebuild the file through another writer.

## Inspect, patch, and verify

Use model inspection for semantic checks and package inspection for OOXML structure:

```js
const semantic = document.inspect({
  kind: "document,paragraph,listItem,table,comment,header,footer,hyperlink,field,image,section,style,layout",
  maxChars: 20_000,
});

const pkg = await DocumentFile.inspectDocx(output, {
  includeText: true,
  maxChars: 20_000,
});
```

`DocumentFile.patchDocx` is an explicit low-level operation for deliberate OOXML part changes. It is not called automatically by import or export, and it is not an alternate high-level writer. Reinspect and render every patched package.

Run the checked-in end-to-end fixture:

```bash
node skills/documents/scripts/run-fixture.mjs \
  --fixture skills/documents/fixtures/business-brief.json \
  --output-dir tmp/document-skill-fixture \
  --native-render auto
```

Verify any DOCX:

```bash
node skills/documents/scripts/verify-document.mjs \
  --input out/readiness-brief.docx \
  --output-dir tmp/readiness-qa \
  --preview-format png \
  --native-render auto
```

For visual regression work, add `--baseline-dir DIR --write-baseline true` once, then rerun without `--write-baseline`. The workflow compares the model preview and, when native tools are available, LibreOffice PDF plus Poppler page PNGs.

## Checked-in fixtures

- `business-brief.json` covers styles, run and paragraph formatting, lists, a fixed table, a hyperlink, a field, a PNG image, a section, first/even variants, a classic comment, and source-bound edits.
- `open-chestnut-merged-table.json` covers supported physical-grid merge geometry and bounded table-format edits.
- `open-chestnut-numbering-edit.json` covers a fixed direct-numbering group and atomic definition edits.
- `open-chestnut-comments.json` covers classic comment creation and fixed-topology metadata/text edits.
- `package-comments.json`, `package-numbering.json`, and `package-settings.json` retain their historical fixture names but now exercise the supported 0.2 classic-comment, direct-numbering, and section/header settings surfaces. They no longer depend on a second Office implementation.

## Delivery gate

A document is ready only when:

1. semantic inspection contains the intended blocks and values;
2. package inspection reports valid DOCX relationships and parts;
3. semantic verification has no blocking issue;
4. the model preview passes visual QA;
5. LibreOffice and Poppler rendering passes when those tools are available;
6. unsupported imported content is either preserved unchanged or rejected on edit.

The fixture runner writes the DOCX, inspection records, verification records, layout JSON, preview, optional pixel diff, native PDF/page PNGs, and a summary JSON into the selected output directory.
