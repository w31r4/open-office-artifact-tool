# Text watermarks and background elements

## Agent contract

Use the typed OpenChestnut path for a greenfield text watermark or a recognized imported canonical text watermark:

```text
intent -> import/inspect -> resolve one watermark -> typed edit/remove
       -> OpenChestnut -> second import -> verify -> native page render -> export
```

Do not infer editability from a visible diagonal label. Word stores watermark-like content in header parts, but real files may use VML text paths, DrawingML, images, shared header parts, or several objects. The public API deliberately recognizes only an exclusive canonical VML text-watermark paragraph whose surrounding header can be hash-protected.

## Create a canonical text watermark

```js
import { DocumentFile, DocumentModel } from "open-office-artifact-tool";

const document = DocumentModel.create({ name: "Review copy", blocks: [] });
document.addParagraph("Controlled review document.");
document.addWatermark("DRAFT", {
  sectionIndex: 0,
  referenceType: "default",
});

const output = await DocumentFile.exportDocx(document);
await output.save("review-copy.docx");
```

There may be only one modeled watermark per zero-based `sectionIndex` and `referenceType` (`default`, `first`, or `even`). A first-page watermark activates the first-page header variant; an even-page watermark activates even/odd headers.

## Inspect and edit an imported watermark

```js
import { DocumentFile, FileBlob } from "open-office-artifact-tool";

const source = await FileBlob.load("input.docx");
const document = await DocumentFile.importDocx(source);
const records = document.inspect({ kind: "document,watermark" });
console.log(records.ndjson);

const matches = document.watermarks.filter((item) =>
  item.text === "DRAFT" &&
  item.sectionIndex === 0 &&
  item.referenceType === "default");
if (matches.length !== 1) throw new Error(`Expected one watermark, found ${matches.length}.`);
const watermark = matches[0];
if (!watermark.sourceBound || !watermark.editable) {
  throw new Error("The selected background object has no safe canonical watermark edit capability.");
}
if (document.resolve(watermark.id) !== watermark) throw new Error("Watermark locator did not resolve.");

watermark.text = "INTERNAL REVIEW";
const output = await DocumentFile.exportDocx(document);
await output.save("reviewed.docx");
```

Use `watermark.remove()` instead of assigning blank text when the user requests removal. On export, OpenChestnut re-proves the exact header part, paragraph, VML shape, source semantics, element hash, and residual header hash. Text editing changes only the text-path string. Removal deletes only that complete watermark paragraph. Changing its section/reference scope, reordering imported watermarks, or adding a watermark to an imported package fails closed because those operations require unproven relationship or section-topology changes.

For a real imported file, prefer the shipped transactional workflow:

```bash
node examples/openchestnut-watermark-workflow.mjs \
  input.docx reviewed.docx watermark-audit.json \
  DRAFT "INTERNAL REVIEW" 0 default edit
```

Use `remove` as the final argument for whole-object removal. The workflow protects the input, refuses overwrite, requires one exact target, permits only one changed `word/headerN.xml` part, reimports, verifies, model-renders, and writes an audit. It still requires native page review before delivery.

## Irregular and image watermarks

If `document.watermarks` is empty while the rendered document visibly contains a watermark, treat the object as opaque. Do not rebuild the file through `DocumentModel`, delete all header shapes, or silently route to a heuristic edit.

The retained Python helpers are explicit advanced package tools:

```bash
python scripts/watermark_audit_remove.py input.docx --mode report
python scripts/watermark_audit_remove.py input.docx --mode remove --contains DRAFT --out candidate.docx
```

The second command is heuristic and can delete a legitimate header graphic whose text happens to match. Run it only on a protected copy, only after the report has identified the intended object, and only when the user accepts the package-patch route. `scripts/watermark_add.py` likewise remains a compatibility helper for existing packages outside the modeled relationship profile; source-free agent authoring should use `document.addWatermark(...)`.

## Mandatory QA

1. Preserve the input and write a distinct output transactionally.
2. Import the result again and inspect `kind: "watermark"`.
3. Run `document.verify({ visualQa: true })`.
4. Inspect the package diff. The bounded edit/remove workflow permits exactly one `word/headerN.xml` content change.
5. Render with the native helper and inspect every page at 100%:

   ```bash
   python render_docx.py reviewed.docx --output_dir rendered
   ```

6. Confirm the requested watermark text/removal, page count, body layout, and every ordinary header/footer element. LibreOffice may not render every VML/DrawingML profile exactly like Word; for delivery that depends on Microsoft Word appearance, repeat the final check in Word.

## Fail-closed boundaries

- shared header part across section/reference scopes;
- multiple recognized watermark paragraphs in one header part;
- image watermarks, DrawingML watermarks, irregular VML, or objects mixed with unrelated content in the same paragraph;
- scope/topology changes or new relationship authoring in an imported package;
- text that is blank, over 256 characters, or not XML-safe;
- any edit whose exact source or residual hash no longer matches.

These are product boundaries, not permission to flatten the package. Preserve opaque content unchanged or use a separately reviewed explicit package workflow.
