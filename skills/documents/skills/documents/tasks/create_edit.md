# Task: Create or edit a DOCX

Use this workflow for ordinary document authoring and semantic edits. Read
`../artifact_tool/API_QUICK_START.md` before implementing the builder.

## Default tool: DocumentModel + OpenChestnut

Use the public `open-office-artifact-tool` package for:

- paragraphs, formatted runs, and named styles;
- real numbered or character-bulleted lists;
- fixed-geometry tables;
- sections, headers, footers, hyperlinks, bounded whole-block bookmarks, plain-text footnotes/endnotes, and simple fields;
- PNG/JPEG inline images, classic comments, and bounded modern root/direct-reply threads;
- DOCX import, export, inspect, resolve, verification, and model preview.

`DocumentFile.importDocx(...)` and `DocumentFile.exportDocx(...)` always use the
bundled OpenChestnut C# WebAssembly codec. Do not pass codec selectors, lossy
options, or a Python authoring fallback.

## Create

1. Resolve Node.js and the package directory with the workspace dependency
   loader.
2. Create a writable task directory and an ES module builder.
3. Choose a design preset from `../references/design_presets.md` and translate
   it into explicit `DocumentModel` styles and geometry.
4. Build semantic paragraphs, lists, and tables. Do not fake lists with text
   markers or use tables as prose layout containers.
5. Export through `DocumentFile.exportDocx(...)`, re-import the resulting DOCX,
   and assert the important content with `inspect()` and `verify()`.
6. Run the render loop in `verify_render.md` and inspect every page.

From the Skill root, the packaged runnable example covers the complete vertical
slice:

```bash
node examples/openchestnut-end-to-end.mjs output.docx
```

## Edit an existing document

```js
import {
  DocumentFile,
  FileBlob,
} from "open-office-artifact-tool";

const document = await DocumentFile.importDocx(await FileBlob.load("input.docx"));
const targets = document.blocks.filter(
  (block) => block.kind === "paragraph" && block.text.includes("old wording"),
);
if (targets.length !== 1) throw new Error(`Expected one target paragraph, found ${targets.length}.`);
const target = targets[0];
if (!target.textEditable && !target.textPatchable) {
  throw new Error("Target text is source-bound and has no safe public edit capability.");
}

const range = document.resolve(`${target.id}/text`);
if (!range) throw new Error("Advertised text range did not resolve.");
range.replace("old wording", "replacement wording");
const report = document.verify({ visualQa: true });
if (!report.ok) throw new Error(report.ndjson || JSON.stringify(report.issues));

const output = await DocumentFile.exportDocx(document);
await output.save("edited.docx");
```

For an imported paragraph or table cell, `textEditable` means the whole modeled
text value can be assigned without changing its verified topology.
`textPatchable` is deliberately narrower: `range.replace(old, next)` must use a
non-empty literal that occurs exactly once in one ordinary native `w:t` node or
across adjacent non-empty ordinary runs whose exact `w:rPr` markup matches.
Mixed formatting, empty-run gaps, paragraph boundaries, duplicate matches,
fields, content controls, revisions, and other complex text graphs fail closed.
A table cell exposes the same operation through
`document.resolve(cell.id + "/text")` or
`table.getCell(row, column).replaceText(old, next)`; whole-cell assignment still
fails when `cell.editable` is false.

For an auditable existing-file transaction, use
`../examples/openchestnut-source-text-patch-workflow.mjs`. It binds a structured
paragraph or physical table-cell target, protects the input, permits only
`word/document.xml` to change, publishes without overwrite, reimports, verifies,
and emits an audit JSON before the required native render review.

When a user explicitly needs one true redline inside an existing ordinary
paragraph, do not use the untracked `textPatchable` route. Use
`DocumentFile.addTrackedReplacement(...)` or
`examples/openchestnut-tracked-replacement-workflow.mjs`; bind the exact source
hash, semantic block index, full paragraph snapshot, and unique literal. The
literal may span adjacent non-empty ordinary runs only when their exact `w:rPr`
markup matches; mixed formatting, empty-run gaps, and broader revision graphs
still require the explicit OOXML or Office-host route.

Preserve the original and make minimal, local changes. If OpenChestnut rejects
an edit, narrow the edit or report the unsupported boundary instead of
flattening or rebuilding the file.

## Explicit low-level package patches

The Python and OOXML scripts in this Skill are allowed only when the requested
operation is explicitly package-level and is not an ordinary model edit, for
example accepting imported tracked revisions, applying the Google Docs title
sanitizer, or auditing package relationships. Run the narrow documented script
after OpenChestnut export, then structurally inspect and re-render the result.
Never switch the whole document to `python-docx` because one construct is
unsupported.

## After every meaningful batch

Export, re-import or structurally inspect, render to PNG, and review every page.
Use `verify_render.md`; an SVG/model preview complements native rendering but
does not replace it.

## Output hygiene

Keep the task output directory clean and return only the requested final
deliverable unless the user asks for QA intermediates.
