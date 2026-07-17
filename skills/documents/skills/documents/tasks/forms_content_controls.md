# Task: Forms / content controls (SDTs)

## When to use

Use this workflow to create an Agent-fillable DOCX template or populate an
existing template that contains Word structured document tags (SDTs).

Choose the route before editing:

- Use public `paragraph.addTextContentControl(...)`,
  `document.contentControls`, and `document.fillContentControls(...)` for
  source-free body paragraphs and recognized imported inline plain-text
  controls.
- Use `scripts/content_controls.py` only for explicit package work such as
  wrapping placeholders in an existing template, controls in headers/footers,
  or inspection of controls outside the bounded public model.
- Detect rich, block, cell, nested, data-bound, dropdown, date, checkbox,
  placeholder-document, and locked controls. Preserve them unchanged or fail
  closed; do not flatten them into plain text.

## Public API golden path

### Author one inline plain-text control

```js
import {
  DocumentFile,
  DocumentModel,
  FileBlob,
} from "open-office-artifact-tool";

const document = DocumentModel.create({ blocks: [] });
const customer = document.addParagraph("Customer: ");
customer.addTextContentControl("{{CUSTOMER_NAME}}", {
  id: "customer-name",
  tag: "CUSTOMER_NAME",
  alias: "Customer name",
  style: { bold: true },
});

const template = await DocumentFile.exportDocx(document);
await template.save("template.docx");
```

OpenChestnut assigns the package-local native `w:id` when it is omitted and
authors canonical run-level `w:sdt` markup with `w:text`, `w:tag`, and
`w:alias`. The model `id` is an Agent locator for the current object graph; it
is not a persistent identity across independent imports.

You may also describe the control in a paragraph's initial `runs` array:

```js
document.addParagraph("", {
  runs: [
    { text: "Account: " },
    {
      text: "{{ACCOUNT_ID}}",
      contentControl: {
        id: "account-id",
        tag: "ACCOUNT_ID",
        alias: "Account ID",
      },
    },
  ],
});
```

### Inspect and fill by tag

```js
const imported = await DocumentFile.importDocx(
  await FileBlob.load("template.docx"),
);

for (const control of imported.contentControls) {
  console.log({
    id: control.id,
    targetId: control.targetId,
    runIndex: control.runIndex,
    tag: control.tag,
    alias: control.alias,
    nativeId: control.nativeId,
    text: control.text,
  });
}

const result = imported.fillContentControls({
  CUSTOMER_NAME: "Ada Lovelace",
  ACCOUNT_ID: "AC-2048",
});
if (result.missingTags.length) throw new Error("Template fields are missing");

const filled = await DocumentFile.exportDocx(imported);
await filled.save("filled.docx");
```

`fillContentControls()` is transactional for unknown tags by default: it
checks every requested tag before changing any text. Duplicate controls with
the same tag are all filled. Pass `{ strict: false }` only when partial
template population is intentional; inspect `missingTags` in the returned
result.

For a single recognized control, mutate `control.text`, `control.tag`, or
`control.alias`. Imported control topology and native identity are source-bound:
adding, removing, or reordering a recognized imported control fails closed.

### Verify and render

```js
const inspection = imported.inspect({
  kind: "document,paragraph,contentControl",
  maxChars: 12_000,
});
if (!inspection.ndjson.includes("CUSTOMER_NAME")) {
  throw new Error("Expected content control was not imported");
}

const verification = imported.verify({ visualQa: true });
if (!verification.ok) throw new Error(verification.ndjson);
```

Export and re-import once more, then render every page:

```bash
python render_docx.py /mnt/data/filled.docx --output_dir /mnt/data/out_forms
```

Inspect every `page-<N>.png` at 100% zoom.

## Existing-template package route

Use this route when the user provides a template whose placeholder text must be
wrapped into SDTs, or when controls live in headers/footers or other parts that
the public body-inline model does not edit.

Keep placeholders such as `{{NAME}}`, `{{DATE}}`, and `{{EMAIL}}` contiguous in
one text run, then wrap them:

```bash
python scripts/content_controls.py /mnt/data/template.docx wrap_placeholders \
  --output /mnt/data/template_sdt.docx
```

Populate matching plain-text controls by tag:

```bash
python scripts/content_controls.py /mnt/data/template_sdt.docx fill \
  --set NAME="Ada Lovelace" \
  --set EMAIL="ada@example.com" \
  --output /mnt/data/filled.docx
```

List part location, tag, alias, and visible text for diagnosis:

```bash
python scripts/content_controls.py /mnt/data/template_sdt.docx list --json
```

This helper is an explicit package patch, not a silent fallback. Protect the
source file, write a distinct output, structurally inspect the result, re-import
through OpenChestnut where possible, and render again.

## Pitfalls and fail-closed boundaries

- Placeholder text split across Word runs may not be wrapped by the helper.
  Retype the token contiguously or perform a reviewed narrow package patch.
- The public model recognizes only one run-level plain-text `w:sdt` containing
  one supported text run and canonical `w:sdtPr` metadata.
- Rich, block, cell, nested, data-bound, dropdown, date, checkbox,
  placeholder-document, locked, or extension-bearing controls remain opaque
  and source-bound. Do not reconstruct them as ordinary text.
- Controls in footnotes, comments, headers, footers, and text boxes are outside
  the public body-inline profile. Route them explicitly or report the boundary.
- Never claim a template is fully populated until requested tags, native
  structure, second-import semantics, and all rendered pages pass review.

## Deliverables

Deliver only the requested final DOCX. PNGs and optional PDFs are QA artifacts
unless the user explicitly requests them.
