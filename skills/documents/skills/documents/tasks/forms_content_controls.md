# Task: Forms / content controls (SDTs)

## When to use

Use this workflow to create an Agent-fillable DOCX template or populate an
existing template that contains Word structured document tags (SDTs).

Choose the route before editing:

- Use public `paragraph.addTextContentControl(...)`,
  `paragraph.addCheckboxContentControl(...)`,
  `paragraph.addDropdownContentControl(...)`,
  `paragraph.addComboBoxContentControl(...)`,
  `paragraph.addDateContentControl(...)`, `document.contentControls`,
  `document.fillContentControls(...)`, `document.setCheckboxContentControls(...)`,
  `document.setDropdownContentControls(...)`, and
  `document.setComboBoxContentControls(...)`, and
  `document.setDateContentControls(...)` for source-free body
  paragraphs and recognized imported inline plain-text, canonical Word 2010+
  checkbox, canonical Word drop-down, canonical Word combo-box, or canonical
  ISO/Gregorian date controls.
- Use `scripts/content_controls.py` only for explicit package work such as
  wrapping placeholders in an existing template, controls in headers/footers,
  or inspection of controls outside the bounded public model.
- Detect rich, block, cell, nested, data-bound, irregular list-control,
  localized/noncanonical date, legacy or custom-symbol checkbox,
  placeholder-document, and locked
  controls. Preserve them unchanged or fail closed; do not flatten them into
  plain text.

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

### Author and set one checkbox control

Use the typed checkbox primitive instead of writing a Unicode box yourself:

```js
const terms = document.addParagraph("Terms accepted: ");
terms.addCheckboxContentControl(false, {
  id: "terms-accepted",
  tag: "TERMS_ACCEPTED",
  alias: "Terms accepted",
});

const update = document.setCheckboxContentControls({
  TERMS_ACCEPTED: true,
});
if (update.missingTags.length) throw new Error("Required checkbox is missing");
```

OpenChestnut owns the visible `☐`/`☒` glyph and the exact
`w14:checkbox`, `w14:checked`, `w14:checkedState`, and
`w14:uncheckedState` declarations. The public value is boolean. Supplying or
editing the visible glyph directly fails closed.

### Author and select one drop-down control

Use one ordered choice table. `displayText` is what Word shows; `value` is the
stable value the Agent sends:

```js
const priority = document.addParagraph("Priority: ");
priority.addDropdownContentControl([
  { displayText: "Low", value: "low" },
  { displayText: "Medium", value: "medium" },
  { displayText: "High", value: "high" },
], {
  id: "priority",
  tag: "PRIORITY",
  alias: "Priority",
  selectedValue: "medium",
});

const selection = document.setDropdownContentControls({ PRIORITY: "high" });
if (selection.missingTags.length) throw new Error("Required drop-down is missing");
```

OpenChestnut authors canonical `w:dropDownList` / `w:listItem` markup. The
public mutable state is only `selectedValue`; visible run text is derived from
the matching `displayText`. Unknown tags and values outside the declared table
fail before mutation. A string choice is shorthand for identical display text
and value.

### Author and set one combo-box control

Use a combo box only when the user may choose a known value **or** type a
bounded custom value. Use a drop-down when values must remain a strict enum:

```js
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

const custom = document.setComboBoxContentControls({
  CONTACT_METHOD: "Pager duty",
});
if (custom.missingTags.length) throw new Error("Required combo box is missing");
```

OpenChestnut authors canonical `w:comboBox` / `w:listItem` markup. The mutable
`value` may match a declared internal value, in which case Word shows that
item's `displayText`, or it may be XML-safe custom text of 1–255 characters,
which is shown verbatim. Choice order and identity remain source-bound after
import. Unknown tags, empty/control-bearing values, and direct visible-text
edits fail before mutation.

### Author and set one date control

Use a date control when the value is a calendar date, not an instant or a
locale-formatted label. The public value is always an exact `YYYY-MM-DD`
string:

```js
const review = document.addParagraph("Review date: ");
review.addDateContentControl("2026-07-21", {
  id: "review-date",
  tag: "REVIEW_DATE",
  alias: "Review date",
});

const dateUpdate = document.setDateContentControls({
  REVIEW_DATE: "2028-02-29",
});
if (dateUpdate.missingTags.length) throw new Error("Required date is missing");
```

OpenChestnut writes a single-run `w:date` with UTC-midnight `w:fullDate`,
`yyyy-MM-dd` display mask, `en-US` language, `date` mapped-data storage, and a
Gregorian calendar. The visible run text is codec-owned and equals
`dateValue`. Invalid leap days, JavaScript `Date` objects, timezone-bearing
timestamps, localized strings, direct run-text edits, and unknown tags fail
before mutation. This intentionally avoids machine-locale and timezone drift.

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
    controlType: control.controlType,
    text: control.text,
    checked: control.checked,
    choices: control.choices,
    selectedValue: control.selectedValue,
    value: control.value,
    dateValue: control.dateValue,
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

For a single recognized text control, mutate `control.text`; for a checkbox,
mutate `control.checked`; for a drop-down, mutate `control.selectedValue`; for
a combo box, mutate `control.value`; for a date, mutate `control.dateValue`.
All types allow `control.tag` and `control.alias`. Imported control type,
drop-down/combo-box choices and order, symbol declaration, native date profile,
topology, and native identity are source-bound: adding, removing, reordering,
redefining, or converting a recognized imported control fails closed.

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
- The public model recognizes one run-level plain-text `w:sdt`, one canonical
  Word 2010+ `w14:checkbox`, one canonical `w:dropDownList`, or one canonical
  `w:comboBox`, or one canonical ISO/Gregorian `w:date`, each containing exactly
  one supported run and canonical `w:sdtPr` metadata. Both list controls are bounded to 1–256 unique
  display/value pairs of at most 255 characters; combo-box custom values are
  bounded to 1–255 characters. Dates must be real `0001-01-01` through
  `9999-12-31` Gregorian dates in exact `YYYY-MM-DD` form.
- Rich, block, cell, nested, data-bound, irregular drop-down/combo-box, localized-date,
  legacy checkbox, custom-symbol checkbox, placeholder-document, locked, or
  unrelated extension-bearing controls remain opaque and source-bound. Do not
  reconstruct them as ordinary text or a canonical control.
- Controls in footnotes, comments, headers, footers, and text boxes are outside
  the public body-inline profile. Route them explicitly or report the boundary.
- Never claim a template is fully populated until requested tags, native
  structure, second-import semantics, and all rendered pages pass review.

## Deliverables

Deliver only the requested final DOCX. PNGs and optional PDFs are QA artifacts
unless the user explicitly requests them.
