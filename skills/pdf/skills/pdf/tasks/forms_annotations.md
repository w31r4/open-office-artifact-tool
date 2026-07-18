# Forms and annotations

Use MuPDF.js for bounded source-bound single-widget text/combo/checkbox updates and source-bound Text-note pins. Use pypdf when radio export values, shared widgets, choice display/export mappings, appearance-state validation, flattening, or more complex AcroForm handling is required. Always open the original PDF directly.

## Inspect first

```bash
python3 scripts/pypdf_edit.py inspect input.pdf \
  --output tmp/pdfs/pypdf-inspect.json
```

Check field hierarchy, widget pages, current values, annotations, encryption, signatures, and DocMDP before mutation.

For a supported MuPDF.js field or text note:

```bash
node scripts/mupdf.mjs probe
node scripts/mupdf.mjs inspect input.pdf
node scripts/mupdf.mjs edit input.pdf tmp/pdfs/form-operations.json tmp/pdfs/filled.pdf \
  --save-policy rewrite
```

Native inspection emits individual `mupdfWidget` records and groups them into
`mupdfFormField` records. For an agent-safe direct field update, select one
field record by semantic name/type/value, then copy **both** the inspection
`summary.sourceSha256` and that record's `id`/`snapshot`. Do not select by
array position or field name alone:

```js
const inspection = await PdfFile.inspectPdf(input);
const field = inspection.records.find((record) => record.kind === "mupdfFormField"
  && record.name === "sender.city");
if (!field?.snapshot) throw new Error("Expected one inspectable city field.");

const edited = await PdfFile.editPdf(input, {
  savePolicy: "incremental",
  operations: [{
    type: "update_form_field",
    sourceSha256: inspection.summary.sourceSha256,
    formFieldId: field.id,
    expected: field.snapshot,
    value: "Shanghai",
  }],
});
```

`update_form_field` accepts exactly one non-password text widget, one
non-multiselect combo whose display and export options are identical, or one
checkbox. The complete snapshot protects name/type/current value/read-only
state/options/visible widget geometry. It verifies the field state before save,
but it is not a durable field identity: re-inspect the output before any second
mutation. It may use unsigned `incremental` save and proves the exact source
prefix; it still does not authorize signed changes.

Radio buttons, shared-widget fields, list or multi-select choices, password
fields, mismatched export values, stale snapshots, and unsupported options fail
closed in this path. Route them to the explicit pypdf workflow below. Signed
PDF incremental edits are also rejected.

## Add one source-bound Text note

Select the target `mupdfPage` record from the same inspection. Use its `bbox`
and `rotation` as an exact coordinate precondition, then place one Text-note
pin with non-empty `contents`:

```js
const page = inspection.records.find((record) => record.kind === "mupdfPage"
  && record.page === 1);
if (!page) throw new Error("Expected an inspectable first page.");

const annotated = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_text_annotation",
    page: page.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: page.bbox, rotation: page.rotation },
    point: [72, 128],
    contents: "Review this assumption.",
    author: "Reviewer",
  }],
});
```

`point` is in the inspected unrotated visible `mupdfPage.bbox` coordinate
space. It is not a request for a specific note rectangle: the provider
normalizes the native icon geometry, verifies exactly one new Text annotation,
and records the actual rectangle in the operation audit. A `text` alias,
`bbox`/`rect`, icon selection, rotated page, stale hash/page snapshot, or
incremental save fails closed. Re-inspect the rewrite before a later
annotation update/deletion; the returned xref is current-source-only.

Before a pypdf mutation, probe and bind the exact route. Change `--task` to `annotate` for notes:

```bash
python3 scripts/pdf_provider.py check --provider pypdf --require
python3 scripts/pdf_provider.py plan \
  --task fill-form --provider pypdf --strategy incremental \
  --input input.pdf --output tmp/pdfs/filled.pdf --require-provider
```

## Fill form with pypdf

```bash
python3 scripts/pypdf_edit.py fill-form input.pdf tmp/pdfs/filled.pdf \
  --strategy incremental \
  --field 'sender.city=Shanghai' \
  --field 'approved=Yes'
```

The script sets `auto_regenerate=False` so the output carries explicit field state rather than asking the viewer to regenerate it. Use `--flatten` only with `rewrite`, after confirming that interactivity should be removed.

The adapter resolves each field type before mutation. Text and choice values remain strings; radio buttons and checkboxes are matched against their real `/AP /N` appearance-state names and written as PDF Names. Unknown button states, read-only fields, signature fields, push buttons, unsupported field types, missing appearances, or a post-write `/V`/`/AS` mismatch fail closed and remove the transactional output. This prevents a radio value from looking filled in field metadata while every widget still renders `/Off`.

## Add annotation with pypdf

```bash
python3 scripts/pypdf_edit.py add-note input.pdf tmp/pdfs/annotated.pdf \
  --strategy incremental \
  --page 1 --rect 72,640,96,664 \
  --text 'Review this assumption.'
```

The optional PyMuPDF specialist script also exposes `add_text_annotation` and `fill_form`, but it is selected explicitly rather than used as a fallback.

## Signed forms

An incremental update can retain signed byte ranges, but it can still violate DocMDP or a field lock. The script refuses signed inputs unless `--allow-signed` is explicit. Run pyHanko validation before and after and compare the reported modifications.

Record the canonical [`open-office-artifact-tool.pdf-audit.v1`](../references/AUDIT_SCHEMA.md) envelope and run `scripts/pdf_audit.py validate` against the exact source and delivered artifact before handoff.
