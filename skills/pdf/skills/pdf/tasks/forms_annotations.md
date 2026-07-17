# Forms and annotations

Use MuPDF.js for bounded text/choice/checkbox form values and text annotations. Use pypdf when radio export values, appearance-state validation, flattening, or more complex AcroForm handling is required. Always open the original PDF directly.

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

Radio buttons fail closed in this path because the API does not expose a trustworthy widget-to-export-value mapping. Signed-PDF incremental edits are also rejected.

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
