# Forms and annotations

Use pypdf for basic AcroForm and annotation operations, or PyMuPDF when widget/appearance/page integration requires its advanced provider. Always open the original PDF directly.

## Inspect first

```bash
python3 scripts/pypdf_edit.py inspect input.pdf \
  --output tmp/pdfs/pypdf-inspect.json
```

Check field hierarchy, widget pages, current values, annotations, encryption, signatures, and DocMDP before mutation.

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

## Add annotation with pypdf

```bash
python3 scripts/pypdf_edit.py add-note input.pdf tmp/pdfs/annotated.pdf \
  --strategy incremental \
  --page 1 --rect 72,640,96,664 \
  --text 'Review this assumption.'
```

For PyMuPDF, use `add_text_annotation` or `fill_form` operations with `scripts/pymupdf_edit.py`.

## Signed forms

An incremental update can retain signed byte ranges, but it can still violate DocMDP or a field lock. The script refuses signed inputs unless `--allow-signed` is explicit. Run pyHanko validation before and after and compare the reported modifications.

Record the canonical [`open-office-artifact-tool.pdf-audit.v1`](../references/AUDIT_SCHEMA.md) envelope and run `scripts/pdf_audit.py validate` against the exact source and delivered artifact before handoff.
