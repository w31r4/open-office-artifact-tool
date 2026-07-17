# Task: Fields + update behavior (TOC / page # / refs)

## Goal
Avoid "looks wrong" renders that are actually **stale Word fields**.

Common fields:
- `PAGE` — current page number
- `NUMPAGES` — total page count
- `TOC` — table of contents
- `REF` / `PAGEREF` — cross references (often "see page X")

## When this matters
- PDF/PNG render shows placeholders (e.g., TOC looks empty, refs show wrong page, page numbers all “1”).
- The doc was modified programmatically (python-docx / OOXML patch) and then exported without a field refresh.
- LibreOffice vs Word disagree.

## What to do
### 1) Scan for fields
Run a quick field inventory:

```bash
python scripts/fields_report.py /mnt/data/input.docx
```

If you see `TOC`, `REF`, `PAGEREF`, `NUMPAGES`, or `PAGE`, plan for a field refresh step.

For a source-free native TOC, prefer the public OpenChestnut path:

```js
document.addTableOfContents({ levels: "1-3" });
```

For canonical inline numbering and references, also prefer the public run path:

```js
paragraph.addField("SEQ Figure \\* ARABIC", "0", { bookmarkName: "fig1" });
paragraph.addField("REF fig1 \\h", "0");
paragraph.addField("PAGEREF fig1 \\h", "0");
```

The second argument is only cached display text. `bookmarkName` on a `SEQ`
field wraps only that cached result, giving `REF`/`PAGEREF` a real caption
target. After importing a canonical field paragraph, the cache and ordinary
run text may change, but field position, instruction, and bookmark identity
remain source-bound.

This sets `document.settings.updateFields = true` and emits native
`w:updateFields`. Treat it only as a host-refresh request. It never means the
cached result is current, and headless export alone is not proof that page
numbers match.

### 2) Render and inspect

```bash
python render_docx.py /mnt/data/input.docx --output_dir /mnt/data/out
```

Inspect all `page-*.png` at 100% zoom.

### 3) If anything is wrong: update fields in a GUI editor
**Fast checklist (Word):**
1. Open the DOCX in **Microsoft Word**
2. `Ctrl+A` (select all)
3. `F9` (Update Fields)
4. Save
5. Re-render with `render_docx.py`

LibreOffice (GUI) can also update fields, but Word is the reference implementation.

## Deterministic rendering workaround (when you can't update fields)
If your goal is **stable PNG regression testing** (not perfect Word semantics), you can
*materialize* some field results into literal text so headless renders won't omit them:

```bash
# Replace REF/PAGEREF blocks with their currently cached visible text
python scripts/flatten_ref_fields.py input.docx --out ref_flattened.docx

# Materialize SEQ/REF results (e.g., caption numbers / cross-refs)
python scripts/fields_materialize.py ref_flattened.docx --out fields_materialized.docx
```

Notes:
- This does **not** refresh TOC/PAGE/NUMPAGES; those still typically require Word/LO GUI.
- Always render and visually verify after materialization.

## Render → PNG review checklist (fields)
- Page numbers increment correctly (footer/header)
- Total page count (`NUMPAGES`) matches the rendered page count
- TOC entries exist, have correct indentation, and page numbers match headings
- Cross references (`REF`/`PAGEREF`) resolve (no "Error! Reference source not found.")
- No placeholder text like “(TOC will populate...)” remains
