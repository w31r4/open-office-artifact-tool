# Imported headers and footers

## Agent contract

Use the public OpenChestnut route only when the imported page-furniture record
explicitly proves it is editable:

```text
intent -> import -> inspect/resolve one header or footer -> capability check
       -> one text edit -> export -> package-scope check -> second import
       -> document.verify -> native page render -> deliver
```

Do not infer editability from text visible at the top or bottom of a page. Word
headers and footers can be shared, inherited, field-driven, rich, or attached
to a section relationship graph. The public model deliberately exposes only a
small, source-bound edit profile.

## Create new page furniture

For a source-free document, use the ordinary public API. First/even variants
activate the necessary document setting when they are active. A single-field
footer can use the legacy `fieldInstruction` shorthand; use ordered segments
when literal text and two or more field displays must occupy one native
paragraph.

```js
document.addHeader("Decision brief | Internal", {
  referenceType: "default",
  sectionIndex: 0,
});
document.addFooter([
  { text: "Page " },
  { field: { instruction: "PAGE", display: "1" } },
  { text: " of " },
  { field: { instruction: "NUMPAGES", display: "1" } },
], {
  referenceType: "default",
  sectionIndex: 0,
});
```

The array must contain 2 through 32 ordered `{ text }` or
`{ field: { instruction, display } }` items, at least one supported simple
field, and no run formatting. `text` is derived from the concatenated literal
and cached field displays; use `footer.setSegments(nextSegments)` to replace a
source-free sequence atomically. Do not combine segments with
`fieldInstruction`, and do not assign `footer.text` directly while segments are
present. A compatible pagination host owns the live PAGE/NUMPAGES result; the
`display` values are the authored fallback and must be rendered/reviewed before
delivery.

This is source-free authoring only. An imported structured field paragraph is
reported for inspect/resolve and preserved exactly on a no-op export, but is
source-bound/read-only. It does not make imported PAGE or NUMPAGES fields
ordinary editable header/footer text.

## Edit one imported text paragraph

```js
import { DocumentFile, FileBlob } from "open-office-artifact-tool";

const source = await FileBlob.load("input.docx");
const document = await DocumentFile.importDocx(source);
const candidates = document.headers.filter((item) =>
  item.referenceType === "default" && item.text === "Northwind | Internal");
if (candidates.length !== 1) {
  throw new Error(`Expected one header target, found ${candidates.length}.`);
}

const header = candidates[0];
if (!header.sourceBound || !header.editable) {
  throw new Error("The selected header has no safe source-bound text-edit capability.");
}
if (document.resolve(header.id) !== header) {
  throw new Error("Header locator did not resolve.");
}

header.text = "Northwind | Reviewed";
const output = await DocumentFile.exportDocx(document);
await output.save("reviewed.docx");

const reimported = await DocumentFile.importDocx(await FileBlob.load("reviewed.docx"));
const verified = reimported.headers.filter((item) => item.text === "Northwind | Reviewed");
if (verified.length !== 1 || !verified[0].sourceBound || !verified[0].editable) {
  throw new Error("Header text did not survive the source-bound round-trip.");
}
if (!reimported.verify({ visualQa: true }).ok) {
  throw new Error("Document verification failed.");
}
```

Use `document.footers` in exactly the same way. `document.inspect({ kind:
"header,footer" })` records `sourceBound`, `editable`, section/reference
scope, relationship ID, and part path so an agent can state why it selected or
rejected a candidate.

For the ordinary imported **header** or **footer** profile, prefer the matching
packaged transaction instead of reproducing its package proof by hand:

```bash
node examples/openchestnut-header-text-edit-workflow.mjs \
  input.docx reviewed.docx audit.json \
  "Northwind | Internal" "Northwind | Reviewed" 0 default
node examples/openchestnut-footer-text-edit-workflow.mjs \
  input.docx reviewed.docx audit.json \
  "Northwind | Internal" "Northwind | Reviewed" 0 default
```

Each requires distinct absent output/audit paths, checks exactly one selected
item, exports through OpenChestnut, permits only its matching
`word/headerN.xml` or `word/footerN.xml` part, normalizes the unique target
`w:t` to prove the part residual, reimports, verifies, model-renders, and
records a byte-bound audit. The two entry points are intentionally separate:
do not use either one to reinterpret PAGE/simple fields or the opposite page
furniture kind as ordinary text.

## Exact editable profile

An imported item is editable only when all of these remain true at export:

- Its HeaderPart or FooterPart is uniquely used after effective section
  inheritance is evaluated.
- The target is one direct ordinary `w:p > w:r > w:t` paragraph with no run
  formatting and no field.
- Its source relationship, part path, paragraph locator, ID/name/style,
  section/reference/variant state, field state, and source hashes still match.
- At most one text edit may target a given source part; no other header/footer
  paragraph in that part may change in the transaction.

The export re-proves element, semantic, paragraph-residual, and
part-residual hashes before modifying the existing `w:t`. It writes no new
relationship and must change exactly the selected `word/headerN.xml` or
`word/footerN.xml` part.

## Fail-closed boundaries

Do not use this profile for:

- PAGE or other simple fields, complex fields, rich/multi-run text, drawings,
  controls, tables, or extension-bearing paragraphs;
- shared, linked, or inherited Header/Footer parts;
- two text edits in one source part, even if both look ordinary;
- changing first/even/default scope, section ownership, style, identity,
  relationship, part path, variant activation, field instruction, or topology;
- adding/removing imported page furniture or rebuilding it from visible text.

If a capability check fails, preserve the source and either narrow the request,
use an explicitly reviewed package-level workflow, or report the refusal. The
public codec must fail closed; do not silently create a replacement
header/footer.

## Watermark coexistence and QA

A recognized canonical text-watermark edit may share a header part with one
eligible header-text edit: both operations independently prove their target and
the remaining header content. This does **not** broaden the profile to image,
DrawingML, shared, or irregular watermarks.

Before delivery:

1. Preserve the source and publish to a different output path.
2. Inspect the ZIP content diff; permit exactly one target Header/Footer part.
3. Import the output again and check the exact text, `sourceBound`, and
   `editable` evidence.
4. Run `document.verify({ visualQa: true })`.
5. Render with `render_docx.py` and inspect every affected page at 100%; check
   page variants, field display, body layout, and all other page furniture.
