# Native PivotTables

Use a PivotTable when the workbook should retain a native, refreshable Excel
summary object rather than only a formula-based report. Keep raw source rows in
one rectangular range with one non-empty, unique header row; place the PivotTable
on a separate summary sheet whenever practical.

## Bounded OpenChestnut profile

The current source-free native profile supports:

- 1 through 8 row fields in tabular order, without native compact layout or
  automatic subtotals;
- zero or one column field;
- 1 through 32 value fields, each using `sum`, `count`, `average`, `min`, or
  `max`; multiple values use the native SpreadsheetML data-layout axis;
- optional row and column grand totals;
- zero through nine exact item filters, with at most one on each configured
  row/column field.
  Each uses one non-empty `include` or `exclude` list with at most 1024 string,
  finite-number, boolean, or `null` items;
- `refreshOnLoad`, `saveData`, `enableRefresh`, `invalid`,
  `missingItemsLimit`, `refreshedBy`, and `refreshedDateIso` cache policy;
- cached worksheet values plus native cache records when `saveData` is true;
- inspect, resolve, SVG preview, OpenChestnut export/import, and a second
  byte-preserving export.

```js
const pivot = summary.pivotTables.add({
  name: "Revenue and units by region",
  sourceRange: "Data!A1:E100",
  targetRange: "A1",
  rowFields: ["Region", "Channel"],
  columnFields: ["Product"],
  valueFields: [
    { field: "Revenue", summarizeBy: "sum", name: "Revenue" },
    { field: "Units", summarizeBy: "sum", name: "Units" },
  ],
  filters: [{ field: "Region", exclude: ["North"] }],
  rowGrandTotals: true,
  columnGrandTotals: true,
});
```

`targetRange` may be one anchor cell or the exact cached-output rectangle.
Styling empty target cells before export is supported; any pre-existing value or
formula in the projected output rectangle is a collision and export fails closed.

## Agent workflow

1. Inspect the source range and confirm unique headers and numeric value data.
2. Add the PivotTable with an explicit provider-independent field config.
3. Inspect `kind: "pivotTable"` and compare every metric in `computedValues()`
   with an independent aggregation or source totals.
4. Render the summary range and review labels, number formats, widths, clipping,
   and grand totals.
5. Export, import again, and confirm the native object and computed matrix.
6. Preserve the original source bytes until final verification succeeds.

Run `examples/openchestnut-pivot-table-workflow.mjs` for the complete author,
inspect, render, export, second-import, and verification path.

## Fail-closed boundaries

Grouping, calculated fields, date/condition filters, filters on fields outside
the native axes, more than 8 row fields, more than one column field, compact or
subtotal-bearing multi-row layouts, more than 32 value fields, and source-free
edits inside an imported workbook are not silently flattened.
They remain useful in the JavaScript calculation/preview facade, but native XLSX
export rejects them with an explicit unsupported-profile/filter diagnostic.
An exact filter that names an unknown source item, exceeds the item budget, or
hides every source row also fails closed.

Recognized imported PivotTables expose their semantic configuration for inspect
and resolve, but the native graph, source range values, and cached output are
hash-bound and read-only in this first profile. Unsupported imported PivotTable
graphs remain opaque and unchanged in the validated source package.

Excel-compatible hosts may omit the optional materialized `rowItems` and
`colItems` axis caches when resaving a multi-row or multi-value PivotTable.
OpenChestnut still recognizes that host-normalized graph when the ordered row
fields, canonical `x=-2` data-layout field when needed, ordered data fields,
cache source, field indexes, and relationships all validate. A present but
inconsistent item list, a compact/subtotal-bearing multi-row field, or a
missing/duplicate data-layout field remains opaque and unchanged.

Exact native filters use standard `pivotField/items/item@h` visibility. A host
may normalize an `include` list to the equivalent complementary `exclude` list;
import reports the same visible item set rather than promising preservation of
the caller's original syntactic mode. The shipped workflow is resaved through
LibreOffice and must retain both the semantic filter and cached totals before a
byte-preserving second export.
