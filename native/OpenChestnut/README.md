# OpenChestnut

OpenChestnut is this repository's clean-room C# codec for XLSX, DOCX, and PPTX. It uses the public Microsoft Open XML SDK and the independently defined protocol in `proto/open_office/artifact/v1/office_artifact.proto`.

It is the only Office codec used by version 0.2. The JavaScript package supplies the public artifact model and wire adapter; it does not contain a fallback Office parser/writer. PDF is outside OpenChestnut.

## Projects

- `OpenChestnut.Codec` implements package validation and the XLSX/DOCX/PPTX codecs.
- `OpenChestnut.Runtime` exposes the byte-in/byte-out WebAssembly entry point.
- `OpenChestnut.Tests` covers supported creation/import/edit/export profiles, opaque preservation, and fail-closed cases.

`OpenChestnut.Runtime` keeps its one `byte[] -> byte[]` JavaScript entry point
as a small explicit registration shim in `OpenChestnutJavaScriptInterop.cs`. This avoids
the .NET 8 `[JSExport]` generator's process-random wrapper-name hash while
retaining the same public WASM ABI. The paired clean-build gate verifies the
generated runtime file inventory and hashes rather than merely trusting an
incremental build.

Protocol version 2 removes `allow_lossy`; the removed field name and number are reserved. Opaque content can be exported only from a validated, hash-bound source package. Unsupported edits, source-evidence mismatch, unsafe OPC paths, invalid relationships/content types, and missing runtime data return structured failures.

## Supported 0.2 slices

- XLSX: cells/formulas/styles, merges/dimensions/freeze panes, tables, PNG/JPEG images, bar/line/pie/standard-area/50%-doughnut charts plus marker-only numeric-X/Y scatter and bounded numeric-X/Y/positive-Size 2D bubble charts, standard Office 2010 line/column/stacked sparklines with reversible range mappings, canonical one-variable/two-variable What-If data tables, a bounded native PivotTable profile, basic data validation and conditional formatting, and bounded Office 2019 threaded comments with one root plus direct replies. Pivot authoring owns the workbook/worksheet/Pivot/cache relationship graph for 1 through 8 ordered tabular row fields without automatic subtotals, zero or one column field, 1 through 32 sum/count/average/min/max value fields, optional grand totals, exact typed include/exclude item filters on those axes, and optional saved cache records; multiple values use the canonical `x=-2` data-layout axis. Recognized imports are source/hash/topology-bound and read-only, including host-normalized graphs that omit the optional `rowItems`/`colItems` materialized axis caches while retaining the canonical fields or normalize an item include list to its equivalent exclude complement. Compact/subtotal-bearing multi-row and multi-column graphs remain opaque. Marker-only scatter emits an explicit no-fill series outline for cross-host fidelity and reserves marker borders for the marker style. Bubble authoring uses canonical 2D, 100%-scale, area sizing; imported non-marker scatter, 3D/negative/custom-scale/non-area bubble profiles, richer Pivot graphs, and data-table topology are source-bound/read-only. Nested/branched replies, mentions, orphan/cross-cell parents, and extended identity graphs remain opaque/source-bound.
- DOCX: styles, paragraph/run formatting, pages/sections, headers/footers, PAGE/simple fields, canonical inline five-run SEQ/REF/PAGEREF fields with an optional narrow bookmark around a SEQ cached result, one canonical one-paragraph complex TOC placeholder plus `updateFields` intent, PNG/JPEG images, lists, links, classic comments plus bounded modern root/direct-reply threads with resolved/durable/UTC/person metadata, fixed-geometry tables, bounded whole-block bookmarks/internal hyperlinks, plain-text paragraph/list-item footnotes/endnotes, standalone whole-paragraph `w:ins`/`w:del` revisions, canonical inline plain-text `w:sdt` controls, and one bounded bibliography `b:Sources` catalog with whole-paragraph `CITATION` fields. A separately advertised `text_patchable` capability permits one source-hash-bound literal replacement in exactly one ordinary direct `w:r/w:t` node of an otherwise read-only paragraph or complex table cell; duplicate/cross-node matches, hyperlinks, fields, controls, revisions, capability tampering, residual-topology changes, and whole-paragraph/cell replacement fail closed. Imported recognized modern threads permit text and resolved-state edits while paragraph/durable/person/UTC identity, root anchors, and topology remain source-bound. Imported canonical inline-field cached results, TOC placeholders, note bodies, content-control values, bibliography source content, and citation display text permit fixed-topology edits while field positions/instructions/bookmark identity, anchors, native identity, source order, and tags remain source-bound; imported whole-block bookmarks remain read-only. Nested/irregular modern comments, automatic field materialization, refreshed cross-paragraph TOCs, and other rich or irregular content-control/bibliography/field/bookmark/note/revision graphs remain opaque/source-bound or explicit package workflows.
- PPTX: direct RGB/theme solid and native style-reference slide backgrounds, text boxes/round rectangles, basic fill/line/shadow, straight/polyline connectors and arrows, source-free bar/line/pie charts, bounded embedded rectangular images with signed DrawingML `a:srcRect` crop, fixed tables, recursive native `p:grpSp` trees with local `chOff/chExt` coordinates, rich text/lists/links, canonical plain-text speaker notes, legacy slide annotations, bounded Office 2021 modern root/direct-reply threads with top-level element or shape-text-range anchors, source-bound Master/Layout preservation, and read-only presentation view metadata (grid spacing, snap flags, and horizontal/vertical guides). Recognized imported modern threads permit existing text/status edits while author/date identity, anchor/range, position, topology, relationships, and source hashes remain fixed; reactions/task fields, nested replies/anchors, rich text, mixed families, and connected comment parts remain opaque/source-bound. Recognized imported direct backgrounds, source rectangles, canonical fixed-topology groups, and the presence-aware `p:sp/@useBgFill` flag are projected without flattening inheritance or surrounding package graphs; `useBgFill` is preview-visible but read-only and source-bound. Eligible top-level OLE objects also allow one source/hash/relationship-bound XLSX payload replacement: the replacement workbook is budgeted, opened by the Open XML SDK, and Office 2021 validated while the OLE shell, preview, and unrelated native graph remain unchanged. JavaScript computes `contain`/`cover` into the native source rectangle before crossing the wire and owns local gridline/guide visibility; the native `viewProps.xml` graph remains source/hash-bound. Advanced background/group/view graphs, shared or ambiguous OLE payloads, external or masked/effect-bearing pictures, and rich or irregular notes remain opaque-preserved and fail closed on mutation.
- PPTX imported placeholder text: a concrete SlidePart placeholder keeps `editable=false` for the surrounding shape and carries a separate `text_editable` source claim only when its local `p:txBody` is fully recognized. Export accepts textual-content deltas with identical paragraph/inline kinds and formatting, then applies them in place through the Open XML SDK. Placeholder identity, geometry, style, layout binding, and unmodeled XML remain hash-bound; capability tampering, formatting changes, topology changes, fields outside the bounded request, and unsupported native text graphs fail closed.

The PPTX combo profile is intentionally small: literal clustered bars plus literal
lines, globally ordered series, duplicated chart-level data labels, and either
one shared primary category/value axis pair or one canonical secondary top/right
pair for every line. Bars stay primary; all lines must share one axis group.
External data, mixed line groups, secondary bars, per-series labels/point
overrides, smooth lines, trendlines, error bars, and irregular mixed plots remain
opaque-preserved or fail closed.

Imported objects outside these modeled profiles remain opaque and unchanged. They do not become permission for lossy rewriting.

## Build and test

```sh
dotnet test native/OpenChestnut/OpenChestnut.sln --configuration Release
npm run build:open-chestnut
npm run verify:open-chestnut-build
```

The bundled consumer runtime is written to `runtime/open-chestnut`. The npm package includes that runtime, manifest, SBOM, and license notices, but excludes this C# source tree and build output. Installed consumers therefore do not need a local .NET SDK.
