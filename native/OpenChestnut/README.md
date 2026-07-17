# OpenChestnut

OpenChestnut is this repository's clean-room C# codec for XLSX, DOCX, and PPTX. It uses the public Microsoft Open XML SDK and the independently defined protocol in `proto/open_office/artifact/v1/office_artifact.proto`.

It is the only Office codec used by version 0.2. The JavaScript package supplies the public artifact model and wire adapter; it does not contain a fallback Office parser/writer. PDF is outside OpenChestnut.

## Projects

- `OpenChestnut.Codec` implements package validation and the XLSX/DOCX/PPTX codecs.
- `OpenChestnut.Runtime` exposes the byte-in/byte-out WebAssembly entry point.
- `OpenChestnut.Tests` covers supported creation/import/edit/export profiles, opaque preservation, and fail-closed cases.

Protocol version 2 removes `allow_lossy`; the removed field name and number are reserved. Opaque content can be exported only from a validated, hash-bound source package. Unsupported edits, source-evidence mismatch, unsafe OPC paths, invalid relationships/content types, and missing runtime data return structured failures.

## Supported 0.2 slices

- XLSX: cells/formulas/styles, merges/dimensions/freeze panes, tables, PNG/JPEG images, bar/line/pie/standard-area/50%-doughnut charts, standard Office 2010 line/column/stacked sparklines with reversible range mappings, basic data validation and conditional formatting, and bounded Office 2019 threaded comments with one root plus direct replies. Nested/branched replies, mentions, orphan/cross-cell parents, and extended identity graphs remain opaque/source-bound.
- DOCX: styles, paragraph/run formatting, pages/sections, headers/footers, PAGE/simple fields, canonical inline five-run SEQ/REF/PAGEREF fields with an optional narrow bookmark around a SEQ cached result, one canonical one-paragraph complex TOC placeholder plus `updateFields` intent, PNG/JPEG images, lists, links, classic comments, fixed-geometry tables, bounded whole-block bookmarks/internal hyperlinks, plain-text paragraph/list-item footnotes/endnotes, standalone whole-paragraph `w:ins`/`w:del` revisions, canonical inline plain-text `w:sdt` controls, and one bounded bibliography `b:Sources` catalog with whole-paragraph `CITATION` fields. Imported canonical inline-field cached results, TOC placeholders, note bodies, content-control values, bibliography source content, and citation display text permit fixed-topology edits while field positions/instructions/bookmark identity, anchors, native identity, source order, and tags remain source-bound; imported whole-block bookmarks remain read-only. Automatic field materialization, refreshed cross-paragraph TOCs, and other rich or irregular content-control/bibliography/field/bookmark/note/revision graphs remain opaque/source-bound or explicit package workflows.
- PPTX: direct RGB/theme solid and native style-reference slide backgrounds, text boxes/round rectangles, basic fill/line/shadow, straight/polyline connectors and arrows, source-free bar/line/pie charts, bounded embedded rectangular images with signed DrawingML `a:srcRect` crop, fixed tables, recursive native `p:grpSp` trees with local `chOff/chExt` coordinates, rich text/lists/links, canonical plain-text speaker notes, source-bound Master/Layout preservation, and read-only presentation view metadata (grid spacing, snap flags, and horizontal/vertical guides). Recognized imported direct backgrounds, source rectangles, canonical fixed-topology groups, and the presence-aware `p:sp/@useBgFill` flag are projected without flattening inheritance or surrounding package graphs; `useBgFill` is preview-visible but read-only and source-bound. Eligible top-level OLE objects also allow one source/hash/relationship-bound XLSX payload replacement: the replacement workbook is budgeted, opened by the Open XML SDK, and Office 2021 validated while the OLE shell, preview, and unrelated native graph remain unchanged. JavaScript computes `contain`/`cover` into the native source rectangle before crossing the wire and owns local gridline/guide visibility; the native `viewProps.xml` graph remains source/hash-bound. Advanced background/group/view graphs, shared or ambiguous OLE payloads, external or masked/effect-bearing pictures, and rich or irregular notes remain opaque-preserved and fail closed on mutation.

Imported objects outside these modeled profiles remain opaque and unchanged. They do not become permission for lossy rewriting.

## Build and test

```sh
dotnet test native/OpenChestnut/OpenChestnut.sln --configuration Release
npm run build:open-chestnut
npm run verify:open-chestnut-build
```

The bundled consumer runtime is written to `runtime/open-chestnut`. The npm package includes that runtime, manifest, SBOM, and license notices, but excludes this C# source tree and build output. Installed consumers therefore do not need a local .NET SDK.
