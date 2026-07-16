# Coverage

This document describes the supported 0.2 boundary. It is not a promise that every legal OOXML or PDF construct is editable.

Status meanings:

- **done**: the modeled 0.2 operation is supported for creation/import, edit, export, and second import where applicable.
- **partial**: a narrower profile is editable, or imported content is preserved but some creation/editing is intentionally rejected.
- **unsupported**: no modeled operation is provided; Office export must fail rather than flatten the object if preservation cannot be proven.

## Architecture

| Surface | Status | Boundary |
| --- | --- | --- |
| XLSX/DOCX/PPTX facade path | done | All six Office facade methods lazily load the canonical OpenChestnut C# WASM adapter. There is no codec selector or JS Office fallback. |
| PDF path | done | Independent PDF implementation; never enters the OpenChestnut OOXML wire. |
| Legacy options | done | `codec`, `allowLossy`, `preferNative`, and `relativeDateAsOf` are rejected explicitly. |
| Opaque Office preservation | done | Imported unmodeled parts are content-type/hash/source bound. Unchanged content is preserved; unsupported edits and missing source snapshots fail closed. |
| Low-level OOXML inspect/patch | done | Explicit, bounded package inspection and patching remain available and are never invoked automatically as a fallback. |
| Wire protocol | done | `open_office.artifact.v1`, protocol version 2; removed `allow_lossy` field name and number are reserved. |
| JavaScript source layering | partial | The root entry preserves the 36-symbol public API while Help, Compose, binary conversion, `FileBlob`, and inspection primitives live in dependency-leaf modules. Format models and the shared OOXML package engine remain to be extracted atomically without changing class identity or facade behavior. |

## Spreadsheets

| Capability | Status | Notes |
| --- | --- | --- |
| Cells and scalar values | done | Text, number, boolean, blank, error, formulas, cached values, and Date-to-Excel-serial conversion. |
| Formula calculation | partial | JavaScript model evaluates the bounded catalog; OpenChestnut writes/reads formulas and honors the XLSX export `recalculate` option. Unsupported Excel functions remain host-calculated or explicit model errors. |
| Static cell styles | done | Number format, font, fill, border, alignment, and protection within the modeled profile. |
| Geometry and views | done | Merges, row heights, column widths, hidden state, and frozen rows/columns. |
| Tables | done | Basic fixed-range worksheet tables and styles. |
| Images | done | Embedded PNG/JPEG worksheet images within the bounded anchor profile. |
| Charts | done | Source-free and recognized imported bar, line, and pie charts with bounded series/categories/title/legend/axis behavior. |
| Data validation | partial | Basic list, whole-number, decimal, date, time, and text-length profiles; complex extension graphs remain source-bound. |
| Conditional formatting | partial | Basic cell/value/formula/color-scale style profiles represented by the public model; complex differential-style and extension graphs remain source-bound. |
| Threaded comments | partial | One root comment per modeled thread. Replies and unsupported identity graphs are preserved only when unchanged and cannot be newly authored. |
| Pivots, QueryTables, connections, sparklines, dynamic arrays | partial | Existing imported objects can remain opaque and unchanged. Source-free authoring or semantic editing is not part of the 0.2 facade boundary. |

## Documents

| Capability | Status | Notes |
| --- | --- | --- |
| Paragraphs and runs | done | Ordered runs, bold/italic/underline, fonts, sizes, colors, and paragraph alignment/spacing/indent within the modeled profile. |
| Styles | done | Document defaults plus paragraph/character styles and `basedOn` relationships within the bounded graph. |
| Sections and pages | done | Page size, orientation, margins, section breaks, and section-level first/even behavior. |
| Headers and footers | done | Default/first/even references, section scoping, text, and PAGE/simple field content within the supported topology. |
| Lists | partial | Numbered and character-bulleted lists in the bounded numbering profile. Picture bullets and complex inherited numbering graphs remain source-bound. |
| Tables | done | Fixed geometry, rectangular cell topology, text, merges, widths, `TableGrid`, and bounded direct formatting. Arbitrary source-free custom table-style graphs are not materialized; topology-changing edits to imported complex tables fail. |
| Links and simple fields | done | External/internal links and bounded simple field instructions including PAGE. |
| Images | done | Embedded PNG/JPEG inline images with bounded geometry and alternative text. |
| Comments | partial | Classic comments and bounded anchors; modern reply/presence graphs are preserved unchanged but not newly authored. |
| Bookmarks, bibliography, tracked changes, advanced settings, complex fields, content controls, drawings | partial | Preserve imported content when source evidence remains valid; creation/editing is unsupported. Only modeled section/header behavior and `evenAndOddHeaders` are writable settings; `trackRevisions`, `updateFields`, `mirrorMargins`, and `documentProtection` fail closed. |

## Presentations

| Capability | Status | Notes |
| --- | --- | --- |
| Slides and basic shapes | done | Text boxes, rectangles/ellipses, and `roundRect` with bounded transforms. |
| Fill, line, shadow | done | Basic solid/no-fill, line color/width/dash, and bounded outer shadow. Complex theme/effect graphs remain source-bound. |
| Rich text and lists | done | Paragraphs, runs, common formatting, bullets/numbering, and bounded links. |
| Connectors | done | Straight and polyline connectors with bounded line/arrow styling. |
| Literal custom geometry | partial | Source-free and recognized imported DrawingML `custGeom` paths support bounded literal move, line, cubic Bézier, and close commands. Guides, handles, connection sites, arcs, quadratic curves, text rectangles, and per-path paint overrides remain source-bound. |
| Images | done | Embedded PNG/JPEG images with bounded placement and alternative text. |
| Tables | done | Source-free fixed rectangular tables and fixed-topology imported cell text/geometry edits. |
| Charts | done | Source-free bar, line, and pie charts with literal categories/values. Complex/combo/external-data chart graphs remain source-bound. |
| Master/Layout fidelity | partial | Imported Master/Layout graphs are preservation-only and read-only. Source-free creation and every semantic, property, or topology edit fail closed. |
| OLE, SmartArt, media, 3D, custom XML | partial | Unchanged imported graphs are opaque-preserved. Source-free creation and semantic editing fail closed. |

## PDF

| Capability | Status | Notes |
| --- | --- | --- |
| Creation and import | done | Multi-page modeled PDF creation and clean-room metadata roundtrip; optional PDF.js parser supports arbitrary-file extraction. |
| Text and tables | done | Positioned/flow text, extraction, modeled tables, spans, headers, and geometry. |
| Images and charts | done | PNG/JPEG images, vector charts, alternative text, and extraction/model records. |
| Page geometry and reading order | done | Stable IDs, explicit logical reading order, inspect, resolve, and layout JSON. |
| Accessibility | partial | Tagged structure, headings, language/title, table semantics, figures/alt text, and artifact marking. This is not a claim of full PDF/UA conformance. |
| Render and visual QA | done | SVG/layout preview plus optional Poppler/Playwright/sharp/canvas render paths and visual comparison. |
| Arbitrary PDF editing | unsupported | Arbitrary native object-stream/content-stream editing is outside the modeled PDF boundary. |

## Reference Skills

The published layout is four native plugin bundles and five Skills; `test/skill-harness` is development-only and is not a public Skill surface. Detailed evidence and gaps are in [reference Skill compatibility](reference-skills.md).

| Skill | Status | Main workflow |
| --- | --- | --- |
| Documents | partial | Native packaging plus the ordinary public-API create/import/edit/export workflow are runnable and tested through canonical OpenChestnut, semantic assertions, and real LibreOffice page QA. Python/OOXML helpers are explicit advanced package-patch or audit tools only. The broader reference tasks for tracked revisions, content controls, complex fields, and other source-bound features remain partial. |
| Spreadsheets | partial | The native plugin and reference-style core workbook example pass canonical XLSX export/import. Extended Quick API parity remains incomplete. |
| Excel live control | partial | Native routing Skill and connector declaration are present; execution depends on a host-provided connected Excel session outside the npm package. |
| Presentations | partial | The native plugin and complete 26-slide built-in template pass canonical OpenChestnut export/import, including bounded custom geometry. The broader reference API guide exceeds the current fail-closed PPTX boundary. |
| PDF | partial | Native plugin packaging and the independent project PDF pipeline pass; the reference guide still needs full public-API adapter convergence. |
