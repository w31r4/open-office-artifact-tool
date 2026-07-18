# Release

## 0.2.0 boundary

0.2.0 is a breaking convergence release:

- OpenChestnut is the only DOCX/XLSX/PPTX codec.
- PDF remains an independent fourth pipeline and never enters OpenChestnut. Required `mupdf@1.28.0` is the runtime-lazy default for arbitrary-file parse, native inspect/render, and bounded direct-original edits; specialist Python/system providers remain explicit task routes.
- The project is licensed under GNU AGPL-3.0-or-later. Normal npm installation resolves MuPDF.js as a direct dependency; there is no PDF postinstall hook or standalone dependency downloader.
- Wire protocol is version 2.
- `allow_lossy` is removed and reserved in the proto.
- `OFFICE_CODEC_IDS`, `office-codec-policy.mjs`, and the `codecs/openxml-wasm` export are removed.
- `codec`, `allowLossy`, `preferNative`, and `relativeDateAsOf` facade options are rejected.
- Old JavaScript Office parsers/writers and dedicated dead helper modules are not packaged.
- Imported advanced Office content is preserved only with validated source evidence; unsupported edits and opaque content without that evidence fail closed.

There is no compatibility window or fallback mode.

## Source and npm distributions

The repository is the authoritative source distribution. It contains OpenChestnut C# source, locked dependencies, protocol definitions, build scripts, tests, Skills, and reproducibility gates.

The npm tarball is the consumer distribution. It contains the JavaScript object models, OpenChestnut adapter, generated wire binding, public proto, bundled runtime, integrity manifest, SBOM, license notices, render/QA helpers, PDF pipeline, and five native plugin bundles containing six Skills, including the local-only Template Creator utility. It excludes C# source, test sources, build output, repository-only build scripts, and the development-only `test/skill-harness` fixtures. MuPDF.js is declared in the required npm dependency graph rather than copied into this project's own tarball, and its WASM runtime is initialized only by a PDF operation.

Installed consumers do not need `dotnet` on `PATH`.

## Required release gates

Run from a clean source checkout with the documented Node and .NET SDK versions:

```sh
npm ci
npm run proto:generate
npm run test:open-chestnut-dotnet
npm run build:open-chestnut
npm test
npm run test:pack
npm run verify:open-chestnut-build
npm run docs:api
npm run release:check
```

The release candidate is acceptable only when all of the following are true:

- the generated JavaScript wire matches protocol 2 and proto lint passes;
- C# DOCX, XLSX, PPTX, package-boundary, and failure-profile tests pass;
- default facade create/import/edit/re-export roundtrips pass for all three Office formats;
- legacy options, old subpath, missing runtime, and opaque-without-source cases fail explicitly;
- all five native plugin manifests validate, the published six-Skill topology is complete, and every workflow promoted to compatible in `docs/reference-skills.md` passes from the public package surface;
- PDF greenfield authoring plus default MuPDF.js import/inspect/render/bounded-edit, lazy-load, pre-WASM budget, exact-prefix incremental-save, signature/redaction/deletion fail-closed, Skill CLI source-protection, and specialist-provider contract tests pass independently;
- when explicitly configured, the real optional-provider test covers ReportLab creation, pdfplumber extraction, type-aware pypdf text/radio/checkbox forms and annotations, typed pypdf merge/reorder/selective watermarking, PyMuPDF rewrite/incremental/page/text/image/form/annotation edits, real redaction/scrub/residue scans, capped numerical text-fit behavior, canonical audit byte binding, and typed Poppler source/output comparison;
- Open XML SDK validation passes for generated Office fixtures;
- configured LibreOffice/Poppler/Playwright/native render gates pass where available;
- a production-only packed clean install completes all three Office roundtrips, PDF smoke, and a real packaged Template Creator invocation while `dotnet` is absent from `PATH`;
- two clean OpenChestnut builds produce the same runtime file set and hashes;
- package contents contain no legacy Office codec files, C# build output, tests, or repository-only scripts;
- package metadata, version `0.2.0`, licenses, third-party notices, SBOM, and integrity manifest agree;
- hosted Linux runs the same required non-optional gates.

The single byte-array managed entry point uses a checked-in, small JavaScript
interop registration shim rather than .NET 8's `[JSExport]` source generator.
That generator derives a wrapper name from the process-random
`String.GetHashCode()` and cannot satisfy the clean-build hash gate. The shim
uses the equivalent stable type-name hash and remains covered by the actual
bundled-WASM protocol smoke.

## Optional native validation

LibreOffice, Poppler, Playwright, and the Windows Office bridge are validation/render tools. Absence of an optional host may skip only the explicitly environment-gated native rendering branch; it must not skip codec, semantic, package, Skill, PDF, or clean-install gates.

The Office bridge does not participate in normal import/export and must never be used to hide a codec failure.

## Current local evidence

### PDF direct-original page rotation

On 2026-07-18, the default direct-original MuPDF.js route added
`PdfFile.editPdf(..., { operations: [{ type: "rotate_page", page, rotation }] })`.
The operation accepts only absolute `0`, `90`, `180`, or `270` degree clockwise
values, reports the prior normalized page value, writes only the page's
`/Rotate` entry, and verifies the value after writing. It does not transform,
reflow, or remove page content.

Native inspection now emits normalized page rotation alongside raw `MediaBox`
and `CropBox` facts. On unsigned input the operation can use byte-prefix-verified
incremental save; malformed inherited rotation, unsupported angles, signed
incremental edits, and every other unsafe save mode fail closed. The library
and published Skill CLI regressions prove raw page-box preservation, the
90-degree rendered dimension swap, absolute reset back to `0`, exact source
prefix preservation, and rejection of `45` degrees.

### PDF direct-original visible CropBox edits

On 2026-07-18, the independent PDF path added one bounded existing-file page
geometry primitive: `PdfFile.editPdf(..., { operations: [{ type:
"set_page_crop", ... }] })`. MuPDF.js now exposes raw `MediaBox` and
`CropBox` evidence for every native inspected page, accepts a raw unrotated
`[x, y, width, height]` CropBox wholly inside the inspected MediaBox, maps it
through any existing crop-origin shift, verifies the raw written box, and
renders the resulting visible size from the original bytes.

This is deliberately a visible-window operation, not content removal. The
operation records `contentRemoved: false`; off-window content remains in the
file, so it is never a redaction or sanitize route. It may use unsigned
byte-prefix-verified incremental save, while deletion/redaction and every
signed incremental edit remain rejected. Rotated pages fail closed rather than
exposing ambiguous page-box coordinates; callers must select an explicit
specialist provider for that case.

Library regressions prove raw box inspection, crop rendering dimensions,
incremental prefix preservation, restoring a pre-existing crop in the original
page coordinate system, out-of-MediaBox rejection, and rotated-page rejection.
The published PDF Skill CLI regression performs the same crop from an
operations JSON, reopens it through `PdfFile.inspectPdf`, and renders it at
72 DPI. The API reference, provider matrix, save-policy guide, and
existing-PDF workflow all state the visible-only boundary.

The complete local gate passed `npm test` (including Playwright), `npm run
docs:api`, `npm run test:pack` (444 files; 9.4 MB packed, 23.6 MB unpacked),
offline `npm run release:check -- --skip-network --allow-dirty`, OfficeBridge
`5/5`, and OpenChestnut `210/210`.

### XLSX criteria extrema formulas

On 2026-07-18, the bounded JavaScript calculation catalog added `MINIFS` and
`MAXIFS`. Each requires a value range followed by one or more
criteria-range/criterion pairs with exactly the same rectangular shape. The
implementation reuses the existing case-insensitive comparison and Excel
wildcard criteria semantics, considers only finite numeric values in matching
value cells, returns `0` when no numeric value matches, and fails with
`#VALUE!` before calculation when a criterion range has a different shape.

The public Help catalog documents both functions and regenerates the API
reference. Spreadsheet smoke tests cover multi-criterion minimum/maximum,
ignored nonnumeric matched cells, the no-match `0` result, and mismatched-range
rejection; Help tests pin the catalog and workbook counts and the numeric
formula schemas.

### XLSX bounded 2D bubble charts

On 2026-07-18, the Spreadsheet public model, protocol-2 wire, OpenChestnut C#
codec, bundled WASM runtime, generated Help/API reference, clean-install pack
probe, and native Spreadsheet Skill converged on real 2D bubble charts.
`sheet.charts.add("bubble", range)` accepts exactly `X | Y | Size` and creates
one unambiguous formula-bound series; a multi-series bubble chart must instead
use explicit `xValues/xFormula`, `values/formula`, and
`bubbleSizes/bubbleSizeFormula` fields. The three vectors must have equal point
counts, and sizes must be finite and positive.

OpenChestnut authors native `c:bubbleChart`, `c:xVal`, `c:yVal`,
`c:bubbleSize`, and two numeric value axes. It uses an explicit 2D,
100%-scale, area-sizing profile. Imported `bubble3D`, negative-bubble,
custom-scale, or non-area-size graphs are detected as source-bound/read-only;
unchanged export retains them exactly and mutation fails closed. Bubble SVG
preview maps the positive size vector proportionally to visible circles rather
than pretending a bubble plot is a scatter plot.

The runnable Skill workflow creates, inspects, verifies, exports, imports,
edits, re-exports, and re-imports one formula-backed bubble chart. Its real
LibreOffice-to-PDF plus Poppler QA passes in one page. The production-only
tarball smoke also creates/imports the bubble chart while `dotnet` is absent
from `PATH`; C# tests validate Office 2021 XML, fixed-topology edits, and
exact preservation/rejection of a noncanonical `bubbleScale` profile.

### XLSX marker-only scatter charts

On 2026-07-18, the Spreadsheet public model, protocol-2 wire, OpenChestnut C#
codec, bundled WASM runtime, generated Help/API reference, and native
Spreadsheet plugin converged on real numeric-X/Y marker-only scatter charts.
Range-backed creation treats the first numeric column as each series' X source
and the remaining columns as Y series; explicit series carry `xValues/xFormula`
and `values/formula`. OpenChestnut authors native `c:scatterChart`, `c:xVal`,
`c:yVal`, and two value axes instead of substituting a category line chart.

Marker-only is a cross-host contract, not only a chart-style label. Every
authored or edited series now carries an explicit DrawingML no-fill outline so
LibreOffice and Microsoft Office do not synthesize connecting lines.
Series-level `line`/`stroke` is rejected; `marker.line` remains the supported
marker-border control. Imported scatter line/smooth variants and marker-style
graphs with an explicit series line remain source-bound/read-only and are exact
preserved until a broader semantic profile exists.

The runnable Skill example creates two formula-bound series, imports, edits,
re-exports, re-imports, inspects, verifies, and renders SVG. Its real local
LibreOffice-to-PDF plus Poppler run produced one complete page with separate
circle/diamond points and no connecting lines; package, semantic, SVG, native
page-count, and raster QA all passed. The package test also exercises scatter
create/import from a production-only clean install with no `dotnet` on `PATH`.

Final candidate `77706e52e72a22a6ff76bc11ca7dda29e9434b01` passed hosted
[CI run 29600831755](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29600831755)
in 4m43s. The run passed deterministic protocol/runtime verification, installed
Chromium/LibreOffice/Poppler, the complete `npm test`, regenerated API-doc diff,
offline release check, clean-install pack gate, OfficeBridge, and OpenChestnut.
Local evidence for the same candidate includes OpenChestnut `186/186`,
OfficeBridge `5/5`, a reproducible 39-file build audit, a 38-file bundled runtime
of 14,329,020 bytes, and a 435-file npm tarball of 9,301,407 compressed bytes /
23,199,173 unpacked bytes. Offline release metadata reports `publishReady: true`;
real npm authentication/publication remains an external release action.

### XLSX What-If data tables

On 2026-07-18, the Spreadsheet public model, protocol-2 wire, OpenChestnut C#
codec, bundled WASM runtime, generated Help/API reference, and native
Spreadsheet plugin converged on canonical Excel What-If data tables.
`sheet.dataTables.add(range, { rowInput, columnInput })` authors row-oriented
and column-oriented one-variable tables plus two-variable tables as native
`<f t="dataTable">` formulas. Import exposes defensive result-range, anchor,
input, orientation, and display-formula evidence without projecting the native
anchor as an ordinary editable cell formula.

The bounded profile is loss-aware: source-free authoring rejects missing
top-left formulas, cross-sheet or out-of-bounds input cells, undersized grids,
and overlapping result ranges. Recognized imports retain fixed count, order,
result range, inputs, and orientation; ordinary surrounding-cell edits and a
second export/import work, while topology or semantic data-table mutation fails
closed. Noncanonical attributes remain source-bound and unchanged. Excel,
LibreOffice, or another compatible host calculates the result values; the
JavaScript evaluator does not pretend to simulate `TABLE`.

The complete candidate at commit
`94f1d2234c93078b0782a52fd95c115be9df86f8` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29594786964](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29594786964)
in 5m07s. The run covered deterministic protocol/runtime verification,
Chromium/LibreOffice/Poppler setup, the full npm suite including the runnable
What-If Skill workflow and fail-closed topology corpus, generated API-doc
cleanliness, offline release metadata, clean-install packing, OfficeBridge
`5/5`, and OpenChestnut `185/185`.

### Neutral template library and Template Creator sync

On 2026-07-18, the observable Skill delta from the remotely reachable reference
revision `207ce094a55d82a37efdca42a1c5e9656f696962`
(`origin/feat/sync-grid-layout-library-template-creator`, based on the
`office-artifact-tool@2.8.24` runtime-sync revision) was adapted into the public
package. The Presentation asset/support trees now use the neutral
`grid-layout-library` identity throughout. Its model-facing registry, 26 exact
plain-JavaScript Compose builders, 27 preview PNGs, content/design tokens, and
reconstruction runner remain unflattened and pass a complete 26-slide
OpenChestnut export. The submodule now pins that exact reachable revision; no
unreachable pointer is recorded.

The new fifth plugin contains the sixth published Skill, Template Creator. It
creates numbered local `artifact-template-*` Skills from one DOCX/PPTX/XLSX
reference and a structurally validated PNG, or updates exactly one named
same-kind template. The retained Office reference and preview are byte-exact.
The implementation performs no network fetch, writes only beneath
`${OFFICE_ARTIFACT_HOME:-~/.office-artifact-tool}/skills`, enforces 512 MiB
reference and 64 MiB preview budgets, rejects symlink-bearing update trees,
serializes writers, recovers interrupted replacements, preserves extra
template-owned files, rolls back failed placement, and proves there is no
stage/backup/lock residue. All five native plugin manifests pass the official
validator, while the reference-compatible versioned Template Creator manifest
is retained alongside the native manifest.

The complete local gate passed `npm test`, `npm run docs:api`,
`npm run proto:check`, `npm run test:pack`, serial
`npm run verify:open-chestnut-build`, OpenChestnut `185/185`, and OfficeBridge
`5/5`. Two clean WASM builds produced the same 39 audited files and the same
manifest-bound 38-file, 14,324,412-byte runtime. The clean-install tarball
contains 434 files, is 9,297,469 bytes compressed and 23,178,713 bytes
unpacked. The pack gate now also invokes Template Creator from that offline,
production-only installation, validates the generated spreadsheet-template
sidecar, and proves the retained XLSX and PNG bytes exactly match its inputs.
The final candidate at commit
`4f3dbbb4f95081e1c86747e88c10b4b16722ad0d` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29596412072](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29596412072)
in 4m40s, covering npm installation, deterministic OpenChestnut verification,
Chromium/native-tool setup, the complete npm/Skill/security suite, generated
API-doc cleanliness, offline release metadata, clean-install packing,
OfficeBridge `5/5`, and OpenChestnut `185/185`.

### XLSX standard-area and fixed-doughnut chart families

On 2026-07-17, the Spreadsheet public model, protocol-2 wire, OpenChestnut C#
codec, bundled WASM runtime, generated Help/API reference, and native
Spreadsheet plugin converged on two additional canonical worksheet-chart
families. Source-free and recognized imported standard-area charts now carry
the existing primary category/value axes; 50%-hole doughnut charts author and
import as circular plots without axes. Both families support bounded
categories, series, titles, legends, source formulas/caches, semantic edits,
second import, and type-aware SVG QA. The worksheet chart preview moved from
the large Spreadsheet domain leaf into its own focused module while preserving
the public `WorksheetChart.toSvg()` contract.

The slice remains deliberately loss-aware. Stacked and percent-stacked area,
non-50% doughnut geometry, exploded points, scatter/bubble/radar/stock/modern
families, and other unmodeled plot options are never substituted. Recognized
advanced area/doughnut graphs import as source-bound and read-only, preserve
their native chart XML unchanged, and reject semantic edits with
`unsupported_spreadsheet_chart_edit`; unsupported source-free types and
illegal axis/line-option combinations fail before crossing the wire. The
regressions also closed two existing chart-patch defects: valid imported
pie/doughnut topology is now editable, and editing a non-line chart cannot
inject line grouping markup.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, isolated `npm run verify:open-chestnut-build`,
OpenChestnut `184/184`, and OfficeBridge `5/5`. Two isolated clean WASM builds
produced the same 39 audited files and the same manifest-bound 38-file,
14,318,780-byte runtime. The clean-install tarball contains 423 files, is
9,276,398 bytes compressed and 23,122,542 bytes unpacked. No npm publish, tag,
or GitHub release operation was attempted.

The completed candidate at commit
`e61bb1ff143f27a9d4afc1ecdd6490634c3dfa4a` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29590185608](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29590185608)
on 2026-07-17. The run completed with conclusion `success` in 4m31s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including five-family chart author/import/edit/preservation
and native Spreadsheet Skill regressions, generated API-doc cleanliness,
offline release metadata, the 423-file clean-install tarball, OfficeBridge
`5/5`, and OpenChestnut `184/184`.

### DOCX transactional SEQ/REF cached-result materialization

On 2026-07-17, the public Documents model and native Documents plugin closed
the bounded automatic cached-result gap for canonical inline `SEQ` and `REF`
fields. `document.materializeFields()` computes independent, case-insensitive
SEQ counters in document order, resolves REF targets from the shared native
bookmark namespace, and updates only cached display text. The operation
supports dry-run evidence and defaults to strict transactional failure before
mutation when a bookmark target is missing, duplicated, or dangling. An
explicit non-strict mode may update resolvable fields while reporting missing
targets.

The primitive does not remove or flatten native field topology. PAGEREF is
reported but never fabricated; explicitly requesting it fails closed because
trustworthy page numbers require a real pagination host. Imported field
positions, instructions, and bookmark identity remain source-bound. The
shipped Documents fixture now materializes one caption SEQ and its REF through
the public model before two OpenChestnut round trips, while retaining an
explicit PAGEREF cache for native-host refresh. JavaScript regressions cover
dry-run immutability, strict rollback, opt-in partial materialization,
case-insensitive targets, multiple counters, and the PAGEREF boundary.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, isolated `npm run verify:open-chestnut-build`,
OpenChestnut `183/183`, OfficeBridge `5/5`, and the Documents Skill validator.
Two isolated clean WASM builds produced the same 39 audited files and the same
manifest-bound 38-file, 14,317,244-byte runtime. The clean-install tarball
contains 422 files, is 9,274,329 bytes compressed and 23,114,497 bytes
unpacked. No npm publish, tag, or GitHub release operation was attempted.

The completed candidate at commit
`fbe86f197850f2bc297c6a95c5b5d5d19e442ab2` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29588303966](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29588303966)
on 2026-07-17. The run completed with conclusion `success` in 4m29s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the native field-materialization workflow and
transactional fail-closed regressions, generated API-doc cleanliness, offline
release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and
OpenChestnut `183/183`.

### DOCX caption-number inline bookmark workflow

On 2026-07-17, the public Documents run model, protocol-2 wire, OpenChestnut
C# codec, bundled WASM runtime, Help catalog, and native Documents plugin
closed the bounded caption-target gap. A source-free canonical `SEQ` run may
now set `bookmarkName` through `paragraph.addField(...)`; OpenChestnut writes
one paired `w:bookmarkStart`/`w:bookmarkEnd` around only that field's cached
result run. Canonical `REF` and `PAGEREF` runs can therefore target the native
caption number without an OOXML patch helper.

Whole-block and inline-field bookmarks share one case-insensitive name and
unsigned native-ID space. The two bounded profiles cannot nest in the same
paragraph. Import collapses the seven-child bookmarked field graph back into
one logical run, preserves name/native identity, permits cached-result and
ordinary-text edits, and rejects any field position, instruction, bookmark
name, or native-ID change. Unsupported placement and duplicate identity fail
closed; automatic field calculation/materialization remains an explicit host
or package-helper workflow rather than a codec claim.

The shipped Documents fixture authors `SEQ Figure \\* ARABIC` with `fig1`, then
targets it with both `REF fig1 \\h` and `PAGEREF fig1 \\h`. It completes two
OpenChestnut round trips, edits all three cached results, asserts that the
native bookmark encloses only the SEQ result run, and passes real
LibreOffice/Poppler page QA. C# and JavaScript regressions cover semantic
roundtrip, source-bound identity, invalid non-SEQ bookmark use, case-insensitive
duplicate names, and bounded-profile nesting rejection.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, isolated `npm run verify:open-chestnut-build`,
OpenChestnut `183/183`, OfficeBridge `5/5`, and the Documents Skill validator.
Two isolated clean WASM builds produced the same 39 audited files and the same
manifest-bound 38-file, 14,317,244-byte runtime. The clean-install tarball
contains 422 files, is 9,273,020 bytes compressed and 23,105,625 bytes
unpacked. No npm publish, tag, or GitHub release operation was attempted.

The completed candidate at commit
`206712aef6691490318e42bfa26553d27b4d1fc9` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29586820145](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29586820145)
on 2026-07-17. The run completed with conclusion `success` in 4m36s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the native caption/cross-reference workflow and
fail-closed identity regressions, generated API-doc cleanliness, offline
release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and
OpenChestnut `183/183`.

### DOCX canonical inline SEQ/REF/PAGEREF field runs

On 2026-07-17, the Documents paragraph/run model, versioned protobuf wire,
OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native Documents
plugin converged on one bounded inline-field profile. A paragraph can mix
ordinary runs with logical `SEQ <label> \\* ARABIC`, `REF <bookmark> \\h`, and
`PAGEREF <bookmark> \\h` fields through `paragraph.addField(...)`.
OpenChestnut materializes each logical field as the canonical five-run native
begin/instruction/separate/cached-result/end graph and imports that exact graph
back into one `run.inlineField` object.

Source-free cached results and ordinary text survive export, import, edits,
second export, and second import. On imported paragraphs, field positions and
instructions remain source-bound while cached display text remains editable;
topology or target changes fail closed. Non-canonical switches, partial or rich
field graphs and automatic SEQ/REF materialization remained opaque or explicit
advanced package workflows at that milestone; the later bounded caption-number
bookmark slice is recorded above.
The shipped native fixture covers all three commands, native XML evidence,
semantic inspection, cache edits, and real LibreOffice/Poppler page QA.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, isolated `npm run verify:open-chestnut-build`,
OpenChestnut `183/183`, and OfficeBridge `5/5`. Two isolated clean WASM builds
produced the same 39 audited files and the same manifest-bound 38-file,
14,309,564-byte runtime. The clean-install tarball contains 422 files, is
9,270,563 bytes compressed and 23,092,374 bytes unpacked. No npm publish, tag,
or GitHub release operation was attempted.

The completed inline-field candidate at commit
`05713031054388f92c0586e7f6547a407544c696` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29584754018](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29584754018)
on 2026-07-17. The run completed with conclusion `success` in 4m58s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the native inline-field Documents fixture and
source-bound topology regression, generated API-doc cleanliness, offline
release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and
OpenChestnut `183/183`.

### DOCX canonical complex TOC and field-refresh workflow

On 2026-07-17, the Documents public model, versioned protobuf wire,
OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native Documents
plugin converged on one bounded complex-field profile. Source-free documents may
author a canonical one-paragraph TOC field with configurable ascending heading
levels and the `\\h`, `\\z`, and `\\u` switches. The public
`document.addTableOfContents(...)` primitive emits native
`w:fldChar`/`w:instrText` topology and enables the `w:updateFields` open-time
refresh hint by default. The hint is modeled explicitly and never asserts that
the cached TOC result is current.

Canonical unrefreshed one-paragraph TOCs survive export, import,
fixed-topology instruction/display edits, second export, and second import.
Refreshed cross-paragraph TOCs, including otherwise plain-looking paragraphs
inside the field span, remain opaque/source-bound and read-only as one graph.
Arbitrary complex REF/PAGEREF/SEQ/external-content fields remain outside the
semantic profile. The shipped Documents fixture proves the public workflow,
native field/settings XML, semantic inspect evidence, LibreOffice rendering,
and fail-closed irregular-graph preservation.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, isolated `npm run verify:open-chestnut-build`,
OpenChestnut `182/182`, and OfficeBridge `5/5`. Two isolated clean WASM builds
produced the same 39 audited files and the same manifest-bound 38-file,
14,303,420-byte runtime. The clean-install tarball contains 422 files, is
9,268,734 bytes compressed and 23,076,840 bytes unpacked. No npm publish, tag,
or GitHub release operation was attempted.

The completed TOC/field-refresh candidate at commit
`f7c83de4fbb83ddade81542cb95df86d01be4939` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29582981065](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29582981065)
on 2026-07-17. The run completed with conclusion `success` in 4m28s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the native TOC Documents workflow and
cross-paragraph fail-closed regression, generated API-doc cleanliness, offline
release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and
OpenChestnut `182/182`.

### DOCX canonical bibliography and citation public workflow

On 2026-07-17, the Documents public model, versioned protobuf wire,
OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native Documents
plugin converged on one reversible bibliography/citation profile. Source-free
documents may author one canonical bibliography Custom XML catalog plus
whole-paragraph `w:fldSimple` `CITATION <tag>` fields. Ordinary Author name
lists or one corporate author, the bounded scalar source catalog, bibliography
settings, and multiple sources survive export, import, fixed-topology content
edits, second export, and second import through the Microsoft Open XML SDK.

Imported source order, GUIDs, tags, citation tags, and field topology remain
source-bound. Other contributor roles, complex field switches/results,
multiple or irregular bibliography parts, bibliography output fields, and
extension graphs stay opaque/read-only and fail closed on semantic mutation.
The package inspector recognizes SDK-authored package-root bibliography
relationship targets without weakening OPC path validation. The shipped
Documents example and fixture author a source and citation, edit the imported
title/author/display result, resolve the final semantic objects, assert native
bibliography and field markup, verify the model, and complete two OpenChestnut
round trips.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, serial `npm run verify:open-chestnut-build`, OpenChestnut
`180/180`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39
audited files and the same manifest-bound 38-file, 14,297,788-byte runtime. The
clean-install tarball contains 422 files, is 9,265,073 bytes compressed and
23,061,507 bytes unpacked. No npm publish, tag, or GitHub release operation was
attempted.

The completed public-workflow candidate at commit
`cec9628f8429934e48820bdbae064d5a2ff32fd5` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29580505711](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29580505711)
on 2026-07-17. The run completed with conclusion `success` in 4m11s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the canonical bibliography/citation Documents
workflow and fail-closed irregular-graph regressions, generated API-doc
cleanliness, offline release metadata, the 422-file clean-install tarball,
OfficeBridge `5/5`, and OpenChestnut `180/180`.

### DOCX inline plain-text content-control public workflow

On 2026-07-17, the already integrated inline plain-text content-control codec
was completed as a truthful public Documents workflow. The generated Help/API
surface now documents `paragraph.addTextContentControl(...)`,
`document.contentControls`, and transactional
`document.fillContentControls(...)`; the coverage and architecture records no
longer describe every content control as unsupported. The native Documents
Skill routes source-free and recognized imported body-inline plain-text SDTs
through the public JavaScript model and OpenChestnut, while retaining the
Python package helper only for explicit existing-template wrapping,
headers/footers, and controls outside the bounded model.

The shipped `openchestnut-end-to-end.mjs` example now authors an `OWNER`
control, exports canonical `w:sdt` markup, imports it, fills it by tag, exports
again, and proves the final tag/alias/text through semantic import, inspect,
native XML, and verification assertions. Unknown tags fail before mutation by
default; duplicate tags fill every match. Imported text/tag/alias edits are
permitted only with fixed run topology and native identity. Rich, block, cell,
nested, data-bound, dropdown, date, checkbox, placeholder-document, locked,
header/footer, and extension-bearing controls remain opaque/source-bound and
fail closed rather than being flattened.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, offline `npm run release:check -- --skip-network
--allow-dirty`, serial `npm run verify:open-chestnut-build`, OpenChestnut
`179/179`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39
audited files and the same manifest-bound 38-file, 14,250,684-byte runtime. The
clean-install tarball contains 422 files, is 9,245,399 bytes compressed and
23,000,796 bytes unpacked. No npm publish, tag, or GitHub release operation was
attempted.

The completed public-workflow candidate at commit
`6231544c88dd5eee461c2b3e4950425807c8add9` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29575996251](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29575996251)
on 2026-07-17. The run completed with conclusion `success` in 4m30s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the published Documents Skill example and native
content-control assertions, generated API-doc cleanliness, offline release
metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and
OpenChestnut `179/179`.

### XLSX direct threaded-comment replies

On 2026-07-17, the Spreadsheet public model, versioned protobuf wire, OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native Spreadsheet Skill converged on the reference workflow's `thread.addReply()` primitive. One root may now carry multiple direct replies with independent native comment/person GUIDs, display/user/provider identity, ISO date, done state, semantic inspect evidence, and fixed-cell parent binding. Source-free replies author native Office 2019 `parentId`; canonical imports expose the replies for text/state edits and a second export/import; unchanged imported threaded-comment parts remain byte-preserved.

The bounded topology is deliberately loss-aware: a reply may point only to its thread root and must target the same cell. Reply-of-reply or branched graphs, orphan/self/cross-cell parents, mentions, invalid GUID/person graphs, and other extensions remain opaque/source-bound; attempts to replace them fail closed. The shipped Spreadsheet fixture now exercises two authors, one direct reply, native XML evidence, semantic import, and the public Skill workflow instead of documenting a model-only method that the codec rejects.

The complete local gate passed `npm test` including Playwright, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, serial `npm run verify:open-chestnut-build`, OpenChestnut `178/178`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,238,396-byte runtime. The clean-install tarball contains 422 files, is 9,234,762 bytes compressed and 22,965,796 bytes unpacked. No npm publish or tag operation was attempted.

The complete direct-reply candidate at commit `f69b5b50ffb6ace73d0e6de53db9c3bb30f7571b` passed the hosted Linux `ci` workflow in [GitHub Actions run 29573634244](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29573634244) on 2026-07-17. The run completed with conclusion `success` in 5m20s and covered deterministic protocol/runtime verification, Chromium/LibreOffice/Poppler tools, the full npm suite including the native direct-reply Skill fixture and topology gates, generated API-doc cleanliness, offline release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `178/178`.

### Reference 2.8.24 public API completion

On 2026-07-17, a public-surface comparison between the previously pinned 2.8.22 package and the synchronized 2.8.24 reference found five additional named exports: `setOfficeFontDesignMetrics`, `registerScopedOfficeFontDesignMetrics`, `resolveOfficeFontDesignMetrics`, `clearOfficeFontDesignMetrics`, and `skiaPaintBaselineCompensationPx`. The clean-room implementation independently provides replacement and scoped metric registries, defensive normalized results, primary-family and exact-style selection, nearest numeric weight selection, idempotent scope disposal, invalid-record filtering, and deterministic finite baseline compensation. It also exposes fresh, sorted, case-insensitively deduplicated `fontFamilies` inventories on Document, Workbook, and Presentation models.

The same observable audit found the read-only `Shape.useBackgroundFill` getter. OpenChestnut now imports presence-aware native `p:sp/@useBgFill`, projects it through protobuf and the JavaScript model, uses it for preview paint, and preserves an unchanged source-bound slide XML byte-for-byte. Source-free authoring and semantic mutation fail closed instead of manufacturing or flattening native state. The native Documents, Spreadsheets, and Presentations Skill guides document the new inventories and imported-shape boundary; generated API documentation contains all nine new Help records.

The complete local gate passed `npm test` including Playwright, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, serial `npm run verify:open-chestnut-build`, OpenChestnut `178/178`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,232,252-byte runtime. The clean-install tarball contains 422 files, is 9,229,500 bytes compressed and 22,954,466 bytes unpacked. `npm run release:check` passed every code, package, license, documentation, JavaScript, and .NET gate; before this evidence commit it reported only the intentionally dirty generated documentation plus the external npm-authentication blocker. No publish or tag operation was attempted.

The completed reference 2.8.24 API-sync candidate at commit `a81d71823a3366805d9cb3a44c322bced386f2c5` passed the hosted Linux `ci` workflow in [GitHub Actions run 29571987958](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29571987958) on 2026-07-17. The run completed with conclusion `success` in 4m07s and covered protocol/runtime verification, Chromium/LibreOffice/Poppler tool checks, the full npm suite, generated API-doc cleanliness, offline release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `178/178`.

### Reference 2.8.24 Skill and Presentation-view convergence

On 2026-07-17, the pinned `reference/office-artifact-tool` submodule advanced from `2.8.22` to `2.8.24` at `2d0e249ea6b62f55cca22a343b832a38e8f7537c`. The published Spreadsheet Skill now uses the reference-native `artifact_tool_docs/API_QUICK_START.md` and `features/charts.md` paths, with package and workflow regressions proving the old locations are absent. The Presentation Skill gained the reference `presentation.view` workflow without importing the private runtime: local gridline/guide visibility matches the observable API, while a new independent protobuf/Open XML SDK slice imports grid spacing, snap flags, and bounded horizontal/vertical guides from `ppt/viewProps.xml`.

Imported guide definitions are returned as defensive copies by `presentation.view.toProto()` and exposed as frozen Master/Layout `slideGuides` projections. Unchanged source-bound export preserves the native view-properties part byte-for-byte; semantic mutation, source-binding mismatch, source-free wire authoring, invalid guide values, and more than 1,024 modeled guides fail closed. The combined release also completed the pending DOCX note edge corpus and public Documents workflow: rich multi-paragraph note bodies remain opaque/source-bound and byte-preserved when unchanged instead of being flattened into the bounded plain-text note model.

The complete local gate passed `npm test`, `npm run docs:api`, clean-HEAD `npm run proto:check`, `npm run test:pack`, serial `npm run verify:open-chestnut-build`, OpenChestnut `177/177`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,230,716-byte runtime. The clean-install tarball contains 421 files, is 9,220,694 bytes compressed and 22,935,038 bytes unpacked. The generic Skill quick validator accepts Documents but rejects the reference package's intentionally preserved `name: Spreadsheets` and `name: Presentations` capitalization; the project-native four-plugin validator and complete reference workflow gate pass those published discovery contracts.

### DOCX whole-block bookmark vertical slice

On 2026-07-17, the Documents public model, versioned protobuf wire, OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native Documents plugin converged on one reversible bookmark profile. A source-free bookmark may wrap exactly one paragraph-like block and supplies a native target for internal `w:hyperlink` anchors. Export writes paired `w:bookmarkStart`/`w:bookmarkEnd`; import exposes name, target, native identity, source position, and semantic binding. Recognized imported whole-block bookmarks remain fixed-topology and read-only, while cross-block, nested, crossing, table-cell, and otherwise irregular graphs stay opaque-preserved and fail closed on mutation.

The shipped Documents example now authors a `DecisionSection` bookmark plus internal jump link, completes two OpenChestnut round trips, resolves the imported target, verifies the final model, and asserts the native bookmark and hyperlink markup in `word/document.xml`. The broader Documents smoke continues to require semantic verification and native LibreOffice/Poppler page rendering. C# and JavaScript regressions cover source-free authoring, internal-link planning before bookmark insertion, unchanged source-preserving export, imported rename rejection, topology-change rejection, and cross-block authoring rejection.

The complete local gate passed `npm test`, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, serial `npm run verify:open-chestnut-build`, OpenChestnut `174/174`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,184,636-byte runtime. The clean-install tarball contains 421 files, is 9,202,409 bytes compressed and 22,856,201 bytes unpacked. The unpacked-size regression budget moved narrowly from 22,850,000 to 22,900,000 bytes to account for the measured codec/runtime addition while retaining less than 44 KiB of headroom.

### Composition root and Spreadsheet domain layering

On 2026-07-17, the remaining stateful Spreadsheet domain moved from the public root entry into `src/spreadsheet/index.mjs`: workbook/worksheet/range ownership, tables/pivots/charts/images/sparklines, formulas and dependency graphs, conditional formatting, inspect/resolve/verify/help, layout/SVG rendering, delimited-file support, metadata restoration, and the `SpreadsheetFile` facade now share one cohesive owner. The root retains exact re-exports rather than wrappers; all four public Spreadsheet bindings are strict-identical between the root and leaf entry. The OpenChestnut Spreadsheet adapter imports the leaf directly, and a source-level regression prevents a back-edge to `src/index.mjs`.

This completed the public-entry decomposition: `src/index.mjs` is now a 42-line composition root, or 0.1% of the 28,109 JavaScript source lines, while preserving all 36 public exports and the cross-format verify/render/visual-QA/help surface. The complete local gate passed `npm test`, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, `npm run verify:open-chestnut-build`, OpenChestnut `173/173`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,166,204-byte runtime. The clean-install tarball contains 421 files, is 9,196,227 bytes compressed and 22,828,311 bytes unpacked.

### Presentation domain layering

On 2026-07-17, the complete JavaScript Presentation domain moved from the public root entry into `src/presentation/index.mjs`: model collections, themes/masters/layouts/placeholders, shapes/connectors/groups/images/tables/charts/native objects, inspect/resolve/verify/help, layout/SVG rendering, and the `PresentationFile` facade now share one cohesive owner. The root retains exact re-exports rather than wrappers; all eight public Presentation bindings are strict-identical between the root and leaf entry. The OpenChestnut Presentation adapter imports the leaf directly, and a source-level regression prevents a back-edge to `src/index.mjs`.

The extraction reduced `src/index.mjs` from 5,694 to 4,305 lines, or 15.3% of the 28,107 JavaScript source lines, while preserving all 36 public root exports. The complete local gate passed `npm test`, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, `npm run verify:open-chestnut-build`, OpenChestnut `173/173`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,166,204-byte runtime. The clean-install tarball contains 420 files, is 9,196,256 bytes compressed and 22,828,153 bytes unpacked.

### DOCX whole-paragraph tracked-change vertical slice

On 2026-07-17, the Documents public model, versioned protobuf wire, OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native plugin workflow converged on one bounded tracked-change profile. `document.addInsertion(...)` and `document.addDeletion(...)` author native whole-paragraph `w:ins`/`w:del` markup containing one text run; import exposes type, text, author, and optional timestamp; source-preserving export permits fixed-topology text/author/date edits while retaining native revision identity and unmodeled formatting. In-paragraph replacements, mixed or nested revisions, moves, property changes, and automatic future-change tracking remain explicit advanced workflows.

The edge-case corpus proves that a multi-run revision is still semantically visible but marked read-only, carries no editable-topology residual hash, round-trips with byte-identical source-package output when unchanged, and rejects semantic mutation with `unsupported_document_edit`. This closes the irregular-topology exception path without weakening fail-closed preservation.

The complete local gate passed `npm test`, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, `npm run verify:open-chestnut-build`, OpenChestnut `173/173`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,166,204-byte runtime. The clean-install tarball contains 419 files, is 9,196,284 bytes compressed and 22,827,449 bytes unpacked. The shipped Documents example completed create, export, import, tracked-change edit, second export/import, semantic/XML assertions, and LibreOffice/Poppler rendering; both rendered pages were reviewed, with the insertion and deletion visibly represented and no clipping or overlap.

### AGPL and default MuPDF.js vertical slice

On 2026-07-17, the focused PDF vertical slice passed `node test/pdf.mjs`, `node test/pdf-provider-skill.mjs`, `node test/reference-skills.mjs`, `node test/release-check.mjs`, and the package-contents gate. The checks cover root-import laziness, first-use MuPDF initialization, arbitrary-PDF import/inspect, native PNG/JPEG render, bounded direct-original editing, input/render/image/object limits, exact-prefix incremental saves, rewrite redaction, signed/redaction/deletion incremental fail-closed behavior, real link and raster extraction, CLI atomic output, nested output creation, and direct/symlink source-overwrite rejection.

This focused record intentionally does not assign the combined worktree's final package file count, tarball size, complete Office/PPTX gate, or hosted CI result to the MuPDF-only change. Those integration measurements belong to the subsequent combined release-evidence update.

On 2026-07-17, the native reference-plugin/OpenChestnut compatibility worktree passed the complete local gate on macOS arm64 with Node 26.5.0, npm 11.17.0, and .NET SDK 8.0.128:

- `npm test` passed the OpenChestnut protocol/facade tests, explicit OOXML inspect/patch tests, four native plugin bundles/five published Skills, the Spreadsheet Range/R1C1/direct-series/standard-sparkline compatibility suites, the runnable Documents create/import/edit/export vertical slice, the 26-slide reference Presentation vertical slice plus recursive-group and embedded-XLSX/OLE fixtures, PDF greenfield and provider-contract suites, render/visual QA, Playwright, examples, release metadata, package contents, and Help catalog.
- OpenChestnut passed `170/170` C# tests, including standard Office 2010 sparkline coverage, literal DrawingML custom-geometry coverage, direct PPTX slide-background author/import/hash-bound add-edit-remove/advanced-source fail-closed coverage, embedded-picture signed `a:srcRect` author/import/add-edit-remove/irregular-source fail-closed coverage, plain-text speaker-notes author/import/hash-bound edit/rich-source fail-closed coverage, recursive native `p:grpSp` author/import/fixed-topology mixed-descendant edit/complex-group opaque coverage, and source/hash/content-type/relationship-bound embedded XLSX payload replacement with Open XML SDK plus Office 2021 validation and unrelated-part preservation; the optional OfficeBridge passed `5/5` protocol tests.
- Buf lint passed, protobuf generation was byte-idempotent, and `npm run docs:api` regenerated the public API reference.
- `npm run test:pack` passed the no-local-dotnet clean-install probe for DOCX/XLSX/PPTX and the independent PDF path.
- The final combined dry-run npm tarball contains 417 files, is 9,190,240 bytes compressed and 22,803,636 bytes unpacked. It includes the rich PDF Skill tasks/references/examples, required lazy MuPDF.js route, six-page tagged-accessibility report workflow, explicit Python-runtime selector, path-safe attachment quarantine, type-aware AcroForm and merge/reorder/selective-watermark adapters, active-content inert scanner, canonical audit validator/schema, typed Poppler source/output comparator, the dependency-leaf cross-format raster registration/diff/visual-QA engine, the shipped Spreadsheet Range and sparkline workflows, the PPTX direct-background, speaker-notes, signed image-crop, recursive native-group, and embedded-XLSX/OLE workflow references, and the dependency-leaf OOXML package, Range, chart-source, sparkline, and Presentation group/image-fit/crop/native-object modules while excluding Python bytecode/cache files, development harnesses, private reference material, Agent PromptBench oracles, and removed Office codec paths.
- `npm run verify:open-chestnut-build` compared 39 audited files; both clean builds produced the same manifest-bound 38-file, 14,153,404-byte runtime payload. The build entry point clears the Release incremental graph before restore/publish so a preceding `dotnet test` cannot make the first WASM publish differ from the second.
- Render-backed gates ran with LibreOfficeDev 26.8.0.0.alpha0, Poppler 26.05.0, and the installed Playwright Chromium runtime.
- The explicit real-provider gate ran with ReportLab 4.4.9, pdfplumber 0.11.9, pypdf 6.10.0, and separately installed PyMuPDF 1.27.2.3 under the approved AGPL route. It proved byte-identical incremental prefixes, non-prefix rewrites, pypdf text/radio/checkbox value and appearance handling, complete-source page selection/reorder/selective watermarking with preserved navigation, pypdf and PyMuPDF annotation operations, bounded text/image/page edits, full redaction/scrub, single-revision/no-residue output, and deliberate reflow rejection. The typed Poppler comparator maps each merged output page back to its immutable source, requires exact pixels on unstamped pages, requires bounded change on stamped pages, and reports blank-state and dark-ratio evidence. The scrub-only active-content fixture proves removal of root/additional JavaScript, Launch/SubmitForm actions, attachments, invisible text, comments, populated form values, personal metadata, and the null active-content dictionary names that PyMuPDF can leave after logical deletion; unfamiliar object serialization and invisible text overlapping visible text fail closed. An image-bearing strict residue scan failed closed because Tesseract is not installed; qpdf, pyHanko, veraPDF, pikepdf, and OCRmyPDF were not executed and remain external/planned as documented.
- The shipped Documents example and an independent Skill forward test each completed two OpenChestnut round trips, semantic assertions, and a one-page LibreOffice render with every page inspected. The audit narrowed custom source-free table styles to `TableGrid` plus direct formatting, documented non-persistent model locators across imports, aligned header/footer distance to the codec's 720-twip default, and routed ordinary classic comments through `DocumentModel.addComment`.
- The shipped Spreadsheet Range example completed R1C1/block-write/formula-evidence/navigation/format/chart/verify/render and two OpenChestnut round trips. A separate forward test authored and visually reviewed a three-sheet operating forecast with formula-driven financials, zero spreadsheet errors, PASS model checks, and a formula-bound line chart. Its findings led to automatic native formulas for `containsText`, text-preserving direct references, and live internal-range cache resolution for formula-only chart inspect/render/export; imported chart persistence snapshots remain deliberately separate and fail closed.
- The shipped Spreadsheet sparkline example and development fixture authored standard Office 2010 `x14:sparklineGroups` line, column, and stacked profiles, exercised vertical row and horizontal column mappings, imported and edited recognized groups with fixed topology, preserved unsupported non-contiguous native groups unchanged, and rejected lossy topology changes. LibreOfficeDev opened the fixture and Poppler rendered all three types across two pages; the JavaScript SVG preview rendered each target cell independently.
- The repository-only Agent PromptBench defines 31 black-box cases (19 PDF and 12 Office; 11 ready and 20 asset-required). The ready Office set includes two generated XLSX workflows: a direct-threaded-comment reply/resolve transaction and a formula-assumption transaction that changes only `Forecast!B9`, recalculates dependent cached values, protects `B10` plus the second-sheet baseline canary, permits only the target worksheet and workbook metadata parts to change, then checks native target-page change/baseline-page stability. It also includes the generated DOCX classic-comment and PPTX title/plain-text-notes workflows. Every ready Office route binds byte-bound audit/trace evidence, second import, and final LibreOffice/Poppler rendering. Existing clean XLSX comment (`d558a924ad63528f2b2dca5e1bbeb1fb0dc120a7`), DOCX (`b06f6ca067666b5774bb81fb23155f1cec50e694`), and PPTX (`450017eb8acb209f6ceb161d247f6c8059ab2571`) candidate trials each passed `100/100`; none is a repeat matrix or reference-Skill comparison. The clean formula-slice candidate at `3df67f5a083758eb1f6fc0c37cdc6c53f228e2eb` likewise passed `100/100` across 5 machine, 2 visual, 2 security, and 6 trace checks, binding tarball SHA-256 `720400c754d9cdbfaa949120c7939467702202d2cf507f7bef91d5068c1a7503`; it is one candidate trial, not a repeat matrix or reference-Skill comparison. A candidate/reference PDF preparation uses byte-identical package tarballs and prompts while varying only the copied Skill. Generic grading covers fail-closed branching, immutable file/directory inputs, read-only prompt/Skill/dependency trees, regular output types, and exact deliverables; the 20 corpus/PKI cases remain explicitly asset-required.
- The bounded contract-ID defect was closed through the official typed path: `replace_text` now preserves the source baseline/default style and accepts the observed `0.0000227pt` provider quantization difference inside a non-configurable tolerance capped at `0.0005pt`, while genuine overflow and rotated text remain fail-closed. The Skill now requires provider probe and route planning before mutation and ships the canonical `open-office-artifact-tool.pdf-audit.v1` schema plus a validator that recomputes source/output hashes. One fixed candidate trial and one same-prompt, same-tarball reference-Skill trial each scored 100/100 across machine, visual, security, and trace evidence. Repeat trials and the remaining case graders/corpus are still open.
- The active-content public-sanitize fixed matrix passes candidate `3/3` and reference Skill `3/3`, all at 100/100 across machine, visual, security, and trace. The six runs bind clean commit `39fa301dcb1005f2848282e6e63da1e934104821`, byte-identical package SHA-256 `e78e18c0f8f1cffe301ae1f2ea17e882bc879b3044914033e24b0b11ac0e8b69`, identical prompt/input/oracle fingerprints, and fixed but distinct candidate/reference Skill fingerprints. Independent pypdf evidence proves root/additional JavaScript, Launch/SubmitForm, five attachments, invisible text, a comment, a populated widget, and personal metadata in each immutable source and zero active structure/canary residue in every output; Poppler changes stay inside the expected form/comment masks. Historical bypass and interpreter-drift runs remain defect evidence rather than passes. The evaluator ignores help-only invocations and permits independently scheduled probe/plan completion only when both precede the real typed edit; low-level bypass and post-mutation gates remain strict.
- The AcroForm adapter now distinguishes text/choice strings from radio/checkbox PDF Name states and validates post-write field/widget appearances before transaction promotion. Its complete generated fixture contains five text widgets, two radio widgets, and one pre-checked checkbox. The clean fixed matrix binds commit `bffd35dbfdb94bb1183717703e7e55bfb83c3f3c`, package SHA-256 `7ab9e6a30035df5d0ef7ee9990f3a0445152877e58a7f0d065ede9ddc1db300b`, and identical prompt/input/oracle fingerprints. Candidate passes `3/3` at 100/100; reference passes `2/3` at 100/100. The remaining reference run earned 70 raw machine/visual points but correctly scored zero after manually writing incremental PDF objects, reporting a non-pypdf provider, omitting canonical audit provenance, and bypassing provider preflight, typed fill, and post-mutation audit gates. All semantically successful fixed runs preserve TIN/signature/unselected-radio/checkbox pixels and editability, preserve the exact source prefix with one appended revision, and render every page through Poppler.
- The pypdf attachment-quarantine primitive covers document/page FileSpecs, duplicate and Unicode names, traversal-safe cross-platform naming, decoded-byte budgets, transactional extraction, raw identity/MIME/size/SHA evidence, and source/manifest provenance without opening payloads. Its fixed matrix binds clean commit `748fbb1d81ccfa14a594d6fed9bc6601866bfa95`, package SHA-256 `9cff93494c5b32e16394ce3b4fcffa1daf76ad6df57326dab0ced47d2a45b5bf`, and identical prompt/input/oracle fingerprints. Candidate passes `3/3` at 100/100; reference passes `2/3` at 100/100. The retained reference miss extracted the correct six payloads safely but used a custom Node parser and alternate manifest/audit contract, so missing pypdf/read-only/no-fallback preflight, typed extraction, canonical schema, and byte-binding evidence correctly forced the run to zero.
- The six-page greenfield tagged-accessibility workflow now has a clean fixed matrix: candidate `3/3` and reference Skill `3/3`, every run at 100/100 across machine, visual, security, and trace. All six records bind clean commit `2323a70331b93781dee37aa05198e4a73a7ec533`, byte-identical package SHA-256 `cfbcf5c76ba5fdb929dae27f2a0295d6da12694eec2150a926cebeecedefccb9`, identical prompt/input/oracle fingerprints, and fixed but distinct candidate/reference Skill fingerprints. Independent pypdf traversal proves title/language, H1-H3, one logical Table across pages 3-4, Figure alt, Link/StructParent/OBJR, reading order, and 12 running artifacts; Poppler/Pillow proves every page is nonblank, unclipped, and contains the expected physical table segments. The workflow separates modeled checks, optional veraPDF machine evidence, and required human PDF/UA judgment rather than claiming automatic conformance.
- The merge/reorder/selective-watermark clean fixed matrix passes candidate `3/3` and reference Skill `3/3`, every run at 100/100 across machine, visual, security, and trace. All six runs follow provider check/inspect/plan, the typed pypdf manifest primitive, typed Poppler comparison, output inspection, and multi-source canonical audit, and bind commit `90cbb9e0a5527a4620a28bb38aad8feeca895a3b`, byte-identical package SHA-256 `c3962993aee732c7a8e60282159409dfe08ca2c7c4f0dd59eb80468c28630ff5`, identical prompt/input/oracle fingerprints, and fixed but distinct candidate/reference Skill fingerprints. The comparator closed a real QA-discipline gap: a prior Agent deleted a correct output after subjectively misreading two thumbnails as black even though their renders were more than 98.7% white and byte-identical to successful trials.
- LibreOffice opened the shipped 26-slide reference template and produced a 26-page PDF; bounded custom-geometry icons rendered visibly. This local LibreOffice build substituted `Helvetica Neue`, so pixel parity with the checked-in preview images is not claimed and remains a visual-fidelity gap.

## Hosted evidence

The reference 2.8.24 Skill sync, Presentation view-properties codec, and completed DOCX note workflow candidate at commit `5cf2eb629f4612314d34c73041a5f1296a9ff145` passed the hosted Linux `ci` workflow in [GitHub Actions run 29569265350](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29569265350) on 2026-07-17. The run completed with conclusion `success` in 4m15s and covered generated protocol plus deterministic bundled-runtime verification, Chromium/LibreOffice/Poppler tool checks, the complete npm suite including the native Spreadsheet path migration, Presentation `view/grid/guides` import/export/fail-closed regressions, Documents note workflow, generated API-doc cleanliness, offline release metadata, the 421-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `177/177`.

The DOCX whole-block bookmark and Documents navigation candidate through commit `b161676d1e68df8acdc7bd3a4a3640fbcf267b3c` passed the hosted Linux `ci` workflow in [GitHub Actions run 29564975672](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29564975672) on 2026-07-17. The run completed with conclusion `success` in 5m09s and covered generated protocol plus deterministic bundled-runtime verification, Chromium/native-tool checks, the complete npm suite including the public bookmark/internal-link Skill workflow and fail-closed topology regressions, generated API-doc cleanliness, offline release metadata, the 421-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `174/174`.

The composition-root and complete Spreadsheet-domain layering candidate at commit `f8a4241a40448cc21ad1dce449edb6747d491473` passed the hosted Linux `ci` workflow in [GitHub Actions run 29563335625](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29563335625) on 2026-07-17. The run completed with conclusion `success` in 4m26s and covered protocol plus deterministic bundled-runtime verification, Chromium/native-tool checks, the complete npm suite including root/leaf Spreadsheet binding identity and codec dependency-direction regressions, generated API-doc cleanliness, offline release metadata, the 421-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `173/173`.

The complete Presentation-domain layering candidate at commit `f8d50cad14ae74ee6cef7610ba58a52ea0438514` passed the hosted Linux `ci` workflow in [GitHub Actions run 29562292218](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29562292218) on 2026-07-17. The run completed with conclusion `success` in 4m18s and covered protocol plus deterministic bundled-runtime verification, Chromium/native-tool checks, the complete npm suite including root/leaf Presentation binding identity and codec dependency-direction regressions, generated API-doc cleanliness, offline release metadata, the 420-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `173/173`.

The DOCX whole-paragraph tracked-change and Documents Skill candidate through commit `98dbdd7e8a0814f79830efb5e8a7e2675dfcf520` passed the hosted Linux `ci` workflow in [GitHub Actions run 29561312029](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29561312029) on 2026-07-17. The run completed with conclusion `success` in 4m11s and covered generated protocol plus deterministic bundled-runtime verification, Chromium/native-tool checks, the complete npm suite including the public tracked-change Skill example and irregular-topology regression, generated API-doc cleanliness, offline release metadata, the 419-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `173/173`.

The AGPL/MuPDF candidate at commit `40fcee0931c541c6f2cb1639ead0d10e2b76c7e6` passed the hosted Linux `ci` workflow in [GitHub Actions run 29558072285](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29558072285) on 2026-07-17. The run completed with conclusion `success` in 4m52s and covered npm install, protocol/runtime verification, Chromium/native-tool checks, the full npm suite, generated API-doc cleanliness, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `170/170`.

The embedded-XLSX/OLE and Presentation native-object layering candidate at commit `6119c54ae05d4b60fe562641e7aef10130581782` passed the hosted Linux `ci` workflow in [GitHub Actions run 29558401718](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29558401718) on 2026-07-17. The run completed with conclusion `success` in 4m22s and covered protocol/runtime verification, Chromium/native-tool checks, the full npm suite including payload-only OLE replacement and Presentation Skill regressions, generated API-doc cleanliness, offline release metadata, the 417-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `170/170`.

The Document-domain and shared text-range layering candidate at commit `765040c5827ee73c3c4645824688470e281c8be5` passed the hosted Linux `ci` workflow in [GitHub Actions run 29559403816](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29559403816) on 2026-07-17. The run completed with conclusion `success` in 5m22s and covered protocol/runtime verification, Chromium/native-tool checks, the full npm suite including root/leaf binding identity and native Documents workflow regressions, generated API-doc cleanliness, offline release metadata, the 419-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `170/170`.

The type-aware AcroForm candidate at commit `bffd35dbfdb94bb1183717703e7e55bfb83c3f3c` passed the hosted Linux `ci` workflow in [GitHub Actions run 29515214432](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29515214432) on 2026-07-16. The run completed with conclusion `success` and covered the full npm suite including AcroForm provider and independent grader regressions, deterministic OpenChestnut verification, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167`.

The active-content PDF sanitization candidate at commit `099bf1ce2f62ab992971e61b82641e6d6712a95d` passed the hosted Linux `ci` workflow in [GitHub Actions run 29501149008](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29501149008) on 2026-07-16. The run completed with conclusion `success` in 4m01s and covered deterministic OpenChestnut verification, Chromium/LibreOffice/Poppler tool checks, the full npm suite including the typed sanitize provider and independent grader regressions, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167`.

The Agent PromptBench scaffold at commit `70b2ddeea642de3729f8b7d7401bf10bace3be69` passed the hosted Linux `ci` workflow in [GitHub Actions run 29493964234](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29493964234) on 2026-07-16. The run completed with conclusion `success` and covered suite validation/tests inside the full npm gate, deterministic OpenChestnut verification, Chromium/LibreOffice/Poppler checks, API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167`.

The OpenChestnut XLSX sparkline candidate at commit `e8aa3e14249de346207f16b8fa24d7cb00b1253f` passed the hosted Linux `ci` workflow in [GitHub Actions run 29492891825](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29492891825) on 2026-07-16. The run completed with conclusion `success` in 4m02s and covered deterministic protocol/runtime verification, Chromium/LibreOffice/Poppler tool checks, the full npm suite, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167` including standard Office 2010 sparkline coverage.

The PDF provider-routing candidate at commit `b405ddd249c7c2f760c659c07e88495f3a3562f3` passed the hosted Linux `ci` workflow in [GitHub Actions run 29487829878](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29487829878) on 2026-07-16. The run completed with conclusion `success` in 3m59s and covered deterministic protocol/runtime verification, Chromium/LibreOffice/Poppler tool checks, the full npm suite including the provider contract tests, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge, and OpenChestnut 163-test execution. Optional Python providers remain an explicit local/environment gate rather than an undeclared hosted dependency.

The Documents native-workflow/OpenChestnut candidate through commit `e07e24382ff0259c7beefe27b0743d908a1f946f` passed the hosted Linux `ci` workflow in [GitHub Actions run 29483188346](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29483188346) on 2026-07-16. The run completed with conclusion `success` in 3m55s and covered protocol/runtime reproducibility, Chromium/native render tools, `npm test`, generated API-doc diff, offline release metadata, the registry-independent clean-install tarball, OfficeBridge, and OpenChestnut.

`npm run release:check` passes the source, documentation, package, license, JavaScript, and .NET gates. Its only remaining blocker is unavailable npm authentication. No `npm publish` or tag/release operation has been performed.

## Publishing

Before publishing:

1. Verify `package.json` and `package-lock.json` both declare `0.2.0`.
2. Rebuild and verify the OpenChestnut runtime from source.
3. Regenerate API documentation after the final public API change.
4. Inspect `npm pack --dry-run --json` for the required runtime/proto/Skill files and forbidden legacy files.
5. Run the tarball clean-install probe, not only source-tree tests.
6. Record the exact commit and hosted gate result used for publication.
