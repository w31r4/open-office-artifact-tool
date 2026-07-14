# Reference runtime architecture and clean-room direction

- Status: accepted direction, XLSX, DOCX, and PPTX WebAssembly vertical slices implemented; migration active
- Evidence snapshot: 2026-07-13
- Reference package: `office-artifact-tool@2.8.22`
- Reference runtime asset package: `@officer/walnut@0.1.210`

## Decision

The long-term Office I/O path for `open-office-artifact-tool` should be a bundled, clean-room .NET WebAssembly codec built with the public Microsoft Open XML SDK. JavaScript remains the agent-facing product layer. The existing Microsoft Office COM sidecar remains an optional renderer and compatibility oracle; it is not the core DOCX/PPTX/XLSX codec.

The target split is:

- JavaScript: artifact models, editing facades, formulas, inspect/resolve/help, layout, preview rendering, workflow orchestration, codec selection, and bounded input preflight.
- C# in bundled .NET WebAssembly: typed OPC/Open XML package reading and writing for DOCX, PPTX, and XLSX, package relationship/content-type handling, format-version validation, and preservation of package data that is outside the semantic model.
- A versioned public wire schema: independently designed messages shared by JavaScript and C#. The schema is an implementation boundary, not a clone of the reference runtime's private protobuf contract.
- Playwright/Skia/sharp/canvas: model preview and raster rendering.
- LibreOffice, Poppler, Open XML SDK validation, and optional Microsoft Office automation: render-backed and application-compatibility QA.
- PDF.js and the clean-room PDF implementation: PDF parsing/writing. The observed reference package does not expose a PDF model or PDF codec from its main module.

The existing JavaScript OOXML codecs remain useful as a fallback, a patch/inspection layer, a test oracle, and a migration safety net. They should not be deleted before the WebAssembly path proves equal or better behavior fixture by fixture.

## Clean-room boundary

This study used only observable and redistributable package facts:

- package manifests and file names;
- runtime boot metadata;
- public ESM exports and public object methods;
- public `.NET WebAssembly` assembly export names;
- shipped README, skills, and smoke tests;
- black-box calls using generated empty artifacts and byte counts.

It did not decompile, disassemble, copy, or translate the reference JavaScript bundle, the `Walnut` assembly, its WebAssembly binaries, native bindings, method bodies, or protobuf implementation. No reference runtime artifact may be copied into this project.

## Package evidence

The checked-out reference submodule is `e6dd700a79dae65ea53b0e775fa72ef6b8cc5376`. Its package manifest declares two bundled runtime dependencies:

- `@officer/walnut@0.1.210`
- `skia-canvas@^3.0.6`

Approximate installed footprint from `du -sk`:

| Area | KiB | Role |
| --- | ---: | --- |
| `dist/` | 10,880 | Compiled ESM product/model/formula/layout layer |
| `node_modules/@officer/walnut` | 12,624 | .NET WebAssembly Open XML codec runtime |
| `node_modules/skia-canvas` | 22,108 | Native text/raster canvas runtime |
| `skills/` | 6,460 | Agent instructions, scripts, templates, and QA workflows |

The main `dist/artifact_tool.mjs` file is 11,118,246 bytes. Public reflection found 1,163 ESM exports, including approximately 498 all-uppercase spreadsheet formula functions and 59 exports whose names explicitly mention proto conversion. These counts are inventory evidence, not a compatibility promise.

Node process-report evidence also shows that importing the main ESM module loads `skia-canvas/lib/skia.node` into the process. Skia is therefore a separate native text/raster dependency, not part of the C# Office codec. This project does not need to copy that eager-loading choice: renderer dependencies should remain optional or lazy so Office model and codec use does not fail on an unsupported native canvas build.

The public model surface shows substantial JavaScript ownership:

- `Workbook`: model collections, formulas/recalculation, trace, inspect/help/resolve/apply, collaboration, render, HTML, image, and export behavior, plus `toProto()`.
- `Presentation`: compose, model collections, inspect/help/resolve/apply, collaboration, render context, `toProto()`, `toPresentationBytes()`, and `toArtifactBundle()`.
- `DocumentModel`: document model/layout/export behavior, `toProto()`, and `toDocumentBytes()`.
- `SpreadsheetFile`, `PresentationFile`, and `DocumentFile`: the public Office file import/export facades.

The public `toProto()` methods return plain JavaScript objects. Top-level examples include workbook sheets/styles/theme/pivot caches, presentation slides/layouts/charts/images/theme, and document elements/sections/styles/comments/notes. `Presentation.toArtifactBundle()` returns an object with `rootArtifactId` and `artifacts`, indicating a second, package-oriented preservation path alongside the semantic presentation message.

## .NET WebAssembly evidence

`@officer/walnut` exposes runtime assets rather than a conventional JavaScript API. Its `blazor.boot.json` reports:

- `mainAssemblyName: "Walnut"`;
- 31 `.wasm` files in the runtime directory;
- `linkerEnabled: true`;
- `globalizationMode: "invariant"`.

The boot resource graph contains:

- `Walnut` at about 2.8 MiB;
- `DocumentFormat.OpenXml` at about 4.2 MiB;
- `DocumentFormat.OpenXml.Framework` at about 276 KiB;
- `Google.Protobuf` at about 312 KiB;
- `System.IO.Packaging`, `System.IO.Compression`, XML, collections, and other .NET assemblies;
- the .NET native/runtime JavaScript and WebAssembly host files.

The separate `DocumentFormat.OpenXml.Framework` and typed `DocumentFormat.OpenXml` resources strongly indicate Open XML SDK 3.x or later, because Microsoft split the framework package starting with 3.0. The exact reference version is not proven by the package metadata and is not required for the clean-room implementation.

Loading the runtime with its standard `dotnet.js` host and calling `getAssemblyExports("Walnut")` exposes only six codec classes and eight canonical methods:

| .NET export class | Method | Observed JS arity | Boundary |
| --- | --- | ---: | --- |
| `DocxReader` | `ExtractDocxProto` | 2 | DOCX bytes to document message bytes |
| `DocxExport` | `ExportProtoToDocx` | 1 | Document message bytes to DOCX bytes |
| `PptxReader` | `ExtractSlidesProto` | 2 | PPTX bytes to presentation message bytes |
| `PptxReader` | `ExtractPresentationArtifactBundle` | 2 | PPTX bytes to package-preserving bundle bytes |
| `PptxExport` | `ExportProtoToPptx` | 1 | Presentation message bytes to PPTX bytes |
| `PptxExport` | `ExportArtifactBundleToPptx` | 1 | Bundle bytes to PPTX bytes |
| `XlsxReader` | `ExtractXlsxProto` | 2 | XLSX bytes to workbook message bytes |
| `XlsxExport` | `ExportProtoToXlsx` | 1 | Workbook message bytes to XLSX bytes |

The second reader argument is optional in the observed empty-file calls. Its precise private meaning was not investigated because it is not necessary to establish the architecture and is outside the clean-room target.

## Black-box boundary proof

Direct public export calls were exercised with empty clean-room model objects. Only byte types, lengths, ZIP signatures, and roundtrip success were recorded.

| Format | Input/package bytes | Message bytes | Re-export bytes | Result |
| --- | ---: | ---: | ---: | --- |
| XLSX | 3,315 | 980 | 3,312 | Returned a valid ZIP/OOXML package |
| PPTX | n/a | 2,477 authored message | 7,668 | Returned a valid ZIP/OOXML package; extracted message was 3,581 bytes |
| DOCX | n/a | 2,086 authored message | 3,973 | Returned a valid ZIP/OOXML package; extracted message was 2,169 bytes |

For the same empty PPTX, the core presentation extraction returned 3,581 bytes and the artifact-bundle extraction returned 3,629 bytes. Both corresponding exporters produced valid OOXML ZIP packages. This proves a dual PPTX I/O contract without requiring knowledge of the private message fields.

Instrumenting only public JavaScript methods during normal file export observed calls to `Workbook.toProto()`, `Presentation.toProto()` plus `toPresentationBytes()`, and `DocumentModel.toProto()` plus `toDocumentBytes()`. This confirms that JavaScript owns model-to-message conversion while the .NET layer consumes serialized bytes.

The runtime also loaded successfully under Node with `PATH=/usr/bin:/bin`, where the Homebrew `dotnet` CLI was not visible, and still exposed all six codec classes. The SDK is required to build the assets, but end users receive and execute the published `dotnet.js`, runtime WebAssembly, and managed assemblies from npm.

One local timing sample measured about 1,058 ms for the large main-module import, 1 ms for an empty presentation model creation, 320 ms for the first PPTX export, and 20 ms for a second export in the same process. These machine-specific numbers are not a performance contract, but the warm-call difference is consistent with a cached runtime/codec instance and supports the requirement to initialize once per process.

## What the Open XML SDK contributes

The Microsoft Open XML SDK provides strongly typed package, part, relationship, and schema element APIs for WordprocessingML, PresentationML, and SpreadsheetML. It removes a large class of hand-maintained namespace, schema-order, relationship, content-type, and package-ownership errors. It also provides validation against Office format versions.

It does not provide the high-level agent artifact model, spreadsheet calculation engine, layout engine, presentation compose system, renderer, or application-equivalent pagination. The official project explicitly describes the SDK as a low-level Open XML/OPC API rather than a high-level productivity abstraction. The reference package's large JavaScript surface is consistent with that division.

Primary public references:

- [Microsoft Open XML SDK getting started](https://learn.microsoft.com/en-us/office/open-xml/getting-started)
- [Microsoft Open XML SDK repository and MIT license](https://github.com/dotnet/Open-XML-SDK)
- [.NET WebAssembly browser/Node application templates and JS interop](https://learn.microsoft.com/en-us/aspnet/core/client-side/dotnet-interop/wasm-browser-app)
- [Protocol Buffers C# generated code guide](https://protobuf.dev/reference/csharp/csharp-generated/)

## Current implementation state

The first source-built boundary now exists:

- `proto/open_office/artifact/v1/office_artifact.proto` is the independently designed, versioned public wire contract. Buf generates the shipped JavaScript binding; `Grpc.Tools` generates the C# binding from the same source.
- `native/OpenChestnut/src/OpenChestnut.Codec` uses `DocumentFormat.OpenXml` 3.5.1 and `Google.Protobuf` 3.35.1 for bounded XLSX/DOCX/PPTX import/export and structured diagnostics.
- `native/OpenChestnut/src/OpenChestnut.Runtime` exposes one `[JSExport]` byte-in/byte-out method and publishes a trimmed `browser-wasm` AppBundle.
- `src/codecs/open-chestnut.mjs` lazily initializes and caches that runtime. Family adapters such as `open-chestnut-presentation.mjs` map JavaScript models to the public wire schema, fail closed when direct authoring leaves a modeled slice, and carry validated source-package preservation state across imported edits.
- `runtime/open-chestnut` contains the reproducible release bundle, integrity manifest, CycloneDX SBOM, .NET license, and upstream third-party notices. It is built from this repository; it contains no reference-package artifact.
- The npm clean-install gate packs the real tarball, installs it in a temporary directory, removes `dotnet` from runtime `PATH`, and performs XLSX, DOCX, and PPTX export/import through the public package subpath.

OpenChestnut is the canonical implementation identity. The old `codecs/openxml-wasm` module and old build/test script names are deprecated compatibility aliases that delegate to the canonical module/scripts. There is no legacy runtime directory or assembly; status and exported file metadata identify the codec as `open-chestnut` even when invoked through an old alias.

The first slice covers workbook date systems, worksheets, primitive and cached-formula cells, merged ranges, row/column dimensions, gridline state, and frozen panes. Import stores one budget-checked source-package snapshot plus opaque part digests and relationship inventory in the public envelope. Second export verifies the snapshot SHA-256 and source identity, updates modeled fields in the original package through Open XML SDK, validates the complete result against Office 2021, and then rejects it unless all opaque part digests and relationships still match. This preserves imported styles, themes, tables, pivots, drawings, comments, validations, defined names, and arbitrary legal targets without pretending those features are semantically modeled by C#. Direct advanced authoring still uses the JavaScript codec, and missing/tampered preservation state remains fail-closed unless the caller explicitly accepts lossy rebuilding.

The DOCX slice models ordered paragraphs, limited direct run formatting, table text/merge geometry/fixed direct formatting, bounded whole-paragraph hyperlinks, bounded whole-paragraph simple fields, and direct or paragraph/numbering-style-linked numbered single-run paragraphs. `DocxHyperlinkCodec` recognizes one direct `w:hyperlink` with one text run, models external http(s) or internal bookmark targets plus text/tooltip/history, and changes those fields in place. `DocxPartContext` owns relationship additions/removals only for the hyperlink being edited and caches read-only `styles.xml`/`numbering.xml` documents without materializing Open XML SDK roots. `DocxFieldCodec` recognizes `w:fldSimple`, exposes its instruction and visible result, and edits one-result-run fields only when their command belongs to an explicit non-fetching catalog. `DocxTableGeometry` maps physical cells onto the effective `w:tblGrid`, including row offsets, `w:gridSpan`, `w:vMerge` restart/continuation state, effective row spans, and per-cell editability. Additive table wire fields leave the existing row text matrix at field 1. Source-free authoring consumes that same public shape only when one record exists for every physical cell, every row covers `grid_columns`, continuations exactly match the preceding row's grid column/span, continuation text is empty, and each origin's declared row span equals the chain. `DocxTableFormatting` owns optional wire field 3 as one complete profile: fixed dxa table width/indent, one positive width per logical grid column, derived merged-cell widths, four logical margins, six uniform direct RGB borders, and a filled/bold first row. Presence is not a claim about effective table-style formatting. Width sums, colors, dimensions, and Word border-size bounds are validated in JS and C# before output. `DocxTableCodec` emits schema-ordered `tblPr`/`tblGrid`/`tcPr`; `DocxDirectStyles` supplies the minimal default `TableGrid` definition. Reimport recognizes only this exact direct profile and exposes it through the JS model, but source-bound export rejects any formatting mutation in both adapters while the residual hash protects all other properties. Arbitrary style graphs, theme/conditional/per-cell formatting, gaps, overlaps, and malformed chains fail closed. Imported geometry remains source-bound: only simple unmerged, horizontally merged, and vertical-origin text nodes are updated in place while all table/row/cell/paragraph/run properties remain residual-hashed. `DocxNumberedParagraphCodec` owns paragraph assignment and fixed-topology text updates; the dedicated `DocxNumberingResolver` selects style-inherited levels through `w:lvl/w:pStyle`, follows `w:styleLink` and recursive `w:numStyleLink` → numbering-style `numId` chains, then applies instance level/start overrides. Paragraph-style `ilvl` is ignored by standard; direct paragraph `ilvl` remains authoritative. Additive wire field 7 carries the resolved numbering-style identity. Only the single text node is editable. Missing, duplicate, cyclic, over-deep, ambiguous, or unresolved style/definition graphs stay source-preserved and read-only instead of falling back to an editable plain paragraph. Imported body blocks carry their original body position, XML digest, semantic digest, residual digest where applicable, and editability evidence. The complex business-brief fixture retains its imported merge/edit/preservation gate; the `open-chestnut-merged-table` fixture starts from the public JS model, authors a three-column merged table with non-default fixed formatting directly through WASM, reimports and edits its origin, then passes package/model plus real Playwright/LibreOffice/Poppler gates.

Source-free numbering is a whole-document transaction. The JavaScript adapter assigns package-local direct `numberingId`/`abstractNumberingId` values when callers omit them and rejects conflicting definitions before crossing the wire. `DocxDirectNumbering` independently preflights every numbered block, groups instances and shared abstract definitions, and emits schema-ordered `w:abstractNum`/`w:lvl`/`w:num` plus direct paragraph assignments. This bounded slice covers text-marker levels 0 through 8; picture-bullet assets and style-linked graphs remain on the JavaScript codec.

Numbering-definition mutation is a separate batch transaction rather than a paragraph-local side effect. `DocxCodec` first verifies every source element/semantic binding, then `DocxNumberingEditPlanner` groups requested changes by direct `numId`/level. A group is editable only when every affected top-level paragraph has fixed editable topology, every requested `numberFormat`/`start`/`levelText` target is identical, and no nested paragraph or other package part references that instance. The planner materializes or updates a complete instance-local `w:lvlOverride`, never the shared `w:abstractNum`; paragraph identity/level/style remain residual-protected. Partial groups, style-inherited or style-linked graphs, duplicate/malformed overrides, and cross-part use fail closed. The `open-chestnut-numbering-edit` fixture now begins with bundled-WASM graph authoring, crosses bundled-WASM import/edit/export, then proves the override through reimport, Playwright, LibreOffice, and Poppler. Hosted Linux run [`29340497565`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29340497565) passes that real-render fixture, all 55 codec tests, the 208-file clean install, and two byte-identical 38-file/12,960,958-byte runtime builds.

The PPTX slice models slide order/size, a hash-bound Slide → Layout → Master package identity chain, direct Master/Layout backgrounds, owner-local Master/Layout placeholder text, top-level rectangle/ellipse shapes, bounded master `p:txStyles` defaults, and a bounded DrawingML paragraph/inline/text-body subset. Imported master, layout, placeholder, and slide IDs are stable locators inside the bound source package; they are not a promise of identity after an unrelated file is re-imported. Every slide records its native layout relationship, every layout records its owning master, and every direct placeholder records its raw owner-local shape-tree position plus XML/semantic hashes. Source export rejects graph, name, placeholder identity/topology/shape-level, layout-content, or slide-binding changes. From-scratch authoring currently creates one canonical master and its internal blank layout; template placeholders still require a validated source package.

Shape paragraphs and shape-local `a:lstStyle` levels 0 through 8 cover alignment, direct margins and signed indents, point/multiplier spacing, tab stops, bounded default-run and inline styles, text/field/line-break inlines, character/auto-number/no-bullet/picture markers, bounded click links, and bounded `a:bodyPr` layout. `PptxParagraphPropertiesCodec` coordinates the shared paragraph-property subset, `PptxListStyleCodec` owns shape-local list-level topology, `PptxBodyPropertiesCodec` owns text-frame properties and AutoFit, `PptxMasterTextStylesCodec` owns title/body/other master levels, `PptxBackgroundCodec` owns direct Master/Layout solid or background-style-reference fills without flattening inheritance, and `PptxPlaceholderCodec` owns only direct `p:ph` identity plus local text. `PptxPartContext` generalizes relationship ownership across Slide, Master, and Layout parts while the shared asset catalog owns content-addressed picture bytes.

Missing master-style and background fields preserve source state. Explicit master-level deletion removes only modeled properties and drops the native level node only when no residual attributes or children remain. Direct background replacement/addition accepts only `p:bgPr/a:solidFill` or `p:bgRef` with one untransformed RGB/theme color and a bounded reference index; direct removal, gradients, patterns, images, groups, effects, and transforms remain source-bound. Placeholder name/type/index and shape-tree position are identity; source-bound editing permits only fixed paragraph/inline-topology local text/body/list/link changes, while geometry/fill/shape style and effective inherited values remain source XML. Fixed-topology edits retain old media and relationships, validate against Office 2021, compare the guarded OPC graph, and require byte equality or residual equality outside explicitly changed slide/master/layout content. Custom-show/mouse-over links, field-local paragraph properties, WordArt/3D/compatibility features, placeholder geometry/shape styles, themes, charts, groups, notes, and arbitrary recursive OPC content remain source-preserved and fail closed when replaced. The runnable `open-chestnut-preservation` fixture crosses JavaScript authoring, bundled-WASM import, master/layout background and owner-local placeholder edits, a Layout-owned hyperlink addition, master-title edit, loss-aware master-body-level deletion, Master-owned external picture-marker addition, shape text/body/list edits, JavaScript semantic/package verification, and Playwright plus LibreOffice/Poppler rendering.

The placeholder slice passes 34 codec tests, a reproducible 38-file/12,864,199-byte runtime build, the 192-file clean-install package, and real Chromium/LibreOffice/Poppler gates. Hosted Linux run [`29298306670`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29298306670) passes that complete gate on `main`.

The subsequent DOCX hyperlink slice passes hosted Linux run [`29299603126`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29299603126). The bounded simple-field slice raises the gate to 39 codec tests, a reproducible 38-file/12,886,215-byte runtime build, the 195-file clean-install package, and real Chromium/LibreOffice/Poppler-backed document checks. Hosted Linux run [`29301993212`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29301993212) passes that complete milestone on `main`.

The OpenChestnut numbering-style-linked numbered-paragraph milestone passes hosted Linux run [`29327443007`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29327443007): canonical and compatibility APIs, 47 C# tests, the 202-file no-local-dotnet clean install, Chromium/LibreOffice/Poppler, OfficeBridge, and two byte-identical 38-file/12,913,854-byte runtime builds all pass on `main`. This evidence covers `w:lvl/w:pStyle`, `styleLink`, recursive `numStyleLink`, instance start overrides, and additive numbering-style wire identity.

The subsequent merge-aware DOCX table slice passes 49 local C# tests plus bundled-WASM and real Documents-skill gates with a 38-file/12,925,116-byte macOS runtime. Hosted Linux run [`29330878218`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29330878218) passes the same 49-test gate, the 203-file no-local-dotnet clean install, real Chromium/LibreOffice/Poppler checks, OfficeBridge tests, and two byte-identical 38-file/12,925,118-byte runtime builds.

The source-free merged-table authoring slice raises the current local gate to 51 C# tests, the 205-file no-local-dotnet clean install, and a 38-file/12,931,772-byte macOS runtime. Its dedicated runnable fixture uses OpenChestnut for initial DOCX creation rather than injecting XML into a JavaScript-authored package, and passes Office 2021 validation plus real Chromium/LibreOffice/Poppler QA. Hosted Linux run [`29336666891`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29336666891) passes the same 51-test/205-file gate, real browser/native render workflow, OfficeBridge, and two byte-identical 38-file/12,931,774-byte runtime builds.

The imported table-formatting edit increment keeps the local codec gate at 56 tests, the 209-file no-local-dotnet clean install, two matching 39-file reproducibility snapshots, and a 38-file/12,981,948-byte macOS runtime. Its public wire shape is unchanged. “Source-bound” now means an imported table must retain the complete recognized profile and fixed geometry, not that its modeled formatting values are immutable: width, indent, grid/cell widths, margins, uniform borders, and header fill update in place. Residual hashing canonicalizes only those exact fields, with explicit tests preserving unknown row and paragraph properties and rejecting profile removal or unrecognized formatting. The merged-table fixture proves text plus formatting edits through bundled WASM, real Chromium, and LibreOffice/Poppler. Hosted Linux runs [`29343501551`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29343501551) and [`29343983983`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29343983983) pass the 56-test/209-file gate, OfficeBridge 5/5, real browser/native-render workflow, and two byte-identical 38-file/12,981,950-byte runtime builds.

The XLSX number-format increment adds `CellArtifact.number_format_code` as additive field 4 and delegates SpreadsheetML style-table work to `XlsxNumberFormatCodec`. Source-free export emits a minimal valid stylesheet and deduplicates built-in/custom format records. Source-preserving export clones the selected cell XF, retaining every unmodeled style property, and avoids loading changes back into `styles.xml` when the effective format is unchanged. Invalid style indices, missing custom formats, unmapped locale-dependent built-ins, oversized codes, and control characters return structured `invalid_cell_number_format` failures. The OpenChestnut→OpenChestnut spreadsheet fixture now carries custom and percentage formats through authoring/import/export plus Playwright and LibreOffice/Poppler. The local codec gate is 60 tests, the 210-file no-local-dotnet clean install passes, two isolated 39-file builds are byte-identical, and the macOS runtime is 38 files/12,997,820 bytes. Hosted run [`29346231840`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29346231840) passes the same gate with a 12,997,822-byte Linux runtime, OpenChestnut 60/60, OfficeBridge 5/5, and real Chromium/LibreOffice/Poppler.

The XLSX native-formula increment adds `CellArtifact.formula_metadata` as additive field 5 with independent shared/legacy-array kind, shared index, and bounded A1 reference. The semantic model stores an expanded formula on every shared member while `XlsxFormulaCodec` derives the top-left master, validates the complete rectangular group, and emits master/follower `t/si/ref` records. Legacy arrays remain one top-left anchor plus `ref`; data-table, malformed, incomplete, overlapping, nested-formula, oversized, or unknown topology fails closed. Import expands followers using source-built relative A1 translation, including absolute references, protected strings/structured selectors, and digit-bearing function names. Source-preserving export compares formula semantics separately from cached value and number format, so those unrelated edits retain the exact source `<f>` XML. A formula edit with unmodeled attributes is rejected; writing through the JavaScript range facade detaches the full shared group before producing normal formulas. Group-range occupancy is registered once per group on both sides, keeping validation linear in represented topology instead of quadratic, and total represented formula topology is capped at 1,048,576 cells before iteration. The OpenChestnut fixture now carries shared and array records through authoring, import, source-preserving export, JavaScript reimport, Playwright, and native LibreOffice/Poppler QA. The local gate is 65 C# tests, 38 runtime files/13,037,756 bytes, two byte-identical isolated 39-file builds, and a 211-file no-local-dotnet package. Hosted runs [`29348628745`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29348628745) and [`29349197886`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29349197886) pass the base slice and topology hardening; the latter proves the same 65-test/211-file/real-render gates plus two byte-identical builds around a 13,037,758-byte Linux runtime.

The XLSX static-style increment adds `CellArtifact.style` as additive field 6. Its independent messages carry font, pattern fill, nine border edges and flags, alignment, protection, and RGB/theme/indexed/automatic colors with optional tint; number-format field 4 remains wire-compatible. `XlsxCellStyleCodec` becomes the sole stylesheet owner so one cell mutation derives exactly one XF while `XlsxNumberFormatCodec` remains a pure built-in-format vocabulary helper. Source-free export emits the required default font/fills/border and deduplicates complete resources/XFs. Source-bound export clones the selected XF and referenced resources, appends modeled replacements, and preserves residual XML such as font schemes and quote-prefix state instead of mutating shared records. Custom workbook themes, gradient fills, named/cell style inheritance, and other unmodeled style semantics remain source-preserved and fail closed when replacement would be lossy. Invalid font/color/fill/border/alignment/protection inputs return structured `invalid_cell_style` errors. Four new C# tests cover complete source-free authoring with Office 2021 validation, source-bound residual preservation, invalid inputs, and deduplication; the local gate is 69 tests, a 38-file/13,096,124-byte bundled runtime, and a 212-file no-local-dotnet package. The runnable spreadsheet fixture carries patterned/theme static styles through OpenChestnut authoring, import, source-preserving export, JavaScript reimport, Playwright, and native LibreOffice/Poppler QA. Hosted Linux run [`29351669274`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29351669274) passes the same 69-test/212-file/real-render gate plus two byte-identical builds around a 38-file/13,096,126-byte runtime.

The XLSX workbook-theme increment adds `WorkbookArtifact.theme` as additive field 4. Its independent message carries the theme name, the fixed DrawingML color-scheme order (`dk1`, `lt1`, `dk2`, `lt2`, six accents, hyperlink, followed hyperlink), and a source binding; it does not reuse style or package identity. `XlsxThemeCodec` writes a complete source-free `ThemePart`, recognizes RGB and system-color-with-`lastClr` inputs, and permits only the modeled name/RGB slots to change on a recognized imported theme. The imported font scheme, format scheme, extensions, and other residual XML remain in place. Per-color transforms remain residual unless that exact modeled slot is replaced. HSL/preset/scheme colors and incomplete or otherwise complex profiles remain exact-preserved but uneditable, so an attempted replacement fails closed instead of flattening their semantics. Four new C# tests cover complete authoring and Office 2021 validation, source-bound editing with residual preservation, unchanged complex-theme preservation and replacement rejection, plus invalid/tampered inputs. The local gate is 73 tests, a 38-file/13,116,092-byte runtime, and a 213-file no-local-dotnet package. Hosted Linux run [`29353440899`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29353440899) passes the complete gate with OpenChestnut 73/73, OfficeBridge 5/5, real Chromium/LibreOffice/Poppler, the 213-file clean install, and two byte-identical 38-file/13,116,094-byte runtime builds.

The XLSX worksheet-table increment adds `WorksheetArtifact.tables` at field 9 plus independent `SpreadsheetTableArtifact` and `SpreadsheetTableSourceBinding` messages. The semantic message owns name/range, headers/totals/filter-button choices, column names, style name, and first/last/row/column stripe toggles; cell data remains in the worksheet. The binding records worksheet/table part paths, relationship ID, XML/semantic hashes, and editability without treating package-local relationship IDs as durable object identity. `XlsxTableCodec` exclusively owns the bounded `tableParts` → TableDefinitionPart graph. Source-free export authors one relationship and complete table XML; source-bound edits retain part path/relationship identity and change only modeled XML. Query relationships, filter columns, differential styles, extensions, unknown attributes/children, and other complex profiles remain byte-exact hidden source slots and reject replacement. Imported add/remove, binding tamper, malformed ranges/names/columns, duplicate workbook-global names, and invalid header/filter profiles fail closed. Four C# tests plus bundled-JS, clean-install, and runnable-skill tests cover Office 2021 validation, semantic reimport, source identity, exact complex preservation, and structured errors. The local gate is 77 tests, a 38-file/13,144,252-byte runtime, and a 214-file package (about 4.8 MB compressed/16,494,888 bytes unpacked). Hosted Linux run [`29355927675`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29355927675) passes the complete gate with OpenChestnut 77/77, OfficeBridge 5/5, real Chromium/LibreOffice/Poppler, the 214-file clean install, and two byte-identical 38-file/13,144,254-byte runtime builds.

The table-column-formula increment keeps the preceding wire fields stable and adds `SpreadsheetTableArtifact.columns` at field 14. Each independent `SpreadsheetTableColumnArtifact` carries its name, calculated-column formula/array flag, totals function/label, and custom totals formula/array flag; formula strings share the public leading-`=` convention with worksheet cells. When both projections are present, legacy field-12 `column_names` remains authoritative for names so existing clients can rename imported columns without constructing the rich message. `XlsxTableCodec` reads/writes the native `calculatedColumnFormula`, `totalsRowFunction`, `totalsRowLabel`, and `totalsRowFormula` vocabulary, validates built-in versus custom/label profiles, includes the rich fields in the source semantic hash, and preserves the TableDefinitionPart path and relationship during edits. Cell values/formulas remain independently owned by `XlsxCodec`, which avoids hidden table XML mutations of the public worksheet grid. Filter columns, query-table graphs, differential styles, extensions, and unknown column attributes/children remain exact-preserved read-only profiles. Three additional C# tests cover Office 2021-valid author/import, identity-preserving source edits, and invalid formula/totals combinations; bundled-JS, JavaScript-codec, clean-install, and runnable-skill tests cover both public projections and real render QA. The complete local gate is 80 C# tests, a 38-file/13,155,516-byte macOS runtime, and the 214-file no-local-dotnet package. Hosted Linux run [`29357896900`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29357896900) passes OpenChestnut 80/80, OfficeBridge 5/5, real Chromium/LibreOffice/Poppler, the same 214-file clean install, and two byte-identical 38-file/13,155,518-byte runtime builds.

The table-filter increment adds `SpreadsheetTableArtifact.filters` at field 15 without changing fields 1–14. Each `SpreadsheetTableFilterArtifact` uses a zero-based table column index and exactly one independently versionable criteria branch: `SpreadsheetTableValueFilterArtifact` carries exact selected strings plus optional blank inclusion, while `SpreadsheetTableCustomFilterArtifact` carries one or two comparisons and explicit OR/AND composition. Supported custom operators are the six SpreadsheetML comparison tokens. `XlsxTableCodec` owns the corresponding `autoFilter/filterColumn/filters/customFilters` XML, includes it in the table semantic hash, and patches a recognized source table without changing its part path or worksheet relationship. Dynamic/date-group/color/icon/Top10 filters, per-column button overrides, sort state, query relationships, extensions, and unknown attributes remain exact-preserved read-only residuals; a dedicated dynamic-filter corpus test proves byte-exact second export. JavaScript and OpenChestnut adapters expose the same public table model, and the runnable fixture crosses both filter kinds through bundled WASM, package inspection, Playwright, and LibreOffice/Poppler. The local gate is 84 C# tests, a 38-file/13,172,924-byte macOS runtime, and a 214-file package at about 4.83 MB compressed/16,578,079 bytes unpacked. Hosted Linux run [`29359575561`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29359575561) passes OpenChestnut 84/84, OfficeBridge 5/5, real Chromium/LibreOffice/Poppler, the same 214-file clean install, and two byte-identical 38-file/13,172,926-byte runtime builds.

The C# `OfficeBridge` remains a separate optional Windows JSON stdin/stdout process for render/convert/application checks; it is not part of the core codec path.

## Target architecture

```mermaid
flowchart LR
    Agent["Agent workflows and skills"] --> JS["JavaScript artifact model"]
    JS --> API["inspect / resolve / help / layout / render / formulas"]
    JS --> Adapter["Office codec adapter"]
    Adapter --> Schema["Versioned public protobuf schema"]
    Schema --> Wasm["C# OpenXML codec in bundled .NET WASM"]
    Wasm --> Office["DOCX / PPTX / XLSX packages"]
    Adapter -. fallback and patch .-> JSCodec["Existing JavaScript OOXML codecs"]
    Office --> QA["LibreOffice / OpenXML validator / optional Microsoft Office QA"]
    API --> Raster["Playwright / sharp / canvas"]
```

### JavaScript owns

- stable agent-facing IDs and public object facades;
- editing and composition semantics;
- spreadsheet calculation, dependency graphs, and formula help;
- inspect, resolve, help, verify, layout, and preview APIs;
- renderer selection and visual QA;
- conversion between the public JS model and the project's own wire schema;
- codec policy such as `wasm`, `js`, or controlled fallback;
- input preflight, output limits, timeouts, and structured errors.

### C# OpenChestnut owns

- opening and creating WordprocessingDocument, PresentationDocument, and SpreadsheetDocument packages from bounded byte streams;
- typed part, content-type, relationship, and schema element handling;
- semantic extraction into the project's public messages;
- semantic export from those messages;
- validation and package-level serialization of native shared/legacy-array formula topology, while formula calculation remains in JavaScript;
- preservation of package parts and relationships not represented by the editable model;
- Open XML format-version validation and structured diagnostic records;
- codec-local part/count/byte/depth budgets.

It must not own agent IDs, workflow policy, formula evaluation, layout, raster rendering, or network access.

### Native Office automation owns

- optional Word/Excel/PowerPoint application render and conversion;
- recalculation, field update, and application-specific compatibility checks that Open XML alone cannot prove;
- Windows-only baselines.

It remains outside the default import/export path and outside the required npm runtime.

## Independent wire contract

The repository should publish its own `.proto` sources and generated C#/JavaScript bindings. The initial contract should include:

- a protocol version and artifact family;
- the semantic document/presentation/workbook payload;
- stable model IDs that are explicitly distinct from native OOXML IDs;
- binary assets with media type, digest, and bounded bytes;
- source-package identities needed for loss-aware editing;
- an opaque package graph for unsupported parts, internal/external relationships, and content types;
- import/export diagnostics and unsupported-feature records;
- deterministic unknown-field and version-skew behavior.

The opaque graph is essential. The reference runtime's separate PPTX artifact-bundle path is observable evidence that a semantic message alone is insufficient for high-fidelity third-party roundtrips.

The schema must be designed from this project's public model, ECMA-376/ISO 29500, OPC, and Open XML SDK types. It must not reuse reference field numbers, message names, binary fixtures, or generated code.

## Runtime packaging requirements

- Build assets from checked-in C# and `.proto` sources in CI.
- Bundle the published `dotnet.js`, native runtime WebAssembly, managed assemblies, boot manifest, and integrity metadata in the npm tarball.
- Load the runtime lazily on the first DOCX/PPTX/XLSX codec call and cache one initialized runtime per process.
- Require no end-user .NET SDK, Microsoft Office, LibreOffice, or network access for normal Office import/export.
- Keep build-time SDK/workload requirements separate from runtime requirements.
- Record all NuGet/runtime licenses in `THIRD_PARTY_NOTICES.md`, generate an SBOM, and fail package tests if required runtime assets are absent.
- Pin reproducible .NET, Open XML SDK, and protobuf versions.
- Do not ship symbols, source maps, debug files, or reference-package assets unless deliberately licensed and required.

The build is pinned by `global.json` to .NET SDK 8.0.128 and uses the 8.0.28 `wasm-tools`/`wasm-experimental` workload packs. NuGet restore runs in locked mode. Native Emscripten and managed wrapper bytes can differ across build hosts even though the output is portable WebAssembly, so CI performs two isolated builds on the same hosted runner and rejects any byte or file-inventory difference between them; the checked-in consumer bundle is separately covered by its integrity manifest, clean-install test, and functional codec gates. Consumers do not run this build toolchain.

## Migration plan

1. **Implemented:** add a source-built `native/OpenChestnut` scaffold and a minimal versioned schema.
2. **Implemented:** prove one end-to-end XLSX vertical slice: JS model to message bytes to C# Open XML SDK to XLSX, then back to the JS model.
3. **Implemented:** package the runtime and run the installed tarball with `dotnet` removed from runtime `PATH`.
4. **Implemented locally and in hosted Linux CI:** exercise the existing spreadsheet formula-summary and arbitrary-path fixtures through both `wasm` and `js` codecs. Compare semantic inspect output, modeled verification, Open XML validation, all-sheet Playwright output, and LibreOffice/Poppler native pages.
5. **Implemented for imported XLSX:** retain a hash-bound source package, edit modeled fields in place, and verify opaque part/relationship preservation after write. Broader third-party corpus roundtrips remain the next XLSX migration gate.
6. **DOCX first slice implemented locally and in hosted Linux CI:** ordered paragraph/run/table semantics, block-level source bindings, advanced-content preservation, complex runnable fixture, and no-local-dotnet package probe. Hosted run [`29264703143`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29264703143) passed the complete runtime/render/package/.NET gate.
7. **PPTX first slice implemented locally and in hosted Linux CI:** simple shape authoring/editing, slide/element source bindings, opaque native-object preservation from the first slice, Office 2021 validation, runnable render-backed fixture, and no-local-dotnet package probe. Hosted run [`29267402425`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29267402425) passed the complete runtime/render/package/.NET gate; broader third-party corpus evidence remains pending.
8. Make WebAssembly the default only after compatibility and failure-mode gates pass. Retain an explicit JavaScript fallback until the migration matrix is complete.
9. Keep PDF and render adapters on their existing independent paths.

## First vertical-slice acceptance criteria

The first WebAssembly migration milestone is not complete until every row is done:

| Acceptance gate | Status |
| --- | --- |
| Runtime artifacts are built only from this repository and public dependencies | done |
| Tarball contains runtime, integrity manifest, SBOM, licenses, source, and no reference artifacts | done locally and in hosted Linux CI |
| Clean temporary npm install imports/exports XLSX with `dotnet` absent from runtime `PATH` | done locally and in hosted Linux CI |
| Minimal public workbook fixture roundtrips through JavaScript model and WebAssembly codec | done |
| Generated XLSX passes package inspection and Open XML SDK Office 2021 validation with zero errors | done |
| Unknown OPC content is detected and second export is fail-closed unless explicitly lossy | done: hash/source/graph-bound snapshot preservation plus post-write opaque digest/inventory verification |
| Malformed/oversized input produces bounded structured errors | partial: byte/ZIP/part/sheet/cell/path/ratio budgets covered; broader malformed corpus remains todo |
| Runtime initialization is lazy, cached, and concurrency-tested | done |
| Bounded `open-chestnut-basic` fixture passes JS/WASM semantic import, inspect/resolve/verify, and skill QA | done locally and in hosted Linux CI |
| Existing formula-summary/arbitrary-path fixture passes inspect/resolve/verify through both codecs | done locally and in hosted run [`29257565233`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29257565233) |
| LibreOffice/Poppler render-backed cross-codec output passes | done locally and in hosted Linux CI for `open-chestnut-basic` plus the three-sheet formula-summary |
| Full npm/docs/package/C# gates and hosted Linux CI pass for the committed milestone | done in hosted run [`29257565233`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29257565233) |

## Explicit non-goals

- Reimplementing the reference `Walnut` assembly or matching its private protobuf binary schema.
- Moving formulas, agent APIs, layout, rendering, or collaboration wholesale into C#.
- Requiring Microsoft Office, Windows, LibreOffice, or a local .NET SDK for normal package I/O.
- Treating Open XML validation as proof of application rendering fidelity.
- Deleting working JavaScript codecs before the WebAssembly replacement is proven.
