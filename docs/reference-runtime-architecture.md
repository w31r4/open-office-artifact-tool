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
- `native/OpenXmlWasm/src/OpenOffice.OpenXmlCodec` uses `DocumentFormat.OpenXml` 3.5.1 and `Google.Protobuf` 3.35.1 for bounded XLSX/DOCX/PPTX import/export and structured diagnostics.
- `native/OpenXmlWasm/src/OpenOffice.OpenXmlWasm` exposes one `[JSExport]` byte-in/byte-out method and publishes a trimmed `browser-wasm` AppBundle.
- `src/codecs/openxml-wasm.mjs` lazily initializes and caches that runtime. Family adapters such as `openxml-wasm-presentation.mjs` map JavaScript models to the public wire schema, fail closed when direct authoring leaves a modeled slice, and carry validated source-package preservation state across imported edits.
- `runtime/openxml-wasm` contains the reproducible release bundle, integrity manifest, CycloneDX SBOM, .NET license, and upstream third-party notices. It is built from this repository; it contains no reference-package artifact.
- The npm clean-install gate packs the real tarball, installs it in a temporary directory, removes `dotnet` from runtime `PATH`, and performs XLSX, DOCX, and PPTX export/import through the public package subpath.

The first slice covers workbook date systems, worksheets, primitive and cached-formula cells, merged ranges, row/column dimensions, gridline state, and frozen panes. Import stores one budget-checked source-package snapshot plus opaque part digests and relationship inventory in the public envelope. Second export verifies the snapshot SHA-256 and source identity, updates modeled fields in the original package through Open XML SDK, validates the complete result against Office 2021, and then rejects it unless all opaque part digests and relationships still match. This preserves imported styles, themes, tables, pivots, drawings, comments, validations, defined names, and arbitrary legal targets without pretending those features are semantically modeled by C#. Direct advanced authoring still uses the JavaScript codec, and missing/tampered preservation state remains fail-closed unless the caller explicitly accepts lossy rebuilding.

The DOCX slice models ordered paragraphs, limited direct run formatting, and table text. Imported body blocks carry their original body position, XML digest, semantic digest, and an editability flag. Unchanged advanced WordprocessingML remains in the hash-bound source package; attempts to modify hyperlinks, fields, comments/bookmarks, drawings, numbering, sections, tables, or opaque top-level blocks fail closed until those semantics are modeled. The complex business-brief fixture crosses this path and then runs the existing JS package/semantic plus LibreOffice/Poppler gates.

The PPTX slice models slide order/size, top-level rectangle/ellipse shapes, and a bounded DrawingML paragraph/inline subset. It authors and imports paragraph level, left/center/right/justified alignment, direct `marL`/signed `indent` list coordinates, direct `lnSpc`/`spcBef`/`spcAft` point-or-multiplier spacing, ordered ordinary text, `a:fld`, and `a:br` inlines, left/center/right/decimal `a:tabLst` stops, the direct list-marker choice (character, all standard auto-number schemes with optional start, explicit no-bullet, or a picture), three independent direct marker-style choices, inline styles, and bounded click links. Picture sources are either a content-addressed envelope `Asset` containing bounded PNG/JPEG/GIF/safe-SVG bytes or a non-fetched absolute http(s) URI; field 16 adds that choice without changing existing bullet, tab-stop, or text wire fields. Untransformed marker theme colors use additive field 17 and preserve one of the 16 public DrawingML scheme tokens rather than resolving it to an RGB snapshot. Paragraph coordinate value/delete choices use additive fields 18 through 21; three spacing oneofs use additive fields 22 through 30 and retain native `spcPts` versus `spcPct` identity. Clearing an imported direct coordinate or spacing slot restores inheritance while absence still preserves unknown source state. A shared C# asset catalog owns durable bytes and media reuse, while each `PptxSlideContext` owns only slide-local relationships. Every imported slide and top-level shape-tree element remains hash-bound to its source. Fixed-topology edits mutate existing PresentationML nodes in place, retain old media/relationships as opaque content, and permit only explicitly tracked image/link relationship and media-part additions. Transformed, unsupported-format, malformed, or unresolved picture bullets remain residual and fail closed if replaced. Transformed or non-sRGB/non-theme marker colors, custom-show/mouse-over links, field-local `a:pPr`, unknown break properties, underline, East Asian font, unsupported alignment, pictures, charts, groups, notes, masters, layouts, themes, and arbitrary recursive OPC graphs stay source-preserved. Post-write validation requires Office 2021 zero errors, exact guarded graph equality after removing only modeled additions, and byte equality for every package part except the changed slide XML, its changed relationship part, `[Content_Types].xml` when necessary, and exact tracked new media. The runnable `openxml-wasm-preservation` fixture crosses JS export, WASM list/layout/spacing/style/text/link/field/break/tab-stop/picture-bullet/theme-color edits, JS semantic/package verification, and Playwright plus LibreOffice/Poppler rendering.

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

### C# OpenXML-WASM owns

- opening and creating WordprocessingDocument, PresentationDocument, and SpreadsheetDocument packages from bounded byte streams;
- typed part, content-type, relationship, and schema element handling;
- semantic extraction into the project's public messages;
- semantic export from those messages;
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

1. **Implemented:** add a source-built `native/OpenXmlWasm` scaffold and a minimal versioned schema.
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
| Bounded `openxml-wasm-basic` fixture passes JS/WASM semantic import, inspect/resolve/verify, and skill QA | done locally and in hosted Linux CI |
| Existing formula-summary/arbitrary-path fixture passes inspect/resolve/verify through both codecs | done locally and in hosted run [`29257565233`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29257565233) |
| LibreOffice/Poppler render-backed cross-codec output passes | done locally and in hosted Linux CI for `openxml-wasm-basic` plus the three-sheet formula-summary |
| Full npm/docs/package/C# gates and hosted Linux CI pass for the committed milestone | done in hosted run [`29257565233`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29257565233) |

## Explicit non-goals

- Reimplementing the reference `Walnut` assembly or matching its private protobuf binary schema.
- Moving formulas, agent APIs, layout, rendering, or collaboration wholesale into C#.
- Requiring Microsoft Office, Windows, LibreOffice, or a local .NET SDK for normal package I/O.
- Treating Open XML validation as proof of application rendering fidelity.
- Deleting working JavaScript codecs before the WebAssembly replacement is proven.
