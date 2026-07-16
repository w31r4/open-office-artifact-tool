# OpenChestnut

OpenChestnut is this repository's clean-room C# codec for XLSX, DOCX, and PPTX. It uses the public Microsoft Open XML SDK and the independently defined protocol in `proto/open_office/artifact/v1/office_artifact.proto`.

It is the only Office codec used by version 0.2. The JavaScript package supplies the public artifact model and wire adapter; it does not contain a fallback Office parser/writer. PDF is outside OpenChestnut.

## Projects

- `OpenChestnut.Codec` implements package validation and the XLSX/DOCX/PPTX codecs.
- `OpenChestnut.Runtime` exposes the byte-in/byte-out WebAssembly entry point.
- `OpenChestnut.Tests` covers supported creation/import/edit/export profiles, opaque preservation, and fail-closed cases.

Protocol version 2 removes `allow_lossy`; the removed field name and number are reserved. Opaque content can be exported only from a validated, hash-bound source package. Unsupported edits, source-evidence mismatch, unsafe OPC paths, invalid relationships/content types, and missing runtime data return structured failures.

## Supported 0.2 slices

- XLSX: cells/formulas/styles, merges/dimensions/freeze panes, tables, PNG/JPEG images, bar/line/pie charts, standard Office 2010 line/column/stacked sparklines with reversible range mappings, basic data validation and conditional formatting, and one-level threaded comments.
- DOCX: styles, paragraph/run formatting, pages/sections, headers/footers, PAGE/simple fields, PNG/JPEG images, lists, links, classic comments, and fixed-geometry tables.
- PPTX: text boxes/round rectangles, basic fill/line/shadow, straight/polyline connectors and arrows, source-free bar/line/pie charts, images, fixed tables, rich text/lists/links, and source-bound Master/Layout preservation.

Imported objects outside these modeled profiles remain opaque and unchanged. They do not become permission for lossy rewriting.

## Build and test

```sh
dotnet test native/OpenChestnut/OpenChestnut.sln --configuration Release
npm run build:open-chestnut
npm run verify:open-chestnut-build
```

The bundled consumer runtime is written to `runtime/open-chestnut`. The npm package includes that runtime, manifest, SBOM, and license notices, but excludes this C# source tree and build output. Installed consumers therefore do not need a local .NET SDK.
