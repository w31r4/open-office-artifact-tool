# Release readiness

Status date: 2026-07-14

## Current package

- Package: `open-office-artifact-tool`
- Version: `0.1.0`
- Publish state: publish-ready, not published from this environment

## Required local gates

The current release gate is:

```sh
npm run proto:check
npm run build:openxml-wasm
npm run verify:openxml-wasm-build
npm run test:openxml-wasm-dotnet
npm test
npm run docs:api
npm run test:pack
dotnet test native/OfficeBridge # when dotnet is installed
npm run test:playwright-renderer # when Playwright/Chromium are installed
```

Use the consolidated checker:

```sh
npm run release:check
```

The checker runs npm tests/docs/pack, conditionally runs dotnet tests when `dotnet` is available, audits every locked dependency license against `scripts/license-policy.json` and `THIRD_PARTY_NOTICES.md`, checks npm auth, checks the published npm version, and reports blockers.

GitHub CI and the manual release workflow install Playwright Chromium, LibreOffice Writer/Calc/Impress, Poppler, and .NET 8 before running these gates. Their environment-probe step is required, so hosted runs exercise real browser and native render branches instead of silently accepting skips.

`test/package-contents.mjs` also validates the actual dry-run tarball manifest. It requires the project-local runnable skills, native bridge source, public protobuf source/generated binding, source-built XLSX/DOCX/PPTX OpenXML-WASM codecs/tests and same-host reproducibility verifier, runtime integrity manifest/SBOM/.NET notices, spreadsheet Pivot/date/group/filter/formula/aggregation/coercion/threaded-comment modules, PresentationML grouped-shape, Office 2021 modern-comment, and opaque native-object graph codecs/validators, the split OpenXML-WASM presentation/error adapters and C# PPTX text/color/bullet codecs, the DOCX bibliography, bookmark/internal-link, and section/header-footer planners, the PDF table-grid/accessibility normalizers, the structured-intersection/OpenXML-WASM fixtures, shared visual-baseline lifecycle, and third-party notices while rejecting reference/handoff material and local dotnet `bin/obj` build output. The current package remains below the reviewed 8 MB/16 MB gates; most of the footprint is the bundled .NET/Open XML SDK runtime required for no-local-dotnet Office I/O. The gate catches accidentally bundled fonts, reference assets, fixtures outside the published skills, symbols/maps, and build artifacts while forcing explicit review when legitimate published capability grows the package.

## Latest observed blockers in this environment

- `npm whoami` fails with `ENEEDAUTH`; this machine is not logged in to npm.
- `npm view open-office-artifact-tool version --json` returns `E404`; the package name/version was not visible on npm from this environment.
- `dotnet test native/OfficeBridge` passes locally with .NET 8 (`5` tests passed on macOS arm64); Windows + Microsoft Office automation remains an external integration gate.
- `dotnet test native/OpenXmlWasm/OpenXmlWasm.sln --configuration Release` passes locally (`17` tests), including PPTX character/auto/none marker authoring and edits, real relationship-backed picture-marker preservation, invalid marker rejection, basic rich-text roundtrip/editing, unmodeled paragraph/run-property retention, fixed-topology rejection, XLSX/DOCX/PPTX hash/source/opaque-graph or element-binding preservation, unsupported-advanced-edit rejection, and fail-closed Office 2021 validation. Hosted evidence for the list-marker increment is pending.
- Playwright renderer smoke tests pass locally with Playwright 1.61.1 and its Chromium runtime, covering PNG, WebP, JPEG, and PDF output.

## Publish command once auth is available

```sh
npm publish --access public
```

Do not publish if `npm run release:check` reports any required blocker, or if the current `package.json` version is already published.

## GitHub release workflow

A manual `.github/workflows/release.yml` workflow is available:

- `publish_npm=false` runs the release gates and `release-check` in dry-run/no-network mode.
- `publish_npm=true` requires `secrets.NPM_TOKEN`, runs `npm run release:check`, creates/pushes tag `v<package.version>`, creates a GitHub release if needed, and runs `npm publish --access public`.

Use `publish_npm=true` only after confirming the package version, changelog/release notes, and npm auth are correct.
