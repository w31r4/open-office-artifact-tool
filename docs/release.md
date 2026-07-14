# Release readiness

Status date: 2026-07-15

## Current package

- Package: `open-office-artifact-tool`
- Version: `0.1.0`
- Publish state: publish-ready, not published from this environment

## Required local gates

The current release gate is:

```sh
npm run proto:check
npm run build:open-chestnut
npm run verify:open-chestnut-build
npm run test:open-chestnut-dotnet
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

The package publishes `codecs/open-chestnut` as the canonical Office codec. `codecs/openxml-wasm` and the old npm script names remain deprecated compatibility aliases and are covered by source-tree and clean-install tests; all runtime assets, C# projects, CI commands, fixtures, and new documentation use OpenChestnut.

`test/package-contents.mjs` also validates the actual dry-run tarball manifest. It requires the project-local runnable skills, native bridge source, public protobuf source/generated binding, source-built XLSX/DOCX/PPTX OpenChestnut codecs/tests and same-host reproducibility verifier, runtime integrity manifest/SBOM/.NET notices, spreadsheet Pivot/date/group/filter/formula/aggregation/coercion/threaded-comment modules, PresentationML grouped-shape, Office 2021 modern-comment, and opaque native-object graph codecs/validators, the split OpenChestnut presentation/error/asset adapters, C# XLSX workbook-theme/static-style/number-format/native-formula/worksheet-table rich-column/filter/sort codecs, C# DOCX direct-numbering, merge-aware table/geometry/formatting, numbered-paragraph, and numbering-edit planner codecs, and C# PPTX asset-catalog/slide-context/text/paragraph-layout/paragraph-spacing/color/bullet/bullet-style/hyperlink codecs, the DOCX bibliography, bookmark/internal-link, and section/header-footer planners, the PDF table-grid/accessibility normalizers, the structured-intersection/OpenChestnut fixtures, shared visual-baseline lifecycle, and third-party notices while rejecting reference/handoff material and local dotnet `bin/obj` build output. The current dry-run package contains 214 files, is about 4.85 MB compressed/16.69 MB unpacked, and remains below the reviewed 8,000,000/16,700,000-byte gates with roughly 10 KB of unpacked headroom. The additive icon-filter/sort wire schema, C# source/tests, adapter, and runnable fixture remain inside the previously reviewed ceiling; no unrelated asset growth is hidden by a gate increase. Most of the footprint remains the bundled .NET/Open XML SDK runtime required for no-local-dotnet Office I/O. The gate catches accidentally bundled fonts, reference assets, fixtures outside the published skills, symbols/maps, and build artifacts while forcing explicit review when legitimate published capability grows the package.

## Latest observed blockers in this environment

- `npm whoami` fails with `ENEEDAUTH`; this machine is not logged in to npm.
- `npm view open-office-artifact-tool version --json` returns `E404`; the package name/version was not visible on npm from this environment.
- `dotnet test native/OfficeBridge` passes locally with .NET 8 (`5` tests passed on macOS arm64); Windows + Microsoft Office automation remains an external integration gate.
- `dotnet test native/OpenChestnut/OpenChestnut.sln --configuration Release` passes locally and in hosted Linux CI with `94` tests. Twenty-one worksheet-table tests cover complete source-free authoring with Office 2021 validation, source-bound name/style, calculated-column/totals, exact/grouped-date/custom/dynamic/Top10/icon AutoFilter and ordered value/icon-sort edits without part/relationship identity churn, unchanged complex/color-filter/color-sort byte preservation with lossy-replacement rejection, and structured invalid-profile rejection. The icon tests cover present-zero and absent/no-icon IDs, per-set cardinality, invalid vocabulary, and semantic edits. The preceding workbook-theme/static-style/native-formula and DOCX/PPTX corpus remains green. Local macOS produces 38 runtime files/13,210,300 bytes; the 214-file no-local-dotnet clean install creates and imports icon filters/sorts alongside the earlier advanced profiles; and the OpenChestnut→OpenChestnut spreadsheet fixture carries all six bounded filter profiles plus value/icon sort states with calculated-column/custom-totals metadata, table relationships, custom theme, static styles, shared/array topology, and number formats through modeled, package, Playwright, and real LibreOffice/Poppler QA. Hosted run [`29365621713`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29365621713) passes the complete gate with OpenChestnut 94/94, OfficeBridge 5/5, the 214-file package, real Chromium/LibreOffice 24.2.7.2/Poppler 24.02.0, and two byte-identical 38-file/13,210,302-byte Linux runtime builds.
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
