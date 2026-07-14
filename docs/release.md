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

`test/package-contents.mjs` also validates the actual dry-run tarball manifest. It requires the project-local runnable skills, native bridge source, public protobuf source/generated binding, source-built XLSX/DOCX/PPTX OpenChestnut codecs/tests and same-host reproducibility verifier, runtime integrity manifest/SBOM/.NET notices, spreadsheet Pivot/date/group/filter/formula/aggregation/coercion/threaded-comment modules, PresentationML grouped-shape, Office 2021 modern-comment, and opaque native-object graph codecs/validators, the split OpenChestnut presentation/error/asset adapters, C# DOCX merge-aware table/geometry and source-bound numbered-paragraph codecs, and C# PPTX asset-catalog/slide-context/text/paragraph-layout/paragraph-spacing/color/bullet/bullet-style/hyperlink codecs, the DOCX bibliography, bookmark/internal-link, and section/header-footer planners, the PDF table-grid/accessibility normalizers, the structured-intersection/OpenChestnut fixtures, shared visual-baseline lifecycle, and third-party notices while rejecting reference/handoff material and local dotnet `bin/obj` build output. The current dry-run package contains 203 files, is 4.7 MB compressed and 15.9 MB unpacked, and remains below the reviewed 8 MB/16 MB gates; most of the footprint is the bundled .NET/Open XML SDK runtime required for no-local-dotnet Office I/O. The gate catches accidentally bundled fonts, reference assets, fixtures outside the published skills, symbols/maps, and build artifacts while forcing explicit review when legitimate published capability grows the package.

## Latest observed blockers in this environment

- `npm whoami` fails with `ENEEDAUTH`; this machine is not logged in to npm.
- `npm view open-office-artifact-tool version --json` returns `E404`; the package name/version was not visible on npm from this environment.
- `dotnet test native/OfficeBridge` passes locally with .NET 8 (`5` tests passed on macOS arm64); Windows + Microsoft Office automation remains an external integration gate.
- `dotnet test native/OpenChestnut/OpenChestnut.sln --configuration Release` passes locally (`51` tests); hosted run [`29330878218`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29330878218) remains green for the preceding `49`-test milestone. Local coverage adds source-free horizontal/vertical merge construction, exact physical-grid/row-span/continuation validation, default `TableGrid` style materialization, invalid authored-graph rejection, and a direct-authoring Documents fixture to the existing all-family, source-bound merge, numbering-style, hyperlink/simple-field/table, PPTX owner-local, exact package-guard, and Office 2021 gates. The current macOS runtime has 38 files/12,931,772 bytes and the clean-install package has 205 files; refreshed hosted reproducibility, Chromium, LibreOffice, Poppler, and OfficeBridge evidence is pending this milestone's push.
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
