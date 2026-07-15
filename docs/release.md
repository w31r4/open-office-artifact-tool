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

`test/package-contents.mjs` validates the actual dry-run tarball manifest. It requires the project-local runnable skills, optional native Office bridge source, public protobuf source/generated binding, runtime integrity manifest/SBOM/.NET notices, consumer API docs, format codecs and validators, published fixtures, shared visual-baseline lifecycle, and third-party notices while rejecting reference/handoff material, development tests, local dotnet `bin/obj` output, repository-only planning notes, OpenChestnut C# source, and development build scripts. The GitHub source checkout remains the authoritative location for C# projects, locked dependencies, build tooling, tests, and same-host reproducibility verification; excluding their second copy from npm does not change any export or runtime lookup. The current chart-series-line consumer package contains 153 files, is 4,837,162 bytes compressed and 16,258,847 bytes unpacked. It remains below the unchanged 8,000,000/16,700,000-byte gates with 441,153 bytes of unpacked headroom, and the no-local-dotnet clean-install gate passes. The current exact-SHA hosted run is pending; the preceding chart-text-size run [`29406358049`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29406358049) passed its 152-file clean install, complete source-checkout reproducibility, real Chromium/LibreOffice 24.2.7.2/Poppler 24.02.0 tooling, generated protocol/API checks, OfficeBridge 5/5, OpenChestnut 132/132, and two byte-identical 38-file/13,573,822-byte Linux runtime builds. Most of the footprint remains the .NET/Open XML SDK runtime required for no-local-dotnet Office I/O.

## Latest observed blockers in this environment

- `npm whoami` fails with `ENEEDAUTH`; this machine is not logged in to npm.
- `npm view open-office-artifact-tool version --json` returns `E404`; the package name/version was not visible on npm from this environment.
- `dotnet test native/OfficeBridge` passes locally with .NET 8 (`5` tests passed on macOS arm64); Windows + Microsoft Office automation remains an external integration gate.
- `dotnet test native/OpenChestnut/OpenChestnut.sln --configuration Release` passes locally with `133` tests. The drawing corpus authors/imports bounded bar, line, and pie charts with title, legend, literal or formula-backed categories, numeric caches, optional per-series direct RGB solid fills, direct RGB lines with five preset dash styles and 0–1,584 point width, bounded chart-title and axis tick-label font sizes, one-cell/two-cell/absolute anchors, and primary category/value axes; preserves residual ChartSpace and unrelated series/text XML; permits fixed-topology source-bound name/title/legend/formula/cache/fill/line/font-size/axis edits; imports formula-referenced series names, unrecognized fill/line graphs, rich/complex text graphs, and logarithmic axes read-only; and rejects chart add/remove, anchor/type/series/point/axis topology mutation, binding tamper, external data, combo/multi-plot profiles, invalid ranges, pie axes, marker/axis-title styling, and unsupported axis profiles. Pictures retain their preceding bounded crop/effects/transform and same-format byte-replacement coverage, including simultaneous picture/chart edits against one pre-mutation DrawingPart hash. Local macOS produces 38 inventoried runtime files/13,584,060 bytes, and two isolated builds are byte-identical across all 39 audited files. `open-chestnut-basic` carries a formula-backed line chart with series-fill/line, title/tick-label font-size, and primary-axis edits, all three picture anchor kinds, source-bound chart and image edits, its `SEQUENCE` spill, and two workbook windows through two WASM exports plus required real Chromium 149.0.7827.55/LibreOfficeDev 26.8/Poppler 26.05 QA; its four native pages pass page-count and pixel gates. JavaScript fallback and the passing no-local-dotnet clean-install gate use the same public model. The current 153-file package is 4,837,162 bytes compressed and 16,258,847 bytes unpacked. Exact-SHA hosted evidence for this chart-series-line slice is pending; the preceding chart-text-size run [`29406358049`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29406358049) remains the latest green hosted baseline.
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
