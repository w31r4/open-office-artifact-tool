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

`test/package-contents.mjs` validates the actual dry-run tarball manifest. It requires the project-local runnable skills, optional native Office bridge source, public protobuf source/generated binding, runtime integrity manifest/SBOM/.NET notices, consumer API docs, format codecs and validators, published fixtures, shared visual-baseline lifecycle, and third-party notices while rejecting reference/handoff material, development tests, local dotnet `bin/obj` output, repository-only planning notes, OpenChestnut C# source, and development build scripts. The GitHub source checkout remains the authoritative location for C# projects, locked dependencies, build tooling, tests, and same-host reproducibility verification; excluding their second copy from npm does not change any export or runtime lookup. The current dry-run consumer package contains 152 files, is 4,813,009 bytes compressed and 16,177,070 bytes unpacked. It remains below the unchanged 8,000,000/16,700,000-byte gates with 522,930 bytes of unpacked headroom, and the no-local-dotnet clean-install gate passes. Hosted run [`29396506905`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29396506905) passes the preceding 151-file package alongside complete source-checkout reproducibility, real render tooling, and codec gates; exact-SHA hosted evidence for the 152-file chart package is pending. Most of the footprint remains the .NET/Open XML SDK runtime required for no-local-dotnet Office I/O.

## Latest observed blockers in this environment

- `npm whoami` fails with `ENEEDAUTH`; this machine is not logged in to npm.
- `npm view open-office-artifact-tool version --json` returns `E404`; the package name/version was not visible on npm from this environment.
- `dotnet test native/OfficeBridge` passes locally with .NET 8 (`5` tests passed on macOS arm64); Windows + Microsoft Office automation remains an external integration gate.
- `dotnet test native/OpenChestnut/OpenChestnut.sln --configuration Release` passes locally with `130` tests. The drawing corpus authors/imports bounded bar, line, and pie charts with title, legend, literal or formula-backed categories, numeric caches, and one-cell/two-cell/absolute anchors; preserves residual ChartSpace XML; permits only fixed-topology source-bound name/title/legend/formula/cache edits; imports formula-referenced series names read-only; and rejects chart add/remove, anchor/type/series/point topology mutation, binding tamper, external data, combo/multi-plot profiles, and unsupported styling. Pictures retain their preceding bounded crop/effects/transform and same-format byte-replacement coverage, including simultaneous picture/chart edits against one pre-mutation DrawingPart hash. Local macOS produces 38 inventoried runtime files/13,547,708 bytes, and two isolated builds are byte-identical across all 39 audited files. `open-chestnut-basic` carries a formula-backed line chart, all three picture anchor kinds, source-bound chart and image edits, its `SEQUENCE` spill, and two workbook windows through two WASM exports plus real Chromium/LibreOffice/Poppler QA; JavaScript fallback and no-local-dotnet clean-install probes use the same public model. The current 152-file package is 4,813,009 bytes compressed and 16,177,070 bytes unpacked. Hosted Linux run [`29398115651`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29398115651) remains the latest evidence for the preceding picture-transform increment; exact-SHA chart evidence is pending.
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
