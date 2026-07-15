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

`test/package-contents.mjs` validates the actual dry-run tarball manifest. It requires the project-local runnable skills, optional native Office bridge source, public protobuf source/generated binding, runtime integrity manifest/SBOM/.NET notices, consumer API docs, format codecs and validators, published fixtures, shared visual-baseline lifecycle, and third-party notices while rejecting reference/handoff material, development tests, local dotnet `bin/obj` output, repository-only planning notes, OpenChestnut C# source, and development build scripts. The GitHub source checkout remains the authoritative location for C# projects, locked dependencies, build tooling, tests, and same-host reproducibility verification; excluding their second copy from npm does not change any export or runtime lookup. The current chart-series-line consumer package contains 153 files, is 4,837,162 bytes compressed and 16,258,847 bytes unpacked. It remains below the unchanged 8,000,000/16,700,000-byte gates with 441,153 bytes of unpacked headroom, and the no-local-dotnet clean-install gate passes. Exact-SHA hosted run [`29408434715`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29408434715) verifies the same 153-file clean install, complete source-checkout reproducibility, real Chromium/LibreOffice 24.2.7.2/Poppler 24.02.0 tooling, generated protocol/API checks, OfficeBridge 5/5, OpenChestnut 133/133, and two byte-identical 38-file/13,584,062-byte Linux runtime builds. Most of the footprint remains the .NET/Open XML SDK runtime required for no-local-dotnet Office I/O.

The current chart-series-marker increment produces a 154-file consumer package of 4,845,101 bytes compressed and 16,282,229 bytes unpacked, leaving 417,771 bytes below the unchanged unpacked-size gate. The no-local-dotnet clean-install probe passes. Its source-built macOS runtime contains 38 files/13,593,788 bytes, and two isolated builds are byte-identical across all 39 audited files. OpenChestnut Release tests pass 134/134 and OfficeBridge passes 5/5. The required `open-chestnut-basic` gate crosses diamond/8 authoring and triangle/10 source-bound editing through two WASM passes; all five model sheets pass Chromium 149.0.7827.55, and all four native pages pass LibreOfficeDev 26.8.0.0.alpha0 plus Poppler 26.05.0. Exact-SHA hosted Linux run [`29410193061`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29410193061) passes the same 154-file clean install, all npm/skill/render/docs/package gates, real LibreOffice 24.2.7.2/Poppler 24.02.0, OpenChestnut 134/134, OfficeBridge 5/5, and two byte-identical 38-file/13,593,790-byte runtime builds.

The current chart-level line-smoothing increment produces a 155-file consumer package of 4,847,772 bytes compressed and 16,292,979 bytes unpacked, leaving 407,021 bytes below the unchanged unpacked-size gate. The no-local-dotnet clean-install probe passes. Its source-built macOS runtime contains 38 files/13,597,884 bytes, and two isolated builds are byte-identical across all 39 audited files. OpenChestnut Release tests pass 135/135 and OfficeBridge passes 5/5. The required `open-chestnut-basic` gate crosses chart-level smooth `true` authoring and explicit `false` source-bound editing through two WASM passes; all five model sheets pass Chromium 149.0.7827.55, and all four native pages pass LibreOfficeDev 26.8.0.0.alpha0 plus Poppler 26.05.0. Exact-SHA hosted Linux run [`29411788218`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29411788218) passes the same full npm/skill/render/docs/package gate with real LibreOffice 24.2.7.2/Poppler 24.02.0, OpenChestnut 135/135, OfficeBridge 5/5, the 155-file clean install, and two byte-identical 38-file/13,597,886-byte runtime builds.

The current chart-level line-grouping increment keeps the consumer package at 155 files, grows it to 4,849,498 bytes compressed and 16,301,411 bytes unpacked, and leaves 398,589 bytes below the unchanged unpacked-size gate. The no-local-dotnet clean-install probe passes. Its source-built macOS runtime contains 38 files/13,600,956 bytes, and two isolated builds are byte-identical across all 39 audited files. OpenChestnut Release tests pass 135/135 and OfficeBridge passes 5/5. The required `open-chestnut-basic` gate crosses `stacked` authoring to `percentStacked` source-bound editing alongside smooth true→false through two WASM passes; all five model sheets pass Chromium 149.0.7827.55, and all four native pages pass LibreOfficeDev 26.8.0.0.alpha0 plus Poppler 26.05.0. Exact-SHA hosted Linux run [`29413451804`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29413451804) passes the same full npm/skill/render/docs/package gate with real LibreOffice 24.2.7.2/Poppler 24.02.0, the 155-file clean install, OpenChestnut 135/135, OfficeBridge 5/5, and two byte-identical 38-file/13,600,958-byte runtime builds.

## Latest observed blockers in this environment

- `npm whoami` fails with `ENEEDAUTH`; this machine is not logged in to npm.
- `npm view open-office-artifact-tool version --json` returns `E404`; the package name/version was not visible on npm from this environment.
- `dotnet test native/OfficeBridge` passes locally with .NET 8 (`5` tests passed on macOS arm64); Windows + Microsoft Office automation remains an external integration gate.
- `dotnet test native/OpenChestnut/OpenChestnut.sln --configuration Release` passes locally and on exact-SHA hosted Linux with `135` tests. The current drawing corpus covers bounded bar/line/pie charts, direct RGB fill/line, line markers, chart-level standard/stacked/percent-stacked grouping, smooth true/false/absence, title/tick font sizes, primary axes, all three anchor kinds, source-bound add/change/remove, complex-graph exact preservation, and Office 2021 zero-error validation. Local macOS produces 38 runtime files/13,600,956 bytes, and hosted Linux produces 38 files/13,600,958 bytes, each with two byte-identical 39-file build audits. The 155-file clean-install package and required Chromium/LibreOffice/Poppler `open-chestnut-basic` gate pass in both environments.
- Playwright renderer smoke tests pass locally with Playwright 1.61.1 and its Chromium runtime, covering PNG, WebP, JPEG, and PDF output.
- The marker increment's complete local and exact-SHA hosted gates pass: OpenChestnut 134/134, OfficeBridge 5/5, all npm/runnable-skill tests, generated API/protobuf checks, the 154-file clean install, two byte-identical runtime builds, and required real browser/native rendering. Hosted evidence is run [`29410193061`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29410193061).
- The line-smoothing increment's complete local and exact-SHA hosted gates pass: OpenChestnut 135/135, OfficeBridge 5/5, all npm/runnable-skill tests, generated API/protobuf checks, the 155-file clean install, two byte-identical runtime builds, and required real browser/native rendering. Hosted evidence is run [`29411788218`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29411788218).
- The line-grouping increment's complete local and exact-SHA hosted gates pass: OpenChestnut 135/135, OfficeBridge 5/5, all npm/runnable-skill tests, generated API/protobuf checks, the 155-file clean install, two byte-identical runtime builds, and required real browser/native rendering. Hosted evidence is run [`29413451804`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29413451804).

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
