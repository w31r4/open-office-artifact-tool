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

`test/package-contents.mjs` validates the actual dry-run tarball manifest. It requires the project-local runnable skills, production native bridge/OpenChestnut source, public protobuf source/generated binding, same-host reproducibility verifier, runtime integrity manifest/SBOM/.NET notices, format codecs and validators, published fixtures, shared visual-baseline lifecycle, and third-party notices while rejecting reference/handoff material, development test source, and local dotnet `bin/obj` output. C# and JavaScript test sources remain in the public GitHub repository but are not npm runtime inputs. The current dry-run package contains 207 files, is 4,821,419 bytes compressed and 16,395,605 bytes unpacked. It remains below the reviewed 8,000,000/16,700,000-byte gates with 304,395 bytes of unpacked headroom after adding the color-filter/sort wire schema, C# codec, adapter, runtime, and runnable fixture; no gate increase hides the new capability. Most of the footprint remains the bundled .NET/Open XML SDK runtime required for no-local-dotnet Office I/O. The gate catches accidentally bundled fonts, reference assets, tests, symbols/maps, and build artifacts while forcing explicit review when legitimate published capability grows the package.

## Latest observed blockers in this environment

- `npm whoami` fails with `ENEEDAUTH`; this machine is not logged in to npm.
- `npm view open-office-artifact-tool version --json` returns `E404`; the package name/version was not visible on npm from this environment.
- `dotnet test native/OfficeBridge` passes locally with .NET 8 (`5` tests passed on macOS arm64); Windows + Microsoft Office automation remains an external integration gate.
- `dotnet test native/OpenChestnut/OpenChestnut.sln --configuration Release` passes locally with `97` tests. Twenty-four XLSX worksheet-table tests cover complete source-free authoring with Office 2021 validation, source-bound name/style, calculated-column/totals, exact/grouped-date/custom/dynamic/Top10/icon/color AutoFilter and ordered value/icon/color-sort edits without part/relationship identity churn, unchanged complex differential-style byte preservation with lossy-replacement rejection, and structured invalid-profile rejection. The color tests cover cell-fill and font targets, RGB and theme+tint values, differential-format deduplication, append-only source edits, invalid target/color combinations, and complex-record fallback. The preceding workbook-theme/static-style/native-formula and DOCX/PPTX corpus remains green. Local macOS produces 38 runtime files/13,220,028 bytes; the 207-file no-local-dotnet clean install creates and imports color filters/sorts alongside the earlier advanced profiles; and the OpenChestnut→OpenChestnut spreadsheet fixture carries all seven bounded filter profiles plus value/icon/color sort states with calculated-column/custom-totals metadata, table relationships, custom theme, static styles, shared/array topology, and number formats through modeled, package, Playwright, and real LibreOffice/Poppler QA. The last fully green hosted baseline before this milestone is run [`29365929087`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29365929087); current color-profile hosted evidence is recorded after the milestone commit passes.
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
