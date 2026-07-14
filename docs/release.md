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

`test/package-contents.mjs` validates the actual dry-run tarball manifest. It requires the project-local runnable skills, production native bridge/OpenChestnut source, public protobuf source/generated binding, same-host reproducibility verifier, runtime integrity manifest/SBOM/.NET notices, format codecs and validators, published fixtures, shared visual-baseline lifecycle, and third-party notices while rejecting reference/handoff material, development test source, and local dotnet `bin/obj` output. C# and JavaScript test sources remain in the public GitHub repository but are not npm runtime inputs. The current dry-run package contains 209 files, is 4,840,227 bytes compressed and 16,502,256 bytes unpacked. It remains below the reviewed 8,000,000/16,700,000-byte gates with 197,744 bytes of unpacked headroom after adding the QueryTable wire schema, dedicated C# deep module, adapter/fallback boundary, rebuilt runtime, and runnable fixture; no gate increase hides the new capability. Most of the footprint remains the bundled .NET/Open XML SDK runtime required for no-local-dotnet Office I/O. The gate catches accidentally bundled fonts, reference assets, tests, symbols/maps, and build artifacts while forcing explicit review when legitimate published capability grows the package.

## Latest observed blockers in this environment

- `npm whoami` fails with `ENEEDAUTH`; this machine is not logged in to npm.
- `npm view open-office-artifact-tool version --json` returns `E404`; the package name/version was not visible on npm from this environment.
- `dotnet test native/OfficeBridge` passes locally with .NET 8 (`5` tests passed on macOS arm64); Windows + Microsoft Office automation remains an external integration gate.
- `dotnet test native/OpenChestnut/OpenChestnut.sln --configuration Release` passes locally with `100` tests. The new QueryTable tests cover Office 2021-valid import, source-bound root-policy editing, table/query relationship and ConnectionsPart preservation, source-free fabrication rejection, missing-connection and binding-tamper rejection, plus byte-exact hidden fallback for unsupported companion relationships. The preceding worksheet table, theme/static-style/native-formula, DOCX, and PPTX corpus remains green. Local macOS produces 38 runtime files/13,278,396 bytes; two isolated builds are byte-identical across 39 audited files. The 209-file no-local-dotnet clean install imports and edits the QueryTable slice alongside earlier profiles. `open-chestnut-query-table` crosses a synthetic public-standard source graph through OpenChestnut import, safe refresh disablement/root edit, two source-preserving exports, package inspection, and Playwright QA. Native office rendering is intentionally disabled for that fixture to avoid allowing a renderer to execute its external connection. Hosted QueryTable evidence is pending; the preceding hosted run [`29367655620`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29367655620) remains green with OpenChestnut 97/97, OfficeBridge 5/5, real Chromium/LibreOffice 24.2.7.2/Poppler 24.02.0, and reproducible Linux runtime builds.
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
