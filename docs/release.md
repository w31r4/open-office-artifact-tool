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

`test/package-contents.mjs` validates the actual dry-run tarball manifest. It requires the project-local runnable skills, production native bridge/OpenChestnut source, public protobuf source/generated binding, same-host reproducibility verifier, runtime integrity manifest/SBOM/.NET notices, consumer API docs, format codecs and validators, published fixtures, shared visual-baseline lifecycle, and third-party notices while rejecting reference/handoff material, development test source, local dotnet `bin/obj` output, and repository-only coverage/release/architecture/subsystem notes. The shipped skills link to the GitHub coverage matrix instead of assuming that implementation-planning history is installed beside runtime files. C# and JavaScript test sources plus decision logs remain public on GitHub but are not npm runtime inputs. The current dry-run package contains 212 files, is 4,878,272 bytes compressed and 16,687,003 bytes unpacked. It remains below the unchanged 8,000,000/16,700,000-byte gates with 12,997 bytes of unpacked headroom after adding selected-tab wire/API/C# coverage and rebuilding the bundled runtime. Most of the footprint remains the .NET/Open XML SDK runtime required for no-local-dotnet Office I/O.

## Latest observed blockers in this environment

- `npm whoami` fails with `ENEEDAUTH`; this machine is not logged in to npm.
- `npm view open-office-artifact-tool version --json` returns `E404`; the package name/version was not visible on npm from this environment.
- `dotnet test native/OfficeBridge` passes locally with .NET 8 (`5` tests passed on macOS arm64); Windows + Microsoft Office automation remains an external integration gate.
- `dotnet test native/OpenChestnut/OpenChestnut.sln --configuration Release` passes locally with `118` tests. The workbook-view corpus now covers coordinated source-free `ActiveTab`/`tabSelected` authoring, grouped selection import/edit, stable worksheet IDs, per-worksheet part/view bindings, explicit-false presence, residual pane/zoom preservation, incomplete selection graphs, nested-binding tamper, legacy active-only edits, opaque multi-window safety, and hidden active/selected rejection; preceding worksheet/calculation/name/sort/connection/table/theme/style/formula/DOCX/PPTX coverage remains green. Local macOS produces 38 inventoried runtime files/13,415,612 bytes. `open-chestnut-basic` crosses active `Icon Rules` plus the selected `Summary`/`Icon Rules` group through two WASM passes and real render QA; fallback and no-local-dotnet clean-install probes use the same public APIs. Hosted run [`29383428603`](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29383428603) remains the latest complete predecessor evidence at OpenChestnut 116/116; hosted evidence for this 118-test increment is pending the pushed commit.
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
