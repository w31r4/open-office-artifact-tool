# Release

## 0.2.0 boundary

0.2.0 is a breaking convergence release:

- OpenChestnut is the only DOCX/XLSX/PPTX codec.
- PDF remains an independent fourth pipeline.
- Wire protocol is version 2.
- `allow_lossy` is removed and reserved in the proto.
- `OFFICE_CODEC_IDS`, `office-codec-policy.mjs`, and the `codecs/openxml-wasm` export are removed.
- `codec`, `allowLossy`, `preferNative`, and `relativeDateAsOf` facade options are rejected.
- Old JavaScript Office parsers/writers and dedicated dead helper modules are not packaged.
- Imported advanced Office content is preserved only with validated source evidence; unsupported edits and opaque content without that evidence fail closed.

There is no compatibility window or fallback mode.

## Source and npm distributions

The repository is the authoritative source distribution. It contains OpenChestnut C# source, locked dependencies, protocol definitions, build scripts, tests, Skills, and reproducibility gates.

The npm tarball is the consumer distribution. It contains the JavaScript object models, OpenChestnut adapter, generated wire binding, public proto, bundled runtime, integrity manifest, SBOM, license notices, render/QA helpers, PDF pipeline, and four native plugin bundles containing five Skills. It excludes C# source, test sources, build output, repository-only build scripts, and the development-only `test/skill-harness` fixtures.

Installed consumers do not need `dotnet` on `PATH`.

## Required release gates

Run from a clean source checkout with the documented Node and .NET SDK versions:

```sh
npm ci
npm run proto:generate
npm run test:open-chestnut-dotnet
npm run build:open-chestnut
npm test
npm run test:pack
npm run verify:open-chestnut-build
npm run docs:api
npm run release:check
```

The release candidate is acceptable only when all of the following are true:

- the generated JavaScript wire matches protocol 2 and proto lint passes;
- C# DOCX, XLSX, PPTX, package-boundary, and failure-profile tests pass;
- default facade create/import/edit/re-export roundtrips pass for all three Office formats;
- legacy options, old subpath, missing runtime, and opaque-without-source cases fail explicitly;
- all four native plugin manifests validate, the published five-Skill topology is complete, and every workflow promoted to compatible in `docs/reference-skills.md` passes from the public package surface;
- PDF create/import, inspect, semantic verify, and render smoke passes independently;
- Open XML SDK validation passes for generated Office fixtures;
- configured LibreOffice/Poppler/Playwright/native render gates pass where available;
- a packed clean install completes all three Office roundtrips and PDF smoke while `dotnet` is absent from `PATH`;
- two clean OpenChestnut builds produce the same runtime file set and hashes;
- package contents contain no legacy Office codec files, C# build output, tests, or repository-only scripts;
- package metadata, version `0.2.0`, licenses, third-party notices, SBOM, and integrity manifest agree;
- hosted Linux runs the same required non-optional gates.

## Optional native validation

LibreOffice, Poppler, Playwright, and the Windows Office bridge are validation/render tools. Absence of an optional host may skip only the explicitly environment-gated native rendering branch; it must not skip codec, semantic, package, Skill, PDF, or clean-install gates.

The Office bridge does not participate in normal import/export and must never be used to hide a codec failure.

## Current local evidence

On 2026-07-16, the native reference-plugin/OpenChestnut compatibility worktree passed the complete local gate on macOS arm64 with Node 26.5.0, npm 11.17.0, and .NET SDK 8.0.128:

- `npm test` passed the OpenChestnut protocol/facade tests, explicit OOXML inspect/patch tests, four native plugin bundles/five published Skills, the runnable Documents create/import/edit/export vertical slice, the 26-slide reference Presentation vertical slice, PDF, render/visual QA, Playwright, examples, release metadata, package contents, and Help catalog.
- OpenChestnut passed `163/163` C# tests, including literal DrawingML custom-geometry author/import/edit/failure coverage; the optional OfficeBridge passed `5/5` protocol tests.
- Buf lint passed, protobuf generation was byte-idempotent, and `npm run docs:api` regenerated the public API reference.
- `npm run test:pack` passed the no-local-dotnet clean-install probe for DOCX/XLSX/PPTX and the independent PDF path.
- The dry-run npm tarball contains 359 files, is 9,011,581 bytes compressed and 22,152,335 bytes unpacked. The additional file is the runnable Documents OpenChestnut example; development harnesses, private reference material, and removed Office codec paths remain excluded.
- `npm run verify:open-chestnut-build` compared 39 audited files; both clean builds produced the same 38-file, 14,018,748-byte runtime payload. The build entry point now clears the Release incremental graph before restore/publish so a preceding `dotnet test` cannot make the first WASM publish differ from the second.
- Render-backed gates ran with LibreOfficeDev 26.8.0.0.alpha0, Poppler 26.05.0, and the installed Playwright Chromium runtime.
- The shipped Documents example and an independent Skill forward test each completed two OpenChestnut round trips, semantic assertions, and a one-page LibreOffice render with every page inspected. The audit narrowed custom source-free table styles to `TableGrid` plus direct formatting, documented non-persistent model locators across imports, aligned header/footer distance to the codec's 720-twip default, and routed ordinary classic comments through `DocumentModel.addComment`.
- LibreOffice opened the shipped 26-slide reference template and produced a 26-page PDF; bounded custom-geometry icons rendered visibly. This local LibreOffice build substituted `Helvetica Neue`, so pixel parity with the checked-in preview images is not claimed and remains a visual-fidelity gap.

## Hosted evidence

The code-bearing native reference-plugin/OpenChestnut compatibility commit `c526ce1534552c4e2deaf732513ddb270b873aab` passed the hosted Linux `ci` workflow in [GitHub Actions run 29480511630](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29480511630) on 2026-07-16. The run completed with conclusion `success` and covered protocol/runtime reproducibility, Chromium/native render tools, `npm test`, generated API-doc diff, offline release metadata, the clean-install tarball, OfficeBridge, and OpenChestnut.

`npm run release:check` passes the source, documentation, package, license, JavaScript, and .NET gates. Its only remaining blocker is unavailable npm authentication. No `npm publish` or tag/release operation has been performed.

## Publishing

Before publishing:

1. Verify `package.json` and `package-lock.json` both declare `0.2.0`.
2. Rebuild and verify the OpenChestnut runtime from source.
3. Regenerate API documentation after the final public API change.
4. Inspect `npm pack --dry-run --json` for the required runtime/proto/Skill files and forbidden legacy files.
5. Run the tarball clean-install probe, not only source-tree tests.
6. Record the exact commit and hosted gate result used for publication.
