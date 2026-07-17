# Release

## 0.2.0 boundary

0.2.0 is a breaking convergence release:

- OpenChestnut is the only DOCX/XLSX/PPTX codec.
- PDF remains an independent fourth pipeline and never enters OpenChestnut. Required `mupdf@1.28.0` is the runtime-lazy default for arbitrary-file parse, native inspect/render, and bounded direct-original edits; specialist Python/system providers remain explicit task routes.
- The project is licensed under GNU AGPL-3.0-or-later. Normal npm installation resolves MuPDF.js as a direct dependency; there is no PDF postinstall hook or standalone dependency downloader.
- Wire protocol is version 2.
- `allow_lossy` is removed and reserved in the proto.
- `OFFICE_CODEC_IDS`, `office-codec-policy.mjs`, and the `codecs/openxml-wasm` export are removed.
- `codec`, `allowLossy`, `preferNative`, and `relativeDateAsOf` facade options are rejected.
- Old JavaScript Office parsers/writers and dedicated dead helper modules are not packaged.
- Imported advanced Office content is preserved only with validated source evidence; unsupported edits and opaque content without that evidence fail closed.

There is no compatibility window or fallback mode.

## Source and npm distributions

The repository is the authoritative source distribution. It contains OpenChestnut C# source, locked dependencies, protocol definitions, build scripts, tests, Skills, and reproducibility gates.

The npm tarball is the consumer distribution. It contains the JavaScript object models, OpenChestnut adapter, generated wire binding, public proto, bundled runtime, integrity manifest, SBOM, license notices, render/QA helpers, PDF pipeline, and four native plugin bundles containing five Skills. It excludes C# source, test sources, build output, repository-only build scripts, and the development-only `test/skill-harness` fixtures. MuPDF.js is declared in the required npm dependency graph rather than copied into this project's own tarball, and its WASM runtime is initialized only by a PDF operation.

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
- PDF greenfield authoring plus default MuPDF.js import/inspect/render/bounded-edit, lazy-load, pre-WASM budget, exact-prefix incremental-save, signature/redaction/deletion fail-closed, Skill CLI source-protection, and specialist-provider contract tests pass independently;
- when explicitly configured, the real optional-provider test covers ReportLab creation, pdfplumber extraction, type-aware pypdf text/radio/checkbox forms and annotations, typed pypdf merge/reorder/selective watermarking, PyMuPDF rewrite/incremental/page/text/image/form/annotation edits, real redaction/scrub/residue scans, capped numerical text-fit behavior, canonical audit byte binding, and typed Poppler source/output comparison;
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

### AGPL and default MuPDF.js vertical slice

On 2026-07-17, the focused PDF vertical slice passed `node test/pdf.mjs`, `node test/pdf-provider-skill.mjs`, `node test/reference-skills.mjs`, `node test/release-check.mjs`, and the package-contents gate. The checks cover root-import laziness, first-use MuPDF initialization, arbitrary-PDF import/inspect, native PNG/JPEG render, bounded direct-original editing, input/render/image/object limits, exact-prefix incremental saves, rewrite redaction, signed/redaction/deletion incremental fail-closed behavior, real link and raster extraction, CLI atomic output, nested output creation, and direct/symlink source-overwrite rejection.

This focused record intentionally does not assign the combined worktree's final package file count, tarball size, complete Office/PPTX gate, or hosted CI result to the MuPDF-only change. Those integration measurements belong to the subsequent combined release-evidence update.

On 2026-07-17, the native reference-plugin/OpenChestnut compatibility worktree passed the complete local gate on macOS arm64 with Node 26.5.0, npm 11.17.0, and .NET SDK 8.0.128:

- `npm test` passed the OpenChestnut protocol/facade tests, explicit OOXML inspect/patch tests, four native plugin bundles/five published Skills, the Spreadsheet Range/R1C1/direct-series/standard-sparkline compatibility suites, the runnable Documents create/import/edit/export vertical slice, the 26-slide reference Presentation vertical slice plus recursive-group and embedded-XLSX/OLE fixtures, PDF greenfield and provider-contract suites, render/visual QA, Playwright, examples, release metadata, package contents, and Help catalog.
- OpenChestnut passed `170/170` C# tests, including standard Office 2010 sparkline coverage, literal DrawingML custom-geometry coverage, direct PPTX slide-background author/import/hash-bound add-edit-remove/advanced-source fail-closed coverage, embedded-picture signed `a:srcRect` author/import/add-edit-remove/irregular-source fail-closed coverage, plain-text speaker-notes author/import/hash-bound edit/rich-source fail-closed coverage, recursive native `p:grpSp` author/import/fixed-topology mixed-descendant edit/complex-group opaque coverage, and source/hash/content-type/relationship-bound embedded XLSX payload replacement with Open XML SDK plus Office 2021 validation and unrelated-part preservation; the optional OfficeBridge passed `5/5` protocol tests.
- Buf lint passed, protobuf generation was byte-idempotent, and `npm run docs:api` regenerated the public API reference.
- `npm run test:pack` passed the no-local-dotnet clean-install probe for DOCX/XLSX/PPTX and the independent PDF path.
- The final combined dry-run npm tarball contains 417 files, is 9,190,240 bytes compressed and 22,803,636 bytes unpacked. It includes the rich PDF Skill tasks/references/examples, required lazy MuPDF.js route, six-page tagged-accessibility report workflow, explicit Python-runtime selector, path-safe attachment quarantine, type-aware AcroForm and merge/reorder/selective-watermark adapters, active-content inert scanner, canonical audit validator/schema, typed Poppler source/output comparator, the dependency-leaf cross-format raster registration/diff/visual-QA engine, the shipped Spreadsheet Range and sparkline workflows, the PPTX direct-background, speaker-notes, signed image-crop, recursive native-group, and embedded-XLSX/OLE workflow references, and the dependency-leaf OOXML package, Range, chart-source, sparkline, and Presentation group/image-fit/crop/native-object modules while excluding Python bytecode/cache files, development harnesses, private reference material, Agent PromptBench oracles, and removed Office codec paths.
- `npm run verify:open-chestnut-build` compared 39 audited files; both clean builds produced the same manifest-bound 38-file, 14,153,404-byte runtime payload. The build entry point clears the Release incremental graph before restore/publish so a preceding `dotnet test` cannot make the first WASM publish differ from the second.
- Render-backed gates ran with LibreOfficeDev 26.8.0.0.alpha0, Poppler 26.05.0, and the installed Playwright Chromium runtime.
- The explicit real-provider gate ran with ReportLab 4.4.9, pdfplumber 0.11.9, pypdf 6.10.0, and separately installed PyMuPDF 1.27.2.3 under the approved AGPL route. It proved byte-identical incremental prefixes, non-prefix rewrites, pypdf text/radio/checkbox value and appearance handling, complete-source page selection/reorder/selective watermarking with preserved navigation, pypdf and PyMuPDF annotation operations, bounded text/image/page edits, full redaction/scrub, single-revision/no-residue output, and deliberate reflow rejection. The typed Poppler comparator maps each merged output page back to its immutable source, requires exact pixels on unstamped pages, requires bounded change on stamped pages, and reports blank-state and dark-ratio evidence. The scrub-only active-content fixture proves removal of root/additional JavaScript, Launch/SubmitForm actions, attachments, invisible text, comments, populated form values, personal metadata, and the null active-content dictionary names that PyMuPDF can leave after logical deletion; unfamiliar object serialization and invisible text overlapping visible text fail closed. An image-bearing strict residue scan failed closed because Tesseract is not installed; qpdf, pyHanko, veraPDF, pikepdf, and OCRmyPDF were not executed and remain external/planned as documented.
- The shipped Documents example and an independent Skill forward test each completed two OpenChestnut round trips, semantic assertions, and a one-page LibreOffice render with every page inspected. The audit narrowed custom source-free table styles to `TableGrid` plus direct formatting, documented non-persistent model locators across imports, aligned header/footer distance to the codec's 720-twip default, and routed ordinary classic comments through `DocumentModel.addComment`.
- The shipped Spreadsheet Range example completed R1C1/block-write/formula-evidence/navigation/format/chart/verify/render and two OpenChestnut round trips. A separate forward test authored and visually reviewed a three-sheet operating forecast with formula-driven financials, zero spreadsheet errors, PASS model checks, and a formula-bound line chart. Its findings led to automatic native formulas for `containsText`, text-preserving direct references, and live internal-range cache resolution for formula-only chart inspect/render/export; imported chart persistence snapshots remain deliberately separate and fail closed.
- The shipped Spreadsheet sparkline example and development fixture authored standard Office 2010 `x14:sparklineGroups` line, column, and stacked profiles, exercised vertical row and horizontal column mappings, imported and edited recognized groups with fixed topology, preserved unsupported non-contiguous native groups unchanged, and rejected lossy topology changes. LibreOfficeDev opened the fixture and Poppler rendered all three types across two pages; the JavaScript SVG preview rendered each target cell independently.
- The repository-only Agent PromptBench validates 26 black-box cases (19 PDF and 7 Office; 7 ready and 19 asset-required). All seven deterministic ready cases materialized isolated candidate trials with an npm-packed install and read-only source hashes. A candidate/reference PDF preparation used byte-identical package tarballs and prompts while varying only the copied Skill. Generic grading covers fail-closed branching, immutable file/directory inputs, read-only prompt/Skill/dependency trees, regular output types, and exact deliverables. All seven ready PDF cases additionally have independent semantic/structural, Poppler visual where applicable, residue/revision/path-safety or conformance-claim, audit-provenance, and provider-trace graders; the 19 corpus/PKI cases remain explicitly asset-required.
- The bounded contract-ID defect was closed through the official typed path: `replace_text` now preserves the source baseline/default style and accepts the observed `0.0000227pt` provider quantization difference inside a non-configurable tolerance capped at `0.0005pt`, while genuine overflow and rotated text remain fail-closed. The Skill now requires provider probe and route planning before mutation and ships the canonical `open-office-artifact-tool.pdf-audit.v1` schema plus a validator that recomputes source/output hashes. One fixed candidate trial and one same-prompt, same-tarball reference-Skill trial each scored 100/100 across machine, visual, security, and trace evidence. Repeat trials and the remaining case graders/corpus are still open.
- The active-content public-sanitize fixed matrix passes candidate `3/3` and reference Skill `3/3`, all at 100/100 across machine, visual, security, and trace. The six runs bind clean commit `39fa301dcb1005f2848282e6e63da1e934104821`, byte-identical package SHA-256 `e78e18c0f8f1cffe301ae1f2ea17e882bc879b3044914033e24b0b11ac0e8b69`, identical prompt/input/oracle fingerprints, and fixed but distinct candidate/reference Skill fingerprints. Independent pypdf evidence proves root/additional JavaScript, Launch/SubmitForm, five attachments, invisible text, a comment, a populated widget, and personal metadata in each immutable source and zero active structure/canary residue in every output; Poppler changes stay inside the expected form/comment masks. Historical bypass and interpreter-drift runs remain defect evidence rather than passes. The evaluator ignores help-only invocations and permits independently scheduled probe/plan completion only when both precede the real typed edit; low-level bypass and post-mutation gates remain strict.
- The AcroForm adapter now distinguishes text/choice strings from radio/checkbox PDF Name states and validates post-write field/widget appearances before transaction promotion. Its complete generated fixture contains five text widgets, two radio widgets, and one pre-checked checkbox. The clean fixed matrix binds commit `bffd35dbfdb94bb1183717703e7e55bfb83c3f3c`, package SHA-256 `7ab9e6a30035df5d0ef7ee9990f3a0445152877e58a7f0d065ede9ddc1db300b`, and identical prompt/input/oracle fingerprints. Candidate passes `3/3` at 100/100; reference passes `2/3` at 100/100. The remaining reference run earned 70 raw machine/visual points but correctly scored zero after manually writing incremental PDF objects, reporting a non-pypdf provider, omitting canonical audit provenance, and bypassing provider preflight, typed fill, and post-mutation audit gates. All semantically successful fixed runs preserve TIN/signature/unselected-radio/checkbox pixels and editability, preserve the exact source prefix with one appended revision, and render every page through Poppler.
- The pypdf attachment-quarantine primitive covers document/page FileSpecs, duplicate and Unicode names, traversal-safe cross-platform naming, decoded-byte budgets, transactional extraction, raw identity/MIME/size/SHA evidence, and source/manifest provenance without opening payloads. Its fixed matrix binds clean commit `748fbb1d81ccfa14a594d6fed9bc6601866bfa95`, package SHA-256 `9cff93494c5b32e16394ce3b4fcffa1daf76ad6df57326dab0ced47d2a45b5bf`, and identical prompt/input/oracle fingerprints. Candidate passes `3/3` at 100/100; reference passes `2/3` at 100/100. The retained reference miss extracted the correct six payloads safely but used a custom Node parser and alternate manifest/audit contract, so missing pypdf/read-only/no-fallback preflight, typed extraction, canonical schema, and byte-binding evidence correctly forced the run to zero.
- The six-page greenfield tagged-accessibility workflow now has a clean fixed matrix: candidate `3/3` and reference Skill `3/3`, every run at 100/100 across machine, visual, security, and trace. All six records bind clean commit `2323a70331b93781dee37aa05198e4a73a7ec533`, byte-identical package SHA-256 `cfbcf5c76ba5fdb929dae27f2a0295d6da12694eec2150a926cebeecedefccb9`, identical prompt/input/oracle fingerprints, and fixed but distinct candidate/reference Skill fingerprints. Independent pypdf traversal proves title/language, H1-H3, one logical Table across pages 3-4, Figure alt, Link/StructParent/OBJR, reading order, and 12 running artifacts; Poppler/Pillow proves every page is nonblank, unclipped, and contains the expected physical table segments. The workflow separates modeled checks, optional veraPDF machine evidence, and required human PDF/UA judgment rather than claiming automatic conformance.
- The merge/reorder/selective-watermark clean fixed matrix passes candidate `3/3` and reference Skill `3/3`, every run at 100/100 across machine, visual, security, and trace. All six runs follow provider check/inspect/plan, the typed pypdf manifest primitive, typed Poppler comparison, output inspection, and multi-source canonical audit, and bind commit `90cbb9e0a5527a4620a28bb38aad8feeca895a3b`, byte-identical package SHA-256 `c3962993aee732c7a8e60282159409dfe08ca2c7c4f0dd59eb80468c28630ff5`, identical prompt/input/oracle fingerprints, and fixed but distinct candidate/reference Skill fingerprints. The comparator closed a real QA-discipline gap: a prior Agent deleted a correct output after subjectively misreading two thumbnails as black even though their renders were more than 98.7% white and byte-identical to successful trials.
- LibreOffice opened the shipped 26-slide reference template and produced a 26-page PDF; bounded custom-geometry icons rendered visibly. This local LibreOffice build substituted `Helvetica Neue`, so pixel parity with the checked-in preview images is not claimed and remains a visual-fidelity gap.

## Hosted evidence

The AGPL/MuPDF candidate at commit `40fcee0931c541c6f2cb1639ead0d10e2b76c7e6` passed the hosted Linux `ci` workflow in [GitHub Actions run 29558072285](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29558072285) on 2026-07-17. The run completed with conclusion `success` in 4m52s and covered npm install, protocol/runtime verification, Chromium/native-tool checks, the full npm suite, generated API-doc cleanliness, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `170/170`.

The embedded-XLSX/OLE and Presentation native-object layering candidate at commit `6119c54ae05d4b60fe562641e7aef10130581782` passed the hosted Linux `ci` workflow in [GitHub Actions run 29558401718](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29558401718) on 2026-07-17. The run completed with conclusion `success` in 4m22s and covered protocol/runtime verification, Chromium/native-tool checks, the full npm suite including payload-only OLE replacement and Presentation Skill regressions, generated API-doc cleanliness, offline release metadata, the 417-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `170/170`.

The Document-domain and shared text-range layering candidate at commit `765040c5827ee73c3c4645824688470e281c8be5` passed the hosted Linux `ci` workflow in [GitHub Actions run 29559403816](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29559403816) on 2026-07-17. The run completed with conclusion `success` in 5m22s and covered protocol/runtime verification, Chromium/native-tool checks, the full npm suite including root/leaf binding identity and native Documents workflow regressions, generated API-doc cleanliness, offline release metadata, the 419-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `170/170`.

The type-aware AcroForm candidate at commit `bffd35dbfdb94bb1183717703e7e55bfb83c3f3c` passed the hosted Linux `ci` workflow in [GitHub Actions run 29515214432](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29515214432) on 2026-07-16. The run completed with conclusion `success` and covered the full npm suite including AcroForm provider and independent grader regressions, deterministic OpenChestnut verification, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167`.

The active-content PDF sanitization candidate at commit `099bf1ce2f62ab992971e61b82641e6d6712a95d` passed the hosted Linux `ci` workflow in [GitHub Actions run 29501149008](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29501149008) on 2026-07-16. The run completed with conclusion `success` in 4m01s and covered deterministic OpenChestnut verification, Chromium/LibreOffice/Poppler tool checks, the full npm suite including the typed sanitize provider and independent grader regressions, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167`.

The Agent PromptBench scaffold at commit `70b2ddeea642de3729f8b7d7401bf10bace3be69` passed the hosted Linux `ci` workflow in [GitHub Actions run 29493964234](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29493964234) on 2026-07-16. The run completed with conclusion `success` and covered suite validation/tests inside the full npm gate, deterministic OpenChestnut verification, Chromium/LibreOffice/Poppler checks, API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167`.

The OpenChestnut XLSX sparkline candidate at commit `e8aa3e14249de346207f16b8fa24d7cb00b1253f` passed the hosted Linux `ci` workflow in [GitHub Actions run 29492891825](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29492891825) on 2026-07-16. The run completed with conclusion `success` in 4m02s and covered deterministic protocol/runtime verification, Chromium/LibreOffice/Poppler tool checks, the full npm suite, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167` including standard Office 2010 sparkline coverage.

The PDF provider-routing candidate at commit `b405ddd249c7c2f760c659c07e88495f3a3562f3` passed the hosted Linux `ci` workflow in [GitHub Actions run 29487829878](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29487829878) on 2026-07-16. The run completed with conclusion `success` in 3m59s and covered deterministic protocol/runtime verification, Chromium/LibreOffice/Poppler tool checks, the full npm suite including the provider contract tests, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge, and OpenChestnut 163-test execution. Optional Python providers remain an explicit local/environment gate rather than an undeclared hosted dependency.

The Documents native-workflow/OpenChestnut candidate through commit `e07e24382ff0259c7beefe27b0743d908a1f946f` passed the hosted Linux `ci` workflow in [GitHub Actions run 29483188346](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29483188346) on 2026-07-16. The run completed with conclusion `success` in 3m55s and covered protocol/runtime reproducibility, Chromium/native render tools, `npm test`, generated API-doc diff, offline release metadata, the registry-independent clean-install tarball, OfficeBridge, and OpenChestnut.

`npm run release:check` passes the source, documentation, package, license, JavaScript, and .NET gates. Its only remaining blocker is unavailable npm authentication. No `npm publish` or tag/release operation has been performed.

## Publishing

Before publishing:

1. Verify `package.json` and `package-lock.json` both declare `0.2.0`.
2. Rebuild and verify the OpenChestnut runtime from source.
3. Regenerate API documentation after the final public API change.
4. Inspect `npm pack --dry-run --json` for the required runtime/proto/Skill files and forbidden legacy files.
5. Run the tarball clean-install probe, not only source-tree tests.
6. Record the exact commit and hosted gate result used for publication.
