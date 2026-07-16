# Reference Skill compatibility

The 0.2 source tree publishes the reference file-type layout as four native Codex plugin bundles, not as the earlier flat project-specific Skills:

```text
skills/
  documents/{.codex-plugin,README.md,assets,skills/documents}
  spreadsheets/{.codex-plugin,.app.json,README.md,assets,skills/{spreadsheets,excel-live-control}}
  presentations/{.codex-plugin,README.md,assets,skills/presentations}
  pdf/{.codex-plugin,README.md,assets,skills/pdf}
```

There are therefore four plugin packages and five Skills. The old fixture runners remain under `test/skill-harness`; they are development tests and are excluded from the npm package.

## Verified compatibility

| Surface | Status | Evidence and remaining boundary |
| --- | --- | --- |
| Plugin manifests and discovery | done | All four manifests pass the Codex plugin validator. Every declared Skill, `agents/openai.yaml`, plugin icon, and Skill icon resolves inside its plugin bundle. |
| Private-package imports | done | Published JavaScript examples and runners import `open-office-artifact-tool`; the package test rejects imports from `office-artifact-tool`. |
| Presentations built-in template | done | The unflattened 26-slide `codex-grid-layout-library` runs through its shipped `create-presentation.mjs`, canonical OpenChestnut export, and second import. Text runs, body properties, absolute frames, 11 literal custom geometries, one prompt placeholder image, and two connectors survive. |
| Presentation workspace helper | done | The shipped setup helper resolves the public package root or bundled runtime, creates a module workspace, and links/imports `open-office-artifact-tool` without a private runtime path. |
| Spreadsheet core example | done | The reference-style create/value/formula/fill/style/chart/SVG-render/export/import example and `Workbook.fromCSV` run against the public package. |
| Full Spreadsheet Quick API | partial | The core vertical slice is compatible. Broader reference calls such as `formulasR1C1`, `displayFormulas`, `formulaInfos`, `Range.write/writeValues`, copy aliases, and several range-navigation aliases still require compatibility work or documentation narrowing. Formula coverage is bounded by the public Help catalog, not the reference document's historical formula-count claim. |
| Excel live control | partial | Routing content, Skill metadata, icon, and `.app.json` connector declaration are present. Execution requires the host-provided connected-document app plus an active Excel add-in session; the npm package does not implement that service. |
| Documents | partial | The native plugin content and render/QA workflow are packaged. Its reference authoring instructions still use Python DOCX/OOXML helpers directly. Final convergence must route ordinary creation/import/edit/export through `DocumentModel` and OpenChestnut, retaining low-level scripts only as explicit package-patch or QA tools. |
| Full Presentation API guide | partial | The built-in source-free template is compatible. Broader reference instructions for advanced Master/Layout, notes/comments, custom shows, groups, and other package graphs exceed the current canonical export boundary and remain fail-closed or preservation-only. |
| PDF | partial | The native PDF plugin content is packaged and the project PDF model/inspect/render/verify tests pass independently. The reference guide still needs a project-API adapter audit so its primary workflow consistently uses the public PDF surface instead of unrelated helper stacks. |

## Compatibility discipline

Plugin packaging and workflow compatibility are separate gates. A plugin is not marked fully compatible merely because its manifest validates or its files appear in the tarball.

For each remaining Skill, convergence requires:

1. run the published instructions from a clean npm install;
2. exercise the public package rather than repository-only test harnesses;
3. keep DOCX/XLSX/PPTX ordinary I/O on the single OpenChestnut facade;
4. retain direct OOXML patching only as an explicit, user-selected low-level operation;
5. render and inspect representative output where the environment supports it;
6. record unsupported reference calls and either implement them or narrow the published instruction honestly;
7. add the workflow to `test/reference-skills.mjs` before promoting its coverage status.

`test/reference-skills.mjs` is the publication-contract smoke test. The older format-specific tests under `test/*-skill.mjs` continue to exercise the deeper development fixtures under `test/skill-harness`.
