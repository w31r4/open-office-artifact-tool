# Clean-room template-library provenance

## Release decision

The Office Template Library is a source-free catalog. It must not contain a third-party Office reference file, preview image, package graph, raw XML, binary fingerprint, or source-derived Skill text unless the project has a clear redistribution authorization for that specific material.

This is a release policy, not a legal opinion. In the observed upstream template plugin, package metadata declared the material proprietary and did not include a redistribution license or notice. The locally observed migration work also retained those Office and PNG files byte-for-byte. That is insufficient for this AGPL repository and npm package, so those assets are excluded. OpenAI's published terms likewise distinguish permitted service use from a blanket right to copy or distribute the service itself; a future authorization must be explicit and asset-specific. [OpenAI Terms of Use](https://openai.com/policies/row-terms-of-use/)

The 20 catalog intents are therefore compatibility requirements only. They are useful descriptions of Agent tasks, not a license to reproduce a visual system or reuse an artifact.

## What ships

`skills/default-template-library/` is a normal AGPL plugin bundle. It contains:

- a `catalog.json` with 20 document, presentation, and workbook intents;
- one routing Skill plus six ready template Skills;
- family-specific reviewed JavaScript generators for Design Report and Strategy Memorandum (DOCX), Operating Review and Project Kickoff (PPTX), and Financial Budget and Project Tracker (XLSX);
- project-authored SVG icons and no Office or preview binaries.

The ready generator is the source of truth. It creates a new file at an explicit path, refuses to overwrite an existing output or audit, verifies the model, exports through OpenChestnut, imports it again, performs a second export/import, renders an SVG preview, and writes a hash-bound audit with `source: null` and `provenance: project-authored-source-free`.

The automated smoke test also makes one bounded edit to each generated DOCX/PPTX/XLSX and repeats its public import/export path. When LibreOffice and Poppler are installed, it converts each result to PDF, checks the intentional page-count contract, and rasterizes every native page. Financial Budget and Project Tracker require exactly one native PDF page per worksheet, so a horizontal table split is a test failure. A missing native tool skips only that native-render check; it does not turn a planned catalog item into a ready one.

## Catalog state

The catalog intentionally has two states:

| State | Meaning | Agent behavior |
| --- | --- | --- |
| `ready` | A self-authored generator and its round-trip tests ship in this package. | Generate it through the declared Skill, inspect the audit, then edit and verify through the matching Office workflow. |
| `planned` | The task intent is recorded, but no self-authored design has cleared the source, codec, and QA gates. | Explain that it is unavailable. Do not substitute a nearby visual design, search a cache, or download a reference file. |

At this milestone, six of twenty entries are `ready`; the other fourteen are intentionally `planned`.

## How a new template becomes ready

1. Design the artifact in this repository from a written, project-owned specification. Do not trace, extract, or imitate a prohibited binary reference.
2. Add a generator using the public `open-office-artifact-tool` API and OpenChestnut. Keep generated files out of the repository unless they are independently authored test fixtures with clear provenance.
3. Add semantic, import/edit/export/second-import, model-render, and available native-render tests.
4. Add provenance assertions that reject Office references, preview PNGs, cache paths, and source-derived binary equality checks.
5. Change the catalog item to `ready` only after the full package gate passes from a clean npm install.

For a user's own DOCX, PPTX, or XLSX reference, use Template Creator instead. It retains the user-supplied source locally in the user's chosen template home and is deliberately separate from the distributable library.
