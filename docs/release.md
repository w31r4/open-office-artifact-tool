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

The npm tarball is the consumer distribution. It contains the JavaScript object models, OpenChestnut adapter, generated wire binding, public proto, bundled runtime, integrity manifest, SBOM, license notices, render/QA helpers, PDF pipeline, the optional buildable OfficeBridge source project, and five native plugin bundles containing six Skills: the four file-type workflows, the separate `excel-live-control` route, and the local-only Template Creator utility. It excludes OpenChestnut C# source, every C# test and solution, all build output, repository-only build scripts, the development-only `test/skill-harness` fixtures, and the MIT-licensed repository-only Default Template Library with its 20 retained Office/PNG assets. MuPDF.js is declared in the required npm dependency graph rather than copied into this project's own tarball, and its WASM runtime is initialized only by a PDF operation.

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
- all five npm-distributed native plugin manifests validate, the published six-Skill topology is complete, and every workflow promoted to compatible in `docs/reference-skills.md` passes from the public package surface; the repository-only Default Template Library separately passes its canonical inventory, integrity, source-bound codec, and package-exclusion gates;
- PDF greenfield authoring plus default MuPDF.js import/inspect/render/bounded-edit, lazy-load, pre-WASM budget, exact-prefix incremental-save, source-bound annotation/link/form-field behavior, signature/redaction/deletion fail-closed, Skill CLI source-protection, and specialist-provider contract tests pass independently, including qpdf structure, pikepdf active/auxiliary structure cleanup, pyHanko local-PKCS#12 signing and signature validation, veraPDF conformance, and OCRmyPDF searchable-layer routing;
- when explicitly configured, the real optional-provider tests cover ReportLab creation, pdfplumber extraction, type-aware pypdf text/radio/checkbox forms and annotations, typed pypdf merge/reorder/selective watermarking, PyMuPDF rewrite/incremental/page/text/image/form/annotation edits, real redaction/scrub/residue scans, capped numerical text-fit behavior, canonical audit byte binding, typed Poppler source/output comparison, and OCRmyPDF/Tesseract recovery of a generated image-only scan with MuPDF second import and Poppler pixel invariance;
- Open XML SDK validation passes for generated Office fixtures;
- configured LibreOffice/Poppler/Playwright/native render gates pass where available;
- a production-only packed clean install completes all three Office roundtrips, PDF smoke, and a real packaged Template Creator invocation while `dotnet` is absent from `PATH`, and proves that the repository-only Default Template Library is absent;
- two clean OpenChestnut builds produce the same runtime file set and hashes;
- package contents contain no legacy Office codec files, OpenChestnut C# source, incomplete OfficeBridge solution, C# build output, tests, or repository-only scripts;
- package metadata, version `0.2.0`, licenses, third-party notices, SBOM, and integrity manifest agree;
- hosted Linux runs the same required non-optional gates.

The single byte-array managed entry point uses a checked-in, small JavaScript
interop registration shim rather than .NET 8's `[JSExport]` source generator.
That generator derives a wrapper name from the process-random
`String.GetHashCode()` and cannot satisfy the clean-build hash gate. The shim
uses the equivalent stable type-name hash and remains covered by the actual
bundled-WASM protocol smoke.

## Optional native validation

LibreOffice, Poppler, Playwright, and the Windows Office bridge are validation/render tools. Absence of an optional host may skip only the explicitly environment-gated native rendering branch; it must not skip codec, semantic, package, Skill, PDF, or clean-install gates.

The Office bridge does not participate in normal import/export and must never be used to hide a codec failure.

## Current local evidence

### Repository structure convergence

On 2026-07-19, a whole-repository reachability, package, generated-artifact,
native-project, and duplicate-asset audit found no orphaned production module:
all 95 `src` modules are reachable from the public package entries, all 98 C#
files belong to their MSBuild projects, and the runtime manifest exactly covers
the bundled OpenChestnut payload. The current tree therefore keeps the source,
generated wire binding, bundled runtime, test harnesses, canonical templates,
and self-contained Skill assets rather than deleting files merely because they
look duplicated.

The audit did remove the obsolete `handoff/2026-07-11` tree: 236 tracked files
and 6,098,763 bytes whose live reference-Skill consumer was PromptBench. The
reference subject now copies directly from the pinned
`reference/office-artifact-tool/skills` submodule after the deterministic
`skills/reference-sync.json` gate passes, so there is one upstream reference
source instead of a drifting second snapshot. Git history remains the archive
for the old handoff. The template starter helper also dropped 190 net lines of
scaffolding, including its fixed-false future implementation; its
current command performs a read-only preflight and fails closed without writing.

The npm manifest no longer publishes `OfficeBridge.sln`, because that solution
references a repository-only test project. It continues to publish the complete,
optional `native/OfficeBridge/src` project. An isolated checkout containing only
this convergence passed `npm test`, API-document regeneration with a clean diff,
the clean-install package smoke, and the offline release check including
OfficeBridge and OpenChestnut .NET tests. Its tarball contains 452 files, no
handoff/reference tree or incomplete solution, and 24,087,732 unpacked bytes.

### Reference Skill source synchronization

On 2026-07-19, the six retained reference plugin trees were audited against the
public `office-artifact-tool` submodule at commit
`256cb31bfe0a07b3cef0051b6b159342be381378`. The repository-only
`skills/reference-sync.json` records a deterministic snapshot of all 343 source
files and 33,186,487 bytes, with separate file counts, byte counts, and SHA-256
tree digests for Documents, Spreadsheets, Presentations, PDF, Template Creator,
and the Default Template Library. `scripts/reference-skill-sync.mjs check`
rejects a changed reference commit or tree and any reference path missing from
the corresponding project compatibility superset; it deliberately does not
claim byte identity for richer project adapters.

The audit exposed one real omission: the Documents
`examples/end_to_end_smoke_test.md` checklist. Before synchronization the new
gate failed with that exact missing path. The checklist is now retained
byte-for-byte, declared by the native Skill manifest, shipped in the npm plugin,
and documented as an optional route for the explicit Python render/package
helpers. The public OpenChestnut create/import/edit/export workflow remains the
default. The canonical 20-template Office/PNG assets continue to use their
separate byte-identity and package-exclusion gate; no self-authored template
generator or fallback was reintroduced.

The complete local gate passed `npm test`, `npm run docs:api`,
`npm run proto:check`, `npm run test:pack`, serial
`npm run verify:open-chestnut-build`, OpenChestnut `283/283`, and OfficeBridge
`5/5`. The official Skill and plugin validators also accepted the Documents
bundle. Two clean WASM builds reproduced all 39 audited package-layer files and
the same manifest-bound 38-file, 14,635,200-byte runtime. The production
clean-install tarball contains 453 files at 9,538,062 compressed bytes and
24,096,300 unpacked bytes; the repository-only sync snapshot/script and
canonical template library remain excluded. The specialist Python PDF provider
test remained contract-only because no explicit
`OPEN_OFFICE_PDF_PROVIDER_PYTHON` was configured; core MuPDF.js,
Playwright/Chromium, LibreOffice/Poppler, canonical template rendering, and all
other npm gates ran locally. No publish or tag operation was attempted.

### XLSX bounded multi-row PivotTables

On 2026-07-19, the Spreadsheet facade, OpenChestnut codec, Help catalog, and
runnable Spreadsheet Skill expanded native Pivot authoring from one row field
to 1 through 8 ordered row fields. The bounded profile is deliberately tabular:
every row field has a distinct cached-output column, automatic subtotals are
disabled, `location@firstDataCol` records the row-axis width, and authored
`rowItems` contain complete uncompressed cache-index tuples. Zero or one column
field, 1 through 32 values, exact filters on any configured axis, grand totals,
and saved-cache policy continue to compose with that hierarchy.

Import accepts the same bounded tabular/no-subtotal graph when a compatible
host omits the optional `rowItems`/`colItems` materialized caches. It preserves
recognized Pivot/cache parts byte for byte and binds ordered axes, filters,
source data, cached output, and topology as read-only. A duplicate/out-of-range
axis, more than 8 row fields, a second ordinary column field, or a compact or
subtotal-bearing multi-row graph remains opaque/source-bound or fails closed;
none is flattened into the tabular model.

The shipped workflow now exercises two row fields, one column field, two value
fields, and an exact region filter with independently checked totals. Its
eight-column summary keeps the previously CI-proven 440-pixel print-width
budget, with a smaller wrapped header, so both sheets remain one native page.
Bundled LibreOfficeDev 26.8 rendered and resaved both sheets. An Ubuntu 24.04
amd64 LibreOffice 24.2.7 resave normalized the native display range and omitted
the optional axis caches, then OpenChestnut recovered `Region -> Channel`, Product,
the filter, and the `440 / 44` grand totals and preserved that host graph. An
independently generated LibreOffice 24.2 DataPilot oracle confirmed the same
ordered `rowFields`, `firstDataCol`, and optional-cache structure. Open XML SDK
Office 2021 validation passes the authored complete-tuple graph.

The complete local gate passed `npm test`, `npm run docs:api`,
`npm run proto:check`, `npm run test:pack`, serial
`npm run verify:open-chestnut-build`, OpenChestnut `283/283`, and OfficeBridge
`5/5`. Two clean WASM builds reproduced all 39 audited package-layer files and
the same manifest-bound 38-file, 14,635,200-byte runtime. The production
clean-install tarball contains 452 files at 9,537,005 compressed bytes and
24,094,042 unpacked bytes; the repository-only canonical template library
remains excluded. The specialist Python PDF provider test remained
contract-only because no explicit `OPEN_OFFICE_PDF_PROVIDER_PYTHON` was
configured; core MuPDF.js, Playwright/Chromium, LibreOffice/Poppler, canonical
template rendering, and all other npm gates ran locally. No publish or tag
operation was attempted.

### XLSX exact native Pivot item filters

On 2026-07-19, the Spreadsheet model, additive protocol-2 wire, generated
binding, OpenChestnut C# codec, Help catalog, and runnable Spreadsheet Skill
gained one bounded native manual-filter profile. A source-free PivotTable may
place one exact `include` or `exclude` filter on each configured row/column
field, with 1 through 1024 string, finite-number, boolean, or blank items.
OpenChestnut validates every item against the typed cache, derives the output
rectangle from rows that survive all filters, writes standard
`pivotField/items/item@h` visibility, and preserves the full source cache for
refresh. Unknown/duplicate/over-budget items, a filter that removes every
record, date/condition filters, and fields outside the native axes fail closed.

Import recovers the visible item set, binds it into the source semantic hash,
and preserves the native Pivot/cache parts byte for byte on an unchanged second
export. A corrupt or incomplete item-index graph remains opaque/source-bound;
filter mutation after import is rejected. The shipped two-measure workflow now
excludes one region, independently verifies its filtered revenue/unit totals,
imports twice, renders both sheets, and resaves through LibreOffice. Bundled
LibreOfficeDev 26.8 and an Ubuntu 24.04 amd64 container with LibreOffice 24.2.7
both normalized the file to the same visible item set and remained semantically
recognized; an `include` list may legitimately return as its equivalent
complementary `exclude` list. Open XML SDK Office 2021 validation passes the
authored graph.

The complete local gate passed `npm test`, `npm run docs:api`,
`npm run proto:check`, `npm run test:pack`, serial
`npm run verify:open-chestnut-build`, OpenChestnut `282/282`, and OfficeBridge
`5/5`. Two clean WASM builds reproduced all 39 audited package-layer files and
the same manifest-bound 38-file, 14,629,056-byte runtime. The production
clean-install tarball contains 452 files at 9,535,218 compressed bytes and
24,086,436 unpacked bytes. The specialist Python PDF provider test remained
contract-only because no explicit `OPEN_OFFICE_PDF_PROVIDER_PYTHON` was
configured; core MuPDF.js, Playwright/Chromium, LibreOffice/Poppler, canonical
template rendering, and all other npm gates ran locally. No publish or tag
operation was attempted.

### XLSX bounded multi-value PivotTables

On 2026-07-19, the Spreadsheet model, protocol-2 wire, OpenChestnut C# codec,
Help catalog, generated API reference, and native Spreadsheet Skill expanded the
bounded source-free PivotTable profile from exactly one value field to 1 through
32 ordered value fields. Each value independently retains its source field,
display name, and `sum`/`count`/`average`/`min`/`max` aggregation. Multi-value
tables author the canonical SpreadsheetML `x=-2` data-layout axis, both with and
without an ordinary column field; the output collision check covers the full
expanded rectangle. Recognized imports preserve all Pivot/cache parts byte for
byte on an unchanged second export. Missing or inconsistent data-layout axes,
more than 32 values, and unmodeled graphs remain opaque/source-bound or fail
closed rather than being flattened.

The shipped workflow independently verifies a two-measure revenue/units matrix,
exports and imports twice, and renders both source and Pivot sheets. A real
LibreOffice XLSX resave retained both measures and aggregations, then reimported
through OpenChestnut with correct totals. LibreOffice and Poppler produced a
clean two-page native review with no clipping on the macOS host and in a Debian
12 amd64 container running LibreOffice 7.4.7. An Ubuntu 24.04 amd64
LibreOffice 24.2.7 resave omitted optional `rowItems`/`colItems` caches while
retaining the canonical `x=-2` and ordered data fields; OpenChestnut recognized
that host-normalized graph and preserved its native parts unchanged. C# tests
separately cover Office 2021 validation for multi-value tables with and without
a column field, omitted axis-item caches, roundtrip identity, malformed-axis
preservation, and the 32-field budget.

The complete local gate passed `npm test`, `npm run docs:api`, `npm run
proto:check`, `npm run test:pack`, and serial `npm run
verify:open-chestnut-build`; OpenChestnut passed `281/281` and OfficeBridge
passed `5/5`. Two clean WASM builds produced the same 39 audited files and the
same manifest-bound 38-file, 14,610,112-byte runtime. The clean-install tarball
contains 452 files, is 9,524,121 bytes compressed and 24,058,724 bytes
unpacked. The optional specialist Python PDF-provider gate remained
contract-only because `OPEN_OFFICE_PDF_PROVIDER_PYTHON` was not explicitly
configured; the required MuPDF.js path and all other npm gates ran locally. No
publish or tag operation was attempted.

### PPTX Office 2021 modern comment threads

On 2026-07-19, the Presentation model, protocol-2 wire, OpenChestnut C# codec,
Help catalog, generated API reference, and native Presentation Skill closed one
bounded Office 2021 modern-comment slice. Source-free decks can author one root
with direct replies, independent person and timestamp metadata, active/resolved/
closed state, a top-level DrawingML element or exact text-range anchor, and an
explicit slide coordinate. Recognized imported graphs permit only existing
comment text and status changes: author/person/date identity, anchor/range,
position, topology/order, relationships, part paths, and source hashes are
re-proved and fixed. Reactions, task fields, extensions, rich text, nested
replies, nested group-child monikers, unknown anchors, connected comment parts,
mixed legacy/modern graphs, and modern-comment slide clones remain opaque and
fail closed rather than being flattened.

The shipped workflow proves source-free authoring, export/import, fixed-topology
root/reply text editing, resolve/reopen state, second import, package inspection,
model rendering, source immutability, atomic output, and a byte-bound audit. C#
tests independently exercise native author/comment parts, Office 2021 Open XML
SDK validation, unchanged-part preservation, identity and anchor mutation
rejection, out-of-bounds ranges, and connected-graph preservation. Legacy PPTX
annotations remain a separate unchanged-only import profile.

The complete local gate passed `npm test` including the canonical 20-template
LibreOffice/Poppler corpus and Playwright, `npm run docs:api`, `npm run
test:pack`, and serial `npm run verify:open-chestnut-build`; OpenChestnut passed
`274/274` and OfficeBridge passed `5/5`. Two clean WASM builds produced the same
39 audited files and the same manifest-bound 38-file, 14,552,252-byte runtime.
The clean-install tarball contains 449 files, is 9,497,480 bytes compressed and
23,966,478 bytes unpacked. The real optional Python PDF-provider gate remained
contract-only because `OPEN_OFFICE_PDF_PROVIDER_PYTHON` was not explicitly
configured; the required MuPDF.js path and the remaining npm gates ran locally.
The repository-only canonical Default Template Library remains byte-bound to
reference commit `256cb31`, excluded from npm, and has no self-authored generator
or visual fallback. No publish or tag operation was attempted.

### Canonical Default Template Library reference pin

On 2026-07-19, the public `reference/office-artifact-tool` submodule advanced
from `207ce094a55d82a37efdca42a1c5e9656f696962` to the canonical MIT template
commit `256cb31bfe0a07b3cef0051b6b159342be381378`. The retained project library
already used the latter commit's 20 templates; this closure makes a fresh
recursive checkout contain the exact authoritative source tree instead of
requiring an adjacent developer checkout. No retained Office or PNG asset was
rewritten, and the reference package/runtime remains excluded from the npm
tarball.

Hosted CI now checks out the fixed submodule and passes its template-library
root to the existing corpus test. That gate compares all 40 Office/PNG assets
byte-for-byte before running integrity budgets, transactional materialization,
all 20 public import/unchanged-export/second-import paths, the bounded edits,
and LibreOffice/Poppler rendering. The local full suite passed with the same
source-root setting, proving the checked-in `integrity.json` aggregate and every
individual asset against the pinned commit.

The complete local gate passed `npm test` including Playwright and native
template rendering, `npm run docs:api`, `npm run proto:check`, `npm run
test:pack`, and serial `npm run verify:open-chestnut-build`; OpenChestnut passed
`260/260` and OfficeBridge passed `5/5`. Two clean WASM builds retained the same
39 audited files and manifest-bound 38-file, 14,428,348-byte runtime. The npm
tarball still excludes the reference submodule and repository-only template
library: it contains 447 files, is 9,440,424 bytes compressed and 23,759,623
bytes unpacked. No publish or tag operation was attempted.

### PPTX bounded imported SlidePart placeholder-text edits

On 2026-07-19, the Presentation model, protocol-2 wire, OpenChestnut C# codec,
Help catalog, native Presentation Skill, and repository-only Default Template
Library converged on one bounded edit for a concrete placeholder owned by an
imported SlidePart. The source binding now reports `text_editable` separately
from whole-element `editable`: a fully recognized local `p:txBody` may change
through the ordinary `TextFrame.set(...)` or structured text-replacement path,
while placeholder identity, geometry, paragraph/run topology and formatting,
layout/master inheritance, relationships, and every other shape property remain
source-bound. A model-side capability flag cannot grant permission; export
re-reads the source package, re-proves its hash and native capability, and
accepts only the exact text-only protobuf delta.

The corpus regression imports all seven canonical retained PPTX templates
through the public facade, edits one visible SlidePart placeholder in each,
exports and reimports, and proves the new text plus stable native placeholder
identity, frame, paragraph/run formatting, and source capability. LibreOffice
and Poppler render both the retained inputs and all seven edited outputs. A
deliberate newline-topology change, formatting change, capability tamper,
unsupported native run graph, inherited layout/master projection, and pending
source clone remain fail-closed.

The complete local gate passed `npm test` including Playwright and the native
template render corpus, `npm run docs:api`, `npm run test:pack`, and serial
`npm run verify:open-chestnut-build`; OpenChestnut passed `260/260` and
OfficeBridge passed `5/5`. Two clean WASM builds produced the same 39 audited
files and the same manifest-bound 38-file, 14,428,348-byte runtime. The
clean-install tarball contains 447 files, is 9,440,298 bytes compressed and
23,759,304 bytes unpacked. `npm run release:check` passed every code,
documentation, package, license, JavaScript, and .NET gate; before this
candidate commit it reported only the intentionally dirty worktree plus the
external npm-authentication blocker. No publish or tag operation was attempted.

### PPTX source-bound slide-name edits

On 2026-07-18, the canonical OpenChestnut PPTX route gained one deliberately
small imported-deck metadata edit: an original, source-bound `slide.name`
changes only that SlidePart's existing `p:cSld/@name` attribute. The source
binding, layout binding, direct-background contract, notes/comments, and fixed
element topology remain preconditions; no relationship or presentation graph is
created, removed, or rebound. The transaction validates the requested name
after write and then proves every other package part is byte-identical.

The JavaScript regression uses a multi-slide imported deck, changes one name,
requires all non-target decoded parts to match byte-for-byte, and reimports the
result. The C# codec regression validates the saved package with Open XML SDK
Office 2021 validation and checks the native name after a second import. A
pending clone still must be an exact source copy, so renaming it before its
export/reimport boundary fails closed.

The complete local gate passed `npm test` (including Playwright), `npm run
docs:api`, `npm run proto:check`, `npm run test:pack`, and serial `npm run
verify:open-chestnut-build`; OpenChestnut passed `211/211` and OfficeBridge
`5/5`. Two clean WASM builds produced the same 39 audited files and the same
manifest-bound 38-file, 14,417,084-byte runtime.

### PPTX bounded source-bound connector clone

On 2026-07-18, the existing imported `slide.duplicate()` transaction gained one
additional inline leaf: a canonical straight or elbow connector. It remains a
closed, unchanged SlidePart copy rather than a general relationship-graph
clone. Every present endpoint must resolve to an element in that same copied
SlidePart tree; a missing or cross-tree target fails before a pending clone can
be inserted.

The pending JavaScript clone exposes fresh clone-local element identities, so
Agent `inspect`/`resolve` operations follow the copied targets naturally. The
first export keeps the source-bound endpoint identities private at the codec
boundary, preserving the original SlidePart bytes while the connector binds to
the copied shapes in the new SlidePart. The new clone XML may be canonically
serialized by Open XML SDK, so its proof is structural/semantic rather than a
false lexical-byte claim. The connector adds no OPC relationship. Unsupported
connector forms, targets outside the clone tree, edits before export/reimport,
and all broader graph edges remain fail-closed.

The JavaScript regression proves cloned group connector endpoints resolve to
the fresh copied child identities, source `slide1.xml` remains byte-identical,
the second import resolves the copied endpoints, immediate clone mutation is
rejected, and an unresolved source target leaves no partial clone. The C#
regression validates the exported package with Open XML SDK Office 2021 and
asserts that the cloned SlidePart contains the connector while the source slide
bytes remain unchanged.

The complete local implementation gate passed `npm test` (including
Playwright), `npm run docs:api`, `npm run proto:check`, `npm run test:pack`,
and serial `npm run verify:open-chestnut-build`; OpenChestnut passed `211/211`
and OfficeBridge passed `5/5`. The reproducibility check kept the runtime at
39 audited files and 38 bundled files / 14,417,084 bytes.

### PPTX auditable source-bound slide duplicate workflow

On 2026-07-18, the Presentation Skill added
`openchestnut-slide-duplicate-workflow.mjs` for the imported-slide clone
profile. It is an Agent transaction, not a generic ZIP copier: its safe default
selects one explicit unique source name, accepts the canonical inline leaf
profile including same-tree straight/elbow connectors, and deliberately refuses
NotesSlide and legacy-comments leaves.

The workflow records input/output hashes, the actual source and clone
`SlidePart` paths, adjacent insertion, the only permitted new parts, retained
source-part byte preservation, reimported structural equality, and model-SVG
visual equality. Fresh `data-*-id` locator attributes are removed only for that
visual comparison; the new clone Slide XML may be canonically serialized and is
not claimed byte-identical. Missing/ambiguous names, closed leaves outside the
selected profile, unresolved connector endpoints, unsupported graph content, or
an unexpected package delta abort and remove temporary outputs/audits.

The same workflow now exposes a separate `allowClosedLeaves: true` /
`--allow-closed-leaves` route; it never activates by inference. It accepts at
most one canonical NotesSlide with exactly its NotesMaster/back-to-source-slide
relationships and at most one canonical legacy comments part with no child
relationship graph plus the immutable presentation-wide author catalog. The
audit names the added notes/comments parts, proves the leaf XML is verbatim
copied, proves the clone notes back-reference points to the clone SlidePart,
and proves the shared immutable resources remain shared after reimport. It does
not make rich/modern comments or arbitrary graph cloning available.

The complete local gate passed `npm test` (including Playwright), `npm run
docs:api`, `npm run proto:check`, `npm run test:pack`, and serial `npm run
verify:open-chestnut-build`; OpenChestnut passed `211/211` and OfficeBridge
`5/5`. The packed clean-install probe executes the installed workflow with
`dotnet` absent from `PATH`, including the explicit closed-leaf route. The
reproducibility check retained 39 audited files and a 38-file,
14,417,084-byte runtime.

### PPTX auditable imported slide-name workflow

On 2026-07-18, the native Presentation plugin added
`openchestnut-slide-name-edit-workflow.mjs` for the preceding bounded
source-bound name operation. It accepts an immutable input, distinct PPTX and
audit outputs, one expected name, and one replacement name. It resolves the
actual `presentation.xml` SlidePart relationship order, refuses duplicate,
missing, fallback-only, or mismatched names, performs only the typed
`slide.name` assignment, and promotes output only after package, reimport,
semantic, model-SVG, and ordinary presentation verification.

Its audit binds source/output hashes, the target SlidePart, requested native
attribute, and explicit no-fallback rewrite policy. It proves package topology
and every non-target part byte-identical. The target `SlidePart` itself may be
XML-canonicalized by Open XML SDK, so its safety proof is the requested native
name plus typed postwrite/reimport semantics rather than a false lexical-byte
promise. The plugin regression runs both the programmatic and CLI paths and
proves a missing target creates neither output nor audit.

The complete local gate passed `npm test` (including Playwright), `npm run
docs:api`, `npm run proto:check`, `npm run test:pack`, serial `npm run
verify:open-chestnut-build`, offline release metadata, OpenChestnut `211/211`,
and OfficeBridge `5/5`. The packed clean install includes the workflow and the
bundled runtime remains reproducible at 38 files / 14,417,084 bytes.

### PDF source-bound single-widget form-field updates

On 2026-07-18, the default direct-original MuPDF.js route added
`update_form_field`. Native inspection now emits individual `mupdfWidget`
records and grouped `mupdfFormField` records whose `id` is a current-source
`mupdf-form-field-<xref>` locator and whose `snapshot` captures the semantic
field state plus every visible widget rectangle.

The operation requires the exact inspection SHA-256, that locator, and the full
snapshot. It deliberately accepts only one non-password text widget, one
non-multiselect combo whose display/export values match, or one checkbox. It
verifies the value and field structure after applying the native mutation. A
shared widget group, radio/list/multi-select/password field, mismatched export
mapping, stale snapshot, unknown choice, or signed incremental request fails
closed and is routed to the explicit pypdf workflow instead.

Unlike annotation and link mutation, this narrow non-destructive form update
may use unsigned byte-prefix-verified incremental save. It remains a current
byte-sequence locator rather than a durable document identity, so every output
must be re-inspected before another field update. The native regression fixture
exercises text, combo, checkbox, shared-radio refusal, stale hash/snapshot,
unknown-choice, widget budget, and exact-prefix behavior without requiring a
Python provider.

### PDF direct-original page rotation

On 2026-07-18, the default direct-original MuPDF.js route added
`PdfFile.editPdf(..., { operations: [{ type: "rotate_page", page, rotation }] })`.
The operation accepts only absolute `0`, `90`, `180`, or `270` degree clockwise
values, reports the prior normalized page value, writes only the page's
`/Rotate` entry, and verifies the value after writing. It does not transform,
reflow, or remove page content.

Native inspection now emits normalized page rotation alongside raw `MediaBox`
and `CropBox` facts. On unsigned input the operation can use byte-prefix-verified
incremental save; malformed inherited rotation, unsupported angles, signed
incremental edits, and every other unsafe save mode fail closed. The library
and published Skill CLI regressions prove raw page-box preservation, the
90-degree rendered dimension swap, absolute reset back to `0`, exact source
prefix preservation, and rejection of `45` degrees.

### PDF direct-original visible CropBox edits

On 2026-07-18, the independent PDF path added one bounded existing-file page
geometry primitive: `PdfFile.editPdf(..., { operations: [{ type:
"set_page_crop", ... }] })`. MuPDF.js now exposes raw `MediaBox` and
`CropBox` evidence for every native inspected page, accepts a raw unrotated
`[x, y, width, height]` CropBox wholly inside the inspected MediaBox, maps it
through any existing crop-origin shift, verifies the raw written box, and
renders the resulting visible size from the original bytes.

This is deliberately a visible-window operation, not content removal. The
operation records `contentRemoved: false`; off-window content remains in the
file, so it is never a redaction or sanitize route. It may use unsigned
byte-prefix-verified incremental save, while deletion/redaction and every
signed incremental edit remain rejected. Rotated pages fail closed rather than
exposing ambiguous page-box coordinates; callers must select an explicit
specialist provider for that case.

Library regressions prove raw box inspection, crop rendering dimensions,
incremental prefix preservation, restoring a pre-existing crop in the original
page coordinate system, out-of-MediaBox rejection, and rotated-page rejection.
The published PDF Skill CLI regression performs the same crop from an
operations JSON, reopens it through `PdfFile.inspectPdf`, and renders it at
72 DPI. The API reference, provider matrix, save-policy guide, and
existing-PDF workflow all state the visible-only boundary.

The complete local gate passed `npm test` (including Playwright), `npm run
docs:api`, `npm run test:pack` (444 files; 9.4 MB packed, 23.6 MB unpacked),
offline `npm run release:check -- --skip-network --allow-dirty`, OfficeBridge
`5/5`, and OpenChestnut `210/210`.

### PDF source-bound annotation deletion

On 2026-07-18, direct-original MuPDF.js editing added bounded deletion of one
imported native annotation. `PdfFile.inspectPdf` now returns a SHA-256 of the
exact input bytes plus bounded `mupdfAnnotation` records. Each record exposes a
page/xref locator and semantic facts such as type, contents, author, subject,
and rectangle.

`delete_annotation` accepts only an inspection-returned locator for the same
page, the exact source SHA-256, and one or more semantic snapshot
preconditions. It rejects ignored precondition keys, stale hashes, mismatched
snapshots, ambiguous/missing locators, signed incremental edits, and every
incremental delete. A successful operation verifies removal before saving,
uses rewrite only, and records source/output hashes plus the matched snapshot
in its operation audit.

The locator is intentionally not a persistent document identity. A rewrite can
renumber or reuse an xref for a different surviving annotation, so every later
annotation mutation must start from a new inspection of the exact current
bytes. This rule avoids mutable annotation-array indexes while also avoiding
the false promise that an xref survives an arbitrary rewrite.

Library and published-Skill CLI tests cover bounded annotation enumeration,
source hash propagation, max-annotation rejection, rewrite-only enforcement,
stale hash and snapshot rejection, ignored-key rejection, exact one-target
deletion, surviving-note preservation, and output reinspection.

### PDF source-bound annotation update

The same route now supports one bounded `update_annotation` operation. It
requires the exact input SHA-256, the inspect-returned
`mupdf-annotation-<page>-<xref>` locator, and one or more semantic snapshot
facts before changing a Text annotation. Its patch is intentionally narrow:
only non-empty `contents`, `author`, and `subject` fields are accepted.
The operation reopens the page-native annotation list after update and verifies
that the same xref remains uniquely addressable and that every requested patch
field persisted before saving.

This is rewrite-only. Incremental revisions retain previous annotation values,
so they cannot satisfy the same audit/security meaning. It likewise requires a
new inspection for every following annotation mutation because an output rewrite
can renumber or reuse xrefs.

The geometry boundary is deliberate and fail-closed. A direct MuPDF probe
showed that a requested Text annotation rectangle is normalized to its native
minimum geometry (for example, a requested 24-by-24 rectangle was exposed as
20-by-20 after creation). Therefore `rect` is accepted only as a source
snapshot precondition; a `patch.rect` fails rather than claiming a move or
resize that native output cannot verify. An agent must explicitly delete and add
a replacement note, or route that task to a specialist provider.

Library and published-Skill CLI tests cover stale snapshot rejection,
incremental refusal, rejected geometry patch, exact text/author/subject update,
fresh-output reinspection, and a subsequent delete that uses only the refreshed
source hash and locator.

### PDF source-bound Text annotation creation

On 2026-07-18, `add_text_annotation` was moved from a direct convenience
mutation to the same agent-safe source-bound contract. It now requires the
exact input SHA-256 and an inspected target `mupdfPage` bbox/rotation snapshot,
accepts only an unrotated in-window `[x,y]` pin with non-empty `contents` and
optional non-empty author/subject, and is rewrite-only.

The contract intentionally does not accept a requested rectangle, `text` alias,
or icon selection. A native probe proved that MuPDF normalizes Text-note
geometry, so the operation verifies one new `Text` annotation at the requested
pin, records the provider's actual normalized rectangle, and requires a fresh
inspection before that result can be updated or deleted. The implementation
also snapshots the native annotation count before creation because MuPDF's
enumeration list is live and otherwise makes a false count delta appear valid.

Library and shipped CLI tests cover source-hash/page-snapshot binding,
incremental refusal, stale page rejection, out-of-window and rotated-page
failure, legacy alias refusal, exactly-one addition, normalized geometry, and
fresh-output inspection.

### PDF source-bound link deletion

On 2026-07-18, the same direct-original route made imported link deletion
agent-safe. Native inspection now emits bounded `mupdfLink` records containing
a page-bound fingerprint over URL, rectangle, and externality. The direct
parser uses that same opaque locator instead of its former link-array index.

`delete_link` requires the exact inspected source SHA-256, the page-bound
locator, and one or more URL/rectangle/external snapshot preconditions. It
rejects the legacy URL/index-only selection shape, stale hashes or snapshots,
ignored precondition keys, duplicate equivalent link fingerprints, all
incremental output, and any ambiguous/missing target. The mutation verifies
removal before rewrite save and records the matched facts in its operation
audit.

MuPDF does not expose a native PDF-object xref for every link handle, so the
fingerprint is deliberately source-byte-bound rather than a durable link ID.
Every rewrite requires a fresh inspect before another link operation. The link
budget protects both parser reconstruction and native inspection from
unbounded link-record expansion.

### PDF source-bound link update

Direct-original MuPDF.js editing now adds one bounded `update_link` operation.
It uses the same exact input SHA-256, inspected page/fingerprint locator, and
URL/rectangle/externality snapshot as link deletion, then replaces only one
safe non-empty internal `#...` or absolute `http`, `https`, or `mailto` target
URL. It reads the native link list again before saving and
requires exactly one retained link with the requested URL and the original
rectangle; otherwise the whole rewrite fails without output.

This is rewrite-only. A link update is not safe for an incremental revision
because the prior target remains in older bytes, and every later link operation
must use a fresh inspected source hash and fingerprint.

The operation deliberately does not expose `patch.bbox`. A direct native probe
called `setBounds([96, 144, 216, 168])`, but after a save/reload the public
inspector reported `[96, 624, 120, 24]`. That coordinate-layer leak is not a
stable public geometry contract, so the API refuses a bounds patch rather than
publishing a misleading move primitive. To move a link, an agent must use one
source-bound `delete_link` then `add_link` rewrite transaction, or a specialist
provider.

Library and published-Skill CLI tests cover stale snapshots, incremental
refusal, rejected bounds patches, URL update with retained rectangle,
fresh-output reinspection, and a subsequent delete that only uses the refreshed
link fingerprint and source hash.

### PDF source-bound link creation

The direct-original route now exposes `add_link` for one bounded imported-PDF
placement. It is not a generic drawing or hit-testing API: the operation binds
to the exact inspected source SHA-256 and to the target `mupdfPage` visible
CropBox/rotation snapshot. It accepts only an unrotated page, a positive
`[x, y, width, height]` rectangle fully inside that CropBox, and either an
internal `#...` destination or an absolute `http`, `https`, or `mailto` URL.
`javascript:`, `file:`, `data:`, malformed URLs, stale page facts, and an
existing identical URL/rectangle pair fail before save.

`add_link` is rewrite-only and rereads the native page link list to prove that
exactly one new matching record exists. Its audit records the source page
snapshot, added URL/rectangle/externality, and pre/post count; callers still
must inspect the new bytes before using any `mupdf-link` locator. This closes
the safe public move path: one operation list can delete an old source-bound
link first, then add its replacement using the same input hash and page
snapshot. Library and shipped CLI coverage prove add, unsafe-scheme refusal,
stale-page-field refusal, fresh inspection, and the combined move transaction.

### XLSX criteria extrema and branch formulas

On 2026-07-18, the bounded JavaScript calculation catalog added `MINIFS` and
`MAXIFS`. Each requires a value range followed by one or more
criteria-range/criterion pairs with exactly the same rectangular shape. The
implementation reuses the existing case-insensitive comparison and Excel
wildcard criteria semantics, considers only finite numeric values in matching
value cells, returns `0` when no numeric value matches, and fails with
`#VALUE!` before calculation when a criterion range has a different shape.

The same catalog now adds `IFS` and `SWITCH` for concise agent-authored branch
logic. `IFS` walks condition/value pairs in order and evaluates only the first
selected result; it returns `#N/A` for no match and `#VALUE!` for malformed
pairs. `SWITCH` matches one expression against ordered value/result pairs,
supports an optional default, and otherwise returns `#N/A`.

The public Help catalog documents all four functions and regenerates the API
reference. Spreadsheet smoke tests cover multi-criterion minimum/maximum,
ignored nonnumeric matched cells, the no-match `0` result, mismatched-range
rejection, short-circuit branch behavior, `SWITCH` default/no-match behavior,
and canonical OpenChestnut XLSX export/import; Help tests pin the catalog and
workbook counts and the numeric formula schemas.

### XLSX bounded 2D bubble charts

On 2026-07-18, the Spreadsheet public model, protocol-2 wire, OpenChestnut C#
codec, bundled WASM runtime, generated Help/API reference, clean-install pack
probe, and native Spreadsheet Skill converged on real 2D bubble charts.
`sheet.charts.add("bubble", range)` accepts exactly `X | Y | Size` and creates
one unambiguous formula-bound series; a multi-series bubble chart must instead
use explicit `xValues/xFormula`, `values/formula`, and
`bubbleSizes/bubbleSizeFormula` fields. The three vectors must have equal point
counts, and sizes must be finite and positive.

OpenChestnut authors native `c:bubbleChart`, `c:xVal`, `c:yVal`,
`c:bubbleSize`, and two numeric value axes. It uses an explicit 2D,
100%-scale, area-sizing profile. Imported `bubble3D`, negative-bubble,
custom-scale, or non-area-size graphs are detected as source-bound/read-only;
unchanged export retains them exactly and mutation fails closed. Bubble SVG
preview maps the positive size vector proportionally to visible circles rather
than pretending a bubble plot is a scatter plot.

The runnable Skill workflow creates, inspects, verifies, exports, imports,
edits, re-exports, and re-imports one formula-backed bubble chart. Its real
LibreOffice-to-PDF plus Poppler QA passes in one page. The production-only
tarball smoke also creates/imports the bubble chart while `dotnet` is absent
from `PATH`; C# tests validate Office 2021 XML, fixed-topology edits, and
exact preservation/rejection of a noncanonical `bubbleScale` profile.

### XLSX marker-only scatter charts

On 2026-07-18, the Spreadsheet public model, protocol-2 wire, OpenChestnut C#
codec, bundled WASM runtime, generated Help/API reference, and native
Spreadsheet plugin converged on real numeric-X/Y marker-only scatter charts.
Range-backed creation treats the first numeric column as each series' X source
and the remaining columns as Y series; explicit series carry `xValues/xFormula`
and `values/formula`. OpenChestnut authors native `c:scatterChart`, `c:xVal`,
`c:yVal`, and two value axes instead of substituting a category line chart.

Marker-only is a cross-host contract, not only a chart-style label. Every
authored or edited series now carries an explicit DrawingML no-fill outline so
LibreOffice and Microsoft Office do not synthesize connecting lines.
Series-level `line`/`stroke` is rejected; `marker.line` remains the supported
marker-border control. Imported scatter line/smooth variants and marker-style
graphs with an explicit series line remain source-bound/read-only and are exact
preserved until a broader semantic profile exists.

The runnable Skill example creates two formula-bound series, imports, edits,
re-exports, re-imports, inspects, verifies, and renders SVG. Its real local
LibreOffice-to-PDF plus Poppler run produced one complete page with separate
circle/diamond points and no connecting lines; package, semantic, SVG, native
page-count, and raster QA all passed. The package test also exercises scatter
create/import from a production-only clean install with no `dotnet` on `PATH`.

Final candidate `77706e52e72a22a6ff76bc11ca7dda29e9434b01` passed hosted
[CI run 29600831755](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29600831755)
in 4m43s. The run passed deterministic protocol/runtime verification, installed
Chromium/LibreOffice/Poppler, the complete `npm test`, regenerated API-doc diff,
offline release check, clean-install pack gate, OfficeBridge, and OpenChestnut.
Local evidence for the same candidate includes OpenChestnut `186/186`,
OfficeBridge `5/5`, a reproducible 39-file build audit, a 38-file bundled runtime
of 14,329,020 bytes, and a 435-file npm tarball of 9,301,407 compressed bytes /
23,199,173 unpacked bytes. Offline release metadata reports `publishReady: true`;
real npm authentication/publication remains an external release action.

### XLSX What-If data tables

On 2026-07-18, the Spreadsheet public model, protocol-2 wire, OpenChestnut C#
codec, bundled WASM runtime, generated Help/API reference, and native
Spreadsheet plugin converged on canonical Excel What-If data tables.
`sheet.dataTables.add(range, { rowInput, columnInput })` authors row-oriented
and column-oriented one-variable tables plus two-variable tables as native
`<f t="dataTable">` formulas. Import exposes defensive result-range, anchor,
input, orientation, and display-formula evidence without projecting the native
anchor as an ordinary editable cell formula.

The bounded profile is loss-aware: source-free authoring rejects missing
top-left formulas, cross-sheet or out-of-bounds input cells, undersized grids,
and overlapping result ranges. Recognized imports retain fixed count, order,
result range, inputs, and orientation; ordinary surrounding-cell edits and a
second export/import work, while topology or semantic data-table mutation fails
closed. Noncanonical attributes remain source-bound and unchanged. Excel,
LibreOffice, or another compatible host calculates the result values; the
JavaScript evaluator does not pretend to simulate `TABLE`.

The complete candidate at commit
`94f1d2234c93078b0782a52fd95c115be9df86f8` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29594786964](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29594786964)
in 5m07s. The run covered deterministic protocol/runtime verification,
Chromium/LibreOffice/Poppler setup, the full npm suite including the runnable
What-If Skill workflow and fail-closed topology corpus, generated API-doc
cleanliness, offline release metadata, clean-install packing, OfficeBridge
`5/5`, and OpenChestnut `185/185`.

### Neutral template library and Template Creator sync

On 2026-07-18, the observable Skill delta from the remotely reachable reference
revision `207ce094a55d82a37efdca42a1c5e9656f696962`
(`origin/feat/sync-grid-layout-library-template-creator`, based on the
`office-artifact-tool@2.8.24` runtime-sync revision) was adapted into the public
package. The Presentation asset/support trees now use the neutral
`grid-layout-library` identity throughout. Its model-facing registry, 26 exact
plain-JavaScript Compose builders, 27 preview PNGs, content/design tokens, and
reconstruction runner remain unflattened and pass a complete 26-slide
OpenChestnut export. The submodule now pins that exact reachable revision; no
unreachable pointer is recorded.

The new fifth plugin contains the sixth published Skill, Template Creator. It
creates numbered local `artifact-template-*` Skills from one DOCX/PPTX/XLSX
reference and a structurally validated PNG, or updates exactly one named
same-kind template. The retained Office reference and preview are byte-exact.
The implementation performs no network fetch, writes only beneath
`${OFFICE_ARTIFACT_HOME:-~/.office-artifact-tool}/skills`, enforces 512 MiB
reference and 64 MiB preview budgets, rejects symlink-bearing update trees,
serializes writers, recovers interrupted replacements, preserves extra
template-owned files, rolls back failed placement, and proves there is no
stage/backup/lock residue. At that milestone, all five native plugin manifests passed the official
validator, while the reference-compatible versioned Template Creator manifest
is retained alongside the native manifest.

The complete local gate passed `npm test`, `npm run docs:api`,
`npm run proto:check`, `npm run test:pack`, serial
`npm run verify:open-chestnut-build`, OpenChestnut `185/185`, and OfficeBridge
`5/5`. Two clean WASM builds produced the same 39 audited files and the same
manifest-bound 38-file, 14,324,412-byte runtime. The clean-install tarball
contains 434 files, is 9,297,469 bytes compressed and 23,178,713 bytes
unpacked. The pack gate now also invokes Template Creator from that offline,
production-only installation, validates the generated spreadsheet-template
sidecar, and proves the retained XLSX and PNG bytes exactly match its inputs.
The final candidate at commit
`4f3dbbb4f95081e1c86747e88c10b4b16722ad0d` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29596412072](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29596412072)
in 4m40s, covering npm installation, deterministic OpenChestnut verification,
Chromium/native-tool setup, the complete npm/Skill/security suite, generated
API-doc cleanliness, offline release metadata, clean-install packing,
OfficeBridge `5/5`, and OpenChestnut `185/185`.

### Earlier clean-room default template library (superseded)

The subsequent Default Template Library is not a binary migration of an
observed template pack. Its source plugin metadata marked the observed material
proprietary, and the locally observed migration retained its Office and PNG
assets byte-for-byte. The public AGPL package therefore ships neither those
files nor source-derived package graphs, hashes, or Skill text.

Instead, the sixth plugin contains a source-free 20-intent catalog. Three
independently authored entries are currently ready: Strategy Memorandum DOCX,
Project Kickoff PPTX, and Financial Budget XLSX. Their bundled generator
creates new output only, uses OpenChestnut, verifies/export/imports/second
imports and model-renders each artifact, then emits a hash-bound audit with no
source file. The direct test edits each result through its public model and,
when LibreOffice and Poppler are available, performs PDF/page/raster QA. The
offline production tarball test invokes all three shipped generators while
`dotnet` is absent from `PATH`. The remaining 17 catalog entries are explicit
`planned` records and fail closed rather than silently selecting an unrelated
design. See [clean-room template-library provenance](template-library-provenance.md).

### MIT reference-backed Default Template Library

Later on 2026-07-18, the template source obtained an explicit public MIT
distribution record. The canonical source is
[`office-artifact-tool` commit `256cb31bfe0a07b3cef0051b6b159342be381378`](https://github.com/w31r4/office-artifact-tool/commit/256cb31bfe0a07b3cef0051b6b159342be381378),
whose root `LICENSE.md` states `Copyright (c) 2026 w31r4`. The current source
tree therefore replaces the earlier source-free generators and planned catalog
with that one canonical library: 20 retained Skills (7 DOCX, 7 PPTX, 6 XLSX),
their original Office/PNG assets, source Skill metadata, an MIT notice, and
individual plus aggregate SHA-256 records. It is a normal forward migration;
the earlier commits are not rewritten.

The library is repository-only, not a sixth consumer-package plugin. A named
template is materialized only to a distinct output after its retained source
hash is verified, with a provenance audit and overwrite refusal. The source
files remain immutable. The complete 20-template matrix passes public-facade
import, unchanged export, second import, and available LibreOffice/Poppler
source/output rendering. Codec regressions cover one bounded edit per family:
PPTX slide-name metadata, DOCX `updateFields`, and ordinary XLSX string cells.
Unmodeled topology remains source-bound and fails closed; in particular, the
Financial Budget partial shared-formula range is preserved but cannot be
edited. The package-content and clean-install gates prove the library does not
enter the npm tarball. See [template-library provenance](template-library-provenance.md).

### XLSX standard-area and fixed-doughnut chart families

On 2026-07-17, the Spreadsheet public model, protocol-2 wire, OpenChestnut C#
codec, bundled WASM runtime, generated Help/API reference, and native
Spreadsheet plugin converged on two additional canonical worksheet-chart
families. Source-free and recognized imported standard-area charts now carry
the existing primary category/value axes; 50%-hole doughnut charts author and
import as circular plots without axes. Both families support bounded
categories, series, titles, legends, source formulas/caches, semantic edits,
second import, and type-aware SVG QA. The worksheet chart preview moved from
the large Spreadsheet domain leaf into its own focused module while preserving
the public `WorksheetChart.toSvg()` contract.

The slice remains deliberately loss-aware. Stacked and percent-stacked area,
non-50% doughnut geometry, exploded points, scatter/bubble/radar/stock/modern
families, and other unmodeled plot options are never substituted. Recognized
advanced area/doughnut graphs import as source-bound and read-only, preserve
their native chart XML unchanged, and reject semantic edits with
`unsupported_spreadsheet_chart_edit`; unsupported source-free types and
illegal axis/line-option combinations fail before crossing the wire. The
regressions also closed two existing chart-patch defects: valid imported
pie/doughnut topology is now editable, and editing a non-line chart cannot
inject line grouping markup.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, isolated `npm run verify:open-chestnut-build`,
OpenChestnut `184/184`, and OfficeBridge `5/5`. Two isolated clean WASM builds
produced the same 39 audited files and the same manifest-bound 38-file,
14,318,780-byte runtime. The clean-install tarball contains 423 files, is
9,276,398 bytes compressed and 23,122,542 bytes unpacked. No npm publish, tag,
or GitHub release operation was attempted.

The completed candidate at commit
`e61bb1ff143f27a9d4afc1ecdd6490634c3dfa4a` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29590185608](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29590185608)
on 2026-07-17. The run completed with conclusion `success` in 4m31s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including five-family chart author/import/edit/preservation
and native Spreadsheet Skill regressions, generated API-doc cleanliness,
offline release metadata, the 423-file clean-install tarball, OfficeBridge
`5/5`, and OpenChestnut `184/184`.

### DOCX transactional SEQ/REF cached-result materialization

On 2026-07-17, the public Documents model and native Documents plugin closed
the bounded automatic cached-result gap for canonical inline `SEQ` and `REF`
fields. `document.materializeFields()` computes independent, case-insensitive
SEQ counters in document order, resolves REF targets from the shared native
bookmark namespace, and updates only cached display text. The operation
supports dry-run evidence and defaults to strict transactional failure before
mutation when a bookmark target is missing, duplicated, or dangling. An
explicit non-strict mode may update resolvable fields while reporting missing
targets.

The primitive does not remove or flatten native field topology. PAGEREF is
reported but never fabricated; explicitly requesting it fails closed because
trustworthy page numbers require a real pagination host. Imported field
positions, instructions, and bookmark identity remain source-bound. The
shipped Documents fixture now materializes one caption SEQ and its REF through
the public model before two OpenChestnut round trips, while retaining an
explicit PAGEREF cache for native-host refresh. JavaScript regressions cover
dry-run immutability, strict rollback, opt-in partial materialization,
case-insensitive targets, multiple counters, and the PAGEREF boundary.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, isolated `npm run verify:open-chestnut-build`,
OpenChestnut `183/183`, OfficeBridge `5/5`, and the Documents Skill validator.
Two isolated clean WASM builds produced the same 39 audited files and the same
manifest-bound 38-file, 14,317,244-byte runtime. The clean-install tarball
contains 422 files, is 9,274,329 bytes compressed and 23,114,497 bytes
unpacked. No npm publish, tag, or GitHub release operation was attempted.

The completed candidate at commit
`fbe86f197850f2bc297c6a95c5b5d5d19e442ab2` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29588303966](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29588303966)
on 2026-07-17. The run completed with conclusion `success` in 4m29s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the native field-materialization workflow and
transactional fail-closed regressions, generated API-doc cleanliness, offline
release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and
OpenChestnut `183/183`.

### DOCX caption-number inline bookmark workflow

On 2026-07-17, the public Documents run model, protocol-2 wire, OpenChestnut
C# codec, bundled WASM runtime, Help catalog, and native Documents plugin
closed the bounded caption-target gap. A source-free canonical `SEQ` run may
now set `bookmarkName` through `paragraph.addField(...)`; OpenChestnut writes
one paired `w:bookmarkStart`/`w:bookmarkEnd` around only that field's cached
result run. Canonical `REF` and `PAGEREF` runs can therefore target the native
caption number without an OOXML patch helper.

Whole-block and inline-field bookmarks share one case-insensitive name and
unsigned native-ID space. The two bounded profiles cannot nest in the same
paragraph. Import collapses the seven-child bookmarked field graph back into
one logical run, preserves name/native identity, permits cached-result and
ordinary-text edits, and rejects any field position, instruction, bookmark
name, or native-ID change. Unsupported placement and duplicate identity fail
closed; automatic field calculation/materialization remains an explicit host
or package-helper workflow rather than a codec claim.

The shipped Documents fixture authors `SEQ Figure \\* ARABIC` with `fig1`, then
targets it with both `REF fig1 \\h` and `PAGEREF fig1 \\h`. It completes two
OpenChestnut round trips, edits all three cached results, asserts that the
native bookmark encloses only the SEQ result run, and passes real
LibreOffice/Poppler page QA. C# and JavaScript regressions cover semantic
roundtrip, source-bound identity, invalid non-SEQ bookmark use, case-insensitive
duplicate names, and bounded-profile nesting rejection.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, isolated `npm run verify:open-chestnut-build`,
OpenChestnut `183/183`, OfficeBridge `5/5`, and the Documents Skill validator.
Two isolated clean WASM builds produced the same 39 audited files and the same
manifest-bound 38-file, 14,317,244-byte runtime. The clean-install tarball
contains 422 files, is 9,273,020 bytes compressed and 23,105,625 bytes
unpacked. No npm publish, tag, or GitHub release operation was attempted.

The completed candidate at commit
`206712aef6691490318e42bfa26553d27b4d1fc9` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29586820145](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29586820145)
on 2026-07-17. The run completed with conclusion `success` in 4m36s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the native caption/cross-reference workflow and
fail-closed identity regressions, generated API-doc cleanliness, offline
release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and
OpenChestnut `183/183`.

### DOCX canonical inline SEQ/REF/PAGEREF field runs

On 2026-07-17, the Documents paragraph/run model, versioned protobuf wire,
OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native Documents
plugin converged on one bounded inline-field profile. A paragraph can mix
ordinary runs with logical `SEQ <label> \\* ARABIC`, `REF <bookmark> \\h`, and
`PAGEREF <bookmark> \\h` fields through `paragraph.addField(...)`.
OpenChestnut materializes each logical field as the canonical five-run native
begin/instruction/separate/cached-result/end graph and imports that exact graph
back into one `run.inlineField` object.

Source-free cached results and ordinary text survive export, import, edits,
second export, and second import. On imported paragraphs, field positions and
instructions remain source-bound while cached display text remains editable;
topology or target changes fail closed. Non-canonical switches, partial or rich
field graphs and automatic SEQ/REF materialization remained opaque or explicit
advanced package workflows at that milestone; the later bounded caption-number
bookmark slice is recorded above.
The shipped native fixture covers all three commands, native XML evidence,
semantic inspection, cache edits, and real LibreOffice/Poppler page QA.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, isolated `npm run verify:open-chestnut-build`,
OpenChestnut `183/183`, and OfficeBridge `5/5`. Two isolated clean WASM builds
produced the same 39 audited files and the same manifest-bound 38-file,
14,309,564-byte runtime. The clean-install tarball contains 422 files, is
9,270,563 bytes compressed and 23,092,374 bytes unpacked. No npm publish, tag,
or GitHub release operation was attempted.

The completed inline-field candidate at commit
`05713031054388f92c0586e7f6547a407544c696` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29584754018](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29584754018)
on 2026-07-17. The run completed with conclusion `success` in 4m58s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the native inline-field Documents fixture and
source-bound topology regression, generated API-doc cleanliness, offline
release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and
OpenChestnut `183/183`.

### DOCX canonical complex TOC and field-refresh workflow

On 2026-07-17, the Documents public model, versioned protobuf wire,
OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native Documents
plugin converged on one bounded complex-field profile. Source-free documents may
author a canonical one-paragraph TOC field with configurable ascending heading
levels and the `\\h`, `\\z`, and `\\u` switches. The public
`document.addTableOfContents(...)` primitive emits native
`w:fldChar`/`w:instrText` topology and enables the `w:updateFields` open-time
refresh hint by default. The hint is modeled explicitly and never asserts that
the cached TOC result is current.

Canonical unrefreshed one-paragraph TOCs survive export, import,
fixed-topology instruction/display edits, second export, and second import.
Refreshed cross-paragraph TOCs, including otherwise plain-looking paragraphs
inside the field span, remain opaque/source-bound and read-only as one graph.
Arbitrary complex REF/PAGEREF/SEQ/external-content fields remain outside the
semantic profile. The shipped Documents fixture proves the public workflow,
native field/settings XML, semantic inspect evidence, LibreOffice rendering,
and fail-closed irregular-graph preservation.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, isolated `npm run verify:open-chestnut-build`,
OpenChestnut `182/182`, and OfficeBridge `5/5`. Two isolated clean WASM builds
produced the same 39 audited files and the same manifest-bound 38-file,
14,303,420-byte runtime. The clean-install tarball contains 422 files, is
9,268,734 bytes compressed and 23,076,840 bytes unpacked. No npm publish, tag,
or GitHub release operation was attempted.

The completed TOC/field-refresh candidate at commit
`f7c83de4fbb83ddade81542cb95df86d01be4939` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29582981065](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29582981065)
on 2026-07-17. The run completed with conclusion `success` in 4m28s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the native TOC Documents workflow and
cross-paragraph fail-closed regression, generated API-doc cleanliness, offline
release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and
OpenChestnut `182/182`.

### DOCX canonical bibliography and citation public workflow

On 2026-07-17, the Documents public model, versioned protobuf wire,
OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native Documents
plugin converged on one reversible bibliography/citation profile. Source-free
documents may author one canonical bibliography Custom XML catalog plus
whole-paragraph `w:fldSimple` `CITATION <tag>` fields. Ordinary Author name
lists or one corporate author, the bounded scalar source catalog, bibliography
settings, and multiple sources survive export, import, fixed-topology content
edits, second export, and second import through the Microsoft Open XML SDK.

Imported source order, GUIDs, tags, citation tags, and field topology remain
source-bound. Other contributor roles, complex field switches/results,
multiple or irregular bibliography parts, bibliography output fields, and
extension graphs stay opaque/read-only and fail closed on semantic mutation.
The package inspector recognizes SDK-authored package-root bibliography
relationship targets without weakening OPC path validation. The shipped
Documents example and fixture author a source and citation, edit the imported
title/author/display result, resolve the final semantic objects, assert native
bibliography and field markup, verify the model, and complete two OpenChestnut
round trips.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, serial `npm run verify:open-chestnut-build`, OpenChestnut
`180/180`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39
audited files and the same manifest-bound 38-file, 14,297,788-byte runtime. The
clean-install tarball contains 422 files, is 9,265,073 bytes compressed and
23,061,507 bytes unpacked. No npm publish, tag, or GitHub release operation was
attempted.

The completed public-workflow candidate at commit
`cec9628f8429934e48820bdbae064d5a2ff32fd5` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29580505711](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29580505711)
on 2026-07-17. The run completed with conclusion `success` in 4m11s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the canonical bibliography/citation Documents
workflow and fail-closed irregular-graph regressions, generated API-doc
cleanliness, offline release metadata, the 422-file clean-install tarball,
OfficeBridge `5/5`, and OpenChestnut `180/180`.

### DOCX inline plain-text content-control public workflow

On 2026-07-17, the already integrated inline plain-text content-control codec
was completed as a truthful public Documents workflow. The generated Help/API
surface now documents `paragraph.addTextContentControl(...)`,
`document.contentControls`, and transactional
`document.fillContentControls(...)`; the coverage and architecture records no
longer describe every content control as unsupported. The native Documents
Skill routes source-free and recognized imported body-inline plain-text SDTs
through the public JavaScript model and OpenChestnut, while retaining the
Python package helper only for explicit existing-template wrapping,
headers/footers, and controls outside the bounded model.

The shipped `openchestnut-end-to-end.mjs` example now authors an `OWNER`
control, exports canonical `w:sdt` markup, imports it, fills it by tag, exports
again, and proves the final tag/alias/text through semantic import, inspect,
native XML, and verification assertions. Unknown tags fail before mutation by
default; duplicate tags fill every match. Imported text/tag/alias edits are
permitted only with fixed run topology and native identity. Rich, block, cell,
nested, data-bound, dropdown, date, checkbox, placeholder-document, locked,
header/footer, and extension-bearing controls remain opaque/source-bound and
fail closed rather than being flattened.

The complete local gate passed `npm test` including Playwright and all five
published Skills, `npm run docs:api`, `npm run proto:check`,
`npm run test:pack`, offline `npm run release:check -- --skip-network
--allow-dirty`, serial `npm run verify:open-chestnut-build`, OpenChestnut
`179/179`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39
audited files and the same manifest-bound 38-file, 14,250,684-byte runtime. The
clean-install tarball contains 422 files, is 9,245,399 bytes compressed and
23,000,796 bytes unpacked. No npm publish, tag, or GitHub release operation was
attempted.

The completed public-workflow candidate at commit
`6231544c88dd5eee461c2b3e4950425807c8add9` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29575996251](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29575996251)
on 2026-07-17. The run completed with conclusion `success` in 4m30s and covered
protocol/runtime determinism, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite including the published Documents Skill example and native
content-control assertions, generated API-doc cleanliness, offline release
metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and
OpenChestnut `179/179`.

### XLSX direct threaded-comment replies

On 2026-07-17, the Spreadsheet public model, versioned protobuf wire, OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native Spreadsheet Skill converged on the reference workflow's `thread.addReply()` primitive. One root may now carry multiple direct replies with independent native comment/person GUIDs, display/user/provider identity, ISO date, done state, semantic inspect evidence, and fixed-cell parent binding. Source-free replies author native Office 2019 `parentId`; canonical imports expose the replies for text/state edits and a second export/import; unchanged imported threaded-comment parts remain byte-preserved.

The bounded topology is deliberately loss-aware: a reply may point only to its thread root and must target the same cell. Reply-of-reply or branched graphs, orphan/self/cross-cell parents, mentions, invalid GUID/person graphs, and other extensions remain opaque/source-bound; attempts to replace them fail closed. The shipped Spreadsheet fixture now exercises two authors, one direct reply, native XML evidence, semantic import, and the public Skill workflow instead of documenting a model-only method that the codec rejects.

The complete local gate passed `npm test` including Playwright, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, serial `npm run verify:open-chestnut-build`, OpenChestnut `178/178`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,238,396-byte runtime. The clean-install tarball contains 422 files, is 9,234,762 bytes compressed and 22,965,796 bytes unpacked. No npm publish or tag operation was attempted.

The complete direct-reply candidate at commit `f69b5b50ffb6ace73d0e6de53db9c3bb30f7571b` passed the hosted Linux `ci` workflow in [GitHub Actions run 29573634244](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29573634244) on 2026-07-17. The run completed with conclusion `success` in 5m20s and covered deterministic protocol/runtime verification, Chromium/LibreOffice/Poppler tools, the full npm suite including the native direct-reply Skill fixture and topology gates, generated API-doc cleanliness, offline release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `178/178`.

### Reference 2.8.24 public API completion

On 2026-07-17, a public-surface comparison between the previously pinned 2.8.22 package and the synchronized 2.8.24 reference found five additional named exports: `setOfficeFontDesignMetrics`, `registerScopedOfficeFontDesignMetrics`, `resolveOfficeFontDesignMetrics`, `clearOfficeFontDesignMetrics`, and `skiaPaintBaselineCompensationPx`. The clean-room implementation independently provides replacement and scoped metric registries, defensive normalized results, primary-family and exact-style selection, nearest numeric weight selection, idempotent scope disposal, invalid-record filtering, and deterministic finite baseline compensation. It also exposes fresh, sorted, case-insensitively deduplicated `fontFamilies` inventories on Document, Workbook, and Presentation models.

The same observable audit found the read-only `Shape.useBackgroundFill` getter. OpenChestnut now imports presence-aware native `p:sp/@useBgFill`, projects it through protobuf and the JavaScript model, uses it for preview paint, and preserves an unchanged source-bound slide XML byte-for-byte. Source-free authoring and semantic mutation fail closed instead of manufacturing or flattening native state. The native Documents, Spreadsheets, and Presentations Skill guides document the new inventories and imported-shape boundary; generated API documentation contains all nine new Help records.

The complete local gate passed `npm test` including Playwright, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, serial `npm run verify:open-chestnut-build`, OpenChestnut `178/178`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,232,252-byte runtime. The clean-install tarball contains 422 files, is 9,229,500 bytes compressed and 22,954,466 bytes unpacked. `npm run release:check` passed every code, package, license, documentation, JavaScript, and .NET gate; before this evidence commit it reported only the intentionally dirty generated documentation plus the external npm-authentication blocker. No publish or tag operation was attempted.

The completed reference 2.8.24 API-sync candidate at commit `a81d71823a3366805d9cb3a44c322bced386f2c5` passed the hosted Linux `ci` workflow in [GitHub Actions run 29571987958](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29571987958) on 2026-07-17. The run completed with conclusion `success` in 4m07s and covered protocol/runtime verification, Chromium/LibreOffice/Poppler tool checks, the full npm suite, generated API-doc cleanliness, offline release metadata, the 422-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `178/178`.

### Reference 2.8.24 Skill and Presentation-view convergence

On 2026-07-17, the pinned `reference/office-artifact-tool` submodule advanced from `2.8.22` to `2.8.24` at `2d0e249ea6b62f55cca22a343b832a38e8f7537c`. The published Spreadsheet Skill now uses the reference-native `artifact_tool_docs/API_QUICK_START.md` and `features/charts.md` paths, with package and workflow regressions proving the old locations are absent. The Presentation Skill gained the reference `presentation.view` workflow without importing the private runtime: local gridline/guide visibility matches the observable API, while a new independent protobuf/Open XML SDK slice imports grid spacing, snap flags, and bounded horizontal/vertical guides from `ppt/viewProps.xml`.

Imported guide definitions are returned as defensive copies by `presentation.view.toProto()` and exposed as frozen Master/Layout `slideGuides` projections. Unchanged source-bound export preserves the native view-properties part byte-for-byte; semantic mutation, source-binding mismatch, source-free wire authoring, invalid guide values, and more than 1,024 modeled guides fail closed. The combined release also completed the pending DOCX note edge corpus and public Documents workflow: rich multi-paragraph note bodies remain opaque/source-bound and byte-preserved when unchanged instead of being flattened into the bounded plain-text note model.

The complete local gate passed `npm test`, `npm run docs:api`, clean-HEAD `npm run proto:check`, `npm run test:pack`, serial `npm run verify:open-chestnut-build`, OpenChestnut `177/177`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,230,716-byte runtime. The clean-install tarball contains 421 files, is 9,220,694 bytes compressed and 22,935,038 bytes unpacked. The generic Skill quick validator accepts Documents but rejects the reference package's intentionally preserved `name: Spreadsheets` and `name: Presentations` capitalization; the project-native four-plugin validator and complete reference workflow gate pass those published discovery contracts.

### DOCX whole-block bookmark vertical slice

On 2026-07-17, the Documents public model, versioned protobuf wire, OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native Documents plugin converged on one reversible bookmark profile. A source-free bookmark may wrap exactly one paragraph-like block and supplies a native target for internal `w:hyperlink` anchors. Export writes paired `w:bookmarkStart`/`w:bookmarkEnd`; import exposes name, target, native identity, source position, and semantic binding. Recognized imported whole-block bookmarks remain fixed-topology and read-only, while cross-block, nested, crossing, table-cell, and otherwise irregular graphs stay opaque-preserved and fail closed on mutation.

The shipped Documents example now authors a `DecisionSection` bookmark plus internal jump link, completes two OpenChestnut round trips, resolves the imported target, verifies the final model, and asserts the native bookmark and hyperlink markup in `word/document.xml`. The broader Documents smoke continues to require semantic verification and native LibreOffice/Poppler page rendering. C# and JavaScript regressions cover source-free authoring, internal-link planning before bookmark insertion, unchanged source-preserving export, imported rename rejection, topology-change rejection, and cross-block authoring rejection.

The complete local gate passed `npm test`, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, serial `npm run verify:open-chestnut-build`, OpenChestnut `174/174`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,184,636-byte runtime. The clean-install tarball contains 421 files, is 9,202,409 bytes compressed and 22,856,201 bytes unpacked. The unpacked-size regression budget moved narrowly from 22,850,000 to 22,900,000 bytes to account for the measured codec/runtime addition while retaining less than 44 KiB of headroom.

### Composition root and Spreadsheet domain layering

On 2026-07-17, the remaining stateful Spreadsheet domain moved from the public root entry into `src/spreadsheet/index.mjs`: workbook/worksheet/range ownership, tables/pivots/charts/images/sparklines, formulas and dependency graphs, conditional formatting, inspect/resolve/verify/help, layout/SVG rendering, delimited-file support, metadata restoration, and the `SpreadsheetFile` facade now share one cohesive owner. The root retains exact re-exports rather than wrappers; all four public Spreadsheet bindings are strict-identical between the root and leaf entry. The OpenChestnut Spreadsheet adapter imports the leaf directly, and a source-level regression prevents a back-edge to `src/index.mjs`.

This completed the public-entry decomposition: `src/index.mjs` is now a 42-line composition root, or 0.1% of the 28,109 JavaScript source lines, while preserving all 36 public exports and the cross-format verify/render/visual-QA/help surface. The complete local gate passed `npm test`, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, `npm run verify:open-chestnut-build`, OpenChestnut `173/173`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,166,204-byte runtime. The clean-install tarball contains 421 files, is 9,196,227 bytes compressed and 22,828,311 bytes unpacked.

### Presentation domain layering

On 2026-07-17, the complete JavaScript Presentation domain moved from the public root entry into `src/presentation/index.mjs`: model collections, themes/masters/layouts/placeholders, shapes/connectors/groups/images/tables/charts/native objects, inspect/resolve/verify/help, layout/SVG rendering, and the `PresentationFile` facade now share one cohesive owner. The root retains exact re-exports rather than wrappers; all eight public Presentation bindings are strict-identical between the root and leaf entry. The OpenChestnut Presentation adapter imports the leaf directly, and a source-level regression prevents a back-edge to `src/index.mjs`.

The extraction reduced `src/index.mjs` from 5,694 to 4,305 lines, or 15.3% of the 28,107 JavaScript source lines, while preserving all 36 public root exports. The complete local gate passed `npm test`, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, `npm run verify:open-chestnut-build`, OpenChestnut `173/173`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,166,204-byte runtime. The clean-install tarball contains 420 files, is 9,196,256 bytes compressed and 22,828,153 bytes unpacked.

### DOCX whole-paragraph tracked-change vertical slice

On 2026-07-17, the Documents public model, versioned protobuf wire, OpenChestnut C# codec, bundled WASM runtime, Help catalog, and native plugin workflow converged on one bounded tracked-change profile. `document.addInsertion(...)` and `document.addDeletion(...)` author native whole-paragraph `w:ins`/`w:del` markup containing one text run; import exposes type, text, author, and optional timestamp; source-preserving export permits fixed-topology text/author/date edits while retaining native revision identity and unmodeled formatting. In-paragraph replacements, mixed or nested revisions, moves, property changes, and automatic future-change tracking remain explicit advanced workflows.

The edge-case corpus proves that a multi-run revision is still semantically visible but marked read-only, carries no editable-topology residual hash, round-trips with byte-identical source-package output when unchanged, and rejects semantic mutation with `unsupported_document_edit`. This closes the irregular-topology exception path without weakening fail-closed preservation.

The complete local gate passed `npm test`, `npm run docs:api`, `npm run proto:check`, `npm run test:pack`, `npm run verify:open-chestnut-build`, OpenChestnut `173/173`, and OfficeBridge `5/5`. Two clean WASM builds produced the same 39 audited files and the same manifest-bound 38-file, 14,166,204-byte runtime. The clean-install tarball contains 419 files, is 9,196,284 bytes compressed and 22,827,449 bytes unpacked. The shipped Documents example completed create, export, import, tracked-change edit, second export/import, semantic/XML assertions, and LibreOffice/Poppler rendering; both rendered pages were reviewed, with the insertion and deletion visibly represented and no clipping or overlap.

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
- The explicit real-provider gate ran with ReportLab 4.4.9, pdfplumber 0.11.9, pypdf 6.10.0, and separately installed PyMuPDF 1.27.2.3 under the approved AGPL route. It proved byte-identical incremental prefixes, non-prefix rewrites, pypdf text/radio/checkbox value and appearance handling, complete-source page selection/reorder/selective watermarking with preserved navigation, pypdf and PyMuPDF annotation operations, bounded text/image/page edits, full redaction/scrub, single-revision/no-residue output, and deliberate reflow rejection. The typed Poppler comparator maps each merged output page back to its immutable source, requires exact pixels on unstamped pages, requires bounded change on stamped pages, and reports blank-state and dark-ratio evidence. The scrub-only active-content fixture proves removal of root/additional JavaScript, Launch/SubmitForm actions, attachments, invisible text, comments, populated form values, personal metadata, and the null active-content dictionary names that PyMuPDF can leave after logical deletion; unfamiliar object serialization and invisible text overlapping visible text fail closed. In that earlier milestone an image-bearing strict residue scan failed closed because Tesseract was not installed, and qpdf, pyHanko, veraPDF, pikepdf, and OCRmyPDF were not executed; later sections record the separately shipped and tested qpdf, pikepdf, pyHanko, veraPDF, and OCRmyPDF adapters.
- The shipped Documents example and an independent Skill forward test each completed two OpenChestnut round trips, semantic assertions, and a one-page LibreOffice render with every page inspected. The audit narrowed custom source-free table styles to `TableGrid` plus direct formatting, documented non-persistent model locators across imports, aligned header/footer distance to the codec's 720-twip default, and routed ordinary classic comments through `DocumentModel.addComment`.
- The shipped Spreadsheet Range example completed R1C1/block-write/formula-evidence/navigation/format/chart/verify/render and two OpenChestnut round trips. A separate forward test authored and visually reviewed a three-sheet operating forecast with formula-driven financials, zero spreadsheet errors, PASS model checks, and a formula-bound line chart. Its findings led to automatic native formulas for `containsText`, text-preserving direct references, and live internal-range cache resolution for formula-only chart inspect/render/export; imported chart persistence snapshots remain deliberately separate and fail closed.
- The shipped Spreadsheet sparkline example and development fixture authored standard Office 2010 `x14:sparklineGroups` line, column, and stacked profiles, exercised vertical row and horizontal column mappings, imported and edited recognized groups with fixed topology, preserved unsupported non-contiguous native groups unchanged, and rejected lossy topology changes. LibreOfficeDev opened the fixture and Poppler rendered all three types across two pages; the JavaScript SVG preview rendered each target cell independently.
- The repository-only Agent PromptBench defines 32 black-box cases (20 PDF and 12 Office; 12 ready and 20 asset-required). The eight ready PDF routes include a source-bound native Text Highlight workflow: the Agent must use one unique MuPDF text selection with the exact inspected source hash/page snapshot, then pypdf independently verifies the saved Highlight's RGB/review metadata and text-bound quadrilaterals while Poppler/Pillow requires a target-text-contained diff and pixel-stable non-target pages. The grader rejects caller coordinates, direct native/object mutation, skipped re-inspection/render, stale page evidence, and unbound audit provenance. A clean candidate run at `535d875504f84ccd469cb05922ce94528cfd14d8` passed `100/100` with all hard gates plus 6 machine, 3 visual, 3 security, and 8 trace checks; it is one run, not the default three-repeat matrix or a reference-Skill comparison. The ready Office set includes two generated XLSX workflows: a direct-threaded-comment reply/resolve transaction and a formula-assumption transaction that changes only `Forecast!B9`, recalculates dependent cached values, protects `B10` plus the second-sheet baseline canary, permits only the target worksheet and workbook metadata parts to change, then checks native target-page change/baseline-page stability. It also includes the generated DOCX classic-comment and PPTX title/plain-text-notes workflows. Every ready Office route binds byte-bound audit/trace evidence, second import, and final LibreOffice/Poppler rendering. Existing clean XLSX comment (`d558a924ad63528f2b2dca5e1bbeb1fb0dc120a7`), DOCX (`b06f6ca067666b5774bb81fb23155f1cec50e694`), and PPTX (`450017eb8acb209f6ceb161d247f6c8059ab2571`) candidate trials each passed `100/100`; none is a repeat matrix or reference-Skill comparison. The clean formula-slice candidate at `3df67f5a083758eb1f6fc0c37cdc6c53f228e2eb` likewise passed `100/100` across 5 machine, 2 visual, 2 security, and 6 trace checks, binding tarball SHA-256 `720400c754d9cdbfaa949120c7939467702202d2cf507f7bef91d5068c1a7503`; it is one candidate trial, not a repeat matrix or reference-Skill comparison. A candidate/reference PDF preparation uses byte-identical package tarballs and prompts while varying only the copied Skill. Generic grading covers fail-closed branching, immutable file/directory inputs, read-only prompt/Skill/dependency trees, regular output types, and exact deliverables; the 20 corpus/PKI cases remain explicitly asset-required.
- PromptBench now has 33 cases (20 PDF, 13 Office) and 13 deterministic ready slices. The newest Presentation case performs the non-visual source-bound `p:cSld/@name` rename in the canonical two-slide package: the independent grader proves the requested name and fixed ordering/semantics, allows only the target SlidePart to differ, requires every non-target part byte-identical, source-part/attribute audit binding, second import, and pixel-identical LibreOffice/Poppler pages. It explicitly permits Open XML SDK canonicalization of the changed target XML. One clean candidate run at `ffdaca79014afd82038dfcbf0002dcaacb51c54d` passed `100/100` with all hard gates plus 5 machine, 2 visual, 2 security, and 6 trace checks; its tarball/Skill/input/oracle fingerprints are recorded in [Agent black-box evaluations](agent-evals.md). It is one trial, not a repeat matrix or a reference-Skill comparison.
- Current PromptBench inventory supersedes those historical counts: 35 cases (21 PDF, 14 Office), 14 deterministic ready cases, and 21 asset-required cases. The ready `pptx-closed-leaf-slide-clone` fixture now drives the shipped duplicate workflow with explicit `--allow-closed-leaves` over one closed literal-data ChartPart plus canonical NotesSlide and legacy SlideComments leaves. Its independent OPC oracle permits only the adjacent clone and exact cloned parts, requires a distinct byte-identical ChartPart with the same slide-local relationship ID/type and no child graph, checks retained-source byte hashes and shared NotesMaster/CommentAuthors catalogs, requires second import, and compares source/clone/appendix LibreOffice/Poppler pixels. One clean autonomous candidate Agent run at `0e8824cb3dac8332ff631b0cb75850ebe2a56f6b` passed every hard gate and scored `100/100` across 6 machine, 2 visual, 2 security, and 7 trace checks; its tarball/Skill/prompt/input/oracle fingerprints are recorded in [Agent black-box evaluations](agent-evals.md). This is one trial, not the default repeat or reference-Skill matrix. The added asset-required `pdf-encrypted-owner-policy-boundary` is deliberately fail-closed: user-password-only input must not trigger owner-password bypass, silent decryption, or a replacement unencrypted PDF.
- The bounded contract-ID defect was closed through the official typed path: `replace_text` now preserves the source baseline/default style and accepts the observed `0.0000227pt` provider quantization difference inside a non-configurable tolerance capped at `0.0005pt`, while genuine overflow and rotated text remain fail-closed. The Skill now requires provider probe and route planning before mutation and ships the canonical `open-office-artifact-tool.pdf-audit.v1` schema plus a validator that recomputes source/output hashes. One fixed candidate trial and one same-prompt, same-tarball reference-Skill trial each scored 100/100 across machine, visual, security, and trace evidence. Repeat trials and the remaining case graders/corpus are still open.
- The active-content public-sanitize fixed matrix passes candidate `3/3` and reference Skill `3/3`, all at 100/100 across machine, visual, security, and trace. The six runs bind clean commit `39fa301dcb1005f2848282e6e63da1e934104821`, byte-identical package SHA-256 `e78e18c0f8f1cffe301ae1f2ea17e882bc879b3044914033e24b0b11ac0e8b69`, identical prompt/input/oracle fingerprints, and fixed but distinct candidate/reference Skill fingerprints. Independent pypdf evidence proves root/additional JavaScript, Launch/SubmitForm, five attachments, invisible text, a comment, a populated widget, and personal metadata in each immutable source and zero active structure/canary residue in every output; Poppler changes stay inside the expected form/comment masks. Historical bypass and interpreter-drift runs remain defect evidence rather than passes. The evaluator ignores help-only invocations and permits independently scheduled probe/plan completion only when both precede the real typed edit; low-level bypass and post-mutation gates remain strict.
- The AcroForm adapter now distinguishes text/choice strings from radio/checkbox PDF Name states and validates post-write field/widget appearances before transaction promotion. Its complete generated fixture contains five text widgets, two radio widgets, and one pre-checked checkbox. The clean fixed matrix binds commit `bffd35dbfdb94bb1183717703e7e55bfb83c3f3c`, package SHA-256 `7ab9e6a30035df5d0ef7ee9990f3a0445152877e58a7f0d065ede9ddc1db300b`, and identical prompt/input/oracle fingerprints. Candidate passes `3/3` at 100/100; reference passes `2/3` at 100/100. The remaining reference run earned 70 raw machine/visual points but correctly scored zero after manually writing incremental PDF objects, reporting a non-pypdf provider, omitting canonical audit provenance, and bypassing provider preflight, typed fill, and post-mutation audit gates. All semantically successful fixed runs preserve TIN/signature/unselected-radio/checkbox pixels and editability, preserve the exact source prefix with one appended revision, and render every page through Poppler.
- The pypdf attachment-quarantine primitive covers document/page FileSpecs, duplicate and Unicode names, traversal-safe cross-platform naming, decoded-byte budgets, transactional extraction, raw identity/MIME/size/SHA evidence, and source/manifest provenance without opening payloads. Its fixed matrix binds clean commit `748fbb1d81ccfa14a594d6fed9bc6601866bfa95`, package SHA-256 `9cff93494c5b32e16394ce3b4fcffa1daf76ad6df57326dab0ced47d2a45b5bf`, and identical prompt/input/oracle fingerprints. Candidate passes `3/3` at 100/100; reference passes `2/3` at 100/100. The retained reference miss extracted the correct six payloads safely but used a custom Node parser and alternate manifest/audit contract, so missing pypdf/read-only/no-fallback preflight, typed extraction, canonical schema, and byte-binding evidence correctly forced the run to zero.
- The six-page greenfield tagged-accessibility workflow now has a clean fixed matrix: candidate `3/3` and reference Skill `3/3`, every run at 100/100 across machine, visual, security, and trace. All six records bind clean commit `2323a70331b93781dee37aa05198e4a73a7ec533`, byte-identical package SHA-256 `cfbcf5c76ba5fdb929dae27f2a0295d6da12694eec2150a926cebeecedefccb9`, identical prompt/input/oracle fingerprints, and fixed but distinct candidate/reference Skill fingerprints. Independent pypdf traversal proves title/language, H1-H3, one logical Table across pages 3-4, Figure alt, Link/StructParent/OBJR, reading order, and 12 running artifacts; Poppler/Pillow proves every page is nonblank, unclipped, and contains the expected physical table segments. The workflow separates modeled checks, optional veraPDF machine evidence, and required human PDF/UA judgment rather than claiming automatic conformance.
- The merge/reorder/selective-watermark clean fixed matrix passes candidate `3/3` and reference Skill `3/3`, every run at 100/100 across machine, visual, security, and trace. All six runs follow provider check/inspect/plan, the typed pypdf manifest primitive, typed Poppler comparison, output inspection, and multi-source canonical audit, and bind commit `90cbb9e0a5527a4620a28bb38aad8feeca895a3b`, byte-identical package SHA-256 `c3962993aee732c7a8e60282159409dfe08ca2c7c4f0dd59eb80468c28630ff5`, identical prompt/input/oracle fingerprints, and fixed but distinct candidate/reference Skill fingerprints. The comparator closed a real QA-discipline gap: a prior Agent deleted a correct output after subjectively misreading two thumbnails as black even though their renders were more than 98.7% white and byte-identical to successful trials.
- LibreOffice opened the shipped 26-slide reference template and produced a 26-page PDF; bounded custom-geometry icons rendered visibly. This local LibreOffice build substituted `Helvetica Neue`, so pixel parity with the checked-in preview images is not claimed and remains a visual-fidelity gap.

## Hosted evidence

The repository-structure convergence candidate at commit
`b11a249b2bcd2397b5355c4986a2cd08f0aee39a` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29675702140](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29675702140)
on 2026-07-19. The run covered npm installation, deterministic protocol and
bundled OpenChestnut runtime verification, Chromium/LibreOffice/Poppler tool
checks, the complete npm suite, generated API-doc cleanliness, offline release
metadata, the production clean-install tarball, and both OfficeBridge and
OpenChestnut .NET suites.

The bounded native XLSX PivotTable candidate at commit
`4aac21752d502ed8b37c6574d205c8f06679d805` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29666171276](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29666171276)
on 2026-07-19. The job completed with conclusion `success` in 10m07s and
covered npm installation, deterministic protocol/OpenChestnut runtime
verification, Chromium/LibreOffice/Poppler tool checks, the complete npm suite
including the native two-page PivotTable Skill render, generated API-doc
cleanliness, offline release metadata, the 452-file clean-install tarball,
OfficeBridge `5/5`, and OpenChestnut `277/277`.

The canonical Default Template Library reference-pin candidate at commit
`2424ea03d28a7a5a9f2c51695a6113b6d219da78` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29656750965](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29656750965)
on 2026-07-19. The run completed with conclusion `success` in 12m14s. It
recursively checked out public `office-artifact-tool` commit `256cb31`, compared
all 40 retained Office/PNG assets byte-for-byte, and then passed deterministic
protocol/runtime verification, Chromium/LibreOffice/Poppler tool checks, the
complete npm suite and native template corpus, generated API-doc cleanliness,
offline release metadata, the 447-file clean-install tarball, OfficeBridge
`5/5`, and OpenChestnut `260/260`.

The bounded imported SlidePart placeholder-text candidate at commit
`f37d8adbcd5d0637efcbf6fcf10ef9821bfb4216` passed the hosted Linux `ci`
workflow in [GitHub Actions run 29655361922](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29655361922)
on 2026-07-19. The run completed with conclusion `success` in 13m22s and
covered npm installation, generated protocol plus deterministic bundled-runtime
verification, Chromium/LibreOffice/Poppler tool checks, the complete npm suite
including all seven retained PPTX placeholder edits and native renders,
generated API-doc cleanliness, offline release metadata, the 447-file
clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `260/260`.

The reference 2.8.24 Skill sync, Presentation view-properties codec, and completed DOCX note workflow candidate at commit `5cf2eb629f4612314d34c73041a5f1296a9ff145` passed the hosted Linux `ci` workflow in [GitHub Actions run 29569265350](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29569265350) on 2026-07-17. The run completed with conclusion `success` in 4m15s and covered generated protocol plus deterministic bundled-runtime verification, Chromium/LibreOffice/Poppler tool checks, the complete npm suite including the native Spreadsheet path migration, Presentation `view/grid/guides` import/export/fail-closed regressions, Documents note workflow, generated API-doc cleanliness, offline release metadata, the 421-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `177/177`.

The DOCX whole-block bookmark and Documents navigation candidate through commit `b161676d1e68df8acdc7bd3a4a3640fbcf267b3c` passed the hosted Linux `ci` workflow in [GitHub Actions run 29564975672](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29564975672) on 2026-07-17. The run completed with conclusion `success` in 5m09s and covered generated protocol plus deterministic bundled-runtime verification, Chromium/native-tool checks, the complete npm suite including the public bookmark/internal-link Skill workflow and fail-closed topology regressions, generated API-doc cleanliness, offline release metadata, the 421-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `174/174`.

The composition-root and complete Spreadsheet-domain layering candidate at commit `f8a4241a40448cc21ad1dce449edb6747d491473` passed the hosted Linux `ci` workflow in [GitHub Actions run 29563335625](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29563335625) on 2026-07-17. The run completed with conclusion `success` in 4m26s and covered protocol plus deterministic bundled-runtime verification, Chromium/native-tool checks, the complete npm suite including root/leaf Spreadsheet binding identity and codec dependency-direction regressions, generated API-doc cleanliness, offline release metadata, the 421-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `173/173`.

The complete Presentation-domain layering candidate at commit `f8d50cad14ae74ee6cef7610ba58a52ea0438514` passed the hosted Linux `ci` workflow in [GitHub Actions run 29562292218](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29562292218) on 2026-07-17. The run completed with conclusion `success` in 4m18s and covered protocol plus deterministic bundled-runtime verification, Chromium/native-tool checks, the complete npm suite including root/leaf Presentation binding identity and codec dependency-direction regressions, generated API-doc cleanliness, offline release metadata, the 420-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `173/173`.

The DOCX whole-paragraph tracked-change and Documents Skill candidate through commit `98dbdd7e8a0814f79830efb5e8a7e2675dfcf520` passed the hosted Linux `ci` workflow in [GitHub Actions run 29561312029](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29561312029) on 2026-07-17. The run completed with conclusion `success` in 4m11s and covered generated protocol plus deterministic bundled-runtime verification, Chromium/native-tool checks, the complete npm suite including the public tracked-change Skill example and irregular-topology regression, generated API-doc cleanliness, offline release metadata, the 419-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `173/173`.

The AGPL/MuPDF candidate at commit `40fcee0931c541c6f2cb1639ead0d10e2b76c7e6` passed the hosted Linux `ci` workflow in [GitHub Actions run 29558072285](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29558072285) on 2026-07-17. The run completed with conclusion `success` in 4m52s and covered npm install, protocol/runtime verification, Chromium/native-tool checks, the full npm suite, generated API-doc cleanliness, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `170/170`.

The embedded-XLSX/OLE and Presentation native-object layering candidate at commit `6119c54ae05d4b60fe562641e7aef10130581782` passed the hosted Linux `ci` workflow in [GitHub Actions run 29558401718](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29558401718) on 2026-07-17. The run completed with conclusion `success` in 4m22s and covered protocol/runtime verification, Chromium/native-tool checks, the full npm suite including payload-only OLE replacement and Presentation Skill regressions, generated API-doc cleanliness, offline release metadata, the 417-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `170/170`.

The Document-domain and shared text-range layering candidate at commit `765040c5827ee73c3c4645824688470e281c8be5` passed the hosted Linux `ci` workflow in [GitHub Actions run 29559403816](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29559403816) on 2026-07-17. The run completed with conclusion `success` in 5m22s and covered protocol/runtime verification, Chromium/native-tool checks, the full npm suite including root/leaf binding identity and native Documents workflow regressions, generated API-doc cleanliness, offline release metadata, the 419-file clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `170/170`.

The type-aware AcroForm candidate at commit `bffd35dbfdb94bb1183717703e7e55bfb83c3f3c` passed the hosted Linux `ci` workflow in [GitHub Actions run 29515214432](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29515214432) on 2026-07-16. The run completed with conclusion `success` and covered the full npm suite including AcroForm provider and independent grader regressions, deterministic OpenChestnut verification, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167`.

The active-content PDF sanitization candidate at commit `099bf1ce2f62ab992971e61b82641e6d6712a95d` passed the hosted Linux `ci` workflow in [GitHub Actions run 29501149008](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29501149008) on 2026-07-16. The run completed with conclusion `success` in 4m01s and covered deterministic OpenChestnut verification, Chromium/LibreOffice/Poppler tool checks, the full npm suite including the typed sanitize provider and independent grader regressions, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167`.

The Agent PromptBench scaffold at commit `70b2ddeea642de3729f8b7d7401bf10bace3be69` passed the hosted Linux `ci` workflow in [GitHub Actions run 29493964234](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29493964234) on 2026-07-16. The run completed with conclusion `success` and covered suite validation/tests inside the full npm gate, deterministic OpenChestnut verification, Chromium/LibreOffice/Poppler checks, API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167`.

The OpenChestnut XLSX sparkline candidate at commit `e8aa3e14249de346207f16b8fa24d7cb00b1253f` passed the hosted Linux `ci` workflow in [GitHub Actions run 29492891825](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29492891825) on 2026-07-16. The run completed with conclusion `success` in 4m02s and covered deterministic protocol/runtime verification, Chromium/LibreOffice/Poppler tool checks, the full npm suite, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge `5/5`, and OpenChestnut `167/167` including standard Office 2010 sparkline coverage.

The PDF provider-routing candidate at commit `b405ddd249c7c2f760c659c07e88495f3a3562f3` passed the hosted Linux `ci` workflow in [GitHub Actions run 29487829878](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29487829878) on 2026-07-16. The run completed with conclusion `success` in 3m59s and covered deterministic protocol/runtime verification, Chromium/LibreOffice/Poppler tool checks, the full npm suite including the provider contract tests, generated API-doc diff, offline release metadata, clean-install tarball, OfficeBridge, and OpenChestnut 163-test execution. Optional Python providers remain an explicit local/environment gate rather than an undeclared hosted dependency.

The Documents native-workflow/OpenChestnut candidate through commit `e07e24382ff0259c7beefe27b0743d908a1f946f` passed the hosted Linux `ci` workflow in [GitHub Actions run 29483188346](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29483188346) on 2026-07-16. The run completed with conclusion `success` in 3m55s and covered protocol/runtime reproducibility, Chromium/native render tools, `npm test`, generated API-doc diff, offline release metadata, the registry-independent clean-install tarball, OfficeBridge, and OpenChestnut.

The bounded DOCX modern-comment candidate was locally closed on 2026-07-19. The public model/protobuf/OpenChestnut path now authors one root with direct replies through native `commentsExtended`, optional comments IDs/extensible/people parts, and supports source-bound imported text/resolved-state edits while refusing nested or irregular graphs. The shipped Documents workflow proves fixed identity/topology, second import, model/native rendering, source immutability, atomic output, and a byte-bound audit. `npm test`, generated API docs, deterministic OpenChestnut reconstruction, the production clean-install/package gate, OpenChestnut `271/271`, and OfficeBridge `5/5` passed. The audited package contains 448 files at 9,472,335 packed bytes and 23,866,257 unpacked bytes; its bundled OpenChestnut runtime contains 38 files at 14,489,788 bytes. The full PDF provider test remained contract-only because no explicit `OPEN_OFFICE_PDF_PROVIDER_PYTHON` was configured; core MuPDF.js, Playwright/Chromium, LibreOffice/Poppler, canonical template corpus, and all other npm gates ran locally.

### Bounded native XLSX PivotTable vertical slice

On 2026-07-19, the Spreadsheet public model, versioned protobuf wire,
OpenChestnut C# codec, Help catalog, and native Spreadsheet Skill converged on
one source-free PivotTable profile: exactly one row field, zero or one column
field, one `sum`/`count`/`average`/`min`/`max` value field, optional row/column
grand totals, explicit refresh/cache policy, derived worksheet cache cells, and
optional native cache records. OpenChestnut owns the workbook, worksheet,
PivotTable, cache-definition, cache-record, content-type, and relationship graph.
Recognized imports expose semantic inspect/resolve/preview but bind that graph,
configuration, source values, and cached output as read-only; richer or
unrecognized graphs remain opaque/source-bound and no static-table fallback is
used.

The complete local gate passed `npm test`, `npm run docs:api`, Buf lint plus a
byte-idempotent generated binding, `npm run test:pack`, serial
`npm run verify:open-chestnut-build`, OpenChestnut `277/277`, and OfficeBridge
`5/5`. The C# suite includes Office 2021 validation, saved and unsaved cache
records, exact second-export Pivot/cache part preservation, semantic-edit and
duplicate-name rejection, and unsupported-profile fail-closed coverage. The JS
suite covers all five aggregations, case-insensitive workbook-wide names,
pre-styled blank output cells, value/formula collisions, source/cache mutation,
and two imports. Two clean WASM builds produced the same 39 audited files and
the same manifest-bound 38-file, 14,607,552-byte runtime.

The production clean-install tarball contains 452 files, is 9,522,788 bytes
compressed and 24,053,767 bytes unpacked; the repository-only canonical Default
Template Library remains excluded. The unpacked-size ceiling moved narrowly
from 24,050,000 to 24,150,000 bytes after the audited payload exceeded the old
limit by 3,767 bytes, retaining less than 97 KB of headroom. The runnable Pivot
workflow passed inspect/verify/SVG review, native Open XML round trips, and a
real LibreOfficeDev 26.8.0.0.alpha0 → two-page PDF → Poppler PNG review: the Data
and Pivot Summary sheets each fit on one page with visible labels, currency
values, borders, and grand totals. That native gate caught and fixed an initial
column-width unit error that had split the summary across four pages. The
specialist Python PDF provider test remained contract-only because no explicit
`OPEN_OFFICE_PDF_PROVIDER_PYTHON` was configured; core MuPDF.js, Playwright,
canonical template rendering, and all other npm gates ran. `npm whoami` still
returns `ENEEDAUTH`, so no publish or tag operation was attempted.

### Bounded qpdf structural provider

On 2026-07-19, the PDF Skill gained a shipped thin adapter for separately
installed qpdf 11+ JSON v2. The adapter exposes only bounded read-only structural
inspection and exact-source-bound full-rewrite repair or linearization. It
records source identity, qpdf diagnostics, bounded topology counts,
encryption/linearization state, and object-level signature, ByteRange, Perms,
DocMDP, and FieldMDP indicators without claiming signature trust. Rewrites use
a private source snapshot, reject encrypted input, require explicit signature
invalidation, re-inspect a clean output, preserve page/form/attachment/outline
counts, re-prove the source hash, and publish a distinct nonexisting output
without replacement. The adapter exposes no arbitrary qpdf flags and is not a
sanitizer, redactor, renderer, text extractor, strict conformance validator, or
password/signature authority.

The hermetic provider suite proved qpdf 10 rejection, registry routing, stale
hash and symlink refusal, missing-provider no-fallback behavior, hard JSON
output budgets, malformed JSON rejection, and timeout termination. With local
qpdf 12.3.2 it also inspected and rewrote a generated PDF, recovered a broken
cross-reference table, preserved an embedded attachment, detected a synthetic
signature/ByteRange/DocMDP graph, blocked a signed rewrite by default, and
required explicit invalidation. Poppler 26.05.0 rendered the source, repaired,
and linearized files to byte-identical PNGs.

The complete local gate passed `npm test`, `npm run docs:api` with no generated
diff, `npm run proto:check`, `npm run test:pack`,
`npm run verify:open-chestnut-build`, `npm run release:check`, OfficeBridge
`5/5`, and OpenChestnut `283/283`. Two deterministic WASM builds matched across
39 files; the bundled runtime contains 38 files at 14,635,200 bytes. The
production clean-install tarball contains 454 files, is 9,546,746 bytes
compressed and 24,123,780 bytes unpacked. The full npm suite also used
LibreOfficeDev 26.8.0.0.alpha0, Poppler 26.05.0, the installed Playwright
Chromium runtime, and the real qpdf adapter. `npm whoami` still returns
`ENEEDAUTH`; no publish or tag operation was attempted.

### Bounded pyHanko signature-validation provider

On 2026-07-19, the PDF Skill gained a shipped read-only adapter for an
explicitly installed pyHanko `>=0.35,<0.36` core runtime. It binds validation to
an exact source SHA-256, copies the PDF and caller-supplied certificate inputs
into private immutable snapshots, limits source/certificate/signature counts,
subprocess time, and output bytes, disables network fetching and implicit system
trust, then re-proves every source input. There is no CLI invocation, PDF
mutation, system-root guess, or provider fallback.

The versioned report keeps ByteRange integrity and cryptographic validity
separate from certificate-path trust, timestamps, revision coverage,
difference-analysis modification level, changed form fields, DocMDP/FieldMDP,
seed values, and pyHanko's selected-policy bottom line. Callers choose
`cryptographic-only` or explicit trust roots, validation time, one of four
revocation policies, and independent required gates. The adapter explicitly
does not claim complete PAdES profile conformance. Signature creation and
private-key/TSA/LTV operations remain external workflows; the later veraPDF
milestone below supplies the separate source-bound conformance adapter.

The real-provider test used pyHanko 0.35.2 and pyhanko-certvalidator 0.31.1 in
an isolated virtual environment.
It generated its own CA/key and exercised unsigned inventory plus
required-signature failure, one trusted certification signature,
cryptographic-only untrusted
evidence, two signatures in revision order, an allowed post-signing LTA metadata
revision, stale hashes, implicit-root refusal, and a byte-tampered signature.
Every validation preserved the source bytes; the tampered fixture failed the
required integrity gate. Hosted CI now creates the same isolated runtime and
pins both versions before the full npm suite.

The complete local gate passed `npm test`, `npm run docs:api`, `npm run
proto:check`, `npm run test:pack`, OfficeBridge `5/5`, OpenChestnut `283/283`,
and two deterministic OpenChestnut WASM builds across 39 audited files. The
bundled runtime remains 38 files at 14,635,200 bytes. The production
clean-install tarball contains 455 files, is 9,556,405 bytes compressed and
24,165,413 bytes unpacked. `npm whoami` remains unavailable, so no publish or
tag operation was attempted.

### Bounded veraPDF conformance-validation provider

On 2026-07-19, the PDF Skill gained a shipped read-only adapter for an
explicitly installed veraPDF `>=1.30,<1.31` CLI. It binds validation to an exact
final-file SHA-256, copies the PDF to a private read-only snapshot, requires one
explicit built-in PDF/A-1/2/3/4 or PDF/UA-1/2 profile, and exposes no automatic
profile selection, custom profiles, passwords, directories, arbitrary flags,
repair, or provider fallback. Input size, execution time, stdout/stderr, failed
rule evidence, and report cardinality are bounded; source and snapshot hashes
are re-proved after the provider exits.

The versioned report verifies the veraPDF component-version range, selected
profile, item size, one-job completion, exit/result consistency, rule/check
counts, and batch error counters. A noncompliant file is a completed validation
with `machineRuleCompliant: false`; `--require-compliant` separately turns that
fact into a delivery failure. PDF/UA results always retain a human-review gate
because machine rules do not establish author intent, semantic quality,
contrast, or actual assistive-technology usability.

The hermetic test suite covers old versions, missing profiles, stale hashes,
hard budgets, timeouts, malformed and contradictory JSON, wrong profiles,
provider-reported exceptions, private-snapshot mutation, and source
immutability. The real-provider gate uses veraPDF 1.30.2 plus the official
CC BY 4.0 veraPDF corpus PDF/A-1b pass fixture; the same bytes pass the explicit
PDF/A-1b gate and produce a bounded noncompliant PDF/UA-1 report. Hosted CI
installs Java 21 and only the veraPDF CLI pack from the pinned 1.30.2 installer,
whose SHA-256 is
`6cc6341cb1af644044054b81f00a6590a7918abb18f762243de115258bcad838`.

The production package now contains 456 files, is 9,565,610 bytes compressed
and 24,200,954 bytes unpacked on the local audit host. The repository-only
4,048-byte conformance fixture is excluded from the npm tarball. Full release
and hosted results are recorded after the final candidate commit; npm publish
remains blocked by unavailable authentication.

### Bounded OCRmyPDF searchable-layer provider

On 2026-07-19, the PDF Skill gained a shipped thin adapter for an explicitly
installed OCRmyPDF `>=17.8,<17.9` runtime, Tesseract 5.x, qpdf 11+, and Poppler
`pdftotext`. It binds one complete input PDF to an exact SHA-256, copies it to a
private read-only snapshot, and invokes OCRmyPDF with fixed standard-PDF,
optimization-zero, one-job, Tesseract, fpdf2, and pypdfium settings. The caller
must select `skip`, `redo`, or `force`, name installed OCR languages, and declare
either trusted input or caller-provided isolation. `redo` and `force`, Tagged
PDF input, force-mode forms or annotations, and signed input each require their
own loss or invalidation acknowledgement. Encrypted input is rejected.

The adapter accepts no arbitrary OCRmyPDF flags, page subsets, provider plugins,
or output replacement. It enforces source, output, sidecar, process, diagnostic,
text, language, and page-image budgets; terminates the provider process group on
timeout; rejects source-prefix output; re-proves the source and snapshot; and
publishes one distinct full-rewrite result without replacement. qpdf supplies
pre/post encryption, signature, tagging, form, annotation, and bounded topology
evidence. Poppler extracts the final searchable text for required-text and
nonempty gates. This route is not sanitize, redaction, OCR-quality proof, or a
malware sandbox; attacker-chosen files still require external process and host
isolation plus final visual review.

The hermetic suite covers missing and out-of-range providers, unavailable
languages, stale source hashes, symlinks and output collisions, trust and
structure acknowledgements, signed/encrypted policy, fixed provider arguments,
timeouts and hard output budgets, source-snapshot mutation, source-prefix
outputs, empty OCR text, and required-text failure. The real-provider gate used
OCRmyPDF 17.8.1, Tesseract 5.5.2, qpdf 12.3.2, and Poppler 26.05.0. It generated
an image-only PDF whose source exposed no text, recovered both required phrases,
re-imported the output through MuPDF.js, preserved the source bytes, and rendered
the source and OCR output to pixel-identical Poppler images at 96 DPI.

The complete local gate passed `npm test` including Playwright and the real OCR
fixture, `npm run docs:api` with no generated diff, `npm run proto:check`,
`npm run test:pack`, serial `npm run verify:open-chestnut-build`, OfficeBridge
`5/5`, and OpenChestnut `283/283`. Two clean WASM builds reproduced 39 audited
files and the same manifest-bound 38-file, 14,635,200-byte runtime. The
production clean-install tarball contains 458 files, is 9,581,955 bytes
compressed and 24,250,484 bytes unpacked on the local audit host. The same file
set measured 9,784,909 compressed bytes on hosted Linux; the cross-platform
packed ceiling moved narrowly to 9,810,000 bytes. The unpacked-size ceiling
moved to 24,275,000 bytes, retaining less than 24 KiB of measured headroom.
The standalone offline release metadata check passes; `npm whoami` still
returns `ENEEDAUTH`, so no publish or tag operation was attempted. Hosted CI
run 29683749916 passed the complete npm, generated-doc, release, package,
OfficeBridge, and OpenChestnut gates for that candidate.

### Bounded pikepdf active/auxiliary structure-clean provider

On 2026-07-19, the PDF Skill gained a shipped thin adapter for an explicitly
installed pikepdf `>=10.10,<10.11` runtime. It binds inspection and cleanup to
one exact source SHA-256, copies the PDF into a private read-only snapshot, and
requires either a trusted-input declaration or caller-provided process/host
isolation. Cleanup accepts only the fixed `active-content` or
`active-and-auxiliary` profile and always requires explicit signature
invalidation. There are no arbitrary pikepdf operations, provider flags,
passwords, parser recovery, incremental saves, output replacement, or silent
fallbacks.

The adapter uses pikepdf's curated sanitizer traversal to remove JavaScript,
external actions, and multimedia references, with an optional broader pass for
attachments/associated files, thumbnails, search indexes, Web Capture,
private application data, and portfolio presentation. It rejects encryption
and parser warnings; bounds input/output/process/diagnostic/page/object work;
re-proves the source and snapshot; preserves page, annotation, form, XFA,
metadata, tag, and outline topology; and publishes one non-prefix,
single-revision full rewrite without replacement. This operation is named
`structure-clean` deliberately: it does not redact page content, scrub
metadata, flatten forms or XFA, remove hidden/OCR text, validate signatures, or
provide a malware sandbox.

The real-provider fixture used Python 3.13.14 and pikepdf 10.10.0. It exercises
root and page JavaScript, Launch/URI/GoToR actions, RichMedia, an embedded and
associated file, a thumbnail, search and private page-piece data, Web Capture,
a portfolio collection, metadata, a populated form, XFA, annotations, tagging,
and an outline. Independent decoded-object assertions prove the selected
canaries disappear while the declared retained structures and visible content
remain. qpdf 12.3.2 re-opens the output and confirms a clean structure and zero
attachments for the broad profile; MuPDF.js performs a second import; Poppler
26.05.0 renders the source and cleaned output to pixel-identical pages. The
hermetic suite separately covers old/missing providers, stale hashes, missing
trust or signature acknowledgements, path collisions and symlinks, encryption,
timeouts, hard output/object/process budgets, source mutation, and
single-revision/non-prefix postconditions.

The complete local gate passed `npm test` including Playwright and the real
pikepdf fixture, `npm run docs:api` with no generated diff, `npm run
proto:check`, `npm run test:pack`, serial `npm run
verify:open-chestnut-build`, OfficeBridge `5/5`, and OpenChestnut `283/283`.
Two clean WASM builds reproduced 39 audited files and the same manifest-bound
38-file, 14,635,200-byte runtime. The production clean-install tarball contains
460 files, is 9,592,194 bytes compressed and 24,301,346 bytes unpacked on the
local audit host. The unpacked ceiling moved narrowly to 24,325,000 bytes,
retaining less than 24 KiB of measured headroom. Hosted results are recorded
after the candidate commit; npm authentication remains unavailable, so no
publish or tag operation was attempted.

### Source-bound local-PKCS#12 pyHanko signing provider

On 2026-07-19, the PDF Skill gained a second shipped pyHanko 0.35.x adapter for
one bounded local-PKCS#12 signature. `probe` records the supported core surface;
`inspect` inventories signature fields, signatures, certification state,
revisions, and optional page geometry on an immutable SHA-bound source; `sign`
adds one approval or first-document certification signature to an existing
field or a newly created invisible/visible field. Visible creation is limited
to an integer box wholly inside an unrotated inspected CropBox.

The transaction requires explicit trusted/caller-isolated input, exact source
and credential SHA-256 values, stdin-only or deliberate no-passphrase, an
expected signature count, explicit acknowledgement before countersigning, and
explicit DocMDP permission for certification. It rejects symlink credentials,
encryption, stale identities, unsupported field policy, source overwrite,
output collision, raised budgets, and secret-bearing argv/env options. Existing
signatures pass cryptographic/DocMDP preflight before the worker runs. Output
must preserve the complete source prefix, append one revision, add exactly one
signature, pass all-signature integrity and DocMDP validation, and publish
without replacement. Signing never establishes certificate trust; the separate
read-only validator remains the explicit-root delivery gate. TSA/LTV/DSS,
PKCS#11/HSM, remote signing, online revocation retrieval, and complete PAdES
conformance remain external.

The real-provider fixture used Python 3.13.14, pyHanko 0.35.2, and
pyhanko-certvalidator 0.31.1. It generated an isolated CA, signer certificate,
encrypted PKCS#12, source PDF, and existing empty signature field. Tests prove
a visible PAdES certification signature, invisible Adobe-detached
countersignature, existing-field fill, explicit-root trust, prior-revision
`form-filling`, exact prefixes, qpdf and MuPDF re-open, bounded Poppler
appearance pixels, invisible-signature pixel invariance, source/credential
immutability, secret non-disclosure and size limits, stale hashes, missing
DocMDP/field geometry, missing countersign acknowledgement, recertification
refusal, collision, symlink, output-budget, and encryption failures.

The complete local gate passed `npm test` with real qpdf, pikepdf 10.10.0,
pyHanko, veraPDF 1.30.2, OCRmyPDF 17.8.1/Tesseract/Poppler, LibreOffice, and
Playwright; `npm run docs:api` remained clean; `npm run proto:check` and `npm
run test:pack` passed; two deterministic OpenChestnut builds matched across 39
files; OfficeBridge passed `5/5`; and OpenChestnut passed `283/283`. The bundled
runtime remains 38 files at 14,635,200 bytes. The production tarball contains
461 files, is 9,605,861 bytes compressed and 24,354,884 bytes unpacked on the
local audit host. The offline release metadata check passes; hosted results are
recorded after the candidate commit. No publish or tag operation was attempted.

### Source-bound PPTX canonical run-hyperlink clone

On 2026-07-19, the strict imported `slide.duplicate()` profile gained one
additional closed relationship-graph slice for canonical run-level clicks. An
eligible copied shape or recursively copied group may now retain absolute
external URI clicks, internal slide jumps to retained source SlideParts, and
the five modeled action-only navigation clicks. OpenChestnut copies the exact
source relationship IDs, URI or SlidePart targets, externality, action markup,
and semantic destinations into the independent clone; it does not retarget a
source or self link to the new slide. The retained source slide XML and
relationships remain byte-identical.

Preflight parses every inline `a:hlinkClick`, proves that all owned hyperlink
and slide relationships are used by the supported graph, bounds both
relationships and markup, and repeats an exact relationship inventory after
serialization. Shape-level clicks, hover links, unknown or malformed actions,
orphan relationships, removed internal targets, links inside
tables/pictures/connectors, and broader source graphs remain fail-closed. A
fresh import is still required before editing the clone.

The shipped Presentation workflow now records exact source/clone hyperlink
graph fingerprints and normalized internal SlidePart targets in its audit. Its
fixture covers external, internal, and action-only clicks, rejects an orphan
relationship without publishing output, reimports the clone, and proves the
retained source and clone with byte-identical LibreOffice/Poppler raster pages.
The clean-install tarball test invokes the same packaged workflow.

The complete local gate passed `npm test` including Playwright and the native
Presentation render path, `npm run docs:api`, `npm run proto:check`, `npm run
test:pack`, OfficeBridge `5/5`, and OpenChestnut `285/285`. Two deterministic
OpenChestnut builds matched across 39 audited files and produced the same
manifest-bound 38-file, 14,644,416-byte runtime. Environment-gated real pikepdf,
pyHanko, veraPDF, and OCRmyPDF repeats were not enabled in this local pass;
their contract and adversarial gates still ran. The production tarball contains
461 files, is 9,613,431 bytes compressed and 24,372,583 bytes unpacked on the
local audit host. The existing 24,380,000-byte unpacked ceiling remains in
force, leaving about 7 KiB rather than hiding this growth behind a wider budget.
Hosted results are recorded after the candidate commit; no publish or tag
operation was attempted.

### Lossless public-Skill PNG package headroom

On 2026-07-19, the production package recovered sustainable growth headroom
without deleting OpenChestnut, any public Skill, or any user-facing asset, and
without widening the existing package ceilings. A pinned repository-only pako
1.0.11 tool now parses every PNG under the four npm file-type plugins with
strict signature, chunk-boundary, CRC, ordering, dimension, and inflated-stream
budgets. It replaces only the consecutive IDAT representation when a level-9
deflate is smaller. Every non-IDAT chunk remains byte-for-byte identical,
including EXIF, XMP, ICC, palette, and physical-size metadata; the complete
inflated filtered scanline stream also remains byte-for-byte identical.

The checker is idempotent and fail-closed. It inventories exactly 40 published
PNG paths, compares all 39 reference-backed paths against the pinned reference
checkout, separately binds the Excel live-control compatibility icon to its
canonical spreadsheet icon, rejects bad CRCs and trailing bytes, and proves the
write path on an unoptimized source image. A dedicated package-category gate
keeps those 40 assets below 3,550,000 bytes, while the total unpacked ceiling
remains 24,380,000 bytes. The repository-only Default Template Library and its
byte-identity hashes are not rewritten.

The PNG payload fell from 4,397,178 to 3,548,674 bytes, recovering exactly
848,504 unpacked bytes. After the added license, package metadata, and Skill
compatibility documentation, the complete local tarball contains the same 461
files, is 8,959,764 bytes compressed and 23,525,110 bytes unpacked: a net
847,473-byte unpacked reduction from the preceding candidate and about 835 KiB
of explicit headroom under the unchanged ceiling. The bundled OpenChestnut
runtime remains the same manifest-bound 38 files and 14,644,416 bytes.

The complete local gate passed `npm test` including Playwright,
LibreOffice/Poppler, qpdf, and the 20-template corpus; `npm run docs:api`, `npm
run proto:check`, `npm run test:pack`, OfficeBridge `5/5`, OpenChestnut
`285/285`, and two-build deterministic OpenChestnut verification also passed.
Explicit real-provider environments for pikepdf, pyHanko, veraPDF, and
OCRmyPDF were not configured in this pass, so their contract/adversarial gates
ran while their environment-gated real-provider repeats remained skipped.
Hosted results are recorded after the candidate commit; no publish or tag
operation was attempted.

### Source-bound PPTX closed native-chart clone

On 2026-07-19, the strict imported `slide.duplicate()` profile gained one
relationship-owning leaf: a recognized literal-data chart whose frame uniquely
consumes one internal relationship to a numbered `ChartPart`. Preflight requires
every accepted ChartPart to have no child, external, hyperlink, or data
relationship and rejects formula or external-data bindings, embedded workbooks,
duplicate use, orphan relationships, connected parts, and unrecognized chart
markup. The pending clone remains immutable until its first export/reimport.

OpenChestnut now allocates a distinct ChartPart under the clone's same
slide-local relationship ID and copies the complete chart payload byte-for-byte;
it never shares mutable chart state with the origin. The retained source
SlidePart and ChartPart remain byte-identical. After reimport, the two ChartParts
have independent package identity, and a chart that advertises the ordinary
fixed-topology edit capability can be edited without changing the origin. The
public wire only received a contract-comment clarification; no schema field or
field number changed.

The shipped Presentation workflow independently inventories source and clone
chart references, proves unique relationship consumption and empty child graphs,
permits exactly the new SlidePart/relationship/ChartPart package delta, compares
chart bytes, reimports semantics, edits the eligible clone fixture, and verifies
that the source chart remains unchanged. Its native lane also requires
LibreOffice/Poppler source/clone pixel equality. Codec regressions cover the
successful independent-copy transaction plus pending model edits, connected
ChartParts, and orphan ChartParts that must fail without output promotion. The
clean-install package test invokes the same packaged workflow.

The complete local gate passed `npm test` including Playwright and the native
Presentation render lane, `npm run docs:api`, `npm run proto:check`, `npm run
test:pack`, OfficeBridge `5/5`, and OpenChestnut `287/287`. Two deterministic
OpenChestnut builds reproduced 39 audited files and the same manifest-bound
38-file, 14,648,000-byte runtime. The production tarball contains 461 files, is
8,963,831 bytes compressed and 23,541,378 bytes unpacked on the local audit
host, leaving 838,622 bytes below the unchanged 24,380,000-byte unpacked
ceiling. The qpdf real-provider lane passed; dedicated real-provider environments
for pikepdf, pyHanko, veraPDF, and OCRmyPDF were not configured, so their
contract/adversarial tests passed while their environment-gated real repeats
were skipped. Hosted results are recorded after the candidate commit; no
publish or tag operation was attempted.

### Native PPTX custom shows

On 2026-07-19, the Presentation model, additive protocol-2 wire,
OpenChestnut C# codec, Help catalog, and native Presentation Skill converged on
one bounded `p:custShowLst` profile. Source-free decks may author an ordered
list of named native shows, each containing an ordered non-empty route through
retained slides; repeated slide references are intentional and survive the
Office package round trip. The inline graph lives only in
`ppt/presentation.xml` and does not create another SlidePart or package edge.

Canonical imports expose custom shows through the public inspect/resolve
facade. A source-bound transaction may change one existing name and replace its
ordered membership with references to retained slides. Show count and order,
facade ID, native `p:custShow/@id`, source-element hash, and imported package
topology remain fixed. Slide reordering is compatible because show entries bind
the retained presentation relationships rather than display indexes; slide
clone/delete and custom-show topology mutation remain fail closed.

Empty, extension-bearing, unknown-child, duplicate-name/native-ID, or
unresolved-reference lists are treated as opaque. OpenChestnut preserves their
exact validated source XML without exposing an incomplete semantic facade and
rejects attempted replacement. Run hyperlinks that target custom shows remain
a separate unsupported action-graph slice rather than being inferred from the
list.

The shipped workflow independently inventories the native list, requires one
exact show and exact slide names, proves that only `ppt/presentation.xml`
changed, keeps every non-target show and native ID fixed, reimports the result,
compares every model SVG, and emits a byte-bound audit. Its native lane compares
source/output LibreOffice/Poppler pages. Codec regressions cover two-show native
authoring, ordered and repeated membership edits, slide movement, source/hash
tampering, fixed-topology rejection, and extension-bearing opaque preservation.

The complete local gate passed `npm test` including Playwright,
LibreOffice/Poppler, qpdf, and the 20-template corpus; `npm run docs:api`, `npm
run proto:check`, `npm run test:pack`, OfficeBridge `5/5`, and OpenChestnut
`289/289` also passed. Two deterministic OpenChestnut builds reproduced 39
audited files and the same manifest-bound 38-file, 14,671,040-byte runtime. The
production tarball contains 463 files, is 8,971,857 bytes compressed and
23,589,246 bytes unpacked, leaving 790,754 bytes below the unchanged
24,380,000-byte ceiling. The qpdf real-provider lane passed; dedicated
real-provider environments for pikepdf, pyHanko, veraPDF, and OCRmyPDF were not
configured, so their contract/adversarial gates passed while their
environment-gated real repeats remained skipped. No publish or tag operation
was attempted.

### Native PPTX custom-show run hyperlinks

On 2026-07-20, the bounded native custom-show slice expanded from the show
list itself to canonical run-level click actions. The additive protocol-2 wire
binds a run to the stable custom-show facade ID and represents
`returnToSlide` with presence-aware boolean semantics. OpenChestnut emits the
relationship-free `ppaction://customshow?id=<native-id>` action, optionally
followed by the canonical lowercase `&return=true|false`, on source-free slide,
Master, and Layout text runs. A second import projects the current show name
back into the public `{ customShow, returnToSlide }` link object.

Imported links retain identity across a canonical show rename: only
`ppt/presentation.xml` changes, while the referring SlidePart and its
relationship part remain byte-identical and the public model resolves the new
name on second import. An explicit retarget resolves another existing canonical
show and preserves an explicitly false return policy. Missing or dangling
targets fail closed. Malformed actions and actions carrying a relationship ID
remain unmodeled and exact-source-preserved under unrelated edits; they cannot
be silently replaced. Slide clone/delete with custom-show identity references
remains outside the bounded graph and fails closed.

The shipped workflow now inventories every run action bound to the fixed native
show ID, checks exact package scope and second-import semantics, and compares a
normalized model SVG whose only ignored difference is the intentionally updated
nonvisual custom-show name annotation. The stronger package assertion keeps the
referring SlidePart byte-identical, and the native lane continues to compare
LibreOffice/Poppler pixels. JavaScript regressions cover source-free slide,
Master, and Layout authoring, rename, explicit retarget, missing targets, clone
rejection, and opaque preservation. OpenChestnut regressions additionally cover
Office 2021 validation, relationship-bearing actions, malformed return values,
and invalid non-custom-show return policies.

The complete local gate passed `npm test` including Playwright,
LibreOffice/Poppler, qpdf, and the 20-template corpus; `npm run docs:api`,
protocol lint and idempotent generation, `npm run test:pack`, OfficeBridge
`5/5`, and OpenChestnut `291/291` also passed. Two deterministic OpenChestnut
builds reproduced 39 audited files and the same manifest-bound 38-file,
14,674,624-byte runtime. The production tarball contains 463 files, is
8,974,701 bytes compressed and 23,601,235 bytes unpacked, leaving 778,765 bytes
below the unchanged 24,380,000-byte ceiling. The qpdf real-provider lane
passed; dedicated real-provider environments for pikepdf, pyHanko, veraPDF,
and OCRmyPDF were not configured, so their contract/adversarial gates passed
while their environment-gated real repeats were skipped. Hosted results are
recorded after the candidate commit. No publish or tag operation was attempted.

### PPTX custom-show run-link cloning

On 2026-07-20, the strict imported `slide.duplicate()` profile gained the
canonical relationship-free custom-show action already supported by the run
hyperlink codec. Clone preflight resolves
`ppaction://customshow?id=<native-id>[&return=true|false]` through the complete
`PptxCustomShowCatalog`; it retains the exact native show ID and return policy,
creates no OPC relationship, and still rejects malformed, relationship-bearing,
or dangling actions. The origin SlidePart and relationship part remain
byte-identical while the clone receives a distinct SlidePart.

Cloning does not imply custom-show route mutation. The same export transaction
must retain every custom show's ordered stable-ID membership, and the new slide
is not inserted implicitly. A same-transaction show rename is accepted because
the link and membership graph are stable-ID-bound; a membership change combined
with the clone fails as `unsupported_presentation_slide_clone` and must be
performed after export/reimport. Opaque or extension-bearing show graphs remain
outside this operation.

The shipped Presentation workflow independently parses the native show catalog,
records every accepted custom-show action, compares source/output membership,
checks the exact relationship-free action on the clone, and proves the boundary
again after second import. The ready PromptBench clone case now carries the same
action and has an evaluator-owned OPC oracle. Adversarial fixtures reject native
membership drift and unmodeled show extensions instead of trusting the workflow
audit.

The complete local gate passed `npm test` including Playwright,
LibreOffice/Poppler, qpdf, and the 20-template corpus; `npm run docs:api`, `npm
run proto:check`, `npm run test:pack`, OfficeBridge `5/5`, and OpenChestnut
`291/291` also passed. Two deterministic OpenChestnut builds reproduced 39
audited files and the same manifest-bound 38-file, 14,676,160-byte runtime. The
production tarball contains 463 files, is 8,975,778 bytes compressed and
23,607,273 bytes unpacked, leaving 772,727 bytes below the unchanged
24,380,000-byte ceiling. The qpdf real-provider lane passed; dedicated
real-provider environments for pikepdf, pyHanko, veraPDF, and OCRmyPDF were not
configured, so their contract/adversarial gates passed while their
environment-gated real repeats remained skipped. Hosted results are recorded
after the candidate commit. No publish or tag operation was attempted.

### PPTX embedded-XLSX OLE slide cloning

On 2026-07-20, the strict imported `slide.duplicate()` profile gained one more
closed PresentationML leaf: an unchanged top-level OLE graphic frame with one
internal XLSX `EmbeddedPackagePart` and one internal preview `ImagePart`. The
package must be uniquely referenced, relationship-free, source-hash-bound, and
semantically recognized by the existing embedded-workbook model. External,
linked, shared, non-XLSX, nested, ambiguous, relationship-bearing, or
replacement-pending OLE graphs fail before package mutation.

OpenChestnut retains the source SlidePart and relationship part byte-for-byte,
preserves the slide-local package relationship ID, copies the XLSX bytes into a
distinct clone-owned package part, and shares only the proven immutable preview
image. Post-write validation rejects package aliasing, byte drift, child graph
growth, or preview-target drift. The clone remains unchanged until an
export/import boundary; afterward, replacing only its embedded workbook leaves
the origin workbook untouched.

The shipped Presentation workflow independently inventories every package
relationship and content type, proves unique inbound ownership and exact XLSX
hashes, records source/output/reimport OLE bindings in its audit, and refuses
silent fallback. The ready `pptx-closed-leaf-slide-clone` PromptBench fixture and
evaluator now carry the same real OLE frame, distinct-package/shared-preview
oracle, second-import checks, and adversarial package-alias rejection. The
historical `0e8824c` autonomous trial predates this OLE extension, so a new
candidate repeat and reference-Skill comparison remain open evidence work.

The complete local gate passed `npm test` including Playwright,
LibreOffice/Poppler, qpdf, and the 20-template corpus; `npm run docs:api`, `npm
run proto:check`, `npm run test:pack`, OfficeBridge `5/5`, and OpenChestnut
`292/292` also passed. Two deterministic OpenChestnut builds reproduced 39
audited files and the same manifest-bound 38-file, 14,688,960-byte runtime. The
production tarball contains 463 files, is 8,989,422 bytes compressed and
23,643,498 bytes unpacked, leaving 736,502 bytes below the unchanged
24,380,000-byte ceiling. The qpdf real-provider lane passed; dedicated
real-provider environments for pikepdf, pyHanko, veraPDF, and OCRmyPDF were not
configured, so their contract/adversarial gates passed while their
environment-gated real repeats remained skipped. The offline metadata check
reported `publishReady: true` with network checks intentionally skipped. Hosted
results are recorded after the candidate commit. No publish or tag operation
was attempted.

### PPTX closed SmartArt slide cloning

On 2026-07-20, the strict imported `slide.duplicate()` profile gained a
canonical SmartArt leaf. Each accepted unchanged top-level `p:graphicFrame`
must contain exactly one `dgm:relIds` binding whose `dm`, `lo`, `qs`, and `cs`
attributes uniquely consume the standard internal diagram-data, layout,
quick-style, and colors relationships. All four typed parts must be non-empty,
use the exact standard content types, and have no child, external, hyperlink,
or data relationship. Nested, incomplete, mistyped, duplicated-binding,
relationship-bearing, or otherwise connected diagram graphs fail closed before
package mutation.

OpenChestnut preflight accounts for every diagram relationship on the source
slide, retains the source SlidePart and its relationship part byte-for-byte,
then creates four clone-owned Open XML SDK typed parts under the same
slide-local relationship IDs. The SDK currently allocates those parts under
`ppt/graphics/{data,layout,quickStyle,colors}N.xml`. Post-write validation
requires distinct clone paths, the expected relationship and content types,
and bytes identical to the corresponding source parts. This is independent
package identity for safe slide cloning, not SmartArt authoring or semantic
editing: after reimport, each SmartArt object remains opaque,
source-bound/read-only.

The shipped Presentation workflow independently inventories the raw OPC
relationships, content types, part hashes, and child graphs before invoking the
public duplicate API. It records all four source/output/reimport bindings in a
byte-bound audit, proves the clone paths are distinct while their bytes remain
equal, and refuses to publish output or audit on a connected graph. Its real
native fixture passes second import plus LibreOffice/Poppler rendering with
pixel-identical source and clone slides. The focused C# and JavaScript suites
also cover source immutability and fail-closed connected/non-graphic-frame
inputs. SmartArt is not yet a separate autonomous PromptBench case.

The complete local gate passed `npm test` including Playwright,
LibreOffice/Poppler, qpdf, and the 20-template corpus; `npm run docs:api`, `npm
run proto:check`, `npm run test:pack`, OfficeBridge `5/5`, and OpenChestnut
`293/293` also passed. Deterministic OpenChestnut verification reproduced 39
audited files and the same manifest-bound 38-file, 14,696,640-byte runtime. The
production tarball contains 464 files, is 8,998,618 bytes compressed and
23,671,680 bytes unpacked, leaving 708,320 bytes below the unchanged
24,380,000-byte ceiling. The qpdf real-provider lane passed; dedicated
real-provider environments for pikepdf, pyHanko, veraPDF, and OCRmyPDF were not
configured, so their contract/adversarial gates passed while their
environment-gated real repeats remained skipped. The offline metadata check
reported `publishReady: true` with network checks intentionally skipped. Hosted
results are recorded after the candidate commit. No publish or tag operation
was attempted.

### PPTX closed InkML content-part slide cloning

On 2026-07-20, the strict imported `slide.duplicate()` profile gained the
standard PresentationML carrier for digital ink. An accepted unchanged object
must be one direct top-level `p:contentPart` with bounded non-visual properties,
a positive transform, no extension list, and exactly one relationship
attribute. That ID must uniquely consume an internal standard or strict OOXML
`customXml` relationship to a non-empty, fully well-formed, relationship-free
`application/inkml+xml` part whose document element is `ink` in the
`http://www.w3.org/2003/InkML` namespace. Nested, ambiguous, mistyped,
wrong-root, malformed/multi-root, extension-bearing, external, or connected
graphs fail before model or package mutation.

OpenChestnut retains the origin SlidePart and relationship part byte-for-byte,
preserves the slide-local relationship ID, and uses the Open XML SDK
`CustomXmlPartType.InkContent` API to allocate a distinct clone-owned
`ppt/customXml/itemN.xml` part. The InkML bytes and content type are copied
exactly. Post-write validation admits only that typed package delta, and second
import proves disjoint source/clone part paths with equal hashes. This is
source-bound preservation for slide cloning; both resulting contentPart objects
remain opaque and read-only, with no claim of ink authoring, stroke editing, or
arbitrary Custom XML support.

The shipped Presentation workflow independently parses the source and output
OPC graphs, supports default or explicitly prefixed standard InkML roots,
records `operation.inkContentParts`, exact source/clone relationships, part
paths, hashes, and package delta, then repeats the binding check after import.
Its real fixture also requires model-render equivalence and pixel-identical
LibreOffice/Poppler source/clone pages when the native tools are present. A
connected InkML part publishes neither output nor audit. Focused C# and
JavaScript regressions additionally cover source immutability, Open XML SDK
validation, wrong-root and multi-root rejection, raw-root tampering, and
pre-mutation failure.

The complete local gate passed `npm test` including Playwright,
LibreOffice/Poppler, qpdf, and the 20-template corpus; `npm run docs:api`, `npm
run proto:check`, `npm run test:pack`, OfficeBridge `5/5`, and OpenChestnut
`294/294` also passed. Two deterministic OpenChestnut builds reproduced 39
audited files and the same manifest-bound 38-file, 14,702,784-byte runtime. The
production tarball contains 465 files, is 9,005,273 bytes compressed and
23,700,096 bytes unpacked, leaving 679,904 bytes below the unchanged
24,380,000-byte ceiling. The qpdf real-provider lane passed; dedicated
real-provider environments for pikepdf, pyHanko, veraPDF, and OCRmyPDF were not
configured, so their contract/adversarial gates passed while their
environment-gated real repeats were skipped. Hosted results are recorded after
the candidate commit. No publish or tag operation was attempted.

### PPTX closed embedded-MP4 slide cloning

On 2026-07-20, the strict imported `slide.duplicate()` profile gained the
canonical PresentationML video-picture leaf. An accepted unchanged object must
be one direct top-level `p:pic` with the exact empty `ppaction://media`
relationship sentinel, one `a:videoFile/@r:link`, one
`p14:media/@r:embed` under the standard media extension, and one distinct
poster `a:blip/@r:embed`. The video and Office-media relationships must point
to the same uniquely owned, non-empty, relationship-free `video/mp4` data part;
the poster must be an internal ImagePart. Linked, shared, nested, non-MP4,
audio, multi-binding, extension-rich, or connected media graphs fail closed.

OpenChestnut retains the origin SlidePart and relationship part byte-for-byte,
preserves both slide-local media relationship IDs, and uses the Microsoft Open
XML SDK to allocate a distinct package-level `MediaDataPart`. It copies the MP4
bytes exactly while sharing the immutable poster. The SDK may allocate the new
part under `media/*.mp4`; source Office files may use `ppt/media/*.mp4`. Both
forms pass the same single-level safe-path and exact-content-type guard. The
postwrite validator and second import prove distinct MP4 paths with equal
hashes, the same poster path, and exactly one video/one media inbound pair.

The runnable Presentation workflow now performs its independent OPC preflight
before semantic import, records `operation.mediaParts` and
`validation.package.mediaParts`, and publishes neither PPTX nor audit when the
media graph is connected or noncanonical. Its real fixture checks source
immutability, second-import model agreement, independent source/clone media
bindings, model-render equality, and LibreOffice/Poppler poster-pixel equality
when available. That render evidence validates the poster only; this milestone
does not claim video authoring, payload edits, transcoding, timing/trim control,
audio support, or playback equivalence. Media remains opaque/source-bound and
read-only after reimport.

The local candidate passed the complete `npm test` suite, including Playwright,
LibreOffice/Poppler, the real qpdf provider, all four packaged Office/PDF Skill
smokes, and the 343-file reference-Skill sync gate. OpenChestnut passed
`295/295`; OfficeBridge passed `5/5`; `npm run proto:check`, deterministic API
documentation generation, and the clean-install/package gate passed. Two clean
source builds produced the same 39 audited build files and the same manifest-
bound 38-file, 14,721,216-byte runtime. The npm dry-run contains 466 files,
9,014,382 compressed bytes, and 23,753,465 unpacked bytes. Real-provider repeats
for pikepdf, pyHanko, veraPDF, and OCRmyPDF were not configured, so their
contract/adversarial tests passed while those environment-gated repeats were
skipped.

`npm run release:check` passes the source, documentation, package, license,
JavaScript, and .NET gates. Its only remaining blocker is unavailable npm
authentication. No `npm publish` or tag/release operation has been performed.

### PDF image-backed OCR redaction

On 2026-07-20, the optional PyMuPDF sanitize provider gained the typed
`redact_ocr_text` primitive for an exact sensitive term that exists only in
raster pixels. The operation requires one explicit unrotated page, a non-empty
term, an expected match count, PyMuPDF `>=1.27.2,<1.28`, and requested
Tesseract language data. OCR is bounded to 72–300 dpi, 100 million page pixels,
1,000 matches, and a 4,096-character term. Every accepted hit must overlap one
native raster placement by at least 90%; absent/mismatched hits, off-image-only
evidence, unsafe language names, missing traineddata, excessive work, and page
rotation fail before output publication without fallback.

Accepted matches become real redaction annotations and are applied with image
pixel removal. The existing high-trust path then removes active content, runs
PyMuPDF scrub, performs a garbage-collected single-revision full rewrite,
rejects the original byte prefix, and scans raw bytes, decoded objects/streams,
metadata/XMP, attachments, annotations, forms, actions, hidden text, extracted
text, and Tesseract OCR for every sensitive term. The real fixture proves
source SHA-256 immutability, one expected image-backed match, zero final OCR
residue, inert structure, qpdf structural checks, and Poppler-rendered change confined to
the original image placement. Separate adversarial cases prove expected-count,
missing-language, unsafe-language, excessive-DPI, and rotated-page refusals
leave no artifact.

The hosted workflow now creates an isolated Python 3.13 environment pinned to
PyMuPDF 1.27.2.3, ReportLab 4.4.9, pdfplumber 0.11.9, pypdf 6.10.0, and Pillow
12.2.0, verifies Tesseract `eng` data, and exports that interpreter so the full
PDF provider smoke executes the real sanitize/OCR route instead of its
contract-only branch. Local validation used those same Python package versions,
Tesseract 5.5.2, and Poppler 26.05.0. Full `npm test`, generated API docs,
`proto:check`, the clean-install/package gate, OfficeBridge `5/5`, OpenChestnut
`295/295`, and two deterministic source builds passed. The OpenChestnut runtime
remains 38 files and 14,721,216 bytes; the production tarball contains 467
files, 9,020,751 compressed bytes, and 23,770,901 unpacked bytes. The complete
offline release gate is publish-ready; `npm whoami` still returns `ENEEDAUTH`,
so no publish or tag/release operation was attempted. Hosted results are
recorded after the candidate commit.

### PDF right-angle OCR redaction

On 2026-07-20, `redact_ocr_text` extended its high-trust raster-only profile to
pages carrying PDF `/Rotate` values of 90, 180, or 270 degrees. The operation
now binds `expected_rotation`; omission retains the backward-compatible zero-
degree precondition, while a stale or unsupported rotation fails before OCR or
output publication.

The provider temporarily clears `/Rotate` only for the Tesseract pass, obtains
exact matches in PyMuPDF's canonical unrotated page coordinate system, and
restores the original value in a `finally` path before adding real redaction
annotations. Its report retains the source rotation and emits both canonical
and display-space match/image rectangles. The independent residue scanner uses
the same temporary orientation normalization, closing the previous possibility
that raster text on a rotated source could escape required OCR evidence.

Real 0/90/180/270-degree fixtures prove source SHA-256 immutability, exact
image-backed match counts, preserved final page rotation, zero sensitive OCR
residue, single-revision inert output, and qpdf re-open. Poppler renders every
source/output pair; the pixel oracle maps the reported display-space geometry
to each rotated raster, requires a visible opaque redaction, and rejects any
changed pixel outside the original image placement. A stale rotation
precondition fails closed without a partial artifact. Manual review of the
90-degree Poppler pair confirmed the public heading and page orientation stay
unchanged while the raster secret becomes an opaque redaction.

The local candidate passed the complete `npm test` suite with the real
PyMuPDF/Tesseract, qpdf, pyHanko, Playwright, LibreOffice, Poppler, and packaged
Skill paths. OpenChestnut passed `295/295`; OfficeBridge passed `5/5`;
`npm run proto:check`, generated API-document cleanliness, two-source-build
OpenChestnut reproducibility, `npm run test:pack`, and the complete offline
release check passed. The manifest-bound OpenChestnut runtime remains 38 files
and 14,721,216 bytes; the npm dry-run contains 467 files, 9,021,747 compressed
bytes, and 23,776,728 unpacked bytes. Real pikepdf, veraPDF, and OCRmyPDF repeats
were not configured, so their contract/adversarial tests passed while those
environment-gated provider executions were skipped. The implementation
candidate at commit `276309a30fbdf6293d0aba6479b59c420a55d9ae` passed the
hosted Linux `ci` workflow in [GitHub Actions run
29714645455](https://github.com/w31r4/open-office-artifact-tool/actions/runs/29714645455)
on 2026-07-20. Its single job completed with conclusion `success` in 17m10s
after exercising all isolated PDF providers, deterministic protocol/runtime
verification, Chromium/LibreOffice/Poppler, the full npm suite, generated-doc
cleanliness, release metadata, clean-install tarball, OfficeBridge, and
OpenChestnut gates. `npm whoami` remains the external publish gate, and no
publish or tag/release operation has been attempted.

## Publishing

Before publishing:

1. Verify `package.json` and `package-lock.json` both declare `0.2.0`.
2. Rebuild and verify the OpenChestnut runtime from source.
3. Regenerate API documentation after the final public API change.
4. Inspect `npm pack --dry-run --json` for the required runtime/proto/Skill files and forbidden legacy files.
5. Run the tarball clean-install probe, not only source-tree tests.
6. Record the exact commit and hosted gate result used for publication.
