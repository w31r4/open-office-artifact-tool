---
name: Presentations
description: Create or edit PowerPoint or Google Slides decks
---

# Slides Skill

Use this skill as reference material when creating or editing presentation slide decks.

## Important Instructions

- [HARD REQUIREMENT] Content quality and storytelling: before planning the deck, read and follow [Content Quality and Narrative Rules](references/content-rules.md). Ensure the deck covers everything the user requested and forms a coherent, audience-appropriate narrative rather than a collection of disconnected facts.
- [HARD REQUIREMENT] Audience-facing copy: visible slide content must be written for the intended audience, not for the person or model producing the deck. Do not expose planning notes, timing scaffolds, talk tracks, content-selection commentary, or other internal process language unless the user explicitly requests it.

- Info density: avoid cramming low-value details onto a single slide. Prefer lower-density slides with high-value content.
  - Title slide: keep the title slide minimal and simple. Avoid cramming in too much information.
- Layout: keep things clean and simple. Avoid low-quality visuals, but also avoid excessive white space. By default, use equal left and right margins on each slide.
- [HARD REQUIREMENT] Overlap: always pay attention to programmatic overlap warnings. Do not assume that overlapping elements in diagrams are intentional, and do not ignore overlap warnings without inspecting them. You MUST fix all unintended overlap errors before delivering the slides. This is critical.
- [HARD REQUIREMENT] Font size: when a template is provided, match its font sizes. When no template or style guidance is given, you MUST use at least 50pt for deck titles, 35pt for slide titles, 24pt for mid-level text such as subheadings, callout headers, and text-box titles, and 16pt for body text.
- Text layout: when there is too much text, shorten it before shrinking the font size. Inspect visually for unexpected text wrapping. NEVER allow a title/banner text box intended for one line to wrap to two lines.
- Narrative copy must fit the chosen layout: shorten it or change layouts rather than adding density or shrinking type.
- Visual assets:
  - [HARD REQUIREMENT] DO NOT use Python to draw images; DO NOT use programmatic vector shapes for visuals; DO NOT use programmatic drawings of any sort. Use image search or image_gen tool instead!
  - [HARD REQUIREMENT] Minimize the use of diagrams. Add them only when requested or when a single diagram materially improves the clarity of complex concepts. Diagram implementation rules: use native PowerPoint shapes for simple diagrams; use Graphviz for complex relational/topological/network-like diagrams; use image_gen for highly aesthetic, illustrative, or scientific infographic diagrams (e.g. chemical structures, circuit diagrams, etc.). When using native PowerPoint shapes with connectors, create connectors (arrows/edges) before creating entity nodes, so edges appear behind nodes and never cross through node shapes or labels. If this ordering is awkward during early iteration, you may create nodes first in the initial draft, then switch to connectors-first in the revised code.
  - Before sourcing or generating visuals, be mindful of the desired aspect ratio, placement, and cropping options on the slide. For example, if you intend to place text to the left of the image containing a person, you should ask image_gen to put the person on the right side of the image.
  - By default, DO NOT reuse the same image more than once (unless it's a background).
  - Prepare visuals for both the main concept and decorative support.
- Default styling: use one composition instead of a collection of UI panels. UI-like styling typically includes card grids, pills, badges, button-like text boxes, tab or navigation patterns, repeated modular panels, dense dashboard-style layouts, and other component-library aesthetics that imply interactivity. Use stylized text boxes sparingly, favoring a flat structure on the canvas.

## Skill Folder Contents

Contents of the `slides/` skill folder:

- `container_tools/`: Standalone python scripts for slides and relevant asset manipulation.
- `references/`: Additional workflow references for specialized presentation tasks.
- `template_following_scripts/`: Helper scripts for exact source-deck/template following.
- `artifact_tool/`: API documentation and coding examples for the artifact tool library.
- `builtin_templates_support/`: Checked-in guidance, manifests, prompts, and reusable scripts for built-in templates. Each template owns its `ARTIFACT.md`; shared runners live once under `builtin_templates_support/scripts/`.
- `assets/builtin_templates/grid-layout-library/`: Blob-managed static assets for the built-in Grid Layout template, including 26 rendered previews, a model-facing registry, structured content tokens, and 26 exact plain-JavaScript artifact-tool Compose reconstructions with no JSX. This directory contains no Markdown, prompts, or reusable runners.

## Container Tools

The following helper scripts are located in the `container_tools/` directory:

- `ensure_raster_image.py`: Ensure images are rasterized; convert to PNG if needed; quick usage `--input_files <img_path1> ...`.
- `render_slides.py`: Render a PowerPoint file into a folder of PNG slides using default sizing; quick usage: `<input.pptx>`. Output files are named `slide-1.png`, `slide-2.png`, ... in a directory with the same name as the input file.
- `create_montage.py`: Build a tiled montage from images in a directory (for viewing multiple image assets or rendered slides at once); quick usage: `--input_dir <imgs_dir> --output_file <montage.png>`. It supports most image formats with auto conversion under the hood.
- `slides_test.py`: Detect content overflowing the original slide canvas; usage: `<input.pptx>`.

## Grid Layout Artifact-Tool Compose Layout Reference

This skill variant does not include the Office template file. Use the distilled layout library as initial design and composition guidance when the user has not supplied a stronger template or brand system.

Before planning slides:

1. Read `builtin_templates_support/grid-layout-library/ARTIFACT.md`, `assets/builtin_templates/grid-layout-library/design_tokens.json`, and `assets/builtin_templates/grid-layout-library/artifact-tool-compose/template-registry.json`.
2. Inspect `assets/builtin_templates/grid-layout-library/assets/previews/layout-library.png`, then shortlist layouts by `templateUse`, `layoutFamily`, `slots`, `densityBudget`, and `typographyBudget`. Do not open all 26 implementation modules by default.
3. For each selected layout, inspect its generated preview and exact `assets/builtin_templates/grid-layout-library/artifact-tool-compose/slide-XX.mjs` reconstruction.
4. Use the selected module's `layers(...)`, `text(...)`, `shape(...)`, `image(...)`, and `table(...)` helper calls as the implementation reference. Keep the output as plain `.mjs` and use `slide.compose(...)`; do not introduce JSX or a transpilation step.
5. Preserve the selected layout's content ownership, spacing, hierarchy, and media frames while replacing instructional sample text with the user's content. Vary silhouettes across the deck instead of repeating one pattern.

The shared `builtin_templates_support/scripts/create-presentation.mjs` runner can materialize any compatible built-in template for validation when passed that template's static asset root. It is not a request to emit every layout in the user's deck. User-provided templates, explicit brand guidance, and exact source evidence always override this default template.

## Workspace

Use the chat mode supplied by Codex. If the chat is not projectless, use the
project-backed layout.

Set:

- `SKILL_DIR=<absolute path to this skill>`
- `THREAD_ID=${CODEX_THREAD_ID:-manual-<timestamp-or-short-random-suffix>}`
- `TASK_SLUG=<sanitized task/deck slug>`
- `TOPIC_SLUG=<sanitized final deck filename slug>`

Select the remaining paths:

| Chat | Scratch workspace | Final PPTX |
| --- | --- | --- |
| Projectless | `$PWD/work/presentations/$TASK_SLUG` | User-requested path, otherwise `$PWD/outputs/$TOPIC_SLUG.pptx` |
| Project-backed | `$SCRATCH_ROOT/codex-presentations/$THREAD_ID/$TASK_SLUG` | User-requested path, repository convention, or `<project-root>/outputs/$TOPIC_SLUG.pptx` |

For project-backed chats, use an external scratch directory supplied by the
host. If none is supplied, compute `SCRATCH_ROOT` with
`node -p "require('node:os').tmpdir()"`; do not hardcode a platform-specific
temp path. Project-backed scratch must remain outside the repository.

An explicit user destination always wins. Set `OUTPUT_DIR` to the directory
containing `FINAL_PPTX`. If a projectless final is outside `outputs/`, an
optional copy under `outputs/` may be created for app surfacing, but the
requested path remains the primary result. Do not modify Git ignore settings
to conceal scratch files.

### Common workspace layout

After selecting `WORKSPACE`, set:

- `TMP_DIR=$WORKSPACE/tmp`
- `SLIDES_DIR=$TMP_DIR/slides`
- `PREVIEW_DIR=$TMP_DIR/preview`
- `LAYOUT_DIR=$TMP_DIR/layout`
- `ASSET_DIR=$TMP_DIR/assets`
- `QA_DIR=$TMP_DIR/qa`

Use absolute paths in scripts and handoffs. Put every generated file under
`$TMP_DIR` except `FINAL_PPTX` and any additional deliverables explicitly
requested by the user. Retain `$WORKSPACE` after delivery so follow-up turns
can inspect and reuse the prior work.

Use `.txt` for every generated intermediate prose artifact in `$TMP_DIR`,
including plans, source notes, prompt records, design notes, QA ledgers, and
fallback reasons. Reserve `.md` for installed skill/reference files such as
`SKILL.md`, `references/*.md`, and templates shipped with the skill. Do not
create generated planning files such as `slide-plan.md`.

## Route the Request Before Authoring

Choose the output path first:

1. **Existing native Google Slides deck**: use the Google Drive plugin's Google
   Slides skill. Do not round-trip it through a local PPTX unless the user asks.
2. **Net-new native Google Slides deck**: build and verify a local PPTX with
   this skill, then import it as described in Google Slides-Targeted Output.
3. **PowerPoint or local deck**: build or edit the PPTX with this skill.

For every deck built with this skill, choose exactly one visual route. The first
matching route wins:

1. **User reference or template skill**: if the user supplies a reference deck,
   asks to follow an existing deck, or invokes a template skill, use only that
   file as the visual source. An existing PPTX being edited also counts as the
   reference. Do not mix in Grid Layout or another template.
2. **Explicit custom formatting**: if there is no reference and the user asks
   for a theme, brand treatment, visual style, mood, or custom formatting,
   create the deck from scratch. Do not use Grid Layout.
3. **No visual direction**: use the bundled Grid Layout Artifact.md layout
   library as the composition reference. Select and adapt layouts using the
   Grid Layout instructions above; do not run PPTX template-following mode.

User-provided references and explicit visual direction always take precedence
over Grid Layout.

## Google Slides-Targeted Output

For a net-new native Google Slides request, create and verify a local `.pptx`
with this skill first. The native Google Slides deliverable must then be
produced by the Google Drive plugin's presentation import action,
`mcp__codex_apps__google_drive_import_presentation`, with
`upload_mode: "native_google_slides"`.

Do not use Computer Use, Browser Use, blank-Google-Slides creation plus Google
Slides write APIs, or another direct-to-Slides construction path for net-new
Google Slides unless the user explicitly asks for that alternate workflow. If
the Google Drive plugin is unavailable, ask the user to install
`google-drive@officer-curated`. If the plugin is available but presentation
import is missing, ask the user to reinstall or refresh the Google Drive plugin
before continuing with the native Google Slides deliverable.

The local `.pptx` creation and native import workflow above applies only to
net-new Google Slides deliverables.

## Implementation

You MUST use `open-office-artifact-tool` from JavaScript ES modules to implement the slide deck.

Read the local docs before coding:

- `artifact_tool/API_QUICK_START.md`
- `artifact_tool/api/API_DOCS.md`

For native charts, read `artifact_tool/api/references/charts.spec.md` before
authoring or editing. Canonical OpenChestnut output covers literal bar, line,
pie, standard area, fixed 50%-hole doughnut, marker-only scatter, bounded 2D
bubble, and the documented clustered bar+line combo. Use
`examples/openchestnut-chart-families-workflow.mjs` as the Agent-facing
create/import/edit/reimport/render/audit pattern. Inspect an imported ChartPart
before mutation, keep its supported topology fixed, render the final slide, and
let formula-backed, external-workbook, connected, or advanced chart graphs fail
closed instead of rebuilding them from visible caches.

For slide backgrounds, use the typed `slide.setBackground(...)` and
`slide.clearBackground()` primitives documented in
`artifact_tool/api/references/slide.spec.md`. Direct solid/style-reference
backgrounds cross the canonical OpenChestnut PPTX path. Never flatten an
inherited Layout/Master background or silently replace an advanced imported
background graph; preserve it unchanged or let the export fail closed.

For compound objects that must retain one ownership tree and local coordinate
space, use native `slide.groups.add(...)` and read
`artifact_tool/api/references/grouping.spec.md`. Canonical recursive groups
cross the OpenChestnut PPTX path; imported topology is fixed, and complex group
shells remain one opaque read-only object rather than being flattened.

For imported deck order, `slide.moveTo(existingZeroBasedIndex)` changes only
the retained source `SlidePart` order in `p:sldIdLst`; it does not copy or
reconstruct slide graphs. `slide.delete()` is separate and intentionally much
narrower: it performs a real OPC delete only for an isolated layout-only,
non-final slide whose source part has exactly its layout relationship, no
inbound relationship, and no custom-show/section/extension or presentation-level
identity reference. It
removes the source slide part and its relationship part while preserving every
survivor. Media, notes, comments, charts, OLE, hyperlinks, data parts, or any
other connected graph fail closed. `slide.duplicate()` is a separate, much
narrower operation: only an original imported slide whose unchanged graph has
canonical shapes, canonical inline fixed-grid tables with bounded rectangular merges, recognized closed
literal-data charts, eligible top-level embedded-XLSX OLE frames, canonical
top-level four-part SmartArt frames, canonical top-level closed InkML
`p:contentPart` objects, canonical top-level closed embedded-MP4 media pictures,
canonical embedded rectangular images, bounded canonical
straight/elbow connectors, plus recursively canonical groups whose descendants
contain only the non-native-graph leaf kinds, exactly one
internal layout relationship, picture-bound image relationships, canonical
run-level click hyperlinks, and optionally
one closed `NotesSlide -> NotesMaster` / back-to-source-slide leaf plus one
canonical legacy `SlideCommentsPart` leaf may receive a new `SlidePart` and
presentation relationship. It deliberately shares the verified layout,
immutable image Parts, NotesMaster, and presentation-wide `CommentAuthorsPart`
through fresh clone-local relationships; it copies accepted NotesSlide and
SlideComments XML byte-for-byte and points only the preserved notes
back-reference at the clone. The comments part and author catalog must have no
connected relationship graph. Accepted tables are inline-only: table fills,
links, and every other package edge remain outside this profile. Every present
connector endpoint must resolve to an element in the same copied `SlidePart`
tree; accepted connectors add no relationship, and their pending clone targets
resolve to fresh clone-local elements. Accepted groups add no relationship
themselves, and every nested picture must consume one exact verified image
relationship. It preserves the origin part and requires export plus reimport before the clone, its notes, or its comments may be edited;
imported legacy comments remain source-bound read-only after that boundary.
Each accepted chart frame must consume one unique internal relationship to a
numbered `ChartPart`; the ChartPart may not own a child, external, hyperlink, or
data relationship. Export byte-copies it into a distinct clone-local ChartPart
rather than sharing mutable chart state. After export/reimport, the two
ChartParts are independent; a chart that advertises the ordinary fixed-topology
edit capability can use that path without affecting the origin. Formula or
external-data charts, embedded workbooks, duplicate/orphan chart relations, and
any connected chart graph fail closed.
Each accepted OLE frame must bind exactly one internal, uniquely inbound XLSX
`EmbeddedPackagePart` with no child relationship graph and exactly one internal
preview `ImagePart`. Export byte-copies the workbook into a distinct clone-local
package under the same slide-local relationship ID while sharing the immutable
preview. After export/reimport, `nativeObject.replaceEmbeddedWorkbook(...)` on
the clone changes only that independent package. Shared/external/non-XLSX,
nested, relationship-bearing, ambiguous, or replacement-pending OLE graphs fail
closed.
Each accepted SmartArt frame must be a top-level `p:graphicFrame` with exactly
one `dgm:relIds` root. Its `r:dm`, `r:lo`, `r:qs`, and `r:cs` bindings must
resolve to internal diagram data, layout, quick-style, and colors parts with
the exact standard content types and no child/external/hyperlink/data graph.
Export byte-copies all four into distinct typed clone-local parts under the
same slide-local relationship IDs. Reimport proves disjoint paths and equal
hashes. This prevents source/clone coupling. A separately recognized canonical
one-paragraph/one-run DiagramDataPart may expose only source-bound
`setDiagramNodeText(modelId, text)` after import; all other SmartArt remains
source-bound and read-only.
Nested, incomplete, duplicated, mistyped, external, relationship-bearing, or
otherwise noncanonical diagram graphs fail closed.
Read `artifact_tool/api/references/smartart-clone.spec.md` before accepting a
slide whose duplicate profile includes SmartArt.
Each accepted InkML object must be one top-level `p:contentPart` with one exact
internal `customXml` relationship to a non-empty, relationship-free
`application/inkml+xml` CustomXmlPart whose document element is `ink` in the
standard InkML namespace. Export preserves the slide-local relationship ID but
byte-copies the payload into a distinct SDK-typed clone part; reimport proves
disjoint paths and equal hashes. This is opaque source-bound preservation, not
ink authoring or Custom XML editing. Nested, extension-bearing, ambiguous,
mistyped, non-InkML-root, or connected content parts fail closed. Read
`artifact_tool/api/references/inkml-content-part-clone.spec.md` before accepting
a slide whose duplicate profile includes InkML.
Each accepted embedded video must be one top-level canonical `p:pic` whose
empty media-action sentinel, `a:videoFile`, `p14:media`, and poster blip bind
exactly three slide-local relationships. The video and media relationships must
share one uniquely owned, non-empty, relationship-free `video/mp4` data part;
the third relationship must bind one internal poster ImagePart. Export keeps
both media relationship IDs, byte-copies the MP4 into a distinct Open XML SDK
`MediaDataPart`, and shares only the immutable poster. Reimport proves different
MP4 paths with equal hashes and the same poster path. This is opaque
source-bound clone preservation, not video authoring, playback validation,
transcoding, timing/trim editing, or audio support. Nested, linked, shared,
non-MP4, extension-bearing, multi-binding, or connected media graphs fail
closed. Read `artifact_tool/api/references/embedded-video-clone.spec.md` before
accepting a slide whose duplicate profile includes media.
Accepted run links are limited to modeled external absolute URIs, internal jumps
to a retained SlidePart, `nextSlide`/`previousSlide`/`firstSlide`/`lastSlide`/
`endShow` actions, and relationship-free custom-show actions whose native ID
resolves through the canonical presentation-wide show catalog. The clone keeps
each relationship-backed link's exact `r:id` and target; internal jumps keep
pointing to the same retained source target. A custom-show link keeps the same
stable native show ID and optional return policy, creates no package
relationship, and never inserts the clone into the show's membership. Every
hyperlink relationship must be consumed by one of the other inline clicks.
Shape-level clicks, hover links, malformed or orphan relationships, links in
tables/pictures/connectors, unknown actions, malformed/relationship-bearing/
dangling custom-show actions, and jumps to a removed slide fail closed.
Imported add, repeat/mutated clone, rich/connected comments, unsupported
connector forms or targets, and every broad graph clone remain unsupported until
an explicit OPC graph-clone transaction is available.

For the bare agent-facing clone profile, use the shipped transaction rather
than copying ZIP parts or rebuilding the slide:

```sh
node "$SKILL_DIR/examples/openchestnut-slide-duplicate-workflow.mjs" \
  input/source.pptx output/source-with-copy.pptx output/clone-audit.json \
  "Unique source slide name"
```

It requires exactly one explicitly named original imported slide and accepts
the closed canonical profile with no NotesSlide or legacy-comments leaf.
Recognized closed ChartParts are included without an opt-in: the workflow
proves one unique frame relationship per chart, no ChartPart child graph, a
distinct clone-local target, and byte-identical chart payload. It proves the
same independent-copy contract for every eligible embedded-XLSX OLE workbook,
including one unique inbound package edge, empty child graph, exact content
type/hash, same slide-local `r:id`, distinct clone package, and shared preview
ImagePart. It also proves every accepted SmartArt frame's exact four
`dm/lo/qs/cs` roles, relationship IDs/types, standard content types, empty
child graphs, distinct clone-local targets, and byte-identical XML. For every
accepted InkML content part it independently proves one exact `customXml`
relationship, the standard content type and root namespace, an empty child
graph, a distinct clone-local `CustomXmlPart`, and byte-identical XML. It proves the
same video/media relationship pairing, unique inbound ownership, exact
`video/mp4` bytes, distinct clone-local `MediaDataPart`, and shared immutable
poster for every accepted embedded video. It proves the
source part order, inserts one adjacent clone, keeps every retained source part
byte-identical except the required package topology records, allows only the
new SlidePart, its relationship part, and the exact cloned ChartParts, XLSX
packages, SmartArt, InkML, and MP4 parts, then checks exact source/clone
external and internal run-link relationship IDs and targets
with no orphan edge, then reimports and compares the source/clone semantics and
model render. Model
SVG comparison ignores fresh `data-*-id` locator attributes only; it is not a
claim that the clone XML is lexically byte-identical. Missing/duplicate names,
notes/comments, unresolved connector endpoints, unsupported link markup,
nonliteral or connected charts, other graph leaves, or any unexpected package
part fail closed without promoting output or audit. In particular, a nested,
incomplete, duplicated-relationship binding, mistyped, external, or connected SmartArt
graph, or a nested, extension-bearing, mistyped, non-InkML-root, ambiguous, or
connected content part, or a nested, linked, shared, non-MP4, multi-binding, or
connected media graph, is rejected before semantic import, `slide.duplicate()`,
or publication.

The default is intentionally bare. To copy only the separately supported,
already-closed relationship leaves, opt in explicitly rather than relying on a
fallback or ZIP manipulation:

```sh
node "$SKILL_DIR/examples/openchestnut-slide-duplicate-workflow.mjs" \
  input/source.pptx output/source-with-copy.pptx output/clone-audit.json \
  "Unique source slide name" --allow-closed-leaves
```

This opt-in accepts at most one canonical `NotesSlide` with exactly its
`NotesMaster` and back-to-source-slide relationships, and at most one canonical
legacy `SlideCommentsPart` with no child relationship graph plus the immutable
presentation-wide `CommentAuthorsPart`. The audit lists every new notes/comment
part, proves NotesSlide and comments XML are verbatim copies, proves the notes
back-reference points to the clone, and proves the immutable master/catalog are
shared. Rich/modern comments, any extra relationship, and any graph outside
that exact profile still fail closed.

For one imported canonical SmartArt document node, do not patch the ZIP or
rebuild the diagram. First inspect `nativeObject.diagramText`; it is present
only when the closed four-part graph has a direct plain
`dgm:t > a:p > a:r > a:t` document-node profile. Then run the public
transaction below, which changes only the bound DiagramDataPart and writes a
no-overwrite audit:

```sh
node "$SKILL_DIR/examples/openchestnut-smartart-text-edit-workflow.mjs" \
  input/source.pptx output/edited.pptx output/edited.audit.json \
  "Closed SmartArt" "{B31B1833-2B65-4D6B-B3D4-9B3988427B21}" "Before" "After"
```

The workflow resolves exactly one object/node/expected-text triple, preserves
the source, verifies that no non-data package part changed, reimports the
requested node list, and fails closed for rich/multi-run, connected, nested, or
ambiguous SmartArt. It does not add/reorder nodes, change layout/style/colors
or geometry, or claim model SVG verification is a native-host rendering check.
Read `artifact_tool/api/references/smartart-clone.spec.md` before using either
the clone or text-edit profile.

For an original imported slide, `slide.name = "Decision review"` is a narrow
in-place metadata edit: OpenChestnut changes only that SlidePart's
`p:cSld/@name`, preserves its relationship graph and all other parts, and
requires reimport for a fresh binding. It is not available for a pending clone,
which must remain an exact source copy until its export/reimport boundary.

When an imported top-level OLE object contains one uniquely bound XLSX package,
read `artifact_tool/api/references/ole-workbooks.spec.md` before changing it.
Only `getEmbeddedWorkbook()` and `replaceEmbeddedWorkbook(...)` are allowed:
the latter replaces the validated workbook bytes while preserving the OLE
shell, relationship topology, preview image, and every unrelated native part.
Do not patch an embedding part directly or present a reconstructed OLE object
as equivalent; ambiguous, shared, malformed, or source-tampered graphs must
fail closed.

For review annotations, read `artifact_tool/api/references/comments.md` before
calling `slide.comments.addThread(...)`. Canonical PPTX export supports only
bounded legacy slide-level comments with `undefined` targets, one author and
text item per annotation, and explicit coordinates. A completely comment-free
imported presentation may advertise `slide.comments.capability.addable`; that
permits creation of a canonical shared author catalog and closed slide-local
comment leaves. Existing legacy graphs remain unchanged-only. Modern threads,
replies, reactions, resolved state, and element/text anchors must stay in their
native family or fail closed; never flatten them into a legacy comment.

Before running any generated presentation module, initialize its workspace so
Node.js can resolve the bundled `open-office-artifact-tool` package:

```bash
node "$SKILL_DIR/container_tools/setup_artifact_tool_workspace.mjs" \
  --workspace "$TMP_DIR"
```

Create the ES module source file (`.mjs`) under `$TMP_DIR` and export the final
PowerPoint deck (`.pptx`) to `$FINAL_PPTX`. The generated source must be plain
JavaScript that runs directly with `node`; do not require a transpiler or build
step.

You MUST NOT use `python-pptx` or the old Python `artifact_tool` API.

### Bounded Imported Slide Name Edit

For one uniquely named original imported slide, use the shipped public
OpenChestnut workflow rather than patching `ppt/slides/slide*.xml` directly:

```bash
node examples/openchestnut-slide-name-edit-workflow.mjs \
  input.pptx output.pptx audit.json \
  "Go-no-go decision" "Go decision: controlled rollout"
```

It checks the exact source name, maps the source presentation relationship list
to the target SlidePart, changes only `slide.name`, and then proves the saved
package has the same part topology, byte-identical non-target parts, and the
requested target `p:cSld/@name`. Open XML SDK may canonicalize the target
SlidePart's XML serialization; the workflow therefore reimports, preserves the
rest of the target slide's semantics, requires a byte-identical model SVG, and
writes a source/output-bound audit. Duplicate/missing names, fallback-only
native names, unexpected package changes, pending clones, and any other
ambiguous edit fail closed. This is not a generic template metadata editor.

### Native Custom Shows

For source-free decks, create all slides and then use
`presentation.customShows.add(nameOrConfig, slides)` to author real
`p:custShowLst` playback routes. For a canonical imported list, only an
existing show's name and ordered retained-slide membership are editable; show
count/order, facade identity, and native ID remain fixed. Read
`artifact_tool/api/references/custom-shows.spec.md` before changing one.
Canonical text runs may target an existing show by exact name and may set
`returnToSlide: true|false`. OpenChestnut binds that run to the show's stable
facade/native identity, so renaming the show keeps the native action and
SlidePart bytes unchanged while the next import exposes the new public name.

For one exact imported show, use the shipped transaction instead of patching
`ppt/presentation.xml`:

```bash
node examples/openchestnut-custom-show-workflow.mjs \
  input.pptx output.pptx audit.json \
  "Board route" "Executive route" "Appendix,Overview,Appendix"
```

The workflow resolves every supplied slide name uniquely, preserves the source,
proves that only `ppt/presentation.xml` changed, retains native show identity
and all non-target shows, counts any run links bound to that fixed identity,
reimports, compares normalized visual SVG content, and writes a
source/output-bound audit. Lists with extensions, unknown children, unresolved
relationships, duplicate identities, or another noncanonical graph remain
opaque and fail closed. Missing targets and malformed, relationship-bearing,
or dangling custom-show actions fail closed. The bounded clone workflow accepts
only the canonical relationship-free run action and proves that show membership
did not change; slide deletion and custom-show topology mutation remain separate
fail-closed operations. Run LibreOffice/Poppler review after delivery when
available.

### Native PowerPoint Sections

PowerPoint sections are not custom shows: sections form the complete ordered
partition of a deck, whereas custom shows are optional playback subsets. For a
new deck, add every slide first and then define the entire partition through
`presentation.sections`:

```js
const opening = presentation.slides.add({ name: "Opening" });
const evidence = presentation.slides.add({ name: "Evidence" });
const decision = presentation.slides.add({ name: "Decision" });

presentation.sections.add("Context", [opening, evidence]);
presentation.sections.add("Decision", [decision]);
```

The export writes the native Office 2010 `p14:sectionLst` extension in
`ppt/presentation.xml`. Each section must have a unique name and at least one
slide; flattening all memberships must reproduce the current slide order
exactly, with no duplicates or omissions. Inspect an imported deck with
`presentation.inspect({ kind: "section" })`, then resolve or look up an
existing section and change only its name or boundary with `setSlides(...)`.
Canonical imports keep section count, order, public identity, and native GUIDs
fixed. Do not patch `ppt/presentation.xml` directly, add/delete/reorder an
imported section, or combine sections with slide insertion/deletion/duplicate:
those operations fail closed. Duplicate, extension-bearing, unresolved, or
otherwise irregular native section graphs are opaque-preserved and cannot be
semantically replaced. Read `artifact_tool/api/references/sections.spec.md`
before editing an imported deck, then reimport and inspect sections after
export; run native render review when available.

### Bounded Slide Transitions

Use direct `p:transition` metadata only for an intentional between-slide
movement. The public profile is deliberately small: `fade` or directional
`push`, `slow`/`medium`/`fast`, click advancement, and an optional bounded
timer. It is not an animation/timing/sound authoring surface:

```js
slide.setTransition({
  effect: "push",
  direction: "left",
  speed: "fast",
  advanceOnClick: false,
  advanceAfterMs: 4_000,
});
```

For an imported deck, inspect `slide,transition`, resolve
`${slide.id}/transition`, and read `transition.capability` before calling
`set(...)` or `clear()`. Only one existing canonical direct fade/push graph is
editable. A source-bound slide with no transition is not an addable surface;
unknown effects, timing trees, sound actions, `p14` duration, or extension
graphs stay opaque-preserved and fail closed on mutation. The strict slide
clone profile may carry one unchanged canonical direct transition, but never a
timing or sound graph.

Always export, reimport, and inspect the transition again. Static
LibreOffice/Poppler review can prove the visible slide content is stable, not
slideshow playback; use a native PowerPoint playback QA lane when timing or
host effect behavior matters. Read
`artifact_tool/api/references/transitions.spec.md` before modifying imported
transition metadata.

### Bounded Imported Speaker-Notes Add

An imported slide whose source SlidePart has no NotesSlide may add plain-text
speaker notes only when `slide.speakerNotes.capability.addable` is true. Inspect
`slide,notes` first, resolve `${slide.id}/notes`, and prefer the shipped
transaction over direct OOXML relationship edits:

```bash
node examples/openchestnut-speaker-notes-add-workflow.mjs \
  input.pptx output/with-notes.pptx output/with-notes.audit.json \
  "Unique target slide name" "Lead with the evidence.\nClose with the decision."
```

The workflow requires exactly one named imported slide with a notes-absent,
explicitly addable capability. It protects the source bytes, writes to a
temporary path, reimports, checks exact notes plus stable visible slide
semantics/order/name, compares model SVG, and audits the OPC graph. An existing
single NotesMaster is reused byte-for-byte; otherwise OpenChestnut creates one
canonical NotesMaster sharing the first ordered SlideMaster's existing
ThemePart. The new NotesSlide must have exactly one NotesMaster relationship and
one back-reference to its owning SlidePart. Export independently re-proves the
source graph, so changing capability data cannot grant write authority.
Inconsistent/multiple NotesMaster graphs, unusable themes, existing/rich notes,
ambiguous slide names, and any unexpected relationship fail closed with no
output promotion. Run native LibreOffice/Poppler source-vs-output comparison
after delivery; speaker notes must not change the visible slides. See
`artifact_tool/api/references/speaker-notes.spec.md`.

### Bounded Imported Legacy Review-Comment Add

For an ordinary imported deck with no legacy or Office 2021 comments anywhere,
inspect `slide.comments.capability` before adding a review annotation. Prefer
the shipped transaction over editing `.rels`, `commentAuthors.xml`, or
`comments/comment*.xml` yourself:

```bash
node examples/openchestnut-legacy-comment-add-workflow.mjs \
  input.pptx output/with-review.pptx output/with-review.audit.json \
  "Unique target slide name" "Confirm the imported evidence." \
  "Review Owner" "2026-07-20T03:04:05Z" 360 240
```

The workflow requires exactly one named source-bound target whose capability is
`{ format: "legacy", partPresent: false, addable: true }`. It protects the
source, adds one slide-level annotation, exports through OpenChestnut, and then
independently proves that only a canonical `CommentAuthorsPart`, one numbered
closed `SlideCommentsPart`, their two collision-free relationships, content
types, and corresponding relationship Parts changed. Slide XML, slide order,
names, and visible semantics remain unchanged. It reimports the exact author
and text, compares model SVG, emits a byte-bound audit, and uses exclusive
output publication. Native LibreOffice/Poppler source-vs-output pages must be
pixel-identical because legacy review comments are nonvisual in slideshow
rendering.

The capability is defensive preflight evidence only. OpenChestnut re-proves the
complete source package and rejects a forged flag, an existing author catalog,
any legacy or modern comments part on any slide, mixed/connected comment graphs,
or a second add after reimport. Existing imported legacy comments remain
read-only; this vertical slice is canonical creation from a comment-free source,
not topology editing. See `artifact_tool/api/references/comments.md`.

### Bounded Imported Title And Speaker-Notes Edit

For one known slide with one known text shape and a canonical plain-text Notes
part, prefer the shipped public-API/OpenChestnut workflow over an ad-hoc package
patch. The title may be an ordinary editable shape or a concrete imported
SlidePart placeholder with a recognized local text body. The latter grants only
fixed-topology character replacement: native placeholder identity, geometry,
formatting, and layout binding remain source-bound. The workflow imports,
checks the exact source title and notes, changes only those two text values,
exports to a distinct path, reimports, verifies the retained slide/title/notes
identities, produces a model SVG check, and writes a byte-bound audit.

```bash
node examples/openchestnut-title-notes-edit-workflow.mjs \
  input.pptx output.pptx audit.json
```

The optional remaining arguments are, in order: slide name, title-shape name,
expected title, replacement title, expected notes, and replacement notes. The
workflow deliberately fails closed for duplicate/missing slide or shape names,
changed expected source text, absent/rich notes, slide-name/order changes, or
any identity/geometry/direct-background change after reimport. A recognized
placeholder title must also retain its native newline/inline topology; complex
multi-run replacements and unrecognized local text graphs fail closed. It does
not claim universal template editing: SmartArt, irregular modern comment
graphs, rich notes, animations, and other connected PresentationML graphs stay
source-bound.

Run the native render/QA route after delivery when LibreOffice/Poppler is
available; the workflow's SVG check is model evidence, not a substitute for a
native-host review.

For native Office 2021 comment threads, read
`artifact_tool/api/references/comments.md` before authoring or editing. Use the
shipped workflow for a complete root/direct-reply create → import → fixed-
topology text/status edit → second import → inspect/render/audit loop:

```bash
node examples/openchestnut-modern-comment-workflow.mjs \
  output/decision-review.pptx output/modern-comments-audit.json
```

This uses `Presentation.create({ commentFormat: "modern" })`, a top-level
element or shape-text-range anchor, independent person/GUID/time metadata, and
`thread.resolve()`/`thread.reopen()`. On imported threads only existing text and
status may change. Author/date identity, anchor and range, position, root/reply
topology, relationships, and source hashes remain fixed. Reactions/task fields,
nested replies, unknown/nested anchors, connected comment parts, and mixed
legacy/modern graphs remain opaque/source-bound and fail closed.

## Template Following

Use template-following mode only when a user-provided source PPTX supplies the
layout, style, or template. Read `references/template-following.md`, use
`$TMP_DIR` from the Workspace section, and set
`TEMPLATE_PPTX="<absolute path to the user-provided PPTX>"`.

Current availability: the reference starter-deck command below still needs a
broad imported-slide graph clone and broad graph delete semantics, so it
deliberately fails closed in the canonical codec. The isolated layout-only
`slide.delete()` and unchanged shape/inline-table/image/recursive-group clone profiles with closed notes and
legacy comments are not substitutes: the latter creates an independent part but
cannot be edited before an export/reimport boundary, and it cannot carry
arbitrary template graph edges. Do not rebuild or share slide parts to emulate
a clone.
Until the broader milestone exists, use this mode only for source inventory,
plan validation, and render/QA evidence; report the clone limitation before
promising a derived starter deck.

Preserve the source deck's typography, palette, spacing, layout, placeholders,
footers, page markers, and brand chrome unless the user explicitly asks to
restyle. Do not use template-following mode for a deck created from scratch.

Create:

- `$TMP_DIR/template-audit.txt`
- `$TMP_DIR/template-frame-map.json`
- `$TMP_DIR/deviation-log.txt`

Keep `$TMP_DIR/source-notes.txt` for content and asset provenance.

Inspect the complete source deck:

```bash
node "$SKILL_DIR/template_following_scripts/inspect_template_deck.mjs" \
  --workspace "$TMP_DIR" \
  --pptx "$TEMPLATE_PPTX"
```

Map each output slide to an inherited source slide and identify element-level
`editTargets`. Then validate the map. The later starter-deck command is retained
for the future broad graph-clone milestone, but currently rejects before writing
an output deck:

```bash
node "$SKILL_DIR/template_following_scripts/validate_template_plan.mjs" \
  --workspace "$TMP_DIR" \
  --map "$TMP_DIR/template-frame-map.json"

node "$SKILL_DIR/template_following_scripts/prepare_template_starter_deck.mjs" \
  --workspace "$TMP_DIR" \
  --pptx "$TEMPLATE_PPTX" \
  --map "$TMP_DIR/template-frame-map.json" \
  --out "$TMP_DIR/template-starter.pptx" \
  --preview-dir "$TMP_DIR/template-starter-preview" \
  --layout-dir "$TMP_DIR/template-starter-layout" \
  --contact-sheet "$TMP_DIR/template-starter-contact-sheet.png"
```

When a future broad graph-clone milestone enables `template-starter.pptx`,
import it with `open-office-artifact-tool` and edit only inherited slides/objects
unless the validated frame map explicitly allows an insertion. Today, if a
source slide cannot support requested content without broad clone/delete or a
parallel rebuild, report the blocker and the closest viable source-slide
options.

## QA Reminder

Before delivery, render every final slide and inspect each slide individually
at full size. Use a contact sheet only to review deck-level flow and consistency,
not as a substitute for full-size layout QA. Fix unintended overlap, clipping,
wrapping, broken connectors, unresolved placeholders, inconsistent footers/page
markers, and chart/data
mismatches before exporting. Verify that researched claims and sourced assets
are traceable, and cite sources if research was used.

## Final Response

Return a short user-visible summary of the completed deck. Mention the sources cited or
used if research informed the deck. Do not attach scratch plans, previews,
layout JSON, or temporary assets unless the user asks for them.

## Codex App final response citations

Use the inline form `:codex-file-citation{...}` and place each citation immediately after the claim it supports.

For read-only Q&A, cite the source deck. For a successful edit or creation, cite the final delivered deck. For a no-op edit, cite the inspected source deck.

For read-only Q&A, inspect the complete relevant slide, including callouts, the exact question or prompt, chart or table titles, displayed totals or sample sizes, and source or methodology footers. State the direct answer first and cite each distinct evidence-bearing object when exact IDs are available.

Unless the user requests an in-place edit, preserve the input PPTX and export a distinct edited copy. Cite every changed slide in the final response. If no requested content is found and no output is modified, cite the inspected source deck with a plain file citation.

For creation, include exactly one standalone Markdown link to the final delivered PPTX. Do not add a file, slide, or object citation.

Use slide citations when slide numbers come from the latest rendered or inspected cited deck:

```text
:codex-file-citation{path="/abs/path/deck.pptx" artifact_kind="presentation" slide_number="3"}
```

Include `slide_id` only when artifact-tool inspection provides the exact stable `sl/...` ID and stable navigation matters:

```text
:codex-file-citation{path="/abs/path/deck.pptx" artifact_kind="presentation" slide_number="1" slide_id="sl/gs5z1kshq0xv"}
```

For a concrete chart, table, image, diagram, or callout, include `object_id` only when inspection provides the exact ID and you can add a useful label:

```text
:codex-file-citation{path="/abs/path/deck.pptx" artifact_kind="presentation" slide_number="1" slide_id="sl/gs5z1kshq0xv" object_id="ch/pz9t1r3ka8vn" label="ARR by segment chart"}
```

Do not cite internal previews, contact sheets, layout JSON, source notes, scratch files, builders, manifests, or QA outputs unless asked. If slide or object IDs are not reliable, cite the slide without object detail rather than guessing.
