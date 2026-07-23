# Presentation Quick Start (TypeScript)

## API Docs

API docs: `api/API_DOCS.md`

`Presentation` is the in-memory API object. A deck is the exported presentation
file. A slide is one page in the deck.

## Imports

```ts
import fs from "node:fs/promises";
import { FileBlob, Presentation, PresentationFile } from "open-office-artifact-tool";
```

## Create, Render, Export

This script creates an editable deck, renders each slide PNG, writes each slide
layout JSON, writes a deck montage with `montage: true`, and exports PPTX.

```ts
import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "open-office-artifact-tool";

async function writeBlob(path: string, blob: Blob): Promise<void> {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

async function main(): Promise<void> {
  await fs.mkdir("output", { recursive: true });

  const presentation = Presentation.create({
    slideSize: { width: 1280, height: 720 },
  });

  const slide = presentation.slides.add();
  slide.background.fill = "slate-50";

  const page = { left: 72, top: 64, width: 1136, height: 592 };
  const gutter = 28;
  const leftCol = 430;
  const rightCol = page.width - leftCol - gutter;

  const eyebrow = slide.shapes.add({
    geometry: "textbox",
    position: { left: page.left, top: page.top, width: 280, height: 28 },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  eyebrow.text = "EDITABLE PRESENTATION";
  eyebrow.text.style = { fontSize: 12, bold: true, color: "slate-500" };

  const title = slide.shapes.add({
    geometry: "textbox",
    position: {
      left: page.left,
      top: page.top + 92,
      width: leftCol,
      height: 184,
    },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  title.text = "Build decks with editable objects";
  title.text.style = { fontSize: 42, bold: true, color: "slate-950" };

  const subtitle = slide.shapes.add({
    geometry: "textbox",
    position: {
      left: page.left,
      top: page.top + 292,
      width: leftCol,
      height: 96,
    },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  subtitle.text =
    "Rows, columns, grids, charts, tables, images, and text remain editable.";
  subtitle.text.style = { fontSize: 20, color: "slate-600" };

  const chartFrame = slide.shapes.add({
    geometry: "roundRect",
    name: "chart-frame",
    position: {
      left: page.left + leftCol + gutter,
      top: page.top + 92,
      width: rightCol,
      height: 388,
    },
    fill: "white",
    line: { style: "solid", fill: "slate-200", width: 1 },
    borderRadius: "rounded-2xl",
    shadow: "shadow-sm",
  });

  slide.charts.add("bar", {
    position: {
      left: chartFrame.position.left + 36,
      top: chartFrame.position.top + 52,
      width: chartFrame.position.width - 72,
      height: 280,
    },
    categories: ["Rows", "Grid", "Tokens"],
    series: [{ name: "Coverage", values: [3, 4, 5], fill: "accent1" }],
    hasLegend: false,
    dataLabels: { showValue: true, position: "outEnd" },
    yAxis: {
      majorGridlines: { style: "solid", fill: "slate-200", width: 1 },
    },
  });

  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    const png = await presentation.export({ slide, format: "png", scale: 1 });
    await writeBlob(`output/${stem}.png`, png);

    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`output/${stem}.layout.json`, await layout.text());
  }

  const montage = await presentation.export({
    format: "webp",
    montage: true,
    scale: 1,
  });
  await writeBlob("output/deck-montage.webp", montage);

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save("output/deck.pptx");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
```

## Canonical Native Chart Families

OpenChestnut writes literal-data native ChartParts for `bar`, `line`, `pie`,
standard `area`, fixed 50%-hole `doughnut`, marker-only `scatter`, bounded 2D
`bubble`, and the combo profile below. Category families use shared
`categories`; scatter and bubble instead require aligned finite per-series
`xValues` and Y `values`. Bubble additionally requires one positive
`bubbleSize` per point.

```ts
slide.charts.add("area", {
  categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "Revenue", values: [42, 53, 68], fill: "#0EA5E9" }],
  xAxis: { title: "Quarter" },
  yAxis: { title: "Revenue", min: 0, max: 80 },
});

slide.charts.add("doughnut", {
  categories: ["North", "Central", "South"],
  series: [{ name: "Share", values: [52, 31, 17] }],
  dataLabels: { showCategoryName: true, showPercent: true, position: "outsideEnd" },
});

slide.charts.add("scatter", {
  series: [{
    name: "Portfolio",
    xValues: [10, 20, 34],
    values: [35, 68, 84],
    marker: { symbol: "diamond", size: 8, fill: "#8B5CF6" },
  }],
  xAxis: { title: "Reach" },
  yAxis: { title: "Return" },
});

slide.charts.add("bubble", {
  series: [{
    name: "Opportunity",
    xValues: [10, 20, 34],
    values: [35, 68, 84],
    bubbleSizes: [4, 9, 16],
    fill: "#F97316",
  }],
  xAxis: { title: "Reach" },
  yAxis: { title: "Return" },
});
```

Pie and doughnut do not accept axes, and percentage labels are limited to those
circular families. Markers are limited to line and scatter series; marker-only
scatter rejects a series line, so use `marker.line` for its border. Formula or
external-workbook data, stacked area, connected/smooth scatter, non-50%
doughnut geometry, bubble 3D/negative/custom-scale semantics, point overrides,
and topology changes fail closed.

Run the complete Agent workflow from this Skill root:

```sh
node examples/openchestnut-chart-families-workflow.mjs \
  output/chart-families.pptx \
  output/chart-families.png \
  output/chart-families.audit.json
```

It creates all four newly supported families, inventories their native XML,
imports and edits each one, exports and imports a second time, runs
inspect/verify, and produces a real PNG through `renderArtifact` with the
explicit Playwright renderer. See `api/references/charts.spec.md` for the full
bounded contract.

## Bounded Native Combo Chart

`combo` is native PPTX output only for one intentionally narrow profile: literal
clustered columns plus literal lines, at least one primary bar and one line, and
no topology changes after import. There are two canonical axis variants:

- Leave every line on the default primary axis group to use one shared
  category/value pair.
- Put **every** line at `axisGroup: "secondary"`, keep every bar primary, and
  provide `axes.secondary.category` plus `axes.secondary.value`. OpenChestnut
  writes that second pair at the top and right of the chart.

```ts
slide.charts.add("combo", {
  name: "revenue-margin",
  title: "Revenue and margin",
  position: { left: 90, top: 120, width: 1080, height: 420 },
  categories: ["Q1", "Q2", "Q3"],
  series: [
    { name: "Revenue", chartType: "bar", values: [42, 48, 57], color: "#2563EB" },
    {
      name: "Margin",
      chartType: "line",
      axisGroup: "secondary",
      values: [12, 15, 18],
      line: { fill: "#16A34A", width: 2 },
      marker: { symbol: "circle", size: 7 },
    },
  ],
  axes: {
    category: { title: "Quarter" },
    value: { title: "Revenue ($M)" },
    secondary: {
      category: { title: "Quarter" },
      value: { title: "Margin (%)", min: 0, max: 25 },
    },
  },
  dataLabels: { showValue: true, position: "top" },
});
```

Do not mix primary and secondary line series, put a bar on `axisGroup:
"secondary"`, omit either secondary axis when using secondary lines, or use
external/embedded workbook data, smooth lines, point overrides, per-series
labels, trendlines, or error bars. Those combinations fail closed rather than
being flattened or silently rebuilt.

## JSX Compose Equivalent

Use JSX when the slide is naturally rows, columns, grids, or overlays.

```tsx
/** @jsxRuntime automatic */
/** @jsxImportSource open-office-artifact-tool/presentation-jsx */

const frame = { left: 72, top: 64, width: 1136, height: 592 };

slide.compose(
  <column name="content-frame" width="fill" height="fill" gap={28}>
    <row width="fill" height={72} align="center" justify="between">
      <paragraph name="eyebrow" className="text-slate-500 text-sm font-bold">
        EDITABLE PRESENTATION
      </paragraph>
      <paragraph className="text-slate-400 text-sm">Q2 planning</paragraph>
    </row>
    <row width="fill" height="fill" gap={28} align="stretch">
      <column width={430} height="fill" gap={18}>
        <paragraph
          name="primary-heading"
          className="text-slate-950 text-5xl font-bold leading-tight"
        >
          Build decks with editable objects
        </paragraph>
        <paragraph className="text-slate-600 text-xl leading-relaxed">
          Rows, columns, grids, charts, tables, images, and text remain
          editable.
        </paragraph>
      </column>
      <box
        name="chart-frame"
        width="fill"
        height="fill"
        className="bg-white rounded-2xl shadow-sm"
        line={{ style: "solid", fill: "slate-200", width: 1 }}
      >
        <chart
          name="coverage-chart"
          chartType="bar"
          categories={["Rows", "Grid", "Tokens"]}
          series={[{ name: "Coverage", values: [3, 4, 5], fill: "accent1" }]}
          hasLegend={false}
          width="fill"
          height="fill"
        />
      </box>
    </row>
  </column>,
  { frame, baseUnit: 8 },
);
```

## Import And Edit PPTX

Load a PPTX, inspect for stable ids, render before/after evidence, make a
focused edit, re-inspect, and export the edited PPTX.

```ts
const presentation = await PresentationFile.importPptx(
  await FileBlob.load("input.pptx"),
);

const before = await presentation.inspect({
  kind: "slide,textbox,shape,image,table,chart,notes,thread,layout",
  search: "Revenue",
  maxChars: 8000,
});
console.log(before.ndjson);

const slide = presentation.resolve(slideIdFromInspect);
await writeBlob(
  "output/before-slide.png",
  await presentation.export({ slide, format: "png", scale: 1 }),
);
await fs.writeFile(
  "output/before-slide.layout.json",
  await (await slide.export({ format: "layout" })).text(),
);
await writeBlob(
  "output/before-montage.webp",
  await presentation.export({ format: "webp", montage: true, scale: 1 }),
);

const target = presentation.resolve(anchorIdFromInspect);
target.text.replace("Revenue", "Updated revenue outlook");

await writeBlob(
  "output/after-slide.png",
  await presentation.export({ slide, format: "png", scale: 1 }),
);
await fs.writeFile(
  "output/after-slide.layout.json",
  await (await slide.export({ format: "layout" })).text(),
);
await writeBlob(
  "output/after-montage.webp",
  await presentation.export({ format: "webp", montage: true, scale: 1 }),
);

const after = await presentation.inspect({
  target: { id: anchorIdFromInspect, beforeLines: 2, afterLines: 2 },
  kind: "textbox,shape,image,table,chart",
  maxChars: 3000,
});
console.log(after.ndjson);

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save("output/edited-deck.pptx");
```

## Bounded Imported Slide Name Edit

When one original imported slide has one explicit, unique native name, use the
shipped transaction rather than manipulating the package:

```ts
import { editPptxSlideName } from "../examples/openchestnut-slide-name-edit-workflow.mjs";

await editPptxSlideName({
  inputPath: "input.pptx",
  outputPath: "output/renamed.pptx",
  auditPath: "output/rename-audit.json",
  expectedName: "Go-no-go decision",
  replacementName: "Go decision: controlled rollout",
});
```

The workflow changes only `slide.name`, which crosses OpenChestnut as the
existing SlidePart's `p:cSld/@name`. It requires an exact source name, maps
the actual `presentation.xml` relationship order to the target part, checks
the saved target name, keeps every non-target part byte-identical, reimports,
and requires a byte-identical model SVG. Open XML SDK may canonicalize the
target SlidePart XML serialization, so the workflow does not promise lexical
byte identity for that one part. It rejects fallback-only names,
duplicate/missing target names, pending clones, and any unexpected package or
semantic change.

## Native Custom Shows

New presentations can author native playback subsets after their slides exist:

```ts
const overview = presentation.slides.add({ name: "Overview" });
const appendix = presentation.slides.add({ name: "Appendix" });
presentation.customShows.add("Board route", [overview, appendix]);

overview.shapes.add({
  position: { left: 44, top: 140, width: 300, height: 48 },
  text: [{ runs: [{ text: "Open board route", link: { customShow: "Board route", returnToSlide: true } }] }],
});
```

For one canonical imported show, use the audited fixed-topology transaction:

```ts
import { editPptxCustomShow } from "../examples/openchestnut-custom-show-workflow.mjs";

await editPptxCustomShow({
  inputPath: "input.pptx",
  outputPath: "output/custom-show.pptx",
  auditPath: "output/custom-show-audit.json",
  expectedName: "Board route",
  replacementName: "Executive route",
  orderedSlideNames: ["Appendix", "Overview", "Appendix"],
});
```

Only the existing name and ordered membership may change. Show count/order,
facade/native identity, all SlideParts, and every other package part stay
fixed. The workflow reimports, verifies native XML and non-target shows,
compares slide model SVGs, and emits a byte-bound audit. Noncanonical lists are
preserved opaque and fail closed; see
`artifact_tool/api/references/custom-shows.spec.md`.

## Native PowerPoint Sections

Sections are native presentation-wide groups, not custom-show routes: together
they must partition every slide once and in deck order. Add all slides first,
then author the complete group list:

```ts
const opening = presentation.slides.add({ name: "Opening" });
const evidence = presentation.slides.add({ name: "Evidence" });
const decision = presentation.slides.add({ name: "Decision" });

presentation.sections.add("Context", [opening, evidence]);
presentation.sections.add("Decision", [decision]);
```

OpenChestnut writes `p14:sectionLst` under the documented PowerPoint section
extension in `ppt/presentation.xml`. A source-free section gets a deterministic
native GUID unless `nativeId` is supplied. For a canonical imported list,
inspect with `presentation.inspect({ kind: "section" })`, resolve its stable
`section/...` ID (or call `presentation.sections.getItem(...)`), and change an
existing name or boundary only:

```ts
const context = presentation.sections.getItem("Context");
context.name = "Background";
context.setSlides([opening]);
presentation.sections.getItem("Decision").setSlides([evidence, decision]);
```

The imported section count/order, facade IDs, and native GUIDs remain fixed.
The resulting groups still must be the exact ordered partition. New/deleted/
reordered imported sections, pending slide insertion/deletion/duplicate, and
irregular or extension-bearing native graphs fail closed; opaque graphs remain
unchanged instead of being reconstructed. Reimport and inspect the output, then
run render QA if native tools are available. See
`artifact_tool/api/references/sections.spec.md` for the exact contract.

## Bounded Slide Transitions

For source-free slides, use one direct `fade` or directional `push` transition:

```ts
slide.setTransition({
  effect: "fade",
  speed: "medium",
  advanceOnClick: true,
  advanceAfterMs: 4_000,
});
```

`push` accepts `direction: "left" | "up" | "right" | "down"`; fade rejects a
direction. Speed defaults to `medium`, click advancement to `true`, and the
optional timer is an integer from `0` through `86400000` milliseconds. Inspect
an imported deck with `kind: "transition"`, resolve `${slide.id}/transition`,
then use `transition.set(...)` or `transition.clear()` only when its capability
reports `editable: true`. A transition-absent imported slide is deliberately
not addable, and timing/sound/extension/other-effect graphs stay opaque and
fail closed rather than being reconstructed. Reimport after export. Static
render QA proves visible slide content only; use a PowerPoint/native-host lane
for playback QA. See
`artifact_tool/api/references/transitions.spec.md` for the complete contract.

## Bounded Imported Slide Duplicate

When an Agent needs an exact source-bound copy of one imported slide rather
than a reconstructed slide, use the shipped transaction. It is deliberately
smaller than the full `slide.duplicate()` codec profile. By default, the
selected source must have one explicit unique name and no NotesSlide or
legacy-comments leaf.

```ts
import { duplicatePptxSlide } from "../examples/openchestnut-slide-duplicate-workflow.mjs";

await duplicatePptxSlide({
  inputPath: "input.pptx",
  outputPath: "output/source-with-copy.pptx",
  auditPath: "output/clone-audit.json",
  expectedName: "Unique source slide name",
});
```

The transaction supports the same canonical leaves used by the bounded clone
profile, including recognized literal-data charts whose unique ChartParts have
no child/external/hyperlink/data relationship, eligible top-level OLE frames
whose uniquely inbound internal XLSX package has no child graph and whose
preview is one internal ImagePart, canonical top-level SmartArt frames whose
single `dgm:relIds` root binds exactly four relationship-free data/layout/
quick-style/colors parts, canonical top-level `p:contentPart` objects whose one
internal `customXml` relationship binds a closed standard InkML part, canonical
top-level embedded-MP4 pictures whose paired video/media relationships uniquely
bind one closed `video/mp4` payload plus one poster ImagePart, and
straight/elbow connectors whose present endpoints
resolve inside the source slide tree and canonical run-level clicks to an
external absolute URI, a retained internal SlidePart, or a supported relative
slide action, plus a relationship-free custom-show action whose native ID must
resolve in the canonical presentation-wide catalog. It maps source `presentation.xml`
relationships, calls `slide.duplicate()`, requires the clone immediately after
the original, and accepts only the necessary presentation/content-type
topology changes plus one new SlidePart, its relationship part, and one
distinct byte-copied ChartPart for every accepted chart plus one distinct
byte-copied EmbeddedPackagePart for every accepted OLE workbook plus four
distinct typed diagram parts for every accepted SmartArt frame. The clone
also receives one distinct byte-identical SDK `CustomXmlPart` for every
accepted InkML object and one distinct byte-identical SDK `MediaDataPart` for
every accepted embedded video. For OLE and media frames it shares only the immutable preview
ImagePart. It retains every slide-local relationship ID; ordinary embedded
ImageParts are also deliberately shared. Every
retained source part, including the source SlidePart, must remain
byte-identical. It requires exact source/clone chart and hyperlink `r:id`, URI,
target SlidePart, and action-only inventories with no orphan relationship. For
a custom-show action it additionally proves exact native ID/return policy,
zero relationship edge, unchanged show identity/membership, and that the clone
was not silently added to the route. It then
reimports and verifies source/clone structural semantics, connector/link
bindings, chart semantics, independent OLE workbook bindings, and model-render
equivalence. SmartArt validation additionally requires the same slide-local
relationship IDs/types, standard content types, byte-identical payload hashes,
and disjoint source/clone part paths after reimport. InkML validation requires
the same evidence for one `application/inkml+xml` part with an `ink` document
element in the standard namespace and no child graph. Embedded-MP4 validation
requires the same slide-local video/media IDs to point at one distinct clone
payload with equal bytes/hash, the same shared poster path, exactly two inbound
media edges, and no child graph. Native pixels validate the poster, not video
playback. The model comparison ignores
fresh `data-*-id` inspection locators, so it does not claim lexical equality
for the new clone XML. Ambiguous names, notes/comments, unsupported graph
leaves, formula/external-data/embedded-workbook/connected/orphan charts,
shared/external/non-XLSX/nested/relationship-bearing/replacement-pending OLE
graphs, nested/incomplete/mistyped/external/relationship-bearing SmartArt,
non-top-level/extension-bearing/mistyped/non-InkML-root/connected content parts,
linked/shared/non-MP4/nested/multi-binding/connected media graphs,
shape-level/hover/unknown/orphan links, malformed/relationship-bearing/dangling
custom-show actions, unresolved endpoints, and unexpected package changes reject without
promoting an output or audit.

See `api/references/smartart-clone.spec.md` for the exact diagram-part roles,
failure modes, and audit evidence.
See `api/references/inkml-content-part-clone.spec.md` for the exact InkML
CustomXmlPart contract and its read-only boundary.
See `api/references/embedded-video-clone.spec.md` for the exact embedded-MP4
relationship, copy, poster-sharing, and playback-validation boundary.

### Edit One Imported SmartArt Plain Node

This is separate from cloning and intentionally much smaller than SmartArt
authoring. An imported top-level SmartArt object exposes `diagramText` only
when the closed four-part graph has a proven DiagramDataPart with direct
`dgm:t > a:p > a:r > a:t` document nodes. `editable` stays false; only the
returned existing `modelId` can receive a replacement string:

```ts
const diagram = presentation.slides.getItem(0).nativeObjects.items.find(
  (object) => object.nativeKind === "diagram" && object.diagramText,
);
const node = diagram?.diagramText.nodes.find((item) => item.text === "Before");
if (!diagram || !node) throw new Error("Expected canonical SmartArt node was not found.");
diagram.setDiagramNodeText(node.id, "After");
```

For source protection, exact target selection, no-overwrite output, package
scope validation, reimport, and audit, use the shipped workflow:

```ts
import { editPptxSmartArtNodeText } from "../examples/openchestnut-smartart-text-edit-workflow.mjs";

await editPptxSmartArtNodeText({
  inputPath: "input/source.pptx",
  outputPath: "output/edited.pptx",
  auditPath: "output/edited.audit.json",
  objectName: "Closed SmartArt",
  nodeId: "{B31B1833-2B65-4D6B-B3D4-9B3988427B21}",
  expectedText: "Before",
  replacementText: "After",
});
```

The C# codec re-proves the source hash, node IDs/order, and graph before it
rewrites only the bound DiagramDataPart. It preserves the graphic frame,
relationship IDs, layout, quick-style, colors, geometry, and every non-data
part. Multiple runs, fields, breaks, connected/nested graphs, node topology
changes, layout/style/color edits, and raw XML mutation fail closed. The
workflow's model verification is structural evidence; use LibreOffice/Poppler
when delivery needs native render review.

For the one opt-in closed-leaf profile, set `allowClosedLeaves: true` explicitly:

```ts
await duplicatePptxSlide({
  inputPath: "input.pptx",
  outputPath: "output/source-with-copy.pptx",
  auditPath: "output/clone-audit.json",
  expectedName: "Unique source slide name",
  allowClosedLeaves: true,
});
```

This is not a general relationship-graph clone. It accepts at most one
canonical NotesSlide whose only relationships are its immutable NotesMaster and
back-reference to the source SlidePart, plus at most one canonical legacy
SlideCommentsPart with no child relationship graph and one immutable
presentation-wide CommentAuthorsPart. The audit records the extra new parts,
proves notes/comments XML is byte-identical, proves the cloned notes
back-reference now points to the clone, and verifies source/clone closed-leaf
semantics after reimport. Rich/modern comments or any extra relationship fail
closed without a fallback.

## Add Speaker Notes To A Notes-Absent Imported Slide

Inspect the target before mutation. A source-bound slide may add a NotesSlide
only when the codec reports a safely extensible package graph:

```ts
const evidence = presentation.inspect({ kind: "slide,notes", maxChars: 8000 });
console.log(evidence.ndjson);

const notes = presentation.resolve(`${slide.id}/notes`);
if (!notes.capability.sourceBound || notes.capability.partPresent || !notes.capability.addable) {
  throw new Error("The imported slide is not eligible for bounded notes creation.");
}
```

Use the shipped source-bound transaction for the actual edit:

```ts
import { addPptxSpeakerNotes } from "../examples/openchestnut-speaker-notes-add-workflow.mjs";

await addPptxSpeakerNotes({
  inputPath: "input.pptx",
  outputPath: "output/with-notes.pptx",
  auditPath: "output/with-notes.audit.json",
  slideName: "Speaker notes target",
  notes: "Lead with the evidence.\nClose with the requested decision.",
});
```

The workflow keeps the source immutable and promotes output only after exact
reimport, visible-semantics/model-SVG equality, semantic verification, and OPC
relationship audit. It reuses an existing single NotesMaster byte-for-byte or
creates one canonical NotesMaster that shares an existing SlideMaster
ThemePart. The new NotesSlide owns only its NotesMaster relationship and a
back-reference to the selected SlidePart. Export re-proves capability from the
source package; there is no mutable-flag authority or silent fallback.
Inconsistent/multiple master graphs, an unusable theme, existing/rich notes, or
ambiguous targets fail closed. Compare source and output through
LibreOffice/Poppler before delivery because notes are nonvisual and every slide
must remain pixel-identical.

## Bounded Imported Title And Speaker-Notes Edit

When the source deck has one uniquely named slide, one uniquely named title
shape, and one canonical plain-text Notes part, use the shipped workflow rather
than reconstructing the slide or editing OOXML directly. The title may be an
ordinary editable shape or a concrete imported SlidePart placeholder with a
recognized local text body; placeholder text replacement preserves its native
identity, geometry, formatting, and layout binding:

```ts
import { editPptxTitleAndNotes } from "../examples/openchestnut-title-notes-edit-workflow.mjs";

await editPptxTitleAndNotes({
  inputPath: "input.pptx",
  outputPath: "output/edited.pptx",
  auditPath: "output/audit.json",
  slideName: "Go-no-go decision",
  titleShapeName: "approval-title",
  expectedTitle: "Decision: hold for legal review",
  replacementTitle: "Decision: approve controlled rollout",
  expectedNotes: "Lead with the pending legal condition.\nClose with the accountable owner.",
  replacementNotes: "Lead with the approved controls.\nClose with the accountable rollout owner.",
});
```

It requires a distinct output path and fails closed if the target is ambiguous,
the source text is not exact, the notes part is missing/rich, or reimport shows
an identity, geometry, formatting, background, slide-order, or slide-name
change. A placeholder whole-text replacement must retain the source newline and
inline topology; use an exact in-run `title.text.replace(...)` when that is the
narrower operation. It emits only model-render evidence; use the native render
route for final visual QA.

## Bounded Imported Rich Speaker-Notes Run Edit

For a known ordinary rich NotesSlide, preserve the paragraph/run tree and edit
one exact existing run through the dedicated transaction. Do not replace
`notes.text` or call `textFrame.setText()` on a multi-run imported body:

```ts
import { editPptxRichSpeakerNotes } from "../examples/openchestnut-rich-speaker-notes-edit-workflow.mjs";

await editPptxRichSpeakerNotes({
  inputPath: "input.pptx",
  outputPath: "output/edited.pptx",
  auditPath: "output/audit.json",
  slideName: "Go-no-go decision",
  titleShapeName: "approval-title",
  expectedTitle: "Decision: hold for legal review",
  replacementTitle: "Decision: approve controlled rollout",
  paragraphIndex: 0,
  runIndex: 1,
  expectedRunText: "the pending legal condition.",
  replacementRunText: "the approved control set.",
});
```

This fixed-topology transaction validates the imported source run text and
direct style, changes that one run, and reimports to prove paragraph/run counts,
bullets, auto-numbering, sibling runs, slide/title/notes identities, direct
background, slide names, and order. It records byte-bound provenance plus the
selected run/style in its audit. Fields, hyperlinks, picture bullets, notes
body/list styles, arbitrary rich reflow, and irregular NotesSlide graphs fail
closed rather than being flattened or silently reconstructed.

## Bounded Legacy Slide Comments

For a simple slide-level review annotation, use the standard legacy PPTX
profile directly. It has one author, one text item, and one explicit slide
coordinate; it is not a substitute for modern threaded comments.

```ts
const review = slide.comments.addThread(
  undefined,
  "Confirm the source before delivery.",
  {
    author: "Presentation Reviewer",
    created: "2026-07-18T09:30:00.000Z",
    position: { x: 1040, y: 84, unit: "px" },
  },
);

const evidence = presentation.inspect({
  kind: "comment",
  target: review.id,
  maxChars: 2000,
});
console.log(evidence.ndjson);
```

Use `undefined` for the target. Replies, resolution state, reactions, and
element/text-range anchors cannot be represented by legacy PresentationML and
therefore fail closed on canonical export.

For an imported deck, inspect the defensive creation capability before adding
a review annotation:

```ts
const capability = slide.comments.capability;
// { sourceBound, format, partPresent, addable }
```

Only a presentation with no legacy or Office 2021 comment graph anywhere can
advertise `{ sourceBound: true, format: "legacy", partPresent: false,
addable: true }`. Export re-proves the source package; changing JS data cannot
grant authority. Use the shipped source-protecting transaction for the common
single-comment review workflow:

```ts
import { addPptxLegacyReviewComment } from "../examples/openchestnut-legacy-comment-add-workflow.mjs";

await addPptxLegacyReviewComment({
  inputPath: "input.pptx",
  outputPath: "output/with-review.pptx",
  auditPath: "output/with-review.audit.json",
  slideName: "Imported review target",
  text: "Confirm the imported evidence before delivery.",
  author: "Review Owner",
  created: "2026-07-20T03:04:05Z",
  position: { x: 360, y: 240, unit: "px" },
});
```

OpenChestnut creates one canonical shared `CommentAuthorsPart` plus the target
slide's closed numbered `SlideCommentsPart`, allocates relationship IDs without
colliding with internal/external/hyperlink/data edges, reimports exact comment
semantics, and leaves slide XML and visible rendering unchanged. The workflow
audits the exact OPC additions, writes an immutable-source/no-overwrite report,
and fails before publication for an existing/mixed/connected comment graph or a
second add after reimport.

Recognized imported legacy comments are visible for inspection but must remain
unchanged. An unchanged canonical
legacy comments leaf may travel with `slide.duplicate()` through one
export/reimport boundary: its
clone-local `SlideCommentsPart` is byte-copied while the verified immutable
presentation-wide author catalog is shared. This is not an in-place comment
edit; both pending clone comments and reimported legacy comments remain
source-bound read-only. See `api/references/comments.md` for the complete
boundary.

## Bounded Office 2021 Comment Threads

Select the native modern wire family explicitly, then anchor one root plus
direct replies to a supported top-level element or shape text range:

```ts
const presentation = Presentation.create({ commentFormat: "modern" });
const slide = presentation.slides.add();
const title = slide.shapes.add({ id: "decision-title", text: "Customer evidence is ready" });

const thread = slide.comments.addThread({
  textMatch: { element: title, query: "Customer evidence", occurrence: 0 },
}, "Confirm the evidence.", {
  id: "{11111111-1111-4111-8111-111111111111}",
  nativeFormat: "modern",
  position: { x: 1234500, y: 2345600, unit: "emu" },
  comments: [{
    nativeId: "{11111111-1111-4111-8111-111111111111}",
    author: "Review Owner",
    person: {
      id: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
      name: "Review Owner",
      initials: "RO",
      userId: "review.owner@example.test",
      providerId: "None",
    },
    text: "Confirm the evidence.",
    created: "2026-07-19T02:55:00Z",
    status: "active",
  }],
});
```

Source-free threads may append direct replies. Recognized imported threads may
change only existing text and status (`thread.resolve()` / `thread.reopen()`).
Author/person/date identity, anchor/range, position, reply count/order,
relationships, and source hashes remain fixed. Reactions, task fields, rich
text, nested replies, connected comment parts, and unknown/nested anchors stay
opaque/source-bound and fail closed. Run
`examples/openchestnut-modern-comment-workflow.mjs` for the complete second-
import, package-inspect, model-render, and audit loop.

## Local Image Bytes

Use byte-backed images for embedded PPTX assets.

```ts
async function readImageBlob(imagePath: string): Promise<ArrayBuffer> {
  const bytes = await fs.readFile(imagePath);
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  );
}

const imageBytes = await readImageBlob("assets/product.png");
slide.images.add({
  blob: imageBytes,
  contentType: "image/png",
  alt: "Product screenshot",
  fit: "cover",
  position: { left: 720, top: 96, width: 420, height: 280 },
  geometry: "roundRect",
  borderRadius: "rounded-xl",
});
```
