# Presentation Facade

## Create And Load

```ts
const presentation = Presentation.create({ slideSize });
const imported = Presentation.load(proto);
```

## Create Inline Type

```ts
type PresentationCreateOptions = {
  slideSize?: { width: number; height: number };
};
```

## Presentation Slide Collection

```ts
const slide = presentation.slides.add({ layout, layoutId });
const inserted = presentation.slides.insert({ after, layout, layoutId });
const byIndex = presentation.slides.getItem(slideIndex);
slide.moveTo(destinationIndex);
const duplicate = slide.duplicate();
slide.delete();
```

## Presentation Slide Collection Inline Types

```ts
type SlideAddOptions = {
  name?: string;
  layout?: string | SlideLayoutTemplate;
  layoutId?: string;
  background?: string | BackgroundConfig;
  notes?: string;
};

type SlideInsertOptions = SlideAddOptions & {
  after?: Slide | number | null;
};
```

`slides.add(options)` appends. `slides.insert({ after, ...options })` inserts
after an existing slide facade or its 0-based index; `after: null` inserts at
the beginning and omitting `after` appends. Both paths resolve a supplied
source-free layout transactionally and materialize its direct-frame text
placeholders. Unknown targets/layouts leave the collection unchanged.

Insertion remains source-free authoring only: inserting into an imported PPTX
would change its source-bound slide topology and is rejected at export rather
than silently reconstructing the deck.

A concrete imported SlidePart `p:sp/p:ph` with a recognized local text body may
replace existing characters through `shape.text.replace(...)` or a
newline-topology-preserving `shape.text.set(...)`. This component capability
is exposed for preflight as `shape.placeholder.textEditable === true`, but is
re-proved from the source binding during export and cannot be granted by
changing the model flag. It does not make the placeholder shape editable:
type/index, name, geometry,
formatting, layout binding, Master/Layout projections, and unmodeled XML remain
source-bound, and ambiguous topology changes fail closed.

`slide.moveTo(destinationIndex)` moves one existing slide to an existing
0-based deck index. On an imported PPTX it changes only
`ppt/presentation.xml`'s `p:sldIdLst` for the retained source SlideParts; it
neither rebuilds slide parts nor copies their relationship graphs.

`slide.delete()` returns `undefined`. It removes any non-final source-free
slide. On an imported PPTX it succeeds only for an isolated layout-only source
SlidePart with no outbound non-layout relationship, inbound relationship, or
presentation identity reference; it then removes the actual part and relation.

`slide.duplicate()` returns a new adjacent `Slide` only under the bounded
imported shape/inline-table/closed-literal-chart/closed-SmartArt/closed-InkML/embedded-image/connector/
recursive-group layout-leaf profile. Its unchanged graph may contain canonical
simple shapes, canonical inline fixed-grid tables, recognized literal-data
charts, eligible top-level embedded-XLSX OLE frames, canonical embedded
rectangular images, bounded straight/elbow
connectors, canonical top-level SmartArt frames whose exact `dm/lo/qs/cs`
relationships bind four closed standard diagram parts,
canonical top-level `p:contentPart` objects whose one `customXml` relationship
binds a closed standard InkML part, plus recursively
canonical groups whose descendants contain only
those same leaf kinds,
exactly one layout relationship, image relationships bound only by those
pictures, canonical run-level external/internal/relative-action click links
whose exact relationship IDs and retained targets are copied, relationship-free
custom-show actions resolved through the fixed native-ID catalog, and optionally one closed `NotesSlide -> NotesMaster` /
back-to-source-slide leaf plus one canonical legacy `SlideCommentsPart` leaf.
Each chart frame must uniquely bind one numbered ChartPart with no child,
external, hyperlink, or data relationship. Each OLE frame must uniquely bind
one closed internal XLSX package plus one immutable preview ImagePart. Each
SmartArt frame must bind exactly the four closed standard diagram parts and
must not be nested. Each InkML object must bind exactly one relationship-free
`application/inkml+xml` CustomXmlPart whose root has the standard InkML
namespace and must not be nested. It creates a distinct native
SlidePart and presentation relationship, byte-copies each accepted chart into
a distinct clone-local ChartPart, each OLE workbook into an independent
EmbeddedPackagePart, and each accepted SmartArt frame into four distinct typed
diagram parts, and each accepted InkML payload into a distinct SDK-typed
CustomXmlPart. It shares the verified layout, immutable ImageParts, NotesMaster,
and presentation-wide
`CommentAuthorsPart`, copies accepted NotesSlide and SlideComments XML
byte-for-byte, and repoints only the notes leaf at the clone while preserving
the origin part. The comments leaf and author catalog must have no child,
external, hyperlink, or data relationship graph. Accepted tables are inline-only
and cannot introduce a fill, link, or another package relationship. Accepted groups add no relationship themselves, and every nested picture must consume one exact verified ImagePart relationship. The clone is intentionally
read-only until it has crossed one export/reimport boundary; it then imports as
its own source-bound slide with independent ChartParts, OLE workbook packages,
SmartArt parts, and InkML parts. A chart that advertises
the ordinary fixed-topology edit capability can use that path without affecting
the origin, with legacy comments still read-only. Imported add,
repeat/mutated clone, immediate clone edit, rich/connected comments, and every
unsupported-connector/formula-or-external-data-or-embedded-workbook-or-connected-or-orphan-chart/OLE/noncanonical-or-connected-SmartArt/noncanonical-or-connected-InkML/shape-level-or-unmodeled-hyperlink/malformed-or-dangling-custom-show/section/extension, external-or-irregular-image,
or otherwise connected clone/delete graph fails closed.

A canonical custom-show run action is copied as an inline, relationship-free
reference to the same stable native show ID. Cloning never changes the
presentation-wide show list or inserts the clone into membership. Custom-show
deletion/topology changes remain outside this profile.

## Custom Shows

```ts
const show = presentation.customShows.add("Board route", [slide1, slide3]);
const byName = presentation.customShows.getItem("Board route");
show.name = "Executive route";
show.setSlides([slide3, slide1]);

slide1.shapes.add({
  position: { left: 80, top: 80, width: 320, height: 48 },
  text: [{ runs: [{ text: "Open route", link: { customShow: "Executive route", returnToSlide: true } }] }],
});
```

Source-free export writes a native `p:custShowLst`. A canonical imported list
permits only existing-show name and ordered membership edits; count/order,
facade IDs, and native IDs remain source-bound. See
[`custom-shows.spec.md`](./custom-shows.spec.md) for budgets, opaque graphs, and
the stable-identity run-link contract, and the audited workflow.

## Discover And Edit

```ts
const snapshot = await presentation.inspect({
  kind,
  search,
  maxChars,
});

const target = presentation.resolve(anchorId);
```

`inspect` returns stable anchor ids for slides, shapes, images, tables, charts,
custom shows, text ranges, speaker notes, and comment threads. `resolve` maps a
returned anchor id to the matching facade. Layout records expose `layoutId` for
search and comparison; pass only model-returned IDs such as `pr/`, `sl/`,
`custom-show/`, `sh/`, `im/`, `tb/`, `ch/`, `nt/`, `th/`, and `tr/` anchors to
`resolve`.

## Inspect Inline Type

```ts
type PresentationInspectOptions = {
  target?: { id: string; beforeLines?: number; afterLines?: number };
  kind?: string; // e.g. "slide,textbox,shape,image,table,chart,notes,thread,layout"
  include?: string;
  exclude?: string;
  search?: string;
  maxChars?: number;
};
```

## Help

```ts
const help = presentation.help(query, {
  search,
  include,
  maxChars,
});
```

## Help Inline Type

```ts
type PresentationHelpOptions = {
  search?: string;
  include?: string[]; // common: ["index", "examples", "notes"]
  maxChars?: number;
};
```

## Font Inventory

`presentation.fontFamilies` returns a fresh sorted, case-insensitively
deduplicated array of explicitly used text and bullet font families. Theme
tokens such as `+mj-lt` are not reported as installed font names.

```ts
const typefaces = presentation.fontFamilies;
```

## Presentation View

Use `presentation.view` to control gridlines and imported PowerPoint guides in
an editor preview.

```ts
presentation.view.showGridlines();
presentation.view.showGuides();

const gridlinesVisible = presentation.view.gridlinesVisible;
const guidesVisible = presentation.view.guidesVisible;
const horizontalGridSpacingEmu = presentation.view.gridSpacingCxEmu;
const verticalGridSpacingEmu = presentation.view.gridSpacingCyEmu;

presentation.view.hideGridlines();
presentation.view.hideGuides();

const nextGridlineState = presentation.view.toggleGridlines();
const nextGuideState = presentation.view.toggleGuides();
```

Visibility is local editor state. `toProto()` serializes guide visibility as
hidden, while imported grid spacing, snap settings, and guide definitions stay
read-only and source-bound in canonical PPTX export.

## Export And Serialized Data

```ts
const imageBlob = await presentation.export({ slide, format, scale });
const montageBlob = await presentation.export({
  format: "webp",
  montage: true,
  scale: 1,
});
const layoutBlob = await slide.export({ format: "layout", scale });
const proto = presentation.toProto();
```

`toProto()` returns presentation data for host adapters. File export and local
resource resolution belong to host adapter docs.

## Export Inline Type

```ts
type PresentationExportOptions = {
  slide?: Slide;
  format?: "png" | "jpeg" | "webp" | "layout";
  width?: number;
  height?: number;
  scale?: number;
  quality?: number;
  montage?:
    | boolean
    | {
        format?: "png" | "jpeg" | "webp";
        width?: number;
        slideWidth?: number;
        padding?: number;
        gap?: number;
        background?: string;
        columns?: number;
      };
};
```

## Scripts

```ts
const result = presentation.scripts.run(scriptKind, scriptOptions);
```

Scripts provide high-level authoring recipes. Use `presentation.help(...)` to discover available script keys and option shapes.

## Cookbook

```ts
// New deck skeleton: create, set theme, add slides, render checks.
const presentation = Presentation.create({
  slideSize: { width: 1280, height: 720 },
});
presentation.theme.colorScheme = {
  name: "Clean Product",
  themeColors: {
    accent1: "#2563eb",
    accent2: "#0f766e",
    accent3: "#f59e0b",
    accent4: "#dc2626",
    accent5: "#7c3aed",
    accent6: "#16a34a",
    bg1: "#ffffff",
    bg2: "#f8fafc",
    tx1: "#0f172a",
    tx2: "#475569",
    dk1: "#000000",
    dk2: "#1e293b",
    lt1: "#ffffff",
    lt2: "#e2e8f0",
    hlink: "#2563eb",
    folHlink: "#7c3aed",
  },
};

const first = presentation.slides.add();
const second = presentation.slides.add();
const third = presentation.slides.add();

await presentation.export({
  slide: first,
  format: "png",
  scale: 1,
});
const snapshot = await presentation.inspect({
  kind: "deck,slide,textbox,chart,table",
  maxChars: 6000,
});
```

```ts
// Existing deck: inspect first, then resolve exact anchors.
const before = await presentation.inspect({
  kind: "slide,textbox,shape,image,table,chart,notes,thread,layout",
  search: "Customer growth",
  maxChars: 8000,
});
const target = presentation.resolve(anchorIdFromBefore);
```
