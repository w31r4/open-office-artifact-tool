# Masters

The source-free OpenChestnut PPTX codec creates exactly one canonical Slide
Master. It owns a direct background, bounded master text styles, and optional
direct-frame text placeholders. Layouts link to that master through `masterId`;
slides then bind a layout and materialize its placeholders.

## Configure the canonical master

```js
const presentation = Presentation.create({
  master: {
    name: "Brand master",
    background: { fill: "#F8FAFC", mode: "solid" },
    textParagraphStyles: {
      title: { 0: { alignment: "left", style: { fontSize: 30, bold: true } } },
      body: { 0: { style: { fontSize: 18, color: "#334155" } } },
    },
    placeholders: [{
      type: "title",
      index: 0,
      position: { left: 72, top: 56, width: 1136, height: 80 },
    }],
  },
});

presentation.master.setBackground({ fill: "accent1", mode: "reference", index: 1001 });

const layout = presentation.layouts.add({
  name: "Title and body",
  type: "obj",
  masterId: presentation.master.id,
});
```

`presentation.master.clearBackground()` removes only the direct master
background. The default presentation theme is used for source-free output;
`master.setTheme(...)` remains a preview-only operation and fails closed at
PPTX export.

## Supported source-free boundary

- Exactly one master.
- Direct RGB/theme solid or style-reference background.
- Bounded `title`, `body`, and `other` paragraph-style levels.
- Direct-frame `title`, `body`, `ctrTitle`, and `subTitle` text placeholders.
- Layout types `blank`, `title`, `titleOnly`, and `obj`.

`presentation.masters.add(...)` is available for model inspection, but adding
a second master deliberately makes source-free export fail closed. Use
`Presentation.create({ master })` or `presentation.master` for the canonical
authoring path.

## Imported decks

Imported masters, layouts, themes, guides, and placeholders are exposed for
inspect/resolve and preserved unchanged with their source package. They are not
a generic editable template graph: changing their semantic master/layout data,
adding/deleting placeholders, rebinding slides, or introducing another master
fails closed rather than rebuilding the template graph.
