# Native grouped shapes

Use `slide.groups.add(...)` when several slide objects must move, scale,
inspect, and round-trip as one recursive DrawingML group. OpenChestnut writes a
real `p:grpSp`; it does not flatten children or synthesize a visual-only parent.

## Coordinate contract

A group owns two rectangles:

- `position`: the group's outer frame in its parent group's coordinates, or in
  slide pixels for a top-level group. It maps to DrawingML `a:off/a:ext`.
- `childFrame`: the local coordinate rectangle used by direct children. It maps
  to `a:chOff/a:chExt` and may use negative `left` or `top` values. Width and
  height must remain positive.

Every direct child's position, connector endpoints, and nested-group frame use
the owning group's local coordinates. A nested group starts a new coordinate
space. `inspect(...)` and layout JSON report both local and absolute frames.

## Source-free authoring

This example puts the connector before its nodes in native z-order while using
explicit IDs and endpoints for stable targeting:

```js
const group = slide.groups.add({
  id: "workflow-group",
  name: "workflow-group",
  position: { left: 120, top: 110, width: 1040, height: 430 },
  childFrame: { left: -80, top: 40, width: 1280, height: 540 },
  children: [
    {
      kind: "connector",
      id: "model-to-codec",
      name: "model-to-codec",
      startTargetId: "model-node",
      endTargetId: "codec-node",
      start: { x: 300, y: 270 },
      end: { x: 450, y: 270 },
      line: { fill: "#475569", width: 2, endArrow: "triangle" },
    },
    {
      kind: "shape",
      id: "model-node",
      name: "model-node",
      geometry: "roundRect",
      position: { left: 0, top: 200, width: 300, height: 140 },
      fill: "#DBEAFE",
      line: { fill: "#2563EB", width: 2 },
      text: "Model",
    },
    {
      kind: "shape",
      id: "codec-node",
      name: "codec-node",
      geometry: "roundRect",
      position: { left: 450, top: 200, width: 300, height: 140 },
      fill: "#DCFCE7",
      line: { fill: "#16A34A", width: 2 },
      text: "OpenChestnut",
    },
    {
      kind: "groupShape",
      id: "qa-group",
      name: "qa-group",
      position: { left: 900, top: 180, width: 300, height: 220 },
      childFrame: { left: 0, top: 0, width: 300, height: 220 },
      shapes: [{
        name: "qa-node",
        geometry: "rect",
        position: { left: 20, top: 30, width: 260, height: 150 },
        fill: "#F3E8FF",
        text: "Render + verify",
      }],
    },
  ],
});
```

The bounded native profile supports recursive groups containing modeled
shapes/textboxes, straight or polyline connectors, embedded images, fixed-grid
tables, and literal bar/line/pie charts. Child IDs are owner-local in the wire,
but remain globally resolvable through the presentation facade.

## Inspect, resolve, and edit

```js
const snapshot = presentation.inspect({
  kind: "slide,groupShape,shape,connector,image,table,chart",
  search: "workflow-group",
  maxChars: 8000,
});

const importedGroup = presentation.resolve(groupIdFromInspect);
importedGroup.position.left += 24;
importedGroup.childFrame.left = -40;
presentation.resolve("model-node").text.set("Updated model");
```

For a canonical imported group, OpenChestnut permits bounded semantic edits to
the group name, outer frame, child frame, and supported descendants. The child
tree topology is fixed: do not add, remove, reorder, or change the native kind
of an imported child. Export rejects with
`presentation_group_topology_changed` instead of rebuilding or flattening the
group.

## Fail-closed boundary

An imported group remains one opaque, read-only native object when its group
shell uses unmodeled fill/effect/lock/rotation/flip/extension semantics, or when
any descendant cannot be modeled safely. Inspect and preserve it unchanged.
Do not extract and re-add only the understood descendants; that would destroy
the original ownership, z-order, coordinate transform, and unknown graph.

After every grouped edit, run semantic verification, export/import once more,
render the whole slide, and inspect for the group plus its descendants.
