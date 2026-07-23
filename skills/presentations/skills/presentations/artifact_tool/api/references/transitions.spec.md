# PowerPoint Slide Transitions

This is a deliberately small `p:transition` contract for Agent workflows. It
controls the transition between slides; it is not a PowerPoint timing or
animation engine, and a static PNG/PDF render cannot prove slideshow playback.

## Source-Free Authoring

```ts
const fade = presentation.slides.add({
  name: "Opening",
  transition: {
    effect: "fade",
    speed: "medium",
    advanceOnClick: true,
    advanceAfterMs: 4_000,
  },
});

const decision = presentation.slides.add({ name: "Decision" });
decision.setTransition({
  effect: "push",
  direction: "left",
  speed: "fast",
  advanceOnClick: false,
});
```

Only these semantic shapes are authored:

| Field | Supported values | Default |
| --- | --- | --- |
| `effect` | `"fade"` or `"push"` | required |
| `direction` | `"left"`, `"up"`, `"right"`, `"down"` for `push` only | `"left"` |
| `speed` | `"slow"`, `"medium"`, `"fast"` | `"medium"` |
| `advanceOnClick` | boolean | `true` |
| `advanceAfterMs` | integer `0..86400000` | omitted |

`fade` rejects `direction`. The codec writes one direct `p:transition` with
explicit `spd` and `advClick`, an optional numeric `advTm`, and exactly one
empty `p:fade` or `p:push/@dir` child. `clearTransition()` removes that direct
element.

## Inspect, Resolve, And Edit

```ts
const inspection = await presentation.inspect({ kind: "slide,transition" });
const transition = presentation.resolve(`${fade.id}/transition`);

if (transition.capability.editable) {
  transition.set({ effect: "push", direction: "right", speed: "slow" });
}
```

Each slide emits a stable `${slide.id}/transition` inspect record, even when it
is not configured. Its capability is evidence for routing, not permission to
edit an arbitrary package:

```ts
{
  sourceBound: boolean,
  partPresent: boolean,
  editable: boolean,
  addable: boolean,
}
```

Source-free slides are addable and editable. A source-bound slide is editable
only when its existing transition fits this exact profile. An imported slide
with no transition is intentionally **not addable**: adding a new transition
to an arbitrary package is not silently treated as a harmless patch.

## Imported and Clone Boundary

The C# Open XML SDK codec accepts one direct `p:transition` only when it has
the explicit attributes and single child described above. `p14:dur`, sound
actions, extension lists, extra/unknown attributes, multiple children, other
effect names, malformed timers, and any broader timing/animation graph remain
opaque. They are preserved byte-for-byte when unrelated supported edits occur,
and `setTransition()` or `clearTransition()` rejects them.

The strict imported `slide.duplicate()` profile may copy one unchanged
canonical direct transition with its SlidePart. It does not copy or interpret a
timing tree, sound relationship, or extension graph. Re-import after export
before making another semantic change.

## Verification

After a mutation, export, re-import, and inspect the transition record again.
For a source-bound edit, verify that the intended SlidePart transition changed
and unrelated package parts stayed within the operation's declared scope.
LibreOffice/Poppler visual QA can still prove that visible static slide content
did not regress, but cannot certify the transition playback itself. Use a real
PowerPoint/native-host slideshow QA lane when playback timing or host-specific
effects matter.
