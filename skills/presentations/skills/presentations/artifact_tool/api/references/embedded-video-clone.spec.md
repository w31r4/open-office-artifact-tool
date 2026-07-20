# Source-bound embedded-MP4 media clone

OpenChestnut imports a PowerPoint video picture as an opaque `nativeObject`
with `nativeKind === "media"`. The bounded profile below preserves one embedded
MP4 while cloning its slide. It is not video authoring, playback automation,
timing/trim editing, or a general audio/video graph copier.

An unchanged object may travel through `slide.duplicate()` only when all of
these preconditions hold:

- it is one direct top-level `p:pic` under the source slide's `p:spTree`, never
  nested in `p:grpSp`;
- the picture is a canonical rectangular poster frame accepted by the bounded
  image reader;
- `a:hlinkClick` contains the exact empty `r:id` media-action sentinel and
  `action="ppaction://media"`;
- one `a:videoFile/@r:link` and one
  `p14:media/@r:embed` under extension URI
  `{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}` resolve to the same internal
  `MediaDataPart`;
- one distinct poster `a:blip/@r:embed` resolves to an internal ImagePart;
- the video relationship uses the standard OOXML `video` type, the media
  relationship uses the Office 2007 `media` type, and the picture has no other
  relationship-valued attribute except the empty action sentinel;
- the shared payload is non-empty `video/mp4`, has a safe single-level
  `ppt/media/*.mp4` or Open XML SDK `media/*.mp4` path, has no child
  relationship graph, and has exactly those two inbound relationships from
  the owning slide; and
- the imported native object and independent OPC inspection agree on all
  three relationship IDs, both part paths, content types, and SHA-256 digests.

```ts
const source = presentation.slides.getItem(0);
const video = source.nativeObjects.items.find(
  (object) => object.nativeKind === "media",
);

if (!video || video.parts.length !== 2) {
  throw new Error("Source does not expose one closed embedded-MP4 picture.");
}

const clone = source.duplicate();
const output = await PresentationFile.exportPptx(presentation);
const rebound = await PresentationFile.importPptx(output);
```

Export retains the source SlidePart and relationship part byte-for-byte. The
clone keeps both slide-local media relationship IDs, but the Open XML SDK
allocates a distinct `MediaDataPart` and copies the exact MP4 bytes into it.
The immutable poster ImagePart remains shared under the same slide-local image
relationship ID. After second import, source and clone expose different MP4
paths with equal hashes and the same poster path.

For an Agent-facing transaction, use
`examples/openchestnut-slide-duplicate-workflow.mjs`. Its independent preflight
runs before semantic import. The audit records `operation.mediaParts`, exact
video/media/poster IDs, source and clone paths/digests, the allowed package
delta, `validation.package.mediaParts`, and
`sourceAndCloneMediaBindingsIndependent`. Model and native render checks prove
the visible poster remains equal; they do not claim media playback equivalence.
A failed precondition publishes neither PPTX nor audit.

The media object remains opaque and read-only after reimport. Audio, linked or
external media, non-MP4 payloads, shared payloads, nested media pictures,
multiple video/media bindings, alternate extension graphs, timing and animation
graphs, bookmarks, trim/loop/volume controls, poster replacement, transcoding,
and payload editing all fail closed. Preserve those inputs unchanged or route a
separately reviewed explicit OPC operation; do not flatten them into an image
or claim that a poster-only render validates playback.
