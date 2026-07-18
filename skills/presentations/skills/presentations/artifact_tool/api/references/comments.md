# Comments

Presentation comments use people, threads, replies, reactions, and thread state.

## OpenChestnut codec boundary

The richer workflow below remains useful for model-side review planning, but
canonical PPTX export has one deliberately narrow interoperable profile:
standard legacy PresentationML comments. Use it only when one slide-level
annotation can be represented as one author, one text item, and one explicit
slide coordinate.

```ts
const slide = presentation.slides.add();
slide.comments.addThread(undefined, "Confirm the source before delivery.", {
  author: "Presentation Reviewer",
  created: "2026-07-18T09:30:00.000Z",
  position: { x: 1040, y: 84, unit: "px" },
});
```

Pass `undefined` as the target. A legacy `p:cm` has no element or text-range
anchor, reply graph, reaction, or resolved state. Imported comments that match
this profile are visible to `inspect`/`resolve` and are byte-preserved only
while unchanged. Do not add replies, resolve/reopen a thread, attach it to an
element or text range, or replace imported content: those requests fail closed.
Modern threaded-comment graphs remain opaque and source-bound. The remaining
reference examples describe richer workflow material; they do not claim that
the narrow legacy codec can serialize modern comment semantics.

The narrow imported-slide `slide.duplicate()` profile may carry one unchanged
canonical legacy comments part. It creates a fresh slide-local comments part by
copying its XML bytes and reuses the original immutable presentation-wide author
catalog; neither part may have a connected relationship graph. This is only a
preservation primitive: the pending clone must not change its comments before
export/reimport, and imported legacy comments remain read-only after it.

## Current Author

```ts
presentation.comments.setSelf({
  displayName,
  initials,
  email,
});
```

## Person Inline Type

```ts
type PersonConfig = {
  id?: string;
  displayName: string;
  initials?: string;
  email?: string;
  avatarUrl?: string;
  userId?: string;
  providerId?: string;
};
```

## Element Thread

```ts
const thread = presentation.comments.addThread({ element }, bodyText, {
  position,
});

const reply = thread.addReply(replyText);
reply.toggleReaction(reactionText);
thread.resolve();
thread.reopen();
```

## Thread Inline Types

```ts
type CommentTarget =
  | { slide: Slide }
  | { element: Shape | ImageElement | Table | ChartElement }
  | { textRange: TextRange }
  | { textMatch: { element: Shape; query: string; occurrence?: number } };

type ThreadAddOptions = {
  author?: Person | PersonConfig | { id: string };
  createdAt?: string;
  position?: { x: number; y: number; unit?: "px" | "emu" };
};
```

## Slide And Text Threads

```ts
presentation.comments.addThread({ slide }, bodyText, { position });

presentation.comments.addThread(
  { textMatch: { element, query, occurrence } },
  bodyText,
);
```

## Cookbook

```ts
// Add a review comment to exact text.
presentation.comments.setSelf({
  displayName: "Presentation Reviewer",
  initials: "PR",
  email: "reviewer@example.com",
});

const thread = presentation.comments.addThread(
  { textMatch: { element: titleShape, query: "Q4", occurrence: 0 } },
  "Check whether this should be FY2026 Q1.",
);
thread.addReply("Leaving this unresolved for owner review.");
```

```ts
// Resolve an imported thread after an edit.
const self = presentation.comments.setSelf({
  displayName: "Presentation Editor",
  initials: "PE",
  email: "editor@example.com",
});

const thread = presentation.resolve(threadAnchorId);
thread.addReply("Updated the chart title and source note.", { author: self });
thread.resolve(self);

const verified = await presentation.inspect({
  kind: "thread",
  target: { id: threadAnchorId, beforeLines: 0, afterLines: 2 },
  maxChars: 2000,
});
```
