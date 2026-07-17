# Embedded XLSX workbooks

An ordinary imported PPTX may contain an Excel workbook behind a top-level OLE
object. OpenChestnut exposes a narrow payload-only operation when the source
graph proves all of the following:

- one top-level native object contains exactly one `p:oleObj`;
- its owner-local relationship resolves to one internal `/package` part;
- that part has the standard XLSX content type;
- the workbook part is not shared by another relationship; and
- the part path, relationship ID, content type, and source SHA-256 still match
  the imported binding.

The native object remains read-only for its XML shell, name, position,
relationship graph, preview image, and unrelated parts. Its inspect record
advertises only `embeddedWorkbook` in `editableFields`.

## Read, inspect, and replace

```js
import {
  PresentationFile,
  SpreadsheetFile,
} from "open-office-artifact-tool";

const presentation = await PresentationFile.importPptx(sourcePptx);
const nativeObjects = presentation.inspect({
  kind: "nativeObject",
  search: "embeddedWorkbook",
  maxChars: 8000,
});

const oleObject = presentation.resolve(nativeObjectIdFromInspect);
const currentXlsx = oleObject.getEmbeddedWorkbook();
const workbook = await SpreadsheetFile.importXlsx(currentXlsx);

workbook.worksheets
  .getItem("Embedded")
  .getRange("B2")
  .values = [["Updated by the agent"]];

const replacementXlsx = await SpreadsheetFile.exportXlsx(workbook);
oleObject.replaceEmbeddedWorkbook(replacementXlsx);

const pending = oleObject.getEmbeddedWorkbook();
if (pending.metadata.pendingReplacement !== true) {
  throw new Error("Expected a pending embedded-workbook replacement");
}

const outputPptx = await PresentationFile.exportPptx(presentation);
const verified = await PresentationFile.importPptx(outputPptx);
```

`replaceEmbeddedWorkbook(...)` accepts `FileBlob`, `Uint8Array`, `ArrayBuffer`,
or another `ArrayBuffer` view. It copies at most 16 MiB defensively. Canonical
export then applies OPC budgets, opens the payload with the Microsoft Open XML
SDK, requires at least one worksheet, and runs Office 2021 validation before
replacing bytes in the bound embedded-package part. This is payload-only:
preserving the OLE shell, relationship topology, preview image, and unrelated
native parts is part of the checked export contract.

Always import the exported PPTX again and inspect the rebound native object.
Its source digest must describe the replacement bytes and
`replacementPending` must be false. Render the containing slide with the native
QA path as well, while remembering that the preserved preview image is not
regenerated from the new workbook.

## Fail-closed boundary

Do not use this operation to:

- create a new OLE object or embedding relationship;
- change the OLE program ID, icon, shell XML, frame, or preview;
- replace a shared, external, ambiguous, or non-XLSX package;
- mutate arbitrary preserved native parts; or
- bypass source-hash and relationship checks with package patching.

Malformed XLSX input fails with `invalid_presentation_ole_workbook` or a more
specific OPC budget/validation diagnostic. Changed bindings fail with
`presentation_ole_workbook_binding_mismatch`; unsupported graph shapes remain
opaque and read-only. There is no lossy reconstruction or silent fallback.
