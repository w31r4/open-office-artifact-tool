# Embedded Office packages

An ordinary imported PPTX may contain an Office package behind a top-level OLE
object. OpenChestnut exposes narrow payload-only operations when the source
graph proves all of the following:

- one top-level native object contains exactly one `p:oleObj`;
- its owner-local relationship resolves to one internal `/package` part;
- that part has the standard XLSX content type;
- the workbook part is not shared by another relationship; and
- the part path, relationship ID, content type, and source SHA-256 still match
  the imported binding.

The legacy XLSX profile remains available through `getEmbeddedWorkbook()` and
`replaceEmbeddedWorkbook(...)`. The generic Office-package API is compatible
with that profile, and currently adds one explicitly bounded DOCX profile.
The native object remains read-only for its XML shell, name, position,
relationship graph, preview image, and unrelated parts. Its inspect record
advertises `embeddedWorkbook` for the XLSX profile or `embeddedOfficePackage`
for the DOCX profile.

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

## One source-bound embedded DOCX package

`getEmbeddedOfficePackage()` and `replaceEmbeddedOfficePackage(...)` are the
additive generic route. They retain the legacy XLSX behavior above, but the
only newly supported kind is a DOCX package with exact content type
`application/vnd.openxmlformats-officedocument.wordprocessingml.document`.
The source proof is deliberately the same narrow top-level OLE shape, plus an
exclusive inbound package relationship. It is not a generic OLE/container API.

For a straightforward Agent change, resolve exactly one native object, edit one
unique ordinary single-run paragraph in its embedded DOCX through the public
Document facade, then replace only that bound payload:

```js
import {
  DocumentFile,
  PresentationFile,
} from "open-office-artifact-tool";

const presentation = await PresentationFile.importPptx(sourcePptx);
const oleObject = presentation.resolve(nativeObjectIdFromInspect);
const currentDocx = oleObject.getEmbeddedOfficePackage();

if (currentDocx.metadata.officePackageKind !== "docx") {
  throw new Error("Expected the bounded DOCX OLE package profile");
}

const document = await DocumentFile.importDocx(currentDocx);
const matches = document.blocks.filter((block) =>
  block.kind === "paragraph" && block.text === "Source wording",
);
if (matches.length !== 1) throw new Error("Refusing an ambiguous embedded DOCX edit");
matches[0].text = "Approved wording";
if (matches[0].runs.length !== 1 || matches[0].runs[0].text !== "Source wording") {
  throw new Error("Refusing rich or fragmented embedded-DOCX text");
}
matches[0].runs[0].text = "Approved wording";

const replacementDocx = await DocumentFile.exportDocx(document);
oleObject.replaceEmbeddedOfficePackage(replacementDocx);

const pending = oleObject.getEmbeddedOfficePackage();
if (pending.metadata.pendingReplacement !== true) {
  throw new Error("Expected a pending embedded Office-package replacement");
}

const outputPptx = await PresentationFile.exportPptx(presentation);
const rebound = await PresentationFile.importPptx(outputPptx);
```

`replaceEmbeddedOfficePackage(...)` accepts `FileBlob`, `Uint8Array`,
`ArrayBuffer`, or another `ArrayBuffer` view, copies at most 16 MiB defensively,
and requires a DOCX `FileBlob` to retain that exact MIME type. Canonical export
re-proves the original part path, relationship ID, MIME type, source digest, and
unique inbound package ownership; it opens the replacement with the Microsoft
Open XML SDK, requires a WordprocessingDocument body, applies the normal OPC
budgets, and runs Office 2021 validation. It changes only the bound DOCX bytes.

Use `examples/openchestnut-ole-office-package-workflow.mjs` for the complete
auditable transaction. It protects the input, resolves exactly one object and
one single-run paragraph, requires an exact expected source string, checks that
the embedded DOCX is the only changed package part, reimports both PPTX and
DOCX, runs presentation verification/model render, and writes
source/output/binding hashes to its audit. It does not regenerate the preserved
OLE preview image.

## Duplicate an eligible OLE slide leaf

An unchanged eligible top-level OLE frame may travel with the bounded imported
`slide.duplicate()` transaction. This is not OLE authoring or a general OPC
graph copy. Clone preflight additionally requires exactly one embed node, one
internal preview picture, exactly the package `r:id` plus preview `r:embed`
relationship attributes, no child/external/hyperlink/data graph on the XLSX
package, and no second inbound package relationship.

On the first export, OpenChestnut preserves the source SlidePart and workbook
bytes, creates a distinct clone-local XLSX `EmbeddedPackagePart`, byte-copies
the workbook, retains the source slide-local package relationship ID, and shares
the immutable preview ImagePart. The pending clone must remain unchanged until
that export has been imported again. After reimport, the two native objects have
different `oleWorkbook.partPath` values and the same source digest; calling
`replaceEmbeddedWorkbook(...)` on the clone then changes only its independent
package.

Use `examples/openchestnut-slide-duplicate-workflow.mjs` for the auditable
transaction. Its independent package checks prove content type, unique inbound
ownership, empty child graph, exact source/clone bytes, distinct package paths,
shared preview binding, retained source parts, second import, and model-render
equivalence. `--allow-closed-leaves` controls NotesSlide/legacy-comments leaves,
not OLE; an eligible OLE workbook is part of the default canonical clone
profile.

The DOCX Office-package profile is intentionally not cloneable in this release.
The generic API is a source-bound replacement boundary, not evidence that a
second OLE relationship graph can be safely authored. Attempting to clone a
DOCX OLE frame fails closed.

## Fail-closed boundary

Do not use this operation to:

- create a new OLE object or embedding relationship;
- change the OLE program ID, icon, shell XML, frame, or preview;
- replace a shared, external, ambiguous, or unsupported package;
- treat `getEmbeddedOfficePackage()` as a general binary/OLE escape hatch;
- clone a DOCX OLE frame, or silently reinterpret it as the XLSX clone profile;
- mutate arbitrary preserved native parts; or
- bypass source-hash and relationship checks with package patching.

Malformed XLSX input fails with `invalid_presentation_ole_workbook`; malformed
DOCX input fails with `invalid_presentation_ole_office_package`, or a more
specific OPC budget/validation diagnostic. Changed bindings fail with
`presentation_ole_workbook_binding_mismatch` or
`presentation_ole_office_package_binding_mismatch`. Unsupported graph shapes
remain opaque and read-only. There is no lossy reconstruction or silent
fallback.
