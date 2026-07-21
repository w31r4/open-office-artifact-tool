# Restrict Editing / Make Read-Only (Document Protection)

## Goal

Ask Word-compatible hosts to limit ordinary editing to one of these modes:

- `readOnly`
- `comments`
- `trackedChanges`
- `forms`
- explicit `none`

This is useful for review copies and templates that should not be casually changed. It is an editing restriction, **not encryption, authentication, or access control**. A non-conforming application or a determined user can ignore or remove passwordless protection.

## Default public workflow

Create or import the document, set protection as document state, export, then reimport and verify:

```js
import { DocumentFile, DocumentModel, FileBlob } from "open-office-artifact-tool";

const document = DocumentModel.create({
  name: "Review copy",
  blocks: [{ kind: "paragraph", text: "Approved wording." }],
});
document.setSettings({ documentProtection: "comments" });

const output = await DocumentFile.exportDocx(document);
await output.save("comments-only.docx");

const checked = await DocumentFile.importDocx(await FileBlob.load("comments-only.docx"));
if (checked.settings.documentProtection?.edit !== "comments") {
  throw new Error("Document protection did not survive the DOCX round trip.");
}
```

String modes default to enforcement on and formatting protection off. Use the object form when both flags must be explicit:

```js
document.setSettings({
  documentProtection: {
    edit: "readOnly",
    enforcement: true,
    formatting: false,
  },
});
```

`edit: "none"` writes an explicit `w:documentProtection w:edit="none"` state. To remove the element, use one of:

```js
document.setSettings({ documentProtection: false });
document.setSettings({ documentProtection: null });
document.setSettings({ documentProtection: "off" });
```

## Imported-document safety boundary

OpenChestnut recognizes only the canonical passwordless element with `edit`, `enforcement`, and `formatting`. Password hashes/verifiers, cryptographic attributes, extension markup, IRM, and permission exceptions remain source-owned:

- importing and exporting without changing protection preserves their package part;
- attempting to replace them through `document.setSettings(...)` must **fail closed**;
- do not strip or rebuild unsupported protection as an automatic fallback;
- obtain a user-approved unprotected source or use a dedicated policy-aware workflow if the restriction must change.

## Explicit compatibility helper

The reference-compatible Python helper remains available for an explicit package-level workflow:

It requires Python with `lxml`; probe that dependency before selecting this optional route.

```bash
python scripts/set_protection.py input.docx --mode readOnly --out protected.docx
python scripts/set_protection.py input.docx --mode comments --out comments_only.docx
python scripts/set_protection.py input.docx --mode trackedChanges --out tracked_only.docx
python scripts/set_protection.py input.docx --mode forms --out forms_only.docx
python scripts/set_protection.py input.docx --mode off --out unprotected.docx
```

Choose this helper deliberately; it is never a silent fallback from the public API. It refuses an existing protection element with password/cryptographic or otherwise unsupported markup.

## Verification

1. Reimport the exported DOCX and inspect `document.settings.documentProtection`.
2. Inspect `word/settings.xml` and confirm the expected `w:edit`, `w:enforcement`, and `w:formatting` values.
3. Render both the unrestricted control and protected result through LibreOffice plus Poppler.
4. Require equal page counts and a zero-pixel diff: protection settings should not alter document layout.
5. Open in the intended host when application-specific enforcement matters; record which host/version was tested.

```bash
python render_docx.py protected.docx --output_dir out_protected
```

## Pitfalls

- A Word editing restriction does not secure confidential content. Use filesystem permissions or encrypted transport/storage for access control.
- `enforcement: false` records the mode without asking conforming hosts to enforce it.
- Some viewers ignore protection entirely.
- Password-protected and policy-managed documents are intentionally outside this bounded public model.
- Some source files do not have `word/settings.xml`; OpenChestnut creates a valid settings part when authoring a supported mode.
