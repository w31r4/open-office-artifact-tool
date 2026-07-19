# PDF product boundaries

These are format and trust boundaries, not a backlog disguised as unsupported code.

## No OpenChestnut PDF codec

OpenChestnut and its protobuf/WASM protocol remain limited to DOCX, XLSX, and PPTX. PDF uses an independent provider-routing layer. This project does not maintain a general-purpose C# PDF parser/writer, `OpenChestnut.Pdf`, or a second universal PDF object model.

The project-owned layer exposes agent-usable primitives: greenfield authoring, runtime-lazy MuPDF.js parsing/inspection/rendering/bounded edits, raw-source provenance, transactional output, save-policy and signature fail-closed rules, plus shared semantic/structural/visual/security QA. Specialist tools remain explicit routes rather than a second universal PDF model.

## Reflow and text replacement

PDF is a final page-description format, not a Word-style authoring layout model. General text replacement with automatic paragraph/page reflow can be promised only when a trusted source model exists, a tightly controlled template defines the geometry, or a well-structured tagged PDF falls inside a tested profile.

For an ordinary imported PDF, use bounded native edits such as redaction plus a replacement overlay, or declare that the task creates a newly reconstructed document. Do not describe reconstruction as lossless in-place editing. Narrow replacement/layout algorithms must publish their geometric and font assumptions.

## Accessibility remediation

Automation can infer and write tags, propose reading order, and run machine-verifiable checks. It cannot recover author intent with certainty for arbitrary headings, table semantics, alternative text, link purpose, or ambiguous visual order. Low-confidence repairs require human review. A green veraPDF report is evidence for the selected machine rules, not a guarantee that every PDF/UA human checkpoint is satisfied.

## Dynamic and application-specific content

Dynamic XFA, complex Acrobat JavaScript, 3D annotations, and RichMedia require application-specific runtimes. The default workflow detects these constructs and preserves them opaquely when a non-destructive operation permits it. The pikepdf `structure-clean` route may explicitly defang JavaScript, external actions, multimedia, and selected auxiliary structures, but does not execute or flatten them and deliberately leaves XFA/forms/metadata outside its scope. Flattening, invoking another specialist provider, or failing closed requires explicit user selection. No shipped adapter executes these programs.

## Signed documents

An incremental update can retain byte-identical data covered by an earlier signature. It cannot make the earlier signer endorse arbitrary new edits. The MuPDF.js primitive does not validate cryptographic trust, DocMDP, FieldMDP, or field locks, so it rejects signed-PDF incremental editing. A deliberate rewrite requires explicit signature invalidation; specialist validation and re-signing happen outside this primitive.

## License boundary

The project and required `mupdf` npm dependency are GNU AGPL-3.0-or-later. Normal npm installation resolves MuPDF.js; its WASM runtime is lazy and no lifecycle installer is used. Optional Python and system providers are installed separately under their own terms. Missing capability or provider evidence is an explicit error, never a reason to fall back to lossy model reconstruction.
