# PDF product boundaries

These are format and trust boundaries, not a backlog disguised as unsupported code.

## No OpenChestnut PDF codec

OpenChestnut and its protobuf/WASM protocol remain limited to DOCX, XLSX, and PPTX. PDF uses an independent provider-routing layer. This project does not maintain a general-purpose C# PDF parser/writer, `OpenChestnut.Pdf`, or a second universal PDF object model.

The project-owned layer is deliberately narrow: provider registry, capability probes, raw-source provenance, transactional output, save policy, signature-policy checks, residue scans, audit records, and semantic/structural/visual/security regression fixtures.

## Reflow and text replacement

PDF is a final page-description format, not a Word-style authoring layout model. General text replacement with automatic paragraph/page reflow can be promised only when a trusted source model exists, a tightly controlled template defines the geometry, or a well-structured tagged PDF falls inside a tested profile.

For an ordinary imported PDF, use bounded native edits such as redaction plus a replacement overlay, or declare that the task creates a newly reconstructed document. Do not describe reconstruction as lossless in-place editing. Narrow replacement/layout algorithms must publish their geometric and font assumptions.

## Accessibility remediation

Automation can infer and write tags, propose reading order, and run machine-verifiable checks. It cannot recover author intent with certainty for arbitrary headings, table semantics, alternative text, link purpose, or ambiguous visual order. Low-confidence repairs require human review. A green veraPDF report is evidence for the selected machine rules, not a guarantee that every PDF/UA human checkpoint is satisfied.

## Dynamic and application-specific content

Dynamic XFA, complex Acrobat JavaScript, 3D annotations, and RichMedia require application-specific runtimes. The default workflow detects these constructs and preserves them opaquely when a non-destructive operation permits it. Flattening, invoking a specialist provider, or failing closed requires explicit user selection. No current shipped adapter executes these programs.

## Signed documents

An incremental update can retain byte-identical data covered by an earlier signature and can be validated as a later revision. It cannot make the earlier signer endorse arbitrary new edits. ByteRange, DocMDP, FieldMDP, field locks, and the exact modification class must be checked before and after. Broad edits should produce a new unsigned version and then be signed again.

## License boundary

The npm package and its JavaScript/C# implementation remain MIT licensed. PyMuPDF is an optional, separately installed GNU AGPL or commercial provider. The user has approved the AGPL route for this project, but installation, deployment, and redistribution still have to comply with the selected license. Provider absence or missing license acknowledgement is an explicit capability error, never a reason to fall back to lossy model reconstruction.
