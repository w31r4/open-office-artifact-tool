# Read and review an existing PDF

Read-only review never needs model re-export.

## Route by evidence need

- `pdfinfo`: page count, page size, encryption, metadata summary.
- `pdftoppm`: final visual truth for every page.
- pdfplumber: text, words with geometry, tables, lines, and rectangles.
- pypdf: metadata, fields, annotations, outlines, attachments, encryption, and object-level quick checks.
- PyMuPDF: advanced page/object/font/image/annotation/widget inspection.
- qpdf: xref/object-stream structure, warnings, JSON/QDF inspection, and repair diagnosis.
- pyHanko: signature and trust validation.
- veraPDF: PDF/A or PDF/UA machine-verifiable rules.

## Extraction

```bash
python3 scripts/pdfplumber_extract.py input.pdf \
  --output tmp/pdfs/extraction.json \
  --max-pages 200
```

Extraction is not layout fidelity. Compare extracted text/table candidates against rendered pages, especially multi-column layouts, rotated text, merged cells, OCR layers, and scanned pages.

`PdfFile.importPdf(..., { parser: createPdfjsParser(), preferParser: true })` is useful for agent-facing inspect/QA of an arbitrary PDF, but the result is a reconstructed model. Never export that model as an edit to the original file.

## Attachment quarantine

Never write an embedded filename directly to disk. A FileSpec may contain `../`, absolute paths, platform separators, control characters, reserved device names, or duplicate names. Inventory and extract through the shipped read-only primitive:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pypdf_edit.py inspect input.pdf \
  --output tmp/pdfs/pypdf-inspect.json
"$PYTHON_BIN" scripts/pdf_provider.py check --provider pypdf --require
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task extract-attachments --provider pypdf --strategy read-only \
  --input input.pdf --require-provider
"$PYTHON_BIN" scripts/pypdf_edit.py extract-attachments input.pdf outputs/quarantine \
  --manifest outputs/attachments.json \
  --max-attachments 1000 \
  --max-total-bytes 1073741824 \
  --max-attachment-bytes 536870912
```

The manifest records provider/version, immutable source hash, scope (`document` or `page`), page/annotation identity, display name, internal key, MIME and its evidence source, decoded byte size, SHA-256, sanitized saved name/path, and transaction validation. Duplicate or colliding names receive deterministic suffixes and remain separate. A malformed FileSpec, unreadable stream, exceeded budget, pre-existing destination, hash mismatch, or source change fails closed and removes partial quarantine output. The primitive never opens, executes, imports, or recursively extracts an attachment.

Create the canonical operation audit with `savePolicy.strategy: "read-only"`, `operation.type: "extract-attachments"`, and `output` bound to the exact `attachments.json` bytes. Validate it with `pdf_audit.py validate --source input.pdf --artifact outputs/attachments.json --require-operation extract-attachments`.

## Review output

Report confirmed facts separately from inferences:

- file structure and metadata;
- visible page content and layout;
- extracted text/table candidates;
- forms, annotations, attachments, and signatures;
- accessibility/conformance evidence;
- warnings, unsupported structures, and missing providers.
