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

## Review output

Report confirmed facts separately from inferences:

- file structure and metadata;
- visible page content and layout;
- extracted text/table candidates;
- forms, annotations, attachments, and signatures;
- accessibility/conformance evidence;
- warnings, unsupported structures, and missing providers.
