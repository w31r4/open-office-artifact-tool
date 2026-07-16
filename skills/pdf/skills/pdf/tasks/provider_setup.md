# Provider setup and probes

Use the Codex workspace dependency loader first. Its bundled Python currently supplies ReportLab, pdfplumber, and pypdf; Poppler supplies `pdftoppm` and `pdfinfo`.

## PyMuPDF

The project has explicitly accepted the GNU AGPL provider path. Install the tested version into the selected Python environment:

```bash
uv pip install --python "$PYTHON" PyMuPDF==1.27.2.3
```

If `uv` is unavailable:

```bash
"$PYTHON" -m pip install PyMuPDF==1.27.2.3
```

Record the environment and license choice:

```bash
export OPEN_OFFICE_PDF_PYMUPDF_LICENSE=AGPL
"$PYTHON" scripts/pdf_provider.py check --provider pymupdf --require
```

## Other providers

- qpdf: install the official CLI appropriate for the platform, then run `qpdf --version`.
- pikepdf: optional qpdf-based Python provider. The registry probes it, but this release does not ship a pikepdf mutation adapter.
- pyHanko: install `pyHanko`/CLI extras appropriate to the signing workflow, then run `pyhanko --version`.
- veraPDF: install the official CLI distribution, then run `verapdf --version`.
- OCR for sanitize: configure Tesseract so PyMuPDF `get_textpage_ocr()` can process image-bearing pages.
- OCRmyPDF: planned scanned-document provider. Availability can be probed, but this release ships no OCRmyPDF adapter.

## Full probe

```bash
python3 scripts/pdf_provider.py check --provider all
```

Missing optional providers are reported individually. An operation requiring a missing provider must fail; never route it to a different implementation automatically.

The probe's `integration` field distinguishes shipped adapters from documented external CLIs and planned providers. `available: true` means the dependency was found; it does not upgrade a planned provider into a shipped integration.
