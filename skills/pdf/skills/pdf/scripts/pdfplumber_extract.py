#!/usr/bin/env python3
"""Extract bounded text, word geometry, and table candidates with pdfplumber."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import sys


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--max-pages", type=int, default=200)
    parser.add_argument("--max-words", type=int, default=50_000)
    parser.add_argument("--max-tables", type=int, default=1_000)
    parser.add_argument("--max-chars", type=int, default=2_000_000)
    parser.add_argument("--max-bytes", type=int, default=512 * 1024 * 1024)
    args = parser.parse_args()
    try:
        import pdfplumber
    except ImportError:
        print(json.dumps({"ok": False, "error": "pdfplumber is not installed", "provider": "pdfplumber", "silentFallback": False}), file=sys.stderr)
        return 2
    try:
        source = args.input.expanduser().resolve()
        if not source.is_file():
            raise ValueError("input must be an existing PDF")
        if args.max_pages < 1 or args.max_words < 1 or args.max_tables < 1 or args.max_chars < 1 or args.max_bytes < 1:
            raise ValueError("all extraction limits must be positive")
        if source.stat().st_size > args.max_bytes:
            raise ValueError(f"PDF is {source.stat().st_size} bytes; max-bytes is {args.max_bytes}")
        pages = []
        total_words = total_tables = total_chars = 0
        with pdfplumber.open(str(source)) as pdf:
            if len(pdf.pages) > args.max_pages:
                raise ValueError(f"PDF has {len(pdf.pages)} pages; max-pages is {args.max_pages}")
            for page_number, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ""
                total_chars += len(text)
                if total_chars > args.max_chars:
                    raise ValueError(f"extracted text exceeds max-chars {args.max_chars}")
                words = page.extract_words() or []
                total_words += len(words)
                if total_words > args.max_words:
                    raise ValueError(f"extracted words exceed max-words {args.max_words}")
                tables = page.extract_tables() or []
                total_tables += len(tables)
                if total_tables > args.max_tables:
                    raise ValueError(f"table candidates exceed max-tables {args.max_tables}")
                pages.append({
                    "page": page_number,
                    "width": page.width,
                    "height": page.height,
                    "text": text,
                    "words": [{key: word.get(key) for key in ("text", "x0", "x1", "top", "bottom", "doctop", "upright", "direction")} for word in words],
                    "tables": tables,
                    "lines": len(page.lines or []),
                    "rects": len(page.rects or []),
                    "images": len(page.images or []),
                })
        payload = {
            "provider": "pdfplumber",
            "strategy": "read-only",
            "source": {"path": str(source), "bytes": source.stat().st_size, "sha256": sha256(source)},
            "summary": {"pages": len(pages), "chars": total_chars, "words": total_words, "tableCandidates": total_tables},
            "pages": pages,
            "warning": "table extraction is heuristic and must be checked against rendered page geometry",
        }
        rendered = json.dumps(payload, ensure_ascii=False, indent=2)
        if args.output:
            output = args.output.expanduser().resolve()
            if output == source:
                raise ValueError("JSON report path must differ from the source PDF")
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_text(rendered + "\n", "utf-8")
        else:
            print(rendered)
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc), "provider": "pdfplumber", "silentFallback": False}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
