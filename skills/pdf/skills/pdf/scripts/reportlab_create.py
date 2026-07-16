#!/usr/bin/env python3
"""Create a greenfield visual PDF from a bounded JSON spec using ReportLab."""

from __future__ import annotations

import argparse
import hashlib
from html import escape
import json
from pathlib import Path
import sys


def load_spec(path: Path) -> dict:
    spec = json.loads(path.read_text("utf-8"))
    if not isinstance(spec, dict):
        raise ValueError("spec must be a JSON object")
    return spec


def build_pdf(spec: dict, output: Path) -> None:
    try:
        from reportlab.graphics.charts.barcharts import VerticalBarChart
        from reportlab.graphics.shapes import Drawing
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER
        from reportlab.lib.pagesizes import A4, LETTER
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import inch
        from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except ImportError as exc:
        raise RuntimeError("ReportLab is required; install it in the selected Python environment") from exc

    page_size = A4 if str(spec.get("pageSize", "LETTER")).upper() == "A4" else LETTER
    output.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(output),
        pagesize=page_size,
        leftMargin=0.72 * inch,
        rightMargin=0.72 * inch,
        topMargin=0.68 * inch,
        bottomMargin=0.62 * inch,
        title=str(spec.get("title") or "Report"),
        author=str(spec.get("author") or "open-office-artifact-tool PDF Skill"),
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="ReportTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=24, leading=28, textColor=colors.HexColor("#123B5D"), spaceAfter=7))
    styles.add(ParagraphStyle(name="ReportSubtitle", parent=styles["Normal"], fontName="Helvetica", fontSize=11, leading=15, textColor=colors.HexColor("#486779"), spaceAfter=16))
    styles.add(ParagraphStyle(name="Section", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=14, leading=18, textColor=colors.HexColor("#123B5D"), spaceBefore=9, spaceAfter=7))
    styles.add(ParagraphStyle(name="BodyCopy", parent=styles["BodyText"], fontName="Helvetica", fontSize=10.5, leading=15, textColor=colors.HexColor("#172033"), spaceAfter=9))
    styles.add(ParagraphStyle(name="Footer", parent=styles["Normal"], fontName="Helvetica", fontSize=8, textColor=colors.HexColor("#647784"), alignment=TA_CENTER))

    story = [
        Paragraph(escape(str(spec.get("title") or "Report")), styles["ReportTitle"]),
        Paragraph(escape(str(spec.get("subtitle") or "Evidence-led summary")), styles["ReportSubtitle"]),
    ]
    for section in list(spec.get("sections") or [])[:40]:
        if not isinstance(section, dict):
            continue
        story.append(Paragraph(escape(str(section.get("heading") or "Section")), styles["Section"]))
        for paragraph in list(section.get("paragraphs") or [])[:100]:
            story.append(Paragraph(escape(str(paragraph)), styles["BodyCopy"]))
        table_rows = section.get("table")
        if isinstance(table_rows, list) and table_rows:
            safe_rows = [[str(cell) for cell in list(row)[:12]] for row in table_rows[:200] if isinstance(row, list)]
            if safe_rows:
                width = (page_size[0] - doc.leftMargin - doc.rightMargin) / max(1, max(len(row) for row in safe_rows))
                table = Table(safe_rows, colWidths=[width] * max(len(row) for row in safe_rows), repeatRows=1)
                table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCEAF3")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#123B5D")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("LEADING", (0, 0), (-1, -1), 12),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#9EB3C6")),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 7),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]))
                story.extend([table, Spacer(1, 12)])
        chart = section.get("chart")
        if isinstance(chart, dict):
            categories = [str(value) for value in list(chart.get("categories") or [])[:20]]
            values = [float(value) for value in list(chart.get("values") or [])[:20]]
            if categories and len(categories) == len(values):
                drawing = Drawing(470, 190)
                bars = VerticalBarChart()
                bars.x = 45
                bars.y = 35
                bars.height = 125
                bars.width = 390
                bars.data = [values]
                bars.categoryAxis.categoryNames = categories
                bars.categoryAxis.labels.fontName = "Helvetica"
                bars.categoryAxis.labels.fontSize = 8
                bars.valueAxis.valueMin = min(0, min(values))
                bars.valueAxis.valueMax = max(values) * 1.15 if max(values) else 1
                bars.bars[0].fillColor = colors.HexColor(str(chart.get("color") or "#0F766E"))
                drawing.add(bars)
                story.append(KeepTogether([Paragraph(escape(str(chart.get("title") or "Chart")), styles["Section"]), drawing]))

    footer_text = str(spec.get("footer") or "Generated with the PDF Skill")

    def page_footer(canvas, document):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#647784"))
        canvas.drawCentredString(page_size[0] / 2, 0.34 * inch, f"{footer_text}  |  {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=page_footer, onLaterPages=page_footer)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--spec", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--max-spec-bytes", type=int, default=10 * 1024 * 1024)
    args = parser.parse_args()
    try:
        spec_path = args.spec.expanduser().resolve()
        output = args.output.expanduser().resolve()
        if not spec_path.is_file():
            raise ValueError("--spec must be an existing JSON file")
        if args.max_spec_bytes < 1:
            raise ValueError("--max-spec-bytes must be positive")
        if spec_path.stat().st_size > args.max_spec_bytes:
            raise ValueError(f"spec is {spec_path.stat().st_size} bytes; max-spec-bytes is {args.max_spec_bytes}")
        if spec_path == output:
            raise ValueError("spec and output paths must differ")
        build_pdf(load_spec(spec_path), output)
        payload = output.read_bytes()
        print(json.dumps({"output": str(output), "bytes": len(payload), "sha256": hashlib.sha256(payload).hexdigest(), "provider": "reportlab", "strategy": "rewrite"}, sort_keys=True))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc), "provider": "reportlab", "silentFallback": False}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
