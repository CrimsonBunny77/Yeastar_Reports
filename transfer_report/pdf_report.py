"""PDF generation with ReportLab (landscape A4)."""
from __future__ import annotations

import io
import logging
import re
from datetime import datetime

import pandas as pd
from reportlab.graphics.charts.barcharts import HorizontalBarChart
from reportlab.graphics.shapes import Drawing
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (CondPageBreak, Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

from .charts import ChartSpec
from .config import NOT_COMPLETED
from .errors import ReportError
from .overlap import overlap_note
from .processor import Audit

log = logging.getLogger(__name__)
NAVY = colors.HexColor("#1F3A5F")
LIGHT = colors.HexColor("#EEF2F7")
PAGE = landscape(A4)
DEFAULT_TITLE = "Yeastar Transfer Report"
MAX_W, MAX_H = 25.0 * cm, 14.5 * cm

_base = getSampleStyleSheet()
TITLE = ParagraphStyle("T", parent=_base["Title"], fontSize=34, textColor=NAVY, leading=40)
SUB = ParagraphStyle("S", parent=_base["Normal"], fontSize=16, textColor=colors.HexColor("#444444"), alignment=1, leading=22)
H1 = ParagraphStyle("H1", parent=_base["Heading1"], fontSize=18, textColor=NAVY, spaceAfter=10)
H2 = ParagraphStyle("H2", parent=_base["Heading2"], fontSize=14, textColor=NAVY, spaceAfter=8)
NOTE = ParagraphStyle("N", parent=_base["Normal"], fontSize=9, textColor=colors.grey)


def _fmt(v) -> str:
    return f"{v:,.1f}" if isinstance(v, float) else f"{v:,}" if isinstance(v, int) else str(v)


def _table(header: list[str], rows: list[list], widths: list[float], total: list | None = None,
           num_from: int = 2, sub_rows: list[int] | None = None) -> Table:
    header = [h.replace(NOT_COMPLETED, "Transfers\nNot\nCompleted") for h in header]  # long heading: wrap in the cell
    body = [header] + [[_fmt(c) for c in r] for r in rows]
    if total:
        body.append([_fmt(c) for c in total])
    t = Table(body, colWidths=widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (num_from, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BBBBBB")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for i in sub_rows or []:  # subtotal rows (index in `rows`; +1 for the header)
        style += [("FONTNAME", (0, i + 1), (-1, i + 1), "Helvetica-Bold"),
                  ("BACKGROUND", (0, i + 1), (-1, i + 1), colors.HexColor("#D6EAE6")),
                  ("LINEABOVE", (0, i + 1), (-1, i + 1), 0.8, NAVY)]
    if total:
        style += [("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D5DEEA"))]
    t.setStyle(TableStyle(style))
    return t


def _totals(detail: pd.DataFrame) -> dict:
    s = detail[["Transfers", "Answered", "No Answer", "Busy", "Voicemail", "Abandoned", NOT_COMPLETED]].sum()
    out = {k: int(v) for k, v in s.items()}
    out["Answer %"] = round(out["Answered"] / out["Transfers"] * 100, 1) if out["Transfers"] else 0.0
    return out


FLAG_COLS = ["Transfers", "Answered", "No Answer", "Busy", "Voicemail", "Abandoned", NOT_COMPLETED]


def _grouped_totals(detail: pd.DataFrame, group: str) -> tuple[list[list], list[int]]:
    """Rows grouped by one side of the transfer, each group closed by a 'Total' row with Answer %.

    group="To Department"   -> transfers *received* by each department, split by sender.
    group="From Department" -> transfers *sent* by each department, split by receiver.
    Returns (rows, indices of total rows)."""
    received = group == "To Department"
    other = "From Department" if received else "To Department"
    rows, totals = [], []
    for dept in sorted(detail[group].unique()):
        g = detail[detail[group] == dept].sort_values(["Transfers", other], ascending=[False, True])
        for _, r in g.iterrows():
            pair = [r[other], dept] if received else [dept, r[other]]
            rows.append(pair + [int(r[c]) for c in FLAG_COLS] + [""])
        t = g[FLAG_COLS].sum()
        pct = round(t["Answered"] / t["Transfers"] * 100, 1) if t["Transfers"] else 0.0
        totals.append(len(rows))
        rows.append((["Total", dept] if received else [dept, "Total"]) + [int(t[c]) for c in FLAG_COLS] + [pct])
    return rows, totals


GROUP_HEADER = ["From Department", "To Department", "Transfers", "Answered", "No Answer", "Busy",
                "Voicemail", "Abandoned", NOT_COMPLETED, "Answer %"]
GROUP_WIDTHS = [5.2 * cm, 5.2 * cm] + [2.1 * cm] * 6 + [2.7 * cm, 2.1 * cm]


def _native_chart(spec: ChartSpec) -> Drawing:
    """ReportLab-only fallback used when Plotly image export (kaleido) is unavailable."""
    w, h = MAX_W, min(MAX_H, 3.2 * cm + 0.7 * cm * len(spec.labels))
    d = Drawing(w, h)
    bc = HorizontalBarChart()
    bc.x, bc.y, bc.width, bc.height = 7 * cm, 0.8 * cm, w - 9 * cm, h - 1.8 * cm
    bc.data = [list(reversed(v)) for v in spec.series.values()]
    bc.categoryAxis.categoryNames = list(reversed(spec.labels))
    bc.categoryAxis.labels.fontSize = 8
    bc.valueAxis.valueMin = 0
    bc.valueAxis.labels.fontSize = 8
    bc.barLabelFormat = "%.1f%%" if spec.percent else "%d"
    bc.barLabels.fontSize = 7
    bc.barLabels.nudge = 12
    for i, c in enumerate(spec.colors[: len(spec.series)]):
        bc.bars[i].fillColor = colors.HexColor(c)
    d.add(bc)
    return d


def _chart_flowable(spec: ChartSpec):
    try:
        png = spec.figure().to_image(format="png", width=spec.width, height=spec.height, scale=2)
        w, h = MAX_W, MAX_W * spec.height / spec.width
        if h > MAX_H:
            w, h = MAX_H * spec.width / spec.height, MAX_H
        return Image(io.BytesIO(png), width=w, height=h)
    except Exception as exc:  # noqa: BLE001 - kaleido/plotly missing or failing
        log.warning("Plotly image export failed (%s); using native ReportLab chart.", exc)
        return _native_chart(spec)


LOGO_MAX_W, LOGO_MAX_H = 9 * cm, 3.5 * cm


def prepare_logo(data: bytes) -> bytes:
    """Validate an uploaded logo and normalise it to a PNG (max 1200 px) so the PDF stays small."""
    from PIL import Image as PILImage

    try:
        img = PILImage.open(io.BytesIO(data))
        img.load()
        img = img.convert("RGBA")
        img.thumbnail((1200, 1200))
        out = io.BytesIO()
        img.save(out, format="PNG")
        return out.getvalue()
    except Exception as exc:  # noqa: BLE001
        raise ReportError("The logo could not be read. Please upload a PNG or JPG image.") from exc


def _logo_flowable(png: bytes) -> Image:
    w, h = ImageReader(io.BytesIO(png)).getSize()
    scale = min(LOGO_MAX_W / w, LOGO_MAX_H / h)
    return Image(io.BytesIO(png), width=w * scale, height=h * scale, hAlign="CENTER", mask="auto")


def _make_footer(label: str):
    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)
        canvas.drawString(1.5 * cm, 1 * cm, label)
        canvas.drawRightString(PAGE[0] - 1.5 * cm, 1 * cm, f"Page {doc.page}")
        canvas.restoreState()
    return _footer


def build_pdf(detail: pd.DataFrame, audit: Audit, period: pd.Period,
              charts: list[ChartSpec], generated: datetime | None = None,
              scope: list[str] | None = None, match: str = "both",
              title: str = DEFAULT_TITLE, status_summary: pd.DataFrame | None = None,
              logo: bytes | None = None) -> bytes:
    generated = generated or datetime.now()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=PAGE, leftMargin=1.5 * cm, rightMargin=1.5 * cm,
                            topMargin=1.5 * cm, bottomMargin=1.8 * cm,
                            title=title, author="Yeastar Transfer Report App")
    tot = _totals(detail)
    story: list = []

    # Cover
    if logo:
        story += [Spacer(1, 0.5 * cm), _logo_flowable(logo), Spacer(1, 2 * cm)]
    else:
        story.append(Spacer(1, 4.5 * cm))
    story += [Paragraph(escape(title), TITLE), Spacer(1, 1 * cm),
              Paragraph(f"<b>Month:</b> {period.strftime('%B %Y')}", SUB),
              Paragraph(f"<b>Generated:</b> {generated.strftime('%d %B %Y, %H:%M')}", SUB)]
    if scope:
        rule = "both departments" if match == "both" else "either department"
        story += [Spacer(1, 0.4 * cm),
                  Paragraph(f"<b>Departments included</b> (transfers where {rule} selected):<br/>{', '.join(scope)}", SUB)]
    story.append(PageBreak())

    # Executive summary (unchanged layout)
    story.append(Paragraph("Executive Summary", H1))
    summ = detail.sort_values(["Transfers", "From Department"], ascending=[False, True])
    rows = summ[["From Department", "To Department", "Transfers", "Answered", "Answer %"]].values.tolist()
    rows = [[a, b, int(c), int(d), float(e)] for a, b, c, d, e in rows]
    story.append(_table(["From Department", "To Department", "Transfers", "Answered", "Answer %"], rows,
                        [7 * cm, 7 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm],
                        ["TOTAL", "", tot["Transfers"], tot["Answered"], tot["Answer %"]]))

    # Transfers received by each department, with outcome breakdown
    story += [CondPageBreak(6 * cm), Spacer(1, 0.8 * cm), Paragraph("Transfers Received by Department", H2)]
    rows, totals = _grouped_totals(detail, "To Department")
    story += [_table(GROUP_HEADER, rows, GROUP_WIDTHS, sub_rows=totals), PageBreak()]

    # Detailed matrix
    story.append(Paragraph("Detailed Transfer Matrix", H1))
    cols = ["From Department", "To Department", "Transfers", "Answered", "No Answer", "Busy",
            "Voicemail", "Abandoned", NOT_COMPLETED, "Answer %"]
    d = detail.sort_values(["From Department", "To Department"])
    rows = [[r["From Department"], r["To Department"]] + [int(r[c]) for c in cols[2:9]] + [float(r["Answer %"])]
            for _, r in d.iterrows()]
    story.append(_table(["From", "To"] + [GROUP_HEADER[i] for i in range(2, 10)], rows, GROUP_WIDTHS,
                        ["TOTAL", ""] + [tot[c] for c in cols[2:9]] + [tot["Answer %"]]))

    # Transfers sent by each department, with outcome breakdown
    story += [CondPageBreak(6 * cm), Spacer(1, 0.8 * cm), Paragraph("Transfers Sent by Department", H2)]
    rows, totals = _grouped_totals(detail, "From Department")
    story += [_table(GROUP_HEADER, rows, GROUP_WIDTHS, sub_rows=totals), PageBreak()]

    # Charts: one per page for readability
    for i, spec in enumerate(charts):
        block = [Paragraph(spec.title, H1)]
        if len(spec.series) > 1:
            legend = "  ".join(f'<font color="{c}">&#9632;</font> {n}' for n, c in zip(spec.series, spec.colors))
            block.append(Paragraph(legend, NOTE))
        block.append(_chart_flowable(spec))
        story.append(KeepTogether(block))
        story.append(PageBreak())

    # Audit
    story.append(Paragraph("Audit", H1))
    rows = [
        ["Files Processed", audit.files_processed],
        ["Total Rows Imported", audit.total_rows],
        ["Call Records", audit.total_calls],
        ["Blind Transfers", audit.blind],
        ["Attended Transfers", audit.attended],
        ["Final Transfer Records", audit.final_records],
        ["Transfers with no matching call leg (counted as " + NOT_COMPLETED + ")", audit.unmatched_status],
        ["Excluded: unmapped extensions", audit.unmapped_removed],
        ["Excluded: same-department transfers", audit.internal_removed],
    ]
    if getattr(audit, "first_call", None) is not None:
        rows.insert(2, ["Data covers", f"{audit.first_call:%d/%m/%Y %H:%M} - {audit.last_call:%d/%m/%Y %H:%M}"])
    if getattr(audit, "overlap_calls", 0):
        rows.insert(3, ["Calls in more than one file (counted once)" if audit.duplicates_removed
                        else "Calls in more than one file (NOT removed)", f"{audit.overlap_calls:,}"])
    if audit.deselected_removed:
        rows.append(["Excluded: departments not selected", audit.deselected_removed])
    if audit.duplicates_removed:
        rows.append(["Excluded: duplicate records", audit.duplicates_removed])
    story += [_table(["Measure", "Value"], rows, [11 * cm, 7 * cm])]
    if getattr(audit, "overlaps", None):
        story.append(Paragraph("Overlapping files: " + escape(overlap_note(audit.overlaps)), NOTE))
    story.append(Spacer(1, 0.6 * cm))
    if audit.per_file:
        story.append(_table(["File", "Rows", "Calls", "Blind", "Attended"],
                            [[f["File"], f["Rows"], f["Calls"], f["Blind"], f["Attended"]] for f in audit.per_file],
                            [11 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm], num_from=1))
    if status_summary is not None and not status_summary.empty:
        top = status_summary.head(15)
        story += [Spacer(1, 0.6 * cm), Paragraph("Last Status values (final transfer records)", H1),
                  _table(["Last Status", "Records", "Counted as"],
                         [[r["Last Status"], int(r["Records"]), r["Counted as"]] for _, r in top.iterrows()],
                         [10 * cm, 4 * cm, 6 * cm], num_from=1)]
    if audit.files_failed:
        story += [Spacer(1, 0.5 * cm), Paragraph("Files skipped: " + "; ".join(n for n, _ in audit.files_failed), NOTE)]

    footer = _make_footer(title)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


def report_filename(period: pd.Period, title: str = DEFAULT_TITLE) -> str:
    safe = re.sub(r"[^\w\- ]", "", title, flags=re.UNICODE).strip().replace(" ", "_") or "Transfer_Report"
    return f"{safe}_{period.strftime('%Y_%m')}.pdf"
