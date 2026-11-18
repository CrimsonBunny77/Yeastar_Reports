"""PDF (ReportLab) and Excel (.xlsx) output for the department extension call stats."""
from __future__ import annotations

import io
from datetime import datetime
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import CondPageBreak, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from .calls_report import _trim, safe_filename
from .ext_stats import TABLE_COLS, ExtStatsResult, total_row
from .overlap import overlap_note
from .pdf_report import H1, H2, LIGHT, NAVY, NOTE, PAGE, SUB, TITLE, _logo_flowable, _make_footer

DEFAULT_TITLE = "Extension Call Stats"
WIDTHS = [2.2 * cm, 6.8 * cm, 2.4 * cm] + [2.2 * cm] * 6 + [2.1 * cm]
NOTES = {
    "All calls": "Every call leg the extension was on, as caller or callee (inbound + outbound{internal}), counted by "
                 "that leg's status - the same counting as the PBX's Extension Call Statistics report. "
                 "Answer % = Answered / Total Calls.",
    "Inbound": "Legs on inbound calls (calls from outside, including when transferred). A call that rings several "
               "extensions counts once for each.",
    "Outbound": "Legs on outbound calls (calls dialled out, including transfers made during them).",
    "Internal": "Legs on internal calls between extensions.",
}


def _fmt_row(r: dict) -> list:
    return [r["Extension"], _trim(r["Name"], 38), f"{r['Total Calls']:,}"] + [f"{r[c]:,}" for c in
            ["Answered", "No Answer", "Busy", "Voicemail", "Abandoned", "Other"]] + [f"{r['Answer %']:.1f}"]


def _table(df: pd.DataFrame) -> Table:
    rows = [_fmt_row(r) for r in df.to_dict("records")]
    t = total_row(df)
    rows.append(["TOTAL", "", f"{t['Total Calls']:,}"] + [f"{t[c]:,}" for c in
                ["Answered", "No Answer", "Busy", "Voicemail", "Abandoned", "Other"]] + [f"{t['Answer %']:.1f}"])
    tbl = Table([TABLE_COLS] + rows, colWidths=WIDTHS, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (2, 0), (-1, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, LIGHT]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BBBBBB")),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D6EAE6")),
        ("LINEABOVE", (0, -1), (-1, -1), 0.8, NAVY),
    ]))
    return tbl


def build_ext_stats_pdf(res: ExtStatsResult, title: str = DEFAULT_TITLE, logo: bytes | None = None,
                        generated: datetime | None = None) -> bytes:
    generated = generated or datetime.now()
    title = title or DEFAULT_TITLE
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=PAGE, leftMargin=1.5 * cm, rightMargin=1.5 * cm, topMargin=1.5 * cm,
                            bottomMargin=1.8 * cm, title=title, author="Yeastar Transfer Report App")
    story: list = []
    story += [Spacer(1, 0.5 * cm), _logo_flowable(logo), Spacer(1, 2 * cm)] if logo else [Spacer(1, 4.5 * cm)]
    story += [Paragraph(escape(title), TITLE), Spacer(1, 1 * cm),
              Paragraph(f"<b>Department:</b> {escape(res.department)}", SUB),
              Paragraph(f"<b>Month:</b> {res.period.strftime('%B %Y')}", SUB),
              Paragraph(f"<b>Generated:</b> {generated.strftime('%d %B %Y, %H:%M')}", SUB), PageBreak()]

    for i, (name, df) in enumerate(res.tables.items()):
        heading = f"Extension Call Stats - {escape(res.department)}" if i == 0 else f"{name} calls"
        story += [CondPageBreak(7 * cm), Paragraph(heading, H1 if i == 0 else H2), _table(df),
                  Paragraph(NOTES[name].format(internal=" + internal" if res.include_internal else ""), NOTE),
                  Spacer(1, 0.8 * cm)]

    a = res.audit
    story += [PageBreak(), Paragraph("Audit", H1)]
    rows = [["Files processed", f"{a.files_processed:,}"], ["Total rows imported", f"{a.total_rows:,}"],
            ["Call records in files", f"{a.total_calls:,}"], ["Extensions in report", f"{len(res.extensions):,}"]]
    if a.first_call is not None:
        rows.append(["Data covers", f"{a.first_call:%d/%m/%Y %H:%M} - {a.last_call:%d/%m/%Y %H:%M}"])
    if getattr(a, "overlap_calls", 0):
        rows.insert(3, ["Calls in more than one file (counted once)" if a.duplicates_removed
                        else "Calls in more than one file (NOT removed)", f"{a.overlap_calls:,}"])

    if a.outside_period:
        rows.append(["Call legs outside the report month (left out)", f"{a.outside_period:,}"])
    if a.duplicates_removed:
        rows.append(["Duplicate records removed", f"{a.duplicates_removed:,}"])
    audit = Table([["Measure", "Value"]] + rows, colWidths=[11 * cm, 7 * cm])
    audit.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                               ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 9),
                               ("ALIGN", (1, 0), (1, -1), "RIGHT"), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
                               ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BBBBBB"))]))
    story.append(audit)
    if a.overlaps:
        story += [Spacer(1, 0.3 * cm), Paragraph("Overlapping files: " + escape(overlap_note(a.overlaps)), NOTE)]
    if a.files_failed:
        story += [Spacer(1, 0.5 * cm), Paragraph("Files skipped: " + "; ".join(n for n, _ in a.files_failed), NOTE)]
    footer = _make_footer(f"{title} - {res.department}")
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


def build_ext_stats_xlsx(res: ExtStatsResult, title: str = DEFAULT_TITLE, generated: datetime | None = None) -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    generated = generated or datetime.now()
    t = total_row(res.all_calls)
    summary = pd.DataFrame([
        ("Report", title or DEFAULT_TITLE), ("Department", res.department), ("Month", res.period.strftime("%B %Y")),
        ("Generated", generated.strftime("%d/%m/%Y %H:%M")), ("Extensions", f"{len(res.extensions):,}"),
        ("Total calls", f"{t['Total Calls']:,}"), ("Answered", f"{t['Answered']:,}"), ("Answer %", f"{t['Answer %']:.1f}"),
        ("Includes internal calls", "Yes" if res.include_internal else "No"),
    ], columns=["Measure", "Value"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        for name, df in res.tables.items():
            tr = total_row(df)
            body = pd.concat([df, pd.DataFrame([{"Extension": "TOTAL", "Name": "", **tr}])], ignore_index=True)
            body.to_excel(xw, sheet_name=name, index=False)
        fill = PatternFill("solid", fgColor="1F3A5F")
        for ws in xw.book.worksheets:
            for cell in ws[1]:
                cell.font, cell.fill, cell.alignment = Font(bold=True, color="FFFFFF"), fill, Alignment(vertical="center")
            ws.freeze_panes = "A2"
            if ws.title != "Summary" and ws.max_row > 2:
                for cell in ws[ws.max_row]:  # total row: bold, shaded, outside the filter range
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill("solid", fgColor="D6EAE6")
                ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row - 1}"
            for col in ws.columns:
                width = max((len(str(c.value)) for c in col[:300] if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 11), 45)
    return buf.getvalue()


def ext_filename(title: str, department: str, period: pd.Period, ext: str) -> str:
    return safe_filename(title or DEFAULT_TITLE, department, period, ext)
