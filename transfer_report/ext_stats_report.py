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
from .config import NOT_COMPLETED
from .ext_stats import OUTSIDE_COLS, TABLE_COLS, ExtStatsResult, total_row
from .overlap import overlap_note
from .pdf_report import H1, H2, LIGHT, NAVY, NOTE, PAGE, SUB, TITLE, _logo_flowable, _make_footer

DEFAULT_TITLE = "Extension Call Stats"
WIDTHS = [2.2 * cm, 6.3 * cm, 2.4 * cm] + [2.2 * cm] * 5 + [2.7 * cm, 2.1 * cm]
NOTES = {
    "All calls": "Every call leg the extension was on, as caller or callee (inbound + outbound{internal}), counted by "
                 "that leg's status - the same counting as the PBX's Extension Call Statistics report. "
                 "Answer % = Answered / Total Calls.",
    "Inbound": "Inbound legs (calls received from outside, including calls transferred to the extension and its voicemail "
               "box), by each leg's own type as in the PBX report. A call that rings several extensions counts once for each.",
    "Outbound": "Outbound legs (calls the extension dialled out), by each leg's own type as in the PBX report.",
    "Internal": "Internal legs between extensions, by each leg's own type.",
}


def _fmt_row(r: dict) -> list:
    return [r["Extension"], _trim(r["Name"], 38), f"{r['Total Calls']:,}"] + [f"{r[c]:,}" for c in
            ["Answered", "No Answer", "Busy", "Voicemail", "Abandoned", NOT_COMPLETED]] + [f"{r['Answer %']:.1f}"]


def _table(df: pd.DataFrame) -> Table:
    rows = [_fmt_row(r) for r in df.to_dict("records")]
    t = total_row(df)
    rows.append(["TOTAL", "", f"{t['Total Calls']:,}"] + [f"{t[c]:,}" for c in
                ["Answered", "No Answer", "Busy", "Voicemail", "Abandoned", NOT_COMPLETED]] + [f"{t['Answer %']:.1f}"])
    tbl = Table([[c.replace(NOT_COMPLETED, "Transfers\nNot\nCompleted") for c in TABLE_COLS]] + rows, colWidths=WIDTHS, repeatRows=1)
    tbl.setStyle(_style())
    return tbl


def _outside_table(df: pd.DataFrame) -> Table:
    """Calls outside office hours by group; the last row of df is already the total."""
    rows = [[r["Period"]] + [f"{r[c]:,}" for c in OUTSIDE_COLS[1:-1]] + [f"{r['Answer %']:.1f}"]
            for r in df.to_dict("records")]
    head = [c.replace(NOT_COMPLETED, "Transfers\nNot\nCompleted") for c in OUTSIDE_COLS]
    tbl = Table([head] + rows, colWidths=[WIDTHS[0] + WIDTHS[1]] + WIDTHS[2:], repeatRows=1)
    tbl.setStyle(_style())
    return tbl


def _style() -> TableStyle:
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (2, 0), (-1, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, LIGHT]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BBBBBB")),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D6EAE6")),
        ("LINEABOVE", (0, -1), (-1, -1), 0.8, NAVY),
    ])


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
              Paragraph(f"<b>Generated:</b> {generated.strftime('%d %B %Y, %H:%M')}", SUB)]
    if res.office_hours:
        story.append(Paragraph("<b>Counts:</b> office hours only; calls after hours, on weekends and on excluded "
                               "days are shown separately", SUB))
    story.append(PageBreak())

    scope = " Office hours only (see the last section for calls outside office hours)." if res.office_hours else ""
    for i, (name, df) in enumerate(res.tables.items()):
        heading = f"Extension Call Stats - {escape(res.department)}" if i == 0 else f"{name} calls"
        story += [CondPageBreak(7 * cm), Paragraph(heading, H1 if i == 0 else H2), _table(df),
                  Paragraph(NOTES[name].format(internal=" + internal" if res.include_internal else "") + scope, NOTE),
                  Spacer(1, 0.8 * cm)]

    if res.office_hours:
        story += [PageBreak(), Paragraph("Calls outside office hours", H1),
                  Paragraph(escape(res.hours_note), NOTE), Spacer(1, 0.3 * cm), _outside_table(res.outside),
                  Paragraph("These call legs are not included in the tables above. A call is counted once: on an "
                            "excluded day first, then on a weekend, then after hours.", NOTE),
                  Spacer(1, 0.6 * cm), CondPageBreak(7 * cm), Paragraph("Outside office hours by extension", H2),
                  _table(res.outside_by_extension),
                  Paragraph("All call legs outside office hours (after hours, weekends and excluded days together).",
                            NOTE)]

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
    if res.office_hours:
        extra = [("Counts", "Office hours only"), ("Office hours", res.hours_note)]
        extra += [(f"Outside office hours: {r['Period']}".replace(": Total outside office hours", " (total)"),
                   f"{r['Total Calls']:,}") for r in res.outside.to_dict("records")]
        summary = pd.concat([summary, pd.DataFrame(extra, columns=["Measure", "Value"])], ignore_index=True)
    tables = dict(res.tables)
    if res.office_hours:
        tables["Outside hours by extension"] = res.outside_by_extension
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        if res.office_hours:
            res.outside.to_excel(xw, sheet_name="Outside office hours", index=False)
        for name, df in tables.items():
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
