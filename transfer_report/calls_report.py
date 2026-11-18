"""PDF (ReportLab) and Excel (.xlsx) output for the department call list."""
from __future__ import annotations

import io
import re
from datetime import datetime
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import CondPageBreak, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from .calls import CallsResult
from .overlap import overlap_note
from .pdf_report import H1, H2, LIGHT, NAVY, NOTE, PAGE, SUB, TITLE, _logo_flowable, _make_footer

DEFAULT_TITLE = "Department Call List"
PDF_ROW_CAP = 10_000  # per table; the Excel file always has every row
DT_FMT = "%d/%m/%Y %H:%M:%S"


def safe_filename(title: str, department: str, period: pd.Period, ext: str) -> str:
    def clean(t: str) -> str:
        return re.sub(r"[^\w\- ]", "", t, flags=re.UNICODE).strip().replace(" ", "_")
    parts = [clean(title) or "Department_Call_List", clean(department), period.strftime("%Y_%m")]
    return "_".join(p for p in parts if p) + f".{ext}"


def _trim(value, n: int) -> str:
    text = "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value)
    return text if len(text) <= n else text[: n - 1] + "…"


def _table(header: list[str], rows: list[list], widths: list[float], right: tuple[int, ...] = ()) -> Table:
    t = Table([header] + rows, colWidths=widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#BBBBBB")),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    style += [("ALIGN", (c, 0), (c, -1), "RIGHT") for c in right]
    t.setStyle(TableStyle(style))
    return t


def _call_rows(df: pd.DataFrame, with_dept: bool = False) -> list[list]:
    rows = []
    for _, r in df.iterrows():
        row = [r["Date/Time"].strftime(DT_FMT) if pd.notna(r["Date/Time"]) else ""]
        if with_dept:
            row.append(_trim(r["Department"], 22))
        row += [r["Extension"], _trim(r["Agent"], 24 if with_dept else 30), r["Number"], r["Status"],
                r["Total Duration"], r["Talk Time"], r["Call ID"]]
        rows.append(row)
    return rows


def _number_rows(df: pd.DataFrame, with_dept: bool = False) -> list[list]:
    rows = []
    for _, r in df.iterrows():
        row = [r["Number"], int(r["Calls"]), int(r["Answered"]),
               r["Last Call"].strftime(DT_FMT) if pd.notna(r["Last Call"]) else "", r["Total Talk Time"]]
        if with_dept:
            row.append(_trim(r["Departments"], 60))
        rows.append(row)
    return rows


def _section(story: list, heading: str, table_df: pd.DataFrame, builder, header, widths, right, note: str) -> None:
    story += [CondPageBreak(5 * cm), Paragraph(heading, H2)]
    if table_df.empty:
        story += [Paragraph("No records.", NOTE), Spacer(1, 0.5 * cm)]
        return
    shown = table_df.head(PDF_ROW_CAP)
    story.append(_table(header, builder(shown), widths, right))
    extra = len(table_df) - len(shown)
    story.append(Paragraph(note + (f" Showing the first {len(shown):,} of {len(table_df):,} rows; "
                                   "the Excel file contains every row." if extra > 0 else ""), NOTE))
    story.append(Spacer(1, 0.6 * cm))


def build_calls_pdf(res: CallsResult, title: str = DEFAULT_TITLE, logo: bytes | None = None,
                    generated: datetime | None = None) -> bytes:
    generated = generated or datetime.now()
    title = title or DEFAULT_TITLE
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=PAGE, leftMargin=1.5 * cm, rightMargin=1.5 * cm, topMargin=1.5 * cm,
                            bottomMargin=1.8 * cm, title=title, author="Yeastar Transfer Report App")
    story: list = []

    # Cover
    story += [Spacer(1, 0.5 * cm), _logo_flowable(logo), Spacer(1, 2 * cm)] if logo else [Spacer(1, 4.5 * cm)]
    story += [Paragraph(escape(title), TITLE), Spacer(1, 1 * cm),
              Paragraph(f"<b>Department:</b> {escape(res.department)}", SUB),
              Paragraph(f"<b>Month:</b> {res.period.strftime('%B %Y')}", SUB),
              Paragraph(f"<b>Generated:</b> {generated.strftime('%d %B %Y, %H:%M')}", SUB), PageBreak()]

    # Summary
    story.append(Paragraph(f"Summary - {escape(res.department)}", H1))
    stats = [[k, str(v) if isinstance(v, str) else f"{v:,}"] for k, v in res.stats.items()]
    story += [_table(["Measure", "Value"], stats, [11 * cm, 5 * cm], right=(1,)), Spacer(1, 0.5 * cm),
              Paragraph(res.scope_text if res.all_departments else "Extensions: " + ", ".join(res.extensions), NOTE),
              Spacer(1, 0.3 * cm),
              Paragraph("Outbound = calls made from the department's extensions (number dialled). Inbound = calls from "
                        "outside numbers that were answered by one of the department's extensions, including calls "
                        "transferred in (number = caller). Internal calls are excluded.", NOTE),
              ]
    if res.all_departments and not res.by_department.empty:
        bd = res.by_department
        story += [CondPageBreak(7 * cm), Paragraph("By department", H2),
                  _table([c.replace(" ", "\n", 1) if c != "Department" else c for c in bd.columns], [[_trim(v, 32) if i == 0 else (f"{v:,}" if isinstance(v, int) else str(v))
                                             for i, v in enumerate(r)] for r in bd.itertuples(index=False)],
                         [6 * cm] + [2.9 * cm] * 7, right=tuple(range(1, 8))),
                  Paragraph("An inbound call answered by more than one department is counted once for each of them "
                            "here, and once in the totals above.", NOTE)]
    story.append(PageBreak())

    # Unique numbers, then the full call logs
    d = res.all_departments
    if d:
        num_w, num_right = [3.8 * cm, 2.2 * cm, 2.2 * cm, 4.2 * cm, 3.2 * cm, 11 * cm], (1, 2)
        hdr = ["Date/Time", "Department", "Extension", "Agent", "Number", "Status", "Total", "Talk Time", "Call ID"]
        w = [3.6 * cm, 3.7 * cm, 1.8 * cm, 4.4 * cm, 3.1 * cm, 2.4 * cm, 1.9 * cm, 1.9 * cm, 3.9 * cm]
        right = (6, 7)
    else:
        num_w, num_right = [4.5 * cm, 3 * cm, 3 * cm, 5 * cm, 4 * cm], (1, 2)
        hdr = ["Date/Time", "Extension", "Agent", "Number", "Status", "Total", "Talk Time", "Call ID"]
        w = [3.8 * cm, 2 * cm, 5.3 * cm, 3.4 * cm, 2.6 * cm, 2.1 * cm, 2.1 * cm, 4.2 * cm]
        right = (5, 6)
    _section(story, "Outbound - numbers dialled", res.outbound_numbers, lambda x: _number_rows(x, d), res.number_cols,
             num_w, num_right, "One row per number, most-called first.")
    _section(story, "Inbound - numbers connected", res.inbound_numbers, lambda x: _number_rows(x, d), res.number_cols,
             num_w, num_right, "One row per caller number, most-calls first.")
    story.append(PageBreak())
    _section(story, "Outbound call log", res.outbound, lambda x: _call_rows(x, d), hdr, w, right, "Oldest first.")
    story.append(PageBreak())
    _section(story, "Inbound call log (answered by the department)" if not d else "Inbound call log (answered)",
             res.inbound, lambda x: _call_rows(x, d), hdr, w, right, "Oldest first.")

    # Audit
    a = res.audit
    story += [PageBreak(), Paragraph("Audit", H1)]
    rows = [["Files processed", f"{a.files_processed:,}"], ["Total rows imported", f"{a.total_rows:,}"],
            ["Call records in files", f"{a.total_calls:,}"]]
    if a.first_call is not None:
        rows.append(["Data covers", f"{a.first_call:%d/%m/%Y %H:%M} - {a.last_call:%d/%m/%Y %H:%M}"])
    if getattr(a, "overlap_calls", 0):
        rows.insert(3, ["Calls in more than one file (counted once)" if a.duplicates_removed
                        else "Calls in more than one file (NOT removed)", f"{a.overlap_calls:,}"])

    if a.duplicates_removed:
        rows.append(["Duplicate records removed", f"{a.duplicates_removed:,}"])
    story.append(_table(["Measure", "Value"], rows, [11 * cm, 7 * cm], right=(1,)))
    if a.overlaps:
        story.append(Paragraph("Overlapping files: " + escape(overlap_note(a.overlaps)), NOTE))
    if a.per_file:
        story += [Spacer(1, 0.5 * cm),
                  _table(["File", "Rows", "Calls", "Outbound", "Inbound"],
                         [[_trim(f["File"], 60), f"{f['Rows']:,}", f"{f['Calls']:,}", f"{f['Outbound']:,}",
                           f"{f['Inbound']:,}"] for f in a.per_file],
                         [12 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm], right=(1, 2, 3, 4))]
    if a.files_failed:
        story += [Spacer(1, 0.5 * cm), Paragraph("Files skipped: " + "; ".join(n for n, _ in a.files_failed), NOTE)]

    footer = _make_footer(f"{title} - {res.department}")
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


def build_calls_xlsx(res: CallsResult, title: str = DEFAULT_TITLE, generated: datetime | None = None) -> bytes:
    """Excel workbook: Summary, Outbound calls, Inbound calls, Outbound numbers, Inbound numbers."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    generated = generated or datetime.now()
    summary = [("Report", title or DEFAULT_TITLE), ("Department", res.department),
               ("Month", res.period.strftime("%B %Y")), ("Generated", generated.strftime("%d/%m/%Y %H:%M")),
               ("Scope" if res.all_departments else "Extensions",
                res.scope_text if res.all_departments else ", ".join(res.extensions))] + [(k, v) for k, v in res.stats.items()]
    summary_df = pd.DataFrame(summary, columns=["Measure", "Value"]).astype({"Value": str})

    sheets = {
        "Summary": summary_df,
        "Outbound calls": res.outbound[res.display_cols],
        "Inbound calls": res.inbound[res.display_cols],
        "Outbound numbers": res.outbound_numbers,
        "Inbound numbers": res.inbound_numbers,
    }
    if res.all_departments:
        sheets = {"Summary": sheets["Summary"], "By department": res.by_department, **{
            k: v for k, v in sheets.items() if k != "Summary"}}
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
        header_fill = PatternFill("solid", fgColor="1F3A5F")
        for ws in xw.book.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
                cell.alignment = Alignment(vertical="center")
            for col in ws.iter_cols(min_row=2):
                if ws.cell(row=1, column=col[0].column).value in ("Date/Time", "Last Call"):
                    for cell in col:
                        cell.number_format = "dd/mm/yyyy hh:mm:ss"
            ws.freeze_panes = "A2"
            if ws.max_row > 1 and ws.max_column > 0:
                ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                width = max((len(str(c.value)) for c in col[:300] if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 12), 60)
    return buf.getvalue()
