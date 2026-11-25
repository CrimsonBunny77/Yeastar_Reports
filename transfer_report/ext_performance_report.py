"""Excel (.xlsx) output for the Extension Performance Detail report (Excel only, no PDF)."""
from __future__ import annotations

import calendar
import io
import re

import pandas as pd

from .errors import ReportError
from .ext_performance import DETAIL_COLS, DURATION_COLS, PerfResult

DEFAULT_TITLE = "Inbound and Outbound Call Report"
MAX_EXCEL_ROWS = 1_048_575  # Excel's row limit less the header row
HEADER_FILL = "1F3864"
TIME_FORMAT = "dd/mm/yyyy hh:mm:ss"
DURATION_FORMAT = "[h]:mm:ss"
BASE_WIDTHS = {"ID": 24, "Time": 20, "Consultants": 20, "Call From": 28, "Call To": 30, "Call Duration": 15,
               "Ring Duration": 15, "Talk Duration": 15, "Communication Type": 20, "Status": 13}
WIDE_COLS = ("Consultants", "Call From", "Call To")  # widened when a value is longer than the standard width
MAX_WIDTH = 45


def _clean(text: str) -> str:
    """Remove characters that are not allowed in file names; keep spaces and hyphens."""
    return re.sub(r"\s+", " ", re.sub(r'[\\/:*?"<>|\r\n\t]', "", text)).strip()


def period_label(period: pd.Period) -> str:
    """'1-30 September 2026'."""
    last = calendar.monthrange(period.year, period.month)[1]
    return f"1-{last} {period.strftime('%B %Y')}"


def perf_filename(title: str, department: str, period: pd.Period, ext: str = "xlsx") -> str:
    """'Claims Call Centre - Inbound and Outbound Call Report - 1-30 September 2026.xlsx'."""
    parts = [_clean(department), _clean(title or DEFAULT_TITLE) or DEFAULT_TITLE, period_label(period)]
    return " - ".join(p for p in parts if p) + f".{ext}"


def _sheet_name(department: str) -> str:
    return re.sub(r"[\[\]:*?/\\]", "", department).strip()[:31] or "Calls"


def build_perf_xlsx(res: PerfResult, title: str = DEFAULT_TITLE) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    df = res.detail
    if len(df) > MAX_EXCEL_ROWS:
        raise ReportError(f"This report has {len(df):,} rows, which is more than an Excel sheet can hold "
                          f"({MAX_EXCEL_ROWS:,}). Choose a smaller department.")
    wb = Workbook()
    ws = wb.active
    ws.title = _sheet_name(res.department)

    head_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor=HEADER_FILL)
    body_font = Font(name="Arial", size=10)
    ws.append(DETAIL_COLS)
    for cell in ws[1]:
        cell.font, cell.fill, cell.alignment = head_font, head_fill, Alignment(horizontal="center", vertical="center")

    centre, left = Alignment(horizontal="center"), Alignment(horizontal="left")
    times = [None if pd.isna(t) else t.to_pydatetime() for t in df["Time"]]
    durations = {c: [d.to_pytimedelta() for d in df[c]] for c in DURATION_COLS}
    for i, rec in enumerate(df[[c for c in DETAIL_COLS if c not in DURATION_COLS + ["Time"]]]
                            .to_dict("records")):
        row = [rec["ID"], times[i], rec["Consultants"], rec["Call From"], rec["Call To"],
               durations["Call Duration"][i], durations["Ring Duration"][i], durations["Talk Duration"][i],
               rec["Communication Type"], rec["Status"]]
        ws.append(row)
        for col, cell in enumerate(ws[ws.max_row], start=1):
            cell.font = body_font
            name = DETAIL_COLS[col - 1]
            if name == "ID":
                cell.number_format, cell.data_type = "@", "s"  # call IDs are text, never numbers
            elif name == "Time":
                cell.number_format, cell.alignment = TIME_FORMAT, left
            elif name in DURATION_COLS:
                cell.number_format, cell.alignment = DURATION_FORMAT, centre

    for col, name in enumerate(DETAIL_COLS, start=1):
        width = BASE_WIDTHS[name]
        if name in WIDE_COLS and len(df):
            width = min(max(width, int(df[name].astype(str).str.len().max()) + 2), MAX_WIDTH)
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(DETAIL_COLS))}{max(ws.max_row, 2)}"
    wb.properties.title = f"{res.department} - {title or DEFAULT_TITLE} - {period_label(res.period)}"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
