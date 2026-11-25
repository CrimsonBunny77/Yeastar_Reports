"""PDF (ReportLab) and Excel (.xlsx) output for the Queue overflow report."""
from __future__ import annotations

import io
from datetime import datetime
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.lib.units import cm
from reportlab.platypus import CondPageBreak, PageBreak, Paragraph, SimpleDocTemplate, Spacer

from .calls_report import DT_FMT, _section, _table, _trim, safe_filename
from .overlap import overlap_note
from .pdf_report import H1, H2, NOTE, PAGE, SUB, TITLE, _logo_flowable, _make_footer
from .queue_overflow import (DETAIL_COLS, FINAL_COLS, PERCENT_COLS, QueueResult, ROUTING_COLS, ROUTING_TYPES,
                             SUMMARY_COLS)

DEFAULT_TITLE = "Queue Overflow Report"


def queue_filename(title: str, period: pd.Period, ext: str) -> str:
    return safe_filename(title or DEFAULT_TITLE, "", period, ext)


def _wrap(text: str) -> Paragraph:
    return Paragraph(escape(text), NOTE)


def _stat_rows(res: QueueResult, generated: datetime) -> list[list]:
    a = res.audit
    rows = [["Report period", res.period.strftime("%B %Y")], ["Generated", generated.strftime("%d/%m/%Y %H:%M")]]
    if a.first_call is not None:
        rows.append(["Data in the uploaded files covers", f"{a.first_call:%d/%m/%Y %H:%M} - {a.last_call:%d/%m/%Y %H:%M}"])
    return rows + [[k, f"{v:,}"] for k, v in res.stats.items()]


def _int(v) -> str:
    return f"{int(v):,}"


def _summary_rows(df: pd.DataFrame) -> list[list]:
    return [[r["Queue Number"], _trim(r["Queue"], 34)]
            + [f"{r[c]:.1f}%" if c in PERCENT_COLS else _int(r[c]) for c in SUMMARY_COLS[2:]]
            for _, r in df.iterrows()]


def _final_rows(df: pd.DataFrame) -> list[list]:
    return [[r["Queue Number"], _trim(r["Queue"], 34)] + [_int(r[c]) for c in FINAL_COLS[2:]]
            for _, r in df.iterrows()]


def _routing_rows(df: pd.DataFrame) -> list[list]:
    return [[r["Queue Number"], _trim(r["Queue"], 34)] + [_int(r[c]) for c in ROUTING_COLS[2:]]
            for _, r in df.iterrows()]


def _detail_rows(df: pd.DataFrame) -> list[list]:
    return [[r["Time"].strftime(DT_FMT) if pd.notna(r["Time"]) else "", r["Call ID"], r["Caller"], r["Queue Number"],
             _trim(r["Agent"], 30), r["Status"], r["Ring Duration"], r["Talk Duration"], r["Routing Type"],
             r["Final Status"]] for _, r in df.iterrows()]


def _exception_rows(df: pd.DataFrame) -> list[list]:
    return [[_trim(r["Type"], 40), r["Queue Number"], r["Call ID"],
             r["Call Time"].strftime(DT_FMT) if pd.notna(r["Call Time"]) else "", _trim(r["Detail"], 78)]
            for _, r in df.iterrows()]


def build_queue_pdf(res: QueueResult, title: str = DEFAULT_TITLE, logo: bytes | None = None,
                    generated: datetime | None = None) -> bytes:
    generated = generated or datetime.now()
    title = title or DEFAULT_TITLE
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=PAGE, leftMargin=1.5 * cm, rightMargin=1.5 * cm, topMargin=1.5 * cm,
                            bottomMargin=1.8 * cm, title=title, author="Yeastar Transfer Report App")
    story: list = []

    # Cover
    queues = ", ".join(f"{n} {name}" for n, name in zip(res.queues["Queue Number"], res.queues["Queue Name"]))
    story += [Spacer(1, 0.5 * cm), _logo_flowable(logo), Spacer(1, 2 * cm)] if logo else [Spacer(1, 4.5 * cm)]
    story += [Paragraph(escape(title), TITLE), Spacer(1, 1 * cm),
              Paragraph(f"<b>Month:</b> {res.period.strftime('%B %Y')}", SUB),
              Paragraph(f"<b>Generated:</b> {generated.strftime('%d %B %Y, %H:%M')}", SUB),
              Spacer(1, 0.4 * cm),
              Paragraph(f"<b>Queues included ({len(res.queues)}):</b><br/>{escape(_trim(queues, 400))}", SUB),
              PageBreak()]

    # Report period, headline figures and the definitions used
    story += [Paragraph("Report period and definitions", H1),
              _table(["Measure", "Value"], _stat_rows(res, generated), [11 * cm, 7 * cm], right=(1,)),
              PageBreak(), Paragraph("Definitions used in this report", H1),
              _table(["Term", "Definition"], [[_wrap(k), _wrap(v)] for k, v in res.definitions],
                     [4.5 * cm, 22 * cm]),
              PageBreak()]

    # Summary by queue (queue visits and how each ended), routing type, then unique calls by final status
    story += [Paragraph("Summary by queue", H1)]
    # Queue Number | Queue | Total Calls | Answered, Missed, Abandoned, Other | 3 rates | Unique | agent | out
    w = [1.6 * cm, 4.4 * cm, 1.8 * cm, 1.8 * cm, 1.6 * cm, 1.9 * cm, 1.4 * cm, 1.9 * cm, 1.7 * cm, 2.0 * cm, 1.7 * cm,
         2.3 * cm, 2.0 * cm]
    head = {"Answered by Queue Agent": "Answered by\nQueue Agent", "Queue": "Queue"}
    story += [_table([head.get(c, c.replace(" ", "\n", 1)) for c in SUMMARY_COLS],
                     _summary_rows(res.summary), w,
                     right=tuple(range(2, len(SUMMARY_COLS)))),
              Spacer(1, 0.15 * cm),
              Paragraph("Total Calls counts queue visits, as the PBX's Queue Performance report does: a call that "
                        "re-enters the queue is counted each time. Answered, Missed and Abandoned are how each visit "
                        "ended. Unique Calls counts each Call ID once; the last row counts a call once even if it "
                        f"entered several selected queues. Visits that rang {res.min_ring_s} second(s) or less are "
                        "not counted.", NOTE),
              CondPageBreak(7 * cm), Spacer(1, 0.6 * cm), Paragraph("Routing type by queue", H2),
              _table(["Queue\nNumber", "Queue", "Unique\nCalls", "Visits"] + ROUTING_TYPES,
                     _routing_rows(res.routing),
                     [1.9 * cm, 6.3 * cm, 2.6 * cm, 2.6 * cm] + [3.3 * cm] * len(ROUTING_TYPES),
                     right=tuple(range(2, len(ROUTING_COLS)))),
              Spacer(1, 0.15 * cm),
              Paragraph("Routing types count queue visits (they add up to the Visits column). A manual transfer is "
                        "never counted as automatic overflow.", NOTE),
              PageBreak(), Paragraph("Unique calls by final status", H1),
              _table([c.replace(" ", "\n", 1) for c in FINAL_COLS], _final_rows(res.final),
                     [1.9 * cm, 6.0 * cm, 2.0 * cm] + [2.1 * cm] * (len(FINAL_COLS) - 3),
                     right=tuple(range(2, len(FINAL_COLS)))),
              Spacer(1, 0.15 * cm),
              Paragraph("Each call is counted once per queue by Call ID, by the Last Status on the call's summary row "
                        "(how the whole call ended). This can differ from the visit status above.", NOTE),
              PageBreak()]

    # Detail, exceptions
    _section(story, "Call detail (one row per queue visit)", res.detail, _detail_rows,
             ["Time", "Call ID", "Caller", "Queue", "Agent", "Status", "Ring", "Talk", "Routing Type", "Final Status"],
             [3.2 * cm, 3.9 * cm, 2.7 * cm, 1.5 * cm, 3.7 * cm, 1.9 * cm, 1.7 * cm, 1.7 * cm, 2.9 * cm, 2.4 * cm],
             (),
             "Oldest first; Time is when the call entered the queue. Routing notes, call type, hold time and "
             "filters are in the Excel file.")
    story.append(PageBreak())
    _section(story, "Exceptions to review", res.exceptions, _exception_rows,
             ["Type", "Queue", "Call ID", "Call Time", "Detail"],
             [6.0 * cm, 1.5 * cm, 3.9 * cm, 3.2 * cm, 11.6 * cm], (),
             "Calls and queue-mapping rows that need a look. They are shown here rather than dropped silently.")

    # Audit
    a = res.audit
    story += [PageBreak(), Paragraph("Audit", H1)]
    rows = [["Files processed", f"{a.files_processed:,}"], ["Total rows imported", f"{a.total_rows:,}"],
            ["Call records in files", f"{a.total_calls:,}"]]
    if a.first_call is not None:
        rows.append(["Data covers", f"{a.first_call:%d/%m/%Y %H:%M} - {a.last_call:%d/%m/%Y %H:%M}"])
    if a.overlap_calls:
        rows.append(["Calls in more than one file (counted once)" if a.duplicates_removed
                     else "Calls in more than one file (NOT removed)", f"{a.overlap_calls:,}"])
    rows += [["Excluded: calls outside the report month", f"{a.outside_period:,}"],
             [f"Not counted: queue visits that rang {res.min_ring_s} second(s) or less", f"{a.short_visits:,}"],
             ["Included: calls that are not inbound (internal / outbound)", f"{a.non_inbound:,}"],
             ["Not reported: inbound calls to a queue missing from the mapping", f"{a.unmapped_queue_calls:,}"]]
    if a.duplicates_removed:
        rows.append(["Duplicate call records removed", f"{a.duplicates_removed:,}"])
    story.append(_table(["Measure", "Value"], rows, [11 * cm, 7 * cm], right=(1,)))
    if a.overlaps:
        story.append(Paragraph("Overlapping files: " + escape(overlap_note(a.overlaps)), NOTE))
    if a.per_file:
        story += [Spacer(1, 0.5 * cm),
                  _table(["File", "Rows", "Calls", "Queue visits"],
                         [[_trim(f["File"], 60), f"{f['Rows']:,}", f"{f['Calls']:,}", f"{f['Queue visits']:,}"]
                          for f in a.per_file], [12 * cm, 3.5 * cm, 3.5 * cm, 3.5 * cm], right=(1, 2, 3))]
    if a.files_failed:
        story += [Spacer(1, 0.5 * cm), Paragraph("Files skipped: " + "; ".join(n for n, _ in a.files_failed), NOTE)]

    footer = _make_footer(f"{title} - {res.period.strftime('%B %Y')}")
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


def build_queue_xlsx(res: QueueResult, title: str = DEFAULT_TITLE, generated: datetime | None = None) -> bytes:
    """Excel workbook: Summary, By queue, Routing type, Final status, Detail (filterable), Exceptions, Definitions."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    generated = generated or datetime.now()
    a = res.audit
    queues = ", ".join(f"{n} {name}" for n, name in zip(res.queues["Queue Number"], res.queues["Queue Name"]))
    summary = [("Report", title or DEFAULT_TITLE), ("Month", res.period.strftime("%B %Y")),
               ("Generated", generated.strftime("%d/%m/%Y %H:%M")), ("Queues included", queues)]
    if a.first_call is not None:
        summary.append(("Data in the uploaded files covers",
                        f"{a.first_call:%d/%m/%Y %H:%M} - {a.last_call:%d/%m/%Y %H:%M}"))
    summary += [(k, v) for k, v in res.stats.items()]
    summary += [("Excluded: calls outside the report month", a.outside_period),
                (f"Not counted: queue visits that rang {res.min_ring_s} second(s) or less", a.short_visits),
                ("Included: calls that are not inbound (internal / outbound)", a.non_inbound),
                ("Not reported: inbound calls to a queue missing from the mapping", a.unmapped_queue_calls),
                ("Duplicate call records removed", a.duplicates_removed)]
    sheets = {
        "Summary": pd.DataFrame(summary, columns=["Measure", "Value"]).astype({"Value": str}),
        "By queue": res.summary,
        "Routing type": res.routing,
        "Final status": res.final,
        "Detail": res.detail[DETAIL_COLS],
        "Exceptions": res.exceptions,
        "Definitions": pd.DataFrame(res.definitions, columns=["Term", "Definition"]),
    }
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
                if ws.cell(row=1, column=col[0].column).value in ("Time", "Call Time"):
                    for cell in col:
                        cell.number_format = "dd/mm/yyyy hh:mm:ss"
            ws.freeze_panes = "A2"
            if ws.max_row > 1 and ws.max_column > 0:
                ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                width = max((len(str(c.value)) for c in col[:300] if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 12), 60)
        defs = xw.book["Definitions"]
        for row in defs.iter_rows(min_row=2, min_col=2, max_col=2):
            row[0].alignment = Alignment(wrap_text=True, vertical="top")
        defs.column_dimensions["B"].width = 110
    return buf.getvalue()
