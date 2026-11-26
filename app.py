"""Yeastar Transfer Report - Streamlit app.  Run with:  streamlit run app.py"""
from __future__ import annotations

import calendar
import io
import traceback
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

import streamlit as st

from transfer_report.calls import ALL_DEPARTMENTS, process_calls
from transfer_report.calls_report import (DEFAULT_TITLE as CALLS_TITLE, build_calls_pdf, build_calls_xlsx,
                                          safe_filename)
from transfer_report.charts import build_chart_specs
from transfer_report.ext_names import DEFAULT_NAMES_PATH, load_extension_names
from transfer_report.ext_performance import preview_frame, process_ext_performance
from transfer_report.ext_performance_report import (DEFAULT_TITLE as PERF_TITLE, build_perf_xlsx, perf_filename)
from transfer_report.ext_stats import process_ext_stats
from transfer_report.ext_stats_report import (DEFAULT_TITLE as EXT_TITLE, build_ext_stats_pdf,
                                              build_ext_stats_xlsx, ext_filename)
from transfer_report.config import (DEFAULT_CDR_DIR, DEFAULT_EXCLUDED_DAYS_PATH, DEFAULT_MAPPING_PATH,
                                    DEFAULT_QUEUE_MAPPING_PATH)
from transfer_report.hours import (CLOSED, DAY_TYPES, DEFAULT_END, DEFAULT_START, build_office_hours, days_for_year,
                                   load_days, save_days)
from transfer_report.errors import MappingError, ReportError
from transfer_report.exports import unmapped_to_xlsx
from transfer_report.mapping import create_template, list_departments, load_mapping
from transfer_report.pdf_report import build_pdf, prepare_logo, report_filename
from transfer_report.processor import MATCH_BOTH, MATCH_EITHER, process_files
from transfer_report.queue_mapping import load_queue_mapping, queue_label
from transfer_report.queue_overflow import DETAIL_COLS as QUEUE_DETAIL_COLS
from transfer_report.queue_overflow import MIN_RING_S, PERCENT_COLS, ROUTING_TYPES, filter_detail, process_queue_overflow
from transfer_report.queue_report import (DEFAULT_TITLE as QUEUE_TITLE, build_queue_pdf, build_queue_xlsx,
                                          queue_filename)

DEFAULT_LOGO_PATH = Path(__file__).resolve().parent / "logo.png"  # shown on every report unless switched off

st.set_page_config(page_title="Yeastar Transfer Report", page_icon="📞", layout="wide")
st.title("📞 Yeastar Reports")
REPORT_TYPES = ["Transfer report", "Department call list", "Extension call stats", "Queue overflow report",
                "Extension Performance Detail"]
mode = st.radio("Report type", REPORT_TYPES, horizontal=True, key="report_type")
TRANSFER = mode == REPORT_TYPES[0]
EXT_STATS = mode == REPORT_TYPES[2]
QUEUE = mode == REPORT_TYPES[3]
EXT_PERF = mode == REPORT_TYPES[4]
st.caption({
    REPORT_TYPES[0]: "Upload Yeastar CDR exports, combine them, and generate an inter-department transfer report.",
    REPORT_TYPES[1]: "Upload Yeastar CDR exports and list every number a department dialled (outbound) or answered (inbound).",
    REPORT_TYPES[2]: "Upload Yeastar CDR exports and see each extension's calls and outcomes for a department.",
    REPORT_TYPES[3]: "Upload Yeastar CDR exports and a queue mapping to see the calls routed to chosen queues and "
                     "whether they overflowed automatically or were transferred by an agent.",
    REPORT_TYPES[4]: "Upload Yeastar CDR exports and get every inbound and outbound call of a department's "
                     "consultants, one row per call, as an Excel workbook.",
}[mode])

# ---------------------------------------------------------------- sidebar
mapping, map_warnings, departments = None, [], []
queue_map, queue_warnings, queue_issues, queues = None, [], None, []
with st.sidebar:
    st.header("Settings")
    if QUEUE:
        mapping_upload = None
        queue_upload = st.file_uploader(
            "Override queue mapping (.xlsx)", type=["xlsx"],
            help="Optional. Defaults to Que_Mapping_Table.xlsx next to app.py. Needs a queue number column "
                 "(e.g. 'Que Number') and a queue name column (e.g. 'Queue Name' or 'Department').")
    else:
        mapping_upload = st.file_uploader("Override department mapping (.xlsx)", type=["xlsx"],
                                          help="Optional. Defaults to department_mapping.xlsx next to app.py.")
    date_order = st.radio("Date format in 'Time' column", ["Day first (31/12/2026)", "Month first (12/31/2026)"])
    dedupe = st.checkbox("Count each call once when files overlap (recommended)", value=True,
                         help="If the same call appears in more than one uploaded file (for example two exports "
                              "that cover the same days), it is counted once. Overlaps are always reported.")
    st.subheader("Logo")
    use_logo = st.checkbox("Show logo on the report cover", value=True)
    logo_file = st.file_uploader("Use a different logo (optional, PNG or JPG)", type=["png", "jpg", "jpeg"],
                                 help="Replaces the default logo (logo.png next to app.py) for this session.")
    logo = None
    if use_logo:
        try:
            if logo_file:
                logo = prepare_logo(logo_file.getvalue())
            elif DEFAULT_LOGO_PATH.exists():
                logo = prepare_logo(DEFAULT_LOGO_PATH.read_bytes())
            if logo:
                st.image(logo, width=160)
        except ReportError as exc:
            st.warning(str(exc))
    st.divider()

    # The mapping is loaded on every run so the department list always reflects the file in use.
    if QUEUE:
        # The queue report uses its own queue mapping instead of the department mapping.
        try:
            source = io.BytesIO(queue_upload.getvalue()) if queue_upload else DEFAULT_QUEUE_MAPPING_PATH
            queue_map, queue_warnings, queue_issues = load_queue_mapping(source)
            queues = list(queue_map["Queue Number"])
        except MappingError as exc:
            st.error(f"Queue mapping problem: {exc}")
    else:
        try:
            source = io.BytesIO(mapping_upload.getvalue()) if mapping_upload else DEFAULT_MAPPING_PATH
            mapping, map_warnings = load_mapping(source)
            departments = list_departments(mapping)
        except MappingError as exc:
            st.error(f"Department mapping problem: {exc}")
            if not DEFAULT_MAPPING_PATH.exists() and st.button("Create sample department_mapping.xlsx"):
                create_template(DEFAULT_MAPPING_PATH)
                st.rerun()

    names, names_warning = None, None
    if EXT_STATS or EXT_PERF:
        st.subheader("Extension names")
        names_upload = st.file_uploader(
            "Use a different names file (CSV or XLSX)", type=["csv", "xlsx"],
            help="Defaults to extension_names.xlsx next to app.py. Accepts an Extension/Name sheet or a Yeastar "
                 "user export (only extension numbers and names are read).")
        try:
            if names_upload:
                names = load_extension_names(io.BytesIO(names_upload.getvalue()), names_upload.name)
            elif DEFAULT_NAMES_PATH.exists():
                names = load_extension_names(DEFAULT_NAMES_PATH)
            else:
                names_warning = "No extension names file found; names will be taken from the CDR where possible."
            if names is not None:
                st.caption(f"{len(names):,} extension names loaded.")
        except MappingError as exc:
            names_warning = f"Extension names problem: {exc} Names will be taken from the CDR where possible."
        if names_warning:
            st.warning(names_warning)

    selected = departments
    match = MATCH_BOTH
    if departments and TRANSFER:
        selected = st.multiselect(
            "Departments to include in the report", departments, default=departments,
            key=f"depts_{hash(tuple(departments))}",
            help="Populated from the department mapping file.",
        )
        if len(selected) < len(departments):
            choice = st.radio("Include transfers where", ["Both departments are selected",
                                                          "Either department is selected"])
            match = MATCH_BOTH if choice.startswith("Both") else MATCH_EITHER
        if not selected:
            st.warning("Select at least one department.")
        elif match == MATCH_BOTH and len(selected) == 1:
            st.warning("With one department and 'Both', no transfers can match. Choose 'Either' or add departments.")

    selected_queues = []
    min_ring_s = MIN_RING_S
    if QUEUE and queue_map is not None:
        queue_labels = {n: queue_label(n, nm) for n, nm in zip(queue_map["Queue Number"], queue_map["Queue Name"])}
        st.caption(f"{len(queues):,} queues loaded from the queue mapping.")
        selected_queues = st.multiselect(
            "Queues to include in the report", queues, default=queues, format_func=queue_labels.get,
            key=f"queues_{hash(tuple(queues))}",
            help="Populated from the queue mapping file (queue number and name). Untick any you want left out.",
        )
        if not selected_queues:
            st.warning("Select at least one queue.")
        min_ring_s = int(st.number_input(
            "Don't count visits that rang for this many seconds or less", min_value=0, max_value=60, value=MIN_RING_S,
            step=1, help="Matches the PBX's 'exclude calls within X seconds' settings (1 by default). "
                         "Excluded visits are listed on the Exceptions tab."))

    # Office hours (Extension call stats and Queue overflow only)
    office_hours = None
    if EXT_STATS or QUEUE:
        st.divider()
        st.subheader("Office hours")
        apply_hours = st.checkbox(
            "Exclude after-hours, weekends and excluded days", value=True, key="apply_hours",
            help="When ticked, the report's tables count only calls that started during office hours. Calls after "
                 "hours, on weekends and on the excluded days below are still shown, in a separate section. "
                 "Untick to count every call, like the PBX's own reports.")
        h1, h2 = st.columns(2)
        open_at = h1.time_input("Opens", value=DEFAULT_START, step=timedelta(minutes=15), key="hours_open",
                                disabled=not apply_hours)
        close_at = h2.time_input("Closes", value=DEFAULT_END, step=timedelta(minutes=15), key="hours_close",
                                 disabled=not apply_hours, help="A call at this time or later is after hours.")
        days_year = int(st.session_state.get("report_year") or (date.today().replace(day=1) - timedelta(days=1)).year)
        base_key = f"days_base_{days_year}"
        if base_key not in st.session_state:  # kept for the session so edits are never applied twice
            st.session_state[base_key] = days_for_year(days_year, load_days(DEFAULT_EXCLUDED_DAYS_PATH))
        with st.expander(f"Excluded and short days ({days_year})", expanded=False):
            st.caption("Pre-filled with South Africa's public holidays for the report year, or with the days you "
                       "saved. Add rows for once-off closures. For a short day choose 'Short day' and enter the "
                       "closing time; calls after it count as after hours.")
            days = st.data_editor(
                st.session_state[base_key], num_rows="dynamic", hide_index=True, key=f"days_editor_{days_year}",
                disabled=not apply_hours, use_container_width=True,
                column_config={
                    "Date": st.column_config.DateColumn("Date", format="DD/MM/YYYY", required=True),
                    "Type": st.column_config.SelectboxColumn("Type", options=DAY_TYPES, default=CLOSED, required=True),
                    "Closes At": st.column_config.TimeColumn("Closes at", format="HH:mm",
                                                             help="Short days only: the early closing time."),
                    "Reason": st.column_config.TextColumn("Reason"),
                })
            if st.button("Save these days for next time", disabled=not apply_hours, key="save_days",
                         help=f"Saves the list for {days_year} to excluded_days.xlsx next to app.py."):
                try:
                    save_days(DEFAULT_EXCLUDED_DAYS_PATH, days_year, days)
                    st.success(f"Saved the excluded and short days for {days_year}.")
                except OSError as exc:
                    st.error(f"Could not save excluded_days.xlsx: {exc}")
        if apply_hours:
            office_hours, hours_warnings = build_office_hours(days, open_at, close_at)
            for w in hours_warnings:
                st.warning(w)

# ---------------------------------------------------------------- upload
files = st.file_uploader("Drag and drop Yeastar CDR files here", type=["xls", "xlsx"],
                         accept_multiple_files=True)
stored_cdr = []  # queue report only: CDR exports stored next to the app (CDR_Example folder)
if QUEUE and not files and DEFAULT_CDR_DIR.exists():
    stored = sorted(p for p in DEFAULT_CDR_DIR.iterdir() if p.suffix.lower() in (".xls", ".xlsx"))
    if stored:
        chosen = st.multiselect("Or select a CDR export stored with the app (CDR_Example folder)",
                                [p.name for p in stored], help="Used only when nothing is uploaded above.")
        stored_cdr = [p for p in stored if p.name in chosen]
if files:
    st.write(f"**{len(files)}** file(s) selected: " + ", ".join(f.name for f in files))
elif stored_cdr:
    st.write(f"**{len(stored_cdr)}** stored file(s) selected: " + ", ".join(p.name for p in stored_cdr))
else:
    st.info("Upload one or more .xls / .xlsx files to begin.")

# Report name and month: set before generating
prev_month = date.today().replace(day=1) - timedelta(days=1)
c1, c2, c3 = st.columns([3, 2, 1])
report_title = c1.text_input("Report name", value={REPORT_TYPES[0]: "Yeastar Transfer Report", REPORT_TYPES[1]: CALLS_TITLE, REPORT_TYPES[2]: EXT_TITLE,
                                    REPORT_TYPES[3]: QUEUE_TITLE, REPORT_TYPES[4]: PERF_TITLE}[mode],
                             key=f"title_{mode}",
                             help="Shown on the cover page and used in the file names.").strip()
month_names = list(calendar.month_name)[1:]
month_name = c2.selectbox("Report month", month_names, index=None, placeholder="Please Select Month")
year = int(c3.number_input("Year", min_value=2000, max_value=2100, value=prev_month.year, step=1, key="report_year"))
# The month is mandatory: no report can be generated until one is chosen.
period = pd.Period(year=year, month=month_names.index(month_name) + 1, freq="M") if month_name else None
if period is None:
    st.warning("Please select the report month to enable report generation.")



# ================================================================ transfer report
def render_transfer_report():
    can_run = period is not None and bool(files) and mapping is not None and bool(selected)
    generate = st.button("Generate Transfer Report", type="primary", disabled=not can_run)

    # ---------------------------------------------------------------- processing
    if generate:
        st.session_state.pop("result", None)
        bar = st.progress(0.0, text="Starting…")
        try:
            payload = [(f.name, f.getvalue()) for f in files]
            result = process_files(
                payload, mapping, dayfirst=date_order.startswith("Day"), dedupe=dedupe,
                departments=selected, match=match, period=period,
                progress=lambda frac, msg: bar.progress(min(frac, 1.0), text=msg),
            )
            bar.progress(0.93, text="Building charts and PDF…")
            specs = build_chart_specs(result.detail)
            pdf = build_pdf(result.detail, result.audit, result.period, specs,
                            scope=result.scope, match=result.match,
                            title=report_title or "Yeastar Transfer Report",
                            status_summary=result.status_summary, logo=logo)
            unmapped_xlsx = (unmapped_to_xlsx(result.unmapped_summary, result.unmapped_records)
                             if result.audit.unmapped_removed else None)
            bar.progress(1.0, text="Done")
            st.session_state["result"] = dict(
                result=result, specs=specs, pdf=pdf, name=report_filename(result.period, report_title or "Yeastar Transfer Report"),
                warnings=map_warnings + result.warnings, unmapped_xlsx=unmapped_xlsx,
            )
        except ReportError as exc:
            bar.empty()
            st.error(str(exc))
        except Exception:  # noqa: BLE001
            bar.empty()
            st.error("Something unexpected went wrong while generating the report.")
            with st.expander("Technical details"):
                st.code(traceback.format_exc())

    # ---------------------------------------------------------------- results
    state = st.session_state.get("result")
    st.divider()
    st.download_button(
        "Download PDF",
        data=state["pdf"] if state else b"",
        file_name=state["name"] if state else "Transfer_Report.pdf",
        mime="application/pdf",
        disabled=state is None,
        type="primary",
    )

    if state:
        result, audit = state["result"], state["result"].audit
        st.success(f"Report ready for {result.period.strftime('%B %Y')}: {state['name']}")
        if result.scope:
            st.info("Departments included: " + ", ".join(result.scope))
        for w in state["warnings"]:
            st.warning(w)
        for name, msg in audit.files_failed:
            st.warning(f"Skipped: {msg}")

        c = st.columns(5)
        c[0].metric("Files processed", audit.files_processed)
        c[1].metric("Rows imported", f"{audit.total_rows:,}")
        c[2].metric("Blind transfers", f"{audit.blind:,}")
        c[3].metric("Attended transfers", f"{audit.attended:,}")
        c[4].metric("Final transfer records", f"{audit.final_records:,}")
        st.caption(
            f"Excluded: {audit.unmapped_removed:,} unmapped · {audit.internal_removed:,} same-department · "
            f"{audit.deselected_removed:,} departments not selected"
            + (f" · {audit.duplicates_removed:,} duplicates" if audit.duplicates_removed else "")
        )

        tabs = st.tabs(["Transfer matrix", "Charts", f"Unmapped extensions ({len(result.unmapped_summary):,})", "Status values"])
        with tabs[0]:
            st.dataframe(result.detail, use_container_width=True, hide_index=True)
        with tabs[1]:
            try:
                for spec in state["specs"]:
                    st.plotly_chart(spec.figure(), use_container_width=True)
            except Exception:  # noqa: BLE001
                st.info("Charts are included in the PDF; interactive preview is unavailable.")
        with tabs[2]:
            if result.unmapped_summary.empty:
                st.success("Every extension in the transfer records was found in the department mapping.")
            else:
                st.write(
                    f"**{audit.unmapped_removed:,}** transfer record(s) were excluded because one or both "
                    "extensions are not in the mapping. Add these extensions to `department_mapping.xlsx` "
                    "(or upload an updated mapping) and generate again."
                )
                st.dataframe(result.unmapped_summary, use_container_width=True, hide_index=True)
                st.download_button(
                    "Download unmapped list (Excel)", data=state["unmapped_xlsx"],
                    file_name="Unmapped_Extensions.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.caption("The Excel file has a summary, every excluded record, and a 'Mapping additions' "
                           "sheet you can fill in and paste into the mapping file.")
        with tabs[3]:
            st.write("Every `Last Status` value found on the final transfer records, and the column it is counted under.")
            st.dataframe(result.status_summary, use_container_width=True, hide_index=True)



# ================================================================ department call list
def render_call_list():
    department = st.selectbox("Department", [ALL_DEPARTMENTS] + departments, key="calls_department",
                              help="Populated from the department mapping file. Choose 'All departments' for "
                                   "one combined list.") if departments else None
    can_run = period is not None and bool(files) and mapping is not None and department is not None
    generate = st.button("Generate Call List", type="primary", disabled=not can_run)

    if generate:
        st.session_state.pop("calls_state", None)
        bar = st.progress(0.0, text="Starting…")
        try:
            payload = [(f.name, f.getvalue()) for f in files]
            res = process_calls(
                payload, mapping, department, dayfirst=date_order.startswith("Day"), dedupe=dedupe,
                period=period, progress=lambda frac, msg: bar.progress(min(frac, 1.0), text=msg),
            )
            bar.progress(0.93, text="Building PDF and Excel…")
            title = report_title or CALLS_TITLE
            st.session_state["calls_state"] = dict(
                res=res,
                pdf=build_calls_pdf(res, title, logo),
                xlsx=build_calls_xlsx(res, title),
                pdf_name=safe_filename(title, department, res.period, "pdf"),
                xlsx_name=safe_filename(title, department, res.period, "xlsx"),
                warnings=map_warnings + res.warnings,
            )
            bar.progress(1.0, text="Done")
        except ReportError as exc:
            bar.empty()
            st.error(str(exc))
        except Exception:  # noqa: BLE001
            bar.empty()
            st.error("Something unexpected went wrong while generating the report.")
            with st.expander("Technical details"):
                st.code(traceback.format_exc())

    state = st.session_state.get("calls_state")
    st.divider()
    d1, d2, _ = st.columns([1, 1, 2])
    d1.download_button("Download PDF", data=state["pdf"] if state else b"",
                       file_name=state["pdf_name"] if state else "Call_List.pdf", mime="application/pdf",
                       disabled=state is None, type="primary", key="dl_calls_pdf")
    d2.download_button("Download Excel (.xlsx)", data=state["xlsx"] if state else b"",
                       file_name=state["xlsx_name"] if state else "Call_List.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       disabled=state is None, key="dl_calls_xlsx")
    if not state:
        return

    res, audit = state["res"], state["res"].audit
    st.success(f"Call list ready for {res.department} - {res.period.strftime('%B %Y')}")
    for w in state["warnings"]:
        st.warning(w)
    for _, msg in audit.files_failed:
        st.warning(f"Skipped: {msg}")

    s = res.stats
    c = st.columns(5)
    c[0].metric("Outbound calls", f"{s['Outbound calls']:,}")
    c[1].metric("Outbound answered", f"{s['Outbound answered']:,}")
    c[2].metric("Unique numbers dialled", f"{s['Unique numbers dialled']:,}")
    c[3].metric("Inbound calls connected", f"{s['Inbound calls connected']:,}")
    c[4].metric("Unique inbound callers", f"{s['Unique inbound callers']:,}")
    st.caption(res.scope_text)

    names = ["Outbound numbers", "Inbound numbers", "Outbound call log", "Inbound call log",
             f"Unmapped extensions ({len(res.unmapped):,})"]
    if res.all_departments:
        names.insert(0, "By department")
    tabs = dict(zip(names, st.tabs(names)))
    if res.all_departments:
        with tabs["By department"]:
            st.dataframe(res.by_department, use_container_width=True, hide_index=True)
            st.caption("An inbound call answered by more than one department is counted once for each of them "
                       "here, and once in the totals above.")
    with tabs["Outbound numbers"]:
        st.dataframe(res.outbound_numbers, use_container_width=True, hide_index=True)
    with tabs["Inbound numbers"]:
        st.dataframe(res.inbound_numbers, use_container_width=True, hide_index=True)
    with tabs["Outbound call log"]:
        st.dataframe(res.outbound[res.display_cols], use_container_width=True, hide_index=True)
    with tabs["Inbound call log"]:
        st.dataframe(res.inbound[res.display_cols], use_container_width=True, hide_index=True)
    with tabs[names[-1]]:
        if res.unmapped.empty:
            st.success("Every extension seen on these calls was found in the department mapping.")
        else:
            st.write("These extensions made or answered calls but are not in the department mapping, so their "
                     "calls cannot be attributed to a department. Add the ones that belong to a department to "
                     "`department_mapping.xlsx` and generate again. (IVRs and queues are already left out.)")
            st.dataframe(res.unmapped, use_container_width=True, hide_index=True)


# ================================================================ extension call stats
def render_ext_stats():
    department = st.selectbox("Department", departments, key="ext_department",
                              help="Populated from the department mapping file.") if departments else None
    o1, o2, _ = st.columns([1, 1, 2])
    include_internal = o1.checkbox("Include internal calls", value=True, key="ext_internal",
                                   help="Counts calls between extensions too. The PBX's own report ('All') includes them.")
    hide_idle = o2.checkbox("Hide extensions with no calls", value=False, key="ext_hide_idle")
    can_run = period is not None and bool(files) and mapping is not None and department is not None
    generate = st.button("Generate Extension Stats", type="primary", disabled=not can_run)

    if generate:
        st.session_state.pop("ext_state", None)
        bar = st.progress(0.0, text="Starting…")
        try:
            payload = [(f.name, f.getvalue()) for f in files]
            res = process_ext_stats(
                payload, mapping, names, department, include_internal=include_internal, hide_idle=hide_idle,
                dedupe=dedupe, period=period, dayfirst=date_order.startswith("Day"),
                progress=lambda frac, msg: bar.progress(min(frac, 1.0), text=msg), office_hours=office_hours,
            )
            bar.progress(0.93, text="Building PDF and Excel…")
            title = report_title or EXT_TITLE
            st.session_state["ext_state"] = dict(
                res=res,
                pdf=build_ext_stats_pdf(res, title, logo),
                xlsx=build_ext_stats_xlsx(res, title),
                pdf_name=ext_filename(title, department, res.period, "pdf"),
                xlsx_name=ext_filename(title, department, res.period, "xlsx"),
                warnings=map_warnings + res.warnings,
            )
            bar.progress(1.0, text="Done")
        except ReportError as exc:
            bar.empty()
            st.error(str(exc))
        except Exception:  # noqa: BLE001
            bar.empty()
            st.error("Something unexpected went wrong while generating the report.")
            with st.expander("Technical details"):
                st.code(traceback.format_exc())

    state = st.session_state.get("ext_state")
    st.divider()
    d1, d2, _ = st.columns([1, 1, 2])
    d1.download_button("Download PDF", data=state["pdf"] if state else b"",
                       file_name=state["pdf_name"] if state else "Extension_Stats.pdf", mime="application/pdf",
                       disabled=state is None, type="primary", key="dl_ext_pdf")
    d2.download_button("Download Excel (.xlsx)", data=state["xlsx"] if state else b"",
                       file_name=state["xlsx_name"] if state else "Extension_Stats.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       disabled=state is None, key="dl_ext_xlsx")
    if not state:
        return

    res, audit = state["res"], state["res"].audit
    t = res.all_calls[["Total Calls", "Answered"]].sum()
    st.success(f"Extension stats ready for {res.department} - {res.period.strftime('%B %Y')}")
    for w in state["warnings"]:
        st.warning(w)
    for _, msg in audit.files_failed:
        st.warning(f"Skipped: {msg}")
    if audit.first_call is not None:
        st.caption(f"Data covers {audit.first_call:%d/%m/%Y %H:%M} - {audit.last_call:%d/%m/%Y %H:%M}")
    c = st.columns(4)
    c[0].metric("Extensions", f"{len(res.extensions):,}")
    c[1].metric("Total calls", f"{int(t['Total Calls']):,}")
    c[2].metric("Answered", f"{int(t['Answered']):,}")
    c[3].metric("Answer %", f"{(t['Answered'] / t['Total Calls'] * 100 if t['Total Calls'] else 0):.1f}")
    names_ = list(res.tables)
    if res.office_hours:
        out_total = int(res.outside["Total Calls"].iloc[-1])
        st.info(f"**Office hours only.** {out_total:,} call leg(s) outside office hours are counted separately on "
                f"the 'Outside office hours' tab. {res.hours_note}")
        names_.append("Outside office hours")
    tabs = dict(zip(names_, st.tabs(names_)))
    for name, df in res.tables.items():
        with tabs[name]:
            st.dataframe(df, use_container_width=True, hide_index=True)
    if res.office_hours:
        with tabs["Outside office hours"]:
            st.dataframe(res.outside, use_container_width=True, hide_index=True)
            st.caption("Not included in the other tabs. A call is counted once: excluded day first, then weekend, "
                       "then after hours.")
            st.markdown("**By extension** (after hours, weekends and excluded days together)")
            st.dataframe(res.outside_by_extension, use_container_width=True, hide_index=True)


# ================================================================ queue overflow report
def render_queue_overflow():
    can_run = period is not None and (bool(files) or bool(stored_cdr)) and queue_map is not None and bool(selected_queues)
    generate = st.button("Generate Queue Overflow Report", type="primary", disabled=not can_run)

    if generate:
        st.session_state.pop("queue_state", None)
        bar = st.progress(0.0, text="Starting…")
        try:
            payload = ([(f.name, f.getvalue()) for f in files] if files
                       else [(p.name, p.read_bytes()) for p in stored_cdr])
            res = process_queue_overflow(
                payload, queue_map, selected_queues, queue_issues, dayfirst=date_order.startswith("Day"),
                dedupe=dedupe, period=period, progress=lambda frac, msg: bar.progress(min(frac, 1.0), text=msg),
                min_ring_s=min_ring_s, office_hours=office_hours,
            )
            bar.progress(0.93, text="Building PDF and Excel…")
            title = report_title or QUEUE_TITLE
            st.session_state["queue_state"] = dict(
                res=res,
                pdf=build_queue_pdf(res, title, logo),
                xlsx=build_queue_xlsx(res, title),
                pdf_name=queue_filename(title, res.period, "pdf"),
                xlsx_name=queue_filename(title, res.period, "xlsx"),
                warnings=queue_warnings + res.warnings,
            )
            bar.progress(1.0, text="Done")
        except ReportError as exc:
            bar.empty()
            st.error(str(exc))
        except Exception:  # noqa: BLE001
            bar.empty()
            st.error("Something unexpected went wrong while generating the report.")
            with st.expander("Technical details"):
                st.code(traceback.format_exc())

    state = st.session_state.get("queue_state")
    st.divider()
    d1, d2, _ = st.columns([1, 1, 2])
    d1.download_button("Download PDF", data=state["pdf"] if state else b"",
                       file_name=state["pdf_name"] if state else "Queue_Overflow_Report.pdf", mime="application/pdf",
                       disabled=state is None, type="primary", key="dl_queue_pdf")
    d2.download_button("Download Excel (.xlsx)", data=state["xlsx"] if state else b"",
                       file_name=state["xlsx_name"] if state else "Queue_Overflow_Report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       disabled=state is None, key="dl_queue_xlsx")
    if not state:
        return

    res, audit = state["res"], state["res"].audit
    st.success(f"Queue overflow report ready for {res.period.strftime('%B %Y')}")
    covers = (f" · the uploaded files cover {audit.first_call:%d/%m/%Y %H:%M} to {audit.last_call:%d/%m/%Y %H:%M}"
              if audit.first_call is not None else "")
    st.info(f"**Report period: {res.period.strftime('%B %Y')}** (queue visits that entered the queue in this month){covers}")
    for w in state["warnings"]:
        st.warning(w)
    for _, msg in audit.files_failed:
        st.warning(f"Skipped: {msg}")

    s = res.stats
    c = st.columns(5)
    c[0].metric("Total calls (queue visits)", f"{s['Total calls (queue visits)']:,}")
    c[1].metric("Answered", f"{s['Answered']:,}")
    c[2].metric("Missed", f"{s['Missed']:,}")
    c[3].metric("Abandoned", f"{s['Abandoned']:,}")
    c[4].metric("Unique calls", f"{s['Unique calls']:,}")
    st.caption(f"{s['Queues selected']} queue(s) selected · {s['Calls with a recorded agent answer']:,} calls with a "
               f"recorded agent answer · {s['Entered by automatic overflow']:,} entered by automatic overflow · "
               f"{s['Entered by manual transfer']:,} by manual transfer · {s['Routing undetermined']:,} with an "
               f"undetermined routing type · {s['Exceptions to review']:,} exceptions to review")
    if res.office_hours:
        st.info(f"**Office hours only.** {len(res.outside_detail):,} queue visit(s) outside office hours are counted "
                f"separately on the 'Outside office hours' tab. {res.hours_note}")

    with st.expander("Definitions used in this report", expanded=True):
        for term, text in res.definitions:
            st.markdown(f"**{term}:** {text}")

    tab_names = ["Summary by queue", "Routing type", "Call detail", f"Exceptions ({len(res.exceptions):,})"]
    if res.office_hours:
        tab_names.insert(3, f"Outside office hours ({len(res.outside_detail):,})")
    tabs = st.tabs(tab_names)
    if res.office_hours:
        with tabs[3]:
            st.dataframe(res.outside, use_container_width=True, hide_index=True)
            st.caption("Queue visits that entered the queue outside office hours. They are not included in the other "
                       "tabs. A visit is counted once: excluded day first, then weekend, then after hours.")
            st.markdown("**By queue**")
            st.dataframe(res.outside_summary, use_container_width=True, hide_index=True,
                         column_config={c: st.column_config.NumberColumn(c, format="%.2f%%") for c in PERCENT_COLS})
            st.markdown("**Detail** (one row per queue visit)")
            st.dataframe(res.outside_detail[QUEUE_DETAIL_COLS], use_container_width=True, hide_index=True)
        tabs = tabs[:3] + tabs[4:]  # the remaining code addresses the exceptions tab as tabs[3]
    with tabs[0]:
        st.dataframe(res.summary, use_container_width=True, hide_index=True,
                     column_config={c: st.column_config.NumberColumn(c, format="%.2f%%") for c in PERCENT_COLS})
        st.caption("Total Calls counts queue visits, as the PBX's Queue Performance report does, so a call that "
                   "re-enters the queue is counted each time. Answered, Missed and Abandoned are how each visit "
                   "ended. Unique Calls counts each Call ID once. Visits that rang "
                   f"{res.min_ring_s} second(s) or less are not counted.")
        st.markdown("**Unique calls by final status** (how the whole call ended)")
        st.dataframe(res.final, use_container_width=True, hide_index=True)
    with tabs[1]:
        st.dataframe(res.routing, use_container_width=True, hide_index=True)
        st.caption("Counts queue visits (they add up to the Visits column). A manual transfer is never counted as "
                   "automatic overflow.")
    with tabs[2]:
        detail = res.detail
        if detail.empty:
            st.info("No queue visits inside office hours.")
            first_day = last_day = res.period.start_time.date()
        else:
            first_day, last_day = detail["Time"].min().date(), detail["Time"].max().date()
        queue_labels = {n: queue_label(n, nm) for n, nm in zip(res.queues["Queue Number"], res.queues["Queue Name"])}
        f1, f2, f3, f4 = st.columns(4)
        day_range = f1.date_input("Date entered queue", value=(first_day, last_day), min_value=first_day, max_value=last_day,
                                  key="queue_f_dates")
        queue_pick = f2.multiselect("Queue", list(res.queues["Queue Number"]), format_func=queue_labels.get,
                                    key="queue_f_queue", help="Leave empty for all queues.")
        outcome_pick = f3.multiselect("Outcome (visit status)", sorted(detail["Status"].unique()),
                                      key="queue_f_outcome", help="Leave empty for all outcomes.")
        type_pick = f4.multiselect("Routing type", [t for t in ROUTING_TYPES if t in set(detail["Routing Type"])],
                                   key="queue_f_type", help="Leave empty for all routing types.")
        # While only one end of the date range has been picked the date filter is not applied.
        start, end = day_range if isinstance(day_range, tuple) and len(day_range) == 2 else (None, None)
        view = filter_detail(detail, start, end, queue_pick or None, outcome_pick or None, type_pick or None)
        st.caption(f"Showing {len(view):,} of {len(detail):,} queue visits ({view['_ckey'].nunique():,} calls). "
                   "Every row carries its Call ID and source file so it can be traced back to the CDR.")
        st.dataframe(view[QUEUE_DETAIL_COLS], use_container_width=True, hide_index=True)
    with tabs[3]:
        if res.exceptions.empty:
            st.success("No exceptions: every call and mapping row could be classified.")
        else:
            st.write("These calls and queue-mapping rows need a look. They are listed here (and in the Excel "
                     "file and PDF) rather than being dropped silently.")
            st.dataframe(res.exceptions, use_container_width=True, hide_index=True)


# ================================================================ extension performance detail
def render_ext_performance():
    department = st.selectbox("Department", departments, key="perf_department",
                              help="Populated from the department mapping file.") if departments else None
    can_run = period is not None and bool(files) and mapping is not None and department is not None
    generate = st.button("Generate Extension Performance Detail", type="primary", disabled=not can_run)

    if generate:
        st.session_state.pop("perf_state", None)
        bar = st.progress(0.0, text="Starting…")
        try:
            payload = [(f.name, f.getvalue()) for f in files]
            res = process_ext_performance(
                payload, mapping, names, department, dedupe=dedupe, period=period,
                dayfirst=date_order.startswith("Day"),
                progress=lambda frac, msg: bar.progress(min(frac, 1.0), text=msg),
            )
            bar.progress(0.93, text="Building Excel…")
            title = report_title or PERF_TITLE
            st.session_state["perf_state"] = dict(
                res=res,
                xlsx=build_perf_xlsx(res, title),
                xlsx_name=perf_filename(title, department, res.period),
                warnings=map_warnings + res.warnings,
            )
            bar.progress(1.0, text="Done")
        except ReportError as exc:
            bar.empty()
            st.error(str(exc))
        except Exception:  # noqa: BLE001
            bar.empty()
            st.error("Something unexpected went wrong while generating the report.")
            with st.expander("Technical details"):
                st.code(traceback.format_exc())

    state = st.session_state.get("perf_state")
    st.divider()
    d1, _ = st.columns([1, 3])
    d1.download_button("Download Excel (.xlsx)", data=state["xlsx"] if state else b"",
                       file_name=state["xlsx_name"] if state else "Extension_Performance_Detail.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       disabled=state is None, type="primary", key="dl_perf_xlsx")
    if not state:
        return

    res, audit = state["res"], state["res"].audit
    st.success(f"Extension performance detail ready for {res.department} - {res.period.strftime('%B %Y')}")
    for w in state["warnings"]:
        st.warning(w)
    for _, msg in audit.files_failed:
        st.warning(f"Skipped: {msg}")
    if audit.first_call is not None:
        st.caption(f"Data covers {audit.first_call:%d/%m/%Y %H:%M} - {audit.last_call:%d/%m/%Y %H:%M}")
    if res.left_out:
        st.caption(f"{res.left_out:,} call leg(s) on the department's extensions were left out because they are "
                   "internal calls between extensions.")
    c = st.columns(4)
    c[0].metric("Call legs", f"{len(res.detail):,}")
    c[1].metric("Consultants", f"{res.consultants:,}")
    c[2].metric("Inbound", f"{res.count('Inbound'):,}")
    c[3].metric("Outbound", f"{res.count('Outbound'):,}")
    preview = preview_frame(res)
    st.dataframe(preview, use_container_width=True, hide_index=True)
    if len(res.detail) > len(preview):
        st.caption(f"Showing the first {len(preview):,} of {len(res.detail):,} rows. The Excel file has every row.")


# ================================================================ run the selected report
if TRANSFER:
    render_transfer_report()
elif EXT_STATS:
    render_ext_stats()
elif QUEUE:
    render_queue_overflow()
elif EXT_PERF:
    render_ext_performance()
else:
    render_call_list()
