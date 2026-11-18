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
from transfer_report.ext_stats import process_ext_stats
from transfer_report.ext_stats_report import (DEFAULT_TITLE as EXT_TITLE, build_ext_stats_pdf,
                                              build_ext_stats_xlsx, ext_filename)
from transfer_report.config import DEFAULT_MAPPING_PATH
from transfer_report.errors import MappingError, ReportError
from transfer_report.exports import unmapped_to_xlsx
from transfer_report.mapping import create_template, list_departments, load_mapping
from transfer_report.pdf_report import build_pdf, prepare_logo, report_filename
from transfer_report.processor import MATCH_BOTH, MATCH_EITHER, process_files

DEFAULT_LOGO_PATH = Path(__file__).resolve().parent / "logo.png"  # shown on every report unless switched off

st.set_page_config(page_title="Yeastar Transfer Report", page_icon="📞", layout="wide")
st.title("📞 Yeastar Reports")
REPORT_TYPES = ["Transfer report", "Department call list", "Extension call stats"]
mode = st.radio("Report type", REPORT_TYPES, horizontal=True, key="report_type")
TRANSFER = mode == REPORT_TYPES[0]
EXT_STATS = mode == REPORT_TYPES[2]
st.caption({
    REPORT_TYPES[0]: "Upload Yeastar CDR exports, combine them, and generate an inter-department transfer report.",
    REPORT_TYPES[1]: "Upload Yeastar CDR exports and list every number a department dialled (outbound) or answered (inbound).",
    REPORT_TYPES[2]: "Upload Yeastar CDR exports and see each extension's calls and outcomes for a department.",
}[mode])

# ---------------------------------------------------------------- sidebar
mapping, map_warnings, departments = None, [], []
with st.sidebar:
    st.header("Settings")
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
    if EXT_STATS:
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

# ---------------------------------------------------------------- upload
files = st.file_uploader("Drag and drop Yeastar CDR files here", type=["xls", "xlsx"],
                         accept_multiple_files=True)
if files:
    st.write(f"**{len(files)}** file(s) selected: " + ", ".join(f.name for f in files))
else:
    st.info("Upload one or more .xls / .xlsx files to begin.")

# Report name and month: set before generating
prev_month = date.today().replace(day=1) - timedelta(days=1)
c1, c2, c3 = st.columns([3, 2, 1])
report_title = c1.text_input("Report name", value={REPORT_TYPES[0]: "Yeastar Transfer Report", REPORT_TYPES[1]: CALLS_TITLE, REPORT_TYPES[2]: EXT_TITLE}[mode],
                             key=f"title_{mode}",
                             help="Shown on the cover page and used in the file names.").strip()
month_names = list(calendar.month_name)[1:]
month_name = c2.selectbox("Report month", month_names, index=prev_month.month - 1)
year = int(c3.number_input("Year", min_value=2000, max_value=2100, value=prev_month.year, step=1))
period = pd.Period(year=year, month=month_names.index(month_name) + 1, freq="M")



# ================================================================ transfer report
def render_transfer_report():
    can_run = bool(files) and mapping is not None and bool(selected)
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
    can_run = bool(files) and mapping is not None and department is not None
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
    can_run = bool(files) and mapping is not None and department is not None
    generate = st.button("Generate Extension Stats", type="primary", disabled=not can_run)

    if generate:
        st.session_state.pop("ext_state", None)
        bar = st.progress(0.0, text="Starting…")
        try:
            payload = [(f.name, f.getvalue()) for f in files]
            res = process_ext_stats(
                payload, mapping, names, department, include_internal=include_internal, hide_idle=hide_idle,
                dedupe=dedupe, period=period, dayfirst=date_order.startswith("Day"),
                progress=lambda frac, msg: bar.progress(min(frac, 1.0), text=msg),
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
    tabs = dict(zip(res.tables, st.tabs(list(res.tables))))
    for name, df in res.tables.items():
        with tabs[name]:
            st.dataframe(df, use_container_width=True, hide_index=True)


# ================================================================ run the selected report
if TRANSFER:
    render_transfer_report()
elif EXT_STATS:
    render_ext_stats()
else:
    render_call_list()
