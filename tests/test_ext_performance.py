"""Tests for the Extension Performance Detail report, using small synthetic CDR exports (no real call data)."""
import io

import openpyxl
import pandas as pd
import pytest

from transfer_report.calls import parse_layers, read_calls_file
from transfer_report.errors import EmptyUploadError, InvalidFileError, ReportError
from transfer_report.ext_performance import (DETAIL_COLS, preview_frame, process_ext_performance,
                                             read_perf_file)
from transfer_report.ext_performance_report import (build_perf_xlsx, period_label, perf_filename)
from transfer_report.ext_stats import process_ext_stats

SUMMARY_HDR = ["ID", "Call Leg ID", "Time", "Communication Type", "Call From", "Call To", "Second Participant",
               "Last Participant", "Last Status", "Queue", "Ring Group", "Ivr", "Call Flow", "Total Duration",
               "Routing Duration", "Handling Duration"]
LEG_HDR = ["ID", "Call Leg ID", "Call Leg Time", "Communication Type", "Call From", "Call To", "Status",
           "Call Duration", "Ring Duration", "Pure Talk Duration", "Hold Duration"]

MAPPING = pd.DataFrame({"Extension": ["010", "074", "200"],
                        "Department": ["Claims Call Centre", "Claims Call Centre", "Finance"]})
NAMES = pd.DataFrame({"Extension": ["010", "074"], "Name": ["Angelina Ramoroa", "Itumeleng Lehasa"]})
PERIOD = pd.Period("2026-09", "M")


def leg(call_id, leg_id, time, ctype, frm, to, status, dur="00:00:30", ring="00:00:10", talk="00:00:20"):
    return [call_id, leg_id, time, ctype, frm, to, status, dur, ring, talk, "00:00:00"]


def cdr(*calls, name="export.xlsx"):
    """calls: (summary_type, call_id, start_time, [leg rows]) -> (name, bytes) of a layered Yeastar export."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(SUMMARY_HDR)
    for ctype, call_id, start, legs in calls:
        ws.append([call_id, None, start, ctype, legs[0][4], legs[0][5], None, None, "ANSWERED", None, None, None,
                   None, "00:01:00", "00:00:00", "00:00:50"])
        ws.append(LEG_HDR)
        for row in legs:
            ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return name, buf.getvalue()


def inbound_call(call_id="A1", start="01/09/2026 08:00:00"):
    return ("Inbound", call_id, start, [
        leg(call_id, "11001", start, "Inbound", "+27111<0111>", "IVR Main IVR<6200>", "ANSWERED"),
        leg(call_id, "12002", start, "Inbound", "+27111<0111>", "Angelina Ramoroa<010>", "NO ANSWER",
            dur="00:00:30", ring="00:00:30", talk="00:00:00"),
        leg(call_id, "13003", "01/09/2026 08:00:31", "Inbound", "+27111<0111>", "Voicemail Angelina Ramoroa<010>",
            "VOICEMAIL", dur="00:00:26", ring="00:00:00", talk="00:00:26"),
        leg(call_id, "14004", "01/09/2026 08:00:05", "Inbound", "+27111<0111>", "Sales Queue<8000>", "ANSWERED"),
    ])


def outbound_call(call_id="B1", start="01/09/2026 09:00:00"):
    return ("Outbound", call_id, start, [
        leg(call_id, "21001", start, "Outbound", "Itumeleng Lehasa<074>", "0829999999", "ANSWERED",
            dur="00:02:05", ring="00:00:05", talk="00:02:00")])


def run(files, **kw):
    kw.setdefault("period", PERIOD)
    return process_ext_performance(files, MAPPING, NAMES, "Claims Call Centre", **kw)


# ------------------------------------------------------------------ selection rules
def test_inbound_outbound_voicemail_legs_and_names():
    res = run([cdr(inbound_call(), outbound_call())])
    d = res.detail
    assert list(d.columns) == DETAIL_COLS
    assert len(d) == 3  # no-answer leg, voicemail leg, outbound leg; IVR and queue legs are ignored
    assert d["Time"].is_monotonic_increasing
    ib = d[d["Communication Type"] == "Inbound"]
    assert set(ib["Consultants"]) == {"Angelina Ramoroa"}
    assert set(ib["Status"]) == {"NO ANSWER", "VOICEMAIL"}
    assert "Voicemail Angelina Ramoroa<010>" in set(ib["Call To"])  # raw CDR text is kept
    ob = d[d["Communication Type"] == "Outbound"].iloc[0]
    assert (ob["Consultants"], ob["Call From"], ob["Call To"], ob["Status"]) == (
        "Itumeleng Lehasa", "Itumeleng Lehasa<074>", "0829999999", "ANSWERED")
    assert ob["Call Duration"] == pd.Timedelta(minutes=2, seconds=5)
    assert ob["Ring Duration"] == pd.Timedelta(seconds=5)
    assert ob["Talk Duration"] == pd.Timedelta(minutes=2)


def test_other_department_extensions_are_ignored():
    other = ("Outbound", "C1", "01/09/2026 10:00:00", [
        leg("C1", "1", "01/09/2026 10:00:00", "Outbound", "Finance Person<200>", "0821111111", "ANSWERED")])
    res = run([cdr(outbound_call(), other)])
    assert len(res.detail) == 1 and set(res.detail["Consultants"]) == {"Itumeleng Lehasa"}


def test_internal_calls_are_left_out_and_counted():
    internal = ("Internal", "I1", "01/09/2026 11:00:00", [
        leg("I1", "1", "01/09/2026 11:00:00", "Internal", "Angelina Ramoroa<010>", "Itumeleng Lehasa<074>",
            "ANSWERED")])
    res = run([cdr(outbound_call(), internal)])
    assert len(res.detail) == 1
    assert res.left_out == 1


def test_leg_type_is_used_not_the_call_type():
    # an outbound call whose transferred-in leg is inbound for the extension: the row says Inbound
    call = ("Outbound", "T1", "01/09/2026 12:00:00", [
        leg("T1", "1", "01/09/2026 12:00:01", "Inbound", "+27222<0222>", "Angelina Ramoroa<010>", "ANSWERED")])
    res = run([cdr(call)])
    assert list(res.detail["Communication Type"]) == ["Inbound"]
    assert list(res.detail["Consultants"]) == ["Angelina Ramoroa"]


def test_extension_on_the_wrong_side_of_a_leg_is_not_listed():
    odd = ("Inbound", "W1", "01/09/2026 13:00:00", [
        leg("W1", "1", "01/09/2026 13:00:00", "Inbound", "Angelina Ramoroa<010>", "+27333<0333>", "ANSWERED")])
    res = run([cdr(outbound_call(), odd)])
    assert len(res.detail) == 1 and res.left_out == 1


# ------------------------------------------------------------------ months, duplicates, ordering
def test_report_covers_the_selected_month_only():
    august = outbound_call("OLD", "31/08/2026 17:00:00")
    august[3][0][2] = "31/08/2026 17:00:00"
    res = run([cdr(outbound_call(), august)])
    assert len(res.detail) == 1
    assert res.audit.outside_period == 1
    assert any("outside September 2026" in w for w in res.warnings)


def test_no_calls_in_the_month_gives_a_clear_error():
    with pytest.raises(ReportError, match="Check the report month"):
        run([cdr(outbound_call())], period=pd.Period("2026-07", "M"))


def test_overlapping_files_count_each_leg_once_unless_dedupe_is_off():
    a, b = cdr(inbound_call(), name="a.xlsx"), cdr(inbound_call(), name="b.xlsx")
    assert len(run([a, b]).detail) == 2
    assert len(run([a, b], dedupe=False).detail) == 4
    assert any("more than one of the uploaded files" in w for w in run([a, b]).warnings)


def test_rows_are_in_time_order_across_files():
    res = run([cdr(outbound_call("L", "20/09/2026 09:00:00")), cdr(outbound_call("E", "02/09/2026 09:00:00"))])
    assert list(res.detail["ID"]) == ["E", "L"]


def test_day_first_and_month_first_dates():
    call = outbound_call("D", "03/09/2026 09:00:00")
    assert len(run([cdr(call)], dayfirst=True).detail) == 1
    with pytest.raises(ReportError):  # read as month first this is 9 March: outside September
        run([cdr(call)], dayfirst=False)


# ------------------------------------------------------------------ names
def test_name_falls_back_to_the_cdr_then_the_extension_number():
    stranger = ("Outbound", "N1", "01/09/2026 14:00:00", [
        leg("N1", "1", "01/09/2026 14:00:00", "Outbound", "Pat Smith<074>", "0821", "ANSWERED"),
        leg("N1", "2", "01/09/2026 14:01:00", "Outbound", "<010>", "0822", "ANSWERED")])
    res = process_ext_performance([cdr(stranger)], MAPPING, None, "Claims Call Centre", period=PERIOD)
    assert set(res.detail["Consultants"]) == {"Pat Smith", "010"}
    warning = " ".join(res.warnings)
    assert "074 (shown as Pat Smith)" in warning and "010 (shown as 010)" in warning


def test_names_file_wins_over_the_cdr_name():
    call = ("Outbound", "N2", "01/09/2026 14:00:00", [
        leg("N2", "1", "01/09/2026 14:00:00", "Outbound", "Different Name<074>", "0821", "ANSWERED")])
    assert set(run([cdr(call)]).detail["Consultants"]) == {"Itumeleng Lehasa"}


# ------------------------------------------------------------------ errors
def test_input_errors():
    with pytest.raises(EmptyUploadError):
        run([])
    with pytest.raises(ReportError, match="No extensions are mapped"):
        process_ext_performance([cdr(outbound_call())], MAPPING, NAMES, "Nowhere", period=PERIOD)
    with pytest.raises(InvalidFileError):
        run([("junk.xlsx", b"not an excel file")])
    with pytest.raises(ReportError, match="No inbound or outbound calls"):
        process_ext_performance([cdr(outbound_call())], MAPPING, NAMES, "Finance", period=PERIOD)


def test_a_bad_file_is_skipped_when_another_one_works():
    res = run([("junk.xlsx", b"not an excel file"), cdr(outbound_call())])
    assert len(res.detail) == 1 and [n for n, _ in res.audit.files_failed] == ["junk.xlsx"]


# ------------------------------------------------------------------ Excel output
def test_excel_layout_matches_the_example_report():
    res = run([cdr(inbound_call(), outbound_call())])
    wb = openpyxl.load_workbook(io.BytesIO(build_perf_xlsx(res)))
    ws = wb.active
    assert ws.title == "Claims Call Centre"
    assert [c.value for c in ws[1]] == DETAIL_COLS
    assert ws.freeze_panes == "A2" and ws.auto_filter.ref == "A1:J4"
    head = ws["A1"]
    assert head.font.b and head.font.name == "Arial" and head.fill.fgColor.rgb.endswith("1F3864")
    first = [c for c in ws[2]]
    assert first[0].value == "A1" and first[0].number_format == "@"
    assert first[1].number_format == "dd/mm/yyyy hh:mm:ss"
    assert [c.number_format for c in first[5:8]] == ["[h]:mm:ss"] * 3
    assert all(c.font.name == "Arial" and c.font.sz == 10 for c in first)
    assert ws["F4"].value == pd.Timedelta(minutes=2, seconds=5).to_pytimedelta()
    assert ws.column_dimensions["A"].width == 24 and ws.column_dimensions["J"].width == 13


def test_ids_stay_text_even_when_they_look_like_numbers():
    call = ("Outbound", "20260901090000123", "01/09/2026 09:00:00", [
        leg("20260901090000123", "1", "01/09/2026 09:00:00", "Outbound", "Itumeleng Lehasa<074>", "0821",
            "ANSWERED")])
    ws = openpyxl.load_workbook(io.BytesIO(build_perf_xlsx(run([cdr(call)])))).active
    assert ws["A2"].value == "20260901090000123" and ws["A2"].data_type == "s"


def test_filename_and_period_label():
    assert period_label(pd.Period("2026-09", "M")) == "1-30 September 2026"
    assert period_label(pd.Period("2028-02", "M")) == "1-29 February 2028"
    assert (perf_filename("Inbound and Outbound Call Report", "Claims Call Centre", PERIOD)
            == "Claims Call Centre - Inbound and Outbound Call Report - 1-30 September 2026.xlsx")
    name = perf_filename("Report: A/B", "Durban/Durban Brokers Division", PERIOD)
    assert not set('\\/:*?"<>|') & set(name)


def test_sheet_name_is_valid_for_awkward_department_names():
    mapping = pd.DataFrame({"Extension": ["074"], "Department": ["Facilities&Emerald Life/HQ [Long name over 31]"]})
    res = process_ext_performance([cdr(outbound_call())], mapping, NAMES, mapping["Department"][0], period=PERIOD)
    title = openpyxl.load_workbook(io.BytesIO(build_perf_xlsx(res))).active.title
    assert len(title) <= 31 and not set("[]:*?/\\") & set(title)


def test_preview_frame_is_readable_text():
    p = preview_frame(run([cdr(outbound_call())]))
    assert p.loc[0, "Time"] == "01/09/2026 09:00:00" and p.loc[0, "Call Duration"] == "00:02:05"


# ------------------------------------------------------------------ the shared parser change
def test_parse_layers_now_also_returns_leg_durations_and_type():
    _, legs, _ = parse_layers(*cdr(outbound_call()))
    row = legs.iloc[0]
    assert (row["dur"], row["ring"], row["leg_talk"], row["ltype"]) == ("00:02:05", "00:00:05", "00:02:00", "Outbound")


def test_existing_reports_are_unchanged_on_the_same_file():
    file = cdr(inbound_call(), outbound_call())
    stats = process_ext_stats([file], MAPPING, NAMES, "Claims Call Centre", period=PERIOD)
    totals = stats.all_calls.set_index("Extension")["Total Calls"].to_dict()
    assert totals == {"010": 2, "074": 1}  # same legs as the new report lists
    assert len(read_perf_file(*file, {"010", "074"}).legs) == 3
    calls = read_calls_file(file[0], file[1], {"010", "074"}, {"010", "074"}, dayfirst=True)
    assert len(calls.outbound) == 1 and calls.inbound.empty  # nobody answered the inbound call
