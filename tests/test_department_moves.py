"""Tests for effective-dated department mapping (an extension moving department part-way through a month).

Scenario used throughout: extension 078 was in Customer Care until 21 Oct 2026 and in Complaints from 22 Oct 2026.
Synthetic data only - no real call records.
"""
import io

import openpyxl
import pandas as pd
import pytest

from transfer_report.calls import ALL_DEPARTMENTS, process_calls
from transfer_report.errors import MappingError, ReportError
from transfer_report.ext_performance import process_ext_performance
from transfer_report.ext_stats import process_ext_stats
from transfer_report.mapping import (department_at, department_extensions, has_dated_rows, load_mapping,
                                     moves_notice)
from transfer_report.processor import process_files

OCT = pd.Period("2026-10", "M")
NOV = pd.Period("2026-11", "M")
MOVE = pd.Timestamp("2026-10-22")


def mapping_frame(rows):
    """rows: (extension, department, effective_from or None)."""
    df = pd.DataFrame(rows, columns=["Extension", "Department", "Effective From"])
    df["Effective From"] = pd.to_datetime(df["Effective From"]).astype("datetime64[ns]")
    return df


DATED = mapping_frame([
    ("078", "Customer Care", None),
    ("078", "Complaints", "2026-10-22"),
    ("031", "Finance", None),
    ("100", "Customer Care", None),
])
UNDATED = mapping_frame([("078", "Complaints", None), ("031", "Finance", None), ("100", "Customer Care", None)])


def mapping_xlsx(rows, header=("Extension", "Department", "Effective From")):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(list(header))
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ------------------------------------------------------------------ loading the mapping file
def test_loader_reads_dated_rows_and_keeps_both_rows_for_a_moved_extension():
    m, warnings = load_mapping(mapping_xlsx([("078", "Customer Care", None), ("078", "Complaints", "2026-10-22"),
                                             ("031", "Finance", None)]))
    assert warnings == []
    assert len(m) == 3 and list(m.columns) == ["Extension", "Department", "Effective From"]
    assert m[m["Extension"] == "078"]["Effective From"].tolist()[1] == MOVE
    assert has_dated_rows(m)


def test_loader_accepts_excel_dates_and_day_first_text_dates():
    import datetime as dt
    m, _ = load_mapping(mapping_xlsx([("078", "Customer Care", None), ("078", "Complaints", dt.date(2026, 10, 22)),
                                      ("079", "Finance", None), ("079", "HR", "22/10/2026")]))
    assert m[m["Extension"] == "078"]["Effective From"].iloc[1] == MOVE
    assert m[m["Extension"] == "079"]["Effective From"].iloc[1] == MOVE


def test_loader_works_without_the_effective_from_column():
    m, warnings = load_mapping(mapping_xlsx([("078", "Complaints"), ("031", "Finance")], header=("Extension", "Department")))
    assert warnings == [] and not has_dated_rows(m)
    assert m["Effective From"].isna().all()


def test_loader_rejects_a_value_that_is_not_a_date():
    with pytest.raises(MappingError, match="not valid dates.*'soon' \\(extension 078\\)"):
        load_mapping(mapping_xlsx([("078", "Customer Care", None), ("078", "Complaints", "soon")]))


def test_loader_warns_about_two_rows_with_the_same_date_and_uses_the_first():
    m, warnings = load_mapping(mapping_xlsx([("078", "Customer Care", None), ("078", "Complaints", None),
                                             ("031", "Finance", None)]))
    assert len(warnings) == 1 and "same 'Effective From' date" in warnings[0]
    assert m[m["Extension"] == "078"]["Department"].tolist() == ["Customer Care"]


# ------------------------------------------------------------------ the lookup
def test_department_at_switches_at_midnight_on_the_effective_date():
    ext = pd.Series(["078"] * 4)
    when = pd.to_datetime(pd.Series(["2026-10-01 08:00:00", "2026-10-21 23:59:59", "2026-10-22 00:00:00",
                                     "2026-10-31 17:00:00"]))
    assert department_at(DATED, ext, when).tolist() == ["Customer Care", "Customer Care", "Complaints", "Complaints"]


def test_department_at_other_cases():
    ext = pd.Series(["078", "031", "999", None, "078"])
    when = pd.to_datetime(pd.Series(["2026-10-05 10:00:00"] * 4 + [None]))
    out = department_at(DATED, ext, when)
    assert out.iloc[0] == "Customer Care" and out.iloc[1] == "Finance"
    assert pd.isna(out.iloc[2]) and pd.isna(out.iloc[3])  # unmapped extension / no extension
    assert out.iloc[4] == "Complaints"  # no time: the extension's latest row
    assert department_at(DATED, ext).iloc[0] == "Complaints"  # no times at all: latest row


def test_extension_is_unmapped_before_its_first_dated_row():
    m = mapping_frame([("200", "Finance", "2026-10-15")])  # joined on 15 Oct
    ext = pd.Series(["200", "200"])
    when = pd.to_datetime(pd.Series(["2026-10-14 12:00:00", "2026-10-15 00:00:01"]))
    out = department_at(m, ext, when)
    assert pd.isna(out.iloc[0]) and out.iloc[1] == "Finance"


def test_department_at_without_dates_is_a_plain_lookup():
    out = department_at(UNDATED, pd.Series(["078", "031", "999"]), pd.to_datetime(pd.Series(["2026-10-05"] * 3)))
    assert out.iloc[0] == "Complaints" and out.iloc[1] == "Finance" and pd.isna(out.iloc[2])


def test_mapping_without_an_effective_from_column_still_works():
    plain = DATED[["Extension", "Department"]].drop_duplicates("Extension", keep="last")
    assert department_at(plain, pd.Series(["078"]), pd.to_datetime(pd.Series(["2026-10-05"]))).iloc[0] == "Complaints"
    assert department_extensions(plain, "Complaints", OCT) == {"078"}


def test_department_extensions_follow_the_report_month():
    assert "078" in department_extensions(DATED, "Customer Care", OCT)
    assert "078" in department_extensions(DATED, "Complaints", OCT)
    assert "078" not in department_extensions(DATED, "Customer Care", NOV)  # moved out before November
    assert "078" not in department_extensions(DATED, "Complaints", pd.Period("2026-09", "M"))  # not yet in
    assert "078" in department_extensions(DATED, "Customer Care")  # no month: any time


def test_moves_notice_only_for_moves_inside_the_month():
    note = moves_notice(DATED, OCT)
    assert len(note) == 1 and "ext 078 Customer Care to Complaints on 22/10/2026" in note[0]
    assert moves_notice(DATED, NOV) == []
    assert moves_notice(DATED, OCT, "Finance") == []  # not involved
    assert len(moves_notice(DATED, OCT, "Complaints")) == 1
    assert moves_notice(UNDATED, OCT) == []


# ------------------------------------------------------------------ synthetic exports
SUMMARY_HDR = ["ID", "Call Leg ID", "Time", "Communication Type", "Call From", "Call To", "Second Participant",
               "Last Participant", "Last Status", "Queue", "Ring Group", "Ivr", "Call Flow", "Total Duration",
               "Routing Duration", "Handling Duration"]
LEG_HDR = ["ID", "Call Leg ID", "Call Leg Time", "Communication Type", "Call From", "Call To", "Status",
           "Call Duration", "Ring Duration", "Pure Talk Duration", "Hold Duration"]


def leg(call_id, leg_id, time, ctype, frm, to, status="ANSWERED"):
    return [call_id, leg_id, time, ctype, frm, to, status, "00:01:00", "00:00:10", "00:00:50", "00:00:00"]


def cdr(*calls, name="export.xlsx"):
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


def out078(call_id, when):
    return ("Outbound", call_id, when, [leg(call_id, call_id + "1", when, "Outbound", "Agent 078<078>", "0821111111")])


def in078(call_id, when):
    return ("Inbound", call_id, when, [leg(call_id, call_id + "1", when, "Inbound", "+27111<0111>", "Agent 078<078>")])


CALLS = [
    out078("O1", "10/10/2026 09:00:00"), in078("I1", "21/10/2026 23:59:00"),   # Customer Care days
    in078("I2", "22/10/2026 00:00:30"), out078("O2", "25/10/2026 10:00:00"),   # Complaints days
]


# ------------------------------------------------------------------ Department call list
def run_calls(department, mapping=DATED):
    return process_calls([cdr(*CALLS)], mapping, department, period=OCT)


def test_call_list_splits_the_moved_extension_by_date():
    cc, cp = run_calls("Customer Care"), run_calls("Complaints")
    assert sorted(cc.outbound["Call ID"]) == ["O1"] and sorted(cc.inbound["Call ID"]) == ["I1"]
    assert sorted(cp.outbound["Call ID"]) == ["O2"] and sorted(cp.inbound["Call ID"]) == ["I2"]
    assert any("ext 078 Customer Care to Complaints on 22/10/2026" in w for w in cc.warnings)
    assert cc.unmapped.empty and cp.unmapped.empty


def test_call_list_all_departments_labels_each_call_with_its_department_on_that_date():
    res = run_calls(ALL_DEPARTMENTS)
    by_call = dict(zip(pd.concat([res.outbound["Call ID"], res.inbound["Call ID"]]),
                       pd.concat([res.outbound["Department"], res.inbound["Department"]])))
    assert by_call == {"O1": "Customer Care", "I1": "Customer Care", "I2": "Complaints", "O2": "Complaints"}


def test_call_list_without_dates_is_unchanged():
    res = run_calls("Complaints", UNDATED)
    assert len(res.outbound) == 2 and len(res.inbound) == 2
    assert not any("Department change" in w for w in res.warnings)
    with pytest.raises(ReportError, match="No outbound or connected inbound calls"):
        run_calls("Customer Care", UNDATED)  # 078 is Complaints throughout, so Customer Care (ext 100) has no calls


def test_call_list_does_not_include_an_extension_that_left_the_department_before_the_month():
    moved_in_september = mapping_frame([("078", "Customer Care", None), ("078", "Complaints", "2026-09-01"),
                                        ("100", "Customer Care", None)])
    calls = [CALLS[0], ("Outbound", "O9", "11/10/2026 11:00:00",
                        [leg("O9", "O91", "11/10/2026 11:00:00", "Outbound", "Agent 100<100>", "0822222222")])]
    res = process_calls([cdr(*calls)], moved_in_september, "Customer Care", period=OCT)
    assert list(res.outbound["Call ID"]) == ["O9"] and "078" not in res.extensions


# ------------------------------------------------------------------ Extension call stats
def stats_row(department, mapping=DATED, period=OCT):
    res = process_ext_stats([cdr(*CALLS)], mapping, None, department, include_internal=False, period=period)
    return res, res.all_calls.set_index("Extension")


def test_extension_stats_count_each_period_under_the_right_department():
    cc, cc_t = stats_row("Customer Care")
    cp, cp_t = stats_row("Complaints")
    assert cc_t.loc["078", "Total Calls"] == 2 and cp_t.loc["078", "Total Calls"] == 2
    assert any("Department change" in w for w in cc.warnings)


def test_extension_stats_do_not_list_an_extension_for_a_department_it_has_left():
    nov = [out078("N1", "03/11/2026 09:00:00"),
           ("Outbound", "N2", "04/11/2026 09:00:00",
            [leg("N2", "N21", "04/11/2026 09:00:00", "Outbound", "Agent 100<100>", "0822222222")])]
    cc = process_ext_stats([cdr(*nov)], DATED, None, "Customer Care", period=NOV)
    assert list(cc.all_calls["Extension"]) == ["100"]  # 078 is a Complaints extension in November
    cp = process_ext_stats([cdr(*nov)], DATED, None, "Complaints", period=NOV)
    assert list(cp.all_calls["Extension"]) == ["078"] and cp.all_calls["Total Calls"].iloc[0] == 1
    assert not any("Department change" in w for w in cc.warnings + cp.warnings)


def test_extension_stats_without_dates_are_unchanged():
    _, table = stats_row("Complaints", UNDATED)
    assert table.loc["078", "Total Calls"] == 4


# ------------------------------------------------------------------ Extension Performance Detail
def test_performance_detail_splits_the_moved_extension_by_date():
    cc = process_ext_performance([cdr(*CALLS)], DATED, None, "Customer Care", period=OCT)
    cp = process_ext_performance([cdr(*CALLS)], DATED, None, "Complaints", period=OCT)
    assert len(cc.detail) == 2 and len(cp.detail) == 2
    assert cc.detail["Time"].max() < MOVE <= cp.detail["Time"].min()
    assert len(process_ext_performance([cdr(*CALLS)], UNDATED, None, "Complaints", period=OCT).detail) == 4


# ------------------------------------------------------------------ Transfer report
def transfer_cdr(*transfers):
    """transfers: (call_id, time, from_ext, to_ext) - blind transfers answered by the target."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Call Leg ID", "Time", "Communication Type", "Call From", "Call To", "Last Status"])
    for call_id, when, frm, to in transfers:
        ws.append([call_id, None, when, "Inbound", "+27111<0111>", f"Agent<{frm}>", "ANSWERED"])
        ws.append(["ID", "Call Leg ID", "Call Leg Time", "Communication Type", "Call From", "Call To", "Status"])
        ws.append([call_id, "L2", when, "Inbound", "+27111<0111>", f"Target<{to}>", "ANSWERED"])
        ws.append([call_id, "L1", "E1", "operate", f'{{"from":"{frm}","blind_to":"{to}"}}', when, None])
    buf = io.BytesIO()
    wb.save(buf)
    return "transfers.xlsx", buf.getvalue()


TRANSFERS = [
    ("T1", "10/10/2026 09:00:00", "078", "031"),  # Customer Care -> Finance
    ("T2", "12/10/2026 09:00:00", "078", "100"),  # Customer Care -> Customer Care: same department, excluded
    ("T3", "25/10/2026 09:00:00", "078", "031"),  # Complaints -> Finance
    ("T4", "26/10/2026 09:00:00", "078", "100"),  # Complaints -> Customer Care
]


def test_transfer_report_attributes_each_side_by_the_transfer_date():
    res = process_files([transfer_cdr(*TRANSFERS)], DATED, period=OCT)
    pairs = {(r["From Department"], r["To Department"]): r["Transfers"] for _, r in res.detail.iterrows()}
    assert pairs == {("Customer Care", "Finance"): 1, ("Complaints", "Finance"): 1, ("Complaints", "Customer Care"): 1}
    assert res.audit.internal_removed == 1 and res.audit.unmapped_removed == 0
    assert any("ext 078 Customer Care to Complaints on 22/10/2026" in w for w in res.warnings)


def test_transfer_report_without_dates_is_unchanged():
    res = process_files([transfer_cdr(*TRANSFERS)], UNDATED, period=OCT)
    pairs = {(r["From Department"], r["To Department"]): r["Transfers"] for _, r in res.detail.iterrows()}
    assert pairs == {("Complaints", "Finance"): 2, ("Complaints", "Customer Care"): 2}
    assert not any("Department change" in w for w in res.warnings)
