"""Tests for the Queue overflow report, using small synthetic (invented) CDR workbooks.

Run from the project root:  python -m pytest tests/test_queue_overflow.py   (or: python tests/test_queue_overflow.py)

The synthetic calls cover each routing rule and each exception type; no real call data is used.
"""
from __future__ import annotations

import io
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from transfer_report.errors import MappingError, ReportError  # noqa: E402
from transfer_report.queue_mapping import load_queue_mapping  # noqa: E402
from transfer_report.queue_overflow import (AUTO, DIRECT, MANUAL, UNDETERMINED, filter_detail,  # noqa: E402
                                            process_queue_overflow)
from transfer_report.queue_report import build_queue_pdf, build_queue_xlsx  # noqa: E402

SUMMARY_HEADER = ["ID", "Call Leg ID", "Time", "Communication Type", "Call From", "Call To", "Second Participant",
                  "Last Participant", "Last Status", "Queue"]
LEG_HEADER = ["ID", "Call Leg ID", "Call Leg Time", "Communication Type", "Call From", "Call To", "Status"]
EVENT_HEADER = ["ID", "Call Leg ID", "Event ID", "Event Type", "Event", "Event Start Time"]
Q1, Q5 = "Queue Queue 1<801>", "Queue Queue 5<810>"
FMT = "%d/%m/%Y %H:%M:%S"


def entry(to, rd):
    return {"skill_based": 0, "ring_strategy_queue": "linear", "to": to, "ring_duration": rd}


def route(to):
    return {"via_inbound_route": "HQ", "destination_type": "queue", "to": to}


def make_call(cid, start, status="ANSWERED", ctype="Inbound", queues="", legs=(), events=(), frm="0821110000"):
    """legs: (offset_s, 'Call To', status); events: (offset_s, type, dict). Returns the call's rows."""
    t = lambda off: (start + timedelta(seconds=off)).strftime(FMT)  # noqa: E731
    rows = [SUMMARY_HEADER, [cid, None, t(0), ctype, frm, "x", None, None, status, queues]]
    rows.append(LEG_HEADER)
    rows += [[cid, f"L{i}", t(o), ctype, frm, to, st] for i, (o, to, st) in enumerate(legs)]
    rows.append(EVENT_HEADER)
    rows += [[cid, "L0", f"{i:03d}", et, json.dumps(js), t(o)] for i, (o, et, js) in enumerate(events)]
    return rows


def workbook(*calls) -> bytes:
    wb = Workbook()
    ws = wb.active
    for rows in calls:
        for r in rows:
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


D = datetime(2026, 10, 18, 9, 0, 0)
AG1, AG2 = "Agent One<101>", "Agent Two<102>"
MAP = pd.DataFrame({"Queue Number": ["801", "810"], "Queue Name": ["Queue 1 - Reception", "Queue 5 - Customer Care"]})
ISSUES = pd.DataFrame([{"Type": "Queue mapping: duplicate queue number", "Queue Number": "810", "Detail": "dup"}])


def scenario_calls():
    return [
        # A: direct to 801, answered by agent 101
        make_call("A", D, queues="Queue 1<801>", legs=[(0, Q1, "ANSWERED"), (20, AG1, "ANSWERED")],
                  events=[(0, "basic", route("801")), (0, "queue", entry("801", 20)),
                          (20, "queue", {"to": "101", "round_id": 1})]),
        # B: 801 times out after 30s, overflows automatically into 810, answered by agent 102
        make_call("B", D + timedelta(minutes=10), queues="Queue 1<801>,Queue 5<810>",
                  legs=[(0, Q1, "NO ANSWER"), (30, Q5, "ANSWERED"), (42, AG2, "ANSWERED")],
                  events=[(0, "basic", route("801")), (0, "queue", entry("801", 30)),
                          (0, "queue", {"to": "101", "ring_duration": 30, "round_id": 1}),
                          (30, "queue", entry("810", 12)), (42, "queue", {"to": "102", "round_id": 1})]),
        # C: answered in 801, agent blind-transfers it into 810 -> manual
        make_call("C", D + timedelta(minutes=20), status="NO ANSWER", queues="Queue 1<801>,Queue 5<810>",
                  legs=[(0, Q1, "ANSWERED"), (20, AG1, "ANSWERED"), (60, Q5, "NO ANSWER")],
                  events=[(0, "basic", route("801")), (0, "queue", entry("801", 20)),
                          (20, "queue", {"to": "101", "round_id": 1}),
                          (60, "operate", {"from": "101", "blind_to": "150"}), (60, "queue", entry("810", 15))]),
        # D: timings look like an overflow, but an agent transfer was logged in between -> manual wins
        make_call("D", D + timedelta(minutes=30), status="NO ANSWER", queues="Queue 1<801>,Queue 5<810>",
                  legs=[(0, Q1, "NO ANSWER"), (30, Q5, "NO ANSWER")],
                  events=[(0, "basic", route("801")), (0, "queue", entry("801", 30)),
                          (15, "operate", {"from": "101", "attended_to": "150"}), (30, "queue", entry("810", 10))]),
        # E: first visit to 810 with no inbound route / IVR event and no transfer -> undetermined
        make_call("E", D + timedelta(minutes=40), status="NO ANSWER", queues="Queue 5<810>",
                  legs=[(0, Q5, "NO ANSWER")], events=[(0, "queue", entry("810", 10))]),
        # F: 801 -> 810 -> 801: three visits, two to 801, one call
        make_call("F", D + timedelta(minutes=50), status="ABANDONED", queues="Queue 1<801>,Queue 5<810>,Queue 1<801>",
                  legs=[(0, Q1, "NO ANSWER"), (10, Q5, "NO ANSWER"), (20, Q1, "NO ANSWER")],
                  events=[(0, "basic", route("801")), (0, "queue", entry("801", 10)), (10, "queue", entry("810", 10)),
                          (20, "queue", entry("801", 10))]),
        # G: blank final status; H: unrecognised final status
        make_call("G", D + timedelta(minutes=60), status=None, queues="Queue 1<801>", legs=[(0, Q1, "NO ANSWER")],
                  events=[(0, "basic", route("801")), (0, "queue", entry("801", 10))]),
        make_call("H", D + timedelta(minutes=70), status="00:00:05", queues="Queue 1<801>",
                  legs=[(0, Q1, "NO ANSWER")], events=[(0, "basic", route("801")), (0, "queue", entry("801", 10))]),
        # J: queue 999 is not in the mapping
        make_call("J", D + timedelta(minutes=80), queues="Queue 9<999>", legs=[(0, "Queue Queue 9<999>", "NO ANSWER")],
                  events=[(0, "basic", route("999")), (0, "queue", entry("999", 10))]),
        # K: an internal call that went through queue 801 (not inbound)
        make_call("K", D + timedelta(minutes=90), ctype="Internal", queues="Queue 1<801>",
                  legs=[(0, Q1, "NO ANSWER")], events=[(0, "queue", entry("801", 10))]),
        # L: inbound call in September (outside the report month)
        make_call("L", datetime(2026, 9, 30, 10, 0, 0), queues="Queue 1<801>", legs=[(0, Q1, "NO ANSWER")],
                  events=[(0, "basic", route("801")), (0, "queue", entry("801", 10))]),
        # M: summary lists queue 801 but the export has no events or legs for the call
        make_call("M", D + timedelta(minutes=100), status="NO ANSWER", queues="Queue 1<801>"),
    ]


def run(files, dedupe=True, selected=("801", "810")):
    return process_queue_overflow(files, MAP, list(selected), ISSUES, dedupe=dedupe, period=pd.Period("2026-10"))


def row(res, cid, queue):
    d = res.detail
    return d[(d["Call ID"] == cid) & (d["Queue Number"] == queue)]


def test_routing_rules():
    # file 1 has a repeated record (N) and starts later than file 2, so file 1's copy of 'A' is the one kept
    n_dup = make_call("H", D + timedelta(minutes=70), status="00:00:05", queues="Queue 1<801>")
    res = run([("one.xlsx", workbook(*scenario_calls(), n_dup))])
    r = lambda cid, q: row(res, cid, q).iloc[0]  # noqa: E731
    assert r("A", "801")["Routing Type"] == DIRECT
    assert r("B", "801")["Routing Type"] == DIRECT and r("B", "810")["Routing Type"] == AUTO
    assert r("B", "810")["Previous Queue"].endswith("<801>")
    assert r("C", "810")["Routing Type"] == MANUAL
    assert r("D", "810")["Routing Type"] == MANUAL, "a manual transfer must never be counted as automatic overflow"
    assert r("E", "810")["Routing Type"] == UNDETERMINED
    assert r("M", "801")["Routing Type"] == UNDETERMINED
    assert list(row(res, "F", "801")["Visit #"]) == [1, 2] and list(row(res, "F", "810")["Visit #"]) == [1]
    assert r("B", "801")["Overflowed To"].endswith("<810>")


def test_answered_by_queue_agent():
    res = run([("one.xlsx", workbook(*scenario_calls()))])
    r = lambda cid, q: row(res, cid, q).iloc[0]  # noqa: E731
    assert r("A", "801")["Answered by Queue Agent"] == "Yes" and "Agent One" in r("A", "801")["Agent"]
    assert r("B", "801")["Answered by Queue Agent"] == "No"
    assert r("B", "810")["Answered by Queue Agent"] == "Yes" and "Agent Two" in r("B", "810")["Agent"]
    assert r("M", "801")["Answered by Queue Agent"] == "Unknown"


def test_summary_counts_queue_visits_like_the_pbx():
    res = run([("one.xlsx", workbook(*scenario_calls()))])
    s = res.summary.set_index("Queue Number")
    # 801: A B C D F(x2) G H K M = 10 visits from 9 calls (F entered twice; K is an internal call and is counted)
    assert s.loc["801", "Total Calls"] == 10 and s.loc["801", "Unique Calls"] == 9
    # each visit's own status: A, C answered; B D F F G H K missed; M has no queue leg (Other)
    assert (s.loc["801", "Answered"], s.loc["801", "Missed"], s.loc["801", "Abandoned"], s.loc["801", "Other"]) \
        == (2, 7, 0, 1)
    assert s.loc["801", "Answered %"] == 20.0 and s.loc["801", "Missed %"] == 70.0
    assert s.loc["801", "Answered by Queue Agent"] == 2  # A and C
    assert s.loc["801", "Overflowed Out"] == 2  # B and F (D's next hop was manual)
    assert s.loc["810", "Total Calls"] == 5 and s.loc["810", "Unique Calls"] == 5
    assert (s.loc["810", "Answered"], s.loc["810", "Missed"]) == (1, 4) and s.loc["810", "Answered by Queue Agent"] == 1
    assert s.loc["", "Total Calls"] == 15 and s.loc["", "Unique Calls"] == 10  # each call once across both queues
    assert s.loc["", ["Answered", "Missed", "Abandoned", "Other"]].sum() == 15
    rt = res.routing.set_index("Queue Number")
    assert rt.loc["801", "Visits"] == 10 and rt.loc["801", "Unique Calls"] == 9
    assert (rt.loc["801", [DIRECT, AUTO, MANUAL, UNDETERMINED]].sum()) == rt.loc["801", "Visits"]
    assert rt.loc["810", AUTO] == 2 and rt.loc["810", MANUAL] == 2 and rt.loc["810", UNDETERMINED] == 1
    # unique calls by final status (how the whole call ended) is still available
    f = res.final.set_index("Queue Number").loc["801"]
    assert (f["Unique Calls"], f["Answered"], f["No Answer"], f["Abandoned"], f["Blank"], f["Other"]) == (9, 3, 3, 1, 1, 1)
    assert res.stats["Total calls (queue visits)"] == 15 and res.stats["Unique calls"] == 10
    assert res.audit.outside_period == 1 and res.audit.non_inbound == 1 and res.audit.unmapped_queue_calls == 1


def test_short_visits_are_not_counted_but_listed():
    short = make_call("S", D + timedelta(minutes=110), status="ABANDONED", queues="Queue 1<801>",
                      legs=[(0, Q1, "ABANDONED")], events=[(0, "basic", route("801")), (0, "queue", entry("801", 1))])
    files = [("one.xlsx", workbook(*scenario_calls(), short))]
    res = run(files)
    assert "S" not in set(res.detail["Call ID"]) and res.audit.short_visits == 1
    ex = res.exceptions[res.exceptions["Type"].str.startswith("Queue visit not counted")]
    assert list(ex["Call ID"]) == ["S"]
    res0 = process_queue_overflow(files, MAP, ["801", "810"], ISSUES, period=pd.Period("2026-10"), min_ring_s=0)
    assert "S" in set(res0.detail["Call ID"]) and res0.audit.short_visits == 0  # a 1s visit counts when the limit is 0


def test_detail_follows_the_pbx_layout():
    header = LEG_HEADER + ["Call Duration", "Ring Duration", "Pure Talk Duration", "Hold Duration"]
    start = D + timedelta(minutes=120)
    t = lambda off: (start + timedelta(seconds=off)).strftime(FMT)  # noqa: E731
    answered = [SUMMARY_HEADER, ["P", None, t(0), "Inbound", "0821110000", "x", None, None, "ANSWERED", "Queue 1<801>"],
                header, ["P", "L0", t(0), "Inbound", "0821110000", Q1, "ANSWERED", "00:02:10", "00:00:10",
                         "00:01:55", "00:00:05"],
                ["P", "L1", t(10), "Inbound", "0821110000", AG1, "ANSWERED", "00:02:00", "00:00:04", "00:01:55",
                 "00:00:05"],
                EVENT_HEADER, ["P", "L0", "000", "basic", json.dumps(route("801")), t(0)],
                ["P", "L0", "001", "queue", json.dumps(entry("801", 10)), t(0)],
                ["P", "L0", "002", "queue", json.dumps({"to": "101", "round_id": 1}), t(10)]]
    missed = make_call("Q", start + timedelta(minutes=5), status="NO ANSWER", queues="Queue 1<801>",
                       legs=[(0, Q1, "NO ANSWER")], events=[(0, "basic", route("801")), (0, "queue", entry("801", 30))])
    res = run([("one.xlsx", workbook(answered, missed))])
    d = res.detail.set_index("Call ID")
    # Ring = the visit's ring time, Talk = call duration less ring time (includes hold), Hold = hold time
    assert (d.loc["P", "Status"], d.loc["P", "Ring Duration"], d.loc["P", "Talk Duration"], d.loc["P", "Hold Duration"]) \
        == ("Answered", "00:00:10", "00:02:00", "00:00:05")
    assert d.loc["P", "Agent"] == "Agent One<101>" and d.loc["Q", "Agent"] == "NONE"
    assert d.loc["Q", "Status"] == "Missed" and d.loc["Q", "Ring Duration"] == "00:00:30"
    assert list(res.detail.columns[:8]) == ["Call ID", "Time", "Caller", "Agent", "Status", "Ring Duration",
                                            "Talk Duration", "Hold Duration"]
    assert d.loc["P", "Time"] == start  # the time the call entered the queue


def test_exceptions_are_reported_not_dropped():
    res = run([("one.xlsx", workbook(*scenario_calls(), make_call("H", D, status="x")))])
    types = set(res.exceptions["Type"])
    for expected in ("Duplicate call record in file", "Final status is blank", "Final status not recognised",
                     "Routing type cannot be determined", "Queue routing events missing",
                     "Queue number not in the queue mapping",
                     "No queue call leg found for this visit", "Queue mapping: duplicate queue number"):
        assert expected in types, expected
    ex = res.exceptions
    assert set(ex.loc[ex["Type"] == "Routing type cannot be determined", "Call ID"]) == {"E", "K", "M"}
    assert list(ex.loc[ex["Type"] == "Queue number not in the queue mapping", "Queue Number"]) == ["999"]


def test_overlapping_files_count_a_call_once():
    early = workbook(*scenario_calls()[:3])  # A B C
    late = workbook(*scenario_calls()[:5])  # A B C D E  (A B C repeated)
    res = run([("early.xlsx", early), ("late.xlsx", late)])
    assert res.detail["_ckey"].nunique() == 5 and len(res.detail) == len(res.detail.drop_duplicates(["Call ID", "Queue Number", "Position"]))
    assert res.audit.overlap_calls == 3 and res.audit.duplicates_removed == 3
    dup = res.exceptions[res.exceptions["Type"] == "Duplicate call record removed"]
    assert set(dup["Call ID"]) == {"A", "B", "C"}
    assert any("overlap" in w.lower() or "more than one" in w for w in res.warnings)
    res_off = run([("early.xlsx", early), ("late.xlsx", late)], dedupe=False)
    assert res_off.detail["_ckey"].nunique() == 8 and res_off.audit.duplicates_removed == 0


def test_filters_and_selection():
    res = run([("one.xlsx", workbook(*scenario_calls()))])
    d = res.detail
    assert set(filter_detail(d, routing_types=[AUTO])["Call ID"]) == {"B", "F"}
    assert set(filter_detail(d, queues=["810"], outcomes=["Missed"])["Call ID"]) == {"C", "D", "E", "F"}
    day = D.date()
    assert len(filter_detail(d, start=day, end=day)) == len(d)
    assert filter_detail(d, start=day + timedelta(days=1)).empty
    only = run([("one.xlsx", workbook(*scenario_calls()))], selected=("810",))
    assert set(only.detail["Queue Number"]) == {"810"} and list(only.summary["Queue Number"]) == ["810", ""]
    try:
        run([("one.xlsx", workbook(*scenario_calls()))], selected=())
    except ReportError:
        pass
    else:
        raise AssertionError("selecting no queue should raise a friendly error")


def test_outputs_build():
    res = run([("one.xlsx", workbook(*scenario_calls()))])
    pdf, xlsx = build_queue_pdf(res, "Queue Overflow Report"), build_queue_xlsx(res, "Queue Overflow Report")
    assert pdf[:5] == b"%PDF-" and xlsx[:2] == b"PK"
    sheets = pd.read_excel(io.BytesIO(xlsx), sheet_name=None, dtype=str)
    assert list(sheets) == ["Summary", "By queue", "Routing type", "Final status", "Detail", "Exceptions",
                           "Definitions"]
    assert {"Call ID", "Time", "Status", "Routing Type", "Routing Note",
            "Answered by Queue Agent"} <= set(sheets["Detail"].columns)
    assert "Automatic overflow" in set(sheets["Definitions"]["Term"])


def test_queue_mapping_loader():
    wb = Workbook()
    ws = wb.active
    ws.append(["Que Number", "Department"])
    for r in (["801", "Reception"], ["801", "Reception"], ["810", "Customer Care"], ["810", "Other name"],
              [None, "No number"], ["802", None], ["abc", "Letters"], [None, None]):
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    frame, warnings, issues = load_queue_mapping(io.BytesIO(buf.getvalue()))
    assert list(frame["Queue Number"]) == ["801", "810", "802", "abc"]
    assert frame.set_index("Queue Number").loc["810", "Queue Name"] == "Customer Care"
    kinds = set(issues["Type"])
    assert {"Queue mapping: duplicate queue number", "Queue mapping: queue number has different names",
            "Queue mapping: blank queue number", "Queue mapping: blank queue name",
            "Queue mapping: queue number is not numeric"} <= kinds
    assert warnings
    bad = Workbook()
    bad.active.append(["Foo", "Bar"])
    out = io.BytesIO()
    bad.save(out)
    try:
        load_queue_mapping(io.BytesIO(out.getvalue()))
    except MappingError:
        pass
    else:
        raise AssertionError("a mapping without queue columns should raise MappingError")


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("ok", name)
