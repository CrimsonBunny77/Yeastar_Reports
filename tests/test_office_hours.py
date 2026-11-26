"""Tests for office hours (after hours, weekends, excluded days and short days), using synthetic CDR data only."""
import io
from datetime import date, datetime, time, timedelta

import pandas as pd

from tests.test_ext_performance import MAPPING, NAMES, PERIOD, cdr, leg
from tests.test_queue_overflow import ISSUES, MAP, Q1, make_call, entry, route, scenario_calls, workbook
from transfer_report.ext_stats import process_ext_stats
from transfer_report.ext_stats_report import build_ext_stats_pdf, build_ext_stats_xlsx
from transfer_report.hours import (AFTER, CLOSED, EXCLUDED, OFFICE, SHORT, WEEKEND, OfficeHours, build_office_hours,
                                   days_for_year, easter_sunday, load_days, sa_public_holidays, save_days)
from transfer_report.queue_overflow import process_queue_overflow
from transfer_report.queue_report import build_queue_pdf, build_queue_xlsx


def hours(*rows):
    oh, warnings = build_office_hours(pd.DataFrame(rows, columns=["Date", "Type", "Closes At", "Reason"]))
    assert not warnings
    return oh


# ------------------------------------------------------------------ the rules
def test_boundaries_weekends_excluded_and_short_days():
    oh = hours(("2026-10-06", SHORT, "13:00", "Year end function"), ("2026-12-16", CLOSED, None, "Holiday"),
               ("2026-10-10", CLOSED, None, "Saturday holiday"))
    t = pd.Series(pd.to_datetime([
        "2026-10-05 07:59:59", "2026-10-05 08:00:00", "2026-10-05 16:29:59", "2026-10-05 16:30:00",  # Monday
        "2026-10-06 12:59:59", "2026-10-06 13:00:00",  # short day, closes at 13:00
        "2026-10-11 10:00:00",  # Sunday
        "2026-10-10 10:00:00",  # a Saturday that is also an excluded day: excluded day wins
        "2026-12-16 07:00:00",  # an excluded day before opening time: excluded day wins
    ]))
    assert list(oh.classify(t)) == [AFTER, OFFICE, OFFICE, AFTER, OFFICE, AFTER, WEEKEND, EXCLUDED, EXCLUDED]
    assert list(OfficeHours().classify(pd.Series([pd.NaT]))) == [OFFICE]  # unreadable times are never moved out


def test_describe_lists_only_the_report_month():
    oh = hours(("2026-09-24", CLOSED, None, "Heritage Day"), ("2026-10-06", SHORT, "13:00", ""))
    text = oh.describe(pd.Period("2026-09"))
    assert "08:00 to 16:30" in text and "Heritage Day" in text and "06 Oct" not in text


def test_short_day_without_a_time_is_reported_and_ignored():
    oh, warnings = build_office_hours(pd.DataFrame([{"Date": "2026-10-06", "Type": SHORT, "Closes At": None}]))
    assert not oh.short_days and len(warnings) == 1


def test_south_african_public_holidays():
    assert easter_sunday(2026) == date(2026, 4, 5) and easter_sunday(2027) == date(2027, 3, 28)
    days = dict(sa_public_holidays(2026))
    assert len(days) == 13
    assert days[date(2026, 4, 3)] == "Good Friday" and days[date(2026, 4, 6)] == "Family Day"
    assert days[date(2026, 8, 10)] == "National Women's Day (observed)"  # 9 August 2026 is a Sunday
    assert date(2026, 12, 28) not in days  # 26 December 2026 is a Saturday: no Monday off


def test_saved_days_replace_the_defaults_for_that_year_only(tmp_path):
    path = tmp_path / "excluded_days.xlsx"
    assert len(days_for_year(2026, load_days(path))) == 13  # nothing saved: public holidays
    mine = pd.DataFrame([{"Date": date(2026, 3, 20), "Type": CLOSED, "Closes At": None, "Reason": "Once-off"},
                         {"Date": date(2026, 12, 24), "Type": SHORT, "Closes At": time(12, 0), "Reason": "Xmas Eve"}])
    save_days(path, 2026, mine)
    back = days_for_year(2026, load_days(path))
    assert list(back["Date"]) == [date(2026, 3, 20), date(2026, 12, 24)]
    assert back.loc[1, "Closes At"] == time(12, 0) and back.loc[1, "Type"] == SHORT
    assert len(days_for_year(2027, load_days(path))) > 0  # other years still get the public holidays


# ------------------------------------------------------------------ Extension call stats
def outbound(call_id, start):
    return ("Outbound", call_id, start, [leg(call_id, "1", start, "Outbound", "Itumeleng Lehasa<074>", "0829999999",
                                             "ANSWERED")])


EXT_FILE = cdr(
    outbound("B1", "01/09/2026 09:00:00"),  # Tuesday, office hours
    outbound("B2", "01/09/2026 17:00:00"),  # after hours
    outbound("B3", "05/09/2026 10:00:00"),  # Saturday
    outbound("B4", "24/09/2026 10:00:00"),  # Heritage Day (excluded)
    outbound("B5", "25/09/2026 12:00:00"),  # short day, before closing
    outbound("B6", "25/09/2026 14:00:00"),  # short day, after closing
    outbound("B7", "01/09/2026 16:30:00"),  # closing time exactly: after hours
)
EXT_HOURS = hours(("2026-09-24", CLOSED, None, "Heritage Day"), ("2026-09-25", SHORT, "13:00", "Early close"))


def test_extension_stats_count_office_hours_and_show_the_rest_separately():
    res = process_ext_stats([EXT_FILE], MAPPING, NAMES, "Claims Call Centre", period=PERIOD, office_hours=EXT_HOURS)
    assert res.all_calls.set_index("Extension").loc["074", "Total Calls"] == 2  # B1, B5
    out = res.outside.set_index("Period")["Total Calls"]
    assert (out[AFTER], out[WEEKEND], out[EXCLUDED], out["Total outside office hours"]) == (3, 1, 1, 5)
    assert res.outside_by_extension.set_index("Extension").loc["074", "Total Calls"] == 5
    assert "Heritage Day" in res.hours_note


def test_extension_stats_unchanged_without_office_hours():
    res = process_ext_stats([EXT_FILE], MAPPING, NAMES, "Claims Call Centre", period=PERIOD)
    assert res.all_calls.set_index("Extension").loc["074", "Total Calls"] == 7 and res.outside is None


def test_extension_stats_hide_idle_keeps_an_extension_with_only_outside_calls():
    file = cdr(outbound("W1", "05/09/2026 10:00:00"))
    res = process_ext_stats([file], MAPPING, NAMES, "Claims Call Centre", period=PERIOD, hide_idle=True,
                            office_hours=EXT_HOURS)
    assert list(res.all_calls["Extension"]) == ["074"] and res.all_calls["Total Calls"].sum() == 0


def test_extension_stats_outputs_build_with_office_hours():
    res = process_ext_stats([EXT_FILE], MAPPING, NAMES, "Claims Call Centre", period=PERIOD, office_hours=EXT_HOURS)
    assert build_ext_stats_pdf(res)[:5] == b"%PDF-"
    sheets = pd.read_excel(io.BytesIO(build_ext_stats_xlsx(res)), sheet_name=None)
    assert {"Outside office hours", "Outside hours by extension"} <= set(sheets)
    assert sheets["Outside office hours"]["Total Calls"].iloc[-1] == 5


# ------------------------------------------------------------------ Queue overflow
def direct_call(cid, start, status="ANSWERED"):
    """A call routed straight to queue 801; answered by agent 101 when status is ANSWERED."""
    answered = status == "ANSWERED"
    return make_call(cid, start, status=status, queues="Queue 1<801>",
                     legs=[(0, Q1, status)] + ([(10, "Agent One<101>", "ANSWERED")] if answered else []),
                     events=[(0, "basic", route("801")), (0, "queue", entry("801", 10)),
                             (10, "queue", {"to": "101", "round_id": 1} if answered
                              else {"to": "101", "ring_duration": 10, "round_id": 1})])


def run_queue(files, office_hours):
    return process_queue_overflow(files, MAP, ["801", "810"], ISSUES, period=pd.Period("2026-10"),
                                  office_hours=office_hours)


def test_queue_visits_split_by_office_hours():
    files = [("q.xlsx", workbook(direct_call("W1", datetime(2026, 10, 19, 10, 0)),  # Monday, office hours
                                 direct_call("W2", datetime(2026, 10, 19, 18, 0), "NO ANSWER"),  # after hours
                                 direct_call("W3", datetime(2026, 10, 24, 10, 0), "NO ANSWER")))]  # Saturday
    res = run_queue(files, OfficeHours())
    s = res.summary.set_index("Queue Number")
    assert s.loc["801", "Total Calls"] == 1 and s.loc["801", "Answered"] == 1
    out = res.outside.set_index("Period")
    assert (out.loc[AFTER, "Total Calls"], out.loc[WEEKEND, "Total Calls"]) == (1, 1)
    assert out.loc["Total outside office hours", "Missed"] == 2
    assert set(res.outside_detail["Call ID"]) == {"W2", "W3"} and list(res.detail["Call ID"]) == ["W1"]
    assert res.stats["Outside office hours (queue visits)"] == 2
    assert any(term == "Office hours" for term, _ in res.definitions)


def test_queue_all_visits_outside_office_hours_still_builds():
    res = run_queue([("one.xlsx", workbook(*scenario_calls()))], OfficeHours())  # the scenario is on a Sunday
    assert res.detail.empty and res.summary["Total Calls"].sum() == 0
    assert res.outside.set_index("Period").loc[WEEKEND, "Total Calls"] == 15
    assert build_queue_pdf(res)[:5] == b"%PDF-"
    sheets = pd.read_excel(io.BytesIO(build_queue_xlsx(res)), sheet_name=None)
    assert {"Outside office hours", "Outside hours by queue", "Outside hours detail"} <= set(sheets)


def test_queue_unchanged_without_office_hours():
    res = run_queue([("one.xlsx", workbook(*scenario_calls()))], None)
    assert len(res.detail) == 15 and res.outside is None and set(res.detail["Office Hours"]) == {OFFICE}
