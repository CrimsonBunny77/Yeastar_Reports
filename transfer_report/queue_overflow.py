"""Queue overflow report: calls routed to selected queues, how each call reached the queue, and how it ended.

Uses the same layered Yeastar export as the other reports (call summary rows, call-leg rows, event rows) and a
separate queue mapping (queue number -> queue name, see queue_mapping.py).

The headline figures follow the PBX's own Queue Performance report: every entry into a queue is a visit and is
counted (Total Calls), each visit is Answered / Missed / Abandoned from the queue's own call leg, all call types are
included, visits that rang for no more than MIN_RING_S seconds are not counted, and the visit's Time is when the call
entered the queue. Unique calls (distinct Call IDs) are shown beside the total.

Every entry into a queue is a *visit*. Visits are read from the call's "queue" events that carry
``ring_strategy_queue`` (one per queue the call enters), ordered by event time (Event IDs restart on every call leg,
so they cannot be used for ordering). Each visit is then classified from the events around it:

  * Manual transfer  - an agent transfer (``blind_to`` / ``attended_to`` operate event) was logged on the call after
                       the previous queue visit began (or earlier on the call, for the first visit) and at or before
                       this visit's entry. Checked first, so a manual transfer is never counted as automatic overflow.
  * Automatic overflow - the visit began within OVERFLOW_TOLERANCE_S of the end of the previous visit's ring time
                       (the previous queue timed out) and no agent transfer was logged in between.
  * Direct           - an inbound route / IVR event names this queue as its destination (not a hand-off from another
                       queue).
  * Undetermined     - none of the above; listed as an exception.

A visit is "answered by a queue agent" when the queue's own member-dial event for that visit names an agent that
answered, and that agent has an ANSWERED call leg starting inside the visit window. Agents are therefore read from the
CDR itself (the queue mapping lists queues, not agents).

Everything is read as text (dtype=str).
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Callable, Iterable

import numpy as np
import pandas as pd

from .calls import (EXT_RE, MAX_COLS, SYSTEM_WORDS, _agent, _clean_number, _read, _status, coverage_info,
                    coverage_warning, merge_days)
from .errors import EmptyUploadError, InvalidFileError, MissingColumnsError, ReportError
from .hours import OFFICE, OUTSIDE, OfficeHours
from .loader import DT_RE, parse_times
from .mapping import normalise_extension
from .overlap import file_label, find_overlaps, latest_first, overlap_warning
from .queue_mapping import queue_label

ProgressFn = Callable[[float, str], None]

# ------------------------------------------------------------------ rules (shown to the user via definitions())
OVERFLOW_TOLERANCE_S = 2  # max gap between a queue's ring time ending and the next queue starting
ANSWER_TAIL_S = 5  # an agent leg may start this long after the visit's ring time ends and still count

DIRECT, AUTO, MANUAL, UNDETERMINED = "Direct", "Automatic overflow", "Manual transfer", "Undetermined"
ROUTING_TYPES = [DIRECT, AUTO, MANUAL, UNDETERMINED]
BLANK = "(BLANK)"
KNOWN_STATUSES = ["ANSWERED", "NO ANSWER", "BUSY", "VOICEMAIL", "ABANDONED", "FAILED"]
STATUS_LABELS = {"ANSWERED": "Answered", "NO ANSWER": "No Answer", "BUSY": "Busy", "VOICEMAIL": "Voicemail",
                 "ABANDONED": "Abandoned", "FAILED": "Failed", "OTHER": "Other", "BLANK": "Blank"}
STATUS_COLS = list(STATUS_LABELS.values())
NO_LEG = "NO QUEUE LEG"
MIN_RING_S = 1  # default: queue visits that rang for this many seconds or less are not counted (as the PBX does)
# The PBX's own labels for how a queue visit ended; any other leg status is shown as Other.
VISIT_LABELS = {"ANSWERED": "Answered", "NO ANSWER": "Missed", "ABANDONED": "Abandoned"}
VISIT_COLS = ["Answered", "Missed", "Abandoned", "Other"]

SUMMARY_COLS = (["Queue Number", "Queue", "Total Calls"] + VISIT_COLS
                + ["Answered %", "Missed %", "Abandoned %", "Unique Calls", "Answered by Queue Agent",
                   "Overflowed Out"])
PERCENT_COLS = ["Answered %", "Missed %", "Abandoned %"]
FINAL_COLS = ["Queue Number", "Queue", "Unique Calls"] + STATUS_COLS
ROUTING_COLS = ["Queue Number", "Queue", "Unique Calls", "Visits"] + ROUTING_TYPES
# First eight columns follow the PBX's Queue Performance detail list (Time = when the call entered the queue).
DETAIL_COLS = ["Call ID", "Time", "Caller", "Agent", "Status", "Ring Duration", "Talk Duration", "Hold Duration",
               "Queue Number", "Queue", "Call Type", "Call Time", "Visit #", "Queue Path", "Previous Queue",
               "Routing Type", "Routing Note", "Final Status", "Answered by Queue Agent", "Overflowed To",
               "Office Hours", "Source File"]
OUTSIDE_COLS = ["Period", "Total Calls"] + VISIT_COLS + ["Unique Calls"]
EXCEPTION_COLS = ["Type", "Queue Number", "Queue", "Call ID", "Call Time", "Detail", "Source File"]

# Header names the reader needs. The summary, leg and event blocks each have their own header row and Yeastar
# firmware versions move columns around, so columns are located by name (columns 0-5 are fixed in every layout).
REQUIRED_SUMMARY = ("Communication Type", "Call From", "Call To", "Last Status")
LEG_DEFAULTS = {"Call To": 5, "Status": 6, "Call Duration": 7, "Ring Duration": 8, "Pure Talk Duration": 9, "Hold Duration": 10}
EVENT_DEFAULTS = {"Event": 4, "Event Start Time": 5}


def definitions(period: pd.Period | None = None, min_ring_s: int = MIN_RING_S,
                office_hours: OfficeHours | None = None) -> list[tuple[str, str]]:
    """The wording shown in the app, the PDF and the Excel workbook (one source so they always agree)."""
    month = period.strftime("%B %Y") if period is not None else "the selected month"
    hours = ([("Office hours", "Only queue visits that entered the queue during office hours are counted in the "
                               "summary, routing and final-status tables. Visits after hours, on weekends and on "
                               "excluded days are counted separately in the Outside office hours tables (a visit is "
                               "counted once: excluded day first, then weekend, then after hours). "
                               + office_hours.describe(period))]
             if office_hours else [])
    return [
        ("Report period", f"Queue visits whose entry time (the Time column) falls in {month}, the month selected "
                          "before generating. Visits outside it are left out and counted on the audit page."),
        *hours,
        ("Total calls", "Queue visits: every time a call entered the queue, as the PBX's own Queue Performance "
                        "report counts them. A call that rings the queue, times out and comes back is counted each "
                        "time. All call types are included (inbound, internal and outbound calls that reach the "
                        "queue; the Call Type column shows which)."),
        ("Unique calls", "Distinct Call IDs. A call in several uploaded files, or one that entered the queue more "
                         "than once, is counted once. A call that entered two selected queues is counted once in "
                         "each queue's row and once in the total row."),
        ("Answered / Missed / Abandoned", "How that queue visit ended, from the queue's own call leg: Answered = an "
                                          "agent answered; Missed = the queue rang out with no answer; Abandoned = "
                                          "the caller hung up while waiting. Any other leg status is Other. These "
                                          "match the PBX report's Status column."),
        ("Short visits excluded", f"Queue visits that rang for {min_ring_s} second(s) or less are not counted, like "
                                  "the PBX's \"exclude calls within X seconds\" settings (the minimum is set in "
                                  "the sidebar). They are listed on the Exceptions tab."),
        ("Final status", "The Last Status on the call's summary row, i.e. how the whole call ended. Shown in the "
                         "detail and in the Unique calls by final status table; it can differ from the visit status "
                         "(a visit that was Missed in this queue may end Answered in another queue)."),
        ("Direct", "The call reached the queue straight from an inbound route or IVR event that names the queue, not "
                   "by moving on from another queue."),
        ("Automatic overflow", f"The call entered this queue within {OVERFLOW_TOLERANCE_S} seconds of the end of the "
                               "previous queue visit's ring time (the previous queue timed out), and no agent "
                               "transfer was logged between the two. Order comes from the routing events' times."),
        ("Manual transfer", "An agent transferred the call (a blind or attended transfer event on the call) after the "
                            "previous queue visit began, or earlier on the call for its first queue visit, and at or "
                            "before it entered this queue. Checked first: a manual transfer is never counted as "
                            "automatic overflow, even if the timings also fit."),
        ("Undetermined", "Neither rule fits (no matching timeout, transfer or inbound route/IVR event, or the "
                         "previous ring time is missing). Each one is listed on the Exceptions tab."),
        ("Answered by queue agent", "Yes = the queue's own member-dial event for that visit names an agent who answered, "
                                    "and that agent has an ANSWERED call leg that began during the visit. Agents are "
                                    "taken from the CDR's queue events because the queue mapping lists queues, not "
                                    "agents. No = no agent answered that visit. Unknown = the queue leg shows "
                                    "ANSWERED (or has no queue leg) but no agent the queue rang has a matching "
                                    "answered leg (for example the call was answered by call pickup); listed as an "
                                    "exception. Counted as calls with at least one Yes visit to that queue."),
        ("Routing type counts", "The routing-type table counts queue visits, so the types add up to the Visits "
                                "column (which equals Total calls)."),
        ("Overflowed out", "Calls whose visit to this queue was followed by an automatic overflow into another "
                           "queue."),
    ]


# ------------------------------------------------------------------ results
@dataclass
class QueueAudit:
    files_processed: int = 0
    files_failed: list[tuple[str, str]] = field(default_factory=list)
    total_rows: int = 0
    total_calls: int = 0
    duplicates_removed: int = 0
    per_file: list[dict] = field(default_factory=list)
    first_call: pd.Timestamp | None = None
    last_call: pd.Timestamp | None = None
    call_days: dict = field(default_factory=dict)
    overlaps: list = field(default_factory=list)
    overlap_calls: int = 0
    outside_period: int = 0  # calls with queue visits that started outside the report month
    non_inbound: int = 0  # calls with queue visits that are not inbound (internal / outbound); they are included
    short_visits: int = 0  # queue visits left out because they rang for no more than the minimum ring time
    unmapped_queue_calls: int = 0  # inbound calls that entered a queue that is not in the queue mapping


@dataclass
class QueueResult:
    period: pd.Period
    queues: pd.DataFrame  # the selected queues: Queue Number, Queue Name
    summary: pd.DataFrame
    routing: pd.DataFrame
    detail: pd.DataFrame
    exceptions: pd.DataFrame
    audit: QueueAudit
    warnings: list[str] = field(default_factory=list)
    final: pd.DataFrame = field(default_factory=pd.DataFrame)  # unique calls by final status
    min_ring_s: int = MIN_RING_S
    # Set only when office hours were applied. detail and the tables above then hold office-hours visits only;
    # outside_detail holds the rest, outside the breakdown by group, outside_summary the same by queue.
    office_hours: OfficeHours | None = None
    outside_detail: pd.DataFrame | None = None
    outside: pd.DataFrame | None = None
    outside_summary: pd.DataFrame | None = None

    @property
    def stats(self) -> dict:
        d = self.detail

        def calls_where(mask: pd.Series) -> int:
            return int(d.loc[mask, "_ckey"].nunique())

        return {
            "Queues selected": len(self.queues),
            "Total calls (queue visits)": len(d),
            "Answered": int((d["Status"] == "Answered").sum()),
            "Missed": int((d["Status"] == "Missed").sum()),
            "Abandoned": int((d["Status"] == "Abandoned").sum()),
            "Unique calls": int(d["_ckey"].nunique()),
            "Calls with a recorded agent answer": calls_where(d["Answered by Queue Agent"] == "Yes"),
            "Entered by automatic overflow": calls_where(d["Routing Type"] == AUTO),
            "Entered by manual transfer": calls_where(d["Routing Type"] == MANUAL),
            "Routing undetermined": calls_where(d["Routing Type"] == UNDETERMINED),
            "Exceptions to review": len(self.exceptions),
            **({"Outside office hours (queue visits)": len(self.outside_detail)} if self.office_hours else {}),
        }

    @property
    def definitions(self) -> list[tuple[str, str]]:
        return definitions(self.period, self.min_ring_s, self.office_hours)

    @property
    def hours_note(self) -> str:
        return self.office_hours.describe(self.period) if self.office_hours else ""


def outside_table(visits: pd.DataFrame) -> pd.DataFrame:
    """Queue visits outside office hours by group, with how each visit ended, and a total row."""
    rows = []
    for label, sub in [(p, visits[visits["Office Hours"] == p]) for p in OUTSIDE] + [
            ("Total outside office hours", visits)]:
        vc = sub["Status"].value_counts()
        rows.append({"Period": label, "Total Calls": len(sub), **{c: int(vc.get(c, 0)) for c in VISIT_COLS},
                     "Unique Calls": int(sub["_ckey"].nunique())})
    return pd.DataFrame(rows, columns=OUTSIDE_COLS)


# ------------------------------------------------------------------ small helpers
def _loads(text) -> dict | None:
    try:
        value = json.loads(text)
    except (TypeError, ValueError):
        return None
    return value if isinstance(value, dict) else None


def _num(value) -> str:
    return "" if value is None else re.sub(r"\.0$", "", str(value).strip())


def _exc(kind: str, call_ids, detail: str, source: str = "", queue_number: str = "") -> pd.DataFrame:
    ids = list(call_ids)
    return pd.DataFrame({"Type": kind, "Queue Number": queue_number, "Call ID": ids, "Detail": detail,
                         "Source File": source})


def _secs(raw: pd.Series) -> pd.Series:
    """'hh:mm:ss' (or a plain number of seconds) -> seconds; unreadable -> NaN."""
    parts = raw.fillna("").astype(str).str.strip().str.extract(r"^(?:(\d+):)?(\d+):(\d+)$")
    hms = parts[0].astype(float).fillna(0) * 3600 + parts[1].astype(float) * 60 + parts[2].astype(float)
    plain = pd.to_numeric(raw, errors="coerce")
    return hms.fillna(plain)


def _hms(seconds) -> str:
    return "" if pd.isna(seconds) else f"{int(seconds) // 3600:02d}:{int(seconds) % 3600 // 60:02d}:{int(seconds) % 60:02d}"


def _final_status(raw: pd.Series) -> pd.Series:
    s = _status(raw)
    return s.where(s != "", BLANK)


def _bucket(status: pd.Series) -> pd.Series:
    return pd.Series(np.select([status.isin(KNOWN_STATUSES), status == BLANK],
                               [status.map(STATUS_LABELS), STATUS_LABELS["BLANK"]], STATUS_LABELS["OTHER"]),
                     index=status.index)


# ------------------------------------------------------------------ reading
def parse_queue_layers(name: str, data: bytes):
    """Read a layered Yeastar CDR export and return (summary rows, call-leg rows, event rows, data row count).

    Same layout rules as calls.parse_layers, plus the event rows that the routing logic needs and the optional
    summary "Queue" column. Columns are located by header name.
    """
    if not data:
        raise InvalidFileError(f"'{name}' is empty.")
    raw = _read(name, data)
    if raw.shape[1] < 7 or raw.empty:
        raise MissingColumnsError(f"'{name}' does not look like a Yeastar CDR export (too few columns or rows).")
    raw.columns = range(raw.shape[1])
    header = {str(v).strip(): i for i, v in raw.iloc[0].items() if pd.notna(v) and str(v).strip()}
    if header.get("ID") != 0 or header.get("Communication Type") != 3:
        raise MissingColumnsError(
            f"'{name}' does not look like a Yeastar CDR export: the first row should start with "
            "'ID, Call Leg ID, Time, Communication Type'."
        )
    missing = [h for h in REQUIRED_SUMMARY if h not in header]
    if missing:
        raise MissingColumnsError(f"'{name}' is missing required column(s): {', '.join(missing)}.")

    def block_positions(mask: pd.Series, defaults: dict) -> dict:
        rows = raw[mask]
        pos = ({str(v).strip(): i for i, v in rows.iloc[0].items() if pd.notna(v) and str(v).strip()}
               if len(rows) else {})
        for key, default in defaults.items():
            pos.setdefault(key, default)
        return pos

    leg_pos = block_positions(raw[2].fillna("").str.strip() == "Call Leg Time", LEG_DEFAULTS)
    ev_pos = block_positions(raw[3].fillna("").str.strip() == "Event Type", EVENT_DEFAULTS)

    data_rows = raw[raw[0] != "ID"]
    if data_rows.empty:
        raise InvalidFileError(f"'{name}' contains no data rows.")
    is_dt = data_rows[2].fillna("").str.match(DT_RE)
    has_leg = data_rows[1].notna()
    s_raw, l_raw, e_raw = data_rows[is_dt & ~has_leg], data_rows[is_dt & has_leg], data_rows[has_leg & ~is_dt]

    queue_col = header.get("Queue")
    summ = pd.DataFrame({
        "id": s_raw[0], "c2": s_raw[2], "type": s_raw[header["Communication Type"]],
        "frm": s_raw[header["Call From"]], "to": s_raw[header["Call To"]], "status": s_raw[header["Last Status"]],
        "queue": s_raw[queue_col] if queue_col is not None else np.nan,
    })
    def leg_col(key: str) -> pd.Series:
        """A leg column by header name; blank when the export does not have it."""
        pos = leg_pos[key]
        return l_raw[pos] if pos in l_raw.columns else pd.Series(np.nan, index=l_raw.index)

    legs = pd.DataFrame({"id": l_raw[0], "c2": l_raw[2], "to": l_raw[leg_pos["Call To"]],
                         "status": l_raw[leg_pos["Status"]], "dur": leg_col("Call Duration"),
                         "ring": leg_col("Ring Duration"), "talk": leg_col("Pure Talk Duration"),
                         "hold": leg_col("Hold Duration")})
    events = pd.DataFrame({"id": e_raw[0], "etype": e_raw[3], "event": e_raw[ev_pos["Event"]],
                           "etime": e_raw[ev_pos["Event Start Time"]]})
    return summ, legs, events, len(data_rows)


@dataclass
class QueueFile:
    name: str
    rows: int
    calls: int
    visits: pd.DataFrame
    exceptions: pd.DataFrame
    first_call: pd.Timestamp | None = None
    last_call: pd.Timestamp | None = None
    call_days: dict = field(default_factory=dict)
    call_ids: frozenset = field(default_factory=frozenset)


def read_queue_file(name: str, data: bytes, dayfirst: bool = True) -> QueueFile:
    """One file -> every queue visit on every call, classified (routing type, visit status, agent answer)."""
    summ, legs, events, rows = parse_queue_layers(name, data)
    exc: list[pd.DataFrame] = []

    summ = summ[summ["id"].notna()]
    dup = summ.duplicated("id", keep="first")
    if dup.any():
        exc.append(_exc("Duplicate call record in file", summ.loc[dup, "id"].drop_duplicates(),
                        "Call ID repeated in this file; the first record was used.", name))
    summ = summ[~dup].copy()
    first_call, last_call, days = coverage_info(summ["c2"], dayfirst)
    summ["t"] = parse_times(summ["c2"], dayfirst)
    call_ids = frozenset(summ["id"])

    # ---- the routing events we care about, parsed once
    et = events["etype"].fillna("").str.strip().str.lower()
    body = events["event"].fillna("")
    is_q = (et == "queue") & body.str.contains("ring_strategy_queue|round_id", regex=True)
    is_tr = (et == "operate") & body.str.contains("blind_to|attended_to", regex=True)
    is_rt = et.isin(["basic", "ivr"]) & body.str.contains("destination_type", regex=False)
    sel = events[is_q | is_tr | is_rt].copy()
    sel["kind"] = et[sel.index]
    sel["j"] = sel["event"].map(_loads)
    bad = sel[sel["j"].isna()]
    if len(bad):
        exc.append(_exc("Unreadable routing event", bad["id"].drop_duplicates(),
                        f"{len(bad)} routing event(s) in this file could not be read; routing may be incomplete.", name))
    sel = sel[sel["j"].notna()].copy()
    sel["t"] = parse_times(sel["etime"], dayfirst)
    sel["_row"] = sel.index

    q = sel[sel["kind"] == "queue"]
    is_entry = q["j"].map(lambda d: "ring_strategy_queue" in d)
    e = q[is_entry]
    ent = pd.DataFrame({
        "id": e["id"].values, "t": e["t"].values, "_row": e["_row"].values,
        "qno": e["j"].map(lambda d: _num(d.get("to"))).values,
        "rd": pd.to_numeric(e["j"].map(lambda d: d.get("ring_duration")), errors="coerce").values,
    })
    ent = ent[ent["qno"] != ""]
    att_src = q[~is_entry & q["j"].map(lambda d: "round_id" in d and "to" in d)]
    att = pd.DataFrame({
        "id": att_src["id"].values, "a_t": att_src["t"].values,
        "agent": att_src["j"].map(lambda d: _num(d.get("to"))).values,
        "answered": att_src["j"].map(lambda d: "ring_duration" not in d).values,
    })
    att = att[att["answered"]]

    # ---- calls whose summary lists queues but whose routing events are missing: keep them, as Undetermined
    chain = summ["queue"].fillna("").map(lambda s: re.findall(r"<([^<>]*)>", s)) if len(summ) else pd.Series(dtype=object)
    chain = pd.Series(chain.values, index=summ["id"].values)
    chain = chain[chain.map(len) > 0]
    ent = ent.sort_values(["id", "t", "_row"], kind="stable")
    ent_seq = ent.groupby("id")["qno"].agg(list)
    synth = chain[~chain.index.isin(ent_seq.index)]
    if len(synth):
        rows_s = [{"id": cid, "t": pd.NaT, "_row": -1, "qno": _num(n), "rd": np.nan}
                  for cid, nums in synth.items() for n in nums]
        ent = pd.concat([ent, pd.DataFrame(rows_s)], ignore_index=True)
        exc.append(_exc("Queue routing events missing", synth.index,
                        "Summary lists queue(s) but the export has no routing events for the call.", name))
    both = chain[chain.index.isin(ent_seq.index)]
    differs = [cid for cid, nums in both.items() if [_num(n) for n in nums] != ent_seq[cid]]
    if differs:
        exc.append(_exc("Summary queue list differs from routing events", differs,
                        "Summary Queue column differs from the routing events; the events were used.", name))

    ent = ent.sort_values(["id", "t", "_row"], kind="stable", na_position="last").reset_index(drop=True)
    ent["vid"] = np.arange(len(ent))
    ent["k"] = ent.groupby("id").cumcount()
    ent["end"] = ent["t"] + pd.to_timedelta(ent["rd"], unit="s")
    g = ent.groupby("id")
    ent["prev_t"], ent["prev_end"] = g["t"].shift(1), g["end"].shift(1)
    ent["prev_qno"], ent["prev_rd"] = g["qno"].shift(1), g["rd"].shift(1)
    ent["next_qno"] = g["qno"].shift(-1)
    ent["next_t"] = g["t"].shift(-1)
    ent["win_end"] = ent["end"].fillna(ent["next_t"]).fillna(ent["t"] + pd.Timedelta(days=1))
    ent["visit_no"] = ent.groupby(["id", "qno"]).cumcount() + 1

    # ---- manual transfer: latest agent transfer inside (previous visit start, this entry]; any earlier one if k == 0
    t_src = sel[sel["kind"] == "operate"]
    tr = pd.DataFrame({
        "id": t_src["id"].values, "tr_t": t_src["t"].values,
        "tr_from": t_src["j"].map(lambda d: _num(d.get("from"))).values,
        "tr_to": t_src["j"].map(lambda d: _num(d.get("blind_to") or d.get("attended_to"))).values,
        "tr_kind": t_src["j"].map(lambda d: "Blind" if d.get("blind_to") else "Attended").values,
    })
    tr = tr[tr["tr_to"] != ""]
    m = ent[["vid", "id", "k", "t", "prev_t"]].merge(tr, on="id")
    m = m[(m["tr_t"] <= m["t"]) & ((m["k"] == 0) | (m["tr_t"] > m["prev_t"]))]
    m = m.sort_values(["vid", "tr_t"], kind="stable").drop_duplicates("vid", keep="last")
    ent = ent.merge(m[["vid", "tr_t", "tr_from", "tr_to", "tr_kind"]], on="vid", how="left")

    # ---- direct: an inbound route / IVR event whose destination is this queue
    r_src = sel[sel["kind"].isin(["basic", "ivr"])]
    r_src = r_src[r_src["j"].map(lambda d: d.get("destination_type") == "queue")]
    rt = pd.DataFrame({
        "id": r_src["id"].values, "rt_t": r_src["t"].values,
        "rt_dest": r_src["j"].map(lambda d: _num(d.get("to") if "to" in d else d.get("ivr_to"))).values,
        "rt_src": [("IVR menu" if k == "ivr" else "inbound route") for k in r_src["kind"]],
    })
    m = ent[["vid", "id", "k", "t", "prev_t", "qno"]].merge(rt, on="id")
    tol = pd.Timedelta(seconds=OVERFLOW_TOLERANCE_S)
    m = m[(m["rt_dest"] == m["qno"]) & (m["rt_t"] <= m["t"] + tol) & ((m["k"] == 0) | (m["rt_t"] > m["prev_t"]))]
    m = m.sort_values(["vid", "rt_t"], kind="stable").drop_duplicates("vid", keep="last")
    ent = ent.merge(m[["vid", "rt_t", "rt_src"]], on="vid", how="left")

    # ---- classify (manual first, so a manual transfer is never counted as automatic overflow)
    ent["gap"] = (ent["t"] - ent["prev_end"]).dt.total_seconds()
    is_manual = ent["tr_t"].notna()
    is_auto = ~is_manual & (ent["k"] > 0) & ent["gap"].abs().le(OVERFLOW_TOLERANCE_S)
    is_direct = ~is_manual & ~is_auto & ent["rt_t"].notna()
    ent["rtype"] = np.select([is_manual, is_auto, is_direct], [MANUAL, AUTO, DIRECT], UNDETERMINED)

    def note(r) -> str:
        if r.rtype == MANUAL:
            return (f"{r.tr_kind} transfer from {r.tr_from or '?'} to {r.tr_to} at {r.tr_t:%H:%M:%S}, "
                    "before entering the queue")
        if r.rtype == AUTO:
            return (f"No answer in queue {r.prev_qno} after {r.prev_rd:.0f}s; entered this queue "
                    f"{abs(r.gap):.0f}s after it timed out")
        if r.rtype == DIRECT:
            return f"Routed here by an {r.rt_src} event"
        if pd.isna(r.t):
            return "No routing events for this call in the export"
        if r.k > 0 and pd.isna(r.prev_end):
            return f"Entered after queue {r.prev_qno}, whose ring time is missing, and no transfer was logged"
        if r.k > 0:
            return (f"Entered after queue {r.prev_qno}, but {abs(r.gap):.0f}s from the end of its ring time, "
                    "with no transfer logged")
        return "No inbound route / IVR event for this queue and no transfer was logged"

    ent["note"] = [note(r) for r in ent.itertuples(index=False)]
    ent["next_rtype"] = ent.groupby("id")["rtype"].shift(-1)

    # ---- visit status: the queue's own call leg, matched on call, queue number and start time
    lg = legs[legs["id"].isin(ent["id"])].copy()
    lg["ext"] = normalise_extension(lg["to"].str.extract(EXT_RE, expand=False))
    lg["l_t"] = parse_times(lg["c2"], dayfirst)
    lg["l_status"] = _status(lg["status"])
    for col in ("dur", "ring", "talk", "hold"):
        lg[col] = _secs(lg[col])
    ql = ent[["vid", "id", "qno", "t"]].merge(lg[["id", "ext", "l_t", "l_status", "dur", "ring", "talk", "hold"]],
                                              left_on=["id", "qno"], right_on=["id", "ext"])
    ql["dt"] = (ql["l_t"] - ql["t"]).abs().dt.total_seconds()
    ql = ql[ql["dt"] <= OVERFLOW_TOLERANCE_S].sort_values(["vid", "dt"], kind="stable").drop_duplicates("vid")
    ent = ent.merge(ql[["vid", "l_status", "dur", "ring", "talk", "hold"]].rename(
        columns={"dur": "q_dur", "ring": "q_ring", "talk": "q_talk", "hold": "q_hold"}),
                    on="vid", how="left")
    ent["l_status"] = ent["l_status"].replace("", np.nan).fillna(NO_LEG)

    # ---- answered by a queue agent: answered member-dial event in the visit window + that agent's ANSWERED leg
    agent_legs = lg[(lg["l_status"] == "ANSWERED") & lg["ext"].notna()].copy()
    agent_legs["name"] = _agent(agent_legs["to"])
    agent_legs = agent_legs[~agent_legs["name"].str.lower().str.contains("|".join(SYSTEM_WORDS), regex=True)]
    w = ent[["vid", "id", "t", "win_end"]].merge(att, on="id")
    pad = pd.Timedelta(seconds=OVERFLOW_TOLERANCE_S)
    w = w[(w["a_t"] >= w["t"] - pad) & (w["a_t"] <= w["win_end"] + pad)]
    w = w.merge(agent_legs[["id", "ext", "l_t", "name"]], left_on=["id", "agent"], right_on=["id", "ext"])
    w = w[(w["l_t"] >= w["t"] - pad) & (w["l_t"] <= w["win_end"] + pd.Timedelta(seconds=ANSWER_TAIL_S))]
    w = w.sort_values(["vid", "l_t"], kind="stable").drop_duplicates("vid")
    w["agent_text"] = [f"{n}<{a}>" if n else a for n, a in zip(w["name"], w["agent"])]  # as the PBX writes it
    ent = ent.merge(w[["vid", "agent_text"]], on="vid", how="left")
    matched = ent["agent_text"].notna()
    ent["answered"] = np.where(matched, "Yes",
                               np.where(ent["l_status"].isin(["ANSWERED", NO_LEG]), "Unknown", "No"))
    ent["agent_text"] = ent["agent_text"].fillna("")
    # Durations as the PBX's Queue Performance detail reports them, from the queue's own call leg: ring = the
    # visit's ring time, talk = call duration less ring time (so it includes hold), hold = hold time.
    ent["ring_s"] = ent["rd"].fillna(ent["q_ring"])
    ent["talk_s"] = np.where(ent["l_status"] == "ANSWERED", ent["q_dur"] - ent["q_ring"], ent["q_talk"])
    ent["hold_s"] = ent["q_hold"]
    # the PBX names an agent only on answered visits
    ent["agent_text"] = np.where(ent["l_status"] == "ANSWERED", ent["agent_text"], "")

    # ---- attach the call-level fields
    info = pd.DataFrame({
        "id": summ["id"].values, "Call Time": summ["t"].values, "Caller": _clean_number(summ["frm"]).values,
        "Call Type": summ["type"].fillna("").str.strip().str.title().values,
        "Final Status": _final_status(summ["status"]).values,
    })
    path = ent.groupby("id")["qno"].agg(" > ".join).rename("Queue Path")
    out = ent.merge(info, on="id", how="inner").merge(path, left_on="id", right_index=True)
    visits = pd.DataFrame({
        "Call ID": out["id"], "Call Time": out["Call Time"], "Caller": out["Caller"], "Call Type": out["Call Type"],
        "Queue Number": out["qno"], "Visit #": out["visit_no"], "Position": out["k"] + 1,
        "Queue Path": out["Queue Path"], "Previous Queue Number": out["prev_qno"].fillna(""),
        "Next Queue Number": out["next_qno"].fillna(""), "Next Routing Type": out["next_rtype"].fillna(""),
        "Routing Type": out["rtype"], "Routing Note": out["note"], "Visit Status": out["l_status"],
        "Final Status": out["Final Status"], "Answered by Queue Agent": out["answered"], "Agent": out["agent_text"],
        "Time": out["t"].fillna(out["Call Time"]), "Ring Sec": out["rd"], "Ring S": out["ring_s"],
        "Talk S": out["talk_s"], "Hold S": out["hold_s"],
    }).reset_index(drop=True)
    return QueueFile(name, rows, len(summ), visits, pd.concat(exc, ignore_index=True) if exc else pd.DataFrame(),
                     first_call, last_call, days, call_ids)


# ------------------------------------------------------------------ combining files
def _period_of(times: pd.Series) -> pd.Period | None:
    times = times.dropna()
    return times.dt.to_period("M").mode().iloc[0] if len(times) else None


def _summarise(sel: pd.DataFrame, queues: pd.DataFrame, names: dict) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """(summary by queue, routing type by queue, unique calls by final status).

    The summary counts queue visits (the PBX's Total Calls) and how each visit ended. The final-status table counts
    unique calls.
    """
    s_rows, r_rows, f_rows = [], [], []

    def counts(frame: pd.DataFrame, number: str, label: str) -> tuple[dict, dict, dict]:
        calls = frame.drop_duplicates("_ckey")
        by_visit = frame["Status"].value_counts()
        by_final = _bucket(calls["Final Status"]).value_counts() if len(calls) else pd.Series(dtype=int)
        by_type = frame["Routing Type"].value_counts()  # visits, so the types add up to the Visits column
        total = len(frame)
        pct = lambda n: round(100 * n / total, 2) if total else 0.0  # noqa: E731
        s = {"Queue Number": number, "Queue": label, "Total Calls": total,
             **{c: int(by_visit.get(c, 0)) for c in VISIT_COLS},
             **{p: pct(int(by_visit.get(p[:-2], 0))) for p in PERCENT_COLS},
             "Unique Calls": len(calls),
             "Answered by Queue Agent": int(frame.loc[frame["Answered by Queue Agent"] == "Yes", "_ckey"].nunique()),
             "Overflowed Out": int(frame.loc[frame["Next Routing Type"] == AUTO, "_ckey"].nunique())}
        r = {"Queue Number": number, "Queue": label, "Unique Calls": len(calls), "Visits": total,
             **{t: int(by_type.get(t, 0)) for t in ROUTING_TYPES}}
        f = {"Queue Number": number, "Queue": label, "Unique Calls": len(calls),
             **{c: int(by_final.get(c, 0)) for c in STATUS_COLS}}
        return s, r, f

    for number in queues["Queue Number"]:
        s, r, f = counts(sel[sel["Queue Number"] == number], number, names.get(number) or number)
        s_rows.append(s)
        r_rows.append(r)
        f_rows.append(f)
    s, r, f = counts(sel, "", "All selected queues")
    s_rows.append(s)
    r_rows.append(r)
    f_rows.append(f)
    return (pd.DataFrame(s_rows, columns=SUMMARY_COLS), pd.DataFrame(r_rows, columns=ROUTING_COLS),
            pd.DataFrame(f_rows, columns=FINAL_COLS))


def process_queue_overflow(
    files: Iterable[tuple[str, bytes]],
    queue_map: pd.DataFrame,
    selected: list[str],
    mapping_issues: pd.DataFrame | None = None,
    dayfirst: bool = True,
    dedupe: bool = True,
    period: pd.Period | None = None,
    progress: ProgressFn | None = None,
    min_ring_s: int = MIN_RING_S,
    office_hours: OfficeHours | None = None,
) -> QueueResult:
    """office_hours: when given, the tables count only queue visits that entered the queue inside office hours;
    visits after hours, on weekends and on excluded days are counted separately (result.outside)."""
    files = list(files)
    if not files:
        raise EmptyUploadError("No CDR files were uploaded. Upload at least one .xls / .xlsx file.")
    if not selected:
        raise ReportError("Select at least one queue to include in the report.")
    names = dict(zip(queue_map["Queue Number"], queue_map["Queue Name"]))
    queues = queue_map[queue_map["Queue Number"].isin(selected)].reset_index(drop=True)
    audit = QueueAudit()
    parsed: list[QueueFile] = []
    for i, (name, data) in enumerate(files):
        if progress:
            progress(0.05 + 0.75 * i / len(files), f"Reading {name} ({i + 1} of {len(files)})…")
        try:
            parsed.append(read_queue_file(name, data, dayfirst))
        except ReportError as exc:
            audit.files_failed.append((name, str(exc)))
    if not parsed:
        raise ReportError("None of the uploaded files could be read. " + " ".join(m for _, m in audit.files_failed))
    if progress:
        progress(0.82, "Combining files…")

    audit.files_processed = len(parsed)
    audit.total_rows = sum(f.rows for f in parsed)
    audit.total_calls = sum(f.calls for f in parsed)
    firsts = [f.first_call for f in parsed]
    audit.first_call = min((x for x in firsts if x is not None), default=None)
    audit.last_call = max((f.last_call for f in parsed if f.last_call is not None), default=None)
    for f in parsed:
        audit.call_days = merge_days(audit.call_days, f.call_days)
    audit.overlaps, audit.overlap_calls = find_overlaps(
        [(file_label(i + 1, f.name), f.call_ids) for i, f in enumerate(parsed)])
    for i, f in enumerate(parsed):
        audit.per_file.append({"File": f.name, "Rows": f.rows, "Calls": f.calls,
                               "Queue visits": len(f.visits)})

    # ---- merge files; the later export's copy of a repeated call is the one kept
    order = latest_first(list(range(len(parsed))), firsts)
    owner: dict[str, int] = {}
    if dedupe:
        for i in order:
            for cid in parsed[i].call_ids:
                owner.setdefault(cid, i)
    visit_parts, exc_parts = [], []
    for i in order:
        f = parsed[i]
        v = f.visits.assign(**{"Source File": f.name, "_f": i})
        if dedupe and len(v):
            keep = v["Call ID"].map(owner) == i
            dropped = v.loc[~keep, "Call ID"].drop_duplicates()
            if len(dropped):
                where = {cid: parsed[owner[cid]].name for cid in dropped}
                d = _exc("Duplicate call record removed", dropped,
                         "", f.name)
                d["Detail"] = [f"The call is also in '{where[c]}'; that copy was kept." for c in d["Call ID"]]
                exc_parts.append(d)
                audit.duplicates_removed += len(dropped)
            v = v[keep]
        visit_parts.append(v)
        if len(f.exceptions):
            exc_parts.append(f.exceptions)
    visits = pd.concat(visit_parts, ignore_index=True) if visit_parts else pd.DataFrame()
    visits["_ckey"] = visits["Call ID"] if dedupe else visits["Call ID"] + "@" + visits["_f"].astype(str)

    call_time = visits.drop_duplicates("_ckey").set_index("_ckey")["Call Time"]
    if period is None:
        period = _period_of(call_time)
        if period is None:
            raise ReportError("No calls with a readable date were found in the uploaded files.")
    unreadable = call_time[call_time.isna()]
    in_period = visits["Time"].dt.to_period("M") == period
    outside = visits[~in_period & visits["Time"].notna()]
    audit.outside_period = outside["_ckey"].nunique()
    if len(unreadable):
        exc_parts.append(_exc("Call time could not be read", [c.split("@")[0] for c in unreadable.index],
                              "Start time could not be read; the call was left out of the report."))
    visits = visits[in_period]

    # ---- every call type that enters a queue is counted (the PBX does); the Call Type column tells them apart
    audit.non_inbound = visits.loc[visits["Call Type"].str.lower() != "inbound", "_ckey"].nunique()
    unmapped = visits[~visits["Queue Number"].isin(names)]
    audit.unmapped_queue_calls = unmapped["_ckey"].nunique()
    if len(unmapped):
        d = _exc("Queue number not in the queue mapping", unmapped["Call ID"],
                 "Queue is not in the queue mapping, so it cannot be named or selected.")
        d["Queue Number"], d["Call Time"], d["Source File"] = unmapped["Queue Number"].values, \
            unmapped["Call Time"].values, unmapped["Source File"].values
        exc_parts.append(d.drop_duplicates(["Call ID", "Queue Number"]))

    sel = visits[visits["Queue Number"].isin(selected)].copy()
    sel["Queue"] = sel["Queue Number"].map(names)
    label = lambda n: queue_label(n, names.get(n)) if n else ""  # noqa: E731
    sel["Previous Queue"] = sel["Previous Queue Number"].map(label)
    sel["Overflowed To"] = np.where(sel["Next Routing Type"] == AUTO, sel["Next Queue Number"].map(label), "")
    sel = sel.sort_values(["Time", "Call ID", "Position"], kind="stable").reset_index(drop=True)
    sel["Status"] = sel["Visit Status"].map(VISIT_LABELS).fillna("Other")
    sel["Ring Duration"], sel["Talk Duration"], sel["Hold Duration"] = (
        sel["Ring S"].map(_hms), sel["Talk S"].map(_hms), sel["Hold S"].map(_hms))
    sel["Agent"] = sel["Agent"].replace("", "NONE")

    # ---- visits that rang for no more than the minimum ring time are not counted, but are listed
    short = sel["Ring Sec"].le(min_ring_s)
    audit.short_visits = int(short.sum())
    if short.any():
        sub = sel[short]
        d = _exc(f"Queue visit not counted (rang {min_ring_s}s or less)", sub["Call ID"],
                 f"Rang {min_ring_s} second(s) or less, so it is not counted. Change the minimum ring time in "
                 "the sidebar.")
        d["Queue Number"], d["Call Time"], d["Source File"] = (sub["Queue Number"].values, sub["Time"].values,
                                                              sub["Source File"].values)
        exc_parts.append(d)
        sel = sel[~short].reset_index(drop=True)

    # ---- exceptions that stay in the report but need review
    def flag(mask: pd.Series, kind: str, detail) -> None:
        sub = sel[mask]
        if sub.empty:
            return
        d = _exc(kind, sub["Call ID"], "")
        d["Detail"] = detail(sub) if callable(detail) else detail
        d["Queue Number"], d["Call Time"], d["Source File"] = (sub["Queue Number"].values, sub["Time"].values,
                                                              sub["Source File"].values)
        exc_parts.append(d)

    first_visit = sel["Visit #"] == 1
    flag(sel["Final Status"] == BLANK, "Final status is blank",
         "Last Status is empty; counted under Blank.")
    flag(first_visit & ~sel["Final Status"].isin(KNOWN_STATUSES + [BLANK]), "Final status not recognised",
         lambda s: ["Last Status is '" + x + "'; counted under Other." for x in s["Final Status"]])
    flag(sel["Routing Type"] == UNDETERMINED, "Routing type cannot be determined",
         lambda s: list(s["Routing Note"]))
    flag((sel["Answered by Queue Agent"] == "Unknown") & (sel["Visit Status"] == "ANSWERED"),
         "Queue answered but agent not matched",
         "Queue leg ANSWERED but no queue-rung agent matched (often call pickup).")
    flag(sel["Visit Status"] == NO_LEG, "No queue call leg found for this visit",
         "No call leg found for this queue visit; visit status unknown.")
    for _, row in (mapping_issues if mapping_issues is not None else pd.DataFrame()).iterrows():
        exc_parts.append(pd.DataFrame([{"Type": row["Type"], "Queue Number": row["Queue Number"], "Call ID": "",
                                        "Detail": row["Detail"], "Source File": "Queue mapping file"}]))

    if sel.empty:
        extra = (f" {audit.outside_period:,} call(s) entered queues outside {period.strftime('%B %Y')}; check the "
                 "report month." if audit.outside_period else "")
        raise ReportError("No calls were routed to the selected queue(s) in "
                          f"{period.strftime('%B %Y')}.{extra}")

    exceptions = pd.concat(exc_parts, ignore_index=True) if exc_parts else pd.DataFrame(columns=EXCEPTION_COLS)
    exceptions["Queue"] = exceptions["Queue Number"].map(names).fillna("")
    exceptions["Call Time"] = pd.to_datetime(exceptions["Call Time"]) if "Call Time" in exceptions else pd.NaT
    exceptions = exceptions.reindex(columns=EXCEPTION_COLS).fillna({"Detail": "", "Source File": "", "Call ID": ""})
    exceptions = exceptions.sort_values(["Type", "Queue Number", "Call Time", "Call ID"],
                                        kind="stable").reset_index(drop=True)

    if progress:
        progress(0.9, "Summarising…")
    sel["Position"] = sel["Position"].astype(int)
    # ---- office hours: by the time the call entered the queue
    sel["Office Hours"] = office_hours.classify(sel["Time"]) if office_hours is not None else OFFICE
    out_sel = sel[sel["Office Hours"] != OFFICE].reset_index(drop=True)
    sel = sel[sel["Office Hours"] == OFFICE].reset_index(drop=True)
    summary, routing, final = _summarise(sel, queues, names)
    keep_cols = DETAIL_COLS + ["_ckey", "Position", "Next Routing Type"]
    detail = sel[keep_cols]
    outside = outside_summary = outside_detail = None
    if office_hours is not None:
        outside_detail = out_sel[keep_cols].reset_index(drop=True)
        outside = outside_table(out_sel)
        outside_summary = _summarise(out_sel, queues, names)[0]

    warnings: list[str] = []
    cov = coverage_warning(audit.first_call, audit.last_call, audit.call_days, period)
    if cov:
        warnings.append(cov)
    ov = overlap_warning(audit.overlaps, audit.overlap_calls, dedupe)
    if ov:
        warnings.append(ov)
    if audit.outside_period:
        warnings.append(f"{audit.outside_period:,} call(s) that entered a queue started outside "
                        f"{period.strftime('%B %Y')} and were left out.")
    if audit.unmapped_queue_calls:
        warnings.append(f"{audit.unmapped_queue_calls:,} call(s) entered a queue that is not in the queue "
                        "mapping; see the Exceptions tab.")
    n_und = int((sel["Routing Type"] == UNDETERMINED).sum())
    if n_und:
        warnings.append(f"{n_und:,} queue visit(s) have a routing type that could not be determined; "
                        "see the Exceptions tab.")
    if progress:
        progress(0.93, "Done reading.")
    return QueueResult(period, queues, summary, routing, detail.reset_index(drop=True), exceptions, audit, warnings,
                       final, min_ring_s, office_hours, outside_detail, outside, outside_summary)


# ------------------------------------------------------------------ filtering (used by the app's detail table)
def filter_detail(detail: pd.DataFrame, start=None, end=None, queues: list[str] | None = None,
                  outcomes: list[str] | None = None, routing_types: list[str] | None = None) -> pd.DataFrame:
    """Filter the detail table by date the call entered the queue (inclusive), queue number, visit status (Answered,
    Missed, Abandoned, Other) and routing type."""
    mask = pd.Series(True, index=detail.index)
    if start is not None:
        mask &= detail["Time"].dt.date >= start
    if end is not None:
        mask &= detail["Time"].dt.date <= end
    if queues is not None:
        mask &= detail["Queue Number"].isin(queues)
    if outcomes is not None:
        mask &= detail["Status"].isin(outcomes)
    if routing_types is not None:
        mask &= detail["Routing Type"].isin(routing_types)
    return detail[mask]
