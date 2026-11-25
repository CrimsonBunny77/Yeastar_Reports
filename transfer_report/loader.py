"""Reading Yeastar CDR exports.

The real export is *layered*, not a flat table. Each call produces:

  * one summary row      (ID | blank | Time | Inbound/Outbound/Internal | From | To | Last Status ...)
  * one row per call leg (ID | Call Leg ID | Leg start time | type | From | To<ext> | Status ...)
  * one row per event    (ID | Call Leg ID | Event ID | Event type | Event JSON | Event time)

with a repeated header row in front of each block. Transfers are ``operate`` *events*; an event has no
status of its own, so the outcome of a transfer is taken from the call leg that was created for the
transfer target (same call ID, leg's "To" extension == transfer target, leg started at/after the transfer time). If there are several
such legs the final outcome is used: ANSWERED if any leg was answered, otherwise the last leg's status.

Everything is read as text (dtype=str) so IDs are never converted to numbers.
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd

from .config import ATTENDED_RE, BLIND_RE, FROM_RE
from .errors import InvalidFileError, MissingColumnsError

NO_LEG = "NO MATCHING CALL LEG"
N_COLS = 7  # ID, Call Leg ID, Time/Event ID, Type, From/Event, To/Event time, Status
DT_RE = r"^\d{1,4}[/\-.]\d{1,2}[/\-.]\d{1,4}[ T]\d{1,2}:\d{2}(?::\d{2})?$"
LEG_TOLERANCE_SECONDS = 2


def call_coverage(c2: pd.Series, dayfirst: bool) -> tuple:
    """(first call, last call, {date: latest call time that day}) from the call summary rows."""
    times = parse_times(c2, dayfirst).dropna()
    times = times[(times.dt.year >= 2000) & (times.dt.year <= 2100)]
    if times.empty:
        return None, None, {}
    return times.min(), times.max(), times.groupby(times.dt.date).max().to_dict()


@dataclass
class FileResult:
    name: str
    rows: int            # data rows (repeated header rows excluded)
    calls: int           # call summary rows
    operate_rows: int    # operate events
    blind: int
    attended: int
    unmatched: int       # transfers whose target call leg could not be found
    frame: pd.DataFrame  # transfer rows only
    first_call: object = None
    last_call: object = None
    call_days: dict = field(default_factory=dict)
    call_ids: frozenset = field(default_factory=frozenset)


def _read_raw(name: str, data: bytes) -> pd.DataFrame:
    # Yeastar often saves .xlsx content with a .xls name, so try both readers.
    engines = ["xlrd", "openpyxl"] if Path(name).suffix.lower() == ".xls" else ["openpyxl", "xlrd"]
    for engine in engines:
        try:
            df = pd.read_excel(io.BytesIO(data), header=None, dtype=str, engine=engine,
                               usecols=lambda i: i < N_COLS)
            return df
        except Exception:  # noqa: BLE001 - includes a missing reader library; try the next engine
            continue
    raise InvalidFileError(
        f"'{name}' could not be read. The file may be corrupted or is not a valid XLS/XLSX export."
    )


def parse_times(series: pd.Series, dayfirst: bool) -> pd.Series:
    """Parse text timestamps; fast path for dd/mm/yyyy or mm/dd/yyyy, falls back to mixed formats."""
    fmt = "%d/%m/%Y %H:%M:%S" if dayfirst else "%m/%d/%Y %H:%M:%S"
    out = pd.to_datetime(series, format=fmt, errors="coerce")
    if len(series) and out.isna().mean() > 0.05:
        out = pd.to_datetime(series, format="mixed", dayfirst=dayfirst, errors="coerce")
    return out


def read_cdr_file(name: str, data: bytes, dayfirst: bool = True) -> FileResult:
    """Read one CDR file and return its transfer events, each with the transfer outcome."""
    if not data:
        raise InvalidFileError(f"'{name}' is empty.")
    raw = _read_raw(name, data)
    if raw.shape[1] < N_COLS or raw.empty:
        raise MissingColumnsError(f"'{name}' does not look like a Yeastar CDR export (too few columns or rows).")
    raw.columns = [f"c{i}" for i in range(N_COLS)]
    first = raw.iloc[0].fillna("").astype(str).str.strip().tolist()
    if first[0] != "ID" or first[3] != "Communication Type":
        raise MissingColumnsError(
            f"'{name}' does not look like a Yeastar CDR export: the first row should start with "
            "'ID, Call Leg ID, Time, Communication Type'."
        )

    raw = raw[raw["c0"] != "ID"]  # drop every repeated header row
    total = len(raw)
    if total == 0:
        raise InvalidFileError(f"'{name}' contains no data rows.")

    is_dt = raw["c2"].fillna("").str.match(DT_RE)
    has_leg = raw["c1"].notna()
    calls = int((is_dt & ~has_leg).sum())

    events = raw[has_leg & ~is_dt & (raw["c3"].str.strip().str.lower() == "operate")]
    operate_rows = len(events)
    cf = events["c4"].fillna("")
    is_blind = cf.str.contains("blind_to", regex=False)
    is_attended = cf.str.contains("attended_to", regex=False) & ~is_blind
    events, cf, is_blind = events[is_blind | is_attended], cf[is_blind | is_attended], is_blind[is_blind | is_attended]

    out = pd.DataFrame({
        "ID": events["c0"], "Call Leg ID": events["c1"], "Time": events["c5"],
        "TransferType": np.where(is_blind, "Blind", "Attended"),
        "TransferFrom": cf.str.extract(FROM_RE, expand=False),
    })
    blind_to = cf.str.extract(BLIND_RE, expand=False)
    attended_to = cf.str.extract(ATTENDED_RE, expand=False)
    out["TransferTo"] = blind_to.where(is_blind, attended_to)

    # Attended transfers are also logged on the receiving leg without a "from" key; those are
    # duplicates of the real event and are ignored.
    out = out[out["TransferFrom"].notna()].reset_index(drop=True)
    blind = int((out["TransferType"] == "Blind").sum())
    attended = len(out) - blind

    # ---- outcome: find the call leg created for each transfer target
    out["Last Status"] = NO_LEG
    if len(out):
        legs = raw.loc[has_leg & is_dt & raw["c0"].isin(out["ID"].unique()), ["c0", "c2", "c5", "c6"]].copy()
        legs["ext"] = legs["c5"].str.extract(r"<([^<>]*)>\s*$", expand=False)
        legs["leg_start"] = parse_times(legs["c2"], dayfirst)
        legs = legs.dropna(subset=["ext", "leg_start"]).rename(columns={"c0": "ID", "c6": "leg_status"})

        out["tid"] = np.arange(len(out))
        out["event_time"] = parse_times(out["Time"], dayfirst)
        m = out[["tid", "ID", "TransferTo", "event_time"]].merge(
            legs[["ID", "ext", "leg_start", "leg_status"]],
            left_on=["ID", "TransferTo"], right_on=["ID", "ext"], how="inner")
        m["gap"] = (m["leg_start"] - m["event_time"]).dt.total_seconds()
        m = m[m["gap"] >= -LEG_TOLERANCE_SECONDS].sort_values(["tid", "gap"], kind="stable")
        # Outcome of the transfer attempt = what finally happened on the legs to the target extension: ANSWERED if
        # the target answered at any point (a later leg never downgrades an answered transfer), otherwise the status
        # of the last leg (so "no answer / busy, then fell through to voicemail" is Voicemail).
        answered = m["leg_status"].str.strip().str.upper().str.replace("_", " ", regex=False).eq("ANSWERED")
        any_answered = answered.groupby(m["tid"]).any()
        last_status = m.drop_duplicates("tid", keep="last").set_index("tid")["leg_status"]
        status = last_status.where(~any_answered.reindex(last_status.index), "ANSWERED")
        out["Last Status"] = out["tid"].map(status).fillna(NO_LEG)
        out = out.drop(columns=["tid", "event_time"])

    out["Last Status"] = (
        out["Last Status"].fillna(NO_LEG).str.strip().str.upper().str.replace("_", " ", regex=False)
    )
    unmatched = int((out["Last Status"] == NO_LEG).sum())
    first, last, days = call_coverage(raw.loc[is_dt & ~has_leg, "c2"], dayfirst)
    return FileResult(name=name, rows=total, calls=calls, operate_rows=operate_rows,
                      blind=blind, attended=attended, unmatched=unmatched, frame=out,
                      first_call=first, last_call=last, call_days=days,
                      call_ids=frozenset(raw.loc[is_dt & ~has_leg, "c0"]))
