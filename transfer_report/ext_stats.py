"""Department extension call statistics.

For every extension of the selected department: Total Calls, Answered, No Answer, Busy, Voicemail, Abandoned,
Transfers Not Completed and Answer %. The counting rule mirrors the PBX's own "Extension Call Statistics" report: every call leg the
extension was on, as caller or as callee, counts once under that leg's own status (including the extension's own
voicemail-box leg). Inbound / Outbound / Internal is the leg's OWN communication type, as in the PBX report (so a
transferred-in or voicemail leg on a call that was dialled out counts as Inbound for the extension that received it).

Uses the same layered Yeastar export as the other reports (see calls.parse_layers).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Iterable

import pandas as pd

from .calls import (EXT_RE, SYSTEM_WORDS, _agent, _status, coverage_info, coverage_warning, merge_days,
                    parse_layers)
from .loader import parse_times
from .overlap import file_label, find_overlaps, latest_first, overlap_warning
from .config import NOT_COMPLETED
from .errors import EmptyUploadError, InvalidFileError, MissingColumnsError, ReportError
from .hours import OFFICE, OUTSIDE, OfficeHours
from .mapping import department_at, department_extensions, moves_notice, normalise_extension

ProgressFn = Callable[[float, str], None]

OUTCOMES = ["Answered", "No Answer", "Busy", "Voicemail", "Abandoned", NOT_COMPLETED]
_OUTCOME_KEYS = {"ANSWERED": "Answered", "NOANSWER": "No Answer", "BUSY": "Busy",
                 "VOICEMAIL": "Voicemail", "ABANDONED": "Abandoned"}
TABLE_COLS = ["Extension", "Name", "Total Calls"] + OUTCOMES + ["Answer %"]
EVENT_COLS = ["Extension", "Direction", "Outcome", "Call ID", "Leg", "Side", "CdrName", "Time"]
# Words marking legs that are not an extension. "Voicemail <name><ext>" legs ARE kept: they are an extension's own
# voicemail box and the PBX counts them under that extension.
NON_EXTENSION_WORDS = tuple(w for w in SYSTEM_WORDS if w != "voicemail")


def categorise(status: pd.Series) -> pd.Series:
    """Map a raw status to one of OUTCOMES ('No Answer', 'NO_ANSWER' and 'noanswer' are all the same)."""
    key = _status(status).str.replace(r"[^A-Z0-9]", "", regex=True)
    return key.map(_OUTCOME_KEYS).fillna(NOT_COMPLETED)


def _ext_sort_key(ext: str):
    return (0, int(ext), ext) if ext.isdigit() else (1, 0, ext)


# ------------------------------------------------------------------ reading
@dataclass
class ExtFile:
    name: str
    rows: int
    calls: int
    events: pd.DataFrame
    first_call: pd.Timestamp | None = None
    last_call: pd.Timestamp | None = None
    call_days: dict = field(default_factory=dict)
    call_ids: frozenset = field(default_factory=frozenset)


def read_ext_file(name: str, data: bytes, dept_exts: set[str], include_internal: bool,
                  dayfirst: bool = True) -> ExtFile:
    """One event per call leg the department's extensions were on, as caller (From) or callee (To).

    This is how the PBX's own "Extension Call Statistics" report counts: a leg's outcome is that leg's Status,
    including the extension's own voicemail-box leg ("Voicemail <name><ext>"). Direction is the leg's own
    Communication Type (Inbound / Outbound / Internal), falling back to the call's type only if a leg has none. Queue, IVR and ring-group legs are not extensions and are ignored.
    """
    summ, legs, rows = parse_layers(name, data)
    call_type = summ.drop_duplicates("id").set_index("id")["type"].fillna("").str.strip().str.lower()
    lg = legs.copy()
    leg_type = lg["ltype"].fillna("").str.strip().str.lower()
    lg["ctype"] = leg_type.where(leg_type != "", lg["id"].map(call_type).fillna(""))
    call_time = pd.Series(parse_times(summ["c2"], dayfirst).values, index=summ["id"].values)
    call_time = call_time[~call_time.index.duplicated()]
    lg["Time"] = lg["id"].map(call_time)
    lg = lg[lg["ctype"].isin(["inbound", "outbound"] + (["internal"] if include_internal else []))]
    pattern = "|".join(NON_EXTENSION_WORDS)

    parts = []
    for side in ("frm", "to"):
        ext = normalise_extension(lg[side].str.extract(EXT_RE, expand=False))
        label = _agent(lg[side])
        keep = ext.isin(dept_exts) & ~label.str.lower().str.contains(pattern, regex=True)
        sub = lg[keep]
        parts.append(pd.DataFrame({
            "Extension": ext[keep], "Direction": sub["ctype"].str.title(), "Outcome": categorise(sub["status"]),
            "Call ID": sub["id"], "Leg": sub["leg"].fillna(""), "Side": side,
            "CdrName": label[keep].str.replace(r"^Voicemail\s+", "", regex=True), "Time": sub["Time"],
        }))
    events = pd.concat(parts, ignore_index=True)
    first, last, days = coverage_info(summ["c2"], dayfirst)
    return ExtFile(name, rows, len(summ), events, first, last, days, frozenset(summ["id"]))


# ------------------------------------------------------------------ tables
def stats_table(events: pd.DataFrame, extensions: list[str], names: dict[str, str]) -> pd.DataFrame:
    """One row per extension (including extensions with no calls)."""
    counts = (events.groupby(["Extension", "Outcome"]).size().unstack(fill_value=0)
              if len(events) else pd.DataFrame())
    counts = counts.reindex(index=extensions, columns=OUTCOMES, fill_value=0).fillna(0).astype(int)
    out = counts.copy()
    out.insert(0, "Total Calls", counts.sum(axis=1))
    out["Answer %"] = (out["Answered"] / out["Total Calls"].where(out["Total Calls"] > 0) * 100).round(1).fillna(0.0)
    out.insert(0, "Name", [names.get(e, "") for e in out.index])
    out = out.rename_axis("Extension").reset_index()
    out.columns.name = None
    return out[TABLE_COLS]


def total_row(table: pd.DataFrame) -> dict:
    t = {c: int(table[c].sum()) for c in ["Total Calls"] + OUTCOMES}
    t["Answer %"] = round(t["Answered"] / t["Total Calls"] * 100, 1) if t["Total Calls"] else 0.0
    return t


OUTSIDE_COLS = ["Period", "Total Calls"] + OUTCOMES + ["Answer %"]


def outside_table(events: pd.DataFrame) -> pd.DataFrame:
    """Calls outside office hours by group (After hours, Weekend, Excluded day) and a total row."""
    rows = []
    for label in OUTSIDE:
        sub = events[events["Period"] == label]
        vc = sub["Outcome"].value_counts()
        rows.append({"Period": label, "Total Calls": len(sub), **{c: int(vc.get(c, 0)) for c in OUTCOMES}})
    out = pd.DataFrame(rows, columns=OUTSIDE_COLS[:-1])
    out = pd.concat([out, pd.DataFrame([{"Period": "Total outside office hours",
                                         **out.drop(columns="Period").sum().to_dict()}])], ignore_index=True)
    out[OUTSIDE_COLS[1:-1]] = out[OUTSIDE_COLS[1:-1]].astype(int)
    out["Answer %"] = (out["Answered"] / out["Total Calls"].where(out["Total Calls"] > 0) * 100).round(1).fillna(0.0)
    return out[OUTSIDE_COLS]


@dataclass
class ExtAudit:
    files_processed: int = 0
    files_failed: list[tuple[str, str]] = field(default_factory=list)
    total_rows: int = 0
    total_calls: int = 0
    duplicates_removed: int = 0
    per_file: list[dict] = field(default_factory=list)
    first_call: pd.Timestamp | None = None
    last_call: pd.Timestamp | None = None
    call_days: dict = field(default_factory=dict)
    outside_period: int = 0
    overlaps: list = field(default_factory=list)
    overlap_calls: int = 0


@dataclass
class ExtStatsResult:
    department: str
    extensions: list[str]
    tables: dict[str, pd.DataFrame]  # "All calls", "Inbound", "Outbound" (+ "Internal")
    audit: ExtAudit
    period: pd.Period
    include_internal: bool
    warnings: list[str] = field(default_factory=list)
    # Set only when office hours were applied: the rules used, calls outside office hours by group, and the same
    # calls by extension. The main tables then count office-hours calls only.
    office_hours: OfficeHours | None = None
    outside: pd.DataFrame | None = None
    outside_by_extension: pd.DataFrame | None = None

    @property
    def all_calls(self) -> pd.DataFrame:
        return self.tables["All calls"]

    @property
    def hours_note(self) -> str:
        return self.office_hours.describe(self.period) if self.office_hours else ""


def process_ext_stats(
    files: Iterable[tuple[str, bytes]],
    mapping: pd.DataFrame,
    names: pd.DataFrame | None,
    department: str,
    include_internal: bool = False,
    hide_idle: bool = False,
    dedupe: bool = True,
    period: pd.Period | None = None,
    dayfirst: bool = True,
    progress: ProgressFn | None = None,
    office_hours: OfficeHours | None = None,
) -> ExtStatsResult:
    """office_hours: when given, the main tables count only call legs whose call started inside office hours;
    legs after hours, on weekends and on excluded days are counted separately (result.outside)."""
    files = list(files)
    if not files:
        raise EmptyUploadError("No files were uploaded. Please upload at least one XLS/XLSX file.")
    # Extensions that belong to the department at some point in the report month. An extension that moved in or
    # out part-way through the month is listed, but only its legs from the dates it was in the department count.
    extensions = sorted(department_extensions(mapping, department, period), key=_ext_sort_key)
    if not extensions:
        raise ReportError(f"No extensions are mapped to '{department}' in the department mapping file.")
    dept_exts = set(extensions)
    report = progress or (lambda *_: None)

    audit = ExtAudit()
    frames = []
    firsts: list = []
    ids: list = []
    n = len(files)
    for i, (name, data) in enumerate(files):
        report(0.05 + 0.70 * i / n, f"Reading {name} ({i + 1}/{n})")
        try:
            res = read_ext_file(name, data, dept_exts, include_internal, dayfirst)
        except (InvalidFileError, MissingColumnsError) as exc:
            audit.files_failed.append((name, str(exc)))
            continue
        audit.files_processed += 1
        audit.total_rows += res.rows
        audit.total_calls += res.calls
        audit.first_call = min([t for t in (audit.first_call, res.first_call) if t is not None], default=None)
        audit.last_call = max([t for t in (audit.last_call, res.last_call) if t is not None], default=None)
        audit.call_days = merge_days(audit.call_days, res.call_days)
        audit.per_file.append({"File": name, "Rows": res.rows, "Calls": res.calls, "Events": len(res.events)})
        frames.append(res.events)
        firsts.append(res.first_call)
        ids.append((file_label(i + 1, name), res.call_ids))
    if not audit.files_processed:
        raise InvalidFileError("None of the uploaded files could be processed. "
                               + " ".join(m for _, m in audit.files_failed))

    report(0.80, "Calculating statistics")
    audit.overlaps, audit.overlap_calls = find_overlaps(ids)
    events = pd.concat(latest_first(frames, firsts), ignore_index=True)
    if dedupe:
        before = len(events)
        events = events.drop_duplicates(["Call ID", "Leg", "Extension", "Side"])
        audit.duplicates_removed = before - len(events)
    events = events[department_at(mapping, events["Extension"], events["Time"]) == department]  # in dept on that date
    if period is not None:  # the report covers the selected month only (like the PBX report)
        inside = (events["Time"] >= period.start_time) & (events["Time"] <= period.end_time)
        audit.outside_period = int((~inside).sum())
        events = events[inside]
        if events.empty and audit.outside_period:
            raise ReportError(
                f"No calls fall inside {period.strftime('%B %Y')}: the files cover "
                f"{audit.first_call:%d/%m/%Y} to {audit.last_call:%d/%m/%Y}. Check the report month and year.")
    if events.empty:
        raise ReportError(
            f"No calls were found for the extensions of '{department}'. Check that the department's extensions "
            "in department_mapping.xlsx match the CDR files."
        )

    # names: the names file wins; fall back to the name shown in the CDR
    lookup = dict(zip(names["Extension"], names["Name"])) if names is not None else {}
    cdr_names = (events[events["CdrName"] != ""].groupby("Extension")["CdrName"]
                 .agg(lambda s: s.value_counts().index[0]).to_dict())
    full_names = {e: (lookup.get(e) or cdr_names.get(e, "")) for e in extensions}

    outside = outside_by_ext = None
    if office_hours is not None:  # split by the time the call started
        events = events.assign(Period=office_hours.classify(events["Time"]))
        out_events = events[events["Period"] != OFFICE]
        outside = outside_table(out_events)
        outside_by_ext = stats_table(out_events, extensions, full_names)
        events = events[events["Period"] == OFFICE]

    subsets = {"All calls": events, "Inbound": events[events["Direction"] == "Inbound"],
               "Outbound": events[events["Direction"] == "Outbound"]}
    if include_internal:
        subsets["Internal"] = events[events["Direction"] == "Internal"]
    tables = {k: stats_table(v, extensions, full_names) for k, v in subsets.items()}
    shown = extensions
    if hide_idle:  # an extension with calls only outside office hours is still shown
        busy = tables["All calls"].set_index("Extension")["Total Calls"]
        if outside_by_ext is not None:
            busy = busy + outside_by_ext.set_index("Extension")["Total Calls"]
        shown = [e for e in extensions if busy.loc[e] > 0]
        tables = {k: v[v["Extension"].isin(shown)].reset_index(drop=True) for k, v in tables.items()}
        if outside_by_ext is not None:
            outside_by_ext = outside_by_ext[outside_by_ext["Extension"].isin(shown)].reset_index(drop=True)

    if period is None:
        period = pd.Timestamp.now().to_period("M")
    warnings = moves_notice(mapping, period, department)
    cov = coverage_warning(audit.first_call, audit.last_call, audit.call_days, period)
    if cov:
        warnings.append(cov)
    ov = overlap_warning(audit.overlaps, audit.overlap_calls, dedupe)
    if ov:
        warnings.append(ov)
    if audit.outside_period:
        warnings.append(f"{audit.outside_period:,} call leg(s) from calls outside {period.strftime('%B %Y')} were left out.")
    missing = [e for e in extensions if not full_names.get(e)]
    if missing:
        warnings.append(f"{len(missing)} extension(s) have no name in the names file or the CDR: " + ", ".join(missing[:15])
                        + (" ..." if len(missing) > 15 else ""))
    return ExtStatsResult(department=department, extensions=shown, tables=tables, audit=audit, period=period,
                          include_internal=include_internal, warnings=warnings, office_hours=office_hours,
                          outside=outside, outside_by_extension=outside_by_ext)
