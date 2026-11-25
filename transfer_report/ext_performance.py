"""Extension Performance Detail: every inbound and outbound call leg of a department's extensions.

One row per call leg, in time order, laid out like the "Inbound and Outbound Call Report" workbook:
ID, Time, Consultants, Call From, Call To, Call Duration, Ring Duration, Talk Duration, Communication Type, Status.

Rules (taken from the layered Yeastar export, see calls.parse_layers):

  * Each row is one call-leg row of the export. Time is the leg's own start time and the three durations are the leg's
    Call / Ring / Pure Talk durations.
  * Communication Type is the leg's own type. On an Inbound leg the department extension is the callee ("Call To");
    on an Outbound leg it is the caller ("Call From"). A leg that does not fit that (e.g. an internal call between
    two extensions) is left out and counted in the audit.
  * "Voicemail <name><ext>" legs are kept (they are the extension's own voicemail box); queue, IVR, ring-group and
    similar legs are not extensions and are ignored.
  * The report covers the selected month, by leg start time.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Iterable

import pandas as pd

from .calls import EXT_RE, _agent, _status, coverage_info, coverage_warning, hms, merge_days, parse_layers
from .errors import EmptyUploadError, InvalidFileError, MissingColumnsError, ReportError
from .ext_stats import NON_EXTENSION_WORDS, ExtAudit, _ext_sort_key
from .loader import parse_times
from .mapping import department_at, department_extensions, moves_notice, normalise_extension
from .overlap import file_label, find_overlaps, latest_first, overlap_warning

ProgressFn = Callable[[float, str], None]

DETAIL_COLS = ["ID", "Time", "Consultants", "Call From", "Call To", "Call Duration", "Ring Duration",
               "Talk Duration", "Communication Type", "Status"]
DURATION_COLS = ["Call Duration", "Ring Duration", "Talk Duration"]
_WORK_COLS = DETAIL_COLS + ["Extension", "Leg", "CdrName"]  # the last three are working columns, not output
DIRECTIONS = ("inbound", "outbound")
PREVIEW_ROWS = 5_000


def _durations(s: pd.Series) -> pd.Series:
    return pd.to_timedelta(s.fillna(""), errors="coerce").fillna(pd.Timedelta(0))


# ------------------------------------------------------------------ reading
@dataclass
class PerfFile:
    name: str
    rows: int
    calls: int
    legs: pd.DataFrame
    left_out: int = 0
    first_call: pd.Timestamp | None = None
    last_call: pd.Timestamp | None = None
    call_days: dict = field(default_factory=dict)
    call_ids: frozenset = field(default_factory=frozenset)


def read_perf_file(name: str, data: bytes, dept_exts: set[str], dayfirst: bool = True) -> PerfFile:
    """The inbound / outbound call legs the department's extensions were on in one CDR file."""
    summ, legs, rows = parse_layers(name, data)
    call_type = summ.drop_duplicates("id").set_index("id")["type"].fillna("").str.strip().str.lower()
    lg = legs.copy()
    direction = lg["ltype"].fillna("").str.strip().str.lower()
    lg["direction"] = direction.where(direction != "", lg["id"].map(call_type).fillna(""))
    lg["Time"] = parse_times(lg["c2"], dayfirst)
    pattern = "|".join(NON_EXTENSION_WORDS)

    def side(col: str):
        text = lg[col].fillna("")
        label = _agent(text)
        ext = normalise_extension(text.str.extract(EXT_RE, expand=False))
        return ext, label, ext.isin(dept_exts) & ~label.str.lower().str.contains(pattern, regex=True)

    ext_to, label_to, ok_to = side("to")
    ext_from, label_from, ok_from = side("frm")
    inbound, outbound = lg["direction"] == "inbound", lg["direction"] == "outbound"
    keep = (inbound & ok_to) | (outbound & ok_from)
    left_out = int(((ok_to | ok_from) & ~keep).sum())  # on a department extension, but not an inbound / outbound leg

    sel = lg[keep]
    ext = ext_to.where(inbound, ext_from)[keep]
    label = label_to.where(inbound, label_from)[keep].str.replace(r"^Voicemail\s+", "", regex=True)
    out = pd.DataFrame({
        "ID": sel["id"], "Time": sel["Time"], "Consultants": "",
        "Call From": sel["frm"].fillna(""), "Call To": sel["to"].fillna(""),
        "Call Duration": _durations(sel["dur"]), "Ring Duration": _durations(sel["ring"]),
        "Talk Duration": _durations(sel["leg_talk"]),
        "Communication Type": sel["direction"].str.title(), "Status": _status(sel["status"]),
        "Extension": ext, "Leg": sel["leg"].fillna(""), "CdrName": label,
    })[_WORK_COLS].reset_index(drop=True)
    first, last, days = coverage_info(summ["c2"], dayfirst)
    return PerfFile(name, rows, len(summ), out, left_out, first, last, days, frozenset(summ["id"]))


# ------------------------------------------------------------------ result
@dataclass
class PerfResult:
    department: str
    extensions: list[str]
    detail: pd.DataFrame  # DETAIL_COLS, in time order
    audit: ExtAudit
    period: pd.Period
    left_out: int = 0
    warnings: list[str] = field(default_factory=list)

    @property
    def consultants(self) -> int:
        return int(self.detail["Consultants"].nunique())

    def count(self, direction: str) -> int:
        return int((self.detail["Communication Type"] == direction).sum())


def preview_frame(res: PerfResult, limit: int = PREVIEW_ROWS) -> pd.DataFrame:
    """The detail table with text times and durations, for showing on screen (first `limit` rows)."""
    df = res.detail.head(limit).copy()
    df["Time"] = df["Time"].dt.strftime("%d/%m/%Y %H:%M:%S")
    for c in DURATION_COLS:
        df[c] = df[c].dt.total_seconds().map(hms)
    return df


def process_ext_performance(
    files: Iterable[tuple[str, bytes]],
    mapping: pd.DataFrame,
    names: pd.DataFrame | None,
    department: str,
    dedupe: bool = True,
    period: pd.Period | None = None,
    dayfirst: bool = True,
    progress: ProgressFn | None = None,
) -> PerfResult:
    files = list(files)
    if not files:
        raise EmptyUploadError("No files were uploaded. Please upload at least one XLS/XLSX file.")
    # Extensions that belong to the department at some point in the report month; a leg only counts if its extension
    # was in the department on the date of the leg (a mid-month move splits the extension's legs by date).
    extensions = sorted(department_extensions(mapping, department, period), key=_ext_sort_key)
    if not extensions:
        raise ReportError(f"No extensions are mapped to '{department}' in the department mapping file.")
    dept_exts = set(extensions)
    report = progress or (lambda *_: None)

    audit = ExtAudit()
    frames, firsts, ids = [], [], []
    left_out = 0
    n = len(files)
    for i, (name, data) in enumerate(files):
        report(0.05 + 0.70 * i / n, f"Reading {name} ({i + 1}/{n})")
        try:
            res = read_perf_file(name, data, dept_exts, dayfirst)
        except (InvalidFileError, MissingColumnsError) as exc:
            audit.files_failed.append((name, str(exc)))
            continue
        audit.files_processed += 1
        audit.total_rows += res.rows
        audit.total_calls += res.calls
        audit.first_call = min([t for t in (audit.first_call, res.first_call) if t is not None], default=None)
        audit.last_call = max([t for t in (audit.last_call, res.last_call) if t is not None], default=None)
        audit.call_days = merge_days(audit.call_days, res.call_days)
        audit.per_file.append({"File": name, "Rows": res.rows, "Calls": res.calls, "Events": len(res.legs)})
        left_out += res.left_out
        frames.append(res.legs)
        firsts.append(res.first_call)
        ids.append((file_label(i + 1, name), res.call_ids))
    if not audit.files_processed:
        raise InvalidFileError("None of the uploaded files could be processed. "
                               + " ".join(m for _, m in audit.files_failed))

    report(0.80, "Building the call list")
    audit.overlaps, audit.overlap_calls = find_overlaps(ids)
    legs = pd.concat(latest_first(frames, firsts), ignore_index=True)
    if dedupe:
        before = len(legs)
        legs = legs.drop_duplicates(["ID", "Leg", "Extension"])
        audit.duplicates_removed = before - len(legs)
    legs = legs[department_at(mapping, legs["Extension"], legs["Time"]) == department]  # in dept on that date
    if period is not None:  # the report covers the selected month only
        inside = (legs["Time"] >= period.start_time) & (legs["Time"] <= period.end_time)
        audit.outside_period = int((~inside).sum())
        legs = legs[inside]
        if legs.empty and audit.outside_period:
            raise ReportError(
                f"No calls fall inside {period.strftime('%B %Y')}: the files cover "
                f"{audit.first_call:%d/%m/%Y} to {audit.last_call:%d/%m/%Y}. Check the report month and year.")
    if legs.empty:
        raise ReportError(
            f"No inbound or outbound calls were found for the extensions of '{department}'. Check that the "
            "department's extensions in department_mapping.xlsx match the CDR files.")

    # Consultant name: the names file wins, then the name shown in the CDR, then the extension number itself.
    lookup = dict(zip(names["Extension"], names["Name"])) if names is not None else {}
    cdr_names = (legs[legs["CdrName"] != ""].groupby("Extension")["CdrName"]
                 .agg(lambda s: s.value_counts().index[0]).to_dict())
    used = sorted(set(legs["Extension"]), key=_ext_sort_key)
    shown = {e: lookup.get(e) or cdr_names.get(e) or e for e in used}
    legs = legs.assign(Consultants=legs["Extension"].map(shown))
    detail = (legs.sort_values("Time", kind="mergesort")[DETAIL_COLS].reset_index(drop=True))

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
        warnings.append(f"{audit.outside_period:,} call leg(s) outside {period.strftime('%B %Y')} were left out.")
    no_name = [e for e in used if not lookup.get(e)]
    if no_name:
        warnings.append(
            f"{len(no_name)} extension(s) are not in the extension names file, so the CDR name or the extension "
            "number is shown instead: " + ", ".join(f"{e} (shown as {shown[e]})" for e in no_name[:15])
            + (" ..." if len(no_name) > 15 else "") + ". Add them to the names file to show proper names.")
    return PerfResult(department=department, extensions=extensions, detail=detail, audit=audit, period=period,
                      left_out=left_out, warnings=warnings)
