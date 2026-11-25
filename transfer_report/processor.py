"""Combine files, map departments, flag statuses and compute transfer metrics."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Iterable

import numpy as np
import pandas as pd

from .config import COUNT_COLUMNS, METRIC_COLUMNS, NOT_COMPLETED, STATUS_FLAGS, STATUS_LABELS
from .errors import EmptyUploadError, InvalidFileError, MissingColumnsError, NoTransfersError, ReportError
from .calls import coverage_warning, merge_days
from .loader import NO_LEG, parse_times, read_cdr_file
from .overlap import file_label, find_overlaps, latest_first, overlap_warning
from .mapping import department_at, moves_notice, normalise_extension

ProgressFn = Callable[[float, str], None]
MATCH_BOTH, MATCH_EITHER = "both", "either"


@dataclass
class Audit:
    files_processed: int = 0
    files_failed: list[tuple[str, str]] = field(default_factory=list)
    total_rows: int = 0
    total_calls: int = 0
    unmatched_status: int = 0
    operate_rows: int = 0
    blind: int = 0
    attended: int = 0
    duplicates_removed: int = 0
    unmapped_removed: int = 0
    internal_removed: int = 0
    deselected_removed: int = 0
    final_records: int = 0
    per_file: list[dict] = field(default_factory=list)
    first_call: object = None
    last_call: object = None
    call_days: dict = field(default_factory=dict)
    overlaps: list = field(default_factory=list)
    overlap_calls: int = 0


@dataclass
class ProcessingResult:
    detail: pd.DataFrame
    audit: Audit
    period: pd.Period
    warnings: list[str] = field(default_factory=list)
    unmapped_summary: pd.DataFrame = field(default_factory=pd.DataFrame)
    unmapped_records: pd.DataFrame = field(default_factory=pd.DataFrame)
    status_summary: pd.DataFrame = field(default_factory=pd.DataFrame)
    scope: list[str] | None = None  # selected departments, or None when all are included
    match: str = MATCH_BOTH


def _squash(series: pd.Series) -> pd.Series:
    """Upper-case and drop everything except letters/digits, so 'No Answer', 'NO_ANSWER' and
    'noanswer' all compare equal."""
    return series.str.upper().str.replace(r"[^A-Z0-9]", "", regex=True)


def add_status_flags(df: pd.DataFrame) -> pd.DataFrame:
    key = _squash(df["Last Status"].fillna(""))
    for flag, value in STATUS_FLAGS.items():
        df[flag] = (key == value.replace(" ", "")).astype("int64")
    df["OtherStatusFlag"] = 1 - df[list(STATUS_FLAGS)].sum(axis=1)
    return df


def summarise_statuses(df: pd.DataFrame) -> pd.DataFrame:
    """Every distinct Last Status value on the final records and which column it is counted under."""
    counts = (df["Last Status"].fillna("").replace("", "(blank)")
              .value_counts().rename_axis("Last Status").reset_index(name="Records"))
    labels = {v.replace(" ", ""): STATUS_LABELS[v] for v in STATUS_LABELS}
    squashed = _squash(counts["Last Status"].where(counts["Last Status"] != "(blank)", ""))
    counts["Counted as"] = squashed.map(labels).fillna(NOT_COMPLETED)
    return counts


def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    grouped = df.groupby(["From Department", "To Department"], sort=True).agg(
        Transfers=("TransferType", "size"),
        Answered=("AnsweredFlag", "sum"),
        **{
            "No Answer": ("NoAnswerFlag", "sum"),
            "Busy": ("BusyFlag", "sum"),
            "Voicemail": ("VoicemailFlag", "sum"),
            "Abandoned": ("AbandonedFlag", "sum"),
            NOT_COMPLETED: ("OtherStatusFlag", "sum"),
        },
    ).reset_index()
    grouped[COUNT_COLUMNS] = grouped[COUNT_COLUMNS].astype("int64")
    grouped["Answer %"] = (grouped["Answered"] / grouped["Transfers"] * 100).round(1)
    return grouped[METRIC_COLUMNS]


def detect_period(frames: pd.DataFrame, dayfirst: bool) -> pd.Period | None:
    """Most common plausible month in the Time column, or None if it can't be determined."""
    if "Time" in frames.columns and frames["Time"].notna().any():
        times = pd.to_datetime(frames["Time"], errors="coerce", dayfirst=dayfirst, format="mixed")
        times = times[(times.dt.year >= 2000) & (times.dt.year <= 2100)]
        periods = times.dropna().dt.to_period("M")
        if not periods.empty:
            return periods.mode().iloc[0]
    return None


def summarise_unmapped(data: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (summary per extension, full list of excluded records)."""
    no_from = data["From Department"].isna()
    no_to = data["To Department"].isna()
    unm = data.loc[no_from | no_to].copy()
    if unm.empty:
        return pd.DataFrame(columns=["Extension", "Unmapped as From", "Unmapped as To", "Total"]), unm

    as_from = data.loc[no_from, "TransferFrom"].fillna("(blank)").value_counts()
    as_to = data.loc[no_to, "TransferTo"].fillna("(blank)").value_counts()
    summary = pd.DataFrame({"Unmapped as From": as_from, "Unmapped as To": as_to}).fillna(0).astype(int)
    summary["Total"] = summary.sum(axis=1)
    summary = (summary.sort_values(["Total"], ascending=False)
               .rename_axis("Extension").reset_index())

    f, t = unm["From Department"].isna(), unm["To Department"].isna()
    unm["Reason"] = np.select([f & t, f], ["From and To not mapped", "From not mapped"], "To not mapped")
    cols = ["ID", "Call Leg ID", "Time", "TransferType", "TransferFrom", "TransferTo",
            "From Department", "To Department", "Last Status", "Reason"]
    return summary, unm[cols].reset_index(drop=True)


def process_files(
    files: Iterable[tuple[str, bytes]],
    mapping: pd.DataFrame,
    dayfirst: bool = True,
    dedupe: bool = True,
    departments: list[str] | None = None,
    match: str = MATCH_BOTH,
    period: pd.Period | None = None,
    progress: ProgressFn | None = None,
) -> ProcessingResult:
    files = list(files)
    if not files:
        raise EmptyUploadError("No files were uploaded. Please upload at least one XLS/XLSX file.")
    if departments is not None and not departments:
        raise ReportError("No departments are selected. Select at least one department to include.")
    report = progress or (lambda *_: None)

    audit = Audit()
    frames: list[pd.DataFrame] = []
    firsts: list = []
    ids: list = []
    n = len(files)
    for i, (name, data) in enumerate(files):
        report(0.05 + 0.60 * i / n, f"Reading {name} ({i + 1}/{n})")
        try:
            res = read_cdr_file(name, data, dayfirst)
        except (InvalidFileError, MissingColumnsError) as exc:
            audit.files_failed.append((name, str(exc)))
            continue
        audit.files_processed += 1
        audit.total_rows += res.rows
        audit.total_calls += res.calls
        audit.first_call = min([t for t in (audit.first_call, res.first_call) if t is not None], default=None)
        audit.last_call = max([t for t in (audit.last_call, res.last_call) if t is not None], default=None)
        audit.call_days = merge_days(audit.call_days, res.call_days)
        audit.unmatched_status += res.unmatched
        audit.operate_rows += res.operate_rows
        audit.blind += res.blind
        audit.attended += res.attended
        audit.per_file.append({"File": name, "Rows": res.rows, "Calls": res.calls, "Blind": res.blind, "Attended": res.attended})
        frames.append(res.frame)
        firsts.append(res.first_call)
        ids.append((file_label(i + 1, name), res.call_ids))

    if not frames:
        details = " ".join(msg for _, msg in audit.files_failed)
        raise InvalidFileError(f"None of the uploaded files could be processed. {details}")

    report(0.70, "Combining files")
    audit.overlaps, audit.overlap_calls = find_overlaps(ids)
    data = pd.concat(latest_first(frames, firsts), ignore_index=True)
    del frames
    if dedupe:
        subset = [c for c in ("ID", "Call Leg ID", "Time", "TransferFrom", "TransferTo") if data[c].notna().any()]
        before = len(data)
        data = data.drop_duplicates(subset=subset)
        audit.duplicates_removed = before - len(data)
        if audit.duplicates_removed:  # keep the audit counts consistent with the de-duplicated data
            audit.blind = int((data["TransferType"] == "Blind").sum())
            audit.attended = len(data) - audit.blind
            audit.unmatched_status = int((data["Last Status"] == NO_LEG).sum())

    report(0.78, "Mapping departments")
    data["TransferFrom"] = normalise_extension(data["TransferFrom"])
    data["TransferTo"] = normalise_extension(data["TransferTo"])
    # Each side is mapped to the department its extension was in at the time of the transfer, so an extension
    # that moved department part-way through the month is attributed correctly on both sides of the move.
    when = parse_times(data["Time"], dayfirst)
    data["From Department"] = department_at(mapping, data["TransferFrom"], when)
    data["To Department"] = department_at(mapping, data["TransferTo"], when)

    unmapped_summary, unmapped_records = summarise_unmapped(data)
    mapped = data.dropna(subset=["From Department", "To Department"])
    audit.unmapped_removed = len(data) - len(mapped)

    external = mapped[mapped["From Department"] != mapped["To Department"]]
    audit.internal_removed = len(mapped) - len(external)
    if external.empty:
        raise NoTransfersError(
            "No inter-department transfers were found. Check that the extensions in "
            "department_mapping.xlsx match those in the CDR files."
        )
    detected = detect_period(external, dayfirst)  # before department filtering
    month_warning = None
    if period is None:
        period = detected or pd.Timestamp.now().to_period("M")
    elif detected is not None and detected != period:
        month_warning = (
            f"The dates in the files mostly fall in {detected.strftime('%B %Y')}, but the report is "
            f"labelled {period.strftime('%B %Y')}. If that is unexpected, check the month you selected "
            "and the date format setting."
        )

    scope = None
    all_depts = set(mapping["Department"])
    if departments is not None and not all_depts.issubset(departments):
        scope = [d for d in mapping["Department"].drop_duplicates() if d in set(departments)]
        sel = set(scope)
        from_in, to_in = external["From Department"].isin(sel), external["To Department"].isin(sel)
        keep = (from_in & to_in) if match == MATCH_BOTH else (from_in | to_in)
        audit.deselected_removed = int((~keep).sum())
        external = external[keep]
        if external.empty:
            raise NoTransfersError(
                "No transfers match the selected departments. Try selecting more departments "
                "or choose 'Either department selected'."
            )
    external = external.copy()
    audit.final_records = len(external)

    report(0.88, "Calculating statistics")
    status_summary = summarise_statuses(external)
    external = add_status_flags(external)
    detail = compute_metrics(external)

    warnings = [month_warning] if month_warning else []
    warnings += moves_notice(mapping, period)
    cov = coverage_warning(audit.first_call, audit.last_call, audit.call_days, period)
    if cov:
        warnings.append(cov)
    ov = overlap_warning(audit.overlaps, audit.overlap_calls, dedupe)
    if ov:
        warnings.append(ov)
    if int(detail["Answered"].sum()) == 0:
        warnings.append(
            "No transfer records were counted as Answered. Open the 'Status values' tab to see which "
            "Last Status values your files contain."
        )
    if audit.unmapped_removed:
        warnings.append(
            f"{audit.unmapped_removed:,} transfer record(s) were excluded because an extension "
            f"is not in the department mapping ({len(unmapped_summary):,} unmapped extension(s) - see the list below)."
        )
    return ProcessingResult(detail=detail, audit=audit, period=period, warnings=warnings,
                            unmapped_summary=unmapped_summary, unmapped_records=unmapped_records,
                            status_summary=status_summary, scope=scope, match=match)
