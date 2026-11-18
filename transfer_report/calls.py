"""Department call list: outbound numbers dialled and inbound callers connected to a department.

Uses the same layered Yeastar export as the transfer report (call summary rows, call-leg rows, event rows):

  * Outbound: summary rows of type Outbound whose "Call From" is one of the department's extensions.
    The number is the dialled "Call To".
  * Inbound: summary rows of type Inbound that have at least one ANSWERED call leg whose "Call To" is one of
    the department's extensions (this includes calls transferred in). The number is the caller ("Call From").

Internal calls are ignored. Everything is read as text (dtype=str).
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

import pandas as pd

from .errors import EmptyUploadError, InvalidFileError, MissingColumnsError, ReportError
from .loader import DT_RE, parse_times
from .mapping import normalise_extension
from .overlap import file_label, find_overlaps, latest_first, overlap_warning

ProgressFn = Callable[[float, str], None]

# The summary, call-leg and event blocks each have their own header row, and Yeastar firmware versions add or
# move columns (e.g. "Second Participant"/"Last Participant" shift Last Status right). Columns are therefore
# located by header name. Columns 0-5 (ID, Call Leg ID, Time, Type, From, To) are the same in every layout.
MAX_COLS = 30
SUMMARY_FIELDS = {"type": "Communication Type", "frm": "Call From", "to": "Call To", "status": "Last Status",
                  "total": "Total Duration", "handling": "Handling Duration"}
LEG_FIELDS = {"frm": "Call From", "to": "Call To", "status": "Status", "talk": "Pure Talk Duration"}
LEG_DEFAULTS = {"Call From": 4, "Call To": 5, "Status": 6, "Pure Talk Duration": 9}  # used only if a file has no leg header row

EXT_RE = r"<([^<>]*)>\s*$"
NAME_RE = r"^(.*?)\s*<[^<>]*>\s*$"
PHONE_RE = r"^\+?\d[\d\s\-()]*$"
SYSTEM_WORDS = ("ivr", "queue", "voicemail", "ring group", "conference", "auto attendant", "parking", "paging")

ALL_DEPARTMENTS = "All departments"
DISPLAY_COLS = ["Date/Time", "Direction", "Extension", "Agent", "Number", "Status",
                "Total Duration", "Talk Time", "Call ID"]
DISPLAY_COLS_ALL = ["Date/Time", "Direction", "Department", "Extension", "Agent", "Number", "Status",
                    "Total Duration", "Talk Time", "Call ID"]
NUMBER_COLS = ["Number", "Calls", "Answered", "Last Call", "Total Talk Time"]
NUMBER_COLS_ALL = NUMBER_COLS + ["Departments"]


# ------------------------------------------------------------------ small helpers
def _status(s: pd.Series) -> pd.Series:
    return s.fillna("").str.strip().str.upper().str.replace("_", " ", regex=False)


def _clean_number(s: pd.Series) -> pd.Series:
    """'+27648990138<0648990138>' -> '0648990138'; '0613864664' -> unchanged."""
    s = s.fillna("").str.strip()
    inner = s.str.extract(EXT_RE, expand=False).fillna("").str.strip()
    name = s.str.extract(NAME_RE, expand=False).fillna("").str.strip()
    out = s.copy()
    out = out.where(~name.str.match(PHONE_RE), name)
    out = out.where(~inner.str.match(PHONE_RE), inner)
    return out


def _agent(s: pd.Series) -> pd.Series:
    return s.fillna("").str.extract(NAME_RE, expand=False).fillna("").str.strip()


def _seconds(s: pd.Series) -> pd.Series:
    return pd.to_timedelta(s, errors="coerce").dt.total_seconds().fillna(0)


def hms(seconds) -> str:
    seconds = int(round(float(seconds)))
    return f"{seconds // 3600:02d}:{seconds % 3600 // 60:02d}:{seconds % 60:02d}"


def coverage_info(c2: pd.Series, dayfirst: bool) -> tuple:
    """(first call, last call, {date: latest call time that day}) from the call times in the data."""
    times = parse_times(c2, dayfirst).dropna()
    times = times[(times.dt.year >= 2000) & (times.dt.year <= 2100)]
    if times.empty:
        return None, None, {}
    days = times.groupby(times.dt.date).max().to_dict()
    return times.min(), times.max(), days


def merge_days(a: dict, b: dict) -> dict:
    out = dict(a)
    for d, t in b.items():
        out[d] = max(t, out[d]) if d in out else t
    return out


def _ranges(days: list) -> str:
    """[28, 29, 30, 31 Oct] -> '28/10/2026 - 31/10/2026'; separate runs are joined with ', '."""
    out, start, prev = [], None, None
    for d in days + [None]:
        if start is None:
            start = prev = d
        elif d is not None and (d - prev).days == 1:
            prev = d
        else:
            out.append(f"{start:%d/%m/%Y}" if start == prev else f"{start:%d/%m/%Y} - {prev:%d/%m/%Y}")
            start = prev = d
    return ", ".join(out)


def coverage_warning(first, last, days: dict, period: pd.Period) -> str | None:
    """Warn when some days of the report month have no calls at all (an incomplete export)."""
    if first is None or last is None:
        return None
    month = [d.date() for d in pd.date_range(period.start_time.normalize(), period.end_time.normalize())]
    missing = [d for d in month if d not in days]
    if not missing:
        return None
    before = [t for d, t in days.items() if d < missing[0]]
    tail = f" The last call before the first gap is at {max(before):%d/%m/%Y %H:%M}." if before else ""
    return (f"No calls were found on {_ranges(missing)} (the files cover {first:%d/%m/%Y %H:%M} to "
            f"{last:%d/%m/%Y %H:%M}).{tail} Figures for {period.strftime('%B %Y')} are therefore incomplete and will "
            "be lower than a PBX report for the whole month. Export the missing days and generate again.")


def _join_unique(values) -> str:
    return "; ".join(dict.fromkeys(v for v in values if v))


# ------------------------------------------------------------------ reading
@dataclass
class CallsFile:
    name: str
    rows: int
    calls: int
    outbound: pd.DataFrame
    inbound: pd.DataFrame
    seen: pd.DataFrame  # extensions seen on these calls that are not in the mapping
    first_call: pd.Timestamp | None = None
    last_call: pd.Timestamp | None = None
    call_days: dict = field(default_factory=dict)
    call_ids: frozenset = field(default_factory=frozenset)


def _read(name: str, data: bytes) -> pd.DataFrame:
    engines = ["xlrd", "openpyxl"] if Path(name).suffix.lower() == ".xls" else ["openpyxl", "xlrd"]
    for engine in engines:
        try:
            return pd.read_excel(io.BytesIO(data), header=None, dtype=str, engine=engine,
                                 usecols=lambda i: i < MAX_COLS)
        except Exception:  # noqa: BLE001 - try the next engine
            continue
    raise InvalidFileError(
        f"'{name}' could not be read. The file may be corrupted or is not a valid XLS/XLSX export."
    )


def parse_layers(name: str, data: bytes) -> tuple[pd.DataFrame, pd.DataFrame, int]:
    """Read a layered Yeastar CDR export and return (call summary rows, call-leg rows, data row count).

    summary columns: id, c2 (time), type, frm, to, status, total, handling
    leg columns:     id, leg, c2 (leg start), frm, to, status, leg_talk
    Columns are located by header name, so exports with extra summary columns are handled."""
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
    missing = [h for h in SUMMARY_FIELDS.values() if h not in header]
    if missing:
        raise MissingColumnsError(f"'{name}' is missing required column(s): {', '.join(missing)}.")
    leg_header_rows = raw[raw[2].fillna("").str.strip() == "Call Leg Time"]
    if len(leg_header_rows):
        leg_pos = {str(v).strip(): i for i, v in leg_header_rows.iloc[0].items() if pd.notna(v) and str(v).strip()}
    else:
        leg_pos = dict(LEG_DEFAULTS)
    for h in LEG_FIELDS.values():
        leg_pos.setdefault(h, LEG_DEFAULTS.get(h, 6))

    raw = raw[raw[0] != "ID"]
    if raw.empty:
        raise InvalidFileError(f"'{name}' contains no data rows.")
    is_dt = raw[2].fillna("").str.match(DT_RE)
    has_leg = raw[1].notna()
    summ_raw, legs_raw = raw[is_dt & ~has_leg], raw[is_dt & has_leg]

    summ = pd.DataFrame({"id": summ_raw[0], "c2": summ_raw[2],
                         **{k: summ_raw[header[h]] for k, h in SUMMARY_FIELDS.items()}})
    legs = pd.DataFrame({"id": legs_raw[0], "leg": legs_raw[1], "c2": legs_raw[2],
                         **{k: legs_raw[leg_pos[h]] for k, h in LEG_FIELDS.items()}})
    legs = legs.rename(columns={"talk": "leg_talk"})
    return summ, legs, len(raw)


def read_calls_file(name: str, data: bytes, dept_exts: set[str], all_exts: set[str], dayfirst: bool,
                    ext_dept: dict[str, str] | None = None) -> CallsFile:
    ext_dept = ext_dept or {}
    if not data:
        raise InvalidFileError(f"'{name}' is empty.")
    summ, legs, rows = parse_layers(name, data)
    ctype = summ["type"].fillna("").str.strip().str.lower()

    # ---------------- outbound: department extension dialled out
    ob_all = summ[ctype == "outbound"].copy()
    ob_all["ext"] = normalise_extension(ob_all["frm"].str.extract(EXT_RE, expand=False))
    ob = ob_all[ob_all["ext"].isin(dept_exts)]
    outbound = pd.DataFrame({
        "Date/Time": parse_times(ob["c2"], dayfirst),
        "Direction": "Outbound",
        "Department": ob["ext"].map(ext_dept).fillna(""),
        "Extension": ob["ext"],
        "Agent": _agent(ob["frm"]),
        "Number": _clean_number(ob["to"]),
        "Status": _status(ob["status"]),
        "Total Duration": ob["total"].fillna(""),
        "Talk Time": ob["handling"].fillna(""),
        "Call ID": ob["id"],
        "_talk_s": _seconds(ob["handling"]),
    })

    # ---------------- inbound: answered by a department extension
    ib = summ[ctype == "inbound"]
    lg = legs[legs["id"].isin(ib["id"])].copy()
    lg["ext"] = normalise_extension(lg["to"].str.extract(EXT_RE, expand=False))
    lg["agent"] = _agent(lg["to"])
    lg = lg[(_status(lg["status"]) == "ANSWERED") & lg["ext"].notna()]
    lg["talk_s"] = _seconds(lg["leg_talk"])
    mine = lg[lg["ext"].isin(dept_exts)].copy()
    # A queue/IVR leg overlaps in time with the agent leg that took the call, so talk time is summed over agent
    # legs when there are any, and over the system legs only when there are none.
    mine["sys"] = mine["agent"].str.lower().str.contains("|".join(SYSTEM_WORDS), regex=True)
    mine["talk_agent"] = mine["talk_s"].where(~mine["sys"], 0)
    mine["dept"] = mine["ext"].map(ext_dept).fillna("")
    g = mine.groupby("id", sort=False)
    agg = pd.DataFrame({
        "Extension": g["ext"].agg(_join_unique), "Agent": g["agent"].agg(_join_unique),
        "Department": g["dept"].agg(_join_unique),
        "_all": g["talk_s"].sum(), "_agent": g["talk_agent"].sum(), "_has_agent": g["sys"].agg(lambda x: (~x).any()),
    })
    agg["_talk_s"] = agg["_agent"].where(agg["_has_agent"], agg["_all"])
    agg = agg[["Extension", "Agent", "Department", "_talk_s"]]
    ib = ib.merge(agg, left_on="id", right_index=True, how="inner")
    inbound = pd.DataFrame({
        "Date/Time": parse_times(ib["c2"], dayfirst),
        "Direction": "Inbound",
        "Department": ib["Department"],
        "Extension": ib["Extension"],
        "Agent": ib["Agent"],
        "Number": _clean_number(ib["frm"]),
        "Status": "ANSWERED",
        "Total Duration": ib["total"].fillna(""),
        "Talk Time": ib["_talk_s"].map(hms),
        "Call ID": ib["id"],
        "_talk_s": ib["_talk_s"],
    })

    # ---------------- extensions on these calls that are not in the mapping (people, not IVRs/queues)
    seen = []
    o = ob_all[~ob_all["ext"].isin(all_exts) & ob_all["ext"].notna()]
    if len(o):
        seen.append(pd.DataFrame({"Extension": o["ext"], "Name": _agent(o["frm"]), "Seen as": "Outbound caller",
                                  "Call ID": o["id"]}))
    u = lg[~lg["ext"].isin(all_exts)]
    u = u[~u["agent"].str.lower().str.contains("|".join(SYSTEM_WORDS), regex=True)]
    if len(u):
        seen.append(pd.DataFrame({"Extension": u["ext"], "Name": u["agent"], "Seen as": "Inbound answered",
                                  "Call ID": u["id"]}))
    seen_df = (pd.concat(seen, ignore_index=True) if seen
               else pd.DataFrame(columns=["Extension", "Name", "Seen as", "Call ID"]))
    first, last, days = coverage_info(summ["c2"], dayfirst)
    return CallsFile(name, rows, len(summ), outbound.reset_index(drop=True),
                     inbound.reset_index(drop=True), seen_df, first, last, days, frozenset(summ["id"]))


# ------------------------------------------------------------------ summaries
def _join_departments(values) -> str:
    return "; ".join(sorted({d for v in values for d in str(v).split("; ") if d}))


def number_summary(df: pd.DataFrame, with_departments: bool = False) -> pd.DataFrame:
    """One row per distinct number: calls, answered calls, last call, total talk time (and, when
    with_departments is set, which departments dialled / received it)."""
    cols = NUMBER_COLS_ALL if with_departments else NUMBER_COLS
    if df.empty:
        return pd.DataFrame(columns=cols)
    agg = {"Calls": ("Number", "size"),
           "Answered": ("Status", lambda s: int((s == "ANSWERED").sum())),
           "Last Call": ("Date/Time", "max"), "_t": ("_talk_s", "sum")}
    if with_departments:
        agg["Departments"] = ("Department", _join_departments)
    g = df.groupby("Number", sort=False).agg(**agg).reset_index()
    g["Total Talk Time"] = g["_t"].map(hms)
    g = g.drop(columns="_t").sort_values(["Calls", "Number"], ascending=[False, True], kind="stable")
    return g[cols].reset_index(drop=True)


def department_summary(outbound: pd.DataFrame, inbound: pd.DataFrame) -> pd.DataFrame:
    """Per-department totals (used for the 'All departments' report)."""
    ob = outbound.assign(_ans=(outbound["Status"] == "ANSWERED").astype(int))
    ob = ob.groupby("Department").agg(**{
        "Outbound calls": ("Call ID", "size"), "Outbound answered": ("_ans", "sum"),
        "Unique numbers dialled": ("Number", "nunique"), "_ot": ("_talk_s", "sum")})
    ib = inbound.assign(Department=inbound["Department"].str.split("; ")).explode("Department")
    ib = ib[ib["Department"].notna() & (ib["Department"] != "")]
    ib = ib.groupby("Department").agg(**{
        "Inbound calls connected": ("Call ID", "size"),
        "Unique inbound callers": ("Number", "nunique"), "_it": ("_talk_s", "sum")})
    df = ob.join(ib, how="outer").fillna(0)
    for c in ("Outbound calls", "Outbound answered", "Unique numbers dialled", "Inbound calls connected",
              "Unique inbound callers"):
        df[c] = df[c].astype(int)
    df["Outbound talk time"] = df["_ot"].map(hms)
    df["Inbound talk time"] = df["_it"].map(hms)
    df = df.assign(_n=df["Outbound calls"] + df["Inbound calls connected"]).sort_values("_n", ascending=False)
    return df.drop(columns=["_ot", "_it", "_n"]).rename_axis("Department").reset_index()


@dataclass
class CallsAudit:
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


@dataclass
class CallsResult:
    department: str
    extensions: list[str]
    outbound: pd.DataFrame
    inbound: pd.DataFrame
    outbound_numbers: pd.DataFrame
    inbound_numbers: pd.DataFrame
    unmapped: pd.DataFrame
    audit: CallsAudit
    period: pd.Period
    warnings: list[str] = field(default_factory=list)
    all_departments: bool = False
    by_department: pd.DataFrame = field(default_factory=pd.DataFrame)

    @property
    def display_cols(self) -> list[str]:
        return DISPLAY_COLS_ALL if self.all_departments else DISPLAY_COLS

    @property
    def number_cols(self) -> list[str]:
        return NUMBER_COLS_ALL if self.all_departments else NUMBER_COLS

    @property
    def scope_text(self) -> str:
        if self.all_departments:
            return f"{len(self.extensions):,} extension(s) across {len(self.by_department):,} department(s)"
        return f"{len(self.extensions)} extension(s) in {self.department}: " + ", ".join(self.extensions)

    @property
    def stats(self) -> dict:
        ob, ib = self.outbound, self.inbound
        return {
            ("Extensions included" if self.all_departments else "Extensions in department"): len(self.extensions),
            "Outbound calls": len(ob),
            "Outbound answered": int((ob["Status"] == "ANSWERED").sum()),
            "Unique numbers dialled": int(ob["Number"].nunique()),
            "Outbound talk time": hms(ob["_talk_s"].sum()),
            "Inbound calls connected": len(ib),
            "Unique inbound callers": int(ib["Number"].nunique()),
            "Inbound talk time": hms(ib["_talk_s"].sum()),
        }


def process_calls(
    files: Iterable[tuple[str, bytes]],
    mapping: pd.DataFrame,
    department: str,
    dayfirst: bool = True,
    dedupe: bool = True,
    period: pd.Period | None = None,
    progress: ProgressFn | None = None,
) -> CallsResult:
    files = list(files)
    if not files:
        raise EmptyUploadError("No files were uploaded. Please upload at least one XLS/XLSX file.")
    all_mode = department == ALL_DEPARTMENTS
    dept_exts = set(mapping["Extension"]) if all_mode else set(mapping.loc[mapping["Department"] == department, "Extension"])
    ext_dept = dict(zip(mapping["Extension"], mapping["Department"]))
    if not dept_exts:
        raise ReportError(f"No extensions are mapped to '{department}' in the department mapping file.")
    all_exts = set(mapping["Extension"])
    report = progress or (lambda *_: None)

    audit = CallsAudit()
    obs, ibs, seen = [], [], []
    firsts: list = []
    ids: list = []
    n = len(files)
    for i, (name, data) in enumerate(files):
        report(0.05 + 0.65 * i / n, f"Reading {name} ({i + 1}/{n})")
        try:
            res = read_calls_file(name, data, dept_exts, all_exts, dayfirst, ext_dept)
        except (InvalidFileError, MissingColumnsError) as exc:
            audit.files_failed.append((name, str(exc)))
            continue
        audit.files_processed += 1
        audit.total_rows += res.rows
        audit.total_calls += res.calls
        audit.first_call = min([t for t in (audit.first_call, res.first_call) if t is not None], default=None)
        audit.last_call = max([t for t in (audit.last_call, res.last_call) if t is not None], default=None)
        audit.call_days = merge_days(audit.call_days, res.call_days)
        audit.per_file.append({"File": name, "Rows": res.rows, "Calls": res.calls,
                               "Outbound": len(res.outbound), "Inbound": len(res.inbound)})
        obs.append(res.outbound), ibs.append(res.inbound), seen.append(res.seen)
        firsts.append(res.first_call)
        ids.append((file_label(i + 1, name), res.call_ids))
    if not audit.files_processed:
        details = " ".join(m for _, m in audit.files_failed)
        raise InvalidFileError(f"None of the uploaded files could be processed. {details}")

    report(0.78, "Combining files")
    audit.overlaps, audit.overlap_calls = find_overlaps(ids)
    outbound = pd.concat(latest_first(obs, firsts), ignore_index=True)
    inbound = pd.concat(latest_first(ibs, firsts), ignore_index=True)
    if dedupe:
        before = len(outbound) + len(inbound)
        outbound = outbound.drop_duplicates(["Call ID", "Extension"])
        inbound = inbound.drop_duplicates(["Call ID", "Extension"])
        audit.duplicates_removed = before - len(outbound) - len(inbound)
    outbound = outbound.sort_values("Date/Time", kind="stable").reset_index(drop=True)
    inbound = inbound.sort_values("Date/Time", kind="stable").reset_index(drop=True)
    if outbound.empty and inbound.empty:
        raise ReportError(
            f"No outbound or connected inbound calls were found for '{department}'. Check that the "
            "department's extensions are in department_mapping.xlsx and match the CDR files."
        )

    report(0.88, "Summarising numbers")
    seen_all = (pd.concat(latest_first(seen, firsts), ignore_index=True) if seen
                else pd.DataFrame(columns=["Extension", "Name", "Seen as", "Call ID"]))
    if dedupe and not seen_all.empty:
        seen_all = seen_all.drop_duplicates(["Call ID", "Extension", "Seen as"])
    if seen_all.empty:
        unmapped = pd.DataFrame(columns=["Extension", "Name", "Seen as", "Calls"])
    else:
        unmapped = (seen_all.groupby(["Extension", "Name", "Seen as"], sort=False).size()
                    .reset_index(name="Calls").sort_values("Calls", ascending=False).reset_index(drop=True))

    if period is None:
        dt = pd.concat([outbound["Date/Time"], inbound["Date/Time"]]).dropna()
        dt = dt[(dt.dt.year >= 2000) & (dt.dt.year <= 2100)]
        period = dt.dt.to_period("M").mode().iloc[0] if len(dt) else pd.Timestamp.now().to_period("M")

    warnings = []
    cov = coverage_warning(audit.first_call, audit.last_call, audit.call_days, period)
    if cov:
        warnings.append(cov)
    ov = overlap_warning(audit.overlaps, audit.overlap_calls, dedupe)
    if ov:
        warnings.append(ov)
    if len(unmapped):
        warnings.append(
            f"{len(unmapped):,} extension(s) seen on these calls are not in the department mapping, so any "
            "calls they made or answered cannot be attributed to a department (see the Unmapped extensions tab)."
        )
    return CallsResult(
        department=department, extensions=sorted(dept_exts), outbound=outbound, inbound=inbound,
        outbound_numbers=number_summary(outbound, all_mode), inbound_numbers=number_summary(inbound, all_mode),
        unmapped=unmapped, audit=audit, period=period, warnings=warnings, all_departments=all_mode,
        by_department=department_summary(outbound, inbound) if all_mode else pd.DataFrame(),
    )
