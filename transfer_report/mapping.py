"""Department mapping loader (department_mapping.xlsx) and template writer.

The mapping has an optional ``Effective From`` column so an extension can move between departments part-way
through a month. Each row applies from its date until the next row for the same extension takes over; a blank
date means "from the beginning". Helpers below answer "which department was extension X in at time T?" so every
report attributes a call to the department the extension belonged to on the day of that call.
"""
from __future__ import annotations

from pathlib import Path
from typing import BinaryIO

import numpy as np
import pandas as pd

from .errors import MappingError

EFFECTIVE = "Effective From"
_EARLIEST = pd.Timestamp.min  # stands in for a blank Effective From ("applies from the beginning")

SAMPLE_MAPPING = {
    "Reception": ["030", "031", "032"],
    "Customer Care": ["100", "101", "102", "103", "104"],
    "Claims Call Centre": ["180", "181", "182", "183", "184"],
    "Complaints": ["190", "191", "192"],
    "Amendments and Cancellations": ["200", "201", "202"],
}


def normalise_extension(series: pd.Series) -> pd.Series:
    """Trim whitespace and strip a trailing '.0' (Excel numeric artefact)."""
    return series.str.strip().str.replace(r"\.0$", "", regex=True)


def create_template(path: Path | str) -> Path:
    """Write a sample mapping workbook. The sample extensions are placeholders."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Mapping"
    ws.append(["Extension", "Department", EFFECTIVE])  # Effective From is optional: leave it blank unless moving
    for dept, exts in SAMPLE_MAPPING.items():
        for ext in exts:
            ws.append([ext, dept])
    for row in ws.iter_rows(min_row=2, max_col=1):
        row[0].number_format = "@"  # keep leading zeros
        row[0].data_type = "s"
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3, max_row=ws.max_row):
        row[0].number_format = "yyyy-mm-dd"
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 16

    ds = wb.create_sheet("Departments")
    ds.append(["Departments"])
    for dept in SAMPLE_MAPPING:
        ds.append([dept])
    ds.column_dimensions["A"].width = 34
    wb.save(path)
    return Path(path)


def load_mapping(source: Path | str | BinaryIO) -> tuple[pd.DataFrame, list[str]]:
    """Return (DataFrame[Extension, Department], warnings).

    The first sheet containing both an ``Extension`` and a ``Department`` column
    is used, so the third "Departments" list column/sheet is optional.
    """
    if isinstance(source, (str, Path)) and not Path(source).exists():
        raise MappingError(
            f"The department mapping file '{Path(source).name}' was not found. "
            "Place it next to app.py or upload one in the sidebar."
        )
    try:
        sheets = pd.read_excel(source, sheet_name=None, dtype=str, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001
        raise MappingError(
            "The department mapping file could not be read. Make sure it is a valid .xlsx file."
        ) from exc

    frame = None
    for df in sheets.values():
        df.columns = [str(c).strip() for c in df.columns]
        if {"Extension", "Department"}.issubset(df.columns):
            keep = ["Extension", "Department"] + ([EFFECTIVE] if EFFECTIVE in df.columns else [])
            frame = df[keep].copy()
            break
    if frame is None:
        raise MappingError(
            "The department mapping file must contain the columns 'Extension' and 'Department'."
        )
    if EFFECTIVE not in frame.columns:
        frame[EFFECTIVE] = np.nan

    frame = frame.dropna(subset=["Extension", "Department"])  # a blank Effective From is allowed
    frame["Extension"] = normalise_extension(frame["Extension"])
    frame["Department"] = frame["Department"].str.strip()
    frame = frame[(frame["Extension"] != "") & (frame["Department"] != "")]
    if frame.empty:
        raise MappingError("The department mapping file contains no usable rows.")
    frame[EFFECTIVE] = _parse_effective(frame)

    warnings: list[str] = []
    dupes = frame[frame.duplicated(["Extension", EFFECTIVE], keep=False)]
    if not dupes.empty:
        warnings.append(
            f"{dupes['Extension'].nunique()} extension(s) appear more than once in the mapping with the same "
            f"'{EFFECTIVE}' date (or no date); the first entry was used. To record a move between departments, "
            f"add a second row for the extension with the date the move took effect."
        )
        frame = frame.drop_duplicates(["Extension", EFFECTIVE], keep="first")
    return frame.reset_index(drop=True), warnings


def _parse_effective(frame: pd.DataFrame) -> pd.Series:
    """The Effective From column as dates (NaT = blank). Excel dates arrive as text like '2026-10-22 00:00:00'."""
    text = pd.Series([None if pd.isna(v) or str(v).strip() == "" else str(v).strip() for v in frame[EFFECTIVE]],
                     index=frame.index, dtype="object")
    parsed = pd.to_datetime(text, errors="coerce", format="mixed", dayfirst=True)
    bad = frame.loc[text.notna() & parsed.isna()]
    if not bad.empty:
        examples = "; ".join(f"'{text[i]}' (extension {frame.at[i, 'Extension']})" for i in bad.index[:5])
        raise MappingError(
            f"The '{EFFECTIVE}' column has value(s) that are not valid dates: {examples}"
            + (" ..." if len(bad) > 5 else "")
            + ". Use a date such as 2026-10-22, or leave the cell blank."
        )
    return parsed.dt.normalize().astype("datetime64[ns]")


def _effective(mapping: pd.DataFrame) -> pd.Series:
    """Effective From as datetime64[ns] (all NaT when the mapping has no such column, e.g. a hand-built frame)."""
    if EFFECTIVE not in mapping.columns:
        return pd.Series(pd.NaT, index=mapping.index, dtype="datetime64[ns]")
    return pd.to_datetime(mapping[EFFECTIVE]).astype("datetime64[ns]")


def has_dated_rows(mapping: pd.DataFrame) -> bool:
    return bool(_effective(mapping).notna().any())


def department_at(mapping: pd.DataFrame, extensions: pd.Series, when: pd.Series | None = None) -> pd.Series:
    """Department of each extension at the time given for it (NaN where the extension is not mapped then).

    A row applies from its ``Effective From`` date (00:00) until the next row for the same extension starts; a
    blank date means from the beginning. A missing time uses the extension's latest row. With no dated rows
    this is a plain extension -> department lookup, exactly as before dates were supported.
    """
    dates = _effective(mapping)
    if not dates.notna().any():
        return extensions.map(dict(zip(mapping["Extension"], mapping["Department"])))

    rows = pd.DataFrame({"Extension": mapping["Extension"].astype(object),
                         "Department": mapping["Department"].astype(object),
                         "_from": dates.fillna(_EARLIEST)}).sort_values("_from", kind="stable")
    latest = dict(zip(rows["Extension"], rows["Department"]))  # last assignment = latest effective row
    ext = extensions.astype(object)
    out = ext.map(latest)  # default (also used when the time is missing)
    if when is None:
        return out

    left = pd.DataFrame({"Extension": ext.to_numpy(), "_t": pd.to_datetime(when).astype("datetime64[ns]").to_numpy(),
                         "_pos": np.arange(len(ext))})
    left = left[left["Extension"].notna() & left["_t"].notna()].sort_values("_t", kind="stable")
    if left.empty:
        return out
    left["Extension"] = left["Extension"].astype(object)
    hit = pd.merge_asof(left, rows, left_on="_t", right_on="_from", by="Extension", direction="backward")
    values = out.to_numpy(dtype=object, copy=True)
    values[hit["_pos"].to_numpy()] = hit["Department"].to_numpy(dtype=object)  # NaN: no row applies yet
    return pd.Series(values, index=extensions.index, dtype=object)


def department_extensions(mapping: pd.DataFrame, department: str, period: pd.Period | None = None) -> set[str]:
    """Extensions that belong to ``department``: at any time when no period is given, otherwise at any point
    during that month (so an extension that has moved out before the month starts is not listed)."""
    own = mapping["Department"] == department
    dates = _effective(mapping)
    if period is None or not dates.notna().any():
        return set(mapping.loc[own, "Extension"])
    rows = pd.DataFrame({"Extension": mapping["Extension"], "Department": mapping["Department"],
                         "_from": dates.fillna(_EARLIEST)}).sort_values(["Extension", "_from"], kind="stable")
    rows["_to"] = rows.groupby("Extension")["_from"].shift(-1)  # NaT = still applies
    inside = ((rows["Department"] == department) & (rows["_from"] <= period.end_time)
              & (rows["_to"].isna() | (rows["_to"] > period.start_time)))
    return set(rows.loc[inside, "Extension"])


def moves_notice(mapping: pd.DataFrame, period: pd.Period | None, department: str | None = None) -> list[str]:
    """A one-line notice (or none) naming the department changes that take effect inside the report month."""
    dates = _effective(mapping)
    if period is None or not dates.notna().any():
        return []
    rows = pd.DataFrame({"Extension": mapping["Extension"], "Department": mapping["Department"],
                         "_from": dates.fillna(_EARLIEST)}).sort_values(["Extension", "_from"], kind="stable")
    moves = []
    for ext, grp in rows.groupby("Extension", sort=False):
        prev = None
        for _, r in grp.iterrows():
            if (prev is not None and r["Department"] != prev and period.start_time <= r["_from"] <= period.end_time
                    and (department is None or department in (prev, r["Department"]))):
                moves.append((r["_from"], ext, prev, r["Department"]))
            prev = r["Department"]
    if not moves:
        return []
    moves.sort(key=lambda m: (m[0], m[1]))
    shown = "; ".join(f"ext {e} {a} to {b} on {d:%d/%m/%Y}" for d, e, a, b in moves[:10])
    more = f" (and {len(moves) - 10} more)" if len(moves) > 10 else ""
    return [f"Department change(s) during {period.strftime('%B %Y')}: {shown}{more}. Calls are attributed to the "
            "department the extension was in on the date of each call, so these extensions' figures are split "
            "between the two departments."]


def list_departments(mapping: pd.DataFrame) -> list[str]:
    """Unique department names in the order they first appear in the mapping file."""
    return mapping["Department"].drop_duplicates().tolist()
