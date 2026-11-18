"""Department mapping loader (department_mapping.xlsx) and template writer."""
from __future__ import annotations

from pathlib import Path
from typing import BinaryIO

import pandas as pd

from .errors import MappingError

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
    ws.append(["Extension", "Department"])
    for dept, exts in SAMPLE_MAPPING.items():
        for ext in exts:
            ws.append([ext, dept])
    for row in ws.iter_rows(min_row=2, max_col=1):
        row[0].number_format = "@"  # keep leading zeros
        row[0].data_type = "s"
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 34

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
            frame = df[["Extension", "Department"]].copy()
            break
    if frame is None:
        raise MappingError(
            "The department mapping file must contain the columns 'Extension' and 'Department'."
        )

    frame = frame.dropna(how="any")
    frame["Extension"] = normalise_extension(frame["Extension"])
    frame["Department"] = frame["Department"].str.strip()
    frame = frame[(frame["Extension"] != "") & (frame["Department"] != "")]
    if frame.empty:
        raise MappingError("The department mapping file contains no usable rows.")

    warnings: list[str] = []
    dupes = frame[frame.duplicated("Extension", keep=False)]
    if not dupes.empty:
        warnings.append(
            f"{dupes['Extension'].nunique()} extension(s) appear more than once in the mapping; "
            "the first entry was used."
        )
        frame = frame.drop_duplicates("Extension", keep="first")
    return frame.reset_index(drop=True), warnings


def list_departments(mapping: pd.DataFrame) -> list[str]:
    """Unique department names in the order they first appear in the mapping file."""
    return mapping["Department"].drop_duplicates().tolist()
