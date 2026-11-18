"""Extension name mapping (extension number -> person's name).

Accepts either
  * a simple sheet/CSV with columns ``Extension`` and ``Name``, or
  * a Yeastar user export (CSV/XLSX) with ``First Name``, ``Last Name`` and ``Extension Number``.

Only those columns are read. A Yeastar user export also contains passwords, PINs and e-mail addresses; they are
never loaded into memory by this module and are not written to the mapping file.
"""
from __future__ import annotations

import io
from pathlib import Path
from typing import BinaryIO

import pandas as pd

from .errors import MappingError
from .mapping import normalise_extension

NEEDED = {"Extension", "Name", "First Name", "Last Name", "Extension Number"}
DEFAULT_NAMES_PATH = Path(__file__).resolve().parent.parent / "extension_names.xlsx"


def _read(source, filename: str) -> pd.DataFrame:
    keep = lambda c: str(c).strip().lstrip("\ufeff") in NEEDED  # noqa: E731
    suffix = Path(filename).suffix.lower()
    try:
        if suffix == ".csv":
            raw = source.read() if hasattr(source, "read") else Path(source).read_bytes()
            for enc in ("utf-8-sig", "cp1252"):
                try:
                    return pd.read_csv(io.BytesIO(raw), dtype=str, encoding=enc, usecols=keep)
                except UnicodeDecodeError:
                    continue
            raise ValueError("unreadable text encoding")
        return pd.read_excel(source, dtype=str, engine="openpyxl", usecols=keep)
    except MappingError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise MappingError(
            "The extension names file could not be read. Upload a CSV or XLSX with 'Extension' and 'Name' "
            "columns, or a Yeastar user export."
        ) from exc


def load_extension_names(source: Path | str | BinaryIO, filename: str | None = None) -> pd.DataFrame:
    """Return DataFrame[Extension, Name] (one row per extension)."""
    if isinstance(source, (str, Path)):
        if not Path(source).exists():
            raise MappingError(f"The extension names file '{Path(source).name}' was not found.")
        filename = filename or str(source)
    df = _read(source, filename or "names.xlsx")
    df.columns = [str(c).strip().lstrip("\ufeff") for c in df.columns]

    if {"Extension", "Name"}.issubset(df.columns):
        out = df[["Extension", "Name"]].copy()
    elif {"Extension Number", "First Name"}.issubset(df.columns):
        last = df["Last Name"] if "Last Name" in df.columns else ""
        name = (df["First Name"].fillna("") + " " + (last.fillna("") if hasattr(last, "fillna") else last)).str.strip()
        out = pd.DataFrame({"Extension": df["Extension Number"], "Name": name})
    else:
        raise MappingError(
            "The extension names file needs 'Extension' and 'Name' columns (or a Yeastar export with "
            "'Extension Number', 'First Name' and 'Last Name')."
        )
    out = out.dropna(subset=["Extension"])
    out["Extension"] = normalise_extension(out["Extension"].astype(str))
    out["Name"] = out["Name"].fillna("").astype(str).str.strip()
    out = out[(out["Extension"] != "")].drop_duplicates("Extension", keep="first")
    return out.reset_index(drop=True)


def save_extension_names(names: pd.DataFrame, path: Path | str) -> Path:
    """Write Extension/Name to an .xlsx (extensions stored as text so leading zeros survive)."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Names"
    ws.append(["Extension", "Name"])
    for ext, name in names[["Extension", "Name"]].itertuples(index=False):
        ws.append([ext, name])
    for row in ws.iter_rows(min_row=2, max_col=1):
        row[0].number_format = "@"
        row[0].data_type = "s"
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 40
    wb.save(path)
    return Path(path)
