"""Queue mapping loader (e.g. Que_Mapping_Table.xlsx): queue number/extension -> queue name.

Used by the Queue overflow report. Nothing about the queues is hard-coded: the numbers and names come from this file.
Column headers are matched loosely (for example 'Que Number', 'Queue Number', 'Queue Extension' for the number and
'Queue Name', 'Name', 'Department' for the name). Problems are returned as reviewable issues, not silently dropped.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import BinaryIO

import pandas as pd

from .errors import MappingError
from .mapping import normalise_extension

NUMBER_HEADERS = ("quenumber", "queuenumber", "queno", "queueno", "queextension", "queueextension", "queext",
                  "queueext", "extension", "number", "ext")
NAME_HEADERS = ("quename", "queuename", "name", "department", "description", "title", "label")
MAPPING_COLS = ["Queue Number", "Queue Name"]
ISSUE_COLS = ["Type", "Queue Number", "Detail"]


def _key(header) -> str:
    return re.sub(r"[^a-z0-9]", "", str(header).lower())


def queue_label(number: str, name: str | None) -> str:
    """'Queue 1 - Reception <801>' (or just the number when the queue is not in the mapping)."""
    return f"{name} <{number}>" if name else str(number)


def _find_columns(df: pd.DataFrame) -> tuple[str, str] | None:
    keys = {c: _key(c) for c in df.columns}
    number = next((c for h in NUMBER_HEADERS for c, k in keys.items() if k == h), None)
    name = next((c for h in NAME_HEADERS for c, k in keys.items() if k == h and c != number), None)
    return (number, name) if number is not None and name is not None else None


def load_queue_mapping(source: Path | str | BinaryIO) -> tuple[pd.DataFrame, list[str], pd.DataFrame]:
    """Return (DataFrame[Queue Number, Queue Name], warnings, issues).

    The first sheet with a queue-number column and a queue-name column is used. `issues` lists every row that was
    dropped or looks inconsistent (blank number, duplicate number, non-numeric number, blank name) so it can be
    shown as a reviewable exception in the report.
    """
    if isinstance(source, (str, Path)) and not Path(source).exists():
        raise MappingError(
            f"The queue mapping file '{Path(source).name}' was not found. "
            "Place it next to app.py or upload one in the sidebar."
        )
    try:
        sheets = pd.read_excel(source, sheet_name=None, dtype=str, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001
        raise MappingError("The queue mapping file could not be read. Make sure it is a valid .xlsx file.") from exc

    frame = cols = None
    for df in sheets.values():
        df.columns = [str(c).strip() for c in df.columns]
        cols = _find_columns(df)
        if cols:
            frame = df[list(cols)].copy()
            break
    if frame is None:
        raise MappingError(
            "The queue mapping file must contain a queue number column (for example 'Que Number' or 'Queue Number') "
            "and a queue name column (for example 'Queue Name' or 'Department')."
        )
    frame.columns = MAPPING_COLS
    frame = frame.dropna(how="all")  # completely empty rows are not worth reporting

    issues: list[dict] = []

    def issue(kind: str, number: str, detail: str) -> None:
        issues.append({"Type": kind, "Queue Number": number, "Detail": detail})

    frame["Queue Number"] = normalise_extension(frame["Queue Number"].fillna(""))
    frame["Queue Name"] = frame["Queue Name"].fillna("").str.strip()

    blank_number = frame["Queue Number"] == ""
    for name in frame.loc[blank_number, "Queue Name"]:
        issue("Queue mapping: blank queue number", "", f"Row for '{name}' has no queue number, so it was ignored.")
    frame = frame[~blank_number]

    for number in frame.loc[~frame["Queue Number"].str.fullmatch(r"\d+"), "Queue Number"]:
        issue("Queue mapping: queue number is not numeric", number,
              "Queue numbers in the CDR are digits, so this queue will not match any call. Check the mapping file.")

    unnamed = frame["Queue Name"] == ""
    for number in frame.loc[unnamed, "Queue Number"]:
        issue("Queue mapping: blank queue name", number, "The queue has no name; its number is shown instead.")

    dupes = frame[frame.duplicated("Queue Number", keep="first")]
    for _, row in dupes.iterrows():
        first = frame[frame["Queue Number"] == row["Queue Number"]].iloc[0]["Queue Name"]
        if first == row["Queue Name"]:
            issue("Queue mapping: duplicate queue number", row["Queue Number"],
                  f"'{row['Queue Name']}' is listed more than once; the first entry was used.")
        else:
            issue("Queue mapping: queue number has different names", row["Queue Number"],
                  f"Listed as '{first}' and as '{row['Queue Name']}'; the first entry ('{first}') was used.")
    frame = frame.drop_duplicates("Queue Number", keep="first")

    if frame.empty:
        raise MappingError("The queue mapping file contains no usable rows.")

    issues_df = pd.DataFrame(issues, columns=ISSUE_COLS)
    warnings = ([f"{len(issues_df)} issue(s) found in the queue mapping file; see the Exceptions tab."]
                if len(issues_df) else [])
    return frame.reset_index(drop=True), warnings, issues_df
