"""Central configuration constants."""
from __future__ import annotations

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_MAPPING_PATH = BASE_DIR / "department_mapping.xlsx"
DEFAULT_QUEUE_MAPPING_PATH = BASE_DIR / "Que_Mapping_Table.xlsx"  # queue number -> queue name (Queue overflow report)
DEFAULT_CDR_DIR = BASE_DIR / "CDR_Example"  # CDR exports stored with the app, selectable in the Queue overflow report
# Excluded days (public holidays, once-off closures) and short days saved from the sidebar; see hours.py.
DEFAULT_EXCLUDED_DAYS_PATH = BASE_DIR / "excluded_days.xlsx"

# Columns that must exist in every CDR file.
REQUIRED_COLUMNS = ("Communication Type", "Call From", "Last Status")
# Columns used when present (ID is always imported as text, never numeric).
OPTIONAL_COLUMNS = ("ID", "Call Leg ID", "Time")
USED_COLUMNS = frozenset(REQUIRED_COLUMNS + OPTIONAL_COLUMNS)

# Flag name -> normalised Last Status value.
STATUS_FLAGS = {
    "AnsweredFlag": "ANSWERED",
    "NoAnswerFlag": "NO ANSWER",
    "BusyFlag": "BUSY",
    "VoicemailFlag": "VOICEMAIL",
    "AbandonedFlag": "ABANDONED",
}

STATUS_LABELS = {
    "ANSWERED": "Answered", "NO ANSWER": "No Answer", "BUSY": "Busy",
    "VOICEMAIL": "Voicemail", "ABANDONED": "Abandoned",
}

# Heading of the column counting transfers whose status is none of the five above (formerly "Other").
NOT_COMPLETED = "Transfers Not Completed"

# Final metric column order.
METRIC_COLUMNS = [
    "From Department", "To Department", "Transfers", "Answered",
    "No Answer", "Busy", "Voicemail", "Abandoned", NOT_COMPLETED, "Answer %",
]
COUNT_COLUMNS = ["Transfers", "Answered", "No Answer", "Busy", "Voicemail", "Abandoned", NOT_COMPLETED]

# Regexes pulling values out of Call From JSON such as {"from":"031","blind_to":"183"}
def _key_re(key: str) -> str:
    return rf'"{key}"\s*:\s*"?([^",}}\s]+)'

FROM_RE = _key_re("from")
BLIND_RE = _key_re("blind_to")
ATTENDED_RE = _key_re("attended_to")
