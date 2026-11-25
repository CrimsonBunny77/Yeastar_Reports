"""Excel export of excluded (unmapped) transfer records."""
from __future__ import annotations

import io

import pandas as pd

MAX_EXCEL_ROWS = 1_000_000


def unmapped_to_xlsx(summary: pd.DataFrame, records: pd.DataFrame) -> bytes:
    """Three sheets: summary by extension, every excluded record, and a ready-to-fill
    'Mapping additions' sheet (Extension + blank Department) to paste into department_mapping.xlsx."""
    additions = summary.loc[summary["Extension"] != "(blank)", ["Extension"]].copy()
    additions["Department"] = ""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Unmapped extensions", index=False)
        records.head(MAX_EXCEL_ROWS).to_excel(xw, sheet_name="Excluded records", index=False)
        additions.to_excel(xw, sheet_name="Mapping additions", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 12), 40)
    return buf.getvalue()
