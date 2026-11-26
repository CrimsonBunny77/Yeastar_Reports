"""Office hours: which calls fall inside office hours, after hours, on a weekend or on an excluded day.

Used by the Extension call stats and Queue overflow reports. When office hours are applied, only calls inside
office hours are counted in the main tables; the rest are counted separately so the report still shows them.

A call's period is decided by its start time (for queue visits, the time the call entered the queue). Office hours
run from ``start`` up to, but not including, ``end`` on working days. If one call fits more than one group it is
counted once, in this order: Excluded day, then Weekend, then After hours.

Excluded days are whole days the office is closed (public holidays, once-off closures). Short days are working days
the office closes earlier than usual: calls after the early closing time count as After hours.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, time, timedelta
from pathlib import Path

import pandas as pd

OFFICE, AFTER, WEEKEND, EXCLUDED = "Office hours", "After hours", "Weekend", "Excluded day"
OUTSIDE = [AFTER, WEEKEND, EXCLUDED]  # the groups shown in the "outside office hours" breakdown
PERIODS = [OFFICE] + OUTSIDE

DEFAULT_START = time(8, 0)
DEFAULT_END = time(16, 30)
WORKDAYS = frozenset(range(5))  # Monday (0) to Friday (4)

CLOSED, SHORT = "Closed all day", "Short day"
DAY_TYPES = [CLOSED, SHORT]
DAY_COLS = ["Date", "Type", "Closes At", "Reason"]


# ------------------------------------------------------------------ rules
@dataclass(frozen=True)
class OfficeHours:
    start: time = DEFAULT_START
    end: time = DEFAULT_END
    excluded: dict = field(default_factory=dict)  # date -> reason (closed all day)
    short_days: dict = field(default_factory=dict)  # date -> (closing time, reason)
    workdays: frozenset = WORKDAYS

    def classify(self, times: pd.Series) -> pd.Series:
        """One of OFFICE / AFTER / WEEKEND / EXCLUDED per timestamp. An unreadable time is left in Office hours
        (such calls are already reported elsewhere and are never silently moved out of the counts)."""
        t = pd.to_datetime(times)
        days = t.dt.date
        secs = t.dt.hour * 3600 + t.dt.minute * 60 + t.dt.second
        start_s = _secs(self.start)
        end_s = pd.Series(_secs(self.end), index=t.index)
        if self.short_days:
            closes = days.map({d: _secs(c) for d, (c, _) in self.short_days.items()})
            end_s = closes.fillna(end_s).astype(int)
        out = pd.Series(OFFICE, index=t.index, dtype=object)
        out[(secs < start_s) | (secs >= end_s)] = AFTER
        out[~t.dt.weekday.isin(self.workdays)] = WEEKEND
        out[days.isin(set(self.excluded))] = EXCLUDED
        out[t.isna()] = OFFICE
        return out

    def describe(self, period: pd.Period | None = None) -> str:
        """Plain-English summary of the rules, listing only the excluded and short days in the report month."""
        text = (f"Office hours are {self.start:%H:%M} to {self.end:%H:%M}, Monday to Friday (a call at "
                f"{self.end:%H:%M} or later counts as after hours). Calls are placed by the time they started "
                "(queue visits by the time they entered the queue).")
        excl, short = self.days_in(period)
        if excl:
            text += " Excluded days: " + "; ".join(f"{d:%a %d %b %Y}" + (f" ({r})" if r else "") for d, r in excl) + "."
        if short:
            text += " Short days: " + "; ".join(f"{d:%a %d %b %Y} closed at {c:%H:%M}" + (f" ({r})" if r else "")
                                              for d, c, r in short) + "."
        return text

    def days_in(self, period: pd.Period | None = None) -> tuple[list, list]:
        """(excluded days [(date, reason)], short days [(date, closes, reason)]) inside the report month."""
        inside = (lambda d: True) if period is None else (lambda d: d.year == period.year and d.month == period.month)
        excl = sorted((d, r) for d, r in self.excluded.items() if inside(d))
        short = sorted((d, c, r) for d, (c, r) in self.short_days.items() if inside(d) and d not in self.excluded)
        return excl, short


def _secs(t: time) -> int:
    return t.hour * 3600 + t.minute * 60 + t.second


def period_counts(labels: pd.Series, keys: pd.Series | None = None) -> dict:
    """Count of rows (or of distinct keys, when given) in each period group."""
    if keys is None:
        vc = labels.value_counts()
        return {p: int(vc.get(p, 0)) for p in PERIODS}
    return {p: int(keys[labels == p].nunique()) for p in PERIODS}


# ------------------------------------------------------------------ South African public holidays
def easter_sunday(year: int) -> date:
    """Gregorian Easter Sunday (anonymous Gregorian algorithm)."""
    a, b, c = year % 19, year // 100, year % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7  # noqa: E741
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = (h + l - 7 * m + 114) % 31 + 1
    return date(year, month, day)


def sa_public_holidays(year: int) -> list[tuple[date, str]]:
    """South Africa's public holidays under the Public Holidays Act, including the Monday off when a holiday falls
    on a Sunday. Once-off holidays (for example election days) are not included; add them by hand."""
    easter = easter_sunday(year)
    fixed = [((1, 1), "New Year's Day"), ((3, 21), "Human Rights Day"), ((4, 27), "Freedom Day"),
             ((5, 1), "Workers' Day"), ((6, 16), "Youth Day"), ((8, 9), "National Women's Day"),
             ((9, 24), "Heritage Day"), ((12, 16), "Day of Reconciliation"), ((12, 25), "Christmas Day"),
             ((12, 26), "Day of Goodwill")]
    days = {date(year, m, d): name for (m, d), name in fixed}
    days[easter - timedelta(days=2)] = "Good Friday"
    days[easter + timedelta(days=1)] = "Family Day"
    for d, name in list(days.items()):
        if d.weekday() == 6:  # Sunday -> the following Monday is a public holiday
            monday = d + timedelta(days=1)
            while monday in days:
                monday += timedelta(days=1)
            days[monday] = f"{name} (observed)"
    return sorted(days.items())


# ------------------------------------------------------------------ the editable list of excluded / short days
def default_days(year: int) -> pd.DataFrame:
    return pd.DataFrame([{"Date": d, "Type": CLOSED, "Closes At": None, "Reason": name}
                         for d, name in sa_public_holidays(year)], columns=DAY_COLS)


def clean_days(df: pd.DataFrame | None) -> pd.DataFrame:
    """Tidy a days table (from the editor or a file): real dates, a known type, a time or None, text reasons."""
    if df is None or df.empty:
        return pd.DataFrame(columns=DAY_COLS)
    df = df.reindex(columns=DAY_COLS).copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    df = df[df["Date"].notna()]
    df["Type"] = df["Type"].where(df["Type"].isin(DAY_TYPES), CLOSED)
    df["Closes At"] = df["Closes At"].map(_as_time)
    df["Reason"] = df["Reason"].fillna("").astype(str).str.strip()
    return df.drop_duplicates("Date", keep="last").sort_values("Date").reset_index(drop=True)


def _as_time(v) -> time | None:
    if v is None or (not isinstance(v, time) and pd.isna(v)):
        return None
    if isinstance(v, time):
        return v.replace(microsecond=0)
    if isinstance(v, pd.Timestamp):
        return v.time().replace(microsecond=0)
    try:
        return pd.to_datetime(str(v).strip(), format="mixed").time().replace(microsecond=0)
    except (ValueError, TypeError):
        return None


def days_for_year(year: int, saved: pd.DataFrame | None) -> pd.DataFrame:
    """The days shown for a year: what was saved for that year, or South Africa's public holidays if nothing has
    been saved for it yet."""
    saved = clean_days(saved)
    mine = saved[saved["Date"].map(lambda d: d.year) == year]
    return mine.reset_index(drop=True) if len(mine) else default_days(year)


def load_days(path: Path) -> pd.DataFrame:
    if not Path(path).exists():
        return pd.DataFrame(columns=DAY_COLS)
    try:
        return clean_days(pd.read_excel(path, dtype=object))
    except Exception:  # noqa: BLE001 - an unreadable file falls back to the defaults
        return pd.DataFrame(columns=DAY_COLS)


def save_days(path: Path, year: int, days: pd.DataFrame) -> None:
    """Replace the saved days of ``year`` (and any date in ``days``) with ``days``; other years are kept."""
    days = clean_days(days)
    saved = load_days(path)
    keep = saved[(saved["Date"].map(lambda d: d.year) != year) & ~saved["Date"].isin(set(days["Date"]))]
    out = clean_days(pd.concat([keep, days], ignore_index=True))
    out = out.assign(**{"Closes At": out["Closes At"].map(lambda t: t.strftime("%H:%M") if t else "")})
    out.to_excel(path, index=False)


def build_office_hours(days: pd.DataFrame, start: time = DEFAULT_START,
                       end: time = DEFAULT_END) -> tuple[OfficeHours, list[str]]:
    """OfficeHours from the days table, plus warnings about rows that could not be used."""
    days = clean_days(days)
    warnings: list[str] = []
    excluded, short = {}, {}
    for r in days.itertuples(index=False):
        if r.Type == SHORT:
            if r[2] is None:
                warnings.append(f"Short day {r.Date:%d/%m/%Y} has no closing time, so normal hours were used.")
                continue
            short[r.Date] = (r[2], r.Reason)
        else:
            excluded[r.Date] = r.Reason
    if end <= start:
        warnings.append("Office hours end at or before they start; check the office hours in the sidebar.")
    return OfficeHours(start=start, end=end, excluded=excluded, short_days=short), warnings
