"""Detect calls that appear in more than one uploaded file (overlapping exports) and describe them."""
from __future__ import annotations

from collections import Counter

import pandas as pd


def file_label(position: int, name: str) -> str:
    short = name if len(name) <= 34 else "…" + name[-33:]
    return f"file {position} ({short})"


def find_overlaps(ids_by_file: list[tuple[str, frozenset]]) -> tuple[list[tuple[str, str, int]], int]:
    """([(label_a, label_b, shared calls), ...], number of distinct calls found in more than one file)."""
    pairs = []
    for i in range(len(ids_by_file)):
        for j in range(i + 1, len(ids_by_file)):
            shared = len(ids_by_file[i][1] & ids_by_file[j][1])
            if shared:
                pairs.append((ids_by_file[i][0], ids_by_file[j][0], shared))
    if not pairs:
        return [], 0
    counts: Counter = Counter()
    for _, ids in ids_by_file:
        counts.update(ids)
    return pairs, sum(1 for n in counts.values() if n > 1)


def latest_first(items: list, firsts: list) -> list:
    """Order per-file results so the file whose data starts latest comes first. After a concat and
    drop_duplicates(keep='first') the later export's copy of a repeated call is the one kept (it is the more
    complete one if the call was still in progress when the earlier file was exported)."""
    key = lambda i: firsts[i] if firsts[i] is not None else pd.Timestamp.min  # noqa: E731
    return [items[i] for i in sorted(range(len(items)), key=key, reverse=True)]


def overlap_note(pairs: list[tuple[str, str, int]]) -> str:
    shown = "; ".join(f"{a} and {b}: {n:,} calls" for a, b, n in pairs[:4])
    return shown + (f"; and {len(pairs) - 4} more pair(s)" if len(pairs) > 4 else "")


def overlap_warning(pairs: list[tuple[str, str, int]], calls: int, removed: bool) -> str | None:
    if not pairs:
        return None
    base = f"{calls:,} call(s) appear in more than one of the uploaded files ({overlap_note(pairs)})."
    if removed:
        return base + " Each call was counted once."
    return base + " They are being counted more than once because duplicate removal is switched off."
