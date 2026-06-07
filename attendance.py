"""
attendance.py — GenMeet attendance consolidation logic
=======================================================
Imported by app.py Tab 3. No Streamlit dependency here —
pure data wrangling so this module is independently testable.

Public API
----------
    load_f2f(file)  → pd.DataFrame   normalised F2F rows
    load_zoom(file) → pd.DataFrame   deduplicated Zoom rows
    consolidate(f2f_df, zoom_df) → pd.DataFrame  Nickname | Final Status
"""

import io
import pandas as pd

# ── Status priority: higher = beats lower ─────────────────────────────────────
_PRIORITY = {"Attendee": 2, "Late": 1, "Absent": 0}


def _top_status(statuses) -> str:
    """Return the highest-priority status from an iterable."""
    return max(statuses, key=lambda s: _PRIORITY.get(s, -1), default="Absent")


def _norm_nick(s) -> str:
    """Lowercase + strip for case-insensitive nickname matching."""
    return str(s).strip().lower()


# ══════════════════════════════════════════════════════════════════════════════
# DURATION PARSING
# ══════════════════════════════════════════════════════════════════════════════

def _parse_duration_seconds(val) -> int:
    """
    Convert a Zoom duration value to total seconds.

    Handles:
      "1:23:45"  → H:MM:SS  (most common Zoom export format)
      "23:45"    → MM:SS
      "90"       → plain integer treated as minutes
      NaN / ""   → 0
    """
    if pd.isna(val):
        return 0
    val = str(val).strip()
    if not val:
        return 0

    parts = val.split(":")
    try:
        if len(parts) == 3:                           # H:MM:SS
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        elif len(parts) == 2:                         # MM:SS
            return int(parts[0]) * 60 + int(parts[1])
        else:                                         # plain number → assume minutes
            return int(float(val)) * 60
    except (ValueError, TypeError):
        return 0


def _format_duration(total_seconds: int) -> str:
    """Format seconds back to H:MM:SS for display."""
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h}:{m:02d}:{s:02d}"


# ══════════════════════════════════════════════════════════════════════════════
# LOADERS
# ══════════════════════════════════════════════════════════════════════════════

def load_f2f(file) -> pd.DataFrame:
    """
    Read an F2F attendance CSV/Excel file.
    Expected columns (case-insensitive): Surname, Nickname, Entry Time, Status
    Returns a cleaned DataFrame with those four columns.
    """
    raw = _read_file(file)
    raw.columns = [c.strip() for c in raw.columns]

    # Normalise column names case-insensitively
    col_map = {c: c for c in raw.columns}
    for col in raw.columns:
        low = col.lower()
        if "surname" in low:
            col_map[col] = "Surname"
        elif "nickname" in low or "nick" in low:
            col_map[col] = "Nickname"
        elif "entry" in low or "time" in low:
            col_map[col] = "Entry Time"
        elif "status" in low:
            col_map[col] = "Status"
    raw = raw.rename(columns=col_map)

    for required in ("Nickname", "Status"):
        if required not in raw.columns:
            raise ValueError(f"F2F file is missing a '{required}' column.")

    # Fill optional columns if absent
    if "Surname" not in raw.columns:
        raw["Surname"] = ""
    if "Entry Time" not in raw.columns:
        raw["Entry Time"] = ""

    raw["Status"] = raw["Status"].str.strip().str.capitalize()
    raw["_nick_key"] = raw["Nickname"].apply(_norm_nick)
    return raw[["Surname", "Nickname", "Entry Time", "Status", "_nick_key"]]


def load_zoom(file) -> pd.DataFrame:
    """
    Read a Zoom attendance CSV file and deduplicate reconnections.
    Expected columns (case-insensitive): Nickname, Duration, Status

    Deduplication:
      - Group by normalised Nickname
      - Sum Duration (seconds)
      - Collapse Status: Attendee > Late > Absent
    Returns one row per unique participant with total Duration as H:MM:SS.
    """
    raw = _read_file(file)
    raw.columns = [c.strip() for c in raw.columns]

    col_map = {c: c for c in raw.columns}
    for col in raw.columns:
        low = col.lower()
        if "nickname" in low or "nick" in low or "name" in low:
            col_map[col] = "Nickname"
        elif "duration" in low or "time" in low:
            col_map[col] = "Duration"
        elif "status" in low:
            col_map[col] = "Status"
    raw = raw.rename(columns=col_map)

    for required in ("Nickname", "Status"):
        if required not in raw.columns:
            raise ValueError(f"Zoom file is missing a '{required}' column.")
    if "Duration" not in raw.columns:
        raw["Duration"] = 0

    raw["Status"] = raw["Status"].str.strip().str.capitalize()
    raw["_nick_key"]   = raw["Nickname"].apply(_norm_nick)
    raw["_duration_s"] = raw["Duration"].apply(_parse_duration_seconds)

    deduped = (
        raw.groupby("_nick_key", sort=False)
        .agg(
            Nickname=("Nickname", "first"),
            _duration_s=("_duration_s", "sum"),
            Status=("Status", lambda x: _top_status(list(x))),
        )
        .reset_index()
    )
    deduped["Duration"] = deduped["_duration_s"].apply(_format_duration)
    return deduped[["_nick_key", "Nickname", "Duration", "Status"]]


# ══════════════════════════════════════════════════════════════════════════════
# CONSOLIDATION
# ══════════════════════════════════════════════════════════════════════════════

def _final_status(f2f: str, zoom: str) -> str:
    """
    Apply consolidation rules:
      Attendee in either  → Attendee
      Late in at least one, not Attendee in either → Late
      Otherwise → Absent
    """
    statuses = [s for s in (f2f, zoom) if s != "N/A"]
    if not statuses:
        return "Absent"
    if "Attendee" in statuses:
        return "Attendee"
    if "Late" in statuses:
        return "Late"
    return "Absent"


def consolidate(
    f2f_df: pd.DataFrame | None,
    zoom_df: pd.DataFrame | None,
) -> pd.DataFrame:
    """
    Merge F2F and Zoom DataFrames (as returned by load_f2f / load_zoom)
    and compute the Final Status for every participant.

    Returns a DataFrame with exactly two columns:
        Nickname | Final Status
    sorted alphabetically by Nickname.
    """
    # Build master dict keyed by normalised nickname
    people: dict[str, dict] = {}

    if f2f_df is not None:
        for _, row in f2f_df.iterrows():
            key = row["_nick_key"]
            people[key] = {
                "Nickname":   row["Nickname"],
                "f2f_status": row["Status"],
                "zoom_status": "N/A",
            }

    if zoom_df is not None:
        for _, row in zoom_df.iterrows():
            key = row["_nick_key"]
            if key in people:
                people[key]["zoom_status"] = row["Status"]
            else:
                # Zoom-only participant (no F2F record)
                people[key] = {
                    "Nickname":   row["Nickname"],
                    "f2f_status": "N/A",
                    "zoom_status": row["Status"],
                }

    rows = [
        {
            "Nickname":     data["Nickname"],
            "Final Status": _final_status(data["f2f_status"], data["zoom_status"]),
        }
        for data in people.values()
    ]

    df = pd.DataFrame(rows, columns=["Nickname", "Final Status"])
    return df.sort_values("Nickname", key=lambda s: s.str.lower()).reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
# INTERNAL HELPER
# ══════════════════════════════════════════════════════════════════════════════

def _read_file(file) -> pd.DataFrame:
    """Read CSV or Excel from a file-like object based on its name."""
    name = getattr(file, "name", "")
    if name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(file)
    return pd.read_csv(file)
