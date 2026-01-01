import re
from typing import Optional
import pandas as pd

# Map common US tz abbreviations to canonical timezones
TZ_ABBR_MAP = {
    "EDT": "America/New_York",
    "EST": "America/New_York",
    "CDT": "America/Chicago",
    "CST": "America/Chicago",
    "MDT": "America/Denver",
    "MST": "America/Denver",
    "PDT": "America/Los_Angeles",
    "PST": "America/Los_Angeles",
}

TZ_TARGET = "America/Tegucigalpa"  # or "UTC"
_tz_pattern = re.compile(r"\b([A-Z]{2,4})\b")  # catches EDT, EST, etc.

def _parse_one(value: object) -> Optional[pd.Timestamp]:
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        ts = value
        if ts.tzinfo is None:
            return ts  # keep naive
        return ts.tz_convert(TZ_TARGET).tz_localize(None)

    s = str(value).strip()
    if not s:
        return pd.NaT

    m = _tz_pattern.search(s)
    tz_zone = None
    if m:
        abbr = m.group(1)
        tz_zone = TZ_ABBR_MAP.get(abbr)

    if tz_zone:
        s_wo_tz = _tz_pattern.sub("", s).strip()
        try:
            ts = pd.to_datetime(s_wo_tz, errors="raise")
            ts = ts.tz_localize(tz_zone, nonexistent="NaT", ambiguous="NaT")
            if pd.isna(ts):
                return pd.NaT
            return ts.tz_convert(TZ_TARGET).tz_localize(None)
        except Exception:
            return pd.NaT
    else:
        try:
            return pd.to_datetime(s, errors="raise")
        except Exception:
            return pd.NaT

def coerce_to_datetime(series: pd.Series) -> pd.Series:
    """Parse a mixed series of datetimes robustly, handling tz abbreviations."""
    if pd.api.types.is_datetime64_any_dtype(series):
        # If tz-aware, normalize to target and drop tz; else leave as-is
        try:
            if getattr(series.dt, "tz", None) is not None:
                return series.dt.tz_convert(TZ_TARGET).dt.tz_localize(None)
        except Exception:
            pass
        return series

    return series.map(_parse_one)
