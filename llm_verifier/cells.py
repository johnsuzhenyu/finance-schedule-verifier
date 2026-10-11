"""Cell / text helpers — pure functions, no package-internal dependencies."""

import re


def _cell_text(v):
    return "" if v is None else str(v).strip()


def _norm(text):
    """Lowercase, `&` -> `and`, collapse non-alphanumerics to single spaces."""
    s = str(text).lower().replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return " ".join(s.split())


def _canon(text):
    """Strip everything but [a-z0-9]; used for tolerant column-name matching."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def _join_row(row):
    # Filter empty cells so a label living in column B (A empty) does not produce a
    # leading space — startswith("checking"/"variance"/"total") must match regardless
    # of which column the label sits in (found in 0827 v3 outputs: Checking in B col).
    return " ".join(t for t in (_cell_text(v) for v in row) if t).lower()


def _is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _has_number(row):
    return any(_is_number(v) for v in row)


def _is_date_like(v):
    """True for datetime/date objects or strings that look like a date."""
    import datetime as _dt
    if isinstance(v, (_dt.datetime, _dt.date)):
        return True
    if isinstance(v, str):
        s = v.strip()
        return bool(re.match(r"^\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}$", s)) or \
               bool(re.match(r"^\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2}$", s))
    return False

