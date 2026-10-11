"""E1b -- copy-from-historical checker (anti-copy guard).

The historical MA is a reference for TYPE and FORMAT only, never for item content.
This module builds an item signature (numbers and dates stripped) and measures how
much of the new schedule body overlaps the previous period; above threshold it
fails as "copied", and it also locates the current workbook's GL sheet so the two
signature sets are comparable.
"""


import re

from .cells import _cell_text, _has_number, _is_date_like, _is_number, _join_row, _norm
from .io_compat import HAS_OPENPYXL, _load_workbook


# ---------------------------------------------------------------------------
# E1b. Copy-from-historical checker (anti-copy guard)
# ---------------------------------------------------------------------------

# Rows that are structural, not item content (v14 block structure:
# title "As at [FYE]" / item table / Checking+Variance / Agent Notes / rollforward lines).
_SKIP_ITEM_PREFIXES = (
    "as at", "as of", "checking", "variance", "note", "notes", "total",
    "composition", "agent", "[ok]", "[!]", "[pending]", "opening", "closing",
    "data source", "source", "generated", "schedule",
)


def _is_item_row(row):
    """A real item row carries both descriptive text and a numeric amount
    (e.g. `2025-01-15 | Apple MacBook Pro — Personal Expense | INV-001 | 3,500.00`)."""
    has_text = any(
        _cell_text(v) and not _is_number(v) and not _is_date_like(v) for v in row
    )
    return has_text and _has_number(row)


def _item_signature(row):
    """Normalize an item row into a comparison signature: lowercase words only,
    with dates and amounts stripped — so structural/descriptive copies are caught
    even when the current period's amounts differ (the 4033 split-copy case)."""
    parts = []
    for v in row:
        s = _cell_text(v)
        if not s or _is_number(v) or _is_date_like(v):
            continue
        s = re.sub(r"(?i)\b(?:sgd|usd|eur|rmb|s\$)\s?\d[\d,.]*", " ", s)  # S$4,000.00
        s = re.sub(r"\$\s?\d[\d,.]*", " ", s)                              # $1,234.56
        parts.append(s)
    text = _norm(" ".join(parts))
    toks = [t for t in text.split() if not re.fullmatch(r"\d[\d,.]*", t)]
    return " ".join(toks)


def _looks_like_source_sheet(title):
    """Heuristic used ONLY in the content-based fallback: does the sheet name read
    like a raw source/helper sheet rather than a schedule output?"""
    t = title.lower()
    markers = ("profit", "loss", "balance", "ledger", "trial", "general ledger",
               "journal", "cash book", "bank statement", "data source",
               "table of contents", "tb", "gl", "pnl", "p&l")
    return any(m in t for m in markers)


def _has_checking_variance_pair(rows):
    """A schedule sheet carries both a row-starting 'Checking' row and a 'Variance' row
    (excel-output-spec §3.5 mandatory pair). Used as the content-based fallback signal."""
    j = [_join_row(r) for r in rows]
    return any(s.startswith("checking") for s in j) and any(s.startswith("variance") for s in j)


def _has_schedule_title_signal(rows):
    """Weak content signal for legacy schedule sheets (pre-§3.5 formats that carry no
    Checking/Variance pair): a row-starting 'As at / As of' title line AND numeric rows —
    the classic legacy block layout `[Schedule Name] / As at [FYE] / <item table>`."""
    j = [_join_row(r) for r in rows]
    has_title = any(s.startswith(("as at", "as of")) for s in j)
    has_num = any(_has_number(r) for r in rows)
    return has_title and has_num


def _resolve_schedule_titles(wb, declared=None):
    """Decide which sheets of a workbook are schedule sheets.

    Declaration-based first: the Agent names them explicitly (it knows — it generated
    them / read the historical MA, v14 §10.1 item 2 / §10.2 Check D). Declared names that
    exist in the workbook win; names that don't exist are simply ignored.

    Fallback (no declaration): content-driven recognition, two tiers —
      1. strong: Checking+Variance row pair (spec-compliant schedules)
      2. weak: 'As at/As of' title + numeric rows (legacy schedules, e.g. v13-era
         "UOB SG Reconciliation Summary" / "Aged Payables Details" that contain neither
         a "schedule" keyword nor Checking/Variance rows)
    Source/helper sheets (GL/BS/PL/TB/Journal/Data Source/ToC…) are always excluded.
    """
    if declared:
        existing = {ws.title for ws in wb.worksheets}
        hits = [n for n in declared if n in existing]
        if hits:
            return hits
        # declared names all missing -> fall through to content-based recognition
    titles = []
    for ws in wb.worksheets:
        if _looks_like_source_sheet(ws.title):
            continue
        rows = [
            list(r) for r in ws.iter_rows(values_only=True)
            if any(v is not None and str(v).strip() != "" for v in r)
        ]
        if _has_checking_variance_pair(rows) or _has_schedule_title_signal(rows):
            titles.append(ws.title)
    return titles


def _schedule_item_signatures(path, declared_titles=None):
    """Extract normalized item signatures from the schedule sheets of a workbook.

    Sheet selection: declared names win; else content-based recognition via
    `_resolve_schedule_titles` (Checking+Variance pair, excluding source/helper sheets).
    """
    wb = _load_workbook(path)
    titles = _resolve_schedule_titles(wb, declared_titles)
    sigs = set()
    for ws in wb.worksheets:
        if ws.title not in titles:
            continue
        for row in ws.iter_rows(values_only=True):
            vals = [v for v in row if v is not None and str(v).strip() != ""]
            if not vals:
                continue
            if _join_row(vals).startswith(_SKIP_ITEM_PREFIXES):
                continue
            if not _is_item_row(vals):
                continue
            sig = _item_signature(vals)
            if len(sig.split()) >= 2:  # drop noise like a lone currency remnant
                sigs.add(sig)
    return sigs, titles


def check_copy(workbook_path, historical_path, schedule_sheets=None, historical_schedule_sheets=None):
    """Detect whether generated items were copied from the historical MA instead of
    being regenerated from the current GL.

    Rule (v14): historical MA provides only TYPE + FORMAT reference; item content must
    be regenerated from the current-period GL. Maggie's walkthrough caught a real case —
    opening 4033 "split" into 2180+872+981 with closing still 4033: a copy error.

    Sheet selection is declaration-based first (--schedule-sheets /
    --historical-schedule-sheets), with content-based fallback (Checking+Variance pair,
    excluding source/helper sheets) — so legacy historical MAs whose schedule sheets use
    non-"schedule" names (e.g. "UOB SG Reconciliation Summary") are still recognized.

    Comparison is signature-based: dates and amounts are stripped, so even items whose
    amounts were tweaked are caught as long as the description/structure was copied.
    overlap_ratio >= 0.80 -> possible_historical_copy (fail-closed).
    """
    report = {
        "checker": "copy_check",
        "pass": True,
        "workbook": str(workbook_path),
        "historical": str(historical_path),
        "overlap_ratio": None,
        "violations": [],
    }
    if not HAS_OPENPYXL:
        report["pass"] = False
        report["violations"].append("openpyxl not installed — cannot compare workbooks")
        return report

    try:
        current_sigs, current_titles = _schedule_item_signatures(workbook_path, schedule_sheets)
        report["current_schedule_sheets"] = current_titles
    except Exception as e:
        report["pass"] = False
        report["violations"].append(
            f"cannot read current workbook — {type(e).__name__}: {e}"
        )
        return report

    try:
        hist_sigs, hist_titles = _schedule_item_signatures(historical_path, historical_schedule_sheets)
        report["historical_schedule_sheets"] = hist_titles
    except Exception as e:
        report["pass"] = False
        report["violations"].append(
            f"cannot read historical workbook — {type(e).__name__}: {e}"
        )
        return report

    if not current_sigs:
        report["pass"] = False
        report["violations"].append(
            "no item rows found in current workbook schedule sheets — copy comparison not possible"
        )
        return report

    if not hist_sigs:
        # Fail-closed: without a comparable historical schedule baseline, copying
        # cannot be ruled out, so the anti-copy gate does not pass.
        report["pass"] = False
        report["violations"].append(
            "no item rows found in historical workbook schedule sheets — copy comparison not possible"
        )
        return report

    overlap = current_sigs & hist_sigs
    ratio = len(overlap) / len(current_sigs)
    report["overlap_ratio"] = round(ratio, 2)
    report["item_count"] = len(current_sigs)
    report["overlap_count"] = len(overlap)

    if ratio >= 0.8:
        report["violations"].append({
            "kind": "possible_historical_copy",
            "detail": f"{ratio:.0%} of current items overlap historical MA items — items may have been copied instead of regenerated from current GL",
        })

    report["pass"] = len(report["violations"]) == 0
    return report
