#!/usr/bin/env python3
"""
Schedule Generate — Guardrail Checkers (E-layer · declaration-based structural enforcement)

Design principle (v2):
    The LLM (Agent) does *semantic* detection during Phase A — it locates the GL sheet,
    finds the header row, maps columns, and records the original source-sheet names.
    This checker does *deterministic structural validation* on top of those DECLARED facts.

    It does NOT re-guess "which sheet is GL" / "which sheets are schedules" / "where is
    the client name" from fragile content heuristics. The fragile assumptions removed in v2:
      - first sheet == client identity source   → now: filename-based matching
      - header is row[0]                        → now: robust header scan (canonical column match)
      - schedule sheet name contains "schedule" → now: declared via --schedule-sheets (explicit)
      - Checking/Variance == schedule signature → now: not relied on (manager-reviewed GL may
        also contain check/variance markers)

Checkers:
  E0. validate_plan        — run after Phase B recommendation and after user confirm:
       whitelist + evidence + status contract on draft_plan.json / confirmed_plan.json
  E2. source_availability  — run before generation (Phase A→B):
       does the declared GL sheet contain usable detail (transaction rows + Description + Reference)?
  E1. pre_write_validate   — run before write (Phase D→E):
       are the DECLARED schedule sheets structurally sound (Checking+Variance paired, non-empty,
       not merged, not summary/note-heavy/silent-downgraded)?
       The Agent declares the schedule sheets explicitly (it created them); source sheets and
       helper sheets (Journal Report, Table of Contents) are simply not listed. No Data Source
       Log sheet is created anymore — the workbook stays clean.
  E1b. copy_check          — run before upload (Phase E→F, only if a historical MA exists):
       do the generated schedule items overlap the historical MA's schedule items?
       Anti-copy guard: historical MA provides only type/format reference, never item content
       (v14 §10.1 item 14 — the 4033 anti-copy case). Items must be regenerated from current GL.
  E3. client_identity      — run at intake + before write:
       do the output filename and uploaded FILENAMES align with the current client?
  E4. upload_gate          — run before upload:
       did E1 + E3 (and optionally E2) all pass?

Usage:
  python schedule_guardrails.py validate_plan --plan <plan.json> [--mode draft|confirmed] [--json]
  python schedule_guardrails.py source_availability --gl <file> [--gl-sheet <name>] [--json]
  python schedule_guardrails.py pre_write_validate --workbook <out> --schedule-sheets <n1> <n2> ... \
      [--plan <confirmed_plan.json>] [--expected-sheets <t1> ...] [--json]
  python schedule_guardrails.py client_identity --client-id <id> --client-name <name> \
      --files <f1> [<f2> ...] [--output <out>] [--json]
  python schedule_guardrails.py copy_check --workbook <out> --historical <hist.xlsx> \
      [--schedule-sheets <n1> <n2> ...] [--historical-schedule-sheets <n1> <n2> ...] [--json]
  python schedule_guardrails.py original_ma_integrity --source <src.xlsx> --workbook <out.xlsx> \
      --original-sheets <n1> [<n2> ...] [--json]
  python schedule_guardrails.py upload_gate --prewrite <r1.json> --identity <r2.json> \
      [--integrity <r3.json>] [--availability <r4.json>] [--json]

Each checker prints a JSON report to stdout and exits 0 (pass) or 1 (fail).
`--json` prints the full report; otherwise a short human summary is printed.
To feed results into upload_gate, capture stdout: `... --json > r1.json`.
"""

import argparse
import re
import sys

from .cells import _cell_text, _has_number, _is_date_like, _is_number, _join_row, _norm
from .identity import check_client_identity
from .integrity import check_original_ma_integrity
from .io_compat import HAS_OPENPYXL, _emit, _load_json, _load_workbook
from .plan import check_validate_plan
from .prewrite import check_pre_write
from .source import check_source_availability


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


# ---------------------------------------------------------------------------
# E4. Upload gate
# ---------------------------------------------------------------------------

def check_upload_gate(prewrite_report, identity_report, availability_report=None,
                      integrity_report=None, plan_report=None):
    """Only allow upload when all prior checks passed."""
    report = {
        "checker": "upload_gate",
        "pass": True,
        "checks": {},
        "reasons": [],
    }
    if prewrite_report:
        ok = bool(prewrite_report.get("pass", False))
        report["checks"]["pre_write_validate"] = ok
        if not ok:
            report["reasons"].append("pre-write validation failed")

    if identity_report:
        ok = bool(identity_report.get("pass", False))
        report["checks"]["client_identity"] = ok
        if not ok:
            report["reasons"].append("client-identity check failed")

    if integrity_report:
        ok = bool(integrity_report.get("pass", False))
        report["checks"]["original_ma_integrity"] = ok
        if not ok:
            report["reasons"].append("original MA integrity failed — original sheets modified")

    if plan_report:
        ok = bool(plan_report.get("pass", False))
        report["checks"]["validate_plan"] = ok
        if not ok:
            report["reasons"].append("plan validation failed")

    if availability_report:
        ok = bool(availability_report.get("pass", True))
        report["checks"]["source_availability"] = ok
        if not ok:
            report["reasons"].append("source-availability check failed")

    if not report["checks"]:
        report["pass"] = False
        report["reasons"].append("no check results supplied — upload gate defaults to fail-closed")
    else:
        report["pass"] = all(report["checks"].values())

    return report


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(description="Schedule Generate guardrail checkers")
    sub = parser.add_subparsers(dest="command", required=True)

    p_plan = sub.add_parser("validate_plan")
    p_plan.add_argument("--plan", required=True, help="draft_plan.json or confirmed_plan.json")
    p_plan.add_argument("--mode", choices=("draft", "confirmed"), default="draft")
    p_plan.add_argument("--json", action="store_true")

    p_src = sub.add_parser("source_availability")
    p_src.add_argument("--gl", required=True)
    p_src.add_argument("--gl-sheet")
    p_src.add_argument("--json", action="store_true")

    p_pre = sub.add_parser("pre_write_validate")
    p_pre.add_argument("--workbook", required=True)
    p_pre.add_argument("--schedule-sheets", nargs="+", default=None)
    p_pre.add_argument("--expected-sheets", nargs="+", default=None,
                       help="user-CONFIRMED generation list from Phase B/C; every confirmed type must have an output block (completeness gate)")
    p_pre.add_argument("--plan", default=None,
                       help="confirmed_plan.json — derive expected types from status=confirmed")
    p_pre.add_argument("--json", action="store_true")

    p_id = sub.add_parser("client_identity")
    p_id.add_argument("--client-id", required=True)
    p_id.add_argument("--client-name", required=True)
    p_id.add_argument("--files", nargs="+", required=True)
    p_id.add_argument("--output")
    p_id.add_argument("--json", action="store_true")

    p_int = sub.add_parser("original_ma_integrity")
    p_int.add_argument("--source", required=True)
    p_int.add_argument("--workbook", required=True)
    p_int.add_argument("--original-sheets", nargs="+", required=True)
    p_int.add_argument("--json", action="store_true")

    p_up = sub.add_parser("upload_gate")
    p_up.add_argument("--prewrite", required=True)
    p_up.add_argument("--identity", required=True)
    p_up.add_argument("--integrity")
    p_up.add_argument("--plan")
    p_up.add_argument("--availability")
    p_up.add_argument("--json", action="store_true")

    p_copy = sub.add_parser("copy_check")
    p_copy.add_argument("--workbook", required=True)
    p_copy.add_argument("--historical", required=True)
    p_copy.add_argument("--schedule-sheets", nargs="*", default=None)
    p_copy.add_argument("--historical-schedule-sheets", nargs="*", default=None)
    p_copy.add_argument("--json", action="store_true")

    args = parser.parse_args()

    if args.command == "validate_plan":
        report = check_validate_plan(args.plan, args.mode)
    elif args.command == "source_availability":
        report = check_source_availability(args.gl, args.gl_sheet)
    elif args.command == "pre_write_validate":
        report = check_pre_write(args.workbook, args.schedule_sheets, args.expected_sheets,
                                 args.plan)
    elif args.command == "client_identity":
        report = check_client_identity(args.client_id, args.client_name, args.files, args.output)
    elif args.command == "original_ma_integrity":
        report = check_original_ma_integrity(
            args.source, args.workbook, args.original_sheets)
    elif args.command == "copy_check":
        report = check_copy(args.workbook, args.historical,
                            args.schedule_sheets, args.historical_schedule_sheets)
    elif args.command == "upload_gate":
        pre = _load_json(args.prewrite)
        ident = _load_json(args.identity)
        integrity = _load_json(args.integrity) if args.integrity else None
        plan_r = _load_json(args.plan) if args.plan else None
        avail = _load_json(args.availability) if args.availability else None
        report = check_upload_gate(pre, ident, avail, integrity, plan_r)

    _emit(report, args.json)
    sys.exit(0 if report.get("pass") else 1)


if __name__ == "__main__":
    main()
