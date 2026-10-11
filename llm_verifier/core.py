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
import json
import re
import sys
from pathlib import Path

from .cells import (
    _canon,
    _cell_text,
    _has_number,
    _is_date_like,
    _is_number,
    _join_row,
    _norm,
)
from .io_compat import (
    HAS_OPENPYXL,
    _emit,
    _iter_sheets_with_rows,
    _load_json,
    _load_workbook,
    openpyxl,  # re-exported by the compat module (None if openpyxl missing)
)
from .plan import check_validate_plan
from .source import check_source_availability
from .vocabulary import _confirmed_types_from_plan


# ---------------------------------------------------------------------------
# E1. Pre-write structural validator (declaration-based)
# ---------------------------------------------------------------------------

def _classify_schedule_rows(rows):
    """Split a schedule sheet's rows into structural buckets."""
    data, total, note, checking, variance = [], [], [], [], []
    for r in rows:
        j = _join_row(r)
        if j.startswith("checking"):
            checking.append(r)
        elif j.startswith("variance"):
            variance.append(r)
        elif j.startswith("total"):
            total.append(r)
        elif j.startswith(("note", "[!]", "[ok]", "[pending]")) or "agent notes" in j:
            note.append(r)
        elif _has_number(r):
            data.append(r)
        # else: title / header / blank — not counted as data
    return data, total, note, checking, variance


def _type_core(expected_name):
    """'Assets · Deposits' -> 'deposits' (the type name after '·'); tolerates no '·'."""
    return _canon(expected_name.split("\u00b7")[-1].strip())


def _block_titles(rows):
    """Heuristic block-title lines inside a schedule sheet (non-numeric, not a control row)."""
    titles = []
    for r in rows:
        j = _join_row(r)
        if not j:
            continue
        if j.startswith(("checking", "variance", "total", "data", "note", "agent notes",
                         "[ok]", "[!]", "[pending]", "[je]", "as at")):
            continue
        if _has_number(r):
            continue
        if len(j) > 80:  # long descriptive line — header/desc, not a block title
            continue
        titles.append(j)
    return titles


def check_pre_write(workbook_path, schedule_sheets=None, expected_sheets=None, plan_path=None):
    """Validate the DECLARED schedule sheets (structural checks).

    `schedule_sheets` is the explicit list of sheets the Agent generated as schedules
    (display groups). Source sheets and helper sheets (Journal Report / Table of Contents)
    are simply not listed, so they are never validated as schedules. No Data Source Log
    sheet is created anymore — the workbook stays clean.

    `expected_sheets` (optional) is the user-CONFIRMED generation list from Phase B/C:
    every confirmed schedule type must have an output block in the declared sheets
    (completeness check — catches a real-world silent omission of 2 confirmed types).

    `plan_path` (optional): load confirmed types from confirmed_plan.json (status=confirmed).
    When both plan_path and expected_sheets are given, expected_sheets wins.

    Per sheet, runs six checks (aligned with SKILL.md upload gate contract):
      1. missing/unpaired Checking+Variance (hard)
      2. empty schedule — no numeric data rows (hard)
      3. summary_only — Total row but almost no data rows (hard)
      4. note_heavy_table_light — notes may be replacing table content (hard)
      5. possible_silent_downgrade — summary-like content with no downgrade disclosure (hard)
      6. merged_schedule — multiple "As at" schedule blocks but only one Checking row:
         several schedule types collapsed into one summary table instead of independent
         blocks (v14 §10.1 item 4; Yangwen's laziness-merge concern)
    """
    if plan_path and not expected_sheets:
        try:
            plan = _load_json(plan_path)
            expected_sheets = _confirmed_types_from_plan(plan)
        except (OSError, json.JSONDecodeError) as e:
            return {
                "checker": "pre_write_validate",
                "pass": False,
                "workbook": str(workbook_path),
                "plan": str(plan_path),
                "violations": [{"issue": "plan_unreadable", "detail": str(e)}],
            }

    report = {
        "checker": "pre_write_validate",
        "pass": True,
        "workbook": str(workbook_path),
        "schedule_sheets": list(schedule_sheets) if schedule_sheets else [],
        "expected_sheets": list(expected_sheets) if expected_sheets else [],
        "plan": str(plan_path) if plan_path else None,
        "violations": [],
    }
    if not HAS_OPENPYXL:
        report["pass"] = False
        report["violations"].append("openpyxl not installed — cannot inspect workbook")
        return report

    wb = _load_workbook(workbook_path)
    all_rows = {title: rows for title, rows in _iter_sheets_with_rows(wb)}
    declared = [s for s in (schedule_sheets or [])]

    # Resolve the schedule-sheet list: declared names win; else fall back to the
    # Checking+Variance signature (degraded mode — unreliable on manager-reviewed files).
    targets = []
    if declared:
        for name in declared:
            if name in all_rows:
                targets.append(name)
            else:
                report["violations"].append({
                    "sheet": name,
                    "issue": "declared_sheet_not_found",
                    "detail": f"declared schedule sheet '{name}' not found in workbook",
                })
    else:
        for title, rows in all_rows.items():
            joined = " ".join(_join_row(r) for r in rows)
            if "checking" in joined and "variance" in joined:
                targets.append(title)
        if not targets:
            report["violations"].append({
                "sheet": "(workbook)",
                "issue": "no_schedule_sheets",
                "detail": "no schedule sheets declared and none detected — pass --schedule-sheets",
            })

    report["schedule_sheets"] = targets

    for title in targets:
        rows = all_rows[title]
        data, total, note, checking, variance = _classify_schedule_rows(rows)

        # Hard check 1: Checking + Variance must both exist and be paired 1:1.
        if not checking and not variance:
            report["violations"].append({
                "sheet": title,
                "issue": "missing_checking_variance",
                "detail": "schedule sheet has no Checking/Variance rows — rollforward traceability broken",
            })
        elif len(checking) != len(variance):
            report["violations"].append({
                "sheet": title,
                "issue": "unpaired_checking_variance",
                "detail": f"{len(checking)} Checking row(s) vs {len(variance)} Variance row(s) — every schedule needs its own pair",
            })

        # Hard check 2: schedule sheet must not be empty of numeric data.
        if not data and not total:
            report["violations"].append({
                "sheet": title,
                "issue": "empty_schedule",
                "detail": "schedule sheet has no numeric data rows",
            })

        # Check 3-5: legacy semantic checks (v14 upload-gate contract:
        # summary-only / note-heavy / silent-downgrade). Their row definitions follow the
        # pre-existing 4-check implementation: `data_rows` = rows not starting with
        # checking/variance/note/agent (title/header/total rows included), `note_rows` =
        # note-prefixed rows. These are intentionally lenient so legitimate single-item
        # schedules (e.g. Director Remuneration with one composition row) are not flagged;
        # they exist to satisfy the SKILL.md §10.2 contract, while the hard structural
        # checks (1/2) and merged-schedule (6) carry the real enforcement.
        data_rows = [r for r in rows if not _join_row(r).startswith(("checking", "variance", "note", "agent"))]
        note_rows = [r for r in rows if _join_row(r).startswith(("note", "[!]", "[ok]", "[pending]", "agent notes"))]
        total_like = [r for r in rows if _join_row(r).startswith("total")]

        if len(data_rows) <= 1 and total_like:
            report["violations"].append({
                "sheet": title,
                "issue": "summary_only",
                "detail": f"only {len(data_rows)} data row(s) but a Total row exists — possible summary-only schedule",
            })

        if len(note_rows) >= len(data_rows) and len(data_rows) <= 2:
            report["violations"].append({
                "sheet": title,
                "issue": "note_heavy_table_light",
                "detail": f"{len(note_rows)} note row(s) vs {len(data_rows)} data row(s) — notes may be replacing table content",
            })

        downgrade_markers = any(
            k in _join_row(r) for r in rows
            for k in ("downgrade", "not uploaded", "summary-level", "insufficient data")
        )
        if (len(data_rows) <= 1 and total_like) and not downgrade_markers:
            report["violations"].append({
                "sheet": title,
                "issue": "possible_silent_downgrade",
                "detail": "summary-like content with no downgrade disclosure",
            })

        # Check 6: merged-schedule — multiple schedule blocks but only one Checking/Variance pair.
        # Each atomic schedule must be an independent block (title + item table + Checking/Variance).
        # Multiple "As at" titles (v14 block title format, excel-output-spec §4.2) with a single
        # Checking row means several types were collapsed into one summary table instead of
        # independent blocks (Yangwen's laziness-merge concern).
        as_at_count = sum(1 for r in rows if _join_row(r).startswith(("as at", "as of")))
        if as_at_count > 1 and len(checking) <= 1:
            report["violations"].append({
                "sheet": title,
                "issue": "merged_schedule",
                "detail": f"{as_at_count} schedule blocks (\"As at\" titles) but only {len(checking)} Checking row(s) — multiple schedule types may be merged into one summary table instead of independent blocks",
            })

    # Check 7: sheet-name validity (Excel hard limit). A name > 31 characters or containing
    # forbidden chars fails at create_sheet/save — catches the "workbook write stopped because
    # the generated sheet name exceeded Excel's 31-character limit" failure BEFORE the save.
    forbidden = re.search(r"[:\\/?*\[\]]", title)
    if len(title) > 31 or forbidden:
        report["violations"].append({
            "sheet": title,
            "issue": "invalid_sheet_name",
            "detail": f"sheet name '{title}' is {len(title)} chars with forbidden chars={bool(forbidden)} — Excel limit is 31 chars and disallows : \\ / ? * [ ] — the save would fail; shorten to ≤31 chars (e.g. 'Schedules - Dir & RP' for 'Schedules - Director & Related Party') before create_sheet",
        })

    # Check 8 (completeness — a real-world silent omission): every user-CONFIRMED
    # schedule type must have an output block in the declared schedule sheets. A confirmed
    # type that is neither generated nor disclosed is a FAIL, even if the declared sheets
    # are structurally valid.
    if expected_sheets:
        title_canons = []
        for t in targets:
            title_canons.extend(_canon(x) for x in _block_titles(all_rows.get(t, [])))
        # Account-name suffixes that a block title may append to the 22-type name
        # (e.g. 'Income Tax Payable' is the 'Income Tax' type). 'Received' is NOT in the
        # list on purpose: 'Deposits Received' is a DIFFERENT 22-type from 'Deposits'.
        _acct_suffix = ("payable", "receivable", "asset", "liability", "control", "account")
        for exp in expected_sheets:
            core = _type_core(exp)
            if not core:
                continue
            matched = False
            for c in title_canons:
                if c == core or c.endswith(core):
                    matched = True
                    break
                if c.startswith(core):
                    rest = c[len(core):]
                    if rest in _acct_suffix:
                        matched = True
                        break
            if not matched:
                report["violations"].append({
                    "sheet": "(expected)",
                    "issue": "expected_schedule_missing",
                    "detail": f"confirmed schedule type '{exp}' has no output block in the workbook — generate it (even at zero balance if historical) or disclose why in the conversation; a confirmed type neither generated nor disclosed is a FAIL",
                })

    report["pass"] = len(report["violations"]) == 0
    return report


# ---------------------------------------------------------------------------
# E3. Client identity / isolation checker (filename-based)
# ---------------------------------------------------------------------------

STOP_WORDS = {
    "pte", "ltd", "limited", "pl", "llp", "inc", "corp", "corporation",
    "co", "company", "the", "and", "sg", "singapore",
}

GENERIC_WORDS = {
    # financial-statement / document terms
    "profit", "loss", "balance", "sheet", "trial", "ledger", "general", "detail",
    "account", "accounts", "statement", "income", "expense", "expenses", "revenue",
    "tb", "bs", "gl", "pl", "pnl", "ma", "fs", "fy", "fye", "fy2024", "fy2025", "fy2026",
    "schedule", "schedules", "journal", "report", "workbook", "data", "source", "log",
    "summary", "of", "for", "the", "year", "ended", "period", "month", "months",
    "as", "at", "with", "without", "no", "and",
    # review / status words (common in filenames)
    "review", "reviewed", "jl", "final", "finalised", "finalized", "draft", "v1", "v2", "v3",
    "comments", "comment", "management", "manager", "delivered", "schedule",
    "internal", "copy", "without", "with", "for",
    # months (long + short)
    "january", "february", "march", "april", "may", "june", "july", "august",
    "september", "october", "november", "december",
    "jan", "feb", "mar", "apr", "jun", "jul", "aug", "sep", "sept", "oct", "nov", "dec",
}


def _significant_tokens(name):
    """Tokens in `name` that look like an entity name (drop stop words, generic words, numbers)."""
    toks = []
    for t in _norm(name).split():
        if t in STOP_WORDS or t in GENERIC_WORDS:
            continue
        if re.fullmatch(r"\d+", t):  # pure numbers / dates
            continue
        toks.append(t)
    return toks


def check_client_identity(client_id, client_name, files, output_path=None):
    """Verify the output filename and each uploaded FILENAME align with the current client."""
    report = {
        "checker": "client_identity",
        "pass": True,
        "client_id": client_id,
        "client_name": client_name,
        "client_tokens": _significant_tokens(client_name),
        "mismatches": [],
    }
    client_tokens = report["client_tokens"]

    # 1. Output filename must contain the client's significant tokens (hard gate).
    if output_path:
        out_name = Path(output_path).stem
        out_tokens = _significant_tokens(out_name)
        missing = [t for t in client_tokens if t not in out_tokens]
        if missing:
            report["mismatches"].append({
                "kind": "output_filename",
                "detail": f"output filename '{out_name}' is missing client tokens {missing} — "
                          f"output must be named for the current client",
            })

    # 2. Filename match-first: PASS if any client token appears in filename.
    for f in files:
        stem = Path(f).stem
        file_tokens = _significant_tokens(stem)
        if not file_tokens:
            continue
        overlap = [t for t in file_tokens if t in client_tokens]
        if overlap:
            continue
        report["mismatches"].append({
            "kind": "possible_foreign_entity",
            "detail": f"filename '{stem}' has tokens {file_tokens} with no overlap to "
                      f"selected client '{client_name}' (tokens: {client_tokens})",
        })

    report["pass"] = len(report["mismatches"]) == 0
    return report


# ---------------------------------------------------------------------------
# E5. Original MA integrity (append-only verification)
# ---------------------------------------------------------------------------

def _cell_values_map(ws):
    """Map coordinate -> value for all non-None cells in used range."""
    vals = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                vals[cell.coordinate] = cell.value
    return vals


def check_original_ma_integrity(source_path, output_path, original_sheets):
    """Verify output preserved every original sheet unchanged (append-only write).

    Checks: (1) all original sheet names exist in output, (2) same relative order
    among originals, (3) every non-empty cell value unchanged per original sheet.
    """
    report = {
        "checker": "original_ma_integrity",
        "pass": True,
        "source": str(source_path),
        "workbook": str(output_path),
        "original_sheets": list(original_sheets or []),
        "violations": [],
    }
    if not original_sheets:
        report["pass"] = False
        report["violations"].append("no original_sheets list provided")
        return report
    if not HAS_OPENPYXL:
        report["pass"] = False
        report["violations"].append("openpyxl not installed")
        return report

    src_wb = openpyxl.load_workbook(source_path, data_only=True)
    out_wb = openpyxl.load_workbook(output_path, data_only=True)

    # 1. All original names exist
    for name in original_sheets:
        if name not in out_wb.sheetnames:
            report["violations"].append({
                "issue": "original_sheet_missing",
                "detail": f"original sheet '{name}' missing from output — sheet may have been deleted",
            })

    # 2. Order among originals preserved (they must appear as a prefix subsequence in order)
    out_orig = [n for n in out_wb.sheetnames if n in original_sheets]
    if out_orig != list(original_sheets):
        report["violations"].append({
            "issue": "original_sheet_order_changed",
            "detail": f"expected order {original_sheets}, found {out_orig} in output",
        })

    # 3. Cell values unchanged on each original sheet
    for name in original_sheets:
        if name not in src_wb.sheetnames or name not in out_wb.sheetnames:
            continue
        src_vals = _cell_values_map(src_wb[name])
        out_vals = _cell_values_map(out_wb[name])
        if src_vals != out_vals:
            changed = []
            for coord, v in src_vals.items():
                if out_vals.get(coord) != v:
                    changed.append(coord)
                    if len(changed) >= 5:
                        break
            for coord, v in out_vals.items():
                if coord not in src_vals:
                    changed.append(f"+{coord}")
                    if len(changed) >= 5:
                        break
            report["violations"].append({
                "issue": "original_sheet_modified",
                "sheet": name,
                "detail": f"cells changed in original sheet '{name}' (sample: {changed[:5]}) — append-only violation",
            })

    report["pass"] = len(report["violations"]) == 0
    return report


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
