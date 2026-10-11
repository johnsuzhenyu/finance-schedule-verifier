"""E1 -- pre-write structural validator (declaration-based).

The Agent declares which sheets it created; this checker judges them on structure
alone: Checking/Variance paired, not empty, not merged, not summary-only, and not
a silent downgrade of a confirmed type. It never re-guesses sheet identity from
content -- declaration is the contract.
"""


import json
import re

from .cells import _canon, _has_number, _join_row
from .io_compat import HAS_OPENPYXL, _iter_sheets_with_rows, _load_json, _load_workbook
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
