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
import sys

from .copycheck import check_copy
from .identity import check_client_identity
from .integrity import check_original_ma_integrity
from .io_compat import _emit, _load_json
from .plan import check_validate_plan
from .prewrite import check_pre_write
from .source import check_source_availability


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
