"""Command-line entry point.

argparse wiring + exit-code contract: 0 when the report passes, 1 when it fails,
so shell callers can gate on the exit status alone.
"""


import argparse
import sys

from .copycheck import check_copy
from .gate import check_upload_gate
from .identity import check_client_identity
from .integrity import check_original_ma_integrity
from .io_compat import _emit, _load_json
from .plan import check_validate_plan
from .prewrite import check_pre_write
from .source import check_source_availability


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
