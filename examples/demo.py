#!/usr/bin/env python3
"""End-to-end demo of the guardrail chain, run entirely against the synthetic fixtures.

This is the 60-second version of "what does this repo actually do":

    python examples/demo.py

It shells out to `python -m llm_verifier <checker>` exactly the way a caller would,
so what you see here is the real CLI contract -- a JSON report on stdout and an exit
code of 0 (pass) or 1 (fail) -- not a re-implementation. Each case carries the outcome
the fixtures were built to produce, so a green run means the chain still behaves.

Layout
------
Phase A (before generating)   E0  validate_plan          plan contract
                              E2  source_availability    is the declared GL usable
Phase D (before writing)      E1  pre_write_validate     are the output blocks sound
Phase E (before upload)       E1b copy_check             anti-copy vs historical MA
                              E3  client_identity        does the filename match the client
                              E5  original_ma_integrity  original sheets untouched, in order
                              E4  upload_gate            fail-closed verdict over all reports

No API key, no network: everything is local, deterministic and synthetic
(client "DEMO PTE LTD" -- no real client name, vendor or amount anywhere).
"""

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIX = ROOT / "fixtures"

# The bad-case workbook deliberately contains a sheet name longer than Excel's
# 31-char limit, which is what the E1 name check exists to catch. Never open it
# in Excel -- it will try to repair and rename the sheet.
OK_WB = FIX / "ma_fye2026_out_ok.xlsx"
BAD_WB = FIX / "ma_fye2026_out_bad.xlsx"
TAMPERED_WB = FIX / "ma_fye2026_out_tampered.xlsx"
INPUT_WB = FIX / "ma_fye2026_input.xlsx"
HIST_WB = FIX / "ma_hist_fye2025.xlsx"
PLAN_OK = FIX / "draft_plan_ok.json"
PLAN_BAD = FIX / "draft_plan_bad.json"

ORIGINALS = ["Balance Sheet", "Profit & Loss", "GL Detail", "Journal Report"]

# The clean workbook appended three schedule blocks; the negative workbook declares
# four, one of them purposely over the 31-char limit (and one intentionally empty).
SCHEDULES_OK = ["Schedules - Assets", "Schedules - Liabilities", "Schedules - Dir & RP"]
SCHEDULES_BAD = ["Schedules - Assets", "Schedules - Liabilities",
                 "Schedules - Equity", "Schedules - Director & Related Party"]

# E3 is pure filename string work -- it never opens a workbook -- so these paths are
# declarative, not fixtures that have to exist on disk.
CLIENT_ID, CLIENT_NAME = "DEMO-PTE-LTD", "DEMO PTE LTD"
IDENTITY_OK_FILES = ["DEMO PTE LTD - MA FY2026.xlsx"]
IDENTITY_BAD_FILES = ["ACME HOLDINGS PTE LTD - MA FY2026.xlsx"]


# ---------------------------------------------------------------------------
# tiny runner: call the CLI the way a shell caller would
# ---------------------------------------------------------------------------

def run(*argv):
    """Run one checker. Returns (exit_code, parsed_report)."""
    cmd = [sys.executable, "-m", "llm_verifier", *argv, "--json"]
    proc = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")
    report = None
    if proc.stdout.strip():
        try:
            report = json.loads(proc.stdout)
        except json.JSONDecodeError:
            report = None
    return proc.returncode, report


def reasons(report):
    """Flatten a report into the short list of things a human needs to read."""
    if not report:
        return []
    out = []
    for key in ("violations", "mismatches", "reasons"):
        for item in report.get(key, []):
            if isinstance(item, dict):
                label = item.get("issue") or item.get("kind") or "?"
                detail = item.get("detail") or item.get("sheet") or ""
                out.append(f"{label}: {detail}".rstrip(": "))
            else:
                out.append(str(item))
    return out


def show(title, code, report, expected_code):
    ok = code == expected_code
    mark = "  [ok]" if ok else "  [!!] UNEXPECTED"
    verdict = "PASS" if code == 0 else "FAIL"
    print(f"\n{mark}  {title}")
    print(f"        exit={code} (expected {expected_code})  ->  {verdict}")
    for line in reasons(report):
        print(f"          - {line}")
    return ok


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def ensure_fixtures():
    missing = [p.name for p in (OK_WB, BAD_WB, TAMPERED_WB, INPUT_WB, HIST_WB)
               if not p.exists()]
    if not missing:
        return
    print(f"fixtures missing ({', '.join(missing)}) -- rebuilding deterministically")
    subprocess.run([sys.executable, str(FIX / "build_ma_fixtures.py")],
                   cwd=ROOT, check=True)


def main():
    print("=" * 74)
    print("  finance-schedule-verifier  |  end-to-end demo (synthetic fixtures)")
    print("  \"The LLM declares. Python verifies.\"")
    print("=" * 74)
    ensure_fixtures()

    results = []

    print("\n--- Phase A: before generating -------------------------------------")
    code, rep = run("validate_plan", "--plan", str(PLAN_OK))
    results.append(show("E0 validate_plan          well-formed draft plan", code, rep, 0))

    code, rep = run("validate_plan", "--plan", str(PLAN_BAD))
    results.append(show("E0 validate_plan          unknown type / bad status / bad source / duplicate",
                        code, rep, 1))

    code, rep = run("source_availability", "--gl", str(INPUT_WB))
    results.append(show("E2 source_availability    auto-locate the GL sheet (13 transaction rows)",
                        code, rep, 0))

    print("\n--- Phase D: before writing ----------------------------------------")
    code, rep = run("pre_write_validate", "--workbook", str(OK_WB),
                    "--schedule-sheets", *SCHEDULES_OK)
    results.append(show("E1 pre_write_validate     clean output workbook", code, rep, 0))

    code, rep = run("pre_write_validate", "--workbook", str(BAD_WB),
                    "--schedule-sheets", *SCHEDULES_BAD)
    results.append(show("E1 pre_write_validate     missing pair / empty / merged / summary-only / bad name",
                        code, rep, 1))

    print("\n--- Phase E: before upload ------------------------------------------")
    code, rep = run("copy_check", "--workbook", str(OK_WB), "--historical", str(HIST_WB))
    results.append(show("E1b copy_check            regenerated items (low overlap)", code, rep, 0))

    code, rep = run("copy_check", "--workbook", str(BAD_WB), "--historical", str(HIST_WB))
    results.append(show("E1b copy_check            items copied from the historical MA", code, rep, 1))

    code, rep = run("client_identity", "--client-id", CLIENT_ID, "--client-name", CLIENT_NAME,
                    "--files", *IDENTITY_OK_FILES)
    results.append(show("E3 client_identity        filenames carry the current client", code, rep, 0))

    code, rep = run("client_identity", "--client-id", CLIENT_ID, "--client-name", CLIENT_NAME,
                    "--files", *IDENTITY_BAD_FILES)
    results.append(show("E3 client_identity        filename belongs to another entity", code, rep, 1))

    code, rep = run("original_ma_integrity", "--source", str(INPUT_WB),
                    "--workbook", str(OK_WB), "--original-sheets", *ORIGINALS)
    results.append(show("E5 original_ma_integrity  originals untouched and in order", code, rep, 0))

    code, rep = run("original_ma_integrity", "--source", str(INPUT_WB),
                    "--workbook", str(TAMPERED_WB), "--original-sheets", *ORIGINALS)
    results.append(show("E5 original_ma_integrity  original sheet modified and renamed", code, rep, 1))

    # ------------------------------------------------------------------ gate
    # E4 opens no file of its own: it consumes the JSON reports of the other
    # checkers and decides. Capture them the way a real caller would.
    print("\n--- E4 upload gate: fail-closed over the collected reports ----------")
    tmp = ROOT / ".demo_reports"
    tmp.mkdir(exist_ok=True)

    def capture(name, *argv):
        code, rep = run(*argv)
        (tmp / f"{name}.json").write_text(json.dumps(rep), encoding="utf-8")
        return code

    capture("prewrite", "pre_write_validate", "--workbook", str(OK_WB),
            "--schedule-sheets", *SCHEDULES_OK)
    capture("identity", "client_identity", "--client-id", CLIENT_ID,
            "--client-name", CLIENT_NAME, "--files", *IDENTITY_OK_FILES)
    capture("integrity", "original_ma_integrity", "--source", str(INPUT_WB),
            "--workbook", str(OK_WB), "--original-sheets", *ORIGINALS)

    code, rep = run("upload_gate",
                    "--prewrite", str(tmp / "prewrite.json"),
                    "--identity", str(tmp / "identity.json"),
                    "--integrity", str(tmp / "integrity.json"))
    results.append(show("E4 upload_gate            every upstream report green -> allow",
                        code, rep, 0))

    # now flip a single upstream report to red and re-run the gate
    (tmp / "integrity_red.json").write_text(json.dumps({
        "checker": "original_ma_integrity",
        "pass": False,
        "violations": [{"issue": "original_sheet_modified",
                        "detail": "cells changed in original sheet 'Balance Sheet' (C7)"}],
    }), encoding="utf-8")

    code, rep = run("upload_gate",
                    "--prewrite", str(tmp / "prewrite.json"),
                    "--identity", str(tmp / "identity.json"),
                    "--integrity", str(tmp / "integrity_red.json"))
    results.append(show("E4 upload_gate            one upstream report red -> block upload",
                        code, rep, 1))

    passed, total = sum(1 for r in results if r), len(results)
    print("\n" + "=" * 74)
    print(f"  {passed}/{total} cases behaved exactly as the fixtures were built to produce.")
    print("  All green: the chain is fail-closed end to end."
          if passed == total else
          "  Something drifted -- inspect the [!!] lines above.")
    print("=" * 74)
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
