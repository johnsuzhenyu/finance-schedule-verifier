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

Checkers (one module each):
  E0.  plan.py       validate_plan         — run after Phase B recommendation and after user confirm:
       whitelist + evidence + status contract on draft_plan.json / confirmed_plan.json
  E2.  source.py     source_availability   — run before generation (Phase A→B):
       does the declared GL sheet contain usable detail (transaction rows + Description + Reference)?
  E1.  prewrite.py   pre_write_validate    — run before write (Phase D→E):
       are the DECLARED schedule sheets structurally sound (Checking+Variance paired, non-empty,
       not merged, not summary/note-heavy/silent-downgraded)?
       The Agent declares the schedule sheets explicitly (it created them); source sheets and
       helper sheets (Journal Report, Table of Contents) are simply not listed. No Data Source
       Log sheet is created anymore — the workbook stays clean.
  E1b. copycheck.py  copy_check            — run before upload (Phase E→F, only if a historical MA exists):
       do the generated schedule items overlap the historical MA's schedule items?
       Anti-copy guard: historical MA provides only type/format reference, never item content
       (v14 §10.1 item 14 — the 4033 anti-copy case). Items must be regenerated from current GL.
  E3.  identity.py   client_identity       — run at intake + before write:
       do the output filename and uploaded FILENAMES align with the current client?
  E5.  integrity.py  original_ma_integrity — the original MA sheets must survive untouched and in order.
  E4.  gate.py       upload_gate           — run before upload:
       did E1 + E3 (and optionally E2) all pass?

Layers:
  cells · io_compat · vocabulary  →  plan · source · prewrite · identity · integrity · copycheck · gate
  →  cli  (the only module that knows about every checker)

Usage:
  python -m llm_verifier validate_plan --plan <plan.json> [--mode draft|confirmed] [--json]
  python -m llm_verifier source_availability --gl <file> [--gl-sheet <name>] [--json]
  python -m llm_verifier pre_write_validate --workbook <out> --schedule-sheets <n1> <n2> ... \
      [--plan <confirmed_plan.json>] [--expected-sheets <t1> ...] [--json]
  python -m llm_verifier client_identity --client-id <id> --client-name <name> \
      --files <f1> [<f2> ...] [--output <out>] [--json]
  python -m llm_verifier copy_check --workbook <out> --historical <hist.xlsx> \
      [--schedule-sheets <n1> <n2> ...] [--historical-schedule-sheets <n1> <n2> ...] [--json]
  python -m llm_verifier original_ma_integrity --source <src.xlsx> --workbook <out.xlsx> \
      --original-sheets <n1> [<n2> ...] [--json]
  python -m llm_verifier upload_gate --prewrite <r1.json> --identity <r2.json> \
      [--integrity <r3.json>] [--availability <r4.json>] [--json]

Each checker prints a JSON report to stdout and exits 0 (pass) or 1 (fail).
`--json` prints the full report; otherwise a short human summary is printed.
To feed results into upload_gate, capture stdout: `... --json > r1.json`.
"""
