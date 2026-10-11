"""E3 -- client identity / isolation checker (filename-based).

Guards against cross-client contamination: the output filename and every uploaded
filename must carry the current client's significant tokens. Pure string work --
no workbook is opened.
"""


import re
from pathlib import Path

from .cells import _norm


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
