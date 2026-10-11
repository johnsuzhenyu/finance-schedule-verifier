"""Schedule-type vocabulary -- the closed sets the plan checker enforces.

Pure data and pure functions: no I/O, no openpyxl, no argparse. Every checker
that talks to the LLM about types (plan validation, pre-write validation) goes
through these same two sets, so the whitelist cannot drift between checkers.

Depends only on the text normalisers in `cells`, which keeps this module near
the bottom of the import graph.
"""

from .cells import _canon, _norm


# ---------------------------------------------------------------------------
# Canonical 22-type whitelist (display name · slug)
# ---------------------------------------------------------------------------

CANONICAL_TYPES = {
    "Assets · Other Receivables": "other-receivables",
    "Assets · Deposits": "deposits",
    "Assets · Insurance Prepayment": "insurance-prepayment",
    "Assets · Prepayments": "prepayments",
    "Assets · Investment Property": "investment-property",
    "Assets · Inventory": "inventory",
    "Assets · Investments": "investments",
    "Liabilities · CPF Payable": "cpf-payable",
    "Liabilities · GST": "gst",
    "Liabilities · Salary Payables": "salary-payables",
    "Liabilities · Accruals": "accruals",
    "Liabilities · Deferred Income": "deferred-income",
    "Liabilities · Deposits Received": "deposits-received",
    "Liabilities · Bank Loans": "bank-loans",
    "Liabilities · Government Grants": "government-grants",
    "Liabilities · Income Tax": "income-tax",
    "Liabilities · Provisions": "provisions",
    "Equity · Fund Reserves": "fund-reserves",
    "Director/Related · Director Current Account": "director-current-account",
    "Director/Related · Shareholder Loans": "shareholder-loans",
    "Director/Related · Intercompany": "intercompany",
    "ROU · Lease / Right-of-Use": "rou",
}

SLUG_TO_DISPLAY = {v: k for k, v in CANONICAL_TYPES.items()}

VALID_SOURCES = frozenset({"bs", "historical_ma", "gl_supplement"})
VALID_STATUSES = frozenset({"recommended", "confirmed", "excluded", "blocked"})


def _normalize_type_entry(raw_type):
    """Map display name or slug to (display, slug) or (None, None)."""
    if not raw_type or not str(raw_type).strip():
        return None, None
    s = str(raw_type).strip()
    if s in CANONICAL_TYPES:
        return s, CANONICAL_TYPES[s]
    slug = _canon(s.replace(" ", "-"))
    # slug keys use hyphens; _canon strips punctuation
    for display, canonical_slug in CANONICAL_TYPES.items():
        if _canon(canonical_slug) == slug or canonical_slug.replace("-", "") == slug:
            return display, canonical_slug
    # tolerate "Director CA" style shorthands — only if unambiguous substring
    lowered = _norm(s)
    hits = []
    for display, canonical_slug in CANONICAL_TYPES.items():
        core = _norm(display.split("\u00b7")[-1])
        if lowered == core or lowered.replace(" ", "") == core.replace(" ", ""):
            hits.append((display, canonical_slug))
    if len(hits) == 1:
        return hits[0]
    return None, None


def _confirmed_types_from_plan(plan):
    """Return display names of types with status confirmed."""
    out = []
    for entry in plan.get("types") or []:
        if entry.get("status") != "confirmed":
            continue
        display, _ = _normalize_type_entry(entry.get("type"))
        if display:
            out.append(display)
    return out
