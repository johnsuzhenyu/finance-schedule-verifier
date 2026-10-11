"""E2 -- source availability checker (+ the header-detection it relies on).

Answers one question before generation starts: does the DECLARED GL sheet really
carry usable detail (transaction rows with Description and Reference)?

`COLUMN_LEXICON` / `detect_header_row` live here because this checker is their
only consumer -- the header scan exists to serve GL readability, and the fragile
"first row is the header" assumption was removed in v2 in favour of this scan.
"""


from .cells import _canon, _is_date_like
from .io_compat import HAS_OPENPYXL, _load_workbook


# ---------------------------------------------------------------------------
# Header detection (robust — scans for the strongest column-signature row)
# ---------------------------------------------------------------------------

# Canonical column alias groups. A header cell matches if its canonical form
# contains any of these canonical aliases (punctuation/case/space-insensitive).
COLUMN_LEXICON = {
    "date": ["date", "transactiondate", "txndate", "journaldate", "invoicedate", "paymentdate"],
    "description": ["description", "narrative", "particulars", "memo", "details",
                    "contactname", "accountname", "name"],
    "reference": ["reference", "ref", "invoiceno", "invoicenumber", "invno",
                  "chequeno", "referenceno"],
    "debit": ["debit", "dr", "debitamount"],
    "credit": ["credit", "cr", "creditamount"],
    "amount": ["amount", "netamount", "value", "amountsgd"],
    "account_code": ["accountcode", "accountno", "glcode", "acctcode", "code", "account"],
    "source_type": ["sourcetype", "source", "type"],
    "due_date": ["duedate", "due"],
    "running_balance": ["runningbalance", "closingbalance", "balance"],
}


def _header_categories(cells):
    """Return the set of column categories present in one header row."""
    cats = set()
    for cell in cells:
        c = _canon(cell)
        if not c:
            continue
        for cat, aliases in COLUMN_LEXICON.items():
            if any(a in c for a in aliases):
                cats.add(cat)
                break
    return cats


def detect_header_row(rows, max_scan=25):
    """Find the row (index into `rows`) that best matches a header signature.

    Returns (best_index, categories_set). best_index is -1 if no row scores >= 2.
    """
    best_i, best_cats = -1, set()
    for i, row in enumerate(rows[:max_scan]):
        cats = _header_categories(row)
        if len(cats) > len(best_cats):
            best_i, best_cats = i, cats
    if len(best_cats) >= 2:
        return best_i, best_cats
    return -1, set()


# ---------------------------------------------------------------------------
# E2. Source availability checker
# ---------------------------------------------------------------------------

def check_source_availability(gl_path, gl_sheet=None):
    """Inspect the (declared or auto-located) GL sheet for usable detail."""
    report = {
        "checker": "source_availability",
        "pass": True,
        "gl_file": str(gl_path),
        "gl_sheet": gl_sheet,
        "header_row": None,
        "gl_has_detail": False,
        "gl_row_count": 0,
        "reference_present": False,
        "description_present": False,
        "source_type_present": False,
        "due_date_present": False,
        "violations": [],
        "warnings": [],
    }
    if not HAS_OPENPYXL:
        report["pass"] = False
        report["violations"].append("openpyxl not installed — cannot inspect GL")
        return report

    wb = _load_workbook(gl_path)

    # 1. Locate the GL sheet: declared name wins; else detect by GL role signature.
    sheets = [(ws.title, list(ws.iter_rows(values_only=True))) for ws in wb.worksheets]
    target_title, target_rows = None, None

    if gl_sheet:
        for title, rows in sheets:
            if title.lower() == gl_sheet.lower():
                target_title, target_rows = title, rows
                break
        if target_title is None:
            report["pass"] = False
            report["violations"].append(
                f"declared GL sheet '{gl_sheet}' not found; available sheets: "
                + ", ".join(t for t, _ in sheets)
            )
            return report
    else:
        # Role detection: GL = a sheet whose header has date + description + (debit|credit|amount).
        best_title, best_score, best_rows = None, 0, None
        for title, rows in sheets:
            _, cats = detect_header_row(rows)
            score = 0
            if "date" in cats:
                score += 1
            if "description" in cats:
                score += 1
            if cats & {"debit", "credit", "amount"}:
                score += 1
            if score > best_score:
                best_title, best_score, best_rows = title, score, rows
        if best_score >= 3:
            target_title, target_rows = best_title, best_rows
        else:
            report["pass"] = False
            report["violations"].append(
                "no GL-like sheet found (need date + description + debit/credit columns); "
                "pass --gl-sheet <name> to declare it explicitly"
            )
            return report

    report["gl_sheet"] = target_title

    # 2. Detect header row (robust scan) and identify columns.
    nonempty = [r for r in target_rows if any(v is not None and str(v).strip() != "" for v in r)]
    header_idx, cats = detect_header_row(nonempty)
    if header_idx < 0:
        report["pass"] = False
        report["violations"].append(
            f"GL sheet '{target_title}' has no recognizable header row (Date/Description/Debit/Credit columns)"
        )
        return report
    report["header_row"] = header_idx + 1  # 1-based among non-empty rows

    report["reference_present"] = "reference" in cats
    report["description_present"] = "description" in cats
    report["source_type_present"] = "source_type" in cats
    report["due_date_present"] = "due_date" in cats

    # 3. Count transaction rows = data rows whose first cell is date-like.
    tx_count = 0
    for r in nonempty[header_idx + 1:]:
        if r and _is_date_like(r[0]):
            tx_count += 1
    report["gl_row_count"] = tx_count
    report["gl_has_detail"] = tx_count > 0

    # 4. Pass/fail + soft warnings.
    if not report["gl_has_detail"]:
        report["pass"] = False
        report["violations"].append(
            f"GL sheet '{target_title}' has header but 0 transaction rows (no date-prefixed detail rows)"
        )
    if report["gl_has_detail"] and not report["description_present"]:
        report["warnings"].append(
            "GL has transactions but no Description column — per-schedule reviewability is weak"
        )

    return report
