#!/usr/bin/env python3
"""Deterministic builder for the synthetic MA (management-accounts) workbook fixtures.

Why a builder script instead of a committed .xlsx binary?
  - the fixtures stay reviewable as code: sheet layout, every number, and the
    INTENDED violation of each bad case are all explicit and diff-able;
  - the build is byte-reproducible: `md5sum fixtures/*.xlsx` is identical across
    runs, so a rebuild leaves `git status` clean and a fixture someone opened in
    Excel and re-saved is immediately visible as a diff. Getting there needs the
    post-processing pass in _normalize_xlsx() -- openpyxl alone is not
    reproducible (see that function for the two reasons).

All data is synthetic. Client "DEMO PTE LTD" (id DEMO-PTE-LTD), FYE 31 December.
No real client name, vendor or amount is used.

What it writes into fixtures/
  ma_hist_fye2025.xlsx            historical MA, FYE 31.12.2025
                                  originals + FY2025 schedules + one LEGACY sheet
                                  ("UOB SG Reconciliation Summary") that carries no
                                  Checking/Variance pair -> exercises the weak
                                  content-signature tier of _resolve_schedule_titles
  ma_fye2026_input.xlsx           current Target MA File: ORIGINAL sheets only
                                  (the upload that original_ma_integrity compares FROM)
  ma_fye2026_out_ok.xlsx          clean output: originals unchanged, schedules
                                  appended, Narrative block -> every checker PASSes
  ma_fye2026_out_bad.xlsx         negative case, 4 structural violations + copied items
  ma_fye2026_out_tampered.xlsx    append-only violation: original sheet cell changed
                                  and an original sheet renamed

Expected outcome per checker (measured, not assumed -- rerun after every edit):
  source_availability    input, auto-locate -> PASS: GL sheet "GL Detail", header_row 4,
                                                13 transaction rows
                         input --gl-sheet GL -> FAIL: declared sheet not found
  pre_write_validate     ok    -> PASS, 0 violations
                         bad   -> FAIL: missing_checking_variance + empty_schedule (Equity)
                                        merged_schedule (Assets)
                                        summary_only + possible_silent_downgrade (Liabilities)
                                        invalid_sheet_name (Dir & RP)
  copy_check             ok  vs hist -> PASS, overlap 0.38 (13 items, 5 shared)
                         bad vs hist -> FAIL, overlap 1.00 (8 items, 8 shared)
  original_ma_integrity  ok       vs input -> PASS
                         tampered vs input -> FAIL: original_sheet_missing +
                                                  original_sheet_order_changed +
                                                  original_sheet_modified (C7)

Two findings these fixtures exposed (both pre-existing in
script/schedule_guardrails.py, not introduced by the package split):

  1. Check 7 (invalid_sheet_name) sits OUTSIDE the `for title in targets:` loop, so it
     only ever validates the last-processed title. Declaring the 36-char sheet first
     makes the check silently report nothing; declaring it last makes it fire. The
     fixture is built so the bad case fails either way (the other four violations are
     per-sheet), but this is the same bug that also lets a real >31-char name through.
  2. copy_check is a *signature* comparison, and recurring vendors legitimately produce
     identical signatures year over year. The "ok" workbook already sits at 0.38 against
     the historical MA. Before the historical fixture carried its own Director Current
     Account, the deliberately-copied workbook landed at exactly 0.75 and PASSED -- i.e.
     the 0.80 threshold tolerates a lot of genuine copying. 0.38 vs 1.00 is the honest
     demonstration pair; anything in the 0.6-0.8 band deserves a manual look.

Note: ma_fye2026_out_bad.xlsx contains a sheet name longer than Excel's 31-char
limit, on purpose (that is what Check 7 exists to stop before the save). Do not
open that file in Excel.

Run:
  python fixtures/build_ma_fixtures.py
"""

import datetime as dt
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font


HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
if str(ROOT) not in sys.path:          # so the verify pass can use the real helpers
    sys.path.insert(0, str(ROOT))

CLIENT = "DEMO PTE LTD"

FYE_2025 = "31 December 2025"
FYE_2026 = "31 December 2026"

# Sheet-name constants (kept <= 31 chars = Excel hard limit, except the bad fixture)
SH_BS = "Balance Sheet"
SH_PL = "Profit & Loss"
SH_GL = "GL Detail"
SH_JR = "Journal Report"
SH_ASSETS = "Schedules - Assets"
SH_LIAB = "Schedules - Liabilities"
SH_EQUITY = "Schedules - Equity"
SH_DIRRP = "Schedules - Dir & RP"
SH_DIRRP_LONG = "Schedules - Director & Related Party"   # 36 chars -> Check 7 target
SH_LEGACY = "UOB SG Reconciliation Summary"
SH_NARRATIVE = "Narrative"

ORIGINAL_SHEETS = [SH_BS, SH_PL, SH_GL, SH_JR]          # for the FY2026 pair
ORIGINAL_SHEETS_HIST = ORIGINAL_SHEETS + [SH_LEGACY]    # FY2025 also ships the recon sheet


# ---------------------------------------------------------------------------
# shared writers
# ---------------------------------------------------------------------------

# openpyxl stamps docProps/core.xml with the build time, so two identical rebuilds
# produce two different files (git reports the .xlsx as modified even though nothing
# about the content changed). Pinning the document properties makes the output
# byte-reproducible -- `md5sum fixtures/*.xlsx` and `git status` must both be stable.
FIXED_TIMESTAMP = dt.datetime(2026, 1, 1, 0, 0, 0)
FIXED_ZIP_STAMP = (2026, 1, 1, 0, 0, 0)     # zip entry headers: 1980..2107 only
FIXED_AUTHOR = "llm-verifier fixtures"


def _normalize_xlsx(path):
    """Rewrite the .xlsx container so that repeated builds are byte-identical.

    Two things inside the container carry the build time and cannot be suppressed
    from the outside (both confirmed by reading the library sources):
      - every zip entry header: zipfile.writestr() defaults to
        time.localtime(time.time())[:6], and Python 3.12's zipfile has no
        SOURCE_DATE_EPOCH support;
      - docProps/core.xml's <dcterms:modified>: openpyxl's save_workbook()
        overwrites it unconditionally (openpyxl/writer/excel.py), so setting
        wb.properties.modified beforehand is pointless.

    So we post-process: same entries, same order, pinned timestamps. Without this
    `git status` reports all five workbooks as modified after every rebuild, and a
    fixture someone hand-edited in Excel is indistinguishable from a regenerated
    one -- which is exactly the property we want a committed binary to have.
    """
    import io
    import re
    import zipfile

    with zipfile.ZipFile(path) as src:
        items = [(i.filename, src.read(i.filename), i.compress_type) for i in src.infolist()]

    # Rebuild in memory and write back in place. (Rewriting to a .tmp file and
    # os.replace()-ing over the target is the tidier pattern, but replacing an
    # existing file is rejected in some sandboxes -- WinError 5 -- and a crash
    # halfway would leave a stray .xlsx.tmp next to the fixtures.)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
        for name, data, compress in items:
            if name == "docProps/core.xml":
                text = re.sub(
                    r"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)",
                    r"\g<1>2026-01-01T00:00:00Z\g<2>",
                    data.decode("utf-8"),
                )
                data = text.encode("utf-8")
            info = zipfile.ZipInfo(name, date_time=FIXED_ZIP_STAMP)
            info.compress_type = compress
            info.create_system = 3          # pin to unix: same bytes on any build machine
            info.external_attr = 0o644 << 16
            dst.writestr(info, data)
    path.write_bytes(buf.getvalue())


def _save(wb, path):
    wb.properties.creator = FIXED_AUTHOR
    wb.properties.lastModifiedBy = FIXED_AUTHOR
    wb.properties.created = FIXED_TIMESTAMP
    wb.save(path)          # the ONLY place that may call wb.save directly
    _normalize_xlsx(path)  # openpyxl stamps <dcterms:modified> itself -- normalise after


def _write(ws, rows, title_rows=(), total_rows=(), date_cols=(), amount_cols=()):
    """Append row lists, then apply the minimal formatting the spec asks for."""
    for r in rows:
        ws.append(r)
    for r_no in title_rows:
        ws.cell(row=r_no, column=1).font = Font(bold=True, size=12)
    for r_no in total_rows:
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r_no, column=c).font = Font(bold=True)
    for r_no in range(1, ws.max_row + 1):
        for col in date_cols:
            cell = ws.cell(row=r_no, column=col)
            if isinstance(cell.value, dt.date):
                cell.number_format = "yyyy-mm-dd"
        for col in amount_cols:
            cell = ws.cell(row=r_no, column=col)
            if isinstance(cell.value, float):
                cell.number_format = "#,##0.00"


def block(title, as_at, header, items, total_value, checking, variance, data_source, notes,
          total_label="Total"):
    """One atomic schedule block per excel-output-spec v10 §3.1 + §3.5.

    title / As at / header / items / Total / Checking / Variance / Data source /
    blank / [Agent Notes] / notes / blank / blank -- the trailing blank row is the
    spec's mandatory block separation (§3.1 block-spacing rule: >=1 blank row), and
    it doubles as the empty-row-skipping case for _iter_sheets_with_rows.
    """
    rows = [[title], [f"As at {as_at}"], [], list(header)]
    for it in items:
        rows.append(list(it))
    rows.append([total_label] + [""] * (len(header) - 2) + [total_value])
    rows.append(["Checking", checking])
    rows.append(["Variance", variance])
    rows.append([data_source])
    rows.append([])
    rows.append(["[Agent Notes]"])
    for n in notes:
        rows.append([n])
    rows.append([])
    rows.append([])
    return rows


def no_check_block(title, as_at, header, items, total_value, data_source):
    """A block that is missing its Checking/Variance pair (negative case)."""
    rows = [[title], [f"As at {as_at}"], [], list(header)]
    for it in items:
        rows.append(list(it))
    rows.append(["Total"] + [""] * (len(header) - 2) + [total_value])
    rows.append([data_source])
    rows.append([])
    return rows


# ---------------------------------------------------------------------------
# original (client-owned) sheets -- identical wherever they must pass integrity
# ---------------------------------------------------------------------------

def rows_balance_sheet(fye, prior_label, bs):
    rows = [
        [CLIENT],
        ["Balance Sheet"],
        [f"As at {fye}"],
        [],
        ["Account", "Code", fye, prior_label],
        ["Assets"],
    ]
    for label, code, cur, prior in bs["assets"]:
        rows.append([label, code, cur, prior])
    rows.append(["Liabilities"])
    for label, code, cur, prior in bs["liabilities"]:
        rows.append([label, code, cur, prior])
    return rows


def rows_pl(fye, data):
    rows = [
        [CLIENT],
        ["Profit & Loss"],
        [f"For the year ended {fye}"],
        [],
        ["Account", "Code", "Amount (S$)"],
    ]
    for label, code, amt in data:
        rows.append([label, code, amt])
    return rows


def rows_gl(start, end, txns):
    rows = [
        [CLIENT],
        ["General Ledger Detail"],
        [f"{start} to {end}"],
        [],
        ["Date", "Description", "Reference", "Account Code", "Debit", "Credit"],
    ]
    for d, desc, ref, code, dr, cr in txns:
        rows.append([d, desc, ref, code, dr, cr])
    return rows


def rows_journal(fye):
    return [
        ["Journal Report"],
        [f"For the year ended {fye}"],
        [],
        ["Date", "Journal No.", "Description", "Amount (S$)"],
        [dt.date(2026, 1, 31), "JE-01", "Accrual - statutory audit fee", 3000.00],
        [dt.date(2026, 5, 31), "JE-02", "GST provision Q1", 3120.00],
        [dt.date(2026, 6, 30), "JE-03", "Accrual - professional body subscription", 1000.00],
    ]


def rows_narrative(fye):
    return [
        [f"NARRATIVE - {CLIENT} - {fye}"],
        [],
        ["Period overview:"],
        ["  Prepayments and deposits held flat year on year; accruals increased with the FY2026"],
        ["  audit fee and tax computation being provided for at year end."],
        [],
        ["Key movements:"],
        ["  - Accruals up S$400.00 on the FY2025 balance as the annual subscription"],
        ["  - Amount due to director increased S$1,070.80 on expenses paid personally"],
        [],
        ["Anomalies & attention:"],
        ["  - Two other-receivable balances are still outstanding and expected to clear in Q1"],
        [],
        ["Next-period attention:"],
        ["  - Confirm the director's personal expense reimbursements before the next review"],
    ]


def rows_legacy_recon(fye):
    """Legacy client sheet: 'As at' title + numeric rows, no Checking/Variance pair.

    This is the weak content-signature tier: it is still recognised as a schedule
    sheet when no --schedule-sheets is declared (core.py _has_schedule_title_signal).
    """
    return [
        [SH_LEGACY],
        [f"As at {fye}"],
        [],
        ["Description", "Reference", "Amount (S$)"],
        ["Bank balance per statement", "STMT-DEC", 41208.55],
        ["Less: unpresented cheques", "CHQ-3312", 1240.00],
        ["Add: deposits in transit", "DEP-2412", 860.00],
        ["Reconciled bank balance", "", 40828.55],
    ]


# ---------------------------------------------------------------------------
# FY2025 (historical) — balances and schedule item content
# ---------------------------------------------------------------------------

BS_FY2025 = {
    "assets": [
        ("Other Receivables", "1-1200", 1850.00, 1600.00),
        ("Deposits", "1-1300", 4500.00, 4500.00),
        ("Prepayments", "1-1400", 660.00, 480.00),
    ],
    "liabilities": [
        ("GST Payable", "2-2200", 2880.00, 3010.00),
        ("CPF Payable", "2-2300", 1674.00, 1590.00),
        ("Accruals", "2-3000", 4800.00, 4100.00),
        ("Amount Due to Director", "2-3100", 1247.60, 3.98),
    ],
}

PL_FY2025 = [
    ("Revenue", "4-1000", 232000.00),
    ("Professional Fees", "5-2000", 17400.00),
    ("Software Subscriptions", "5-2100", 3600.00),
    ("Secretarial Fees", "5-2200", 1800.00),
]

GL_FY2025 = [
    (dt.date(2025, 2, 14), "Northgate Secretarial - annual secretarial renewal", "INV-114", "1-1400", 360.00, None),
    (dt.date(2025, 3, 3), "Cloudline Workspace - software subscription annual", "INV-118", "1-1400", 300.00, None),
    (dt.date(2025, 4, 30), "Watermark Facilities - refundable tenancy deposit", "DEP-03", "1-1300", 4500.00, None),
    (dt.date(2025, 6, 30), "Redhill Audit LLP - statutory audit fee", "INV-050", "2-3000", None, 3000.00),
    (dt.date(2025, 7, 31), "Tax Agent Co - tax computation", "INV-052", "2-3000", None, 1200.00),
    (dt.date(2025, 9, 30), "Professional Body - annual subscription", "INV-055", "2-3000", None, 600.00),
    (dt.date(2025, 10, 31), "GST control - net payable per F5", "GST-2025", "2-2200", None, 2880.00),
    (dt.date(2025, 12, 31), "CPF Board - December contribution", "CPF-12", "2-2300", None, 1674.00),
    (dt.date(2025, 11, 20), "Meridian Trading - advance billing", "INV-201", "1-1200", 1050.00, None),
    (dt.date(2025, 11, 28), "Staff advance - K. Tan, travel", "ADV-02", "1-1200", 800.00, None),
]

HDR_4 = ["Item / Supplier", "Reference", "Coverage Period", "Amount Deferred (S$)"]
HDR_OR = ["Item / Counterparty", "Reference", "Nature", "Amount (S$)"]
HDR_3 = ["Item / Description", "Reference", "Amount (S$)"]

ITEMS_HIST_PREPAY = [
    ["Northgate Secretarial", "INV-114", "1 Jan 2025 - 31 Dec 2025", 360.00],
    ["Cloudline Workspace", "INV-118", "1 Jan 2025 - 31 Dec 2025", 300.00],
]
ITEMS_HIST_OR = [
    ["Meridian Trading", "INV-201", "Advance billing, offset in Q1 FY2026", 1050.00],
    ["Staff Advance - K. Tan", "ADV-02", "Travel advance, unacquitted", 800.00],
]
ITEMS_HIST_DEPOSITS = [
    ["Watermark Facilities", "DEP-03", "Refundable tenancy deposit", 4500.00],
]
ITEMS_HIST_ACCRUALS = [
    ["Redhill Audit LLP", "INV-050", "Statutory audit FY2025", 3000.00],
    ["Tax Agent Co", "INV-052", "Tax computation FY2025", 1200.00],
    ["Professional Body", "INV-055", "Annual subscription 2025", 600.00],
]
ITEMS_HIST_GST = [
    ["GST control - net payable per F5", "GST-2025", 2880.00],
]
ITEMS_HIST_CPF = [
    ["CPF payable - December contribution", "CPF-12", 1674.00],
]


def rows_hist_assets():
    return (
        block("Prepayments", FYE_2025, HDR_4, ITEMS_HIST_PREPAY, 660.00,
              "S$660.00", "S$0.00",
              "Data source: BS Prepayments 1-1400 - GL Reference# INV-114, INV-118",
              ["[OK] Composition S$660.00 = BS Prepayments S$660.00 - consistent",
               "[OK] 2 items; vendor, service and coverage period explicit",
               "[Pending] next period journal: Dr Professional Fees / Cr Prepayments S$360"])
        + block("Other Receivables", FYE_2025, HDR_OR, ITEMS_HIST_OR, 1850.00,
                "S$1,850.00", "S$0.00",
                "Data source: BS Other Receivables 1-1200 - GL Reference# INV-201, ADV-02",
                ["[OK] Composition S$1,850.00 = BS Other Receivables S$1,850.00 - consistent",
                 "[Pending] next period journal: Dr Bank / Cr Other Receivables S$1,050",
                 "Calculation logic: sum of un-offset advances and overpayments at year end"])
        + block("Deposits", FYE_2025, HDR_3, ITEMS_HIST_DEPOSITS, 4500.00,
                "S$4,500.00", "S$0.00",
                "Data source: BS Deposits 1-1300 - GL Reference# DEP-03",
                ["[OK] Composition S$4,500.00 = BS Deposits S$4,500.00 - consistent",
                 "[OK] single refundable tenancy deposit, no movement in the year"])
    )


def rows_hist_liabilities():
    return (
        block("Accruals", FYE_2025, HDR_3, ITEMS_HIST_ACCRUALS, 4800.00,
              "S$4,800.00", "S$0.00",
              "Data source: BS Accruals 2-3000 - GL Reference# INV-050, INV-052, INV-055",
              ["[OK] Composition S$4,800.00 = BS Accruals S$4,800.00 - consistent",
               "[OK] 3 items; vendor, service and period all explicit",
               "[Pending] next period journal: Dr Professional Fees / Cr Accruals S$4,800"])
        + block("GST", FYE_2025, HDR_3, ITEMS_HIST_GST, 2880.00,
                "S$2,880.00", "S$0.00",
                "Data source: BS GST Payable 2-2200 - GL Reference# GST-2025",
                ["[OK] Composition S$2,880.00 = BS GST Payable S$2,880.00 - consistent",
                 "Calculation logic: output tax less input tax for the December quarter"])
        + block("CPF Payable", FYE_2025, HDR_3, ITEMS_HIST_CPF, 1674.00,
                "S$1,674.00", "S$0.00",
                "Data source: BS CPF Payable 2-2300 - GL Reference# CPF-12",
                ["[OK] Composition S$1,674.00 = BS CPF Payable S$1,674.00 - consistent",
                 "[Pending] next period journal: Dr CPF Expense / Cr CPF Payable on payment"])
    )


# ---------------------------------------------------------------------------
# FY2026 (current) — balances and schedule item content
# ---------------------------------------------------------------------------

BS_FY2026 = {
    "assets": [
        ("Other Receivables", "1-1200", 2430.00, 1850.00),
        ("Deposits", "1-1300", 4500.00, 4500.00),
        ("Prepayments", "1-1400", 660.00, 660.00),
    ],
    "liabilities": [
        ("GST Payable", "2-2200", 3120.00, 2880.00),
        ("CPF Payable", "2-2300", 1908.00, 1674.00),
        ("Accruals", "2-3000", 5200.00, 4800.00),
        ("Amount Due to Director", "2-3100", 2318.40, 1247.60),
    ],
}

PL_FY2026 = [
    ("Revenue", "4-1000", 246000.00),
    ("Professional Fees", "5-2000", 18200.00),
    ("Software Subscriptions", "5-2100", 4200.00),
    ("Secretarial Fees", "5-2200", 1800.00),
]

GL_FY2026 = [
    (dt.date(2026, 1, 15), "Northbridge Corporate - annual secretarial renewal", "INV-364", "1-1400", 360.00, None),
    (dt.date(2026, 1, 20), "Cloudline Workspace - software subscription annual", "INV-370", "1-1400", 300.00, None),
    (dt.date(2026, 2, 28), "Redhill Audit LLP - statutory audit fee", "INV-210", "2-3000", None, 3000.00),
    (dt.date(2026, 3, 31), "Tax Agent Co - tax computation", "INV-215", "2-3000", None, 1200.00),
    (dt.date(2026, 6, 30), "Professional Body - annual subscription", "INV-220", "2-3000", None, 1000.00),
    (dt.date(2026, 4, 30), "Watermark Facilities - refundable tenancy deposit", "DEP-11", "1-1300", 4500.00, None),
    (dt.date(2026, 5, 31), "GST control - net payable per F5", "GST-Q1", "2-2200", None, 3120.00),
    (dt.date(2026, 5, 31), "CPF Board - May contribution", "CPF-05", "2-2300", None, 1908.00),
    (dt.date(2026, 7, 10), "Meridian Trading - advance billing", "INV-401", "1-1200", 1200.00, None),
    (dt.date(2026, 8, 22), "Harbourline Logistics - overpayment to be refunded", "INV-402", "1-1200", 780.00, None),
    (dt.date(2026, 10, 6), "Staff advance - K. Tan, travel", "ADV-07", "1-1200", 450.00, None),
    (dt.date(2026, 4, 18), "Company expenses paid personally by director", "DC-04", "2-3100", None, 862.40),
    (dt.date(2026, 9, 5), "Reimbursement of travel advance to director", "DC-09", "2-3100", None, 208.40),
]

ITEMS_PREPAY = [
    ["Northbridge Corporate", "INV-364", "1 Jan 2026 - 31 Dec 2026", 360.00],
    ["Cloudline Workspace", "INV-370", "1 Jan 2026 - 31 Dec 2026", 300.00],
]
ITEMS_OR = [
    ["Meridian Trading", "INV-401", "Advance billing, offset in Q1 FY2027", 1200.00],
    ["Harbourline Logistics", "INV-402", "Overpayment to be refunded", 780.00],
    ["Staff Advance - K. Tan", "ADV-07", "Travel advance, unacquitted", 450.00],
]
ITEMS_DEPOSITS = [
    ["Watermark Facilities", "DEP-11", "Refundable tenancy deposit", 4500.00],
]
ITEMS_ACCRUALS = [
    ["Redhill Audit LLP", "INV-210", "Statutory audit FY2026", 3000.00],
    ["Tax Agent Co", "INV-215", "Tax computation FY2026", 1200.00],
    ["Professional Body", "INV-220", "Annual subscription 2026", 1000.00],
]
ITEMS_GST = [
    ["GST control - net payable per F5", "GST-Q1", 3120.00],
]
ITEMS_CPF = [
    ["CPF payable - December contribution", "CPF-05", 1908.00],
]
# Director current account: FY2025 (historical) and FY2026 (current) use DIFFERENT
# wording on purpose -- the anti-copy checker only has the previous year to compare
# against, so current-period items must be describable in their own terms.
DIR_CA_LINES_HIST = [
    ["Opening Balance", "", "", 3.98],
    [dt.date(2025, 3, 31), "Company expenses paid personally by director", "DC-01", 900.00],
    [dt.date(2025, 9, 30), "Reimbursement of travel advance to director", "DC-02", 343.62],
    ["Closing Balance", "", "", 1247.60],
]
DIR_CA_CLOSING_HIST = "S$1,247.60"

DIR_CA_LINES_CUR = [
    ["Opening Balance", "", "", 1247.60],
    [dt.date(2026, 4, 18), "Office supplies purchased personally by director", "DC-04", 862.40],
    [dt.date(2026, 9, 5), "Client entertainment reimbursed to director", "DC-09", 208.40],
    ["Closing Balance", "", "", 2318.40],
]
DIR_CA_CLOSING_CUR = "S$2,318.40"


def rows_hist_dirrp():
    return rows_dirrp(
        FYE_2025, DIR_CA_LINES_HIST, DIR_CA_CLOSING_HIST, "DC-01, DC-02",
        ["[OK] Ledger closing S$1,247.60 = BS Amount Due to Director S$1,247.60 - consistent",
         "[OK] Payroll and CPF transactions excluded from the body (see CPF Payable)",
         "[Pending] next period journal: Dr Director Current Account / Cr Bank on reimbursement"],
    )


def rows_out_assets(item_src="current"):
    prepay = ITEMS_PREPAY if item_src == "current" else ITEMS_HIST_PREPAY
    orows = ITEMS_OR if item_src == "current" else ITEMS_HIST_OR
    return (
        block("Prepayments", FYE_2026, HDR_4, prepay, 660.00,
              "S$660.00", "S$0.00",
              "Data source: BS Prepayments 1-1400 - GL Reference# INV-364, INV-370",
              ["[OK] Composition S$660.00 = BS Prepayments S$660.00 - consistent",
               "[OK] 2 items; vendor, service and coverage period explicit",
               "[Pending] next period journal: Dr Professional Fees / Cr Prepayments S$360"])
        + block("Other Receivables", FYE_2026, HDR_OR, orows, 2430.00,
                "S$2,430.00", "S$0.00",
                "Data source: BS Other Receivables 1-1200 - GL Reference# INV-401, INV-402, ADV-07",
                ["[OK] Composition S$2,430.00 = BS Other Receivables S$2,430.00 - consistent",
                 "[OK] 3 items; counterparty and nature explicit",
                 "[Pending] next period journal: Dr Bank / Cr Other Receivables S$1,200"])
        + block("Deposits", FYE_2026, HDR_3, ITEMS_DEPOSITS, 4500.00,
                "S$4,500.00", "S$0.00",
                "Data source: BS Deposits 1-1300 - GL Reference# DEP-11",
                ["[OK] Composition S$4,500.00 = BS Deposits S$4,500.00 - consistent",
                 "[OK] single refundable tenancy deposit, unchanged in the year"])
    )


def rows_out_liabilities(item_src="current"):
    acc = ITEMS_ACCRUALS if item_src == "current" else ITEMS_HIST_ACCRUALS
    return (
        block("Accruals", FYE_2026, HDR_3, acc, 5200.00,
              "S$5,200.00", "S$0.00",
              "Data source: BS Accruals 2-3000 - GL Reference# INV-210, INV-215, INV-220",
              ["[OK] Composition S$5,200.00 = BS Accruals S$5,200.00 - consistent",
               "[OK] 3 items; vendor, service and period all explicit",
               "[Pending] next period journal: Dr Professional Fees / Cr Accruals S$5,200"])
        + block("GST", FYE_2026, HDR_3, ITEMS_GST, 3120.00,
                "S$3,120.00", "S$0.00",
                "Data source: BS GST Payable 2-2200 - GL Reference# GST-Q1",
                ["[OK] Composition S$3,120.00 = BS GST Payable S$3,120.00 - consistent",
                 "Calculation logic: output tax less input tax for the December quarter"])
        + block("CPF Payable", FYE_2026, HDR_3, ITEMS_CPF, 1908.00,
                "S$1,908.00", "S$0.00",
                "Data source: BS CPF Payable 2-2300 - GL Reference# CPF-05",
                ["[OK] Composition S$1,908.00 = BS CPF Payable S$1,908.00 - consistent",
                 "[Pending] next period journal: Dr CPF Expense / Cr CPF Payable on payment"])
    )


def rows_dirrp(as_at, lines, closing, refs, notes):
    """Director current account block -- the one type where a rollforward IS the body."""
    out = [["Director Current Account"], [f"As at {as_at}"], [], ["Date", "Description", "Reference", "S$"]]
    for line in lines:
        out.append(list(line))
    out.append(["Checking", closing])
    out.append(["Variance", "S$0.00"])
    out.append([f"Data source: BS Amount Due to Director 2-3100 - GL Reference# {refs}"])
    out.append([])
    out.append(["[Agent Notes]"])
    for n in notes:
        out.append([n])
    out.append([])
    out.append([])
    return out


def rows_hist_but_consistent_dirrp():
    """The copied-looking variant used by the negative fixture: last year's body, last
    year's figures, still internally consistent -- and still wrong for this period."""
    return rows_dirrp(
        FYE_2026, DIR_CA_LINES_HIST, DIR_CA_CLOSING_HIST, "DC-01, DC-02",
        ["[OK] Ledger closing S$1,247.60 = BS Amount Due to Director S$1,247.60 - consistent",
         "[Pending] next period journal: Dr Director Current Account / Cr Bank on reimbursement"],
    )


# ---------------------------------------------------------------------------
# builders
# ---------------------------------------------------------------------------

def build_hist_fye2025(path):
    wb = Workbook()
    _write(wb.active, rows_balance_sheet(FYE_2025, "31.12.2024", BS_FY2025), title_rows=(1, 2))
    wb.active.title = SH_BS
    _write(wb.create_sheet(SH_PL), rows_pl(FYE_2025, PL_FY2025), title_rows=(1, 2))
    _write(wb.create_sheet(SH_GL), rows_gl("1 Jan 2025", "31 Dec 2025", GL_FY2025),
           title_rows=(1, 2), date_cols=(1,), amount_cols=(5, 6))
    _write(wb.create_sheet(SH_JR), rows_journal(FYE_2025), title_rows=(1,), date_cols=(1,))
    _write(wb.create_sheet(SH_LEGACY), rows_legacy_recon(FYE_2025), title_rows=(1,), amount_cols=(3,))
    _write(wb.create_sheet(SH_ASSETS), rows_hist_assets(), amount_cols=(4,))
    _write(wb.create_sheet(SH_LIAB), rows_hist_liabilities(), amount_cols=(3,))
    _write(wb.create_sheet(SH_DIRRP), rows_hist_dirrp(), date_cols=(1,), amount_cols=(4,))
    _write(wb.create_sheet(SH_NARRATIVE), rows_narrative("FYE 31.12.2025"), title_rows=(1,))
    _save(wb, path)


def build_input_fye2026(path):
    """Current Target MA File: the client's own sheets, nothing appended yet."""
    wb = Workbook()
    _write(wb.active, rows_balance_sheet(FYE_2026, "31.12.2025", BS_FY2026), title_rows=(1, 2))
    wb.active.title = SH_BS
    _write(wb.create_sheet(SH_PL), rows_pl(FYE_2026, PL_FY2026), title_rows=(1, 2))
    _write(wb.create_sheet(SH_GL), rows_gl("1 Jan 2026", "31 Dec 2026", GL_FY2026),
           title_rows=(1, 2), date_cols=(1,), amount_cols=(5, 6))
    _write(wb.create_sheet(SH_JR), rows_journal(FYE_2026), title_rows=(1,), date_cols=(1,))
    _save(wb, path)


def _open_input_copy(path):
    """Start every FY2026 output from the untouched input workbook."""
    wb = load_workbook(path)
    return wb


def build_out_ok(input_path, path):
    wb = _open_input_copy(input_path)
    _write(wb.create_sheet(SH_ASSETS), rows_out_assets(), amount_cols=(4,))
    _write(wb.create_sheet(SH_LIAB), rows_out_liabilities(), amount_cols=(3,))
    _write(wb.create_sheet(SH_DIRRP), rows_dirrp(
        FYE_2026, DIR_CA_LINES_CUR, DIR_CA_CLOSING_CUR, "DC-04, DC-09",
        ["[OK] Ledger closing S$2,318.40 = BS Amount Due to Director S$2,318.40 - consistent",
         "[OK] Payroll and CPF transactions excluded from the body (see CPF Payable)",
         "[Pending] next period journal: Dr Director Current Account / Cr Bank on reimbursement"],
    ), date_cols=(1,), amount_cols=(4,))
    _write(wb.create_sheet(SH_NARRATIVE), rows_narrative("FYE 31.12.2026"), title_rows=(1,))
    _save(wb, path)


def build_out_bad(input_path, path):
    """Negative case. Four independent structural violations + verbatim-copied items."""
    wb = _open_input_copy(input_path)

    # 1. merged_schedule: two "As at" blocks in one sheet, only one Checking row.
    bad_assets = (
        block("Prepayments", FYE_2026, HDR_4, ITEMS_HIST_PREPAY, 660.00,   # <- items copied from FY2025
              "S$660.00", "S$0.00",
              "Data source: BS Prepayments 1-1400 - GL Reference# INV-114, INV-118",
              ["[OK] Composition S$660.00 = BS Prepayments S$660.00 - consistent"])
        + no_check_block("Other Receivables", FYE_2026, HDR_OR, ITEMS_HIST_OR, 1850.00,
                         "Data source: BS Other Receivables 1-1200 - GL Reference# INV-201, ADV-02")
    )
    _write(wb.create_sheet(SH_ASSETS), bad_assets, amount_cols=(4,))

    # 2. summary_only + possible_silent_downgrade: a Total with no item rows behind it.
    bad_liab = (
        block("Accruals", FYE_2026, HDR_3, [], 5200.00,
              "S$5,200.00", "S$0.00",
              "Data source: BS Accruals 2-3000",
              ["[OK] Composition S$5,200.00 = BS Accruals S$5,200.00 - consistent"])
        + block("GST", FYE_2026, HDR_3, ITEMS_HIST_GST, 2880.00,        # <- copied items again
                "S$2,880.00", "S$0.00",
                "Data source: BS GST Payable 2-2200 - GL Reference# GST-2025",
                ["[OK] Composition S$2,880.00 = BS GST Payable S$2,880.00 - consistent"])
        + block("CPF Payable", FYE_2026, HDR_3, ITEMS_HIST_CPF, 1674.00,
                "S$1,674.00", "S$0.00",
                "Data source: BS CPF Payable 2-2300 - GL Reference# CPF-12",
                ["[OK] Composition S$1,674.00 = BS CPF Payable S$1,674.00 - consistent"])
    )
    _write(wb.create_sheet(SH_LIAB), bad_liab, amount_cols=(3,))

    # 3. missing_checking_variance + empty_schedule: no numbers at all, no pair.
    _write(wb.create_sheet(SH_EQUITY), [
        ["Fund Reserves"], [f"As at {FYE_2026}"], [],
        ["Note: no movement in the year - brought forward only."],
    ])

    # 4. invalid_sheet_name: 36 chars (Excel hard limit is 31) - the save would fail.
    #    Body is the verbatim FY2025 director current account -> copied item content.
    _write(wb.create_sheet(SH_DIRRP_LONG), rows_hist_but_consistent_dirrp(),
           date_cols=(1,), amount_cols=(4,))

    _write(wb.create_sheet(SH_NARRATIVE), rows_narrative("FYE 31.12.2026"), title_rows=(1,))
    _save(wb, path)


def build_out_tampered(ok_path, path):
    """Append-only violation: an original cell overwritten + an original sheet renamed."""
    wb = load_workbook(ok_path)
    ws = wb[SH_BS]
    ws["C7"] = 999999.00          # was 2,430.00 -> original sheet modified
    wb[SH_JR].title = "Journal (revised)"   # original sheet missing / order changed
    _save(wb, path)


# ---------------------------------------------------------------------------
# entry point + self-manifest
# ---------------------------------------------------------------------------

def main():
    hist = HERE / "ma_hist_fye2025.xlsx"
    inp = HERE / "ma_fye2026_input.xlsx"
    ok = HERE / "ma_fye2026_out_ok.xlsx"
    bad = HERE / "ma_fye2026_out_bad.xlsx"
    tampered = HERE / "ma_fye2026_out_tampered.xlsx"

    build_hist_fye2025(hist)
    build_input_fye2026(inp)
    build_out_ok(inp, ok)
    build_out_bad(inp, bad)
    build_out_tampered(ok, tampered)

    from llm_verifier.io_compat import _iter_sheets_with_rows, _load_workbook

    print("fixture manifest -- (sheet, non-empty rows) via _iter_sheets_with_rows\n")
    for f in (hist, inp, ok, bad, tampered):
        wb = _load_workbook(f)
        sheets = list(_iter_sheets_with_rows(wb))
        raw = sum(ws.max_row for ws in wb.worksheets)
        kept = sum(len(rows) for _, rows in sheets)
        print(f"  {f.name}  <-- {len(sheets)} sheets, {kept}/{raw} rows kept "
              f"(blank rows dropped)")
        for name, rows in sheets:
            print(f"      {name!r}: {len(rows)} rows")
        print()
    print(f"originals for the FY2026 integrity pair: {ORIGINAL_SHEETS}")
    print(f"originals of the FY2025 historical MA:   {ORIGINAL_SHEETS_HIST}")


if __name__ == "__main__":
    main()
