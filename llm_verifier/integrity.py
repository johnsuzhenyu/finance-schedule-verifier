"""E5 -- original MA integrity (append-only verification).

The generated workbook must still contain every original sheet, unchanged and in
the original order -- schedules are ADDED, never rewritten. Works on cell values,
so formatting-only differences do not raise false alarms.
"""


from .io_compat import (
    HAS_OPENPYXL,
    openpyxl,  # re-exported by the compat module (None if openpyxl missing)
)


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
