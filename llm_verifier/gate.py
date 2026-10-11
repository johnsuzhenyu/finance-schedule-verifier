"""E4 -- upload gate.

The only checker that does not look at a file: it consumes the JSON reports of the
other checkers and decides whether the workbook may be uploaded. Fail-closed by
construction -- a missing or failed report blocks the upload.
"""



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
