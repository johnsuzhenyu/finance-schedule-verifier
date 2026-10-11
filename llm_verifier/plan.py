"""E0 -- plan validator.

Enforces the *declarative* scope contract of draft_plan.json / confirmed_plan.json:
the 22-type whitelist, source/status legality, and evidence presence. It never
judges whether the evidence is factually right -- user confirmation is that gate.

Depends on `vocabulary` for the whitelist and on `io_compat` for reading JSON;
knows nothing about workbooks.
"""

import json

from .io_compat import _load_json
from .vocabulary import VALID_SOURCES, VALID_STATUSES, _normalize_type_entry


# ---------------------------------------------------------------------------
# E0. Plan validator (declarative scope — no alias table / no sheet-name matching)
# ---------------------------------------------------------------------------

def check_validate_plan(plan_path, mode="draft"):
    """Validate draft_plan.json or confirmed_plan.json structure and whitelist.

    Agent declares each type with source + evidence (semantic detection).
    Script enforces: 22-type whitelist, evidence present, status contract.
    Does NOT judge whether evidence is correct — user confirm is that gate.
    """
    report = {
        "checker": "validate_plan",
        "pass": True,
        "plan": str(plan_path),
        "mode": mode,
        "violations": [],
        "confirmed_types": [],
        "recommended_types": [],
    }
    try:
        plan = _load_json(plan_path)
    except (OSError, json.JSONDecodeError) as e:
        report["pass"] = False
        report["violations"].append({"issue": "plan_unreadable", "detail": str(e)})
        return report

    if plan.get("plan_version") != "1":
        report["violations"].append({
            "issue": "bad_plan_version",
            "detail": f"plan_version must be '1', got {plan.get('plan_version')!r}",
        })

    if not plan.get("client_id"):
        report["violations"].append({
            "issue": "missing_client_id",
            "detail": "client_id is required",
        })

    types = plan.get("types")
    if not types or not isinstance(types, list):
        report["violations"].append({
            "issue": "missing_types",
            "detail": "types array is required and must be non-empty",
        })
        report["pass"] = len(report["violations"]) == 0
        return report

    seen = set()
    for i, entry in enumerate(types):
        if not isinstance(entry, dict):
            report["violations"].append({
                "issue": "invalid_type_entry",
                "detail": f"types[{i}] must be an object",
            })
            continue

        display, slug = _normalize_type_entry(entry.get("type"))
        if not display:
            report["violations"].append({
                "issue": "unknown_type",
                "detail": f"types[{i}].type {entry.get('type')!r} not in 22-type whitelist",
            })
            continue

        if slug in seen:
            report["violations"].append({
                "issue": "duplicate_type",
                "detail": f"duplicate type {display}",
            })
        seen.add(slug)

        source = entry.get("source")
        if source not in VALID_SOURCES:
            report["violations"].append({
                "issue": "invalid_source",
                "detail": f"{display}: source must be one of {sorted(VALID_SOURCES)}, got {source!r}",
            })

        evidence = (entry.get("evidence") or "").strip()
        if len(evidence) < 8:
            report["violations"].append({
                "issue": "missing_evidence",
                "detail": f"{display}: evidence required (min 8 chars) — cite BS account, MA block, or GL",
            })

        status = entry.get("status")
        if status not in VALID_STATUSES:
            report["violations"].append({
                "issue": "invalid_status",
                "detail": f"{display}: status must be one of {sorted(VALID_STATUSES)}, got {status!r}",
            })

        if status == "blocked" and not (entry.get("gap_reason") or "").strip():
            report["violations"].append({
                "issue": "blocked_without_reason",
                "detail": f"{display}: blocked status requires gap_reason",
            })

        if status == "recommended":
            report["recommended_types"].append(display)
        if status == "confirmed":
            report["confirmed_types"].append(display)

    if mode == "draft":
        if not report["recommended_types"] and not report["confirmed_types"]:
            report["violations"].append({
                "issue": "empty_recommendation",
                "detail": "draft plan needs at least one type with status recommended or confirmed",
            })

    if mode == "confirmed":
        if not plan.get("confirmed_at"):
            report["violations"].append({
                "issue": "missing_confirmed_at",
                "detail": "confirmed plan requires confirmed_at (ISO-8601 timestamp)",
            })
        if not report["confirmed_types"]:
            report["violations"].append({
                "issue": "no_confirmed_types",
                "detail": "confirmed plan needs at least one type with status confirmed",
            })
        # excluded/blocked must not appear as confirmed — already enforced by status field
        stray = [t for t in report["recommended_types"]]
        if stray:
            report["violations"].append({
                "issue": "unconfirmed_recommended",
                "detail": f"types still 'recommended' after confirm — set to confirmed/excluded/blocked: {stray}",
            })

    report["pass"] = len(report["violations"]) == 0
    return report
