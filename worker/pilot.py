"""Validate and enrich a private NY alumni pilot without touching the MatchIQ UI.

Dry-run is the default. Live RocketReach calls require both --live and the
explicit --confirm-credits YES flag.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from difflib import SequenceMatcher
import json
import os
from pathlib import Path
import sys
import time
from typing import Any

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
import requests


LOOKUP_URL = "https://api.rocketreach.co/api/v2/person/lookup"
CHECK_STATUS_URL = "https://api.rocketreach.co/api/v2/person/checkStatus"
SCRAPE_STATUS_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbyI4HSVEi9ohVAz1-EBZ7U2SpJ2Aa5S-CGSp6PKwhaUXkS_X77ZLcUquAJ6uvQ_gSQ/exec"
)

HEADERS = {
    "first_name": "Person Account: First Name (Advancement)",
    "last_name": "Person Account: Last Name (Advancement)",
    "account_name": "Account Name",
    "home_city": "Person Account: Home City",
    "home_state": "Person Account: Home State/Province",
    "home_country": "Person Account: Home Country",
    "current_title": "Person Account: Current Job Title",
    "current_employer": "Person Account: Current Employer Name",
    "degree": "Primary Degree",
    "reunion_year": "Reunion Year",
}

RESULT_HEADERS = [
    "pilot_source_id",
    "lookup_timestamp_utc",
    "lookup_seconds",
    "lookup_status",
    "lookup_error",
    "rr_profile_id",
    "rr_name",
    "rr_company",
    "rr_industry",
    "rr_email",
    "rr_location",
    "rr_linkedin_url",
    "name_similarity",
    "company_changed",
    "location_changed",
    "review_queue",
    "human_decision",
    "human_comment",
]

# Indexes into each RESULT_HEADERS row (keep in sync with RESULT_HEADERS).
RV_LOOKUP_STATUS = 3
RV_LOOKUP_ERROR = 4
RV_RR_NAME = 6
RV_REVIEW_QUEUE = 15

SCRAPE_FOCUS_HEADERS = [
    "name",
    "graduation_year",
    "company_name",
    "industry",
    "email",
    "lookup_status",
    "review_queue",
    "outcome",
]

OPTIONAL_INPUT_HEADERS = {
    "business_email": "Person Account: Business Email",
    "preferred_email": "Person Account: Preferred Email",
    "industry": "Industry",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate an NY sample; optionally enrich it via RocketReach."
    )
    parser.add_argument("--input", required=True, type=Path, help="Private .xlsx input")
    parser.add_argument("--output", type=Path, help="Output .xlsx for live results")
    parser.add_argument("--sheet", help="Worksheet name (defaults to active sheet)")
    parser.add_argument("--limit", type=int, default=100, help="Maximum people to process")
    parser.add_argument(
        "--live",
        action="store_true",
        help="Call RocketReach. Without this flag, validation only is performed.",
    )
    parser.add_argument(
        "--confirm-credits",
        default="NO",
        help='Required for live calls: --confirm-credits YES',
    )
    parser.add_argument(
        "--region",
        default="tri-state",
        choices=("all", "tri-state", "massachusetts", "west-coast"),
        help="Only include alumni whose home state matches this rollout region",
    )
    parser.add_argument("--min-reunion-year", type=int, help="Minimum reunion/class year (inclusive)")
    parser.add_argument("--max-reunion-year", type=int, help="Maximum reunion/class year (inclusive)")
    parser.add_argument(
        "--sleep-seconds",
        type=float,
        default=1.0,
        help="Pause between live RocketReach lookups (helps stay near ~500/day)",
    )
    parser.add_argument(
        "--home-state",
        help="Optional single state within the region (e.g. NY, then NJ, then CT for tri-state)",
    )
    parser.add_argument(
        "--after-source-row",
        type=int,
        default=0,
        help="Only process sheet rows after this Excel row number (for batch continuation)",
    )
    return parser.parse_args()


def require_real_xlsx(path: Path) -> None:
    if not path.exists():
        raise ValueError(f"Input file does not exist: {path}")
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Input must be an .xlsx workbook.")
    with path.open("rb") as handle:
        signature = handle.read(4)
    if signature != b"PK\x03\x04":
        raise ValueError(
            "This is not a real Excel workbook. It appears to be an Excel web page "
            "saved with an .xlsx name. In Excel Online use File > Save As > "
            "Download a Copy, then select that downloaded file."
        )


def normalized(value: Any) -> str:
    return " ".join(str(value or "").strip().lower().split())


REGION_STATES: dict[str, set[str]] = {
    "tri-state": {"ny", "new york", "nj", "new jersey", "ct", "connecticut"},
    "massachusetts": {"ma", "massachusetts"},
    "west-coast": {"ca", "california", "wa", "washington", "or", "oregon", "nv", "nevada"},
}


def state_is_ny(value: Any) -> bool:
    return normalized(value) in {"ny", "new york"}


def normalize_state_code(value: Any) -> str:
    raw = str(value or "").strip().upper()
    if len(raw) == 2:
        return raw
    mapping = {
        "NEW YORK": "NY",
        "NEW JERSEY": "NJ",
        "CONNECTICUT": "CT",
        "MASSACHUSETTS": "MA",
        "CALIFORNIA": "CA",
        "WASHINGTON": "WA",
        "OREGON": "OR",
        "NEVADA": "NV",
    }
    return mapping.get(raw, raw)


def state_in_region(value: Any, region: str) -> bool:
    if region == "all":
        return True
    allowed = REGION_STATES.get(region)
    if not allowed:
        raise ValueError(f"Unknown region '{region}'. Choose from: all, tri-state, massachusetts, west-coast")
    return normalized(value) in allowed


def state_matches_home_filter(value: Any, home_state: str | None) -> bool:
    if not home_state:
        return True
    return normalize_state_code(value) == normalize_state_code(home_state)


def reunion_year(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def reunion_in_range(value: Any, min_year: int | None, max_year: int | None) -> bool:
    year = reunion_year(value)
    if year is None:
        return min_year is None and max_year is None
    if min_year is not None and year < min_year:
        return False
    if max_year is not None and year > max_year:
        return False
    return True


def changed(old: Any, new: Any) -> str:
    if not normalized(new):
        return ""
    return "YES" if normalized(old) != normalized(new) else "NO"


def similarity(left: Any, right: Any) -> float:
    return round(100 * SequenceMatcher(None, normalized(left), normalized(right)).ratio(), 1)


def resolve_sheet_name(workbook, sheet_name: str | None) -> str | None:
    if not sheet_name:
        return None
    wanted = sheet_name.strip()
    if wanted in workbook.sheetnames:
        return wanted
    by_trim = {name.strip(): name for name in workbook.sheetnames}
    if wanted in by_trim:
        return by_trim[wanted]
    raise ValueError(
        f"Worksheet '{sheet_name}' was not found. Available: {workbook.sheetnames}"
    )


def read_sample(
    path: Path,
    sheet_name: str | None,
    limit: int,
    region: str = "tri-state",
    min_reunion_year: int | None = None,
    max_reunion_year: int | None = None,
    home_state: str | None = None,
    after_source_row: int = 0,
):
    require_real_xlsx(path)
    workbook = load_workbook(path, read_only=True, data_only=True)
    resolved = resolve_sheet_name(workbook, sheet_name)
    if resolved:
        sheet = workbook[resolved]
    else:
        sheet = workbook.active

    header_values = [str(cell.value or "").strip() for cell in sheet[1]]
    indexes = {name: header_values.index(name) for name in HEADERS.values() if name in header_values}
    missing = [name for name in HEADERS.values() if name not in indexes]
    if missing:
        raise ValueError("Required columns are missing:\n- " + "\n- ".join(missing))
    for key, column in OPTIONAL_INPUT_HEADERS.items():
        if column in header_values:
            indexes[column] = header_values.index(column)

    people = []
    for source_row, values in enumerate(
        sheet.iter_rows(min_row=2, values_only=True), start=2
    ):
        record = {header: values[index] for header, index in indexes.items()}
        if not any(
            record.get(HEADERS[key])
            for key in ("first_name", "last_name", "account_name")
        ):
            continue
        home = record.get(HEADERS["home_state"])
        if not state_in_region(home, region):
            continue
        if not state_matches_home_filter(home, home_state):
            continue
        if not reunion_in_range(
            record.get(HEADERS["reunion_year"]),
            min_reunion_year,
            max_reunion_year,
        ):
            continue
        if after_source_row and source_row <= after_source_row:
            continue
        record["_source_row"] = source_row
        record["_all_values"] = list(values[: len(header_values)])
        people.append(record)
        if len(people) >= limit:
            break

    workbook.close()
    if not people:
        raise ValueError(
            "No alumni rows matched the region/year filters. "
            "Try --region all or widen the reunion year range."
        )

    out_of_region = [
        row["_source_row"]
        for row in people
        if not state_in_region(row[HEADERS["home_state"]], region)
    ]
    return header_values, people, out_of_region


def display_validation(
    headers: list[str],
    people: list[dict[str, Any]],
    mismatches: list[int],
    region: str,
    min_reunion_year: int | None,
    max_reunion_year: int | None,
    home_state: str | None = None,
) -> None:
    print("\nPILOT VALIDATION")
    print(f"People loaded       : {len(people)}")
    print(f"Original columns    : {len(headers)}")
    print(f"Region filter       : {region}")
    if home_state:
        print(f"Home state filter   : {normalize_state_code(home_state)}")
    if min_reunion_year is not None or max_reunion_year is not None:
        print(f"Reunion year range  : {min_reunion_year or '...'} - {max_reunion_year or '...'}")
    print(f"Region mismatches   : {len(mismatches)}")
    if mismatches:
        print(f"Mismatch source rows: {mismatches[:20]}")
    print(f"Validation result   : {'PASS' if not mismatches else 'FAIL'}")
    if people:
        print("Selection order     : Excel sheet row 2 -> end (top to bottom, not random)")
        print("First rows selected   :")
        for person in people[:10]:
            year = reunion_year(person.get(HEADERS["reunion_year"]))
            year_label = year if year is not None else "?"
            state = normalize_state_code(person.get(HEADERS["home_state"]))
            print(
                f"  - row {person['_source_row']}: {legacy_name(person)} "
                f"(class {year_label}, {state})"
            )
        if len(people) > 10:
            print(f"  ... and {len(people) - 10} more in this batch")


def _parse_rr_json(response: requests.Response) -> dict[str, Any]:
    try:
        payload = response.json()
    except ValueError:
        return {"status": "error", "message": (response.text or "")[:500]}
    if isinstance(payload, dict):
        return payload
    if isinstance(payload, list) and payload:
        first = payload[0]
        return first if isinstance(first, dict) else {"status": "error", "message": str(first)}
    return {"status": "error", "message": "Unexpected RocketReach response shape"}


def lookup_person(api_key: str, name: str, company: str) -> dict[str, Any]:
    """Match Apps Script: always send current_employer (may be empty)."""
    headers = {"accept": "application/json", "Api-Key": api_key}
    params = {"name": name, "current_employer": company or ""}

    response = requests.get(LOOKUP_URL, headers=headers, params=params, timeout=45)
    if response.status_code == 404:
        body = _parse_rr_json(response)
        body.setdefault("status", "not found")
        return body
    if response.status_code == 400:
        body = _parse_rr_json(response)
        message = json.dumps(body).lower()
        if "current_employer" in message:
            return {
                "status": "skipped_no_employer",
                "message": "RocketReach requires current_employer on the lookup request",
            }
        response.raise_for_status()
    if response.status_code != 200:
        response.raise_for_status()
    result = _parse_rr_json(response)

    attempts = 0
    while normalized(result.get("status")) in {"searching", "progress", "waiting"}:
        if attempts >= 4:
            break
        time.sleep(2)
        lookup_id = result.get("id")
        if not lookup_id:
            break
        response = requests.get(
            CHECK_STATUS_URL,
            headers=headers,
            params={"ids": lookup_id},
            timeout=45,
        )
        if response.status_code != 200:
            break
        checked = _parse_rr_json(response)
        if checked.get("error"):
            break
        result = checked
        attempts += 1
    return result


def legacy_name(record: dict[str, Any]) -> str:
    return str(
        record.get(HEADERS["account_name"])
        or f"{record.get(HEADERS['first_name'], '')} {record.get(HEADERS['last_name'], '')}".strip()
    ).strip()


def legacy_email(record: dict[str, Any]) -> str:
    for column in (
        OPTIONAL_INPUT_HEADERS["preferred_email"],
        OPTIONAL_INPUT_HEADERS["business_email"],
    ):
        value = record.get(column)
        if value and str(value).strip():
            return str(value).strip()
    return ""


def extract_rr_email(result: dict[str, Any]) -> str:
    for key in (
        "recommended_email",
        "recommended_professional_email",
        "current_work_email",
        "email",
    ):
        value = result.get(key)
        if value and str(value).strip():
            return str(value).strip()
    emails = result.get("emails")
    if isinstance(emails, list):
        for item in emails:
            if isinstance(item, dict):
                addr = item.get("email") or item.get("address")
                if addr:
                    return str(addr).strip()
            elif item:
                return str(item).strip()
    return ""


def extract_rr_industry(result: dict[str, Any], record: dict[str, Any]) -> str:
    for key in ("industry", "company_industry", "current_employer_industry"):
        value = result.get(key)
        if value and str(value).strip():
            return str(value).strip()
    legacy = record.get(OPTIONAL_INPUT_HEADERS["industry"], "")
    return str(legacy or "").strip()


def pick_email(record: dict[str, Any], result: dict[str, Any]) -> str:
    return extract_rr_email(result) or legacy_email(record)


def scrape_focus_row(
    record: dict[str, Any],
    result: dict[str, Any],
    status: str,
    queue: str,
) -> list[Any]:
    year = reunion_year(record.get(HEADERS["reunion_year"]))
    company = str(
        result.get("current_employer") or record.get(HEADERS["current_employer"]) or ""
    ).strip()
    return [
        legacy_name(record),
        year or "",
        company,
        extract_rr_industry(result, record),
        pick_email(record, result),
        status,
        queue,
        "",
    ]


def review_queue(record: dict[str, Any], result: dict[str, Any]) -> tuple[float, str]:
    old_name = record.get(HEADERS["account_name"]) or (
        f"{record.get(HEADERS['first_name'], '')} {record.get(HEADERS['last_name'], '')}"
    )
    score = similarity(old_name, result.get("name"))
    status = normalized(result.get("status"))
    if status in {"not found", "not_found", "failed", "error"} or not result.get("name"):
        return score, "NEEDS_RESEARCH"
    if score >= 92:
        return score, "QUICK_REVIEW"
    return score, "MANUAL_REVIEW"


def cell_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def lookup_outcome(status: str, error: str, result: dict[str, Any]) -> str:
    status = cell_text(status)
    error = cell_text(error)
    if normalized(status) in {"skipped_no_employer", "skipped"}:
        return "SKIPPED_NO_EMPLOYER"
    if error or normalized(status) in {"error", "failed"}:
        return "ERROR"
    if normalized(status) in {"not found", "not_found"}:
        return "NOT_FOUND"
    if not result.get("name"):
        return "NOT_FOUND"
    return "FOUND"


def summarize_run(processed: list[tuple[dict[str, Any], list[Any], list[Any]]]) -> dict[str, Any]:
    counts = {
        "total": len(processed),
        "found": 0,
        "not_found": 0,
        "error": 0,
        "skipped_no_employer": 0,
        "quick_review": 0,
        "manual_review": 0,
        "needs_research": 0,
    }
    for _record, result_values, _focus in processed:
        status = cell_text(result_values[RV_LOOKUP_STATUS] if len(result_values) > RV_LOOKUP_STATUS else "")
        error = cell_text(result_values[RV_LOOKUP_ERROR] if len(result_values) > RV_LOOKUP_ERROR else "")
        rr_name = result_values[RV_RR_NAME] if len(result_values) > RV_RR_NAME else ""
        queue = cell_text(result_values[RV_REVIEW_QUEUE] if len(result_values) > RV_REVIEW_QUEUE else "")
        outcome = lookup_outcome(status, error, {"name": rr_name}).lower()
        if outcome == "found":
            counts["found"] += 1
        elif outcome == "not_found":
            counts["not_found"] += 1
        elif outcome == "skipped_no_employer":
            counts["skipped_no_employer"] += 1
        else:
            counts["error"] += 1
        qkey = queue.lower()
        if qkey == "quick_review":
            counts["quick_review"] += 1
        elif qkey == "manual_review":
            counts["manual_review"] += 1
        elif qkey == "needs_research":
            counts["needs_research"] += 1
    total = counts["total"] or 1
    found = counts.get("found", 0)
    skipped = counts.get("skipped_no_employer", 0)
    attempted = total - skipped or 1
    failed = counts.get("not_found", 0) + counts.get("error", 0)
    review = (
        counts["quick_review"] + counts["manual_review"] + counts["needs_research"]
    )
    counts["pct_rr_success"] = round(100 * found / attempted, 1)
    counts["pct_rr_failed"] = round(100 * failed / attempted, 1)
    counts["pct_flagged_for_review"] = round(100 * review / total, 1)
    counts["pct_needs_attention"] = round(100 * (total - found) / total, 1)
    return counts


def print_run_summary(counts: dict[str, Any]) -> None:
    print("\nRUN SUMMARY")
    print(f"Total processed     : {counts['total']}")
    print(
        f"Found (RR match)    : {counts.get('found', 0)} "
        f"({counts.get('pct_rr_success', 0)}% of lookups attempted)"
    )
    print(f"Not found           : {counts.get('not_found', 0)}")
    print(f"Errors              : {counts.get('error', 0)}")
    print(f"Skipped (no employer): {counts.get('skipped_no_employer', 0)} (no API call)")
    print(
        f"Failed (not found+error): {counts.get('not_found', 0) + counts.get('error', 0)} "
        f"({counts.get('pct_rr_failed', 0)}% of lookups attempted)"
    )
    print(
        f"Needs attention (no RR match): {counts.get('pct_needs_attention', 0)}%"
    )
    print(f"Quick review queue  : {counts['quick_review']}")
    print(f"Manual review queue : {counts['manual_review']}")
    print(f"Needs research      : {counts['needs_research']}")
    print(
        f"Rows in a review queue: {counts['quick_review'] + counts['manual_review'] + counts['needs_research']} "
        f"({counts.get('pct_flagged_for_review', 0)}%)"
    )


def append_run_summary_sheet(
    output_path: Path,
    counts: dict[str, Any],
    region: str,
    home_state: str | None,
    min_reunion_year: int | None,
    max_reunion_year: int | None,
    processed: list[tuple[dict[str, Any], list[Any], list[Any]]],
) -> None:
    workbook = load_workbook(output_path)
    for sheet_name in ("Run Summary", "Success", "Needs review"):
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]

    summary = workbook.create_sheet("Run Summary", 0)
    rows = [
        ("metric", "value"),
        ("total_processed", counts["total"]),
        ("found_rr_match", counts.get("found", 0)),
        ("pct_rr_success", counts.get("pct_rr_success", 0)),
        ("not_found", counts.get("not_found", 0)),
        ("skipped_no_employer", counts.get("skipped_no_employer", 0)),
        ("errors", counts.get("error", 0)),
        ("pct_rr_failed", counts.get("pct_rr_failed", 0)),
        ("pct_needs_attention", counts.get("pct_needs_attention", 0)),
        ("quick_review", counts["quick_review"]),
        ("manual_review", counts["manual_review"]),
        ("needs_research", counts["needs_research"]),
        ("pct_flagged_for_review", counts.get("pct_flagged_for_review", 0)),
        ("region_filter", region),
        ("home_state_filter", normalize_state_code(home_state) if home_state else "all"),
        ("min_reunion_year", min_reunion_year if min_reunion_year is not None else ""),
        ("max_reunion_year", max_reunion_year if max_reunion_year is not None else ""),
        ("generated_utc", datetime.now(timezone.utc).isoformat()),
    ]
    for row in rows:
        summary.append(row)

    success = workbook.create_sheet("Success")
    needs = workbook.create_sheet("Needs review")
    success.append(SCRAPE_FOCUS_HEADERS)
    needs.append(SCRAPE_FOCUS_HEADERS)
    for _record, result_values, focus_row in processed:
        status = cell_text(result_values[RV_LOOKUP_STATUS] if len(result_values) > RV_LOOKUP_STATUS else "")
        error = cell_text(result_values[RV_LOOKUP_ERROR] if len(result_values) > RV_LOOKUP_ERROR else "")
        rr_name = result_values[RV_RR_NAME] if len(result_values) > RV_RR_NAME else ""
        if lookup_outcome(status, error, {"name": rr_name}) == "FOUND":
            success.append(focus_row)
        else:
            needs.append(focus_row)

    workbook.save(output_path)
    workbook.close()


def build_output(
    output_path: Path,
    original_headers: list[str],
    processed: list[tuple[dict[str, Any], list[Any], list[Any]]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Pilot Results"
    sheet.append(original_headers + RESULT_HEADERS)
    focus = workbook.create_sheet("Scrape Focus")
    focus.append(SCRAPE_FOCUS_HEADERS)
    for record, result_values, focus_row in processed:
        sheet.append(record["_all_values"] + result_values)
        focus.append(focus_row)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    focus.freeze_panes = "A2"
    focus.auto_filter.ref = focus.dimensions
    workbook.save(output_path)


def preflight_output_path(output_path: Path, original_headers: list[str]) -> None:
    """Verify Excel output is writable before any RocketReach credits are used."""
    try:
        build_output(output_path.resolve(), original_headers, [])
    except PermissionError as exc:
        raise ValueError(
            f"Cannot write to output file (close it in Excel first): {output_path}\n"
            "Use a new --output filename if needed, e.g. "
            "NY_Pilot_Results_2.xlsx"
        ) from exc
    except OSError as exc:
        if getattr(exc, "winerror", None) == 32 or getattr(exc, "errno", None) == 13:
            raise ValueError(
                f"Cannot write to output file (close it in Excel first): {output_path}\n"
                "Use a new --output filename if needed."
            ) from exc
        raise


def append_backup_jsonl(
    backup_path: Path,
    record: dict[str, Any],
    result_values: list[Any],
    focus_row: list[Any],
) -> None:
    payload = {
        "source_row": record.get("_source_row"),
        "legacy_name": legacy_name(record),
        "result": result_values,
        "scrape_focus": focus_row,
        "saved_utc": datetime.now(timezone.utc).isoformat(),
    }
    with backup_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(payload, default=str) + "\n")


def post_scrape_status(payload: dict[str, Any]) -> None:
    """Tell the dashboard how far the live batch has gotten. Never fail the batch."""
    try:
        requests.post(SCRAPE_STATUS_URL, json={"type": "scrape_status", **payload}, timeout=20)
    except Exception:
        return


def run_live(
    people: list[dict[str, Any]],
    original_headers: list[str],
    output_path: Path,
    sleep_seconds: float = 1.0,
    region: str = "tri-state",
    home_state: str | None = None,
    min_reunion_year: int | None = None,
    max_reunion_year: int | None = None,
) -> None:
    load_dotenv(Path(__file__).with_name(".env"))
    api_key = os.getenv("ROCKETREACH_API_KEY", "").strip()
    if not api_key:
        raise ValueError(
            "ROCKETREACH_API_KEY is missing. Copy .env.example to worker/.env "
            "and add the key there."
        )

    output_path = output_path.resolve()
    preflight_output_path(output_path, original_headers)
    backup_path = output_path.with_suffix(output_path.suffix + ".jsonl")
    print(f"Output file OK     : {output_path}")
    print(f"Backup log         : {backup_path}")

    label = f"{normalize_state_code(home_state) if home_state else region} {min_reunion_year or ''}-{max_reunion_year or ''}".strip()
    running = {"found": 0, "not_found": 0, "skipped": 0, "errors": 0}
    post_scrape_status({
        "phase": "running",
        "label": label,
        "current_name": "",
        "processed": 0,
        "total": len(people),
        "found": 0,
        "not_found": 0,
        "skipped": 0,
        "errors": 0,
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "message": f"Starting {len(people)} lookups.",
    })

    processed: list[tuple[dict[str, Any], list[Any], list[Any]]] = []
    total_start = time.perf_counter()
    for position, record in enumerate(people, start=1):
        name = legacy_name(record)
        company = str(record.get(HEADERS["current_employer"]) or "").strip()
        print(f"[{position}/{len(people)}] RocketReach lookup: {name}")
        started = time.perf_counter()
        error = ""
        if not company:
            result = {}
            status = "skipped_no_employer"
            error = "No current employer on file (RocketReach requires current_employer)"
            elapsed = 0.0
        else:
            try:
                result = lookup_person(api_key, str(name), company)
                status = str(result.get("status") or "returned")
            except Exception as exc:
                result = {}
                status = "error"
                error = str(exc)
            elapsed = round(time.perf_counter() - started, 2)
        name_score, queue = review_queue(record, result)
        timestamp = datetime.now(timezone.utc).isoformat()
        result_values = [
            f"LEGACY-{record['_source_row']}",
            timestamp,
            elapsed,
            status,
            error,
            result.get("id", ""),
            result.get("name", ""),
            result.get("current_employer", ""),
            extract_rr_industry(result, record),
            extract_rr_email(result),
            result.get("location", ""),
            result.get("linkedin_url", ""),
            name_score,
            changed(record.get(HEADERS["current_employer"]), result.get("current_employer")),
            changed(
                f"{record.get(HEADERS['home_city'], '')}, "
                f"{record.get(HEADERS['home_state'], '')}",
                result.get("location"),
            ),
            queue,
            "",
            "",
        ]
        focus_row = scrape_focus_row(record, result, status, queue)
        focus_row[-1] = lookup_outcome(status, error, result)
        outcome = focus_row[-1]
        if outcome == "FOUND":
            running["found"] += 1
        elif outcome == "NOT_FOUND":
            running["not_found"] += 1
        elif outcome == "SKIPPED_NO_EMPLOYER":
            running["skipped"] += 1
        else:
            running["errors"] += 1
        post_scrape_status({
            "phase": "running",
            "label": label,
            "current_name": name,
            "processed": position,
            "total": len(people),
            "found": running["found"],
            "not_found": running["not_found"],
            "skipped": running["skipped"],
            "errors": running["errors"],
            "updated_at": timestamp,
            "message": f"Looking up {name}.",
        })
        processed.append((record, result_values, focus_row))
        append_backup_jsonl(backup_path, record, result_values, focus_row)
        try:
            build_output(output_path, original_headers, processed)
        except (PermissionError, OSError) as exc:
            raise ValueError(
                f"RocketReach lookup completed for {name}, but Excel save failed "
                f"({exc}). Recovery data is in {backup_path}. "
                f"Close {output_path.name} in Excel and re-run, or use a new --output path."
            ) from exc
        if position < len(people) and sleep_seconds > 0:
            time.sleep(sleep_seconds)

    elapsed_values: list[float] = []
    for _record, result_values, _focus in processed:
        if len(result_values) <= 2:
            continue
        try:
            elapsed_values.append(float(result_values[2]))
        except (TypeError, ValueError):
            continue
    total_elapsed = round(time.perf_counter() - total_start, 2)
    counts = summarize_run(processed)
    print_run_summary(counts)
    post_scrape_status({
        "phase": "done",
        "label": label,
        "current_name": "",
        "processed": counts["total"],
        "total": len(people),
        "found": counts.get("found", 0),
        "not_found": counts.get("not_found", 0),
        "skipped": counts.get("skipped_no_employer", 0),
        "errors": counts.get("error", 0),
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "message": "Scrape finished.",
    })
    append_run_summary_sheet(
        output_path,
        counts,
        region,
        home_state,
        min_reunion_year,
        max_reunion_year,
        processed,
    )
    print(f"\nOutput              : {output_path}")
    print(f"Total elapsed sec   : {total_elapsed}")
    if elapsed_values:
        avg = round(sum(elapsed_values) / len(elapsed_values), 2)
        per_100 = round(avg * 100 / 60, 1)
        print(f"Avg lookup sec/rec  : {avg}")
        print(f"Est. API min/100 rec: {per_100} (lookup only; add manual verification separately)")


def main() -> int:
    args = parse_args()
    try:
        headers, people, mismatches = read_sample(
            args.input.resolve(),
            args.sheet,
            args.limit,
            args.region,
            args.min_reunion_year,
            args.max_reunion_year,
            args.home_state,
            args.after_source_row,
        )
        display_validation(
            headers,
            people,
            mismatches,
            args.region,
            args.min_reunion_year,
            args.max_reunion_year,
            args.home_state,
        )
        if mismatches:
            print("\nNo API calls made because validation failed.")
            return 2
        if not args.live:
            print("\nDRY RUN ONLY: no RocketReach calls and no UI/Sheet changes.")
            return 0
        if args.confirm_credits != "YES":
            raise ValueError(
                "Live mode can consume RocketReach credits. Re-run with "
                "--live --confirm-credits YES after checking --limit."
            )
        if not args.output:
            raise ValueError("--output is required in live mode.")
        preflight_output_path(args.output.resolve(), headers)
        run_live(
            people,
            headers,
            args.output.resolve(),
            args.sleep_seconds,
            args.region,
            args.home_state,
            args.min_reunion_year,
            args.max_reunion_year,
        )
        return 0
    except (ValueError, OSError) as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        print("No UI or Google Sheet changes were made.", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
