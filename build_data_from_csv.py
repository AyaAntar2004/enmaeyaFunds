"""
build_data_from_csv.py

Same as build_data.py, but reads the flattened CSV export instead of the
.ods workbook. Use this when the source of truth is a CSV (e.g. a file
downloaded/edited outside the .ods), such as:

    C:\\Users\\AyaAntar\\Downloads\\funding updated.csv

The CSV has no real hyperlinks (unlike the .ods, where "Source / Link"
cells carry a hyperlink), so sourceLink falls back to the cell's display
text (e.g. "Open source \u25b8").

Run:
    python build_data_from_csv.py "C:\\Users\\AyaAntar\\Downloads\\funding updated.csv"
"""

import csv
import json
import re
import sys
from datetime import datetime
from pathlib import Path

from build_data import (
    EXPECTED_HEADER,
    FIELD_NAMES,
    clean,
    split_list,
    normalize_date,
    parse_amount_usd_millions,
    classify_region,
    slugify_code,
    MINISTRY_MAP,
    DEFAULT_SECTOR,
)

ROOT = Path(__file__).resolve().parent
DATA_JSON_PATH = ROOT / "data.json"

DEFAULT_CSV_PATH = Path(r"C:\Users\AyaAntar\Downloads\funding updated.csv")

_MDY_RE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")


def normalize_date_flex(raw):
    """Like build_data.normalize_date, but also accepts the M/D/YYYY
    text Excel/Sheets produces when a CSV is exported (the .ods path
    never hits this since pandas reads real date cells as Timestamps)."""
    text = clean(raw)
    if not text:
        return None
    m = _MDY_RE.match(text)
    if m:
        month, day, year = (int(g) for g in m.groups())
        return datetime(year, month, day).strftime("%Y-%m-%d")
    return normalize_date(raw)


def load_csv_rows(csv_path):
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))

    header_idx = None
    for i, row in enumerate(rows):
        if row and row[0].strip() == "Ministry":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Could not find a 'Ministry' header row in the CSV")

    header_vals = [c.strip() for c in rows[header_idx][: len(EXPECTED_HEADER)]]
    if header_vals != EXPECTED_HEADER:
        print(f"  warning: CSV header differs from expected:\n    got: {header_vals}")

    data_rows = [
        r for r in rows[header_idx + 1 :]
        if r and r[0].strip()
    ]
    # Pad/truncate every row to exactly len(FIELD_NAMES) columns.
    n = len(FIELD_NAMES)
    normalized = []
    for r in data_rows:
        r = list(r[:n]) + [""] * (n - len(r))
        normalized.append(r)
    return normalized


def build_records(data_rows):
    records = []
    unmapped = set()
    dupes_seen = {}

    for row in data_rows:
        rec = dict(zip(FIELD_NAMES, row))
        ministry = clean(rec["ministry"]) or "Not specified"

        code, sector = MINISTRY_MAP.get(ministry, (None, None))
        if code is None:
            code, sector = slugify_code(ministry), DEFAULT_SECTOR
            unmapped.add(ministry)

        amount_raw = clean(rec["amountRaw"]) or "Not specified"
        project = clean(rec["project"]) or ""

        key = (ministry, project)
        dupes_seen[key] = dupes_seen.get(key, 0) + 1

        source_link = clean(rec["sourceLink_display"]) or ""
        source_link = source_link.replace("\u25b8", "").strip()

        records.append({
            "ministry": ministry,
            "ministryCode": code,
            "sector": sector,
            "project": project,
            "status": clean(rec["status"]) or "Unknown",
            "funders": split_list(rec["funders_raw"]),
            "amountRaw": amount_raw,
            "amountUSDm": parse_amount_usd_millions(amount_raw),
            "partners": split_list(rec["partners_raw"]),
            "region": classify_region(rec["region_raw"]),
            "date": normalize_date_flex(rec["published_date_raw"]),
            "confidence": clean(rec["confidence"]) or "Unknown",
            "sourceFeed": clean(rec["sourceFeed"]) or "",
            "description": clean(rec["description"]) or "",
            "sourceLink": source_link,
            "note": "",
            "fundingDate": normalize_date_flex(rec["funding_date_raw"]),
            "dateBasis": clean(rec["date_basis"]) or "",
            "dateEvidence": clean(rec["date_evidence"]) or "",
            "projectPeriod": clean(rec["project_period"]) or "",
            "in2025_26": clean(rec["in_2025_26"]) or "",
            "enmaeyaLink": clean(rec["enmaeya_link"]) or "",
            "origin": clean(rec["origin"]) or "",
            "sourceChecked": clean(rec["source_checked"]) or "",
            "livePageEvidence": clean(rec["live_page_evidence"]) or "",
            "fundingType": clean(rec["funding_type"]) or "",
        })

    if unmapped:
        print(f"  warning: {len(unmapped)} ministry name(s) not in MINISTRY_MAP, "
              f"defaulted to sector '{DEFAULT_SECTOR}':")
        for m in sorted(unmapped):
            print(f"    - {m}")

    dupes = {k: v for k, v in dupes_seen.items() if v > 1}
    if dupes:
        print(f"  note: {len(dupes)} ministry+project pair(s) appear more than once "
              f"(kept as separate records, not merged):")
        for (m, p), n in sorted(dupes.items()):
            print(f"    - [{n}x] {m} \u2014 {p}")

    return records


def main():
    csv_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_CSV_PATH
    print(f"Reading {csv_path}")

    data_rows = load_csv_rows(csv_path)
    print(f"  {len(data_rows)} records found")

    records = build_records(data_rows)
    print(f"  built {len(records)} records")

    print(f"Updating {DATA_JSON_PATH}")
    with open(DATA_JSON_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    before = len(data.get("ministryProjects", []))
    data["ministryProjects"] = records
    with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")
    print(f"  ministryProjects: {before} -> {len(records)} records")


if __name__ == "__main__":
    main()
