"""
sync_daily.py

Rebuilds data.json from:
  1. funding_updated.ods          - the hand-checked base records, and
  2. daily/Funding_YYYY-MM-DD.xlsx - the AI-extracted funding records that the
     daily news scraper (fatimaelkhishen/Funds-Enmaeya) copies in every morning.

Sectors: base records take their sector from sector_overrides.csv (made once by
classify_sectors.py, editable in Excel) and daily records from the "Sector" column
the AI writes; anything without one falls back to build_data's ministry table.

Daily records are added on top of the base. A daily record is skipped when its
ministry is not one of the Lebanese ministries in build_data.MINISTRY_MAP (this
drops "Not specified", foreign ministries and misspelled names), or when its
source link, or its ministry + project name, is already in the data (from the
base or from an earlier day), so re-running this script never duplicates
anything. Running it again with the same files gives the same data.json.

Then it applies data_fixes.py: English names for funders/partners (no Arabic
entries in the Funder filter) and the approved rows of corrections.csv.

It also regenerates downloads/funding-data.xlsx (the "Download Excel" file) from
the same records, and moves meta.lastRecord and the end of meta.reportingPeriod forward to the
newest published date in the data.

Run (from the repository folder):
    pip install -r requirements.txt
    python sync_daily.py

The daily scraper's GitHub workflow runs this automatically and pushes the
result, which redeploys the site.
"""

import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

import build_data as bd
from classify_sectors import SECTOR_IDS, load_overrides, record_key
import data_fixes

ROOT = Path(__file__).resolve().parent
DAILY_DIR = ROOT / "daily"

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def to_iso(value):
    """'7/21/2026' (M/D/YYYY, as the daily files write it), '2026-07-21',
    a datetime or a pandas Timestamp -> '2026-07-21'. Anything else as-is."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if not text:
        return None
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", text)
    if m:
        mo, d, y = map(int, m.groups())
        return f"{y:04d}-{mo:02d}-{d:02d}"
    return bd.normalize_date(text)


def read_daily_file(path):
    """Return (DataFrame with the 22 template columns in order, list of real
    source URLs, list of dashboard sector ids or None) for one daily Funding_*.xlsx file."""
    wb = load_workbook(path)
    ws = wb.active
    header_row = None
    for r in range(1, min(ws.max_row, 20) + 1):
        if str(ws.cell(r, 1).value or "").strip() == "Ministry":
            header_row = r
            break
    if header_row is None:
        print(f"  warning: no 'Ministry' header row in {path.name}, skipped")
        return None, [], []
    header = [str(ws.cell(header_row, c).value or "").strip() for c in range(1, ws.max_column + 1)]
    missing = [h for h in bd.EXPECTED_HEADER if h not in header]
    if missing:
        print(f"  warning: {path.name} is missing column(s) {missing}, skipped")
        return None, [], []
    idx = [header.index(h) + 1 for h in bd.EXPECTED_HEADER]
    link_col = header.index("Source / Link") + 1
    # optional extra column (newer daily files): the AI's sector for each project
    sector_col = header.index("Sector") + 1 if "Sector" in header else None

    rows, links, sectors = [], [], []
    for r in range(header_row + 1, ws.max_row + 1):
        if not str(ws.cell(r, 1).value or "").strip():
            continue
        values = [ws.cell(r, c).value for c in idx]
        # dates come out of the daily file as M/D/YYYY text; build_records
        # expects ISO text or Timestamps, so convert them here
        for name in ("Published Date", "Funding Date"):
            i = bd.EXPECTED_HEADER.index(name)
            values[i] = to_iso(values[i])
        rows.append(values)
        cell = ws.cell(r, link_col)
        links.append(cell.hyperlink.target if cell.hyperlink else (cell.value or ""))
        label = str(ws.cell(r, sector_col).value or "").strip() if sector_col else ""
        sectors.append(SECTOR_IDS.get(label))
    return pd.DataFrame(rows), links, sectors


def norm_key(text):
    return re.sub(r"\s+", " ", str(text or "")).strip().lower()


def long_date(iso):
    d = datetime.strptime(iso, "%Y-%m-%d")
    return f"{d.day} {d.strftime('%B %Y')}"


def main():
    # 1. base records from the .ods (this also restores the real source URLs)
    print(f"Reading {bd.ODS_PATH.name}")
    xl = pd.ExcelFile(bd.ODS_PATH, engine="odf")
    header_idx, data_rows = bd.load_sheet_rows(xl, bd.SHEET_NAME)
    real_links = bd.extract_column_hrefs(bd.ODS_PATH, bd.SOURCE_LINK_COL_IDX, header_idx + 1)
    records = bd.build_records(data_rows, real_links)

    # sectors by project (sector_overrides.csv, made by classify_sectors.py) instead of
    # build_data's one-sector-per-ministry table
    overrides = load_overrides()
    moved = 0
    for r in records:
        sid = overrides.get(record_key(r["ministry"], r["project"], r["sourceLink"]))
        if sid:
            moved += sid != r["sector"]
            r["sector"] = sid
    if overrides:
        print(f"  sector_overrides.csv: {len(overrides)} row(s), {moved} base record(s) moved sector")
    print(f"  base: {len(records)} records")

    seen_links = {norm_key(r["sourceLink"]) for r in records if str(r["sourceLink"]).startswith("http")}
    seen_pairs = {(norm_key(r["ministry"]), norm_key(r["project"])) for r in records}

    # 2. daily AI records, oldest day first
    files = sorted(DAILY_DIR.glob("Funding_*.xlsx")) if DAILY_DIR.is_dir() else []
    print(f"Reading {len(files)} daily file(s) from {DAILY_DIR.name}/")
    added = skipped = dropped = 0
    for path in files:
        df, links, day_sectors = read_daily_file(path)
        if df is None or df.empty:
            continue
        day_records = bd.build_records(df, links)
        for rec, sid in zip(day_records, day_sectors):
            if sid:
                rec["sector"] = sid
        day_added = 0
        for rec in day_records:
            if rec["ministry"] not in bd.MINISTRY_MAP or rec["ministry"] == "Not specified":
                dropped += 1
                continue
            link = norm_key(rec["sourceLink"])
            pair = (norm_key(rec["ministry"]), norm_key(rec["project"]))
            if (link and link in seen_links) or pair in seen_pairs:
                skipped += 1
                continue
            seen_links.add(link)
            seen_pairs.add(pair)
            records.append(rec)
            day_added += 1
        added += day_added
        print(f"  {path.name}: {len(day_records)} record(s), {day_added} new")
    print(f"  daily: {added} added, {skipped} skipped as already present, "
          f"{dropped} dropped (no Lebanese ministry from the dashboard's list)")

    # 3. clean-ups on every record: English funder/partner names, approved
    #    corrections (corrections.csv), placeholder text in date fields
    data_fixes.apply_corrections(records)
    data_fixes.fix_names(records)
    data_fixes.tidy(records)

    # 4. write data.json, keeping everything else in it as it was
    data_path = bd.DATA_JSON_PATH
    with open(data_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    before = len(data.get("ministryProjects", []))
    data["ministryProjects"] = records

    dates = sorted(r["date"] for r in records if r.get("date") and re.match(r"^\d{4}-\d{2}-\d{2}$", r["date"]))
    if dates:
        latest = long_date(dates[-1])
        meta = data.setdefault("meta", {})
        meta["lastRecord"] = latest
        period = meta.get("reportingPeriod") or ""
        if "–" in period:
            meta["reportingPeriod"] = period.split("–")[0] + "–" + latest
        print(f"  latest published date: {latest}")

    with open(data_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")
    print(f"Updated {data_path.name}: ministryProjects {before} -> {len(records)} records")

    write_download(records)


DOWNLOAD_PATH = ROOT / "downloads" / "funding-data.xlsx"
DOWNLOAD_COLUMNS = [
    ("Ministry", "ministry"), ("Sector", "sector"), ("Project / Initiative", "project"),
    ("Status", "status"), ("Funder(s)", "funders"), ("Amount / Scale", "amountRaw"),
    ("Amount (USD millions, approx.)", "amountUSDm"), ("Partner(s)", "partners"),
    ("Region", "region"), ("Published Date", "date"), ("Funding Date", "fundingDate"),
    ("In 2025-26?", "in2025_26"), ("Description", "description"), ("Source", "sourceLink"),
    ("Confidence", "confidence"), ("Funding Type", "fundingType"), ("Source Feed", "sourceFeed"),
    ("Origin", "origin"),
]


def write_download(records):
    """Regenerate the file behind the site's "Download Excel" button so it
    always holds the same records the dashboard shows."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws.title = "Funding"
    ws.append([c for c, _ in DOWNLOAD_COLUMNS])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for r in records:
        row = []
        for _, key in DOWNLOAD_COLUMNS:
            v = r.get(key)
            row.append(", ".join(v) if isinstance(v, list) else v)
        ws.append(row)
        link = r.get("sourceLink") or ""
        if link.startswith("http"):
            cell = ws.cell(ws.max_row, [k for _, k in DOWNLOAD_COLUMNS].index("sourceLink") + 1)
            cell.hyperlink = link
            cell.font = Font(color="0563C1", underline="single")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    DOWNLOAD_PATH.parent.mkdir(exist_ok=True)
    wb.save(DOWNLOAD_PATH)
    print(f"Updated {DOWNLOAD_PATH.relative_to(ROOT)} ({len(records)} records)")


if __name__ == "__main__":
    main()
