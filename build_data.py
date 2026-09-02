"""
build_data.py

Rebuilds the `ministryProjects` array in data.json from the Lebanon
Ministry Funding workbook (new format: one sheet, 22 columns, 526 records).

Column layout (in order):
  Ministry, Project / Initiative, Status, Funder(s), Amount / Scale,
  Partner(s), Region, Funding Date, Date Basis, Date Evidence (from
  article), Project Period, In 2025-26?, Published Date, Description,
  Source / Link, Enmaeya Reporting Link, Confidence, Source Feed,
  Origin, Source Checked, Live-page Evidence, Funding Type.

The site (index.html / lebanon-funding-map.html) only reads these keys
off each ministryProjects record: ministry, ministryCode, sector,
project, status, funders, amountRaw, amountUSDm, partners, region.
Everything else (date, description, sourceLink, confidence, sourceFeed,
and the new extra columns) is carried through for reference/future use
but isn't required for the charts to render.

"date" is populated from the "Published Date" column (per project
decision), not "Funding Date" — many rows don't have a confirmed
funding date, but every row has a published date.

The "Source / Link" column displays as a placeholder label ("Open
source \u25b8") with the real URL stored as a hyperlink on the cell, not
as visible text. This script reads content.xml directly to pull out the
actual href for each row instead of using the display text.

Run:
    python build_data.py
"""

import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd

ODS_PATH = Path(__file__).resolve().parent / "funding_updated.ods"
ROOT = Path(__file__).resolve().parent
DATA_JSON_PATH = ROOT / "data.json"

SHEET_NAME = "Sheet1"

EXPECTED_HEADER = [
    "Ministry", "Project / Initiative", "Status", "Funder(s)",
    "Amount / Scale", "Partner(s)", "Region", "Funding Date", "Date Basis",
    "Date Evidence (from article)", "Project Period", "In 2025-26?",
    "Published Date", "Description", "Source / Link",
    "Enmaeya Reporting Link", "Confidence", "Source Feed", "Origin",
    "Source Checked", "Live-page Evidence", "Funding Type",
]
FIELD_NAMES = [
    "ministry", "project", "status", "funders_raw", "amountRaw",
    "partners_raw", "region_raw", "funding_date_raw", "date_basis",
    "date_evidence", "project_period", "in_2025_26", "published_date_raw",
    "description", "sourceLink_display", "enmaeya_link", "confidence",
    "sourceFeed", "origin", "source_checked", "live_page_evidence",
    "funding_type",
]
SOURCE_LINK_COL_IDX = EXPECTED_HEADER.index("Source / Link")

# ministry -> (short code, dashboard sector). Ministries not listed here
# fall back to sector "development" and get an auto-generated code; a
# warning is printed so they can be reviewed and added above.
MINISTRY_MAP = {
    "Ministry of Agriculture": ("MoA", "environment"),
    "Ministry of Culture": ("MoC", "development"),
    "Ministry of Economy and Trade": ("MoET", "economy"),
    "Ministry of Education and Higher Education": ("MoEHE", "education"),
    "Ministry of Energy and Water": ("MoEW", "environment"),
    "Ministry of Environment": ("MoEnv", "environment"),
    "Ministry of Finance": ("MoF", "economy"),
    "Ministry of Foreign Affairs and Emigrants": ("MoFA", "development"),
    "Ministry of Industry": ("MoInd", "economy"),
    "Ministry of Information": ("MoInfo", "development"),
    "Ministry of Interior and Municipalities": ("MoIM", "development"),
    "Ministry of Justice": ("MoJ", "development"),
    "Ministry of Labour": ("MoL", "economy"),
    "Ministry of Public Health": ("MoPH", "health"),
    "Ministry of Public Works and Transport": ("MoPWT", "development"),
    "Ministry of Social Affairs": ("MoSA", "development"),
    "Ministry of State for Administrative Development": ("MoSAD", "ai"),
    "Ministry of State for Administrative Reform (OMSAR)": ("OMSAR", "ai"),
    "Ministry of Telecommunications": ("MoTel", "ai"),
    "Ministry of Tourism": ("MoTour", "economy"),
    "Ministry of National Defense": ("MoND", "development"),
    "Not specified": ("N/A", "development"),
}
DEFAULT_SECTOR = "development"

# --- amount parsing --------------------------------------------------
CURRENCY_RATES_TO_USD = {"USD": 1.0, "EUR": 1.08, "GBP": 1.27, "JPY": 0.0067}
CURRENCY_PRIORITY = ["USD", "EUR", "GBP", "JPY"]
MULTIPLIERS = {"thousand": 0.001, "million": 1.0, "billion": 1000.0}
AMOUNT_RE = re.compile(
    r"(USD|US\s*\$|EUR|GBP|JPY)\s*\$?\s*((?:\d[\d,.\s]*\d)|\d)(?:\s*(thousand|million|billion))?",
    re.IGNORECASE,
)
# Fallback for "20 million euros" / "600,000 US dollars" style phrasing,
# where the currency word trails the number instead of leading it.
CURRENCY_WORDS = {
    "dollars": "USD", "us dollars": "USD", "usd": "USD",
    "euros": "EUR", "eur": "EUR",
    "pounds sterling": "GBP", "gbp": "GBP",
    "yen": "JPY", "jpy": "JPY",
}
TRAILING_AMOUNT_RE = re.compile(
    r"((?:\d[\d,.\s]*\d)|\d)\s*(thousand|million|billion)?\s*"
    r"(dollars|us dollars|usd|euros|eur|pounds sterling|gbp|yen|jpy)",
    re.IGNORECASE,
)


def parse_amount_usd_millions(raw):
    if not raw or not isinstance(raw, str) or raw.strip().lower() == "not specified":
        return None

    by_currency = {}
    for m in AMOUNT_RE.finditer(raw):
        cur = m.group(1).upper().replace(" ", "")
        cur = "USD" if cur.startswith("US") else cur
        by_currency.setdefault(cur, ("code_first", m))

    for m in TRAILING_AMOUNT_RE.finditer(raw):
        cur = CURRENCY_WORDS.get(m.group(3).lower())
        if cur:
            by_currency.setdefault(cur, ("word_after", m))

    for cur in CURRENCY_PRIORITY:
        entry = by_currency.get(cur)
        if not entry:
            continue
        kind, m = entry
        if kind == "code_first":
            number = float(m.group(2).replace(",", "").replace(" ", ""))
            mult_word = (m.group(3) or "").lower()
        else:  # word_after: group(1)=number, group(2)=multiplier
            number = float(m.group(1).replace(",", "").replace(" ", ""))
            mult_word = (m.group(2) or "").lower()
        millions = number * MULTIPLIERS.get(mult_word, 0.000001)
        usd_millions = round(millions * CURRENCY_RATES_TO_USD[cur], 2)
        return int(usd_millions) if usd_millions == int(usd_millions) else usd_millions

    return None


# --- region classification --------------------------------------------
REGION_PATTERNS = [
    (re.compile(r"\bakkar\b|\bhalba\b|\bakroum\b|\bmhammara\b|\bkleyate\b", re.I),
     "Akkar"),
    (re.compile(r"north lebanon|\btripoli\b|\bbatroun\b|\bzgharta\b|\bkoura\b|"
                r"\bbsharri\b|\bminieh\b|\bdinniyeh\b|\bdanniyeh\b|\bdouma\b|deir ammar", re.I),
     "North Lebanon"),
    (re.compile(r"south lebanon|southern lebanon|\bsaida\b|\bsidon\b|\btyre\b|\bsour\b|"
                r"\bzahrani\b|\bjezzine\b|\bghazieh\b|\bqana\b", re.I),
     "South Lebanon"),
    (re.compile(r"baalbek|baalbeck|\bhermel\b", re.I), "Baalbek-Hermel"),
    (re.compile(r"\bnabatiyeh\b|\bnabatieh\b|\bbint jbeil\b|\bmarjeyoun\b|\bmarjayoun\b|\bhasbaya\b|\bhabboush\b", re.I), "Nabatieh"),
    (re.compile(r"\bbekaa\b|\bbeqaa\b|\bzahle\b|\brachaya\b|\brashaya\b|\brashayya\b|\banjar\b|west bekaa", re.I),
     "Bekaa"),
    (re.compile(r"\bbeirut\b|\bkarantina\b|\bverdun\b|burj al-barajneh|bourj hammoud", re.I), "Beirut"),
    (re.compile(r"mount lebanon|\bchouf\b|\baley\b|\bbaabda\b|\bkeserwan\b|\bmetn\b|\bjbeil\b|"
                r"\bbyblos\b|jal el dib|\bhadath\b|sin el fil|\bzouk\b|\bsibline\b|\bjounieh\b", re.I),
     "Mount Lebanon"),
    (re.compile(r"\bsouth governorate\b", re.I), "South Lebanon"),
]


def classify_region(raw):
    if not raw or not isinstance(raw, str):
        return "National"
    text = raw.strip()
    if not text or text.lower() in ("lebanon", "national", "not specified"):
        return "National"
    matched = {gov for pattern, gov in REGION_PATTERNS if pattern.search(text)}
    return matched.pop() if len(matched) == 1 else "National"


# --- small helpers -------------------------------------------------------

def clean(v):
    """Normalize NaN/NaT/None/lone-NBSP cells to None."""
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        if pd.isna(v):
            return None
        return v.strftime("%Y-%m-%d")
    text = str(v).replace("\xa0", " ").strip()
    if not text or text.lower() == "nan" or text.lower() == "nat":
        return None
    return text


def split_list(raw):
    """'FAO, WFP' -> ['FAO', 'WFP']; strips explanatory parentheticals."""
    text = clean(raw)
    if not text or text.lower() in ("not specified", "none"):
        return []
    text = re.sub(r"\([^)]*\)", "", text)
    return [p.strip() for p in text.split(",") if p.strip()]


def normalize_date(raw):
    """Accepts a pandas Timestamp (from the Published Date column) or a
    plain string, returns YYYY-MM-DD or None."""
    if raw is None:
        return None
    if isinstance(raw, pd.Timestamp):
        if pd.isna(raw):
            return None
        return raw.strftime("%Y-%m-%d")
    text = clean(raw)
    if not text:
        return None
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", text)
    if m:
        y, mo, d = m.groups()
        return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"
    return text


def slugify_code(ministry):
    words = [w for w in re.findall(r"[A-Za-z]+", ministry) if w.lower() not in ("of", "and", "the", "for")]
    return "".join(w[0].upper() for w in words) or "N/A"


# --- hyperlink extraction (the "Source / Link" column shows a display
# label, not the real URL - the URL is a hyperlink on the cell) --------

TABLE_NS = {
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}
XLINK_HREF = "{http://www.w3.org/1999/xlink}href"


def extract_column_hrefs(ods_path, col_idx, num_header_rows):
    """Returns a list of hrefs (or display text if no link), one per data
    row, aligned with pandas' row order for the given column index."""
    with zipfile.ZipFile(ods_path) as z:
        content = z.read("content.xml")
    root = ET.fromstring(content)
    table = root.find(".//table:table", TABLE_NS)
    rows = table.findall("table:table-row", TABLE_NS)
    data_rows = rows[num_header_rows:]

    results = []
    for row in data_rows:
        cells = row.findall("table:table-cell", TABLE_NS)
        if len(cells) <= col_idx:
            results.append(None)
            continue
        cell = cells[col_idx]
        href, texts = None, []
        for p in cell.findall("text:p", TABLE_NS):
            texts.append("".join(p.itertext()))
            a = p.find("text:a", TABLE_NS)
            if a is not None:
                href = a.get(XLINK_HREF)
        results.append(href or (" ".join(texts).strip() or None))
    return results


# --- workbook loading ------------------------------------------------

def load_sheet_rows(xl, sheet_name):
    """Locate the header row (col 0 == 'Ministry') and return the header
    row index plus data rows below it."""
    raw = xl.parse(sheet_name, header=None)
    header_idx = None
    for i in range(len(raw)):
        if str(raw.iat[i, 0]).strip() == "Ministry":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(f"Could not find a 'Ministry' header row in sheet '{sheet_name}'")

    header_vals = [str(v).strip() for v in raw.iloc[header_idx].tolist()]
    if header_vals != EXPECTED_HEADER:
        print(f"  warning: '{sheet_name}' header differs from expected:\n    got: {header_vals}")

    data = raw.iloc[header_idx + 1:].reset_index(drop=True)
    data = data[data[0].notna() & (data[0].astype(str).str.strip() != "")]
    return header_idx, data.reset_index(drop=True)


def build_records(data_rows, real_links):
    records = []
    unmapped = set()
    dupes_seen = {}

    for i, row in enumerate(data_rows.itertuples(index=False, name=None)):
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
            "date": normalize_date(rec["published_date_raw"]),
            "confidence": clean(rec["confidence"]) or "Unknown",
            "sourceFeed": clean(rec["sourceFeed"]) or "",
            "description": clean(rec["description"]) or "",
            "sourceLink": real_links[i] if i < len(real_links) else clean(rec["sourceLink_display"]) or "",
            "note": "",
            # extra columns from the new sheet, kept for reference even
            # though the current charts don't read them yet
            "fundingDate": normalize_date(rec["funding_date_raw"]),
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
            print(f"    - [{n}x] {m} — {p}")

    return records


def main():
    print(f"Reading {ODS_PATH}")
    xl = pd.ExcelFile(ODS_PATH, engine="odf")
    print(f"  sheets found: {xl.sheet_names}")

    header_idx, data_rows = load_sheet_rows(xl, SHEET_NAME)
    print(f"  '{SHEET_NAME}': {len(data_rows)} records found")

    real_links = extract_column_hrefs(ODS_PATH, SOURCE_LINK_COL_IDX, header_idx + 1)

    records = build_records(data_rows, real_links)
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

    print("\nSample record:")
    print(json.dumps(records[0], indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()