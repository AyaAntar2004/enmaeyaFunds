"""
build_data.py

Rebuilds the `ministryProjects` array in data.json from the Lebanon
Ministry Funding 2026 workbook.

The workbook has four tabs:
  - "Ministries" — the master data sheet (437 records), 13 columns:
    Ministry, Project / Initiative, Status, Funder(s), Amount / Scale,
    Partner(s), Region, Date, Description, Source / Link, Confidence,
    Source Feed, Note. This is the source of truth for ministryProjects.
  - "Sheet2" — a smaller data sheet with the same 13 columns (a curated
    subset/preview of the same records). Read and reported on, but not
    merged into data.json since "Ministries" already is the superset.
  - "Sheet1" — manual QA notes keyed by spreadsheet row number in the
    "Ministries" tab (e.g. "wrong date 12/05/2026"). Applied as corrections
    before the records are normalized.
  - "Summary" — ministry-level rollup (records, statuses, high-confidence
    count). Exported as-is to sources/ministry-funding-summary.json.

Only the ministryProjects key in data.json is touched; hero/metrics/
scatter/heatmap/etc. are hand-curated and left alone.

Run:
    python build_data.py
"""

import json
import re
from pathlib import Path

import pandas as pd

ODS_PATH = Path(r"C:\Users\AyaAntar\Downloads\Lebanon Ministry Funding 2026 (AI Gen).ods")
ROOT = Path(__file__).resolve().parent
DATA_JSON_PATH = ROOT / "data.json"
SUMMARY_JSON_PATH = ROOT / "sources" / "ministry-funding-summary.json"

DATA_SHEET = "Ministries"
PREVIEW_SHEET = "Sheet2"
QA_SHEET = "Sheet1"
SUMMARY_SHEET = "Summary"

EXPECTED_HEADER = [
    "Ministry", "Project / Initiative", "Status", "Funder(s)",
    "Amount / Scale", "Partner(s)", "Region", "Date", "Description",
    "Source / Link", "Confidence", "Source Feed", "Note",
]
FIELD_NAMES = [
    "ministry", "project", "status", "funders_raw", "amountRaw",
    "partners_raw", "region_raw", "date", "description", "sourceLink",
    "confidence", "sourceFeed", "note",
]

# ministry -> (short code, dashboard sector). Carried over from the current
# data.json so the refreshed export slots into the same sectors used by
# index.html / lebanon-funding-map.html. Ministries not listed here fall
# back to sector "development" and get an auto-generated code; a warning
# is printed so they can be reviewed and added above.
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
# "Amount / Scale" is free text (e.g. "EUR 5 million", "USD 1,515,000
# (about USD 1,515,000)"). We take the first figure in whichever currency
# is most directly USD (USD > EUR > GBP > JPY), convert with a fixed
# approximate rate, and round to 2dp — a headline-figure, not an attempt
# to sum every number mentioned in a record.
CURRENCY_RATES_TO_USD = {"USD": 1.0, "EUR": 1.08, "GBP": 1.27, "JPY": 0.0067}
CURRENCY_PRIORITY = ["USD", "EUR", "GBP", "JPY"]
MULTIPLIERS = {"thousand": 0.001, "million": 1.0, "billion": 1000.0}
AMOUNT_RE = re.compile(
    r"(USD|US\s*\$|EUR|GBP|JPY)\s*\$?\s*((?:\d[\d,.\s]*\d)|\d)(?:\s*(thousand|million|billion))?",
    re.IGNORECASE,
)


def parse_amount_usd_millions(raw):
    if not raw or not isinstance(raw, str) or raw.strip().lower() == "not specified":
        return None

    by_currency = {}
    for m in AMOUNT_RE.finditer(raw):
        cur = m.group(1).upper().replace(" ", "")
        cur = "USD" if cur.startswith("US") else cur
        by_currency.setdefault(cur, m)

    for cur in CURRENCY_PRIORITY:
        m = by_currency.get(cur)
        if not m:
            continue
        number = float(m.group(2).replace(",", "").replace(" ", ""))
        mult_word = (m.group(3) or "").lower()
        millions = number * MULTIPLIERS.get(mult_word, 0.000001)
        usd_millions = round(millions * CURRENCY_RATES_TO_USD[cur], 2)
        return int(usd_millions) if usd_millions == int(usd_millions) else usd_millions

    return None


# --- region classification --------------------------------------------
# The sheet's "Region" column is free-text (towns, districts, comma lists).
# data.json only ever uses Lebanon's governorates + "National". We match
# known district/town keywords to a governorate; anything that matches
# zero or more-than-one distinct governorate falls back to "National"
# rather than guessing.
REGION_PATTERNS = [
    (re.compile(r"north lebanon|\btripoli\b|\bakkar\b|\bbatroun\b|\bzgharta\b|\bkoura\b|"
                r"\bbsharri\b|\bminieh\b|\bdinniyeh\b|\bdanniyeh\b|\bdouma\b|deir ammar|\bkleyate\b", re.I),
     "North Lebanon"),
    (re.compile(r"south lebanon|southern lebanon|\bsaida\b|\bsidon\b|\btyre\b|\bsour\b|"
                r"\bzahrani\b|\bjezzine\b|\bghazieh\b|\bqana\b", re.I),
     "South Lebanon"),
    (re.compile(r"baalbek|baalbeck|\bhermel\b", re.I), "Baalbek-Hermel"),
    (re.compile(r"nabat(i|ie)yeh|\bbint jbeil\b|\bmarjeyoun\b|\bhasbaya\b|\bhabboush\b", re.I), "Nabatieh"),
    (re.compile(r"\bbekaa\b|\bbeqaa\b|\bzahle\b|\brachaya\b|\brashaya\b|\brashayya\b|\banjar\b|west bekaa", re.I),
     "Bekaa"),
    (re.compile(r"\bbeirut\b|\bkarantina\b|\bverdun\b", re.I), "Beirut"),
    (re.compile(r"mount lebanon|\bchouf\b|\baley\b|\bbaabda\b|\bkeserwan\b|\bmetn\b|\bjbeil\b|"
                r"\bbyblos\b|jal el dib|\bhadath\b|sin el fil|bourj hammoud|\bzouk\b|\bsibline\b", re.I),
     "Mount Lebanon"),
]


def classify_region(raw):
    if not raw or not isinstance(raw, str):
        return "National"
    text = raw.strip()
    if not text or text.lower() in ("lebanon", "national"):
        return "National"
    matched = {gov for pattern, gov in REGION_PATTERNS if pattern.search(text)}
    return matched.pop() if len(matched) == 1 else "National"


# --- small helpers -------------------------------------------------------

def split_list(raw):
    """'FAO, WFP' -> ['FAO', 'WFP']; strips explanatory parentheticals."""
    if not raw or not isinstance(raw, str):
        return []
    text = raw.strip()
    if text.lower() in ("not specified", "none", ""):
        return []
    text = re.sub(r"\([^)]*\)", "", text)
    return [p.strip() for p in text.split(",") if p.strip()]


def normalize_date(raw):
    """Zero-pads the odd 'YYYY-M-D' cell (source has one: '2025-03-6')."""
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", str(raw).strip())
    if not m:
        return str(raw).strip() or None
    y, mo, d = m.groups()
    return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"


def slugify_code(ministry):
    words = [w for w in re.findall(r"[A-Za-z]+", ministry) if w.lower() not in ("of", "and", "the", "for")]
    return "".join(w[0].upper() for w in words) or "N/A"


# --- workbook loading ------------------------------------------------

def load_sheet_rows(xl, sheet_name):
    """Locate the header row (col 0 == 'Ministry') and return data rows below it."""
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
    return raw, header_idx, data.reset_index(drop=True)


def apply_qa_corrections(xl, raw_ministries):
    """Sheet1 flags a handful of rows (by spreadsheet row number) as having
    the wrong date, e.g. '9 Ministry of Agriculture ... wrong date 12/05/2026'.
    Row numbers are 1-indexed against the raw (untrimmed) 'Ministries' sheet."""
    qa = xl.parse(QA_SHEET, header=None)
    fixed = 0
    for _, row in qa.iterrows():
        cell0, note = str(row[0]).strip(), str(row[2]).strip()
        m_row = re.match(r"^(\d+)\s+", cell0)
        m_fix = re.match(r"wrong date\s+(\d{1,2})/(\d{1,2})/(\d{4})", note, re.IGNORECASE)
        if not (m_row and m_fix):
            continue
        row_idx = int(m_row.group(1)) - 1  # spreadsheet row N -> raw.iloc[N-1]
        dd, mm, yyyy = m_fix.groups()
        corrected = f"{int(yyyy):04d}-{int(mm):02d}-{int(dd):02d}"
        old = raw_ministries.iat[row_idx, 7]
        raw_ministries.iat[row_idx, 7] = corrected
        fixed += 1
        print(f"  QA fix: row {row_idx + 1} date {old!r} -> {corrected!r}")
    print(f"  applied {fixed} QA date correction(s) from '{QA_SHEET}'")
    return raw_ministries


def build_records(data_rows):
    records = []
    unmapped = set()
    for row in data_rows.itertuples(index=False, name=None):
        rec = dict(zip(FIELD_NAMES, row))
        ministry = str(rec["ministry"]).strip()

        code, sector = MINISTRY_MAP.get(ministry, (None, None))
        if code is None:
            code, sector = slugify_code(ministry), DEFAULT_SECTOR
            unmapped.add(ministry)

        amount_raw = str(rec["amountRaw"]).strip() or "Not specified"

        records.append({
            "ministry": ministry,
            "ministryCode": code,
            "sector": sector,
            "project": str(rec["project"]).strip(),
            "status": str(rec["status"]).strip(),
            "funders": split_list(rec["funders_raw"]),
            "amountRaw": amount_raw,
            "amountUSDm": parse_amount_usd_millions(amount_raw),
            "partners": split_list(rec["partners_raw"]),
            "region": classify_region(rec["region_raw"]),
            "date": normalize_date(rec["date"]),
            "confidence": str(rec["confidence"]).strip(),
            "sourceFeed": str(rec["sourceFeed"]).strip(),
            "description": str(rec["description"]).strip(),
            "sourceLink": str(rec["sourceLink"]).strip(),
            "note": str(rec["note"]).strip() if rec["note"] and str(rec["note"]).lower() != "nan" else "",
        })

    if unmapped:
        print(f"  warning: {len(unmapped)} ministry name(s) not in MINISTRY_MAP, "
              f"defaulted to sector '{DEFAULT_SECTOR}':")
        for m in sorted(unmapped):
            print(f"    - {m}")

    return records


def build_summary(xl):
    raw = xl.parse(SUMMARY_SHEET, header=None)
    header_idx = None
    for i in range(len(raw)):
        if str(raw.iat[i, 0]).strip() == "Ministry":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Could not find header row in Summary sheet")

    columns = [str(v).strip() for v in raw.iloc[header_idx].tolist()]
    rows = raw.iloc[header_idx + 1:]
    rows = rows[rows[0].notna() & (rows[0].astype(str).str.strip() != "")]

    summary = []
    for row in rows.itertuples(index=False, name=None):
        label = str(row[0]).strip()
        if label.upper() in ("TOTAL", "NOTES") or label.startswith("•"):
            continue
        entry = dict(zip(columns, row))
        entry = {k: (int(v) if isinstance(v, float) and v.is_integer() else v) for k, v in entry.items()}
        summary.append(entry)
    return summary


def main():
    print(f"Reading {ODS_PATH}")
    xl = pd.ExcelFile(ODS_PATH, engine="odf")
    print(f"  sheets found: {xl.sheet_names}")

    raw_ministries, header_idx, _ = load_sheet_rows(xl, DATA_SHEET)
    raw_ministries = apply_qa_corrections(xl, raw_ministries)
    data_rows = raw_ministries.iloc[header_idx + 1:].reset_index(drop=True)
    data_rows = data_rows[data_rows[0].notna() & (data_rows[0].astype(str).str.strip() != "")]
    records = build_records(data_rows)
    print(f"  '{DATA_SHEET}': {len(records)} records normalized")

    _, _, preview_rows = load_sheet_rows(xl, PREVIEW_SHEET)
    print(f"  '{PREVIEW_SHEET}': {len(preview_rows)} rows (preview sheet, not merged into data.json)")

    summary = build_summary(xl)
    print(f"  '{SUMMARY_SHEET}': {len(summary)} ministry rollup rows")

    print(f"Updating {DATA_JSON_PATH}")
    with open(DATA_JSON_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    before = len(data.get("ministryProjects", []))
    data["ministryProjects"] = records
    with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")
    print(f"  ministryProjects: {before} -> {len(records)} records")

    SUMMARY_JSON_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(SUMMARY_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
        f.write("\n")
    print(f"Wrote {SUMMARY_JSON_PATH}")

    print("\nSample record:")
    print(json.dumps(records[0], indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
