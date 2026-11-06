"""
data_fixes.py

Clean-ups that sync_daily.py applies to every record after it rebuilds the data:

1. Funder / partner names (NAME_FIXES): Arabic entries are replaced by their
   English names, and a few spelling variants / list-splitting leftovers are
   merged, so each funder appears once in the dashboard's "Funder" filter.
   Any Arabic name not in the table is dropped from the list and reported, so
   no Arabic entry reaches the filter. To add or change a mapping, edit the
   table below.

2. corrections.csv: field-level corrections for single records, one row per
   change. Rows come from research_missing.py (AI web research, Status
   "pending") or are typed in by hand. Only rows whose Status is "approved" are
   applied, so a reviewer decides what reaches the dashboard.

   Columns: Status, Source link, Project, Field, Current value, New value,
            Evidence URL, Evidence quote, Method, Notes
   Field is one of: date (published date), fundingDate, funders, partners,
            amountRaw, projectPeriod, region, status
   A row matches the record with the same Source link (and the same Project,
   when Project is filled in - needed where one article covers two projects).
"""

import csv
import re
from pathlib import Path

import build_data as bd

ROOT = Path(__file__).resolve().parent
CORRECTIONS_PATH = ROOT / "corrections.csv"
CORRECTION_FIELDS = ["Status", "Source link", "Project", "Field", "Current value", "New value",
                     "Evidence URL", "Evidence quote", "Method", "Notes"]
EDITABLE_FIELDS = {"date", "fundingDate", "funders", "partners", "amountRaw", "projectPeriod",
                   "region", "status"}

ARABIC = re.compile(r"[؀-ۿ]")

# name as it appears in the data -> English name(s) to use instead ("" drops it)
NAME_FIXES = {
    # Arabic entries (all from the base sheet)
    "البنك الدولي": "World Bank",
    "الاتحاد الأوروبي": "European Union",
    "البنك الإسلامي": "Islamic Development Bank",
    "البنك الإسلامي للتنمية": "Islamic Development Bank",
    "وزارة الصحة العامة": "Ministry of Public Health",
    "بنك الاستثمار الأوروبي": "European Investment Bank",
    "الرئيس نجيب ميقاتي": "Prime Minister Najib Mikati",
    "جمعية العزم والسعادة الاجتماعية": "Al-Azm wa Al-Saadeh Social Association",
    "مجلس الوزراء": "Council of Ministers",
    "الحكومة الإيطالية": "Government of Italy",
    "شركة أسترازينيكا": "AstraZeneca",
    "المؤسسة العالمية للسكري": "Global Diabetes Foundation",
    "الوكالة الفرنسية للتنمية": "Agence Française de Développement",
    "الجمهورية الإسلامية الإيرانية": "Islamic Republic of Iran",
    "منظمة الصحة العالمية": "World Health Organization",
    "مصر": "Egypt",
    "الأردن": "Jordan",
    "قطر": "Qatar",
    "فرنسا": "France",
    "اليونان": "Greece",
    "الدولة اللبنانية/وزارة الصحة العامة": ["Lebanese State", "Ministry of Public Health"],
    "جمعية موطن صحي": "Healthy Home Association",
    "دائرة طب العائلة في الجامعة الأميركية في بيروت":
        "American University of Beirut - Department of Family Medicine",
    # spelling variants and leftovers from splitting names on commas
    "Agence Francaise de Developpement": "Agence Française de Développement",
    "French Development Agency": "Agence Française de Développement",
    "Gavi the Vaccine Alliance": "Gavi",
    "the Vaccine Alliance": "Gavi",
    "The Pandemic Fund": "Pandemic Fund",
    "Lebanese government": "Lebanese Government",
    "represented by NCA": "NCA",
}


def _fix_names(names):
    out, dropped = [], []
    for n in names or []:
        n = str(n).strip()
        fixed = NAME_FIXES.get(n, n)
        for v in (fixed if isinstance(fixed, list) else [fixed]):
            if not v:
                continue
            if ARABIC.search(v):
                dropped.append(v)
                continue
            if v not in out:
                out.append(v)
    return out, dropped


def fix_names(records):
    """Apply NAME_FIXES to funders and partners of every record (in place)."""
    changed, dropped_all = 0, []
    for r in records:
        for key in ("funders", "partners"):
            new, dropped = _fix_names(r.get(key))
            if new != r.get(key):
                changed += 1
                r[key] = new
            dropped_all += dropped
    if changed:
        print(f"  names: fixed funder/partner lists on {changed} record field(s)")
    if dropped_all:
        print(f"  names: dropped {len(dropped_all)} Arabic name(s) with no English mapping "
              f"(add them to NAME_FIXES in data_fixes.py): {sorted(set(dropped_all))}")


def _norm(text):
    return re.sub(r"\s+", " ", str(text or "")).strip().lower().rstrip("/")


def load_corrections(status="approved"):
    if not CORRECTIONS_PATH.exists():
        return []
    with open(CORRECTIONS_PATH, encoding="utf-8-sig", newline="") as f:
        return [row for row in csv.DictReader(f)
                if (row.get("Status") or "").strip().lower() == status]


def _apply_value(rec, field, value):
    value = str(value or "").strip()
    if field in ("funders", "partners"):
        rec[field] = bd.split_list(value)
    elif field == "amountRaw":
        rec["amountRaw"] = value or "Not specified"
        rec["amountUSDm"] = bd.parse_amount_usd_millions(rec["amountRaw"])
    elif field in ("date", "fundingDate"):
        rec[field] = bd.normalize_date(value) if value else None
    elif field == "region":
        rec["region"] = bd.classify_region(value)
    else:
        rec[field] = value


def apply_corrections(records):
    """Apply the approved rows of corrections.csv to matching records (in place)."""
    rows = load_corrections("approved")
    if not rows:
        return
    by_link = {}
    for r in records:
        by_link.setdefault(_norm(r.get("sourceLink")), []).append(r)
    applied, unmatched, bad_field = 0, 0, 0
    for row in rows:
        field = (row.get("Field") or "").strip()
        if field not in EDITABLE_FIELDS:
            bad_field += 1
            continue
        targets = by_link.get(_norm(row.get("Source link")), [])
        project = _norm(row.get("Project"))
        if project:
            targets = [r for r in targets if _norm(r.get("project")) == project] or targets
        if not targets:
            unmatched += 1
            continue
        for r in targets:
            _apply_value(r, field, row.get("New value"))
            if field == "fundingDate" and row.get("Evidence quote"):
                r["dateBasis"] = r.get("dateBasis") or "Researched"
                r["dateEvidence"] = row["Evidence quote"]
        applied += 1
    print(f"  corrections.csv: {applied} approved correction(s) applied"
          + (f", {unmatched} matched no record" if unmatched else "")
          + (f", {bad_field} with an unknown Field" if bad_field else ""))


def tidy(records):
    """Small consistency fixes: placeholder text in date fields becomes empty."""
    for r in records:
        for key in ("date", "fundingDate"):
            if str(r.get(key) or "").strip().lower() in ("not specified", "unknown", "n/a"):
                r[key] = None
