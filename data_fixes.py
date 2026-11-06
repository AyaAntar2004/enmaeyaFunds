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
            amountRaw, projectPeriod, region, status, description, fundingType,
            exclude (New value "yes": remove the record - not funding, or a duplicate)
   An empty New value clears the cell (e.g. a funding date the source does not support).
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
                   "region", "status", "description", "fundingType", "exclude"}

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
    elif field == "exclude":
        rec["_exclude"] = value.lower() in ("yes", "y", "true", "1", "x")
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
            if field == "fundingDate":
                if not str(row.get("New value") or "").strip():
                    r["dateBasis"], r["dateEvidence"] = "Not stated in article", ""
                elif row.get("Evidence quote"):
                    r["dateBasis"] = "Researched"
                    r["dateEvidence"] = row["Evidence quote"]
        applied += 1
    print(f"  corrections.csv: {applied} approved correction(s) applied"
          + (f", {unmatched} matched no record" if unmatched else "")
          + (f", {bad_field} with an unknown Field" if bad_field else ""))


def drop_excluded(records):
    """Remove records marked exclude=yes in corrections.csv. Returns the kept list."""
    kept = [r for r in records if not r.pop("_exclude", False)]
    if len(kept) != len(records):
        print(f"  corrections.csv: {len(records) - len(kept)} record(s) removed (exclude)")
    return kept


# How a funding date was worked out (the record's dateBasis). Validation (Nov 2026)
# found that these give the date of an event, a launch or the article - not of the
# funding - so such funding dates are cleared ("left unspecified").
NOT_FUNDING_DATE_BASES = ("same-day report", "in-text: launched", "from live page: event date",
                          "from live page: same-day report", "from live page: published",
                          "from live page: launched")


def tidy(records):
    """Consistency fixes: placeholder text in date fields becomes empty, and funding
    dates that are really event / launch / publication dates are cleared."""
    cleared = 0
    for r in records:
        for key in ("date", "fundingDate"):
            if str(r.get(key) or "").strip().lower() in ("not specified", "unknown", "n/a"):
                r[key] = None
        basis = str(r.get("dateBasis") or "").strip().lower()
        if r.get("fundingDate") and basis != "researched" and (
                basis.startswith(NOT_FUNDING_DATE_BASES)
                or (basis == "in-text date" and r["fundingDate"] == r.get("date"))):
            r["fundingDate"], r["dateBasis"], r["dateEvidence"] = None, "Not stated in article", ""
            cleared += 1
    if cleared:
        print(f"  dates: cleared {cleared} funding date(s) that were event/launch/publication dates")


DUPLICATES_PATH = ROOT / "possible_duplicates.csv"
_DUP_STOP = set("with from that this their have been will under project programme program lebanon "
                "lebanese support ministry funding initiative phase through".split())


def report_duplicates(records):
    """Write possible_duplicates.csv: pairs of records of the same ministry that look like
    the same project (similar names, or fairly similar names published within a month of
    each other with the same funders). Nothing is removed automatically - a reviewer marks
    the extra one with Field "exclude" in corrections.csv."""
    import itertools
    from datetime import date

    def words(r):
        return {w for w in re.findall(r"[a-z]{4,}", str(r.get("project") or "").lower())
                if w not in _DUP_STOP}

    def years(r):
        return set(re.findall(r"\b20\d\d\b", str(r.get("project") or "")))

    def day(r):
        try:
            return date.fromisoformat(str(r.get("date"))[:10])
        except ValueError:
            return None

    rows = []
    for i, j in itertools.combinations(range(len(records)), 2):
        a, b = records[i], records[j]
        if a.get("ministry") != b.get("ministry"):
            continue
        ya, yb = years(a), years(b)
        if ya and yb and ya != yb:  # e.g. Response Plan 2025 vs 2026
            continue
        wa, wb = words(a), words(b)
        if not wa or not wb:
            continue
        sim = len(wa & wb) / len(wa | wb)
        da, db = day(a), day(b)
        close = bool(da and db and abs((da - db).days) <= 31)
        same_funders = set(a.get("funders") or []) == set(b.get("funders") or [])
        if sim >= 0.5 or (sim >= 0.3 and close and same_funders):
            rows.append({
                "Similarity": f"{sim:.2f}", "Ministry": a.get("ministry", ""),
                "Project A": a.get("project", ""), "Date A": a.get("date") or "",
                "Link A": a.get("sourceLink", ""),
                "Project B": b.get("project", ""), "Date B": b.get("date") or "",
                "Link B": b.get("sourceLink", ""),
                "Same funders": "yes" if same_funders else "no",
            })
    rows.sort(key=lambda r: -float(r["Similarity"]))
    with open(DUPLICATES_PATH, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Similarity", "Ministry", "Project A", "Date A", "Link A",
                                          "Project B", "Date B", "Link B", "Same funders"])
        w.writeheader()
        w.writerows(rows)
    print(f"  duplicates: {len(rows)} possible duplicate pair(s) listed in {DUPLICATES_PATH.name}")
