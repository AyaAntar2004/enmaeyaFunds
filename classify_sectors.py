"""
classify_sectors.py  (run once, by hand)

Re-classifies the sector of every base record in funding_updated.ods by what the
PROJECT does, instead of by which ministry runs it (build_data.py's MINISTRY_MAP
gives every project of a ministry the same sector).

Uses a small OpenAI model (gpt-4o-mini) on each record's project name,
description, ministry, funders and partners, with the same sector definitions as
the daily AI extraction in the news-scraper repo. Writes:

    sector_overrides.csv   one row per record: ministry, project, source link,
                           old (ministry-based) sector, new AI sector, short reason

sync_daily.py applies the "AI sector" column when it rebuilds data.json, so the
CSV is the single place to review or correct sectors: open it in Excel, change
any "AI sector" cell to one of Health / Education / Economy / Environment /
Artificial Intelligence / Development, save as CSV, commit. Delete a row to fall
back to the ministry-based sector for that record.

Usage (from the repository folder, with your OpenAI key):
    pip install -r requirements.txt requests
    set OPENAI_API_KEY=sk-...          (Windows Command Prompt)
    python classify_sectors.py

It is safe to stop and run again: records already in sector_overrides.csv are
skipped. Use --redo to classify everything again. Cost: well under $0.10 for the
526 records.
"""

import argparse
import csv
import json
import os
import sys
import time

import pandas as pd
import requests

import build_data as bd

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OVERRIDES_PATH = bd.ROOT / "sector_overrides.csv"
CSV_FIELDS = ["Ministry", "Project", "Source link", "Ministry-based sector", "AI sector", "Reason"]

# Same definitions as SECTORS in the scraper repo's scraper/extract_funding.py.
SECTORS = {
    "Health": "hospitals, primary and community health care, medicines and medical equipment, "
              "public health, disease control, vaccination, mental health, digital health",
    "Education": "schools, universities, access to education, crisis/emergency education, "
                 "teachers, vocational and technical training (TVET), future skills, citizenship",
    "Economy": "economic recovery, jobs and employment, livelihoods, entrepreneurship, SMEs, "
               "private sector, trade, industry, investment, finance, tourism, agricultural value "
               "chains and farmer incomes",
    "Environment": "climate, water and wastewater, waste management, biodiversity, forests and "
                   "land, clean/renewable energy, electricity sector, sustainable agriculture and "
                   "natural resources",
    "Artificial Intelligence": "AI, digital government and e-services, digital transformation of "
                               "public administration, data systems, telecoms, AI skills, "
                               "innovation and tech",
    "Development": "governance, institutions, security forces and justice, social protection and "
                   "social inclusion, humanitarian, food and cash assistance, shelter, refugees, "
                   "youth, gender, municipalities, reconstruction and general recovery",
}
# dashboard sector ids (data.json sectorFilters)
SECTOR_IDS = {"Health": "health", "Education": "education", "Economy": "economy",
              "Environment": "environment", "Artificial Intelligence": "ai",
              "Development": "development"}
ID_TO_LABEL = {v: k for k, v in SECTOR_IDS.items()}

PROMPT = ("You assign ONE dashboard sector to a funded project in Lebanon, based on what the "
          "PROJECT does, not on which ministry runs it. E.g. a job-training project with the "
          "Ministry of Social Affairs is Economy; a digital farmers' registry is Artificial "
          "Intelligence; a forest project is Environment. Sectors:\n"
          + "\n".join(f"- {k}: {v}" for k, v in SECTORS.items())
          + "\nReturn the sector and a reason of at most 15 words.")

SCHEMA = {
    "name": "sector",
    "strict": True,
    "schema": {
        "type": "object", "additionalProperties": False,
        "required": ["sector", "reason"],
        "properties": {"sector": {"type": "string", "enum": list(SECTORS)},
                       "reason": {"type": "string"}},
    },
}


def record_key(ministry, project, link):
    return (str(ministry or "").strip().lower(), str(project or "").strip().lower(),
            str(link or "").strip().lower())


def load_overrides():
    """{(ministry, project, link) lowercased: dashboard sector id}. Missing file -> {}."""
    if not OVERRIDES_PATH.exists():
        return {}
    out = {}
    with open(OVERRIDES_PATH, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            sid = SECTOR_IDS.get((row.get("AI sector") or "").strip())
            if sid:
                out[record_key(row.get("Ministry"), row.get("Project"), row.get("Source link"))] = sid
    return out


def classify(api_key, model, rec):
    user = json.dumps({k: rec.get(k) for k in
                       ("ministry", "project", "description", "funders", "partners", "amountRaw")},
                      ensure_ascii=False)
    body = {"model": model, "temperature": 0,
            "messages": [{"role": "system", "content": PROMPT}, {"role": "user", "content": user}],
            "response_format": {"type": "json_schema", "json_schema": SCHEMA}}
    for attempt in range(5):
        try:
            r = requests.post("https://api.openai.com/v1/chat/completions", json=body, timeout=60,
                              headers={"Authorization": f"Bearer {api_key}"})
            if r.status_code == 200:
                return json.loads(r.json()["choices"][0]["message"]["content"])
            if r.status_code in (400, 401, 403, 404):
                raise SystemExit(f"OpenAI error HTTP {r.status_code}: {r.text[:300]}")
            err = f"HTTP {r.status_code}"
        except requests.RequestException as exc:
            err = str(exc)
        time.sleep(5 * (attempt + 1))
        print(f"  retrying ({err})")
    raise SystemExit("OpenAI kept failing - run again later; finished rows are kept.")


def main():
    p = argparse.ArgumentParser(description="Classify base records' sectors by project with AI.")
    p.add_argument("--model", default=os.environ.get("OPENAI_MODEL") or "gpt-4o-mini")
    p.add_argument("--redo", action="store_true", help="Classify every record again.")
    args = p.parse_args()
    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise SystemExit("Set OPENAI_API_KEY first (Windows: set OPENAI_API_KEY=sk-...).")

    xl = pd.ExcelFile(bd.ODS_PATH, engine="odf")
    header_idx, data_rows = bd.load_sheet_rows(xl, bd.SHEET_NAME)
    links = bd.extract_column_hrefs(bd.ODS_PATH, bd.SOURCE_LINK_COL_IDX, header_idx + 1)
    records = bd.build_records(data_rows, links)

    done = set() if args.redo else set(load_overrides())
    mode = "w" if args.redo or not OVERRIDES_PATH.exists() else "a"
    changed = 0
    with open(OVERRIDES_PATH, mode, encoding="utf-8-sig" if mode == "w" else "utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        if mode == "w":
            w.writeheader()
        todo = [r for r in records if record_key(r["ministry"], r["project"], r["sourceLink"]) not in done]
        print(f"{len(records)} base records, {len(todo)} to classify")
        for i, r in enumerate(todo, 1):
            res = classify(api_key, args.model, r)
            old = ID_TO_LABEL.get(r["sector"], r["sector"])
            if res["sector"] != old:
                changed += 1
            w.writerow({"Ministry": r["ministry"], "Project": r["project"], "Source link": r["sourceLink"],
                        "Ministry-based sector": old, "AI sector": res["sector"], "Reason": res["reason"]})
            f.flush()
            print(f"  [{i}/{len(todo)}] {old:24s} -> {res['sector']:24s} {r['project'][:60]}")
    print(f"\nDone. {changed} of {len(todo)} record(s) changed sector. Review {OVERRIDES_PATH.name}, "
          f"then run: python sync_daily.py")


if __name__ == "__main__":
    main()
