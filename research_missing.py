"""
research_missing.py

Second AI pass for missing cells: for each record with gaps, an OpenAI model
searches the web (starting from the record's own source link, then other
coverage of the same project, e.g. the National News Agency, donor and UN
pages) and proposes values for the missing fields only. Every proposal carries
the page it came from and a short exact quote, so it can be checked.

Proposals are added to corrections.csv with Status "pending". Nothing reaches
the dashboard until a reviewer changes Status to "approved" (then sync_daily.py
applies it). Reject a row by setting Status to "rejected".

Fields researched (only when empty): date (publication date of the source
article), fundingDate, funders, amountRaw, partners, projectPeriod.

Usage (from the dashboard repository folder):
    pip install -r requirements.txt
    set OPENAI_API_KEY=sk-...                 (Windows Command Prompt)
    python research_missing.py --limit 10     # try 10 records first and check the results
    python research_missing.py                # all records with gaps (resumable)
    python research_missing.py --fields date  # only one kind of gap

Records already researched are remembered in research_log.csv and skipped next
time (use --redo to research them again).

Cost: about $0.01-0.03 per record (web-search calls are billed per call), so
roughly $5-15 for every record with gaps. --limit keeps a run small.
"""

import argparse
import csv
import json
import os
import re
import sys
import time
from datetime import datetime, timezone

import requests

import build_data as bd
from data_fixes import CORRECTIONS_PATH, CORRECTION_FIELDS

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

LOG_PATH = bd.ROOT / "research_log.csv"
LOG_FIELDS = ["Source link", "Project", "Researched at", "Fields asked", "Fields found",
              "Search calls", "Model"]

FIELD_HELP = {
    "date": "publication date of the source article (YYYY-MM-DD)",
    "fundingDate": "date the FUNDING was signed / approved / disbursed / allocated / pledged - not a "
                   "launch, inauguration, event or publication date (YYYY-MM-DD, or YYYY-MM / YYYY)",
    "funders": "donor(s) providing the money, English names, comma-separated",
    "amountRaw": "amount with currency as written, e.g. 'USD 2.5 million' or 'EUR 10 million'",
    "partners": "implementing / co-implementing organisations, English names, comma-separated",
    "projectPeriod": "start -> end, e.g. '2023-03 -> 2026-09' or '2024 -> 2027'",
}
ALL_FIELDS = list(FIELD_HELP)

PROMPT = """You complete missing information about ONE funded project in Lebanon for a public
funding dashboard. Research ONLY the fields listed as missing.

How to research:
1. Open the record's source link first.
2. Then search the web for the same project (project name, ministry, donor, Arabic name if
   useful): official ministry or donor pages, UN agency pages, the National News Agency
   (nna-leb.gov.lb), reliefweb.int, reputable news.
3. Only report a value if a page you actually opened states it explicitly for THIS project.
   No guessing, no estimates, no values from a different project or a different year.

Field rules:
- amountRaw: money committed to THIS project (grant, loan, budget, contribution) with currency.
  NOT damages, losses, needs, appeal targets, numbers of people/workdays/tons.
- fundingDate and date: ONE date (YYYY-MM-DD, or YYYY-MM / YYYY if that is all that is stated).
  Not a range, not "ongoing".
- fundingDate is ONLY the date of the funding act itself (agreement signed, grant/loan approved,
  money disbursed or allocated, amount pledged). A launch, inauguration, ceremony, workshop,
  training, event, project start or article date is NOT a funding date - leave it out.
- date is the publication date printed on the source article (dateline / "Published on"), not
  an event date mentioned in it.
- projectPeriod: start -> end, both stated for THIS project - never assumed.
- funders: organisations or governments that PAY for the project - not implementers, not the
  ministry receiving it (unless it funds from its own budget).
- An analysis, report launch or assessment is not funding: only report a funder or amount if
  the page says who financed it.

For each value give the URL it came from and a short exact quote (original language) that
shows that value - the quote must contain the date, amount or name you report. Write values in
English. If you cannot find a field, LEAVE IT OUT - never return "Not specified" or similar.
Confidence: High only if the quote states the value for this exact project; Medium if the page
is clearly about this project but the value needs light interpretation; otherwise Low.
"""

SCHEMA = {
    "name": "research",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "required": ["found"],
        "properties": {
            "found": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "required": ["field", "value", "evidence_url", "evidence_quote", "confidence"],
                    "properties": {
                        "field": {"type": "string", "enum": ALL_FIELDS},
                        "value": {"type": "string"},
                        "evidence_url": {"type": "string"},
                        "evidence_quote": {"type": "string"},
                        "confidence": {"type": "string", "enum": ["High", "Medium", "Low"]},
                    },
                },
            }
        },
    },
}


PLACEHOLDER = re.compile(r"^(not specified|not stated|unknown|n/?a|none|ongoing|not disclosed|"
                         r"unspecified|-|\?)$", re.I)
CURRENCY = re.compile(r"(USD|US\$|\$|EUR|€|euro|GBP|£|JPY|¥|CHF|SEK|NOK|DKK|CAD|AUD|LBP|dollar|"
                      r"دولار|يورو|ليرة)", re.I)
NOT_FUNDING = re.compile(r"damage|loss|needs?\b|appeal|people|individuals|families|workdays|tons?\b",
                         re.I)
ARABIC = re.compile(r"[\u0600-\u06FF]")


def check(field, value, quote):
    """Reject proposals that are placeholders or that the quote does not back up.
    Returns a reason string, or '' if the proposal looks usable."""
    v, q = value.strip(), quote.strip()
    if not v or PLACEHOLDER.match(v) or "not specified" in v.lower():
        return "placeholder value"
    if not q or q.lower() == v.lower() or PLACEHOLDER.match(q):
        return "no real quote"
    if field in ("date", "fundingDate"):
        if not re.match(r"^\d{4}(-\d{2}(-\d{2})?)?$", v):
            return "not a single date"
        if v[:4] not in q and not ARABIC.search(q) and not re.search(r"\d", q):
            return "date not in quote"
    if field == "projectPeriod":
        parts = [x.strip() for x in v.split("->")]
        if len(parts) != 2 or not all(re.search(r"\d{4}", x) for x in parts):
            return "period needs a start and an end year"
    if field == "amountRaw":
        if not re.search(r"\d", v) or not CURRENCY.search(v):
            return "amount without number/currency"
        if NOT_FUNDING.search(v):
            return "not a funding amount"
        digits = re.findall(r"\d+", v)
        if digits and not any(d in re.sub(r"[,.\s]", "", q) or d in q for d in digits):
            return "amount not in quote"
    if field in ("funders", "partners") and not ARABIC.search(q):
        names = [n.strip() for n in v.split(",") if n.strip()]
        words = lambda n: [w for w in re.findall(r"[A-Za-z]{4,}", n)]
        if names and not any(any(w.lower() in q.lower() for w in words(n)) or n.lower() in q.lower()
                             for n in names):
            return "name not in quote"
    return ""


def is_empty(value):
    if value is None:
        return True
    if isinstance(value, list):
        return not [v for v in value if str(v).strip().lower() not in ("", "not specified")]
    return str(value).strip().lower() in ("", "not specified", "unknown", "n/a")


def missing_fields(rec, wanted):
    return [f for f in wanted if is_empty(rec.get(f))]


def show(value):
    return ", ".join(value) if isinstance(value, list) else ("" if value is None else str(value))


def ask(api_key, model, rec, fields, usage):
    known = {k: show(rec.get(k)) for k in ("ministry", "project", "description", "status", "funders",
                                           "partners", "amountRaw", "region", "date", "fundingDate",
                                           "projectPeriod", "sourceFeed")}
    user = ("RECORD (what we already know):\n" + json.dumps(known, ensure_ascii=False, indent=1)
            + f"\nSource link: {rec.get('sourceLink')}\n\nMISSING FIELDS to research:\n"
            + "\n".join(f"- {f}: {FIELD_HELP[f]}" for f in fields))
    body = {
        "model": model,
        "instructions": PROMPT,
        "input": user,
        "tools": [{"type": "web_search"}],
        "text": {"format": {"type": "json_schema", **SCHEMA}},
    }
    for attempt in range(5):
        try:
            r = requests.post("https://api.openai.com/v1/responses", json=body, timeout=300,
                              headers={"Authorization": f"Bearer {api_key}"})
        except requests.RequestException as exc:
            err = str(exc)
        else:
            if r.status_code == 200:
                data = r.json()
                calls = sum(1 for o in data.get("output", []) if o.get("type") == "web_search_call")
                u = data.get("usage") or {}
                usage["in"] += u.get("input_tokens", 0)
                usage["out"] += u.get("output_tokens", 0)
                usage["calls"] += calls
                text = ""
                for o in data.get("output", []):
                    if o.get("type") == "message":
                        for c in o.get("content", []):
                            if c.get("type") == "output_text":
                                text += c.get("text", "")
                return json.loads(text or '{"found": []}'), calls
            err = f"HTTP {r.status_code}: {r.text[:300]}"
            if r.status_code in (400, 401, 403, 404):
                raise SystemExit(f"OpenAI error - {err}")
        wait = 10 * (attempt + 1)
        print(f"    retrying in {wait}s ({err})")
        time.sleep(wait)
    raise SystemExit("OpenAI kept failing - run again later; finished records are kept.")


def load_done():
    if not LOG_PATH.exists():
        return set()
    with open(LOG_PATH, encoding="utf-8-sig", newline="") as f:
        return {(r["Source link"].strip().lower(), r["Project"].strip().lower()) for r in csv.DictReader(f)}


def append_csv(path, fields, rows):
    new = not path.exists()
    with open(path, "a", encoding="utf-8-sig" if new else "utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        if new:
            w.writeheader()
        w.writerows(rows)


def main():
    p = argparse.ArgumentParser(description="AI web research for missing cells -> corrections.csv")
    p.add_argument("--model", default=os.environ.get("OPENAI_RESEARCH_MODEL") or "gpt-4.1-mini")
    p.add_argument("--fields", default=",".join(ALL_FIELDS),
                   help="Comma-separated subset of: " + ", ".join(ALL_FIELDS))
    p.add_argument("--limit", type=int, default=0, help="Research at most this many records.")
    p.add_argument("--redo", action="store_true", help="Research records again even if done before.")
    args = p.parse_args()

    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise SystemExit("Set OPENAI_API_KEY first (Windows: set OPENAI_API_KEY=sk-...).")
    wanted = [f.strip() for f in args.fields.split(",") if f.strip() in FIELD_HELP]

    with open(bd.DATA_JSON_PATH, encoding="utf-8") as f:
        records = json.load(f)["ministryProjects"]
    done = set() if args.redo else load_done()
    todo = []
    for r in records:
        gaps = missing_fields(r, wanted)
        key = (str(r.get("sourceLink") or "").strip().lower(), str(r.get("project") or "").strip().lower())
        if gaps and key not in done:
            todo.append((r, gaps))
    print(f"{len(records)} records, {len(todo)} with gaps still to research"
          + (f" (doing the first {args.limit})" if args.limit else ""))
    if args.limit:
        todo = todo[:args.limit]

    usage = {"in": 0, "out": 0, "calls": 0}
    proposed = 0
    for i, (r, gaps) in enumerate(todo, 1):
        result, calls = ask(api_key, args.model, r, gaps, usage)
        rows, rejected = [], []
        for item in result.get("found", []):
            if item["field"] not in gaps or not item["evidence_url"].startswith("http"):
                continue
            why = check(item["field"], item["value"], item["evidence_quote"])
            if why:
                rejected.append(f"{item['field']}: {why}")
                continue
            rows.append({
                "Status": "pending", "Source link": r.get("sourceLink", ""), "Project": r.get("project", ""),
                "Field": item["field"], "Current value": show(r.get(item["field"])),
                "New value": item["value"].strip(), "Evidence URL": item["evidence_url"],
                "Evidence quote": item["evidence_quote"].strip(),
                "Method": f"AI web research ({args.model}), confidence {item['confidence']}", "Notes": "",
            })
        if rows:
            append_csv(CORRECTIONS_PATH, CORRECTION_FIELDS, rows)
        append_csv(LOG_PATH, LOG_FIELDS, [{
            "Source link": r.get("sourceLink", ""), "Project": r.get("project", ""),
            "Researched at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"),
            "Fields asked": ", ".join(gaps), "Fields found": ", ".join(x["Field"] for x in rows),
            "Search calls": calls, "Model": args.model}])
        proposed += len(rows)
        print(f"  [{i}/{len(todo)}] found {len(rows)}/{len(gaps)} ({', '.join(x['Field'] for x in rows) or '-'})"
              f"  {r.get('project', '')[:60]}"
              + (f"\n      skipped: {'; '.join(rejected)}" if rejected else ""))

    cost = usage["calls"] * 0.01 + usage["in"] / 1e6 * 0.40 + usage["out"] / 1e6 * 1.60
    print(f"\nDone: {proposed} proposed value(s) added to {CORRECTIONS_PATH.name} as 'pending'.")
    print(f"Web searches: {usage['calls']}, tokens: {usage['in']:,} in / {usage['out']:,} out "
          f"(about ${cost:.2f} with {args.model}).")
    print("Next: open corrections.csv, set Status to 'approved' or 'rejected', save as CSV UTF-8, "
          "then run: python sync_daily.py")


if __name__ == "__main__":
    main()
