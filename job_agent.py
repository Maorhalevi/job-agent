#!/usr/bin/env python3
"""
=====================================================================
 DAILY JOB AGENT  —  Product Manager / Cybersecurity  (Israel focus)
=====================================================================
Finds relevant Product Manager roles every day and writes an Excel
file you can open and act on.

SOURCES (combined):
  1. Direct ATS feeds  (Greenhouse / Lever / Ashby) — curated companies
  2. Adzuna aggregator (optional) — broad coverage, country-filtered to IL

PIPELINE:
  collect  ->  Israel-only location filter  ->  PM-title filter
           ->  Claude match score (0-100%)  ->  keep >= MIN_SCORE_PCT
           ->  write jobs.xlsx  (Company | Role | Location | Match% |
                                 Why | Link | Contact)

DELIVERY:
  Designed for a Claude Code Routine: the script writes jobs.xlsx,
  and the routine sends that file to you via the Gmail / Slack connector.

RUN:
  pip install requests anthropic openpyxl
  python job_agent.py
=====================================================================
"""

import os
import json
import datetime as dt

import requests

# ─────────────────────────────────────────────────────────────────
# 1. CONFIG
# ─────────────────────────────────────────────────────────────────
CANDIDATE_PROFILE = """
Product Manager, 5+ years, cybersecurity & technical analysis.
Core expertise: Exposure Management, Cloud Protection, attack-surface
reduction, automated remediation (Auto-Mitigation / Remediation Plan),
API-driven workflows, SaaS security (CASB), Design Partner programs.
Seniority target: Senior PM / Principal / Lead PM.
Background companies: Cymulate, Proofpoint.
"""

TITLE_MUST_INCLUDE = ["product manager", "product owner", "product lead",
                      "head of product", "director of product",
                      "group product", "principal product", "vp product"]

RELEVANCE_KEYWORDS = ["security", "cyber", "exposure", "cloud", "saas",
                      "remediation", "threat", "casb", "posture", "risk",
                      "vulnerability", "detection", "soc", "iam", "api"]

# --- HARD location filter: Israel only (plus Israel-eligible remote) ---
ISRAEL_TERMS = ["israel", "tel aviv", "tel-aviv", "herzliya", "haifa",
                "jerusalem", "ramat gan", "petah tikva", "petach tikva",
                "be'er sheva", "beer sheva", "netanya", "raanana", "ra'anana",
                "yokneam", "caesarea", "rehovot", "kiryat", "il"]
REMOTE_TERMS = ["remote", "anywhere", "global", "worldwide", "emea"]
ALLOW_REMOTE = True   # include remote roles that look open to IL/EMEA/global

MIN_SCORE_PCT = 60          # keep only roles scored >= 60% match
USE_CLAUDE = True           # AI scoring (needs ANTHROPIC_API_KEY)
ANTHROPIC_MODEL = "claude-haiku-4-5-20251001"

OUTPUT_FILE = "jobs.xlsx"
HTML_FILE = "digest.html"

# --- Secrets (set as environment variables in the routine) ---
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ADZUNA_APP_ID  = os.environ.get("ADZUNA_APP_ID", "")     # optional
ADZUNA_APP_KEY = os.environ.get("ADZUNA_APP_KEY", "")    # optional

# ─────────────────────────────────────────────────────────────────
# 2. COMPANIES  —  ATS boards to scan
#    type: greenhouse | greenhouse_eu | lever | ashby
#    Wrong/closed slugs are skipped safely (logged, never crash).
#    [V] = verified live earlier. Prune any that log FAILED.
# ─────────────────────────────────────────────────────────────────
COMPANIES = [
    ("greenhouse",    "tenableinc",     "Tenable"),        # [V]
    ("greenhouse_eu", "guardz",         "Guardz"),         # [V]
    ("greenhouse",    "armissecurity",  "Armis"),          # [V]
    ("greenhouse",    "cymulate",       "Cymulate"),
    ("greenhouse",    "wiz",            "Wiz"),
    ("greenhouse",    "orcasecurity",   "Orca Security"),
    ("greenhouse",    "varonis",        "Varonis"),
    ("greenhouse",    "checkpoint",     "Check Point"),
    ("greenhouse",    "claroty",        "Claroty"),
    ("greenhouse",    "cybereason",     "Cybereason"),
    ("greenhouse",    "snyk",           "Snyk"),
    ("greenhouse",    "perimeterx",     "HUMAN (PerimeterX)"),
    ("greenhouse",    "salt",           "Salt Security"),
    ("greenhouse",    "noname",         "Noname Security"),
    ("greenhouse",    "transmitsecurity","Transmit Security"),
    ("lever",         "sentinelone",    "SentinelOne"),
    ("lever",         "aquasecurity",   "Aqua Security"),
    ("ashby",         "torq",           "Torq"),
    ("ashby",         "island",         "Island"),
    ("greenhouse",    "paloaltonetworks","Palo Alto Networks"),
]

# ─────────────────────────────────────────────────────────────────
# 3. FETCHERS
# ─────────────────────────────────────────────────────────────────
TIMEOUT = 25
UA = {"User-Agent": "personal-job-agent/2.0"}

def _job(title, location, url, company, contact=""):
    return {"title": (title or "").strip(), "location": (location or "").strip(),
            "url": url or "", "company": company, "contact": contact}

def fetch_greenhouse(slug, company):
    url = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=false"
    r = requests.get(url, headers=UA, timeout=TIMEOUT); r.raise_for_status()
    return [_job(j.get("title"), (j.get("location") or {}).get("name"),
                 j.get("absolute_url"), company) for j in r.json().get("jobs", [])]

def fetch_lever(slug, company):
    url = f"https://api.lever.co/v0/postings/{slug}?mode=json"
    r = requests.get(url, headers=UA, timeout=TIMEOUT); r.raise_for_status()
    return [_job(j.get("text"), (j.get("categories") or {}).get("location"),
                 j.get("hostedUrl"), company) for j in r.json()]

def fetch_ashby(slug, company):
    url = f"https://api.ashbyhq.com/posting-api/job-board/{slug}"
    r = requests.get(url, headers=UA, timeout=TIMEOUT); r.raise_for_status()
    return [_job(j.get("title"), j.get("location"),
                 j.get("jobUrl"), company) for j in r.json().get("jobs", [])]

FETCHERS = {"greenhouse": fetch_greenhouse, "greenhouse_eu": fetch_greenhouse,
            "lever": fetch_lever, "ashby": fetch_ashby}

def fetch_adzuna():
    """Optional broad source, filtered to Israel. Skipped if no keys."""
    if not (ADZUNA_APP_ID and ADZUNA_APP_KEY):
        print("  [Adzuna] skipped (no API keys set)")
        return []
    out = []
    for page in (1, 2):
        url = (f"https://api.adzuna.com/v1/api/jobs/il/search/{page}"
               f"?app_id={ADZUNA_APP_ID}&app_key={ADZUNA_APP_KEY}"
               f"&what=product%20manager&results_per_page=50&content-type=application/json")
        try:
            r = requests.get(url, headers=UA, timeout=TIMEOUT); r.raise_for_status()
            for j in r.json().get("results", []):
                out.append(_job(j.get("title"),
                                (j.get("location") or {}).get("display_name"),
                                j.get("redirect_url"),
                                (j.get("company") or {}).get("display_name", "Unknown")))
        except Exception as e:
            print(f"  [Adzuna] page {page} failed: {e}")
    print(f"  [Adzuna] {len(out)} roles")
    return out

def collect_all_jobs():
    jobs = []
    for ats, slug, name in COMPANIES:
        try:
            found = FETCHERS[ats](slug, name)
            jobs.extend(found)
            print(f"  [{name}] {len(found)} roles")
        except Exception as e:
            print(f"  [{name}] FAILED ({ats}/{slug}): {e}")
    jobs.extend(fetch_adzuna())
    return jobs

# ─────────────────────────────────────────────────────────────────
# 4. FILTERS
# ─────────────────────────────────────────────────────────────────
def is_pm_role(title):
    t = title.lower()
    return any(k in t for k in TITLE_MUST_INCLUDE)

def in_israel(location):
    l = location.lower()
    if any(term in l for term in ISRAEL_TERMS):
        return True
    if ALLOW_REMOTE and any(term in l for term in REMOTE_TERMS):
        return True
    return False

def prefilter(jobs):
    seen, kept = set(), []
    for j in jobs:
        key = (j["company"].lower(), j["title"].lower())
        if key in seen or not j["url"]:
            continue
        seen.add(key)
        if is_pm_role(j["title"]) and in_israel(j["location"]):
            kept.append(j)
    return kept

# ─────────────────────────────────────────────────────────────────
# 5. SCORING  (0-100 %)
# ─────────────────────────────────────────────────────────────────
def score_with_claude(jobs):
    from anthropic import Anthropic
    client = Anthropic(api_key=ANTHROPIC_API_KEY)
    listing = "\n".join(f"{i}. {j['title']} @ {j['company']} ({j['location']})"
                        for i, j in enumerate(jobs))
    prompt = (
        f"You are a recruiting assistant. Candidate profile:\n{CANDIDATE_PROFILE}\n\n"
        f"For each job, give a match score 0-100 (100 = perfect fit) and a reason "
        f"of max 10 words. Return ONLY a JSON array: "
        f'[{{"i":<index>,"score":<0-100>,"why":"<reason>"}}]. No prose, no markdown.\n\n'
        f"Jobs:\n{listing}"
    )
    msg = client.messages.create(model=ANTHROPIC_MODEL, max_tokens=2000,
                                 messages=[{"role": "user", "content": prompt}])
    text = "".join(b.text for b in msg.content if b.type == "text").strip()
    text = text.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
    scores = {d["i"]: d for d in json.loads(text)}
    for i, j in enumerate(jobs):
        s = scores.get(i, {"score": 0, "why": ""})
        j["score"], j["why"] = int(s["score"]), s["why"]
    return sorted(jobs, key=lambda x: x["score"], reverse=True)

def score_with_keywords(jobs):
    for j in jobs:
        hits = sum(k in j["title"].lower() for k in RELEVANCE_KEYWORDS)
        j["score"] = min(100, 50 + 12 * hits)   # rough 0-100 proxy
        j["why"] = f"{hits} keyword match(es)" if hits else "PM role"
    return sorted(jobs, key=lambda x: x["score"], reverse=True)

# ─────────────────────────────────────────────────────────────────
# 6. EXCEL OUTPUT
# ─────────────────────────────────────────────────────────────────
def write_excel(jobs, path=OUTPUT_FILE):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = Workbook(); ws = wb.active; ws.title = "Matches"
    headers = ["Company", "Role", "Location", "Match %", "Why it fits",
               "Apply link", "Contact (if public)"]
    ws.append(headers)

    navy = PatternFill("solid", fgColor="16263F")
    white_bold = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    thin = Side(style="thin", color="DDDDDD")
    border = Border(bottom=thin)
    for c in ws[1]:
        c.fill = navy; c.font = white_bold
        c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22

    for j in jobs:
        ws.append([j["company"], j["title"], j["location"],
                   j["score"] / 100, j["why"], j["url"], j.get("contact", "")])
        row = ws.max_row
        link = ws.cell(row=row, column=6)
        if j["url"]:
            link.hyperlink = j["url"]; link.value = "Open ➜"
            link.font = Font(name="Arial", color="1A4FD6", underline="single")
        ws.cell(row=row, column=4).number_format = "0%"
        for col in range(1, 8):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            if cell.font.name is None or col != 6:
                cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=(col in (2, 5)))

    widths = [22, 40, 22, 9, 34, 12, 24]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.freeze_panes = "A2"

    if not jobs:
        ws.append(["No strong matches today — agent ran successfully.",
                   "", "", "", "", "", ""])

    wb.save(path)
    print(f"Wrote {path} with {len(jobs)} rows")

# ─────────────────────────────────────────────────────────────────
# 6b. HTML DIGEST OUTPUT
# ─────────────────────────────────────────────────────────────────
def write_html(jobs, path=HTML_FILE):
    """Write an email-ready HTML digest of today's matched roles."""
    import html as _html
    today = dt.date.today().isoformat()

    if not jobs:
        body = (
            '<p style="font:15px Arial,sans-serif;color:#16263F;">'
            'No strong matches found today — the agent ran successfully.</p>'
        )
    else:
        rows = []
        for j in jobs:
            company = _html.escape(j["company"])
            title = _html.escape(j["title"])
            location = _html.escape(j["location"])
            why = _html.escape(j.get("why", ""))
            url = _html.escape(j["url"], quote=True)
            score = int(j["score"])
            link = (f'<a href="{url}" style="color:#1A4FD6;text-decoration:none;">'
                    f'Open ➜</a>') if j["url"] else ""
            rows.append(
                '<tr>'
                f'<td style="padding:8px 10px;border-bottom:1px solid #eee;">{company}</td>'
                f'<td style="padding:8px 10px;border-bottom:1px solid #eee;">{title}</td>'
                f'<td style="padding:8px 10px;border-bottom:1px solid #eee;">{location}</td>'
                f'<td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:center;font-weight:bold;">{score}%</td>'
                f'<td style="padding:8px 10px;border-bottom:1px solid #eee;">{why}</td>'
                f'<td style="padding:8px 10px;border-bottom:1px solid #eee;">{link}</td>'
                '</tr>'
            )
        headers = "".join(
            f'<th style="padding:8px 10px;text-align:left;color:#fff;font-weight:bold;">{h}</th>'
            for h in ["Company", "Role", "Location", "Match %", "Why it fits", "Apply"]
        )
        body = (
            '<table style="border-collapse:collapse;width:100%;'
            'font:14px Arial,sans-serif;color:#16263F;">'
            f'<thead><tr style="background:#16263F;">{headers}</tr></thead>'
            f'<tbody>{"".join(rows)}</tbody></table>'
        )

    doc = (
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        '<title>Daily PM jobs</title></head>'
        '<body style="margin:0;padding:20px;background:#f6f7f9;">'
        '<h2 style="font:20px Arial,sans-serif;color:#16263F;margin:0 0 4px;">'
        'Daily PM jobs</h2>'
        f'<p style="font:13px Arial,sans-serif;color:#667;margin:0 0 16px;">{today} · '
        f'{len(jobs)} matched role(s)</p>'
        f'{body}</body></html>'
    )

    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)
    print(f"Wrote {path} with {len(jobs)} matched role(s)")

# ─────────────────────────────────────────────────────────────────
# 7. MAIN
# ─────────────────────────────────────────────────────────────────
def main():
    print("Collecting jobs...")
    try:
        jobs = collect_all_jobs()
    except Exception as e:
        jobs = []; print(f"Collection error: {e}")
    print(f"Raw roles: {len(jobs)}")

    jobs = prefilter(jobs)
    print(f"PM roles in Israel after filter: {len(jobs)}")

    if jobs:
        try:
            jobs = (score_with_claude(jobs) if (USE_CLAUDE and ANTHROPIC_API_KEY)
                    else score_with_keywords(jobs))
        except Exception as e:
            print(f"Claude scoring failed, using keywords: {e}")
            jobs = score_with_keywords(jobs)
        jobs = [j for j in jobs if j["score"] >= MIN_SCORE_PCT]
        print(f"Roles >= {MIN_SCORE_PCT}%: {len(jobs)}")

    write_html(jobs)
    try:
        write_excel(jobs)
    except Exception as e:
        print(f"Excel output skipped: {e}")

if __name__ == "__main__":
    main()

# =====================================================================
# OPTIONAL — turn on the broad Adzuna source:
#   1. Register (free) at developer.adzuna.com  -> get app_id + app_key
#   2. Add env vars to the routine: ADZUNA_APP_ID, ADZUNA_APP_KEY
#   Without them the script just uses the ATS company list above.
#
# ROUTINE PROMPT (delivery):
#   Run `python job_agent.py`. It creates `jobs.xlsx` with today's matched
#   Product Manager roles in Israel. Attach that file and email it to me via
#   the Gmail connector, subject "Daily PM jobs (Israel)". If the file has no
#   data rows, send a one-line note that there were no strong matches today.
#
# NETWORK ALLOWLIST (routine environment):
#   boards-api.greenhouse.io, api.lever.co, api.ashbyhq.com,
#   api.anthropic.com, api.adzuna.com
# =====================================================================
