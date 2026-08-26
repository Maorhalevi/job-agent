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
Maor Halevi — Product Manager · Cybersecurity. 5+ years experience.

PROFILE: Strategic Product Manager with 5+ years in cybersecurity and technical
analysis. Expertise in Exposure Management, Cloud Protection, and API-driven
workflows. Bridges technical engineering with business goals, delivering
automated remediation for complex security environments.

EXPERIENCE
- Product Manager, Exposure Management @ Cymulate (Jun 2024 - Mar 2026):
  Led research and integration of Cloud Protection capabilities into the platform;
  enriched the Exposure Management offering (attack-surface reduction); drove
  Auto-Mitigation and Remediation Plan from planning to launch (action-oriented
  platform); partnered with large/mid enterprise customers; led Design Partner
  program.
- Technical Product Manager @ Proofpoint (Jun 2022 - Apr 2024):
  Expanded the CASB product's SaaS application protection offering (market trends,
  customer needs, competitive landscape); led API research and integration
  understanding from a business angle; owned features end-to-end with engineering.
- Inbound Product Manager @ Menorah Mivtachim (2020 - Jun 2022):
  Product Owner for health & insurance systems; requirements + cross-dept integration.
- Full Stack Developer @ Menorah Mivtachim (2018 - 2020): Big Data (Elasticsearch,
  Kibana), end-to-end full-stack.
- Founder @ MindCETEX (2015 - 2018): AI-based EdTech app, full product lifecycle.
- Full Stack Developer @ Sapiens (2014 - 2017): Java ERP/insurance modules.

DOMAINS: Cybersecurity & SaaS Security, Exposure Management, Cloud Protection,
API Strategy & Integrations.
PRODUCT SKILLS: Product Strategy, Roadmap Prioritization, Design Partner Programs,
Jira, Confluence, Pendo, Figma, Postman.
TECHNICAL: SQL, APIs, ERD, AWS, Cloud Infrastructure, Elasticsearch, Kibana, Java.
AI TOOLS: Cursor, Claude, ChatGPT, v0, Lovable, Gemini, Perplexity, Copilot.

TARGET: Senior / Principal / Lead Product Manager in cybersecurity or cloud.
LOCATION: Israel (Tel Aviv area) or remote open to Israel / EMEA.
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
    # All slugs below verified live (cyber/cloud companies hiring PMs in Israel).
    ("greenhouse",    "tenableinc",   "Tenable"),         # exposure mgmt — top fit
    ("greenhouse",    "cymulate",     "Cymulate"),        # your domain
    ("ashby",         "orca",         "Orca Security"),   # agentless cloud security
    ("greenhouse",    "wizinc",       "Wiz"),             # cloud security (CNAPP)
    ("greenhouse_eu", "catonetworks", "Cato Networks"),   # SASE / network security
    ("greenhouse",    "armissecurity","Armis"),           # asset / threat mgmt
    ("greenhouse_eu", "guardz",       "Guardz"),          # SMB security
    # To add more: open the company's careers page and copy the slug from the URL
    #   boards.greenhouse.io/SLUG  /  job-boards.eu.greenhouse.io/...for=SLUG (use greenhouse_eu)
    #   jobs.lever.co/SLUG (lever)  /  jobs.ashbyhq.com/SLUG (ashby)
]

# ─────────────────────────────────────────────────────────────────
# 3. FETCHERS
# ─────────────────────────────────────────────────────────────────
TIMEOUT = 25
UA = {"User-Agent": "personal-job-agent/2.0"}

import re, html as _html
def _strip(text):
    """Strip HTML tags/entities and collapse whitespace; cap length for cost."""
    if not text:
        return ""
    text = re.sub(r"<[^>]+>", " ", text)
    text = _html.unescape(text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:1800]

def _job(title, location, url, company, desc="", contact=""):
    return {"title": (title or "").strip(), "location": (location or "").strip(),
            "url": url or "", "company": company,
            "desc": _strip(desc), "contact": contact}

def fetch_greenhouse(slug, company):
    url = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true"
    r = requests.get(url, headers=UA, timeout=TIMEOUT); r.raise_for_status()
    return [_job(j.get("title"), (j.get("location") or {}).get("name"),
                 j.get("absolute_url"), company, j.get("content"))
            for j in r.json().get("jobs", [])]

def fetch_lever(slug, company):
    url = f"https://api.lever.co/v0/postings/{slug}?mode=json"
    r = requests.get(url, headers=UA, timeout=TIMEOUT); r.raise_for_status()
    return [_job(j.get("text"), (j.get("categories") or {}).get("location"),
                 j.get("hostedUrl"), company,
                 j.get("descriptionPlain") or j.get("description"))
            for j in r.json()]

def fetch_ashby(slug, company):
    url = f"https://api.ashbyhq.com/posting-api/job-board/{slug}"
    r = requests.get(url, headers=UA, timeout=TIMEOUT); r.raise_for_status()
    return [_job(j.get("title"), j.get("location"), j.get("jobUrl"), company,
                 j.get("descriptionPlain") or j.get("description"))
            for j in r.json().get("jobs", [])]

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
    """Score each role 0-100 by matching the FULL job description to the CV.
    Processed in batches to keep each request small and reliable."""
    from anthropic import Anthropic
    client = Anthropic(api_key=ANTHROPIC_API_KEY)
    BATCH = 8
    for start in range(0, len(jobs), BATCH):
        batch = jobs[start:start + BATCH]
        blocks = "\n\n".join(
            f"[{i}] {j['title']} @ {j['company']} ({j['location']})\n"
            f"Description: {j['desc'] or '(no description available)'}"
            for i, j in enumerate(batch))
        prompt = (
            f"You are a recruiting assistant. Here is the candidate's CV:\n"
            f"{CANDIDATE_PROFILE}\n\n"
            f"For each job below, judge how well the candidate's CV matches the "
            f"role's actual requirements (seniority, domain, skills). Give a match "
            f"score 0-100 (100 = strong fit) and a reason of max 12 words citing "
            f"the key matching or missing requirement.\n"
            f'Return ONLY a JSON array: [{{"i":<index>,"score":<0-100>,"why":"..."}}]. '
            f"No prose, no markdown.\n\nJobs:\n{blocks}")
        try:
            msg = client.messages.create(model=ANTHROPIC_MODEL, max_tokens=1500,
                                         messages=[{"role": "user", "content": prompt}])
            text = "".join(b.text for b in msg.content if b.type == "text").strip()
            text = text.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
            scores = {d["i"]: d for d in json.loads(text)}
        except Exception as e:
            print(f"  scoring batch {start} failed: {e}")
            scores = {}
        for i, j in enumerate(batch):
            s = scores.get(i, {"score": 0, "why": "not scored"})
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
# 6b. HTML DIGEST OUTPUT  (for email delivery)
# ─────────────────────────────────────────────────────────────────
HTML_FILE = "digest.html"

def write_html(jobs, path=HTML_FILE):
    """Write an email-friendly HTML digest of matched roles.

    The delivery routine reads this file and emails its contents. When there
    are no matches the file is still written (with an empty-state note) so the
    routine can detect the no-jobs case."""
    today = dt.date.today().strftime("%B %d, %Y")
    n = len(jobs)
    rows = ""
    for j in jobs:
        score = int(round(j.get("score", 0)))
        rows += f"""
      <tr>
        <td style="padding:12px 14px;border-bottom:1px solid #e6e8eb;vertical-align:top;">
          <div style="font-weight:600;color:#16263F;font-size:15px;">{_html.escape(j['title'])}</div>
          <div style="color:#5b6470;font-size:13px;margin-top:2px;">{_html.escape(j['company'])} &middot; {_html.escape(j['location'])}</div>
        </td>
        <td style="padding:12px 14px;border-bottom:1px solid #e6e8eb;vertical-align:top;text-align:center;white-space:nowrap;">
          <span style="display:inline-block;background:#eaf1ff;color:#1A4FD6;font-weight:700;font-size:13px;padding:4px 10px;border-radius:12px;">{score}%</span>
        </td>
        <td style="padding:12px 14px;border-bottom:1px solid #e6e8eb;vertical-align:top;text-align:right;white-space:nowrap;">
          <a href="{_html.escape(j['url'])}" style="display:inline-block;background:#16263F;color:#ffffff;text-decoration:none;font-size:13px;font-weight:600;padding:8px 16px;border-radius:6px;">Apply &rarr;</a>
        </td>
      </tr>"""

    if jobs:
        body = f"""
      <table style="width:100%;border-collapse:collapse;">{rows}
      </table>"""
        subtitle = f"{n} strong match{'es' if n != 1 else ''} (&ge;{MIN_SCORE_PCT}%)"
    else:
        body = """
      <p style="color:#5b6470;font-size:14px;padding:8px 14px;">No strong matches today &mdash; agent ran successfully.</p>"""
        subtitle = "No strong matches today"

    doc = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>
<body style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;">
  <div style="max-width:640px;margin:0 auto;padding:24px 16px;">
    <div style="background:#16263F;border-radius:10px 10px 0 0;padding:22px 24px;">
      <h1 style="margin:0;color:#ffffff;font-size:20px;">Daily PM Jobs &mdash; Cybersecurity / Israel</h1>
      <div style="color:#a9b6c9;font-size:13px;margin-top:4px;">{today} &middot; {subtitle}</div>
    </div>
    <div style="background:#ffffff;border-radius:0 0 10px 10px;padding:8px 10px 16px;box-shadow:0 1px 3px rgba(0,0,0,0.06);">{body}
    </div>
    <div style="text-align:center;color:#8a939e;font-size:12px;margin-top:16px;">
      Generated by your daily job agent.
    </div>
  </div>
</body>
</html>"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)
    print(f"Wrote {path} with {n} rows")

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

    write_excel(jobs)
    write_html(jobs)

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
