#!/usr/bin/env python3
"""
=====================================================================
 DAILY JOB AGENT  —  Product Manager / Cybersecurity
=====================================================================
Pulls open roles every day directly from companies' PUBLIC ATS feeds
(Greenhouse / Lever / Ashby), filters to relevant Product Manager
roles, scores each one against YOUR profile with Claude, and emails
you a clean digest with direct apply links.

WHY THIS APPROACH WORKS BEST
----------------------------
  * It reads jobs straight from the source (the company's own ATS),
    so data is fresh and complete — no scraping, no LinkedIn blocks.
  * One HTTP request per company returns ALL their open roles.
  * These endpoints are public & unauthenticated (the same ones the
    companies' own career pages call).

HOW TO RUN
----------
  1.  pip install requests anthropic
  2.  Fill in the CONFIG section below (email + Anthropic API key).
  3.  python job_agent.py            # runs once, sends you a digest
  4.  Schedule it daily (cron / GitHub Actions — see bottom of file).

Tech is replaceable; the logic is the asset. Edit COMPANIES freely.
=====================================================================
"""

import os
import smtplib
import json
import datetime as dt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import requests

# ─────────────────────────────────────────────────────────────────
# 1. CONFIG  —  edit this block
# ─────────────────────────────────────────────────────────────────

CANDIDATE_PROFILE = """
Product Manager, 5+ years, cybersecurity & technical analysis.
Core expertise: Exposure Management, Cloud Protection, attack-surface
reduction, automated remediation (Auto-Mitigation / Remediation Plan),
API-driven workflows, SaaS security (CASB), Design Partner programs.
Seniority target: Senior PM / Principal / Lead PM.
Location preference: Tel Aviv / Israel, or fully remote.
Background companies: Cymulate, Proofpoint.
"""

# Roles must contain at least one of these to be considered a PM role.
TITLE_MUST_INCLUDE = ["product manager", "product owner", "product lead",
                      "head of product", "director of product", "pm,"]

# Boost / relevance keywords (used by the pre-filter and shown to Claude).
RELEVANCE_KEYWORDS = ["security", "cyber", "exposure", "cloud", "saas",
                      "remediation", "threat", "casb", "posture", "risk",
                      "vulnerability", "detection", "soc", "iam"]

# Only keep roles in these locations (substring match, case-insensitive).
# Leave as [] to keep every location.
LOCATION_FILTER = ["israel", "tel aviv", "remote"]

MIN_SCORE = 6          # 0-10; only email roles Claude scores >= this
USE_CLAUDE = False     # set False to skip AI scoring (keyword-only mode)
ANTHROPIC_MODEL = "claude-haiku-4-5-20251001"   # cheap + fast for scoring

# --- Secrets: prefer environment variables over hard-coding ---
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
EMAIL_FROM   = os.environ.get("EMAIL_FROM",   "you@gmail.com")
EMAIL_TO     = os.environ.get("EMAIL_TO",     "you@gmail.com")
EMAIL_PASS   = os.environ.get("EMAIL_PASS",   "")   # Gmail App Password
SMTP_HOST    = os.environ.get("SMTP_HOST",    "smtp.gmail.com")
SMTP_PORT    = int(os.environ.get("SMTP_PORT", "587"))

# ─────────────────────────────────────────────────────────────────
# 2. COMPANIES  —  the list of ATS boards to scan
# ─────────────────────────────────────────────────────────────────
# Each entry: ("ats_type", "company_slug", "Display Name")
# ats_type is one of: "greenhouse", "greenhouse_eu", "lever", "ashby"
#
# HOW TO FIND A SLUG: open the company's careers page. The URL tells you:
#   boards.greenhouse.io/ACME        -> ("greenhouse", "acme", ...)
#   job-boards.eu.greenhouse.io/ACME -> ("greenhouse_eu", "acme", ...)
#   jobs.lever.co/ACME               -> ("lever", "acme", ...)
#   jobs.ashbyhq.com/ACME            -> ("ashby", "acme", ...)
#
# The slugs below were verified live. Add as many cyber/PM-heavy
# companies as you like — more boards = more coverage.
COMPANIES = [
    ("greenhouse",    "tenableinc",       "Tenable"),       # exposure mgmt
    ("greenhouse_eu", "guardz",           "Guardz"),        # SMB security
    ("greenhouse",    "armissecurity",    "Armis"),         # asset/threat
    # --- add more verified slugs here, e.g.: ---
    # ("greenhouse",  "cymulate",         "Cymulate"),
    # ("greenhouse",  "wiz",              "Wiz"),
    # ("greenhouse",  "orcasecurity",     "Orca Security"),
    # ("lever",       "sentinelone",      "SentinelOne"),
    # ("ashby",       "torq",             "Torq"),
    # ("greenhouse",  "varonis",          "Varonis"),
]

# ─────────────────────────────────────────────────────────────────
# 3. ATS FETCHERS  —  one public endpoint per platform
# ─────────────────────────────────────────────────────────────────
TIMEOUT = 20
UA = {"User-Agent": "personal-job-agent/1.0"}

def _norm(title, location, url, company):
    return {"title": title.strip(), "location": (location or "").strip(),
            "url": url, "company": company}

def fetch_greenhouse(slug, company, eu=False):
    base = "https://boards-api.greenhouse.io/v1/boards"
    url = f"{base}/{slug}/jobs?content=false"
    r = requests.get(url, headers=UA, timeout=TIMEOUT)
    r.raise_for_status()
    out = []
    for j in r.json().get("jobs", []):
        loc = (j.get("location") or {}).get("name", "")
        out.append(_norm(j.get("title", ""), loc, j.get("absolute_url", ""), company))
    return out

def fetch_lever(slug, company):
    url = f"https://api.lever.co/v0/postings/{slug}?mode=json"
    r = requests.get(url, headers=UA, timeout=TIMEOUT)
    r.raise_for_status()
    out = []
    for j in r.json():
        loc = (j.get("categories") or {}).get("location", "")
        out.append(_norm(j.get("text", ""), loc, j.get("hostedUrl", ""), company))
    return out

def fetch_ashby(slug, company):
    url = f"https://api.ashbyhq.com/posting-api/job-board/{slug}"
    r = requests.get(url, headers=UA, timeout=TIMEOUT)
    r.raise_for_status()
    out = []
    for j in r.json().get("jobs", []):
        out.append(_norm(j.get("title", ""), j.get("location", ""),
                         j.get("jobUrl", ""), company))
    return out

FETCHERS = {
    "greenhouse":    lambda s, c: fetch_greenhouse(s, c, eu=False),
    "greenhouse_eu": lambda s, c: fetch_greenhouse(s, c, eu=True),
    "lever":         fetch_lever,
    "ashby":         fetch_ashby,
}

def collect_all_jobs():
    jobs = []
    for ats, slug, name in COMPANIES:
        try:
            found = FETCHERS[ats](slug, name)
            jobs.extend(found)
            print(f"  [{name}] {len(found)} roles")
        except Exception as e:
            print(f"  [{name}] FAILED ({ats}/{slug}): {e}")
    return jobs

# ─────────────────────────────────────────────────────────────────
# 4. FILTERING  —  keep only relevant PM roles
# ─────────────────────────────────────────────────────────────────
def is_pm_role(title):
    t = title.lower()
    return any(k in t for k in TITLE_MUST_INCLUDE)

def location_ok(location):
    if not LOCATION_FILTER:
        return True
    l = location.lower()
    return any(k in l for k in LOCATION_FILTER)

def keyword_relevant(title):
    t = title.lower()
    return any(k in t for k in RELEVANCE_KEYWORDS)

def prefilter(jobs):
    seen, kept = set(), []
    for j in jobs:
        if j["url"] in seen:
            continue
        seen.add(j["url"])
        if is_pm_role(j["title"]) and location_ok(j["location"]):
            kept.append(j)
    return kept

# ─────────────────────────────────────────────────────────────────
# 5. SCORING  —  rank with Claude against your profile
# ─────────────────────────────────────────────────────────────────
def score_with_claude(jobs):
    from anthropic import Anthropic
    client = Anthropic(api_key=ANTHROPIC_API_KEY)
    listing = "\n".join(f"{i}. {j['title']} @ {j['company']} ({j['location']})"
                        for i, j in enumerate(jobs))
    prompt = (
        f"You are a recruiting assistant. Candidate profile:\n{CANDIDATE_PROFILE}\n\n"
        f"Score each job 0-10 for fit (10 = perfect). Return ONLY a JSON array of "
        f'objects: [{{"i": <index>, "score": <0-10>, "why": "<8 words max>"}}]. '
        f"No prose, no markdown.\n\nJobs:\n{listing}"
    )
    msg = client.messages.create(
        model=ANTHROPIC_MODEL, max_tokens=1500,
        messages=[{"role": "user", "content": prompt}],
    )
    text = "".join(b.text for b in msg.content if b.type == "text")
    text = text.strip().removeprefix("```json").removeprefix("```").removesuffix("```")
    scores = {d["i"]: d for d in json.loads(text)}
    for i, j in enumerate(jobs):
        s = scores.get(i, {"score": 0, "why": ""})
        j["score"], j["why"] = s["score"], s["why"]
    return sorted(jobs, key=lambda x: x["score"], reverse=True)

def score_with_keywords(jobs):
    for j in jobs:
        hits = sum(k in j["title"].lower() for k in RELEVANCE_KEYWORDS)
        j["score"] = min(10, 4 + 2 * hits)
        j["why"] = "keyword match" if hits else "PM role"
    return sorted(jobs, key=lambda x: x["score"], reverse=True)

# ─────────────────────────────────────────────────────────────────
# 6. DIGEST  —  build + send the email
# ─────────────────────────────────────────────────────────────────
def build_html(jobs):
    today = dt.date.today().strftime("%d %b %Y")
    if not jobs:
        return f"""
    <div style="font-family:Arial,sans-serif;max-width:680px;margin:auto;">
      <h2 style="color:#16263f;">Your daily job matches · {today}</h2>
      <p style="color:#555;">No strong matches today. The agent ran successfully —
        no new Product Manager roles cleared the relevance bar.</p>
      <p style="color:#999;font-size:12px;">Tip: add more companies to the
        COMPANIES list or lower MIN_SCORE to widen the net.</p>
    </div>"""
    rows = ""
    for j in jobs:
        rows += f"""
        <tr>
          <td style="padding:10px 8px;border-bottom:1px solid #eee;">
            <a href="{j['url']}" style="color:#16263f;font-weight:600;text-decoration:none;">
              {j['title']}</a><br>
            <span style="color:#666;font-size:13px;">{j['company']} · {j['location']}</span>
          </td>
          <td style="padding:10px 8px;border-bottom:1px solid #eee;text-align:center;
                     color:#16263f;font-weight:700;">{j['score']}/10</td>
          <td style="padding:10px 8px;border-bottom:1px solid #eee;color:#555;
                     font-size:13px;">{j['why']}</td>
        </tr>"""
    return f"""
    <div style="font-family:Arial,sans-serif;max-width:680px;margin:auto;">
      <h2 style="color:#16263f;">Your daily job matches · {today}</h2>
      <p style="color:#555;">{len(jobs)} relevant Product Manager roles found.</p>
      <table style="width:100%;border-collapse:collapse;">
        <tr style="background:#16263f;color:#fff;">
          <th style="padding:8px;text-align:left;">Role</th>
          <th style="padding:8px;">Fit</th>
          <th style="padding:8px;text-align:left;">Why</th>
        </tr>{rows}
      </table>
      <p style="color:#999;font-size:12px;margin-top:16px;">
        Pulled from public ATS feeds. Apply by clicking the role title.</p>
    </div>"""

def send_email(html):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Daily PM jobs · {dt.date.today():%d %b}"
    msg["From"], msg["To"] = EMAIL_FROM, EMAIL_TO
    msg.attach(MIMEText(html, "html"))
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(EMAIL_FROM, EMAIL_PASS)
        s.sendmail(EMAIL_FROM, [EMAIL_TO], msg.as_string())
    print("Digest sent.")

# ─────────────────────────────────────────────────────────────────
# 7. MAIN
# ─────────────────────────────────────────────────────────────────
def write_digest(html):
    """Always write digest.html — the routine reads this file every run."""
    with open("digest.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("Wrote digest.html")

def main():
    print("Collecting jobs from ATS feeds...")
    try:
        jobs = collect_all_jobs()
    except Exception as e:
        jobs = []
        print(f"Collection error: {e}")
    print(f"Raw roles: {len(jobs)}")

    jobs = prefilter(jobs)
    print(f"PM roles after filter: {len(jobs)}")

    if jobs:
        try:
            jobs = (score_with_claude(jobs) if (USE_CLAUDE and ANTHROPIC_API_KEY)
                    else score_with_keywords(jobs))
        except Exception as e:
            print(f"Scoring error, falling back to keywords: {e}")
            jobs = score_with_keywords(jobs)
        jobs = [j for j in jobs if j["score"] >= MIN_SCORE]
        print(f"Roles >= {MIN_SCORE}/10: {len(jobs)}")

    # Build + persist the digest on EVERY run (jobs may be empty — that's fine).
    html = build_html(jobs)
    write_digest(html)

    # Optional: only send via SMTP if a password is configured. When used inside
    # a Claude Code routine, leave EMAIL_PASS empty and let the routine send
    # digest.html through the Gmail / Slack connector instead.
    if EMAIL_PASS and jobs:
        try:
            send_email(html)
        except Exception as e:
            print(f"Email send failed (digest.html still written): {e}")

if __name__ == "__main__":
    main()

# =====================================================================
# SCHEDULE IT DAILY
# ---------------------------------------------------------------------
# Option A — cron (Mac/Linux), runs every day at 08:00:
#   crontab -e
#   0 8 * * *  cd /path/to/agent && /usr/bin/python3 job_agent.py
#
# Option B — GitHub Actions (free, no server). Create
# .github/workflows/jobs.yml :
#   name: daily-jobs
#   on:
#     schedule: [{cron: "0 6 * * *"}]   # 06:00 UTC daily
#     workflow_dispatch:
#   jobs:
#     run:
#       runs-on: ubuntu-latest
#       steps:
#         - uses: actions/checkout@v4
#         - uses: actions/setup-python@v5
#           with: {python-version: "3.12"}
#         - run: pip install requests anthropic
#         - run: python job_agent.py
#           env:
#             ANTHROPIC_API_KEY: ${{ secrets.ANTHROPIC_API_KEY }}
#             EMAIL_FROM: ${{ secrets.EMAIL_FROM }}
#             EMAIL_TO:   ${{ secrets.EMAIL_TO }}
#             EMAIL_PASS: ${{ secrets.EMAIL_PASS }}
# =====================================================================
