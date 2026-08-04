#!/usr/bin/env python3
"""Build digest.html from the jobs.xlsx produced by job_agent.py."""
import datetime as dt
import html
from openpyxl import load_workbook

wb = load_workbook("jobs.xlsx")
ws = wb.active

rows = []
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
    company = ws.cell(row=i, column=1).value
    role = ws.cell(row=i, column=2).value
    location = ws.cell(row=i, column=3).value
    match = ws.cell(row=i, column=4).value
    why = ws.cell(row=i, column=5).value
    link_cell = ws.cell(row=i, column=6)
    url = link_cell.hyperlink.target if link_cell.hyperlink else ""
    # Skip the "no matches" placeholder row (no url, no company match%)
    if company and match is not None:
        rows.append({
            "company": company, "role": role, "location": location,
            "match": match, "why": why, "url": url,
        })

today = dt.date.today().strftime("%A, %B %d, %Y")

def esc(x):
    return html.escape(str(x if x is not None else ""))

parts = []
parts.append(f"""<div style="font-family:Arial,Helvetica,sans-serif;max-width:720px;margin:0 auto;color:#16263F;">
  <h2 style="color:#16263F;margin:0 0 4px;">Daily PM Jobs — {esc(today)}</h2>
  <p style="color:#555;margin:0 0 16px;font-size:14px;">{len(rows)} matched Product Manager role(s) in Israel (match &ge; 60%).</p>
  <table style="border-collapse:collapse;width:100%;font-size:14px;">
    <thead>
      <tr style="background:#16263F;color:#fff;text-align:left;">
        <th style="padding:8px;">Company</th>
        <th style="padding:8px;">Role</th>
        <th style="padding:8px;">Location</th>
        <th style="padding:8px;">Match</th>
        <th style="padding:8px;">Why</th>
        <th style="padding:8px;">Apply</th>
      </tr>
    </thead>
    <tbody>""")

for j in rows:
    pct = f"{round(j['match']*100)}%"
    apply = (f'<a href="{esc(j["url"])}" style="color:#1A4FD6;">Open &#10148;</a>'
             if j["url"] else "")
    parts.append(f"""      <tr style="border-bottom:1px solid #ddd;">
        <td style="padding:8px;">{esc(j['company'])}</td>
        <td style="padding:8px;font-weight:bold;">{esc(j['role'])}</td>
        <td style="padding:8px;">{esc(j['location'])}</td>
        <td style="padding:8px;">{pct}</td>
        <td style="padding:8px;color:#555;">{esc(j['why'])}</td>
        <td style="padding:8px;">{apply}</td>
      </tr>""")

parts.append("""    </tbody>
  </table>
  <p style="color:#999;font-size:12px;margin-top:16px;">Generated automatically by job_agent.py.</p>
</div>""")

digest = "\n".join(parts)
with open("digest.html", "w", encoding="utf-8") as f:
    f.write(digest)
print(f"Wrote digest.html with {len(rows)} jobs")
