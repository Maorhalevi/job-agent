from openpyxl import load_workbook
import datetime as dt, html

wb = load_workbook('jobs.xlsx'); ws = wb.active
rows = []
for r in range(2, ws.max_row+1):
    company = ws.cell(row=r,column=1).value
    role    = ws.cell(row=r,column=2).value
    loc     = ws.cell(row=r,column=3).value
    pct     = ws.cell(row=r,column=4).value
    why     = ws.cell(row=r,column=5).value
    link    = ws.cell(row=r,column=6).hyperlink
    url     = link.target if link else ''
    if company and 'No strong matches' not in str(company):
        rows.append((company, role, loc, pct, why, url))

today = dt.date.today().strftime('%B %d, %Y')
e = html.escape
trs = ""
for company, role, loc, pct, why, url in rows:
    pct_txt = f"{round((pct or 0)*100)}%"
    trs += f"""    <tr>
      <td style="padding:10px 12px;border-bottom:1px solid #e6e6e6;font-weight:600;">{e(str(company))}</td>
      <td style="padding:10px 12px;border-bottom:1px solid #e6e6e6;"><a href="{e(url)}" style="color:#1a4fd6;text-decoration:none;">{e(str(role))}</a></td>
      <td style="padding:10px 12px;border-bottom:1px solid #e6e6e6;color:#555;">{e(str(loc))}</td>
      <td style="padding:10px 12px;border-bottom:1px solid #e6e6e6;text-align:center;font-weight:600;color:#16a766;">{pct_txt}</td>
      <td style="padding:10px 12px;border-bottom:1px solid #e6e6e6;color:#555;">{e(str(why))}</td>
    </tr>
"""

doc = f"""<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#f4f5f7;font-family:Arial,Helvetica,sans-serif;color:#16263f;">
  <div style="max-width:760px;margin:0 auto;padding:24px;">
    <h2 style="margin:0 0 4px;">Daily PM Jobs &mdash; {len(rows)} match{'es' if len(rows)!=1 else ''}</h2>
    <p style="margin:0 0 18px;color:#777;">Product Manager / Cybersecurity roles in Israel &middot; {today}</p>
    <table style="border-collapse:collapse;width:100%;background:#fff;border:1px solid #e6e6e6;border-radius:6px;overflow:hidden;">
      <thead>
        <tr style="background:#16263f;color:#fff;text-align:left;">
          <th style="padding:10px 12px;">Company</th>
          <th style="padding:10px 12px;">Role</th>
          <th style="padding:10px 12px;">Location</th>
          <th style="padding:10px 12px;text-align:center;">Match</th>
          <th style="padding:10px 12px;">Why it fits</th>
        </tr>
      </thead>
      <tbody>
{trs}      </tbody>
    </table>
    <p style="margin:18px 0 0;color:#999;font-size:12px;">Generated automatically by job_agent.py. Match scores are keyword-based (AI scoring unavailable this run).</p>
  </div>
</body>
</html>
"""

open('digest.html','w').write(doc)
print(f"Wrote digest.html with {len(rows)} jobs")
