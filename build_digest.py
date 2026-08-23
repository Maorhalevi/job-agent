from openpyxl import load_workbook
import datetime as dt

wb = load_workbook('jobs.xlsx')
ws = wb.active
jobs = []
for r in range(2, ws.max_row+1):
    company = ws.cell(row=r, column=1).value
    if not company or "No strong matches" in str(company):
        continue
    role = ws.cell(row=r, column=2).value
    loc = ws.cell(row=r, column=3).value
    pct = ws.cell(row=r, column=4).value
    why = ws.cell(row=r, column=5).value
    link_cell = ws.cell(row=r, column=6)
    url = link_cell.hyperlink.target if link_cell.hyperlink else ""
    jobs.append((company, role, loc, pct, why, url))

today = dt.date.today().strftime("%A, %B %d, %Y")

rows = ""
for company, role, loc, pct, why, url in jobs:
    pct_disp = f"{round((pct or 0)*100)}%"
    rows += f"""
      <tr>
        <td style="padding:14px 16px;border-bottom:1px solid #e6e8ec;vertical-align:top;">
          <div style="font-weight:600;color:#16263F;font-size:15px;">{role}</div>
          <div style="color:#5b6472;font-size:13px;margin-top:2px;">{company} &middot; {loc}</div>
        </td>
        <td style="padding:14px 16px;border-bottom:1px solid #e6e8ec;text-align:center;vertical-align:top;">
          <span style="display:inline-block;background:#eef4ff;color:#1a4fd6;font-weight:700;font-size:14px;padding:4px 10px;border-radius:12px;">{pct_disp}</span>
        </td>
        <td style="padding:14px 16px;border-bottom:1px solid #e6e8ec;text-align:right;vertical-align:top;">
          <a href="{url}" style="display:inline-block;background:#16263F;color:#ffffff;text-decoration:none;font-size:13px;font-weight:600;padding:8px 16px;border-radius:6px;">Apply &rarr;</a>
        </td>
      </tr>"""

html = f"""<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#f4f5f7;font-family:Arial,Helvetica,sans-serif;">
  <div style="max-width:640px;margin:0 auto;padding:24px 12px;">
    <div style="background:#16263F;border-radius:10px 10px 0 0;padding:22px 24px;">
      <div style="color:#ffffff;font-size:20px;font-weight:700;">Daily PM Jobs</div>
      <div style="color:#9db0cc;font-size:13px;margin-top:4px;">{today} &middot; {len(jobs)} strong match{'es' if len(jobs)!=1 else ''} &middot; Product Manager / Cybersecurity, Israel</div>
    </div>
    <div style="background:#ffffff;border-radius:0 0 10px 10px;padding:8px 8px 16px;">
      <table style="width:100%;border-collapse:collapse;">
        {rows}
      </table>
    </div>
    <div style="text-align:center;color:#9098a4;font-size:11px;margin-top:16px;">
      Generated automatically by your job agent.
    </div>
  </div>
</body>
</html>"""

with open('digest.html','w') as f:
    f.write(html)
print(f"Wrote digest.html with {len(jobs)} jobs")
