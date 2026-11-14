#!/usr/bin/env python3
"""Build digest.html from jobs.xlsx matched roles."""
import datetime as dt
from html import escape
from openpyxl import load_workbook

wb = load_workbook("/home/user/job-agent/jobs.xlsx")
ws = wb.active

jobs = []
for r in range(2, ws.max_row + 1):
    company = ws.cell(row=r, column=1).value
    if not company:
        continue
    role = ws.cell(row=r, column=2).value
    loc = ws.cell(row=r, column=3).value
    pct = ws.cell(row=r, column=4).value
    why = ws.cell(row=r, column=5).value
    link_cell = ws.cell(row=r, column=6)
    url = link_cell.hyperlink.target if link_cell.hyperlink else ""
    jobs.append({
        "company": company, "role": role, "loc": loc,
        "pct": round((pct or 0) * 100), "why": why, "url": url,
    })

today = dt.date.today().strftime("%A, %d %B %Y")

cards = []
for j in sorted(jobs, key=lambda x: x["pct"], reverse=True):
    cards.append(f"""
    <tr>
      <td style="padding:16px 18px;border:1px solid #e4e8ef;border-radius:10px;background:#ffffff;">
        <div style="font-size:12px;font-weight:700;color:#1a4fd6;text-transform:uppercase;letter-spacing:.4px;">{escape(str(j['company']))}</div>
        <div style="font-size:17px;font-weight:700;color:#16263f;margin:4px 0 6px;">{escape(str(j['role']))}</div>
        <div style="font-size:13px;color:#55607a;margin-bottom:10px;">📍 {escape(str(j['loc']))} &nbsp;·&nbsp; <strong style="color:#16263f;">{j['pct']}% match</strong> &nbsp;·&nbsp; {escape(str(j['why']))}</div>
        <a href="{escape(str(j['url']))}" style="display:inline-block;background:#16263f;color:#ffffff;text-decoration:none;font-size:13px;font-weight:600;padding:8px 16px;border-radius:6px;">View &amp; apply ➜</a>
      </td>
    </tr>
    <tr><td style="height:12px;line-height:12px;">&nbsp;</td></tr>""")

html = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f4f6fa;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6fa;padding:24px 0;">
    <tr><td align="center">
      <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;font-family:Arial,Helvetica,sans-serif;">
        <tr><td style="padding:0 18px 18px;">
          <div style="font-size:22px;font-weight:800;color:#16263f;">Daily PM Jobs</div>
          <div style="font-size:13px;color:#55607a;margin-top:4px;">{today} &nbsp;·&nbsp; {len(jobs)} strong match{'es' if len(jobs)!=1 else ''} (≥60%) — Product Manager, Cybersecurity / Cloud, Israel focus</div>
        </td></tr>
        {''.join(cards)}
        <tr><td style="padding:14px 18px 0;font-size:11px;color:#9aa3b5;border-top:1px solid #e4e8ef;">
          Scored by keyword relevance (no AI scoring key set). Sources: Tenable, Cymulate, Orca, Wiz, Cato Networks, Armis, Guardz.
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>"""

with open("/home/user/job-agent/digest.html", "w", encoding="utf-8") as f:
    f.write(html)

print(f"Wrote digest.html with {len(jobs)} jobs")
