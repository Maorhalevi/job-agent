from openpyxl import load_workbook
import html, datetime as dt

wb = load_workbook('jobs.xlsx')
ws = wb.active
rows = []
for r in range(2, ws.max_row + 1):
    company = ws.cell(row=r, column=1).value
    role    = ws.cell(row=r, column=2).value
    loc     = ws.cell(row=r, column=3).value
    pct     = ws.cell(row=r, column=4).value
    why     = ws.cell(row=r, column=5).value
    linkc   = ws.cell(row=r, column=6)
    url     = linkc.hyperlink.target if linkc.hyperlink else ''
    if not company:  # skip the "no matches" placeholder row
        continue
    rows.append((company, role, loc, pct, why, url))

today = "2026-07-09"
def esc(x): return html.escape(str(x if x is not None else ''))

cards = []
for company, role, loc, pct, why, url in rows:
    pct_disp = f"{round((pct or 0)*100)}%"
    cards.append(f"""
      <tr>
        <td style="padding:14px 16px;border-bottom:1px solid #e6e8eb;">
          <div style="font-size:16px;font-weight:600;color:#16263F;">{esc(role)}</div>
          <div style="font-size:13px;color:#5a6472;margin-top:2px;">{esc(company)} &middot; {esc(loc)}</div>
          <div style="font-size:12px;color:#7a8290;margin-top:6px;">Why it fits: {esc(why)}</div>
        </td>
        <td style="padding:14px 16px;border-bottom:1px solid #e6e8eb;text-align:center;white-space:nowrap;">
          <span style="display:inline-block;background:#eaf1ff;color:#1A4FD6;font-weight:700;font-size:14px;padding:4px 10px;border-radius:12px;">{pct_disp}</span>
        </td>
        <td style="padding:14px 16px;border-bottom:1px solid #e6e8eb;text-align:right;white-space:nowrap;">
          <a href="{esc(url)}" style="display:inline-block;background:#16263F;color:#ffffff;text-decoration:none;font-size:13px;font-weight:600;padding:8px 14px;border-radius:6px;">Open &#10142;</a>
        </td>
      </tr>""")

doc = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;">
  <div style="max-width:640px;margin:0 auto;padding:24px 12px;">
    <div style="background:#16263F;color:#fff;padding:22px 24px;border-radius:10px 10px 0 0;">
      <div style="font-size:20px;font-weight:700;">Daily PM Jobs &mdash; Cybersecurity / Israel</div>
      <div style="font-size:13px;color:#b9c4d6;margin-top:4px;">{today} &middot; {len(rows)} strong match{'es' if len(rows)!=1 else ''} today</div>
    </div>
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#ffffff;border-radius:0 0 10px 10px;border-collapse:collapse;">
      {''.join(cards)}
    </table>
    <div style="font-size:11px;color:#98a1ae;margin-top:14px;line-height:1.5;">
      Sourced from Greenhouse / Ashby ATS feeds (Tenable, Cymulate, Orca, Wiz, Cato, Armis, Guardz).
      Match % is a keyword-based proxy (Claude scoring was unavailable this run &mdash; no API key set).
    </div>
  </div>
</body></html>"""

with open('digest.html', 'w', encoding='utf-8') as f:
    f.write(doc)
print(f"Wrote digest.html with {len(rows)} jobs, {len(doc)} bytes")
