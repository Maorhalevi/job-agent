from openpyxl import load_workbook
import html, datetime as dt

wb = load_workbook('jobs.xlsx')
ws = wb.active
rows = list(ws.iter_rows())
jobs = []
for r in rows[1:]:
    company, role, loc, match, why = r[0].value, r[1].value, r[2].value, r[3].value, r[4].value
    link = r[5].hyperlink.target if r[5].hyperlink else (r[5].value or "")
    if company is None:  # skip empty/placeholder
        continue
    jobs.append((company, role, loc, match, why, link))

today = "2026-07-26"
parts = []
parts.append('<div style="font-family:Arial,Helvetica,sans-serif;color:#16263F;max-width:760px;margin:0 auto;">')
parts.append(f'<h2 style="color:#16263F;margin-bottom:2px;">Daily PM Jobs — {today}</h2>')
parts.append(f'<p style="color:#555;margin-top:0;font-size:14px;">{len(jobs)} strong match(es) in Israel today.</p>')
parts.append('<table style="border-collapse:collapse;width:100%;font-size:14px;">')
parts.append('<tr style="background:#16263F;color:#fff;text-align:left;">'
             '<th style="padding:8px;">Company</th><th style="padding:8px;">Role</th>'
             '<th style="padding:8px;">Location</th><th style="padding:8px;">Match</th>'
             '<th style="padding:8px;">Why</th><th style="padding:8px;">Apply</th></tr>')
for i, (company, role, loc, match, why, link) in enumerate(jobs):
    bg = "#f6f8fb" if i % 2 else "#ffffff"
    pct = f"{round(match*100)}%" if isinstance(match, (int, float)) else str(match)
    parts.append(
        f'<tr style="background:{bg};">'
        f'<td style="padding:8px;border-bottom:1px solid #ddd;"><b>{html.escape(str(company))}</b></td>'
        f'<td style="padding:8px;border-bottom:1px solid #ddd;">{html.escape(str(role))}</td>'
        f'<td style="padding:8px;border-bottom:1px solid #ddd;">{html.escape(str(loc))}</td>'
        f'<td style="padding:8px;border-bottom:1px solid #ddd;font-weight:bold;color:#1A4FD6;">{pct}</td>'
        f'<td style="padding:8px;border-bottom:1px solid #ddd;color:#555;">{html.escape(str(why))}</td>'
        f'<td style="padding:8px;border-bottom:1px solid #ddd;"><a href="{html.escape(str(link))}" style="color:#1A4FD6;">Open &#10132;</a></td>'
        f'</tr>')
parts.append('</table>')
parts.append('<p style="color:#888;font-size:12px;margin-top:16px;">Sources: Greenhouse / Ashby ATS feeds for curated cyber &amp; cloud companies. Scored against your CV.</p>')
parts.append('</div>')

htmlout = "\n".join(parts)
with open('digest.html', 'w', encoding='utf-8') as f:
    f.write(htmlout)
print("Wrote digest.html", len(htmlout), "chars,", len(jobs), "jobs")
