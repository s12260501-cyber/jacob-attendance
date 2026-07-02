#!/usr/bin/env python3
"""Build script for Netlify: read Excel, generate index.html"""
import openpyxl
from pathlib import Path

EXCEL_PATH = Path("雅各組出席表.xlsx")
HTML_PATH = Path("index.html")

wb = openpyxl.load_workbook(EXCEL_PATH)
ws = wb["工作表3"]

rows = []
for row in ws.iter_rows(min_row=5, max_col=5, values_only=True):
    seq, name, apr, may, jun = row
    if not seq or str(seq).strip() == "總計":
        break
    rows.append({
        "seq": int(seq),
        "name": name,
        "apr": bool(apr),
        "may": bool(may),
        "jun": bool(jun),
    })

total_apr = sum(1 for r in rows if r["apr"])
total_may = sum(1 for r in rows if r["may"])
total_jun = sum(1 for r in rows if r["jun"])
total_people = len(rows)

def check_cell(present: bool) -> str:
    cls = "check present" if present else "check absent"
    return f'<td class="{cls}"></td>'

tbody_rows = []
for r in rows:
    total = sum([r["apr"], r["may"], r["jun"]])
    tbody_rows.append(
        f'<tr><td>{r["seq"]}</td><td>{r["name"]}</td>'
        f'{check_cell(r["apr"])}'
        f'{check_cell(r["may"])}'
        f'{check_cell(r["jun"])}'
        f'<td>{total}</td></tr>'
    )

tbody_rows.append(
    f'<tr style="background: var(--gray-100); font-weight: 600;">'
    f'<td colspan="2">總計</td>'
    f'<td class="check">{total_apr}</td>'
    f'<td class="check">{total_may}</td>'
    f'<td class="check">{total_jun}</td>'
    f'<td>{total_apr + total_may + total_jun}</td></tr>'
)

html = f'''<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>雅各組 出席表 Dashboard</title>
<style>
:root {{
  --font-sans: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans TC', sans-serif;
  --gray-50: #fafafa;
  --gray-100: #f5f5f5;
  --gray-200: #e5e5e5;
  --gray-400: #a3a3a3;
  --gray-600: #525252;
  --gray-800: #262626;
  --blue-500: #3b82f6;
  --green-500: #22c55e;
  --red-500: #ef4444;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ 
  font-family: var(--font-sans); 
  background: var(--gray-50); 
  color: var(--gray-800); 
  padding: 32px; 
  line-height: 1.6;
}}
.container {{ max-width: 960px; margin: 0 auto; }}
.header {{ 
  margin-bottom: 32px; 
  border-bottom: 1px solid var(--gray-200); 
  padding-bottom: 16px;
}}
.header h1 {{ 
  font-size: 24px; 
  font-weight: 600; 
  letter-spacing: -0.02em;
}}
.header .period {{ 
  font-size: 14px; 
  color: var(--gray-600); 
  margin-top: 4px;
}}
.summary {{ 
  display: grid; 
  grid-template-columns: repeat(3, 1fr); 
  gap: 16px; 
  margin-bottom: 32px;
}}
.card {{ 
  background: white; 
  border: 1px solid var(--gray-200); 
  padding: 20px; 
  text-align: center;
}}
.card .value {{ 
  font-size: 36px; 
  font-weight: 700; 
  color: var(--blue-500);
}}
.card .label {{ 
  font-size: 13px; 
  color: var(--gray-600); 
  margin-top: 4px;
}}
.card .rate {{ 
  font-size: 12px; 
  color: var(--gray-400); 
  margin-top: 2px;
}}
.table-container {{ 
  background: white; 
  border: 1px solid var(--gray-200); 
  overflow: hidden;
}}
table {{ 
  width: 100%; 
  border-collapse: collapse; 
  font-size: 14px;
}}
th, td {{ 
  padding: 12px 16px; 
  text-align: left; 
  border-bottom: 1px solid var(--gray-200);
}}
th {{ 
  background: var(--gray-100); 
  font-weight: 500; 
  font-size: 12px; 
  color: var(--gray-600); 
  text-transform: uppercase; 
  letter-spacing: 0.05em;
}}
th.month {{ width: 80px; text-align: center; }}
td.month {{ 
  text-align: center; 
  font-weight: 600; 
  width: 80px;
}}
td.check {{ 
  text-align: center; 
  font-size: 16px;
}}
td.check.present::before {{ 
  content: "●"; 
  color: var(--green-500);
}}
td.check.absent::before {{ 
  content: "○"; 
  color: var(--gray-200);
}}
tr:last-child td {{ border-bottom: none; }}
.footer {{ 
  margin-top: 24px; 
  font-size: 12px; 
  color: var(--gray-400); 
  text-align: right;
}}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>雅各組 出席統計</h1>
    <div class="period">2026年4月 - 6月</div>
  </div>

  <div class="summary">
    <div class="card">
      <div class="value">{total_apr}</div>
      <div class="label">4月出席人數</div>
      <div class="rate">/ {total_people} ({round(total_apr/total_people*100)}%)</div>
    </div>
    <div class="card">
      <div class="value">{total_may}</div>
      <div class="label">5月出席人數</div>
      <div class="rate">/ {total_people} ({round(total_may/total_people*100)}%)</div>
    </div>
    <div class="card">
      <div class="value">{total_jun}</div>
      <div class="label">6月出席人數</div>
      <div class="rate">/ {total_people} ({round(total_jun/total_people*100)}%)</div>
    </div>
  </div>

  <div class="table-container">
    <table>
      <thead>
        <tr>
          <th width="50">#</th>
          <th>姓名</th>
          <th class="month">4月</th>
          <th class="month">5月</th>
          <th class="month">6月</th>
          <th width="60">總次</th>
        </tr>
      </thead>
      <tbody>
{''.join(tbody_rows)}
      </tbody>
    </table>
  </div>

  <div class="footer">雅各組 出席表 Dashboard | 2026年7月1日</div>
</div>
</body>
</html>'''

HTML_PATH.write_text(html, encoding="utf-8")
print(f"Generated {HTML_PATH}")