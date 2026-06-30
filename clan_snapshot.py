import os
import urllib.request
from datetime import datetime, timezone, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
import json
import time
import sys

API_URL = "https://playninjarift.com/api/detail_clan_website.php?clan_id=2527"
TARGET_TZ = timezone(timedelta(hours=8))
EXCEL_FILE = "clan_2527.xlsx"
CLAN_ID = 2527

def fetch_clan():
    req = urllib.request.Request(API_URL, headers={"User-Agent": "clan-snapshot/1.0"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode())

def get_previous_sheet_names(wb):
    names = [s.title for s in wb.worksheets]
    names.sort()
    return names

def load_prev_from_xlsx(filename, before_date):
    prev_data = []
    if not os.path.exists(filename):
        return prev_data
    wb = load_workbook(filename)
    prev_names = [s.title for s in wb.worksheets]
    prev_names.sort()
    prev_sheet_name = None
    for n in reversed(prev_names):
        if n < before_date:
            prev_sheet_name = n
            break
    if prev_sheet_name and prev_sheet_name in wb.sheetnames:
        ps = wb[prev_sheet_name]
        for row in ps.iter_rows(min_row=4, max_col=2, values_only=True):
            if row[0] and row[1] is not None:
                prev_data.append({"character_name": row[0], "member_reputation": int(row[1])})
    return prev_data

def compute_diff(members, prev_data):
    prev_map = {m["character_name"]: m["member_reputation"] for m in prev_data}
    result = []
    for m in members:
        name = m["character_name"]
        reps = m["member_reputation"]
        if name in prev_map:
            diff = reps - prev_map[name]
            diff_str = f"+{diff}" if diff > 0 else str(diff)
        else:
            diff_str = "N/A"
        result.append((name, reps, diff_str))
    return result

def write_sheet(ws, data, prev_data, now):
    rows = compute_diff(data["members"], prev_data)
    ws.title = now.strftime("%Y-%m-%d")

    names = [r[0] for r in rows]
    reps = [str(r[1]) for r in rows]
    diffs = [r[2] for r in rows]

    max_name = max((len(n) for n in names), default=10)
    max_reps = max((len(r) for r in reps), default=4)
    max_diff = max((len(d) for d in diffs), default=3)

    ws.column_dimensions["A"].width = max(max_name + 5, 10)
    ws.column_dimensions["B"].width = max(max_reps + 3, 8)
    ws.column_dimensions["C"].width = max(max_diff + 3, 14)

    clan_name = data.get("clan_name", "Unknown")
    ws.merge_cells("A1:C1")
    ws["A1"] = f"Clan: {clan_name} ({CLAN_ID})"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:C2")
    ws["A2"] = f"Timestamp: {now.strftime('%Y-%m-%d %H:%M:%S')}"
    ws["A2"].font = Font(bold=True, size=11)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Name", "Reps", "Reps Difference"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, (name, reps_val, diff_val) in enumerate(rows, 4):
        ws.cell(row=row_idx, column=1, value=name).alignment = Alignment(vertical="center")
        ws.cell(row=row_idx, column=2, value=reps_val).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_idx, column=3, value=diff_val).alignment = Alignment(horizontal="center", vertical="center")

def save_xlsx(data, prev_data, now):
    sheet_name = now.strftime("%Y-%m-%d")

    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
    else:
        wb = Workbook()

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    ws = wb.create_sheet(title=sheet_name)
    write_sheet(ws, data, prev_data, now)

    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb["Sheet"]

    wb.save(EXCEL_FILE)
    print(f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] Saved sheet '{sheet_name}' to {EXCEL_FILE}")

def save_html(data, prev_data, now, all_dates):
    rows = compute_diff(data["members"], prev_data)
    clan_name = data.get("clan_name", "Unknown")
    date_str = now.strftime("%Y-%m-%d")
    ts_str = now.strftime("%Y-%m-%d %H:%M:%S")

    def css_color(diff_str):
        if diff_str.startswith("+"):
            return "#4caf50"
        elif diff_str == "N/A":
            return "#888"
        elif diff_str.startswith("-"):
            return "#f44336"
        return "#888"

    archive_links = "".join(
        f'<a href="{d}.html" class="{"active" if d == date_str else ""}">{d}</a>'
        for d in sorted(all_dates, reverse=True)
    )

    table_rows = "".join(
        f"""<tr>
          <td>{name}</td>
          <td class="num">{reps}</td>
          <td class="num" style="color:{css_color(diff_str)}">{diff_str}</td>
        </tr>"""
        for name, reps, diff_str in rows
    )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{clan_name} - Clan Snapshot</title>
<style>
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{
    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
    background: #0d0d0d;
    color: #e0e0e0;
    min-height: 100vh;
    display: flex;
    justify-content: center;
    padding: 40px 16px;
  }}
  .container {{ max-width: 900px; width: 100%; }}
  .header {{
    text-align: center;
    padding: 28px 20px;
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    border-radius: 12px 12px 0 0;
    border-bottom: 3px solid #e94560;
  }}
  .header h1 {{ font-size: 22px; color: #fff; margin-bottom: 6px; }}
  .header p {{ font-size: 14px; color: #aaa; }}
  .archive {{
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    justify-content: center;
    padding: 12px 20px;
    background: #161616;
    border-bottom: 1px solid #222;
  }}
  .archive a {{
    color: #888;
    text-decoration: none;
    font-size: 13px;
    padding: 4px 10px;
    border-radius: 4px;
    transition: 0.2s;
  }}
  .archive a:hover {{ background: #222; color: #fff; }}
  .archive a.active {{ color: #e94560; font-weight: 700; background: #1a1a2e; }}
  table {{
    width: 100%;
    border-collapse: collapse;
    background: #111;
    border-radius: 0 0 12px 12px;
    overflow: hidden;
  }}
  th {{
    background: #1a1a2e;
    padding: 12px 16px;
    text-align: left;
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    color: #e94560;
  }}
  th:nth-child(2), th:nth-child(3) {{ text-align: center; }}
  td {{
    padding: 10px 16px;
    border-bottom: 1px solid #1a1a2e;
    font-size: 14px;
  }}
  tr:nth-child(even) td {{ background: #0d0d0d; }}
  tr:hover td {{ background: #1a1a2e; }}
  td.num {{ text-align: center; font-variant-numeric: tabular-nums; }}
  .footer {{
    text-align: center;
    padding: 20px;
    color: #555;
    font-size: 12px;
  }}
  .footer a {{ color: #e94560; text-decoration: none; }}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>{clan_name}</h1>
    <p>Clan ID: {CLAN_ID} &mdash; Snapshot: {ts_str}</p>
  </div>
  {f'<div class="archive">{archive_links}</div>' if archive_links else ""}
  <table>
    <thead>
      <tr><th>Name</th><th>Reps</th><th>Reps Difference</th></tr>
    </thead>
    <tbody>{table_rows}</tbody>
  </table>
  <div class="footer">
    Auto-updated daily at 13:01 GMT+8 via
    <a href="https://github.com/nixervo/Clairvoyant-Reps" target="_blank">GitHub Actions</a>
  </div>
</div>
</body>
</html>"""

    index_path = "index.html"
    with open(index_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[{ts_str}] Saved {index_path}")

    archive_path = f"{date_str}.html"
    with open(archive_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[{ts_str}] Saved {archive_path}")

def save_snapshot(data):
    now = datetime.now(TARGET_TZ)
    sheet_name = now.strftime("%Y-%m-%d")

    prev_data = load_prev_from_xlsx(EXCEL_FILE, sheet_name)

    save_xlsx(data, prev_data, now)

    existing_html = [f.replace(".html", "") for f in os.listdir(".") if f.endswith(".html") and f[:4].isdigit() and f != "index.html"]
    all_dates = set(existing_html)
    all_dates.add(sheet_name)
    save_html(data, prev_data, now, sorted(all_dates))

def run():
    print("Clan snapshot daemon started. Will run at 13:01 GMT+8 daily.")
    while True:
        now = datetime.now(TARGET_TZ)
        target = now.replace(hour=13, minute=1, second=0, microsecond=0)
        if now >= target:
            target += timedelta(days=1)
        sleep_sec = (target - now).total_seconds()
        time.sleep(sleep_sec)
        try:
            data = fetch_clan()
            save_snapshot(data)
        except Exception as e:
            print(f"[{datetime.now(TARGET_TZ).strftime('%Y-%m-%d %H:%M:%S')}] ERROR: {e}")

def fetch_once():
    data = fetch_clan()
    save_snapshot(data)

if __name__ == "__main__":
    if "--once" in sys.argv:
        fetch_once()
    else:
        run()
