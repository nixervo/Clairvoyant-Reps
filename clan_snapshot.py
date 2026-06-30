import os
import urllib.request
from datetime import datetime, timezone, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
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

def compute_diff(data, prev_data):
    prev_map = {m["character_name"]: m["member_reputation"] for m in prev_data}
    today_map = {m["character_name"]: m["member_reputation"] for m in data["members"]}
    result = []
    for m in data["members"]:
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
    rows = compute_diff(data, prev_data)
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

def save_snapshot(data):
    now = datetime.now(TARGET_TZ)
    sheet_name = now.strftime("%Y-%m-%d")

    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        prev_names = get_previous_sheet_names(wb)
        prev_sheet_name = None
        for n in reversed(prev_names):
            if n < sheet_name:
                prev_sheet_name = n
                break
        prev_data = []
        if prev_sheet_name and prev_sheet_name in wb.sheetnames:
            ps = wb[prev_sheet_name]
            for row in ps.iter_rows(min_row=4, max_col=2, values_only=True):
                if row[0] and row[1] is not None:
                    prev_data.append({"character_name": row[0], "member_reputation": int(row[1])})
    else:
        wb = Workbook()
        prev_data = []
        prev_sheet_name = None

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    ws = wb.create_sheet(title=sheet_name)
    write_sheet(ws, data, prev_data, now)

    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb["Sheet"]

    wb.save(EXCEL_FILE)
    print(f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] Saved sheet '{sheet_name}' to {EXCEL_FILE}")

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
