"""
update_scores.py  — อัปเดต สรุปคะแนน.html จาก สะสมคะแนน.xlsx อัตโนมัติ
"""
import re, json, sys, os
import openpyxl
from datetime import datetime, timedelta

BASE  = os.path.dirname(os.path.abspath(__file__))
XLSX  = os.path.join(BASE, "สะสมคะแนน.xlsx")
HTML  = os.path.join(BASE, "สรุปคะแนน.html")

# ── 1. อ่าน Excel ───────────────────────────────────────────
print("📖 กำลังอ่าน", XLSX)
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb.active

rows = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: continue
    dv, name, menu = row[0], row[1], row[2]
    hc, fd, lb, ef, et, cl = row[3], row[4], row[5], row[6], row[7], row[9]
    if not dv or not name or not isinstance(dv, datetime): continue
    name = name.strip()
    cl   = int(cl) if isinstance(cl, (int, float)) else 0
    rows.append({
        "date":        dv.strftime("%Y-%m-%d"),
        "name":        name,
        "menu":        (menu or "").strip(),
        "homeChef":    int(hc) if hc else 0,
        "foodie":      int(fd) if fd else 0,
        "lunchbox":    int(lb) if lb else 0,
        "extraFriend": int(ef) if ef else 0,
        "extraTask":   int(et) if isinstance(et, (int, float)) else 0,
        "countLike":   cl,
    })

if not rows:
    print("❌ ไม่พบข้อมูลใน Excel"); sys.exit(1)

dates    = sorted({r["date"] for r in rows})
min_date = dates[0]; max_date = dates[-1]
print(f"✅ พบข้อมูล {len(rows)} rows  ({min_date} → {max_date})")

# ── 2. สร้าง JS block ───────────────────────────────────────
lines = ["  " + json.dumps(r, ensure_ascii=False) + "," for r in rows]
lines[-1] = lines[-1].rstrip(",")
raw_js = "const RAW = [\n" + "\n".join(lines) + "\n];"

# auto WEEKS ทุก 7 วัน
dt_min = datetime.strptime(min_date, "%Y-%m-%d")
dt_max = datetime.strptime(max_date, "%Y-%m-%d")
weeks = []; wk = 1; start = dt_min
while start <= dt_max:
    end = min(start + timedelta(days=6), dt_max)
    weeks.append({"label": f"Week {wk}", "from": start.strftime("%Y-%m-%d"), "to": end.strftime("%Y-%m-%d")})
    wk += 1; start = end + timedelta(days=1)
weeks_all = [{"label": "รวมทั้งหมด", "from": min_date, "to": max_date}] + weeks
wlines = ['  { "label": "' + w["label"] + '", "from": "' + w["from"] + '", "to": "' + w["to"] + '" },' for w in weeks_all]
wlines[-1] = wlines[-1].rstrip(",")
weeks_js = "const WEEKS = [\n" + "\n".join(wlines) + "\n];"

# ── 3. แทนที่ใน HTML ────────────────────────────────────────
with open(HTML, encoding="utf-8") as f:
    html = f.read()

html_new = re.sub(r"const RAW = \[[\s\S]*?\];",   raw_js,   html)
html_new = re.sub(r"const WEEKS = \[[\s\S]*?\];", weeks_js, html_new)

# ตรวจว่า pattern เจอหรือไม่ (ไม่ดูจาก content เพราะอาจเหมือนกัน)
if not re.search(r"const RAW = \[", html_new):
    print("⚠️  ไม่พบ RAW block ใน HTML"); sys.exit(1)

with open(HTML, "w", encoding="utf-8") as f:
    f.write(html_new)

print(f"🎉 อัปเดตสำเร็จ!")
for w in weeks_all:
    print(f"   {w['label']:15s}  {w['from']} → {w['to']}")
