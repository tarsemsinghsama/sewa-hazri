import io, json, requests
from pathlib import Path
from openpyxl import load_workbook

# आपकी OneDrive Excel sharing link
ONEDRIVE_URL = "https://1drv.ms/x/c/0a8fa620daf4b126/IQBGkzekdy3FR5DLok9EPZgJAd-vX91T1slxFO89cAGYBnA"

ROOT = Path(__file__).resolve().parent
out = ROOT / "data.json"

# OneDrive से नवीनतम Excel डाउनलोड
resp = requests.get(ONEDRIVE_URL, params={"download":"1"}, allow_redirects=True, timeout=60)
resp.raise_for_status()
content = resp.content

# FINAL sheet से B,C,D,H,I,J,K,L,M
wb = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
if "FINAL" not in wb.sheetnames:
    raise RuntimeError("Excel में FINAL sheet नहीं मिली।")

ws = wb["FINAL"]
records = []

def clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v

for row in range(5, ws.max_row + 1):
    badge = clean(ws.cell(row, 2).value)  # B
    if str(badge).strip() == "":
        continue
    records.append({
        "badge": badge,
        "name": clean(ws.cell(row, 3).value),      # C
        "father": clean(ws.cell(row, 4).value),    # D
        "day": clean(ws.cell(row, 8).value),       # H
        "night": clean(ws.cell(row, 9).value),     # I
        "saturday": clean(ws.cell(row,10).value),  # J
        "sunday": clean(ws.cell(row,11).value),    # K
        "lipai": clean(ws.cell(row,12).value),     # L
        "beas": clean(ws.cell(row,13).value),      # M
    })

new = json.dumps(records, ensure_ascii=False, indent=2)

old = out.read_text(encoding="utf-8") if out.exists() else ""
if old != new:
    out.write_text(new, encoding="utf-8")
    print(f"Updated data.json: {len(records)} records")
else:
    print("No attendance change.")
