import base64, io, json, requests, openpyxl

# OneDrive sharing link supplied for the Excel workbook
ONEDRIVE_URL = "https://1drv.ms/x/c/0a8fa620daf4b126/IQBGkzekdy3FR5DLok9EPZgJAZGsCcEmBoMFxZ0Q8NH6FZo"
SHEET = "FINAL"

def download_excel():
    token = base64.urlsafe_b64encode(ONEDRIVE_URL.encode()).decode().rstrip("=")
    urls = [
        f"https://api.onedrive.com/v1.0/shares/u!{token}/root/content",
        ONEDRIVE_URL + ("&" if "?" in ONEDRIVE_URL else "?") + "download=1"
    ]
    for url in urls:
        try:
            r = requests.get(url, timeout=90, allow_redirects=True,
                             headers={"User-Agent":"Mozilla/5.0"})
            r.raise_for_status()
            if r.content[:2] == b"PK":
                return r.content
        except Exception:
            continue
    raise RuntimeError("OneDrive workbook download failed. Check that the share link is accessible without sign-in.")

def cv(v):
    if v is None: return ""
    if isinstance(v,float) and v.is_integer(): return int(v)
    return v

data = download_excel()
wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
if SHEET not in wb.sheetnames:
    raise RuntimeError("FINAL sheet not found: " + ", ".join(wb.sheetnames))
ws = wb[SHEET]

# FINAL has two data blocks: rows 5-115 and 134-179.
records=[]
for r in list(range(5,116)) + list(range(134,180)):
    badge=cv(ws.cell(r,2).value)
    if str(badge).strip():
        records.append({
            "badge":badge, "name":cv(ws.cell(r,3).value), "father":cv(ws.cell(r,4).value),
            "day":cv(ws.cell(r,8).value), "night":cv(ws.cell(r,9).value),
            "saturday":cv(ws.cell(r,10).value), "sunday":cv(ws.cell(r,11).value),
            "lipai":cv(ws.cell(r,12).value), "beas":cv(ws.cell(r,13).value)
        })

with open("data.json","w",encoding="utf-8") as f:
    json.dump({"records":records},f,ensure_ascii=False,indent=2)

print("FINAL records:",len(records))
if len(records) != 157:
    raise RuntimeError(f"Expected 157 records, got {len(records)}")
