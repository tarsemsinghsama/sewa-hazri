import base64, io, json, os
from openpyxl import load_workbook
import requests

ONEDRIVE_URL = "https://1drv.ms/x/c/0a8fa620daf4b126/IQBGkzekdy3FR5DLok9EPZgJAZGsCcEmBoMFxZ0Q8NH6FZo"
OUT = "data.json"

def share_token(url: str) -> str:
    raw = base64.b64encode(url.encode("utf-8")).decode("ascii")
    return raw.rstrip("=").replace("/", "_").replace("+", "-")

def download_xlsx():
    token = share_token(ONEDRIVE_URL)
    api = f"https://api.onedrive.com/v1.0/shares/u!{token}/root/content"
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(api, headers=headers, timeout=60, allow_redirects=True)
    r.raise_for_status()
    data = r.content
    if data[:2] != b"PK":
        # Fallback used by some OneDrive share links.
        r2 = requests.get(ONEDRIVE_URL + "?download=1", headers=headers, timeout=60, allow_redirects=True)
        r2.raise_for_status()
        data = r2.content
    if data[:2] != b"PK":
        raise RuntimeError("OneDrive did not return an XLSX file. Make sure the share link allows anyone with the link to view/download.")
    return data

def clean(v):
    return 0 if v is None or v == "" else v

def main():
    xlsx = download_xlsx()
    wb = load_workbook(io.BytesIO(xlsx), data_only=True, read_only=True)
    if "FINAL" not in wb.sheetnames:
        raise RuntimeError("FINAL sheet not found in workbook")
    ws = wb["FINAL"]
    records = []
    for row in range(5, ws.max_row + 1):
        badge = ws.cell(row, 2).value
        if badge is None or str(badge).strip() == "":
            continue
        records.append({
            "badge": str(badge).strip().upper(),
            "name": "" if ws.cell(row, 3).value is None else str(ws.cell(row, 3).value).strip(),
            "parent": "" if ws.cell(row, 4).value is None else str(ws.cell(row, 4).value).strip(),
            "day": clean(ws.cell(row, 8).value),
            "night": clean(ws.cell(row, 9).value),
            "saturday": clean(ws.cell(row, 10).value),
            "sunday": clean(ws.cell(row, 11).value),
            "lipai": clean(ws.cell(row, 12).value),
            "beas": clean(ws.cell(row, 13).value),
        })
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    print(f"Wrote {len(records)} records to {OUT}")

if __name__ == "__main__":
    main()
