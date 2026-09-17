import base64
import io
import json
import os
import tempfile
import requests
from openpyxl import load_workbook

# Public OneDrive Excel link supplied for this project.
ONEDRIVE_URL = "https://1drv.ms/x/c/0a8fa620daf4b126/IQBGkzekdy3FR5DLok9EPZgJAZGsCcEmBoMFxZ0Q8NH6FZo"

OUTPUT_FILE = "data.json"
SHEET_NAME = "FINAL"

def get_excel_bytes(url: str) -> bytes:
    headers = {"User-Agent": "Mozilla/5.0"}
    urls = []

    # First try the shared link itself.
    urls.append(url)

    # Then try Microsoft's public Shares API. This works with public OneDrive share links.
    token = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")
    urls.append(f"https://api.onedrive.com/v1.0/shares/u!{token}/root/content")

    # Finally try the common download query.
    sep = "&" if "?" in url else "?"
    urls.append(url + sep + "download=1")

    last_error = None
    for candidate in urls:
        try:
            r = requests.get(candidate, headers=headers, allow_redirects=True, timeout=60)
            r.raise_for_status()
            content = r.content
            # XLSX files are ZIP containers and normally start with PK.
            if content[:2] == b"PK":
                return content
            last_error = f"Not an XLSX response from {candidate} (content-type={r.headers.get('content-type')})"
        except Exception as exc:
            last_error = str(exc)

    raise RuntimeError(
        "OneDrive Excel डाउनलोड नहीं हो सकी। Share link को 'Anyone with the link can view' "
        "पर रखें और फिर GitHub Actions दोबारा चलाएँ. Last error: " + str(last_error)
    )

def clean_value(v):
    if v is None:
        return 0
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v

def build_records(xlsx_bytes: bytes):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(xlsx_bytes)
        temp_path = tmp.name

    try:
        wb = load_workbook(temp_path, data_only=True, read_only=True)
        if SHEET_NAME not in wb.sheetnames:
            raise RuntimeError(f"Excel में '{SHEET_NAME}' sheet नहीं मिली।")

        ws = wb[SHEET_NAME]
        records = []

        # FINAL sheet में B,C,D,H,I,J,K,L,M से ही database बनाया जाता है.
        # Header/blank rows अपने-आप skip हो जाती हैं.
        for row in ws.iter_rows(min_row=1, values_only=True):
            b = row[1] if len(row) > 1 else None
            c = row[2] if len(row) > 2 else None
            d = row[3] if len(row) > 3 else None

            if not b or not c:
                continue

            badge = str(b).strip()
            if badge in ("बैज नं0", "बैज नं"):
                continue

            # Excel columns: B=2, C=3, D=4, H=8, I=9, J=10, K=11, L=12, M=13
            def col(idx):
                return clean_value(row[idx - 1]) if len(row) >= idx else 0

            records.append({
                "badge": badge,
                "name": str(c).strip(),
                "parent": str(d or "").strip(),
                "day": col(8),
                "night": col(9),
                "saturday": col(10),
                "sunday": col(11),
                "lipai": col(12),
                "beas": col(13)
            })

        wb.close()
        return records
    finally:
        try:
            os.remove(temp_path)
        except OSError:
            pass

def main():
    xlsx = get_excel_bytes(ONEDRIVE_URL)
    records = build_records(xlsx)

    if not records:
        raise RuntimeError("FINAL sheet से कोई valid record नहीं मिला।")

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    print(f"Updated {OUTPUT_FILE}: {len(records)} records")

if __name__ == "__main__":
    main()
