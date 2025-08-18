import csv
import re
import time
import requests
import os
from bs4 import BeautifulSoup
from dotenv import load_dotenv  # <-- Add this import

load_dotenv()  # <-- Load environment variables from .env

# ===== Config =====
INPUT_FILE = "file.csv"
OUTPUT_FILE = "file_with_addresses.csv"
URL = "https://mgvdisisorinis.registrucentras.lt/ivn/paieska-pagal-asmeni"

COOKIE = os.environ.get("RC_COOKIE", "")
if not COOKIE:
    raise RuntimeError("Environment variable RC_COOKIE is not set.")

HEADERS = {
    "Cookie": COOKIE,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Content-Type": "application/x-www-form-urlencoded"
}

# ===== Extract Address =====
def extract_address(html):
    soup = BeautifulSoup(html, "html.parser")
    li = soup.find(string=re.compile(r"Deklaravo gyvenamąją vietą:"))
    if not li:
        return "", ""
    text = li.strip()
    m = re.search(r"Deklaravo gyvenamąją vietą:\s*\d{4}-\d{2}-\d{2}\s*(.+)", text)
    if not m:
        return "", ""
    
    full_addr = m.group(1).strip()
    # Split postal code if present
    postal_match = re.search(r"(.*),\s*(LT-\d{5})$", full_addr)
    if postal_match:
        address = postal_match.group(1).strip()
        postal_code = postal_match.group(2).strip()
        return address, postal_code
    else:
        return full_addr, ""

# ===== Payload Builder =====
def build_payload(pastaba, vardas, pavarde, gim_data):
    return {
        "XML[uzk_parametrai][paieskos_tikslas][VALUE]": "10018",
        "page_type": "perziura",
        "XML[uzk_parametrai][paieskos_pastaba][VALUE]": pastaba,
        "XML[uzk_parametrai][asm_kodas][VALUE]": "",
        "XML[uzk_parametrai][vardas][VALUE]": vardas.upper(),
        "XML[uzk_parametrai][pavarde][VALUE]": pavarde.upper(),
        "XML[uzk_parametrai][tiksli_gim_data][VALUE]": gim_data,
        "XML[uzk_parametrai][gim_metai_nuo][VALUE]": "",
        "XML[uzk_parametrai][gim_metai_iki][VALUE]": ""
    }

# ===== Main =====
def main():
    results = []
    with open(INPUT_FILE, newline="", encoding="utf-8") as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:  # No header skip
            if len(row) < 4:
                continue
            pastaba, vardas, pavarde, gim_data = row[:4]

            payload = build_payload(pastaba, vardas, pavarde, gim_data)
            response = requests.post(URL, headers=HEADERS, data=payload)
            
            if response.status_code == 200:
                address, postal_code = extract_address(response.text)
            else:
                address, postal_code = "", ""
            
            results.append([pastaba, vardas, pavarde, gim_data, address, postal_code])

            time.sleep(1)  # Delay between requests

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as outcsv:
        writer = csv.writer(outcsv)
        writer.writerows(results)

    print(f"Done. Saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
