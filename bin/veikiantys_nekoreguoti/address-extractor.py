import os
import csv
import re
import time
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()

# ===== Config =====
INPUT_DIR = os.environ.get("INPUT_DIRECTORY", "out")
INPUT_FILE = os.path.join(INPUT_DIR, "output.csv")
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
        "XML[uzk_parametrai][paieskos_tikslas][VALUE]": "10026",
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
    updated_rows = []
    cache = {}
    request_count = 0

    with open(INPUT_FILE, newline="", encoding="utf-8") as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            if len(row) < 8:
                updated_rows.append(row + ["", ""])
                continue

            pastaba = row[0]
            vardas = row[5]
            pavarde = row[6]
            gim_data = row[7]
            cache_key = (vardas, pavarde, gim_data)

            if cache_key in cache:
                address, postal_code = cache[cache_key]
            else:
                payload = build_payload(pastaba, vardas, pavarde, gim_data)
                response = requests.post(URL, headers=HEADERS, data=payload)
                request_count += 1
                if response.status_code == 200:
                    address, postal_code = extract_address(response.text)
                else:
                    address, postal_code = "", ""
                cache[cache_key] = (address, postal_code)
                time.sleep(3)

            updated_rows.append(row + [address, postal_code])

    with open(INPUT_FILE, "w", newline="", encoding="utf-8") as outcsv:
        writer = csv.writer(outcsv)
        writer.writerows(updated_rows)

    print(f"Failas papildytas: {INPUT_FILE}")
    print(f"Requests made to website: {request_count}")

if __name__ == "__main__":
    main()