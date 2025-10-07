import os
import csv
import re
import time
import requests
from bs4 import BeautifulSoup
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ===== Config =====
ETAPAS_DIR = os.environ.get("DIR_ETAPAS")
if not ETAPAS_DIR:
    raise RuntimeError("Environment variable DIR_ETAPAS is not set.")

ETAPAS_FILENAME = os.environ.get("ETAPAS_OUTPUT_FILE_NAME", "aggregated_output.csv")
INPUT_FILE = Path(ETAPAS_DIR) / ETAPAS_FILENAME

MOCK_EXTRACTION = os.environ.get("MOCK_EXTRACTION", "FALSE").upper() == "TRUE"

URL = "https://mgvdisisorinis.registrucentras.lt/ivn/paieska-pagal-asmeni"

COOKIE = os.environ.get("RC_COOKIE", "")
if not COOKIE and not MOCK_EXTRACTION:
    raise RuntimeError("Environment variable RC_COOKIE is not set.")

HEADERS = {
    "Cookie": COOKIE,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Content-Type": "application/x-www-form-urlencoded"
}

# ===== Extract Address =====
def extract_address(html):
    """Extract address from HTML response."""
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
    """Build request payload for Registru Centras API."""
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

# ===== Mock Address =====
def mock_address(vardas, pavarde):
    """Generate a mock address for testing."""
    return f"Mock adresas {vardas} {pavarde}", "LT-12345"

# ===== Main =====
def main():
    if not Path(INPUT_FILE).exists():
        print(f"Error: File {INPUT_FILE} not found.")
        exit(1)

    print(f"Processing file: {INPUT_FILE}")
    
    if MOCK_EXTRACTION:
        print("MOCK MODE: No actual API calls will be made to Registru Centras.")
        print("This will only log what would be extracted.")

    # First, identify all unique individuals that need addresses
    unique_individuals = {}  # (vardas, pavarde, gim_data) -> row indices
    updated_rows = []
    
    with open(INPUT_FILE, newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.reader(csvfile)
        rows = list(reader)
        
    # Keep the header
    header = rows[0]
    updated_rows.append(header)
    
    # Find all unique individuals and track their row indices
    for i, row in enumerate(rows[1:], 1):  # Skip header row
        if len(row) < 12:  # Not enough columns
            updated_rows.append(row)
            continue
        
        vardas = row[5]
        pavarde = row[6]
        gim_data = row[7]
        tipas = row[8]
        
        if tipas.lower() == "fizinis":
            person_key = (vardas, pavarde, gim_data)
            if person_key not in unique_individuals:
                unique_individuals[person_key] = []
            unique_individuals[person_key].append(i)
    
    print(f"Found {len(unique_individuals)} unique individuals requiring address lookup")
    
    # Now process each unique individual once
    cache = {}
    request_count = 0
    
    for person_key, row_indices in unique_individuals.items():
        vardas, pavarde, gim_data = person_key
        
        if MOCK_EXTRACTION:
            address, postal_code = mock_address(vardas, pavarde)
            print(f"MOCK: Would extract address for {vardas} {pavarde}, {gim_data}")
            print(f"MOCK: Would return: {address}, {postal_code}")
        else:
            registro_nr = rows[row_indices[0]][0]  # Use registro_nr from first occurrence
            
            payload = build_payload(registro_nr, vardas, pavarde, gim_data)
            
            # Add detailed debug information before the request
            print(f"\nMaking request for: {vardas} {pavarde} ({gim_data})")
            print(f"URL: {URL}")
            print(f"Payload: {payload}")
            
            response = requests.post(URL, headers=HEADERS, data=payload)
            request_count += 1
            
            print(f"Response status: {response.status_code}")
            
            if response.status_code == 200:
                address, postal_code = extract_address(response.text)
                print(f"Extracted for {vardas} {pavarde}: {address}, {postal_code}")
            else:
                address, postal_code = "", ""
                print(f"Failed to extract for {vardas} {pavarde}. Status code: {response.status_code}")
            time.sleep(2)  # Prevent rate limiting
        
        cache[person_key] = (address, postal_code)
    
    # Now apply cached results to all rows
    for i, row in enumerate(rows[1:], 1):  # Skip header row
        if len(row) < 12:  # Not enough columns
            continue
            
        vardas = row[5]
        pavarde = row[6]
        gim_data = row[7]
        tipas = row[8]
        
        if tipas.lower() == "fizinis":
            person_key = (vardas, pavarde, gim_data)
            if person_key in cache:
                address, postal_code = cache[person_key]
                
                # Ensure row is long enough
                while len(row) < 12:
                    row.append("")
                
                # Always overwrite address and postal code
                if len(row) >= 14:
                    row[12] = address
                    row[13] = postal_code
                else:
                    row = row[:12] + [address, postal_code]

        updated_rows.append(row)

    # Modified part: Always write to file, even in MOCK mode
    with open(INPUT_FILE, "w", newline="", encoding="utf-8-sig") as outcsv:
        writer = csv.writer(outcsv)
        writer.writerows(updated_rows)

    print(f"File updated: {INPUT_FILE}")
    if MOCK_EXTRACTION:
        print("MOCK MODE: File was updated with mock address values for testing.")
    else:
        print(f"Requests made to website: {request_count}")
    print(f"Unique individuals processed: {len(cache)}")
    print(f"Total individuals in file: {len(unique_individuals)}")

if __name__ == "__main__":
    main()