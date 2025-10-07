import csv
import re
from pathlib import Path

# CONFIG
ETAPAS_DIR = "C:/Users/SauliusSteponavičius/ETP/ETP - ! Objektai/2024-72-XX_330kV Marijampolė/04_VE/!Viešinimas_SLD/!INFO/02_Patašinės sen._9 vnt"
ETAPAS_FILENAME = "aggregated_output.csv"
INPUT_FILE = Path(ETAPAS_DIR) / ETAPAS_FILENAME

TERMINAL_FILE = "real_addresses_terminal.txt"  # Path to your saved terminal output

# 1. Parse terminal output to build a lookup dictionary
address_map = {}
with open(TERMINAL_FILE, encoding="utf-8") as f:
    lines = f.readlines()

current_person = None
for line in lines:
    # Match: Making request for: VARDENIS PAVARDENIS (1999-09-09)
    m = re.match(r"Making request for: (.+) (.+) \((\d{4}-\d{2}-\d{2})\)", line.strip())
    if m:
        vardas, pavarde, gim_data = m.group(1), m.group(2), m.group(3)
        current_person = (vardas, pavarde, gim_data)
    # Match: Extracted for VARDENIS PAVARDENIS: Marijampolė, Marijampolės g. 99, LT-99999
    m2 = re.match(r"Extracted for (.+) (.+): (.+)", line.strip())
    if m2 and current_person:
        address_full = m2.group(3)
        # Split address and postal code
        addr_match = re.match(r"(.*),\s*(LT-\d{5})$", address_full)
        if addr_match:
            address = addr_match.group(1).strip()
            postal_code = addr_match.group(2).strip()
        else:
            address = address_full.strip()
            postal_code = ""
        address_map[current_person] = (address, postal_code)
        current_person = None

print(f"Loaded {len(address_map)} real addresses from terminal output.")

# 2. Update aggregated_output.csv
with open(INPUT_FILE, newline="", encoding="utf-8-sig") as csvfile:
    rows = list(csv.reader(csvfile))

header = rows[0]
updated_rows = [header]

for row in rows[1:]:
    if len(row) < 12:
        updated_rows.append(row)
        continue
    vardas = row[5]
    pavarde = row[6]
    gim_data = row[7]
    tipas = row[8]
    if tipas.lower() == "fizinis":
        key = (vardas, pavarde, gim_data)
        if key in address_map:
            address, postal_code = address_map[key]
            # Ensure row is long enough
            while len(row) < 12:
                row.append("")
            if len(row) >= 14:
                row[12] = address
                row[13] = postal_code
            else:
                row = row[:12] + [address, postal_code]
    updated_rows.append(row)

with open(INPUT_FILE, "w", newline="", encoding="utf-8-sig") as outcsv:
    writer = csv.writer(outcsv)
    writer.writerows(updated_rows)

print("aggregated_output.csv updated with real addresses from terminal output.")