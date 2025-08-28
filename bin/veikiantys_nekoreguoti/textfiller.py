import os
import csv
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

CSV_DIRECTORY = os.environ.get("CSV_DIRECTORY")
TXT_DIRECTORY = os.environ.get("TXT_DIRECTORY")

if not CSV_DIRECTORY or not TXT_DIRECTORY:
    print("Please set CSV_DIRECTORY and TXT_DIRECTORY in your .env file.")
    exit(1)

csv_folder = Path(CSV_DIRECTORY)
txt_folder = Path(TXT_DIRECTORY)

# Surask pirmą .csv failą
csv_files = list(csv_folder.glob("*.csv"))
if not csv_files:
    print("No .csv files found in the directory.")
    exit(1)
csv_path = csv_files[0]

# Surask pirmą .txt failą
txt_files = list(txt_folder.glob("*.txt"))
if not txt_files:
    print("No .txt files found in the directory.")
    exit(1)
txt_path = txt_files[0]

# Ištrauk reikiamas reikšmes iš txt failo
info = {"BENDRAS_NR": "", "PROJEKTO_NR": "", "PAVADINIMAS": ""}
with open(txt_path, "r", encoding="utf-8-sig") as f:
    for line in f:
        for key in info.keys():
            if line.startswith(f"{key}="):
                info[key] = line.strip().split("=", 1)[1]

# Perskaityk csv failą ir pašalink BOM iš pirmos eilutės, jei yra
with open(csv_path, "r", newline="", encoding="utf-8-sig") as csvfile:
    reader = list(csv.reader(csvfile))
    # Pašalink BOM iš pirmo stulpelio kiekvienos eilutės, jei yra
    for row in reader:
        if row and row[0]:
            row[0] = row[0].lstrip('\ufeff')

# Pridėk reikiamą informaciją prie kiekvienos csv eilutės
updated_rows = []
for row in reader:
    updated_rows.append(row + [info["BENDRAS_NR"], info["PROJEKTO_NR"], info["PAVADINIMAS"]])

# Įrašyk atnaujintas eilutes atgal į csv failą be BOM
with open(csv_path, "w", newline="", encoding="utf-8") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerows(updated_rows)

print(f"Pridėta informacija iš {txt_path} prie kiekvienos eilutės faile: {csv_path}")