import csv
from pathlib import Path

csv_path = Path(__file__).parent.parent / "resources" / "nameinfo.csv"
output_folder = Path(__file__).parent.parent / "out"
output_folder.mkdir(exist_ok=True)

with open(csv_path, encoding="utf-8") as f:
    reader = csv.reader(f, delimiter=";")
    for row in reader:
        if len(row) < 3:
            continue
        filename = f"{row[0]}.txt"
        txt_path = output_folder / filename
        print(f"Writing to {txt_path}: {row}")  # Debug print
        with open(txt_path, "w", encoding="utf-8") as txt_file:
            txt_file.write(f"BENDRAS_NR={row[0]}\n")
            txt_file.write(f"PROJEKTO_NR={row[1]}\n")
            txt_file.write(f"PAVADINIMAS={row[2]}\n")