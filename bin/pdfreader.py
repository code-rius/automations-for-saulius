from pathlib import Path
import pdfplumber
import csv
import re

# Folder containing PDFs = same folder as this script
folder_path = Path(__file__).parent

output_csv = "output.csv"

patterns = {
    "registro_nr": re.compile(r"Registro Nr\.:?\s*([^\n]+)"),
    "unikalus_nr": re.compile(r"Unikalus daikto numeris:?\s*([^\n]+)"),
    "pavadinimas": re.compile(r"pavadinimas:?\s*([^\n]+)"),
    "role": re.compile(r"(Savininkas|Nuomininkas|Patikėtinis):?\s*([^\n]+)")
}

rows = []

for pdf_file in folder_path.glob("*.pdf"):
    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    lines = full_text.splitlines()

    registro_nr = (m.group(1).strip() if (m := re.search(patterns["registro_nr"], full_text)) else "")

    adresas = ""
    for i, line in enumerate(lines):
        if "Žemės sklypas" in line:
            if i + 1 < len(lines):
                adresas = lines[i + 1].strip()
            break

    unikalus_nr = (m.group(1).strip() if (m := re.search(patterns["unikalus_nr"], full_text)) else "")
    pavadinimas = (m.group(1).strip() if (m := re.search(patterns["pavadinimas"], full_text)) else "")

    for match in re.finditer(patterns["role"], full_text):
        role = match.group(1).strip()
        name_field = match.group(2).strip()

        if ", gim." in name_field.lower():
            entry_type = "fizinis"
        elif ", a.k." in name_field.lower():
            entry_type = "jurinis"
        else:
            entry_type = ""

        split_match = re.split(r",\s*(?:gim\.|a\.k\.)\s*", name_field, maxsplit=1, flags=re.IGNORECASE)
        if len(split_match) == 2:
            name_clean = split_match[0].strip()
            id_or_date = split_match[1].strip()
        else:
            name_clean = name_field
            id_or_date = ""

        if entry_type == "fizinis":
            parts = name_clean.split(maxsplit=1)
            first_name = parts[0]
            surname = parts[1] if len(parts) > 1 else ""
        else:
            first_name = name_clean
            surname = ""

        rows.append((registro_nr, adresas, unikalus_nr, pavadinimas, role, first_name, surname, id_or_date, entry_type))

rows = sorted(set(rows))

with open(output_csv, "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerows(rows)

print(f"Saved {len(rows)} unique rows to {output_csv}")
