from pathlib import Path
import pdfplumber
import re
import csv

folder_path = Path(__file__).parent.parent / "resources"
output_csv = Path(__file__).parent.parent / "out/output.csv"
info_txt = folder_path / "info.txt"

patterns = {
    "registro_nr": re.compile(r"Registro Nr\.:?\s*([^\n]+)"),
    "unikalus_nr": re.compile(r"Unikalus daikto numeris:?\s*([^\n]+)"),
    "kadastro_nr": re.compile(r"pavadinimas:?\s*([^\n]+)"),
    "panaudos_gavejas": re.compile(r"Panaudos gavėjas:?\s*([^\n]+)")
}

# Read info.txt values
info_values = ["", "", ""]
if info_txt.exists():
    with open(info_txt, encoding="utf-8") as f:
        lines = f.readlines()
        for i, line in enumerate(lines):
            if i < 3:
                parts = line.strip().split("=", 1)
                info_values[i] = parts[1] if len(parts) == 2 else ""

rows = []

def split_name(name_clean, entry_type):
    parts = name_clean.split()
    if entry_type == "fizinis":
        if len(parts) >= 3:
            first_name = " ".join(parts[:-1])
            surname = parts[-1]
        elif len(parts) == 2 and "-" in parts[1]:
            first_name = parts[0]
            surname = parts[1]
        elif len(parts) == 2:
            first_name = parts[0]
            surname = parts[1]
        else:
            first_name = name_clean
            surname = ""
    else:
        first_name = name_clean
        surname = ""
    return first_name, surname

def process_role_block(lines, start_keyword, role_label, registro_nr, adresas, unikalus_nr, kadastro_nr):
    for i, line in enumerate(lines):
        if re.match(fr"{start_keyword}:?\s*", line.strip(), re.IGNORECASE):
            first_line = line.strip()
            name_field = re.sub(fr"{start_keyword}:?\s*", "", first_line, flags=re.IGNORECASE).strip()
            entries_raw = [name_field]

            for j in range(i + 1, len(lines)):
                next_line = lines[j].strip()
                if re.match(r".+,\s*(gim\.|a\.k\.)", next_line, re.IGNORECASE):
                    entries_raw.append(next_line)
                else:
                    break

            for entry in entries_raw:
                if entry.lower() == "lietuvos respublika, a.k. 111105555".lower():
                    continue

                if ", gim." in entry.lower():
                    entry_type = "fizinis"
                elif ", a.k." in entry.lower():
                    entry_type = "juridinis"
                else:
                    entry_type = ""

                split_match = re.split(r",\s*(?:gim\.|a\.k\.)\s*", entry, maxsplit=1, flags=re.IGNORECASE)
                if len(split_match) == 2:
                    name_clean = split_match[0].strip()
                    id_or_date = split_match[1].strip()
                else:
                    name_clean = entry
                    id_or_date = ""

                first_name, surname = split_name(name_clean, entry_type)

                rows.append((
                    registro_nr, adresas, unikalus_nr, kadastro_nr, role_label,
                    first_name, surname, id_or_date, entry_type,
                    info_values[0], info_values[1], info_values[2]
                ))

for pdf_file in folder_path.glob("*.pdf"):
    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    lines = full_text.splitlines()

    registro_nr = (m.group(1).strip() if (m := re.search(patterns["registro_nr"], full_text)) else "")
    unikalus_nr_raw = (m.group(1).strip() if (m := re.search(patterns["unikalus_nr"], full_text)) else "")
    unikalus_nr = re.sub(r"Sudarymo data:[^\n]*", "", unikalus_nr_raw).strip()

    kadastro_raw = (m.group(1).strip() if (m := re.search(patterns["kadastro_nr"], full_text)) else "")
    kadastro_match = re.match(r"^\d+/\d+:\d+", kadastro_raw)
    kadastro_nr = kadastro_match.group(0) if kadastro_match else kadastro_raw

    adresas = ""
    for i, line in enumerate(lines):
        if "Sudarymo data:" in line:
            if i + 1 < len(lines):
                adresas_line = lines[i + 1].strip()
                adresas = re.sub(r"^Teritorija:\s*", "", adresas_line, flags=re.IGNORECASE).replace('"', '').strip()
            break

    process_role_block(lines, "Savininkas", "Savininkas", registro_nr, adresas, unikalus_nr, kadastro_nr)
    process_role_block(lines, "Nuomininkas", "Nuomininkas", registro_nr, adresas, unikalus_nr, kadastro_nr)
    process_role_block(lines, "Patikėtinis", "Patikėtinis", registro_nr, adresas, unikalus_nr, kadastro_nr)

    for match in re.finditer(patterns["panaudos_gavejas"], full_text):
        name_field = match.group(1).strip()
        if "sudarymo data" in name_field.lower():
            continue
        if name_field.lower() == "lietuvos respublika, a.k. 111105555".lower():
            continue

        if ", gim." in name_field.lower():
            entry_type = "fizinis"
        elif ", a.k." in name_field.lower():
            entry_type = "juridinis"
        else:
            entry_type = ""

        split_match = re.split(r",\s*(?:gim\.|a\.k\.)\s*", name_field, maxsplit=1, flags=re.IGNORECASE)
        if len(split_match) == 2:
            name_clean = split_match[0].strip()
            id_or_date = split_match[1].strip()
        else:
            name_clean = name_field
            id_or_date = ""

        first_name, surname = split_name(name_clean, entry_type)

        rows.append((
            registro_nr, adresas, unikalus_nr, kadastro_nr, "Panaudos gavėjas",
            first_name, surname, id_or_date, entry_type,
            info_values[0], info_values[1], info_values[2]
        ))

rows = sorted(set(rows))

with open(output_csv, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
    for row in rows:
        writer.writerow(row)

print(f"Saved {len(rows)} unique rows to {output_csv}")