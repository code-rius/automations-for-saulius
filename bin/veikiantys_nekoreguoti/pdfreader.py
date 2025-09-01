import os
from pathlib import Path
import pdfplumber
import re
import csv
from dotenv import load_dotenv

load_dotenv()

def get_elektrine_directories():
    """Get all directories listed in DIR_ELEKTRINE_X environment variables."""
    elektrine_dirs = []
    i = 1
    while True:
        dir_key = f"DIR_ELEKTRINE_{i}"
        dir_path = os.environ.get(dir_key)
        if not dir_path:
            break
        elektrine_dirs.append(dir_path)
        i += 1
    
    if not elektrine_dirs:
        print("No DIR_ELEKTRINE_X entries found in .env file.")
        exit(1)
    
    return elektrine_dirs

def read_info_file(folder_path):
    """Read info from a .txt file named after the folder."""
    folder_name = folder_path.name
    info_txt_path = folder_path / f"{folder_name}.txt"
    info_values = ["", "", ""]
    
    if info_txt_path.exists():
        print(f"Found info file: {info_txt_path}")
        with open(info_txt_path, encoding="utf-8-sig") as f:
            info = {"BENDRAS_NR": "", "PROJEKTO_NR": "", "PAVADINIMAS": ""}
            for line in f:
                for i, key in enumerate(info.keys()):
                    if line.startswith(f"{key}="):
                        info[key] = line.strip().split("=", 1)[1]
                        info_values[i] = info[key]
    else:
        print(f"Warning: Info file {info_txt_path} not found. Using empty values for extra columns.")
    
    return info_values

def split_name(name_clean, entry_type):
    """Split a name into first name and surname based on entry type."""
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

def process_role_block(lines, start_keyword, role_label, registro_nr, adresas, unikalus_nr, kadastro_nr, info_values):
    """Process a block of text to extract people with specific roles."""
    rows = []
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
    return rows

def process_pdf_file(pdf_file, patterns, info_values):
    """Process a single PDF file and extract all relevant information."""
    rows = []
    print(f"Reading PDF: {pdf_file.name}")
    
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

    # Process each role type
    rows.extend(process_role_block(lines, "Savininkas", "Savininkas", registro_nr, adresas, unikalus_nr, kadastro_nr, info_values))
    rows.extend(process_role_block(lines, "Nuomininkas", "Nuomininkas", registro_nr, adresas, unikalus_nr, kadastro_nr, info_values))
    rows.extend(process_role_block(lines, "Patikėtinis", "Patikėtinis", registro_nr, adresas, unikalus_nr, kadastro_nr, info_values))

    # Process panaudos gavėjas entries
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
    
    return rows

def process_directory(folder_path):
    """Process a single directory containing PDF files."""
    print(f"\n=== Processing directory: {folder_path} ===")
    output_csv = folder_path / "output.csv"
    
    # Read info from txt file
    info_values = read_info_file(folder_path)
    
    patterns = {
        "registro_nr": re.compile(r"Registro Nr\.:?\s*([^\n]+)"),
        "unikalus_nr": re.compile(r"Unikalus daikto numeris:?\s*([^\n]+)"),
        "kadastro_nr": re.compile(r"pavadinimas:?\s*([^\n]+)"),
        "panaudos_gavejas": re.compile(r"Panaudos gavėjas:?\s*([^\n]+)")
    }
    
    rows = []
    pdf_files = list(folder_path.glob("*.pdf"))
    if not pdf_files:
        print(f"Warning: No PDF files found in {folder_path}")
        return None
    
    # Process each PDF file
    for pdf_file in pdf_files:
        rows.extend(process_pdf_file(pdf_file, patterns, info_values))
    
    # Remove duplicates and sort
    rows = sorted(set(rows))
    
    # Define header row
    header_row = [
        "Registro Nr", 
        "Sklypo adresas", 
        "Unikalus Nr", 
        "Kadastro Nr", 
        "Rolė", 
        "Vardas", 
        "Pavardė", 
        "ĮK/Data", 
        "Tipas", 
        "Elektrinės Nr", 
        "Projekto Nr", 
        "Projekto pavadinimas",
        "Deklaruotas adresas",
        "Pašto kodas"
    ]
    
    # Write output file
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(header_row)
        writer.writerows(rows)
    
    print(f"Saved {len(rows)} unique rows to {output_csv}")
    return output_csv

def aggregate_files(etapas_dir, processed_files):
    """Aggregate all processed CSV files into one file."""
    if not processed_files:
        print("No files to aggregate.")
        return
    
    etapas_path = Path(etapas_dir)
    if not etapas_path.exists():
        print(f"Creating directory: {etapas_path}")
        etapas_path.mkdir(parents=True, exist_ok=True)
    
    output_path = etapas_path / "aggregated_output.csv"
    
    # Collect all rows and headers
    all_rows = []
    header = None
    
    for csv_path in processed_files:
        if not csv_path or not Path(csv_path).exists():
            continue
        
        print(f"Adding data from: {csv_path}")
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            file_header = next(reader)  # Read header
            
            # Use the first file's header as the main header
            if header is None:
                header = file_header
            
            # Add all data rows
            rows_count = 0
            for row in reader:
                all_rows.append(row)
                rows_count += 1
            
            print(f"Added {rows_count} rows from {csv_path}")
    
    if not all_rows:
        print("No data found in any CSV files!")
        return
    
    # Write the aggregated data to the output file
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(all_rows)
    
    print(f"Successfully aggregated {len(all_rows)} rows into {output_path}")

def main():
    """Main function to orchestrate the entire process."""
    # Get directories from environment
    elektrine_dirs = get_elektrine_directories()
    print(f"Found {len(elektrine_dirs)} directories to process")
    
    # Process each directory
    processed_files = []
    for elektrine_dir in elektrine_dirs:
        folder_path = Path(elektrine_dir)
        if not folder_path.exists():
            print(f"Directory not found: {folder_path}")
            continue
        
        output_file = process_directory(folder_path)
        if output_file:
            processed_files.append(output_file)
    
    # Aggregate all files
    etapas_dir = os.environ.get("DIR_ETAPAS")
    if etapas_dir:
        print(f"\n=== Aggregating files to {etapas_dir} ===")
        aggregate_files(etapas_dir, processed_files)
    else:
        print("\nSkipping aggregation: DIR_ETAPAS not set in .env file")
    
    print("\nAll operations completed successfully!")

if __name__ == "__main__":
    main()