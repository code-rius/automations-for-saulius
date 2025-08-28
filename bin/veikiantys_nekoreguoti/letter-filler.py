from pathlib import Path
from dotenv import load_dotenv
from docx import Document
import csv
import os

load_dotenv()

DOCX_DIRECTORY = os.environ.get("DOCX_DIRECTORY")
CSV_DIRECTORY = os.environ.get("CSV_DIRECTORY")

if not DOCX_DIRECTORY or not CSV_DIRECTORY:
    print("Please set DOCX_DIRECTORY and CSV_DIRECTORY in your .env file.")
    exit(1)

docx_folder = Path(DOCX_DIRECTORY)
csv_folder = Path(CSV_DIRECTORY)

# Find first .docx file
docx_files = list(docx_folder.glob("*.docx"))
if not docx_files:
    print("No .docx files found in the directory.")
    exit(1)
docx_path = docx_files[0]

# Find first .csv file
csv_files = list(csv_folder.glob("*.csv"))
if not csv_files:
    print("No .csv files found in the directory.")
    exit(1)
csv_path = csv_files[0]

# Read all rows from CSV
with open(csv_path, "r", encoding="utf-8") as f:
    reader = csv.reader(f)
    rows = list(reader)

for idx, row in enumerate(rows):
    # Prepare values from CSV columns
    gavejas_1 = f"{row[5]} {row[6]}"
    adresas_2 = row[12]
    pasto_kodas_3 = row[13]
    proj_nr_4 = row[15]
    proj_pav_5 = row[16]
    registro_7 = row[0]
    unikalus_8 = row[2]
    kadastro_9 = row[3]
    sklypo_adresas_10 = row[1]

    replacements = {
        "gavejas_1": gavejas_1,
        "adresas_2": adresas_2,
        "pasto_kodas_3": pasto_kodas_3,
        "proj_nr_4": proj_nr_4,  # <-- corrected key
        "proj_pav_5": proj_pav_5,
        "registro_7": registro_7,
        "unikalus_8": unikalus_8,
        "kadastro_9": kadastro_9,
        "sklypo_adresas_10": sklypo_adresas_10
    }

    doc = Document(docx_path)

    # Replace in paragraphs (preserving formatting)
    for para in doc.paragraphs:
        for key, value in replacements.items():
            if key in para.text:
                for run in para.runs:
                    run.text = run.text.replace(key, value)

    # Replace in tables (preserving formatting)
    for table in doc.tables:
        for row_table in table.rows:
            for cell in row_table.cells:
                for key, value in replacements.items():
                    if key in cell.text:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                run.text = run.text.replace(key, value)

    # Save updated docx to the same directory, with unique name
    output_path = docx_folder / f"filled_{idx+1}_{docx_path.name}"
    doc.save(output_path)

print(f"Created {len(rows)} docx files.")