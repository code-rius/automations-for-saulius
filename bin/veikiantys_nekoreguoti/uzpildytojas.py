import os
import csv
import re
from pathlib import Path
from dotenv import load_dotenv
from docxtpl import DocxTemplate

load_dotenv()

CSV_DIR = os.getenv("DIR_UZPILDYTOJAS_XLS")
DOCX_DIR = os.getenv("DIR_UZPILDYTOJAS_docx")

if not CSV_DIR or not DOCX_DIR:
    raise SystemExit("Set DIR_UZPILDYTOJAS_XLS and DIR_UZPILDYTOJAS_docx in .env")

csv_dir = Path(CSV_DIR)
docx_dir = Path(DOCX_DIR)

template_path = next(docx_dir.glob("*.docx"), None)
if not template_path:
    raise SystemExit("No .docx template found in template dir")

# get prefix from template name
template_stem = template_path.stem
if "_" in template_stem:
    template_prefix = template_stem.rsplit("_", 1)[0]
else:
    template_prefix = template_stem

csv_files = list(csv_dir.glob("*.csv"))
if not csv_files:
    raise SystemExit("No .csv files found in CSV dir")


def safe_name(s: str) -> str:
    s = str(s or "")
    s = re.sub(r'[<>:"/\\|?*\n\r\t]', "_", s).strip()
    return s or "row"


for csv_path in csv_files:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        print("Detected headers:", reader.fieldnames)

        for i, row in enumerate(reader, start=1):
            # fill docx
            context = {
                "Bendras_Nr__1": row.get("Bendras Nr.", "") or "",
                "Adresas_1": row.get("Adresas", "") or "",
                "Pavadinimas_1": row.get("Pavadinimas", "") or "",
            }

            doc = DocxTemplate(str(template_path))
            doc.render(context)

            # 📛 filename: use Bendras Nr. (this is your VE29_M64)
            code = row.get("Bendras Nr.", "") or f"row_{i}"
            code = safe_name(code)

            final_name = f"{template_prefix}_{code}.docx"
            out_path = csv_dir / final_name

            counter = 1
            while out_path.exists():
                out_path = csv_dir / f"{template_prefix}_{code}_{counter}.docx"
                counter += 1

            doc.save(out_path)
            print("Saved:", out_path)
