from pathlib import Path
from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt
import csv
import os
from datetime import date
from collections import defaultdict
import copy

load_dotenv()

def replace_placeholder_in_paragraph(paragraph, replacements):
    """Replace placeholders in a paragraph while preserving formatting."""
    text = paragraph.text
    for key, value in replacements.items():
        if key in text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)

def replace_placeholder_in_cell(cell, replacements):
    """Replace placeholders in a table cell while preserving formatting."""
    for paragraph in cell.paragraphs:
        replace_placeholder_in_paragraph(paragraph, replacements)

def find_paragraph_with_text(doc, text_to_find):
    """Find a paragraph containing specific text."""
    for i, para in enumerate(doc.paragraphs):
        if text_to_find in para.text:
            return i
    return -1

def duplicate_paragraph(doc, para):
    """Create a copy of a paragraph with identical formatting."""
    new_para = doc.add_paragraph()
    new_para.paragraph_format.alignment = para.paragraph_format.alignment
    new_para.paragraph_format.left_indent = para.paragraph_format.left_indent
    new_para.paragraph_format.right_indent = para.paragraph_format.right_indent
    new_para.paragraph_format.space_before = para.paragraph_format.space_before
    new_para.paragraph_format.space_after = para.paragraph_format.space_after
    new_para.style = para.style
    
    for run in para.runs:
        new_run = new_para.add_run(run.text)
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.underline = run.underline
        if run.font.size:
            new_run.font.size = run.font.size
        new_run.font.name = run.font.name
    
    return new_para

def duplicate_paragraph_after(doc, paragraph, new_text=None):
    """Create an exact copy of a paragraph and insert it directly after the original."""
    # Create a new paragraph with same style
    new_paragraph = doc.add_paragraph()
    new_paragraph.style = paragraph.style
    new_paragraph.paragraph_format.alignment = paragraph.paragraph_format.alignment
    new_paragraph.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
    new_paragraph.paragraph_format.right_indent = paragraph.paragraph_format.right_indent
    new_paragraph.paragraph_format.space_before = paragraph.paragraph_format.space_before
    new_paragraph.paragraph_format.space_after = paragraph.paragraph_format.space_after
    
    # Copy all runs with their formatting
    for run in paragraph.runs:
        new_run = new_paragraph.add_run(run.text)
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.underline = run.underline
        if run.font.size:
            new_run.font.size = run.font.size
        new_run.font.name = run.font.name
    
    # Replace text if provided
    if new_text:
        for key, value in new_text.items():
            for run in new_paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)
    
    # Get paragraph index to insert after
    for i, p in enumerate(doc.paragraphs):
        if p == paragraph:
            # Move the new paragraph from the end to directly after the original
            p_element = doc._element.body
            p_element.insert(i+1, new_paragraph._element)
            
            # Remove the duplicate that was added at the end
            doc._element.body.remove(doc.paragraphs[-1]._element)
            
            return doc.paragraphs[i+1]
    
    return new_paragraph

def fill_table_with_plots(doc, plots_data):
    """Fill the table with plot data, adding rows as needed."""
    if not doc.tables or not plots_data:
        return
    
    table = doc.tables[0]
    if len(table.rows) < 2:
        return
    
    # Get the template row (second row)
    template_row = table.rows[1]
    
    # Remove template row placeholder text
    for cell in template_row.cells:
        cell.text = ""
    
    # Fill in first plot data
    if plots_data:
        first_plot = plots_data[0]
        template_row.cells[0].text = first_plot[0]  # Registro Nr
        template_row.cells[1].text = first_plot[2]  # Unikalus Nr
        template_row.cells[2].text = first_plot[3]  # Kadastro Nr
        template_row.cells[3].text = first_plot[1]  # Sklypo adresas
    
    # Add additional rows for the rest of the plots
    for plot in plots_data[1:]:
        new_row = table.add_row()
        new_row.cells[0].text = plot[0]  # Registro Nr
        new_row.cells[1].text = plot[2]  # Unikalus Nr
        new_row.cells[2].text = plot[3]  # Kadastro Nr
        new_row.cells[3].text = plot[1]  # Sklypo adresas

def insert_paragraph_after(doc, paragraph, text, replacements=None):
    """Insert a new paragraph with given text after the reference paragraph."""
    # Find the index of the reference paragraph
    for i, p in enumerate(doc.paragraphs):
        if p == paragraph:
            # Create a new paragraph at the correct position
            new_p = doc.add_paragraph()
            
            # Copy style and formatting from the reference paragraph
            new_p.style = paragraph.style
            new_p.paragraph_format.alignment = paragraph.paragraph_format.alignment
            new_p.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
            new_p.paragraph_format.right_indent = paragraph.paragraph_format.right_indent
            new_p.paragraph_format.space_before = paragraph.paragraph_format.space_before
            new_p.paragraph_format.space_after = paragraph.paragraph_format.space_after
            
            # Add the text with proper formatting by copying runs from the reference
            if len(paragraph.runs) > 0:
                # First run formatting (usually contains the opening quote)
                first_run = new_p.add_run(paragraph.runs[0].text)
                first_run.bold = paragraph.runs[0].bold
                first_run.italic = paragraph.runs[0].italic
                first_run.font.name = paragraph.runs[0].font.name
                if paragraph.runs[0].font.size:
                    first_run.font.size = paragraph.runs[0].font.size
                
                # Main content with replacements
                main_text = text
                if replacements:
                    for key, value in replacements.items():
                        main_text = main_text.replace(key, value)
                
                # Main run formatting (content)
                if len(paragraph.runs) > 1:
                    main_run = new_p.add_run(main_text)
                    main_run.bold = paragraph.runs[1].bold
                    main_run.italic = paragraph.runs[1].italic
                    main_run.font.name = paragraph.runs[1].font.name
                    if paragraph.runs[1].font.size:
                        main_run.font.size = paragraph.runs[1].font.size
                
                # Last run formatting (usually contains the closing quote and semicolon)
                if len(paragraph.runs) > 2:
                    last_run = new_p.add_run(paragraph.runs[2].text)
                    last_run.bold = paragraph.runs[2].bold
                    last_run.italic = paragraph.runs[2].italic
                    last_run.font.name = paragraph.runs[2].font.name
                    if paragraph.runs[2].font.size:
                        last_run.font.size = paragraph.runs[2].font.size
            else:
                # Fallback if no runs in reference paragraph
                new_p.text = text
                if replacements:
                    for key, value in replacements.items():
                        new_p.text = new_p.text.replace(key, value)
            
            # Move the paragraph to the correct position (right after the reference paragraph)
            doc._body._element.insert(i+1, new_p._element)
            # Remove it from the end where it was initially added
            doc._body._element.remove(doc.paragraphs[-1]._element)
            
            # Return the newly inserted paragraph
            return doc.paragraphs[i+1]
    
    return None

def main():
    # Get directories and filenames from environment variables
    etapas_dir = os.environ.get("DIR_ETAPAS")
    template_filename = os.environ.get("TEMPLATE_FILE_NAME")
    csv_filename = os.environ.get("ETAPAS_OUTPUT_FILE_NAME", "aggregated_output.csv")
    
    if not etapas_dir or not template_filename:
        print("Please set DIR_ETAPAS and TEMPLATE_FILE_NAME in your .env file.")
        exit(1)
    
    etapas_path = Path(etapas_dir)
    template_path = etapas_path / template_filename
    csv_path = etapas_path / csv_filename
    
    # Check if files exist
    if not template_path.exists():
        print(f"Template file not found: {template_path}")
        exit(1)
    
    if not csv_path.exists():
        print(f"CSV file not found: {csv_path}")
        exit(1)
    
    print(f"Using template: {template_path}")
    print(f"Reading data from: {csv_path}")
    
    # Create output folder for documents
    output_folder = etapas_path / "letters"
    output_folder.mkdir(exist_ok=True)
    
    # Read the first few bytes to check for delimiter
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        sample = f.read(4096)
        dialect = csv.Sniffer().sniff(sample)
        delimiter = dialect.delimiter
    
    print(f"Detected delimiter: '{delimiter}'")
    
    # Read all rows from CSV with detected delimiter
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=delimiter)
        header = next(reader)  # Skip header row
        rows = list(reader)
    
    print(f"Found {len(rows)} rows in the CSV file")
    
    # Get today's date in YYYY-MM-DD format
    today_date = date.today().strftime("%Y-%m-%d")
    
    # Group rows by unique individual (Vardas, Pavardė, ĮK/Data)
    individuals = defaultdict(list)
    
    for row in rows:
        if len(row) < 9:  # Need at least through the Tipas column
            continue
        
        tipas = row[8].lower()
        
        if tipas == "fizinis":
            # For individuals: combine first name and last name
            vardas = row[5]
            pavarde = row[6]
            id_or_date = row[7]
            individual_key = (vardas, pavarde, id_or_date)
        elif tipas == "juridinis":
            # For companies: use only the company name
            vardas = row[5]
            pavarde = ""
            id_or_date = row[7]
            individual_key = (vardas, "", id_or_date)
        else:
            continue
            
        individuals[individual_key].append(row)
    
    print(f"Found {len(individuals)} unique individuals/companies")
    
    # Process each unique individual
    processed_count = 0
    
    for individual_key, individual_rows in individuals.items():
        vardas, pavarde, id_or_date = individual_key
        tipas = individual_rows[0][8].lower()
        
        # Skip if all entries have no address
        if all(len(row) <= 12 or not row[12] for row in individual_rows):
            print(f"Skipping {vardas} {pavarde} - no address information")
            continue
        
        # Get the first row with an address
        address_row = next((row for row in individual_rows if len(row) > 12 and row[12]), None)
        if not address_row:
            continue
            
        # Create recipient name based on type
        if tipas == "fizinis":
            gavejas_1 = f"{vardas} {pavarde}"
        else:
            gavejas_1 = vardas
            
        # Get address info from the first row with address
        adresas_2 = address_row[12]
        pasto_kodas_3 = address_row[13] if len(address_row) > 13 and address_row[13] else ""
        
        # Collect unique projects and plot data
        unique_projects = {}
        unique_plots = set()
        
        for row in individual_rows:
            # Clean up the elektrine_nr to remove any BOM characters or whitespace
            elektrine_nr = row[9].strip().replace('\ufeff', '')
            
            # Only add unique projects
            if elektrine_nr not in unique_projects:
                unique_projects[elektrine_nr] = {
                    "projekt_nr": row[10],
                    "projekt_pav": row[11]
                }
            
            # Add unique plots (registro_nr, sklypo_adresas, unikalus_nr, kadastro_nr)
            plot_tuple = (row[0], row[1], row[2], row[3])
            unique_plots.add(plot_tuple)
        
        # Convert plots to a list for easier handling
        plot_list = list(unique_plots)
        
        # Common replacements for all projects
        base_replacements = {
            "gavejas_1": gavejas_1,
            "adresas_2": adresas_2,
            "pasto_kodas_3": pasto_kodas_3,
            "proj_data": today_date
        }
        
        # Create a new document from the template
        doc = Document(template_path)
        
        # Replace basic info in the document
        for para in doc.paragraphs:
            replace_placeholder_in_paragraph(para, base_replacements)
        
        # Replace in tables
        for table in doc.tables:
            for row_table in table.rows:
                for cell in row_table.cells:
                    replace_placeholder_in_cell(cell, base_replacements)
        
        # Fill the table with plot data
        fill_table_with_plots(doc, plot_list)
        
        # Find paragraphs containing key placeholders
        proj_pav_para_index = find_paragraph_with_text(doc, "proj_pav_5")
        elektrine_para_index = find_paragraph_with_text(doc, "elektrines_numeris_11")
        
        # Handle project descriptions and attestation paragraphs
        if proj_pav_para_index >= 0 and elektrine_para_index >= 0:
            proj_pav_para = doc.paragraphs[proj_pav_para_index]
            elektrine_para = doc.paragraphs[elektrine_para_index]
            
            # Get ordered list of projects
            project_list = list(unique_projects.items())
            
            if project_list:
                # Process first project by replacing in original paragraphs
                first_elektrine_nr, first_project_info = project_list[0]
                first_project_text = first_project_info["projekt_pav"]
                
                # Replace placeholders in original paragraphs
                for run in proj_pav_para.runs:
                    if "proj_pav_5" in run.text:
                        run.text = run.text.replace("proj_pav_5", first_project_text)
                
                for run in elektrine_para.runs:
                    if "elektrines_numeris_11" in run.text:
                        run.text = run.text.replace("elektrines_numeris_11", first_elektrine_nr)
                
                # Process additional projects
                index_for_proj = proj_pav_para_index + 1
                
                for elektrine_nr, project_info in project_list[1:]:
                    # Create a simpler paragraph with proper indentation
                    new_proj_para = doc.add_paragraph()
                    new_proj_para.style = proj_pav_para.style
                    new_proj_para.paragraph_format.alignment = proj_pav_para.paragraph_format.alignment
                    new_proj_para.paragraph_format.left_indent = proj_pav_para.paragraph_format.left_indent
                    new_proj_para.paragraph_format.right_indent = proj_pav_para.paragraph_format.right_indent
                    new_proj_para.paragraph_format.space_before = proj_pav_para.paragraph_format.space_before
                    new_proj_para.paragraph_format.space_after = proj_pav_para.paragraph_format.space_after
                    new_proj_para.paragraph_format.first_line_indent = proj_pav_para.paragraph_format.first_line_indent
                    
                    # Use just 3 runs for consistent formatting
                    quote_open = new_proj_para.add_run("„")
                    if len(proj_pav_para.runs) > 0:
                        quote_open.font.name = "Arial"  # Explicitly set to Arial
                        quote_open.font.size = Pt(11)   # Set size to 11pt
                        quote_open.bold = proj_pav_para.runs[0].bold
                        quote_open.italic = proj_pav_para.runs[0].italic
                    
                    # Simplified content with just the elektrine number (not the full path)
                    content = f"Energijos iš atsinaujinančių išteklių gamybos paskirties inžinerinio statinio, vėjo elektrinės {elektrine_nr}, statybos projektas"
                    content_run = new_proj_para.add_run(content)
                    if len(proj_pav_para.runs) > 1:
                        content_run.font.name = "Arial"  # Explicitly set to Arial
                        content_run.font.size = Pt(11)   # Set size to 11pt
                        content_run.bold = proj_pav_para.runs[1].bold
                        content_run.italic = proj_pav_para.runs[1].italic
                    
                    quote_close = new_proj_para.add_run("\";")
                    if len(proj_pav_para.runs) > 2:
                        quote_close.font.name = "Arial"  # Explicitly set to Arial
                        quote_close.font.size = Pt(11)   # Set size to 11pt
                        quote_close.bold = proj_pav_para.runs[2].bold
                        quote_close.italic = proj_pav_para.runs[2].italic
                    
                    # Insert at the correct position
                    doc._body._element.insert(index_for_proj, new_proj_para._element)
                    # Remove from end
                    doc._body._element.remove(doc.paragraphs[-1]._element)
                    
                    # Update index for next insertion
                    index_for_proj += 1
                
                # Now add attestation paragraphs after "Pridedama:" with proper bullet points
                pridedama_index = -1
                for i, para in enumerate(doc.paragraphs):
                    if para.text.strip() == "Pridedama:":
                        pridedama_index = i
                        break
                
                if pridedama_index > 0:
                    # Get the style and formatting from the first attestation paragraph
                    for i, para in enumerate(doc.paragraphs):
                        if i > pridedama_index and "Skelbimas apie" in para.text:
                            first_attestation_para = para
                            attestation_index = i + 1
                            break
                    else:
                        # If not found, use defaults
                        first_attestation_para = elektrine_para
                        attestation_index = pridedama_index + 1
                    
                    # Add attestation paragraphs for additional projects
                    for elektrine_nr, project_info in project_list[1:]:
                        # Create new attestation paragraph with bullet point
                        new_attestation = doc.add_paragraph()
                        new_attestation.style = first_attestation_para.style
                        new_attestation.paragraph_format.alignment = first_attestation_para.paragraph_format.alignment
                        new_attestation.paragraph_format.left_indent = first_attestation_para.paragraph_format.left_indent
                        new_attestation.paragraph_format.right_indent = first_attestation_para.paragraph_format.right_indent
                        new_attestation.paragraph_format.space_before = first_attestation_para.paragraph_format.space_before
                        new_attestation.paragraph_format.space_after = first_attestation_para.paragraph_format.space_after
                        new_attestation.paragraph_format.first_line_indent = first_attestation_para.paragraph_format.first_line_indent
                        
                        # Copy bullet points/numbering by using the numbering reference instead of direct XML assignment
                        if hasattr(first_attestation_para, '_p') and first_attestation_para._p.pPr is not None:
                            p_pr = first_attestation_para._p.pPr
                            if p_pr.numPr is not None:
                                # Get the numId and ilvl from source paragraph
                                if p_pr.numPr.numId is not None and p_pr.numPr.ilvl is not None:
                                    num_id = p_pr.numPr.numId.val
                                    ilvl = p_pr.numPr.ilvl.val
                                    
                                    # Apply to the new paragraph through the document's numbering part
                                    from docx.oxml.ns import qn
                                    from docx.oxml import OxmlElement
                                    
                                    # Make sure paragraph has a paragraph properties element
                                    if new_attestation._p.pPr is None:
                                        new_attestation._p.get_or_add_pPr()
                                    
                                    # Add numPr element if it doesn't exist
                                    num_pr = new_attestation._p.pPr.get_or_add_numPr()
                                    
                                    # Set the numId - identifies the numbering definition
                                    num_id_element = OxmlElement('w:numId')
                                    num_id_element.set(qn('w:val'), str(num_id))
                                    num_pr.append(num_id_element)
                                    
                                    # Set the ilvl - identifies the numbering level
                                    ilvl_element = OxmlElement('w:ilvl')
                                    ilvl_element.set(qn('w:val'), str(ilvl))
                                    num_pr.append(ilvl_element)
                        
                        # Add content with consistent formatting for attestation paragraphs
                        text = f"Skelbimas apie energijos iš atsinaujinančių išteklių gamybos paskirties inžinerinio statinio, vėjo elektrinės {elektrine_nr}, projektinių pasiūlymų viešinimą (2 lapai);"
                        run = new_attestation.add_run(text)
                        run.font.name = "Arial"  # Explicitly set to Arial
                        run.font.size = Pt(11)   # Set size to 11pt
                        
                        # Insert at the correct position
                        doc._body._element.insert(attestation_index, new_attestation._element)
                        # Remove from end
                        doc._body._element.remove(doc.paragraphs[-1]._element)
                        
                        # Update index for next insertion
                        attestation_index += 1
        
        # Generate a filename using the recipient name
        safe_name = gavejas_1.replace(" ", "_").replace("/", "-").replace('"', '')
        output_filename = f"{safe_name}.docx"
        output_path = output_folder / output_filename
        
        # Save the document
        doc.save(output_path)
        processed_count += 1
        print(f"Created document: {output_filename} (with {len(unique_projects)} projects and {len(unique_plots)} plots)")
    
    print(f"\nGenerated {processed_count} documents in: {output_folder}")

if __name__ == "__main__":
    main()