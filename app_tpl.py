import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import os
import re
from datetime import datetime, timedelta

def extract_text_by_page(docx_path):
    """Extracts text from the Word document page-wise."""
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        document_xml = docx_zip.read("word/document.xml")
    
    root = ET.fromstring(document_xml)
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    
    paragraphs = root.findall(".//w:p", namespace)
    text_by_page = {}
    page_number = 1
    text_by_page[page_number] = []
    
    for para in paragraphs:
        texts = [node.text for node in para.findall(".//w:t", namespace) if node.text]
        para_text = " ".join(texts).strip()
        
        if para_text:
            text_by_page[page_number].append(para_text)
        
        if para_text.startswith("Page ") and para_text.split()[-1].isdigit():
            page_number += 1
            text_by_page[page_number] = []

    return text_by_page

def extract_section_names(docx_path):
    """Extract section names from the document."""
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        document_xml = docx_zip.read("word/document.xml")

    root = ET.fromstring(document_xml)
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    
    section_names = []
    paragraphs = root.findall(".//w:p", namespace)
    
    for para in paragraphs:
        texts = [node.text for node in para.findall(".//w:t", namespace) if node.text]
        text = " ".join(texts).strip()
        
        if text:
            section_names.append(text)

    return section_names

def extract_table_content(docx_path):
    """Extracts key-value pairs from tables in the document."""
    table_data = []
    
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        document_xml = docx_zip.read("word/document.xml")

    root = ET.fromstring(document_xml)
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    tables = root.findall(".//w:tbl", namespace)

    for table in tables:
        table_rows = []
        for row in table.findall(".//w:tr", namespace):
            cells = row.findall(".//w:tc", namespace)
            row_data = []
            for cell in cells:
                texts = [node.text for node in cell.findall(".//w:t", namespace) if node.text]
                row_data.append(" ".join(texts).strip())

            if row_data:
                table_rows.append(row_data)

        if table_rows:
            table_data.append(table_rows)
    
    return table_data

def validate_revision_history(docx_path):
    """Validates Document Revision History for recent date and non-blank Author."""
    tables = extract_table_content(docx_path)
    revision_history = None

    for table in tables:
        if "Document Revision History" in table[0][0]:  # Checking if it's the right table
            revision_history = table[1:]  # Skip header row
            break
    
    if not revision_history:
        print("❌ Document Revision History table not found!")
        return False

    recent_date = None
    author_missing = False

    for row in revision_history:
        try:
            revision_number = row[0]
            author = row[1]
            revision_date = row[2]

            if not author.strip():
                author_missing = True

            parsed_date = datetime.strptime(revision_date, "%m/%d/%Y")  # Adjust format as per doc
            if not recent_date or parsed_date > recent_date:
                recent_date = parsed_date
        except (IndexError, ValueError):
            continue

    if not recent_date:
        print("❌ Missing or incorrect revision date format!")
    else:
        print(f"✅ Most recent revision date: {recent_date.strftime('%m/%d/%Y')}")

    if author_missing:
        print("❌ Some entries in 'Author' column are blank!")
    else:
        print("✅ All 'Author' entries are filled.")

    return not author_missing and recent_date is not None

def extract_text_from_docx(docx_path):
    """Extracts raw text from a DOCX file by parsing its XML content."""
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        with docx_zip.open("word/document.xml") as xml_file:
            tree = ET.parse(xml_file)
            root = tree.getroot()

            # Extract all text elements from the XML
            namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            texts = [node.text for node in root.findall(".//w:t", namespaces) if node.text]

            return " ".join(texts)  # Join text with spaces for better readability
         

def extract_key_values(text):
    """Extracts key-value pairs from Page 1 text correctly."""
    key_value_pairs = {}

    # Define regex pattern with lookahead to stop at the next key
    patterns = {
        "Project Name": r"Project Name:\s*([^\n]+?)(?=\s+Project ID|$)",
        "Project ID": r"Project ID:\s*([^\n]+?)(?=\s+ReleaseID|$)",
        "ReleaseID": r"ReleaseID\s*:\s*([^\n]+?)(?=\s+Release|$)",
        "Release": r"Release:\s*([^\n]+?)(?=\s+Workstream|$)",
        "Workstream": r"Workstream\s*:\s*([^\n]+?)(?=\s+Document Revision History|$)"
    }

    # Apply regex to extract values
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            key_value_pairs[key] = match.group(1).strip()

    return key_value_pairs


def extract_page1_text(docx_path):
    """Extracts text from the first page using an approximation."""
    full_text = extract_text_from_docx(docx_path)
    return full_text[:200].strip()  # Approximates page 1 text


def read_config(config_path, sheet_name):
    """Reads config Excel file and filters rows starting with 'Page_1_'."""
    df = pd.read_excel(config_path, sheet_name=sheet_name, engine="openpyxl")

    # Debugging: Show column names
    print(f"🔍 Available columns in '{sheet_name}': {df.columns.tolist()}")

    # Ensure columns are correctly named
    expected_columns = ["Key", "Value"]
    df.columns = df.columns.str.strip()  # Trim whitespace
    if not all(col in df.columns for col in expected_columns):
        raise ValueError(f"❌ Excel sheet must contain columns: {expected_columns}. Found: {df.columns.tolist()}")

    # Convert 'Key' and 'Value' to a dictionary
    config_dict = dict(zip(df["Key"], df["Value"]))

    return config_dict

def compare_values(extracted, config):
    """Compares extracted page 1 values with config file values."""
    for key, value in extracted.items():
        config_key = f"page1_{key}"  # Convert key to match config format
        config_value = config.get(config_key)

        if config_value is None:
            print(f"⚠️ {key} not found in config!")
        elif value != config_value:
            print(f"❌ Mismatch: {key} - Extracted: {value}, Config: {config_value}")
        else:
            print(f"✅ Match: {key} - {value}")



def validate_page1_key_values(docx_path, config_file, sheet_name):
    """Validates key-value pairs from Page 1 text against the config file."""
    print(f"🔍 Reading config from: {config_file} | Sheet: {sheet_name}")

    # Load expected key-value pairs from Excel
    config = read_config(config_file, sheet_name)

    # Extract Page 1 text
    page1_text = extract_page1_text(docx_path)

    print("\n📌 Extracted Page 1 Text:")
    print(page1_text)  # Debugging

    # Extract key-value pairs from the text
    extracted_values = extract_key_values(page1_text)

    print("\n📌 Extracted Key-Value Pairs from Page 1:")
    for k, v in extracted_values.items():
        print(f"  - {k}: {v}")

    # Load expected key-value pairs from Excel
    config_df = pd.read_excel(config_file, sheet_name=sheet_name, engine="openpyxl")

    # Ensure required columns exist
    if "Key" not in config_df.columns or "Value" not in config_df.columns:
        raise ValueError("Excel sheet must contain 'Key' and 'Value' columns")

    # Convert DataFrame to dictionary
    config_dict = config_df.set_index("Key")["Value"].to_dict()  # ✅ FIXED VARIABLE NAME

    # Normalize config keys to match extracted format (strip "page_1_" and lowercase)
    normalized_config = {
        key.lower().replace("page_1_", ""): str(value).strip()  # Convert expected values to strings for uniform comparison
        for key, value in config_dict.items()
        if key.lower().startswith("page_1_")  # Only compare relevant fields
    }

    # Normalize extracted keys (convert "Project Name" → "projectname")
    normalized_extracted = {
        key.lower().replace(" ", ""): value.strip()
        for key, value in extracted_values.items()
    }

    # Compare extracted values with expected config values
    mismatches = {}
    for key, expected_value in normalized_config.items():
        if key in normalized_extracted:
            found_value = normalized_extracted[key]
            if found_value != expected_value:
                mismatches[key] = (found_value, expected_value)
        else:
            mismatches[key] = ("❌ Missing Key", expected_value)

    # Output mismatches
    if mismatches:
        print("\n❌ Mismatches found in first-page key-value pairs:")
        for key, (found, expected) in mismatches.items():
            print(f"  - {key}: Found '{found}', Expected '{expected}'")
    else:
        print("\n✅ First-page key-value pairs match the config.")


def extract_toc_sections(docx_path):
    """Extracts section names with heading levels from the Table of Contents."""
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        document_xml = docx_zip.read("word/document.xml")
    root = ET.fromstring(document_xml)
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    toc_sections = []
    for para in root.findall(".//w:p", namespace):
        texts = [node.text for node in para.findall(".//w:t", namespace) if node.text]
        text = " ".join(texts).strip()
        match = re.match(r"(\d+(\.\d+)*)?\s*(.+?)\s+\d+$", text)
        if match:
            heading_level = match.group(1)
            section_name = match.group(3).strip()
            level = heading_level.count(".") + 1 if heading_level else 1
            toc_sections.append((level, section_name))
    return toc_sections

def validate_sections_using_toc(docx_path, config_sections):
    """Validates extracted TOC sections against expected sections while maintaining order."""
    extracted_sections = extract_toc_sections(docx_path)

    def normalize(text):
        return re.sub(r'[^a-zA-Z0-9 ]', '', text).strip().lower()

    expected_sections = [normalize(sec) for sec in config_sections]
    extracted_sections = [(level, normalize(sec)) for level, sec in extracted_sections]  # Normalize names

    missing_sections = [sec for sec in expected_sections if sec not in [s[1] for s in extracted_sections]]
    unexpected_sections = [sec for sec in extracted_sections if sec[1] not in expected_sections]

    return missing_sections, unexpected_sections


def check_embedded_excels(docx_path):
    """Checks for embedded Excel files in the Word document."""
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        return [file for file in docx_zip.namelist() if "embeddings" in file.lower() and file.endswith((".xls", ".xlsx", ".xlsm"))]

def extract_embedded_excel(docx_path):
    """Extracts embedded Excel files from the Word document."""
    extracted_files = []
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        for file in docx_zip.namelist():
            if "embeddings" in file.lower() and file.endswith((".xls", ".xlsx", ".xlsm")):
                output_path = os.path.basename(file)
                with docx_zip.open(file) as src, open(output_path, "wb") as dest:
                    dest.write(src.read())
                extracted_files.append(output_path)
    return extracted_files


def validate_excel_content(excel_path, config):
    """Validate that the Excel file contains required sheets and correct values for Project ID & Release ID."""
    try:
        xls = pd.ExcelFile(excel_path)
        sheet_names = [name.lower() for name in xls.sheet_names]  # Normalize for comparison
        required_sheets = {"summary", "nonfunctional requirement", "logs", "contacts"}

        print(f"\n🔍 Checking {excel_path}...")
        print(f"📄 Found Sheets: {xls.sheet_names}")

        # ✅ Check for required sheets
        missing_sheets = required_sheets - set(sheet_names)
        if missing_sheets:
            print(f"❌ Missing Sheets: {', '.join(missing_sheets)}")
            # return False

        # ✅ Open the Summary sheet
        df = pd.read_excel(xls, sheet_name="Summary", header=None)

        # ✅ Extract expected values from config
        expected_project_id = str(config.get("Page_1_ProjectID", "")).strip().lower()
        expected_release_id = str(config.get("Page_1_ReleaseID", "")).strip().lower()

        # ✅ Ensure enough rows exist before checking
        if df.shape[0] < 8 or df.shape[1] < 2:
            print("❌ Excel does not have enough rows/columns for validation.")
            return False

        # ✅ Read values from A2 and B8
        # ✅ Ensure A2 and B8 are not NaN

        
        # Extract values from fixed cell locations
        project_id_value = str(df.iloc[1, 0]).strip().lower()  # A2
        release_id_value = str(df.iloc[7, 1]).strip().lower()  # B8

        # Print extracted values
        print(f"Extracted Project ID: {project_id_value if project_id_value else 'Not Found'}")
        print(f"Extracted Release ID: {release_id_value if release_id_value else 'Not Found'}")
        # project_id_value = df.iloc[1, 0] if pd.notna(df.iloc[1, 0]) else ""  # A2
        # release_id_value = df.iloc[7, 1] if pd.notna(df.iloc[7, 1]) else ""  # B8

        # ✅ Convert values to lowercase strings after handling NaN
        project_id_value = str(project_id_value).strip().lower()
        release_id_value = str(release_id_value).strip().lower()

        print(f"Project ID (A2): '{project_id_value}'")
        print(f"Release ID (B8): '{release_id_value}'")

        if project_id_value != expected_project_id:
            print(f"❌ A2 (Project ID) Mismatch: Expected '{expected_project_id}', Found '{project_id_value}'")
            return False

        if release_id_value != expected_release_id:
            print(f"❌ B8 (Release ID) Mismatch: Expected '{expected_release_id}', Found '{release_id_value}'")
            return False

        print(f"✅ Validation Passed: A2='{expected_project_id}', B8='{expected_release_id}'")
        return True

    except Exception as e:
        print(f"⚠️ Error reading Excel file: {e}")
        return False

def extract_revision_history(docx_path):
    """Extracts the Document Revision History table from a Word document using XML parsing."""

    # Open the .docx file as a zip archive
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        with docx_zip.open("word/document.xml") as doc_xml:
            xml_content = doc_xml.read()

    # Parse XML
    root = ET.fromstring(xml_content)
    namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    # Find all text elements
    paragraphs = root.findall(".//w:t", namespaces)

    found_section = False
    revision_table = None

    # Iterate through all text elements to locate "Document Revision History"
    for para in paragraphs:
        text = para.text.strip() if para.text else ""

        # Check if we found "Document Revision History"
        if "Document Revision History" in text:
            found_section = True
            continue

        # Once found, locate the next table <w:tbl>
        if found_section:
            for table in root.findall(".//w:tbl", namespaces):
                revision_table = table
                break  # Stop at the first table found after the heading

            if revision_table is not None and len(revision_table)>0:
                break  # Stop searching once the table is found

    # Handle case when no table is found
    if revision_table is None:
        print("❌ Document Revision History table not found!")
        return None

    # Extract rows from the table
    table_data = []
    for row in revision_table.findall(".//w:tr", namespaces):
        row_data = []

        # Extract text from each table cell <w:tc>
        for cell in row.findall(".//w:tc", namespaces):
            # Combine all <w:t> elements within the cell
            cell_text = " ".join(t.text.strip() for t in cell.findall(".//w:t", namespaces) if t.text)
            cell_text = re.sub(r"\s*/\s*", "/", cell_text)  # Remove spaces around slashes (for dates)
            row_data.append(cell_text)

        if any(row_data):  # Ignore empty rows
            table_data.append(row_data)

    return table_data

def extract_footer_text(docx_path):
    """Extracts footer text from a Word document (.docx)."""
    footer_text = ""

    try:
        with zipfile.ZipFile(docx_path, "r") as docx_zip:
            for file in docx_zip.namelist():
                if "footer" in file.lower() and file.endswith(".xml"):
                    with docx_zip.open(file) as f:
                        xml_content = f.read()
                        root = ET.fromstring(xml_content)

                        # Iterate through all XML elements
                        for elem in root.iter():
                            if elem.text and "PAGE" not in elem.text:  # Ignore PAGE fields
                                footer_text += elem.text.strip() + " "
        return footer_text.strip()

    except Exception as e:
        print(f"⚠️ Error extracting footer: {e}")
        return None

def validate_footer_contains_project(docx_path, config):
    """Checks if the footer CONTAINS the Project Name from the config."""
    extracted_footer = extract_footer_text(docx_path)
    expected_project_name = str(config.get("Page_1_ProjectName", "")).strip()

    print(f"🔍 Extracted Footer: {extracted_footer if extracted_footer else 'Not Found'}")
    print(f"🔍 Expected Project Name: {expected_project_name}")

    if not extracted_footer:
        return False, "❌ Footer not found."

    if expected_project_name.lower() in extracted_footer.lower():
        return True, "✅ Footer contains the Project Name."
    else:
        return False, f"❌ Footer does not contain Project Name. Found: '{extracted_footer}'"



def validate_document(docx_path, excel_path):
    """Validates the Word document using a config extracted from an Excel file."""
    file_name = os.path.basename(docx_path).replace(".docx", "").replace("-", "_")
    
    config = pd.read_excel(excel_path, sheet_name=file_name, engine="openpyxl").set_index("Key")["Value"].to_dict()
    
    if "Sections" in config:
        config["Sections"] = [s.strip() for s in str(config["Sections"]).split(",")]

    text_by_page = extract_text_by_page(docx_path)
    extracted_sections = extract_section_names(docx_path)
    tables = extract_table_content(docx_path)
    missing_sections, extra_sections = validate_sections_using_toc(docx_path, config.get("Sections", []))
    
    embedded_excels = extract_embedded_excel(docx_path)

    print(f"\n=== Validation Report for {file_name} ===")
    print(f"\n✅ Section Validation:")
    # print(f"  - Missing Sections: {missing_sections if missing_sections else 'None'}")
    # print(f"  - Extra Sections: {extra_sections if extra_sections else 'None'}")
    print(f"  - Missing Sections: {[name.title() for name in missing_sections] if missing_sections else 'None'}")
    print(f"  - Extra Sections: {[name.title() for _, name in extra_sections] if extra_sections else 'None'}")


    # print(f"\n✅ Extracted Table Content: {tables}")
    
    print(f"\n✅ Embedded Excel Files: {embedded_excels if embedded_excels else 'None'}")


    if embedded_excels:
        for excel_file in embedded_excels:
            result = validate_excel_content(excel_file, config)
            
            if result:
                print(f"✅ Excel Validation Passed for {excel_file}")
            else:
                print(f"❌ Excel Validation Failed for {excel_file}")
    else:
        print("❌ No embedded Excel files found in the document.")

    

    # print("\n✅ Checking Document Revision History...")
    # validate_revision_history(docx_path)

    today = datetime.today()
    one_week_ago = today - timedelta(days=7)
    revision_history = extract_revision_history(docx_path)

    if revision_history:
        print("\n✅ Extracted Document Revision History Table:")
        for row in revision_history:
            print(" | ".join(row))
        
        # Extract second-row author (if available) and last revision date
        author_exists = bool(revision_history[1][1].strip()) if len(revision_history) > 1 else False

        try:
            revision_date = datetime.strptime(revision_history[1][2].strip(), "%m/%d/%Y")
            recent_date = revision_date >= one_week_ago
        except ValueError:
            recent_date = False  # Invalid date format or missing date
    
        # Print summary
        print("\n✅ Revision History Validation:")
        print(f"  - Author Present: {'✅ Yes' if author_exists else '❌ No'}")
        print(f"  - Recent Revision (within last 7 days): {'✅ Yes' if recent_date else '❌ No'}")
    else:
        print("\n❌ Document Revision History table not found!")    

    print("\n✅ Validating first-page key-value pairs...")
    validate_page1_key_values(docx_path, excel_path, sheet_name.lower())
    result, message = validate_footer_contains_project(docx_path, config)
    print(message)

# Example Usage
# validate_document("sample.docx", "config.xlsx")

# Example usage
doc_path = "D:\\Desktop 2024\\PycharmProjects\\RESTAPI\\AutomatedDocumentReview\\performance-testing-strategy.docx"
excel_path = "D:\\Desktop 2024\\PycharmProjects\\RESTAPI\\AutomatedDocumentReview\\config.xlsx"  # Excel file with multiple tabs for different Word documents
sheet_name = "Performance_Testing_Strategy"  # Name of the tab to validate

validate_document(doc_path, excel_path)

