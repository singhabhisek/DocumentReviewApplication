import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import os
import openpyxl
import re
from datetime import datetime, timedelta
from st_aggrid import AgGrid, GridOptionsBuilder

# st.set_page_config(layout="wide", page_title="Word Validation App", page_icon="📊")


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



# def validate_page1_key_values(docx_path, config_file, sheet_name):
#     """Validates key-value pairs from Page 1 text against the config file."""
#     print(f"🔍 Reading config from: {config_file} | Sheet: {sheet_name}")

#     # Load expected key-value pairs from Excel
#     config = read_config(config_file, sheet_name)

#     # Extract Page 1 text
#     page1_text = extract_page1_text(docx_path)

#     print("\n📌 Extracted Page 1 Text:")
#     print(page1_text)  # Debugging

#     # Extract key-value pairs from the text
#     extracted_values = extract_key_values(page1_text)

#     print("\n📌 Extracted Key-Value Pairs from Page 1:")
#     for k, v in extracted_values.items():
#         print(f"  - {k}: {v}")

#     # Load expected key-value pairs from Excel
#     config_df = pd.read_excel(config_file, sheet_name=sheet_name, engine="openpyxl")

#     # Ensure required columns exist
#     if "Key" not in config_df.columns or "Value" not in config_df.columns:
#         raise ValueError("Excel sheet must contain 'Key' and 'Value' columns")

#     # Convert DataFrame to dictionary
#     config_dict = config_df.set_index("Key")["Value"].to_dict()  # ✅ FIXED VARIABLE NAME

#     # Normalize config keys to match extracted format (strip "page_1_" and lowercase)
#     normalized_config = {
#         key.lower().replace("page_1_", ""): str(value).strip()  # Convert expected values to strings for uniform comparison
#         for key, value in config_dict.items()
#         if key.lower().startswith("page_1_")  # Only compare relevant fields
#     }

#     # Normalize extracted keys (convert "Project Name" → "projectname")
#     normalized_extracted = {
#         key.lower().replace(" ", ""): value.strip()
#         for key, value in extracted_values.items()
#     }

#     # Compare extracted values with expected config values
#     mismatches = {}
#     for key, expected_value in normalized_config.items():
#         if key in normalized_extracted:
#             found_value = normalized_extracted[key]
#             if found_value != expected_value:
#                 mismatches[key] = (found_value, expected_value)
#         else:
#             mismatches[key] = ("❌ Missing Key", expected_value)

#     # Output mismatches
#     if mismatches:
#         print("\n❌ Mismatches found in first-page key-value pairs:")
#         for key, (found, expected) in mismatches.items():
#             print(f"  - {key}: Found '{found}', Expected '{expected}'")
#     else:
#         print("\n✅ First-page key-value pairs match the config.")

# def validate_page1_key_values(docx_path, config_file, sheet_name):
#     """Validates key-value pairs from Page 1 text against the config file and returns validation results."""
    
#     # Load expected key-value pairs from Excel
#     config = read_config(config_file, sheet_name)

#     # Extract Page 1 text
#     page1_text = extract_page1_text(docx_path)

#     # Extract key-value pairs from the text
#     extracted_values = extract_key_values(page1_text)

#     # Load expected key-value pairs from Excel
#     config_df = pd.read_excel(config_file, sheet_name=sheet_name, engine="openpyxl")

#     # Ensure required columns exist
#     if "Key" not in config_df.columns or "Value" not in config_df.columns:
#         raise ValueError("Excel sheet must contain 'Key' and 'Value' columns")

#     # Convert DataFrame to dictionary
#     config_dict = config_df.set_index("Key")["Value"].to_dict()  # ✅ FIXED VARIABLE NAME

#     # Normalize config keys to match extracted format (strip "page_1_" and lowercase)
#     normalized_config = {
#         key.lower().replace("page_1_", ""): str(value).strip()  # Convert expected values to strings for uniform comparison
#         for key, value in config_dict.items()
#         if key.lower().startswith("page_1_")  # Only compare relevant fields
#     }

#     # Normalize extracted keys (convert "Project Name" → "projectname")
#     normalized_extracted = {
#         key.lower().replace(" ", ""): value.strip()
#         for key, value in extracted_values.items()
#     }

#     # Compare extracted values with expected config values
#     mismatches = {}
#     for key, expected_value in normalized_config.items():
#         if key in normalized_extracted:
#             found_value = normalized_extracted[key]
#             if found_value != expected_value:
#                 mismatches[key] = (found_value, expected_value)
#         else:
#             mismatches[key] = ("❌ Missing Key", expected_value)

#     # Return mismatches (or an empty dictionary if no mismatches)
#     if mismatches:
#         return {"status": "❌ Mismatches found", "details": mismatches}
#     else:
#         return {"status": "✅ First-page key-value pairs match the config.", "details": {}}

def validate_page1_key_values(docx_path, selected_row):
    """Validates key-value pairs from Page 1 text against the selected row data and returns validation results."""
    
    # Extract Page 1 text from the document
    page1_text = extract_page1_text(docx_path)

    # Extract key-value pairs from the Page 1 text
    extracted_values = extract_key_values(page1_text)

    # Normalize the keys for both the selected row and the extracted values from the document
    normalized_selected_row = {
        'releaseid': selected_row['Release ID'],  # The first column value of the selected row
        'releasename': selected_row['Release Name'],  # The second column value of the selected row
        'projectid': selected_row['Project ID'],  # The third column value of the selected row
        'projectname': selected_row['Project Name'],  # The fourth column value of the selected row
        'appid': selected_row['Application ID'],  # The fifth column value of the selected row
        'appname': selected_row['Application Name']  # The sixth column value of the selected row
    }

    # Normalize extracted keys from Page 1 (convert spaces to lowercase)
    normalized_extracted = {
        key.lower().replace(" ", ""): value.strip()
        for key, value in extracted_values.items()
    }

    # Compare the extracted values from the document with the selected row values
    mismatches = {}
    for key, expected_value in normalized_selected_row.items():
        if key in normalized_extracted:
            found_value = normalized_extracted[key]
            if found_value != expected_value:
                mismatches[key] = (found_value, expected_value)
        else:
            mismatches[key] = ("❌ Missing Key", expected_value)

    # Return mismatches (or an empty dictionary if no mismatches)
    if mismatches:
        return {"status": "❌ Mismatches found", "details": mismatches}
    else:
        return {"status": "✅ All key-value pairs match the document.", "details": {}}

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

def validate_footer_contains_project(docx_path, projectName): #config
    """Checks if the footer CONTAINS the Project Name from the config."""
    extracted_footer = extract_footer_text(docx_path)
    # expected_project_name = str(config.get("Page_1_ProjectName", "")).strip()
    expected_project_name = str(projectName).strip()

    print(f"🔍 Extracted Footer: {extracted_footer if extracted_footer else 'Not Found'}")
    print(f"🔍 Expected Project Name: {expected_project_name}")

    if not extracted_footer:
        return False, "❌ Footer not found."

    if expected_project_name.lower() in extracted_footer.lower():
        return True, f"✅ Footer contains the Project Name as {expected_project_name}"
    else:
        return False, f"❌ Footer does not contain Project Name as {expected_project_name}. Found: '{extracted_footer}'"



def extract_excel_data_from_embedded(file_path):
    """Extracts data from embedded Excel file for specific sheets and cells."""
    sheets_to_check = [
        "summary", "logs", "contacts", "architecture", "nonfunctional requirement", "test data"
    ]
    extracted_data = {}
    matching_sheets = []

    try:
        # Load the embedded Excel file
        wb = openpyxl.load_workbook(file_path)

        # Iterate through each sheet in the Excel file
        for sheet_name in wb.sheetnames:
            if sheet_name.lower() in sheets_to_check:
                sheet = wb[sheet_name]
                
                # Extract values from A2 and B8
                a2_value = sheet["A2"].value
                b8_value = sheet["B8"].value
                
                # Store the extracted data
                extracted_data[sheet_name] = {"A2": a2_value, "B8": b8_value}
                matching_sheets.append(sheet_name)
                
            # Stop searching once we have 3 matching sheets
            if len(matching_sheets) >= 3:
                break

    except Exception as e:
        print(f"Error processing embedded Excel: {e}")

    return extracted_data, matching_sheets


# Define the mapping for more human-readable names
key_mapping = {
    "projectid": "Project Id",
    "projectname": "Project Name",
    "releaseid": "Release ID",
    "workstream": "Workstream",
    "author": "Author",
    "revisiondate": "Revision Date",
    "revisionnumber": "Revision Number",
    "description": "Description",
    "performancetestplanversion": "Test Plan Version",
    "appid" : "Application ID",
    "applicationid": "Application ID",
    "appname":"Application Name",
    "applicationname": "Application Name",
    "releasename": "Release Name"
}


# Main validation function
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

    # Section 1: Section Validation
    st.subheader("1. Section Validation")
    st.write(">>> Matching Sections:")

    # Extracted section names from the document
    extracted_sections = extract_section_names(docx_path)

    # Get the relevant sections from the config (make sure to strip and split them correctly)
    configured_sections = config.get("Sections", [])

    # Check if extracted sections match the configured sections
    matching_sections = [section for section in extracted_sections if section.strip().lower() in [configured_section.lower() for configured_section in configured_sections]]

    if matching_sections:
        st.write(f"  - Matched Sections: {', '.join([section.title() for section in matching_sections])}")
    else:
        st.write("  - No matching sections found.")

    # Check for missing and extra sections
    # missing_sections = [section for section in configured_sections if section.strip().lower() not in [s.lower() for s in extracted_sections]]
    # extra_sections = [section for section in extracted_sections if section.strip().lower() not in [s.lower() for s in configured_sections]]

    missing_sections, extra_sections = validate_sections_using_toc(docx_path, config.get("Sections", []))

    if missing_sections:
        st.write(f"  - Missing Sections: {[name.title() for name in missing_sections] if missing_sections else 'None'}")  # Print missing sections
        st.warning("Please manually check the missing sections in the document.")
    else:
        st.write("  - All required sections are present.")

    if extra_sections:
        st.write(f"  - Extra Sections: {[name.title() for _, name in extra_sections] if extra_sections else 'None'}")  # Print extra sections
        st.warning("Please manually check the extra sections that were found in the document.")
    st.write("\n")  # Spacer


    # Section 2: Document Revision History
    st.subheader("2. Document Revision History Section")
    today = datetime.today()
    one_week_ago = today - timedelta(days=7)
    revision_history = extract_revision_history(docx_path)

    if revision_history:
        st.write(">>> Revision History Found:")
        for row in revision_history:
            st.write(" | ".join(row))
        
        # Extract second-row author (if available) and last revision date
        author_exists = bool(revision_history[1][1].strip()) if len(revision_history) > 1 else False

        try:
            revision_date = datetime.strptime(revision_history[1][2].strip(), "%m/%d/%Y")
            recent_date = revision_date >= one_week_ago
        except ValueError:
            recent_date = False  # Invalid date format or missing date
    
        # Summary of revision history check
        st.write("\n>>> Revision History Validation:")
        st.write(f"  - Author Present: {'✅ Yes' if author_exists else '❌ No'}")
        st.write(f"  - Recent Revision (within last 7 days): {'✅ Yes' if recent_date else '❌ No'}")
    else:
        st.write("\n❌ Document Revision History table not found!")
    st.write("\n")  # Spacer

    # Section 3: Page 1 Details Validation
    st.subheader("3. Page 1 Details Validation Section")
    st.write(">>> Checking key-value pairs on Page 1...")
    validation_result = validate_page1_key_values(docx_path,  selected_row ) #excel_path,sheet_name.lower()

    # Display the result in Streamlit
    st.write(validation_result["status"])

    # If there are mismatches, display them as bullet points
    if validation_result["details"]:
        st.write("❌ Mismatches found in the first-page key-value pairs:")
        for key, (found, expected) in validation_result["details"].items():
            # st.write(f"  - {key}: Found '{found}', Expected '{expected}'")
            # Use the key_mapping to get the human-readable key name
            human_readable_key = key_mapping.get(key, key)  # If no match, use the original key
            # st.write(f"  - {human_readable_key}: Found '{found}', Expected '{expected}'")
            if found == "❌ Missing Key":  # Case for missing key
                st.write(f"  - {human_readable_key}: The key was not found, Expected '{expected}'")
            else:  # Case for mismatched values
                st.write(f"  - {human_readable_key}: Found '{found}', Expected '{expected}'")
    else:
        st.write("✅ All key-value pairs match the configuration.")
    result, message = validate_footer_contains_project(docx_path, selected_row[df.columns[3]]) #config
    st.write(message)  # Output the validation result message
    st.write("\n")  # Spacer

    # Section 4: Table of Content Validation
    st.subheader("4. Table of Content Validation Section")
    st.write(">>> Checking for Mandatory Sections in Table of Contents...")
    missing_sections_toc = validate_sections_using_toc(docx_path, config.get("Sections", []))[0]  # Get missing sections
    if missing_sections_toc:
        st.write(f"  - Missing Sections in TOC: {[name.title() for name in missing_sections_toc]}")
        st.warning("Please manually check the missing sections in the Table of Contents.")
    else:
        st.write("  - All mandatory sections are present in the Table of Contents.")
    st.write("\n")  # Spacer

    # Section 5: Embedded Excel Check
    st.subheader("5. Embedded Excel Validation Section")

    embedded_excels = extract_embedded_excel(docx_path)  # Function to get embedded Excel files
    
    if not embedded_excels:
        st.write("❌ No embedded Excel files found in the document.")
        return  # Exit if no embedded Excel files are found

    for excel_file in embedded_excels:
        extracted_data, matching_sheets = extract_excel_data_from_embedded(excel_file)
        
        if len(matching_sheets) >= 3:
            # Once we find an Excel file with at least 3 matching sheets, break out of the loop
            st.write(f"✅ Embedded Excel File: {excel_file} contains the required sheets: {', '.join(matching_sheets)}")
            
            # Now, validate A2 and B8 values with Project ID and Release ID from the config file
            config = read_config(excel_path, sheet_name.lower())  # Assuming you have a method to read the config
            project_id = config.get("Project ID")
            release_id = config.get("Release ID")
            
            # Check if values from A2 and B8 match Project ID and Release ID
            a2_value = extracted_data.get(matching_sheets[0], {}).get("A2")
            b8_value = extracted_data.get(matching_sheets[0], {}).get("B8")
            
            if a2_value == project_id:
                st.write(f"  - ✅ A2 matches the Project ID: {a2_value}")
            else:
                st.write(f"  - ❌ A2 does not match the Project ID: {a2_value}")
                
            if b8_value == release_id:
                st.write(f"  - ✅ B8 matches the Release ID: {b8_value}")
            else:
                st.write(f"  - ❌ B8 does not match the Release ID: {b8_value}")
            
            # Break the loop after finding the first valid Excel file
            break
    else:
        st.write("❌ Please check if you have attached the correct Non-Functional Requirement sheet template.")

    # if embedded_excels:
    #     st.write(">>> Checking Embedded Excel Files for Matching Values...")
    #     for excel_file in embedded_excels:
    #         result = validate_excel_content(excel_file, config)
    #         if result:
    #             st.write(f"✅ Excel Validation Passed for {excel_file}")
    #         else:
    #             st.write(f"❌ Excel Validation Failed for {excel_file}")
    # else:
    #     st.write("❌ No embedded Excel files found in the document.")
    st.write("\n")  # Spacer

    # Streamlit UI for file upload
st.title("Test Plan Validation Application - Word Format")

# Step 1: Read the Excel file into a pandas DataFrame
file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'SampleReleases.xlsx')
df = pd.read_excel(file_path)

# Step 2: Set up Ag-Grid options
grid_options_builder = GridOptionsBuilder.from_dataframe(df)

# Add checkboxes column for row selection (single selection mode with checkboxes)
grid_options_builder.configure_selection('single', use_checkbox=True)  
grid_options = grid_options_builder.build()

# Step 3: Display the table using AgGrid with selection enabled
st.subheader("📋 Select a Release for Validation")
response = AgGrid(df, gridOptions=grid_options, height=300, width='100%',
    fit_columns_on_grid_load=True)

# Step 5: Get the selected row
selected_rows = response.get('selected_rows', [])

# Step 6: Check if a row is selected
if isinstance(selected_rows, pd.DataFrame) and not selected_rows.empty:  # Ensure selected_rows is a non-empty DataFrame
    selected_row = selected_rows.iloc[0]  # Get the first (and only) selected row from the DataFrame

    # st.write("### Selected Row Data:")
    # st.write(selected_row)  # Debugging the selected row

    # Step 7: Extract values from the selected row to variables
    try:
        releaseID = selected_row[df.columns[0]]  # First column (Release ID)
        releaseName = selected_row[df.columns[1]]  # Second column (Release Name)
        projectID = selected_row[df.columns[2]]  # Third column (Project ID)
        projectName = selected_row[df.columns[3]]  # Fourth column (Project Name)
        appID = selected_row[df.columns[4]]  # Fifth column (App ID)
        appName = selected_row[df.columns[5]]  # Sixth column (App Name)

        # # Step 8: Display the selected values
        # st.write("### Selected Row Details:")
        # st.write(f"Release ID: {releaseID}")
        # st.write(f"Release Name: {releaseName}")
        # st.write(f"Project ID: {projectID}")
        # st.write(f"Project Name: {projectName}")
        # st.write(f"Application ID: {appID}")
        # st.write(f"App Name: {appName}")

    except KeyError as e:
        st.write(f"Error: Column {e} is missing in the selected row.")
else:
    st.write("No row selected. Please select a row.")

# File uploader for DOCX and Excel files
docx_file = st.file_uploader("Upload Word Document (DOCX)", type="docx")
excel_file = st.file_uploader("Upload Excel Configuration File", type="xlsx")

# Dropdown for selecting the sheet name from the Excel file
if excel_file:
    xls = pd.ExcelFile(excel_file)
    sheet_names = xls.sheet_names
    sheet_name = st.selectbox("Select the sheet name", sheet_names)

# Button to trigger validation
if docx_file and excel_file and sheet_name:
    validate_button = st.button("Validate Document")

    if validate_button:
        with st.spinner("Validating document..."):
            # Ensure temp directory exists
            temp_dir = os.getcwd()  # Get the exact current working directory
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)  # Create the directory if it doesn't exist

            # Debugging the current working directory
            # st.write(f"Debug: Current working directory is: {temp_dir}")

            # Save uploaded files temporarily
            docx_path = os.path.join(temp_dir, docx_file.name)
            excel_path = os.path.join(temp_dir, excel_file.name)

            # Debugging the paths to check if they are correct
            # st.write(f"Debug: DOCX file will be saved to: {docx_path}")
            # st.write(f"Debug: Excel file will be saved to: {excel_path}")
            
            # Write the uploaded files to the disk
            with open(docx_path, "wb") as f:
                f.write(docx_file.getbuffer())
            
            with open(excel_path, "wb") as f:
                f.write(excel_file.getbuffer())

            # Call validation function
            validate_document(docx_path, excel_path)

    