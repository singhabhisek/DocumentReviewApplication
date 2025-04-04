from io import BytesIO
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

st.markdown(
    """
    <style>
        .block-container { padding-top: 0.5rem; } /* Reduce top padding */
    </style>
    """,
    unsafe_allow_html=True
)

# Apply custom CSS for styling
st.markdown(
    """
    <style>
        .custom-subheader {
            font-size: 22px !important;
            font-weight: bold;
            color: #333;
        }
    </style>
    """, 
    unsafe_allow_html=True
)

st.markdown(
        """
        <style>
        .streamlit-expanderHeader span p {
            font-size: 20px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

# Custom CSS to increase the expander label font size
st.markdown(
    """
    <style>
    details > summary {
        font-size: 20px; /* Adjust the font size as needed */
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("""
  <style>
     /* Streamlit class name of the div that holds the expander's title*/
    .css-q8sbsg p {
      font-size: 32px;
      color: red;
      }
    
     /* Streamlit class name of the div that holds the expander's text*/
    .css-nahz7x p {
      font-family: bariol;
      font-size: 20px;
      }
  </style>
""", unsafe_allow_html=True)
# Define the path for the config file (assumes it's in a "config" folder next to the script)
CONFIG_FOLDER = os.path.join(os.getcwd(), "config")
CONFIG_FILE = os.path.join(CONFIG_FOLDER, "config.xlsx")
SHEET_NAME = "performance_testing_strategy"  # Assuming a single sheet for all word documents
file_path = os.path.join(CONFIG_FOLDER,'SampleReleases.xlsx')
temp_dir = os.path.join(os.getcwd(), "temp")  # Create 'temp' folder path

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


# def extract_section_names(docx_path):
#     """Extract section names from the document."""
#     with zipfile.ZipFile(docx_path, "r") as docx_zip:
#         document_xml = docx_zip.read("word/document.xml")

#     root = ET.fromstring(document_xml)
#     namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    
#     section_names = []
#     paragraphs = root.findall(".//w:p", namespace)
    
#     for para in paragraphs:
#         texts = [node.text for node in para.findall(".//w:t", namespace) if node.text]
#         text = " ".join(texts).strip()
        
#         if text:
#             section_names.append(text)

#     return section_names


def extract_section_names(docx_path):
    """Extract section names (headings and bold text) from the document."""
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        document_xml = docx_zip.read("word/document.xml")

    root = ET.fromstring(document_xml)
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    section_names = []
    paragraphs = root.findall(".//w:p", namespace)

    for para in paragraphs:
        # Extract all text inside this paragraph (even if not inside <w:r>)
        texts = [node.text.strip() for node in para.findall(".//w:t", namespace) if node.text]
        text = " ".join(texts).strip()  # Properly join words with spaces

        # Check if the paragraph has a heading style
        p_style = para.find(".//w:pPr/w:pStyle", namespace)
        is_heading = False
        if p_style is not None:
            style_val = p_style.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "")
            if "Heading" in style_val:  # Detects Heading1, Heading2, etc.
                is_heading = True

        # Check if any text is bold
        has_bold = any(r.find(".//w:rPr/w:b", namespace) is not None for r in para.findall(".//w:r", namespace))

        # Add text if it's a heading OR contains bold text
        if text and (is_heading or has_bold):
            section_names.append(text)

    # Remove duplicate or extra spaces
    section_names = list(set(name.replace("  ", " ") for name in section_names))

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
            # print (" ".join(texts))
            return " ".join(texts)  # Join text with spaces for better readability
         

def extract_key_values(text):
    """Extracts key-value pairs from the document text correctly."""
    key_value_pairs = {}

    # Define regex pattern with lookahead to stop at the next key
    patterns = {
        "Project Name": r"Project Name:\s*([^\n]+?)(?=\s+Release|$)",
        "Release": r"Release:\s*([^\n]+?)(?=\s+Project ID|$)",
        "Project ID": r"Project ID:\s*([^\n]+?)(?=\s+Enterprise Release ID|$)",
        "Enterprise Release ID": r"Enterprise Release ID:\s*([^\n]+?)(?=\s+Application Name|$)",
        "Application Name": r"Application Name:\s*([^\n]+?)(?=\s+Application ID|$)",
        "Application ID": r"Application ID\s*:\s*([^\n]+?)(?=\s+Document Change History|$)"
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
    return full_text[:500].strip()  # Extracts more characters for better accuracy


def read_config(config_path, sheet_name):
    """Reads the config Excel file and processes its key-value pairs correctly."""
    df = pd.read_excel(config_path, sheet_name=sheet_name, engine="openpyxl")

    # Debugging: Show column names
    print(f"🔍 Available columns in '{sheet_name}': {df.columns.tolist()}")

    # Ensure columns are correctly named
    expected_columns = ["Key", "Value"]
    df.columns = df.columns.str.strip()  # Trim whitespace
    if not all(col in df.columns for col in expected_columns):
        raise ValueError(f"❌ Excel sheet must contain columns: {expected_columns}. Found: {df.columns.tolist()}")

    # Convert 'Key' and 'Value' to a dictionary
    config_dict = {}
    
    for _, row in df.iterrows():
        key = row["Key"].strip()
        value = str(row["Value"]).strip()  # Convert to string and trim whitespace
        
        if key == "Sections":
            # Convert sections into a list
            config_dict[key] = [section.strip() for section in value.split(",")]
        else:
            # Store all other key-value pairs
            config_dict[key] = value

    return config_dict

def compare_values(extracted, config):
    """Compares extracted values with config file values."""
    for key, value in extracted.items():
        config_key = f"page1_{key}"  # Convert key to match config format
        config_value = config.get(config_key)

        # Handle missing keys in config
        if config_value is None:
            print(f"⚠️ WARNING: {key} not found in config!")
            continue

        # Normalize case and whitespace for comparison
        extracted_value = str(value).strip().lower()
        expected_value = str(config_value).strip().lower()

        # Handle list comparison (e.g., Sections)
        if isinstance(config_value, list):
            extracted_list = [item.strip().lower() for item in extracted_value.split(",")]
            expected_list = [item.strip().lower() for item in config_value]
            
            if sorted(extracted_list) != sorted(expected_list):
                print(f"❌ Mismatch in {key}: Extracted: {extracted_list}, Expected: {expected_list}")
            else:
                print(f"✅ Match: {key}")
        else:
            # Standard string comparison
            if extracted_value != expected_value:
                print(f"❌ Mismatch in {key}: Extracted: '{value}', Expected: '{config_value}'")
            else:
                print(f"✅ Match: {key}")

def validate_page1_key_values(docx_path, selected_row, config):
    """Validates key-value pairs from Page 1 text against the selected row data using configurable key validation."""

    # print(f"DEBUG: selected_row type = {type(selected_row)}, value = {selected_row}")

    # ✅ Convert selected_row to dictionary if it's a Pandas Series
    if isinstance(selected_row, pd.Series):
        selected_row = selected_row.to_dict()

    if not isinstance(selected_row, dict):
        raise TypeError(f"Expected selected_row to be a dictionary, but got {type(selected_row).__name__}")

    # print("Converted selected_row:", selected_row)

    # Extract and normalize required keys from config
    raw_mandatory_fields = config.get('Page1_MandatoryFieldsToValidate', [])

    # Ensure raw_mandatory_fields is always a list
    if isinstance(raw_mandatory_fields, str):
        raw_mandatory_fields = raw_mandatory_fields.split(',')
    elif not isinstance(raw_mandatory_fields, list):
        raise TypeError(f"Expected Page1_MandatoryFieldsToValidate to be a list or string, but got {type(raw_mandatory_fields).__name__}")

    # Normalize required keys
    required_keys = set(key.strip().lower().replace(" ", "") for key in raw_mandatory_fields)
    # print("✅ Required Keys:", required_keys)

    # print("Required Keys:", required_keys)

    # Extract text from Page 1 of the document
    page1_text = extract_page1_text(docx_path)
    
    # Extract key-value pairs from the Page 1 text
    extracted_values = extract_key_values(page1_text)

    # print("🔍 DEBUG: extracted_values type =", type(extracted_values), extracted_values)


    # Normalize keys for case-insensitive comparison, but keep original case for UI
    normalized_selected_row = {
        key: str(value).strip()
        for key, value in selected_row.items()
        if key.lower().replace(" ", "") in required_keys
    }

    # print("Normalized Selected Row:", normalized_selected_row)

    normalized_extracted = {
        key.strip(): value.strip()
        for key, value in extracted_values.items()
    }

    # print("Normalized Extracted Values:", normalized_extracted)

    # Ensure normalized_extracted is a dictionary
    if not isinstance(normalized_extracted, dict):
        print("🚨 ERROR: normalized_extracted is not a dictionary! It is:", type(normalized_extracted))
        normalized_extracted = {}
    # Compare extracted values with selected row values
    results = {}

    for key, expected_value in normalized_selected_row.items():
        found_value = normalized_extracted.get(key, "Missing")
        # print(">>>" + key)
        if found_value == "Missing":
            status, reason = "❌ Not Matched", "Key not found in document"
        elif expected_value.lower() == found_value.lower(): #or found_value.lower() in expected_value.lower():
            status, reason = "✅ Matched", "Values match"
        elif key =="Application ID":
            found_value_cleaned = found_value.replace(" ", "").lower()
            expected_value_cleaned = expected_value.replace(" ", "").lower()
    
            # If found_value is numeric, prefix it with "APP-"
            if found_value_cleaned.isnumeric():
                found_value_cleaned = f"appid-{found_value_cleaned}"

            # If expected_value_cleaned is numeric, prefix it with "APP-"
            if expected_value_cleaned.isnumeric():
                expected_value_cleaned = f"appid-{expected_value_cleaned}"
            
            if found_value_cleaned == expected_value_cleaned:
                status, reason = "✅ Matched", "Values match (Partial Match Allowed)"
        else:
            status, reason = "❌ Not Matched", "Value mismatch"

        # Store structured result
        results[key] = {
            "status": status,
            "found": found_value,
            "expected": expected_value,
            "reason": reason
        }

    # print(results)

    # Return structured dictionary
    return {
        "status": "✅ All matched" if all(r["status"] == "✅ Matched" for r in results.values()) else "❌ Mismatches found",
        "details": results  # Now this is a dictionary!
    }

    # # Determine overall validation status
    # overall_status = "✅ All matched" if all("✅ Matched" in r for r in results) else "❌ Mismatches found"
    # print(results)
    # return {
    #     "status": overall_status,
    #     "details": results  # List of formatted messages for UI
    # }


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
                output_path = os.path.join(temp_dir, os.path.basename(file))
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

import zipfile
import xml.etree.ElementTree as ET
import re

def extract_revision_history(docx_path):
    """Extracts the Document Revision History table, handling merged title rows correctly."""

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

    # Iterate through all text elements to locate "Document Change History and Management"
    for para in paragraphs:
        text = para.text.strip() if para.text else ""

        # Check if we found "Document Change History and Management"
        if "Document Change History and Management" in text:
            found_section = True
            continue

        # Once found, locate the next table <w:tbl>
        if found_section:
            for table in root.findall(".//w:tbl", namespaces):
                revision_table = table
                break  # Stop at the first table found after the heading

            if revision_table is not None:
                break  # Stop searching once the table is found

    # Handle case when no table is found
    if revision_table is None:
        print("❌ Document Revision History table not found!")
        return None

    # Extract rows from the table
    table_data = []
    rows = revision_table.findall(".//w:tr", namespaces)

    if len(rows) < 2:
        print("⚠️ Table does not have enough rows to extract data!")
        return None

    # Ignore the first row (merged title row) and take the second row as headers
    headers = []
    for cell in rows[1].findall(".//w:tc", namespaces):
        cell_text = " ".join(t.text.strip() for t in cell.findall(".//w:t", namespaces) if t.text)
        headers.append(cell_text)

    # Extract data rows
    for row in rows[2:]:  # Skip first (title) and second (header) rows
        row_data = []
        for cell in row.findall(".//w:tc", namespaces):
            cell_text = " ".join(t.text.strip() for t in cell.findall(".//w:t", namespaces) if t.text)
            cell_text = re.sub(r"\s*/\s*", "/", cell_text)  # Normalize date formatting
            row_data.append(cell_text)

        if any(row_data):  # Ignore empty rows
            table_data.append(dict(zip(headers, row_data)))  # Convert row into a dictionary

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
def validate_document(docx_path, config_file, sheet_name):
    """Validates the Word document using a config extracted from an Excel file."""
    config = pd.read_excel(config_file, sheet_name=sheet_name, engine="openpyxl").set_index("Key")["Value"].to_dict()
    
    if "Sections" in config:
        config["Sections"] = [s.strip() for s in str(config["Sections"]).split(",")]

    # Extract text, sections, and tables
    text_by_page = extract_text_by_page(docx_path)
    extracted_sections = extract_section_names(docx_path)
    tables = extract_table_content(docx_path)
    missing_sections, extra_sections = validate_sections_using_toc(docx_path, config.get("Sections", []))

    # Initialize results list to store validation outcomes
    validation_results = {
        "Section Validation": [],
        "Document Revision History": [],
        "Page 1 Summary Details": [],
        "Embedded Excel Validation": []
    }

    # Section 1: Section Validation
    results = []
    normalized_extracted = {section.strip().lower(): section for section in extracted_sections}
    normalized_configured = {section.strip().lower(): section for section in config.get("Sections", [])}

    for config_key, config_section in normalized_configured.items():
        if config_key in normalized_extracted:
            results.append(f"✅ {config_section}: Matched (Found in document)")
        else:
            results.append(f"❌ {config_section}: Not Found (Expected but missing)")

    extra_sections = [section for key, section in normalized_extracted.items() if key not in normalized_configured]
    # if extra_sections:
    #     results.append(f"⚠️ Extra Sections: {', '.join(extra_sections)} (Not in config)")
    
    validation_results["Section Validation"] = results

    # Section 2: Document Revision History
    today = datetime.today()
    one_week_ago = today - timedelta(days=7)
    revision_history = extract_revision_history(docx_path)

    def parse_revision_date(date_str):
        formats = [
            "%d-%B-%Y", "%d-%b-%Y", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d",
            "%Y/%m/%d", "%d.%m.%Y", "%A, %d %B %Y", "%d %B %Y", "%d-%m-%Y %H:%M:%S",
        ]
        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        return None

    revision_results = []
    if revision_history:
        for row in revision_history:
            revision_results.append(f"📄 Revision {row.get('Revision Number', 'N/A')}: Author = {row.get('Author', 'N/A')}, Date = {row.get('Revision Date', 'N/A')}")
        
        if len(revision_history) > 0:
            second_row = revision_history[0]
            author_exists = bool(second_row.get("Author", "").strip())
            revision_date_str = second_row.get("Revision Date", "").strip()
            try:
                revision_date = parse_revision_date(revision_date_str) if revision_date_str else None
                recent_date = revision_date and revision_date >= one_week_ago
            except ValueError:
                recent_date = False

            revision_results.append(f"✅ **Author Present:** {'Yes' if author_exists else '❌ No'}")
            revision_results.append(f"🗓️ **Recent Revision (within last 7 days):** {'✅ Yes' if recent_date else '❌ No'}")
        else:
            revision_results.append("⚠️ Not enough data to check revision history.")
    else:
        revision_results.append("❌ **Document Revision History table not found!**")
    
    validation_results["Document Revision History"] = revision_results

    # Section 3: Page 1 Details Validation
    page1_results = []

    page1_validation = validate_page1_key_values(docx_path, selected_row, config)
    

    # Iterate over each key in the expected config
    if "details" in page1_validation:
        for key, value in page1_validation["details"].items():
            status = value.get("status", "❌ Unknown Status")
            found = value.get("found", "N/A")
            expected = value.get("expected", "N/A")
            reason = value.get("reason", "No reason provided")

            page1_results.append(f"{status} {key}: Found '{found}', Expected '{expected}' → {reason}")

    # validation_results["Page 1 Details"] = page1_results

    # Store the detailed results in validation output
    validation_results["Page 1 Summary Details"] = page1_results

    # Section 4: Embedded Excel Check
    embedded_excels = extract_embedded_excel(docx_path)
    embedded_excel_results = []

    if embedded_excels:
        for excel_file in embedded_excels:
            extracted_data, matching_sheets = extract_excel_data_from_embedded(excel_file)
            if len(matching_sheets) >= 3:
                embedded_excel_results.append(f"✅ Embedded Excel File: {excel_file} contains the required sheets: {', '.join(matching_sheets)}")
                project_id = config.get("Project ID")
                release_id = config.get("Release ID")
                a2_value = extracted_data.get(matching_sheets[0], {}).get("A2")
                b8_value = extracted_data.get(matching_sheets[0], {}).get("B8")

                if a2_value == project_id:
                    embedded_excel_results.append(f"  - ✅ A2 matches the Project ID: {a2_value}")
                else:
                    embedded_excel_results.append(f"  - ❌ A2 does not match the Project ID: {a2_value}")

                if b8_value == release_id:
                    embedded_excel_results.append(f"  - ✅ B8 matches the Release ID: {b8_value}")
                else:
                    embedded_excel_results.append(f"  - ❌ B8 does not match the Release ID: {b8_value}")
                
                break
        else:
            embedded_excel_results.append("❌ Please check if you have attached the correct Non-Functional Requirement sheet template.")
    
    validation_results["Embedded Excel Validation"] = embedded_excel_results

    return validation_results

st.title("📑 Test Plan Validation Application - Word Format")

# Load the Excel file
# file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'SampleReleases.xlsx')
df = pd.read_excel(file_path)

st_col1, st_col2 = st.columns([0.8,0.2])
# Add a search bar for filtering
with st_col2:
    search_text = st.text_input("", placeholder="🔍 Search...")
    escaped_search_text = re.escape(search_text)  # Escape special regex characters

# Filter the DataFrame dynamically
if search_text:
    df = df[df.apply(lambda row: row.astype(str).str.contains(escaped_search_text, case=False, na=False).any(), axis=1)]

# Set up Ag-Grid options
grid_options_builder = GridOptionsBuilder.from_dataframe(df)
grid_options_builder.configure_selection('single', use_checkbox=True)  # Enable single-row selection
grid_options = grid_options_builder.build()

# Display table using AgGrid
st.subheader("📋 Select a Release for Validation")
response = AgGrid(df, gridOptions=grid_options, height=300, width='100%', fit_columns_on_grid_load=True)

# Extract the selected row
selected_rows = response.get('selected_rows', [])

# ✅ Ensure row selection is handled correctly
if isinstance(selected_rows, list) and selected_rows:  # Case 1: List of dictionaries
    selected_row = selected_rows[0]  # Extract first row as dictionary
elif isinstance(selected_rows, pd.DataFrame) and not selected_rows.empty:  # Case 2: DataFrame
    selected_row = selected_rows.iloc[0].to_dict()  # Convert first row to dictionary
else:
    selected_row = None  # No selection

# Process selected row
if selected_row:
    try:
        releaseID = selected_row.get(df.columns[0], "N/A")  
        releaseName = selected_row.get(df.columns[1], "N/A")  
        projectID = selected_row.get(df.columns[2], "N/A")  
        projectName = selected_row.get(df.columns[3], "N/A")  
        appID = selected_row.get(df.columns[4], "N/A")  
        appName = selected_row.get(df.columns[5], "N/A")  

        
    except Exception as e:
        st.error(f"⚠️ Error extracting row data: {e}")

else:
    st.warning("⚠️ No row selected. Please select a release.")

# File uploader for DOCX file
docx_file = st.file_uploader("📂 Upload Word Document (DOCX)", type="docx")

# Initialize validation status
validation_completed = False
validation_result = None

# Layout for Validate and Export buttons
col1, col2 = st.columns([0.8, 0.2])

with col1:
    validate_button = st.button("🚀 Validate Document", disabled=not docx_file)

# Initially disable the export button
export_button_disabled = True

# Initialize session state variables
if "validation_completed" not in st.session_state:
    st.session_state["validation_completed"] = False
if "export_clicked" not in st.session_state:
    st.session_state["export_clicked"] = False


if validate_button and docx_file:
    with st.spinner("🔍 Validating document... Please wait."):
        docx_path = os.path.join(temp_dir, docx_file.name)
        with open(docx_path, "wb") as f:
            f.write(docx_file.getbuffer())

        validation_result = validate_document(docx_path, CONFIG_FILE, SHEET_NAME)

        if validation_result:
            st.session_state["validation_completed"] = True  # Store state
            validation_completed = True
            st.toast("✅ Validation Completed!")
            st.write("### Validation Results:")
            for section, results in validation_result.items():
                st.write(f"#### {section}:")
                if isinstance(results, list):
                    for result in results:
                        st.write(f"- {result}")
                else:
                    st.write(f"- {results}")
            st.write("\n")

print(f"🔹 Validation completed state: {validation_completed}")


# ✅ Show Export button only after validation is completed
if st.session_state["validation_completed"]:
    if st.session_state.get("export_clicked"):
        print("🔹 Running export logic!")  

        if not validation_result:
            # st.warning("⚠️ Validation result is empty. Please run validation first.")
            print("⚠️ Warning: Validation result is empty.")  
        else:
            print("✅ Validation result exists, preparing export...")  

            # Convert validation results to a structured DataFrame
            report_data = []
            for section, results in validation_result.items():
                if isinstance(results, list):
                    for result in results:
                        report_data.append({"Section": section, "Result": result})
                else:
                    report_data.append({"Section": section, "Result": results})

            report_df = pd.DataFrame(report_data)

            # Convert DataFrame to Excel format
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                report_df.to_excel(writer, index=False, sheet_name="Validation Report")
            processed_data = output.getvalue()

            # Provide download button
            st.download_button(
                label="📥 Download Excel Report",
                data=processed_data,
                file_name="Validation_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="export_download"
            )

            st.success("✅ Report ready for download!")
