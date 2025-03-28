import datetime
import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
import xml.etree.ElementTree as ET
from io import BytesIO
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import re
from datetime import datetime  # Correct import

# ✅ Set Streamlit to Full-Width Mode
st.set_page_config(layout="wide", page_title="PPT Validation App", page_icon="📊")

# ✅ Header Bar with Logo & Title
st.markdown(
    """
    <style>
        .header-bar {
            display: flex;
            align-items: center;
            background-color: #f8f9fa;
            padding: 10px;
            border-radius: 8px;
        }
        .header-bar img {
            width: 50px;
            margin-right: 15px;
        }
        .header-bar h1 {
            font-size: 24px;
            margin: 0;
        }
    </style>
    <div class="header-bar">
        <img alt="Logo">
        <h1>PPT Validation App</h1>
    </div>
    """,
    unsafe_allow_html=True
)

# ✅ Sidebar with Navigation
st.sidebar.title("📌 Navigation")
page = st.sidebar.radio("Go to:", ["🏠 Home", "📊 PPT Validation"])


# ✅ Handle Navigation
if page == "🏠 Home":
    st.write("## 🏠 Welcome to the PPT Validation App")
    st.write("Use the sidebar to navigate to different sections of the app.")
elif page == "📊 PPT Validation":
    st.write("### Select a Release for Validation")

SAMPLE_RELEASES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'SampleReleases.xlsx')


# Load existing sample releases
def load_sample_releases():
    if os.path.exists(SAMPLE_RELEASES_FILE):
        return pd.read_excel(SAMPLE_RELEASES_FILE)
    else:
        st.error("SampleReleases.xlsx not found. Please place the file in the correct location.")
        return pd.DataFrame()

# Extract text from named shapes in a slide
def extract_named_shapes(zip_path, slide_number):
    shape_texts = {}
    slide_file = f"ppt/slides/slide{slide_number}.xml"

    with zipfile.ZipFile(zip_path, "r") as pptx_zip:
        if slide_file in pptx_zip.namelist():
            with pptx_zip.open(slide_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
                      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

                for sp in root.findall(".//p:sp", namespaces=ns):
                    name_elem = sp.find(".//p:nvSpPr/p:cNvPr", namespaces=ns)
                    if name_elem is not None and "name" in name_elem.attrib:
                        shape_name = name_elem.attrib["name"]
                        text_elem = sp.findall(".//a:t", namespaces=ns)
                        text_content = " ".join([t.text for t in text_elem if t.text])
                        shape_texts[shape_name] = text_content

    return shape_texts

# Check if embedded Excel files exist
def check_embedded_excel(zip_path):
    with zipfile.ZipFile(zip_path, "r") as pptx_zip:
        return any(f.startswith("ppt/embeddings/") and f.endswith(".xlsx") for f in pptx_zip.namelist())


def extract_tables_from_slide(zip_path, slide_number):
    """
    Extracts tables from a given slide in the PowerPoint (.pptx) file.

    Args:
        zip_path (str): Path to the PPTX file (as a zip archive).
        slide_number (int): The slide number to extract tables from.

    Returns:
        list: A list of tables, where each table is a list of rows, and each row is a list of cell values.
    """
    tables = []
    
    slide_path = f"ppt/slides/slide{slide_number}.xml"  # Locate the slide XML file
    
    with zipfile.ZipFile(zip_path, 'r') as pptx:
        if slide_path not in pptx.namelist():
            return tables  # If slide XML is missing, return an empty list
        
        slide_xml = pptx.read(slide_path)
        root = ET.fromstring(slide_xml)

        # Define namespaces to search for table elements
        ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
              'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}

        # Find all tables in the slide
        for table in root.findall(".//a:tbl", ns):
            extracted_table = []
            
            # Find all rows in the table
            for row in table.findall(".//a:tr", ns):
                extracted_row = []
                
                # Find all cells in the row
                for cell in row.findall(".//a:tc", ns):
                    # Extract text from each cell
                    text_elem = cell.find(".//a:t", ns)
                    extracted_row.append(text_elem.text.strip() if text_elem is not None else "")
                
                extracted_table.append(extracted_row)  # Add row to table
            
            tables.append(extracted_table)  # Add table to list of tables
    
    return tables

def extract_embedded_files(zip_path, slide_number, output_dir="embedded_files"):
    """
    Extracts embedded files (Excel, CSV, etc.) from a specific slide in a PowerPoint file.

    :param zip_path: Path to the PPTX zip archive.
    :param slide_number: The slide number to check for embedded files.
    :param output_dir: Directory to store extracted files.
    :return: List of extracted file paths.
    """
    extracted_files = []
    os.makedirs(output_dir, exist_ok=True)  # Ensure directory exists

    with zipfile.ZipFile(zip_path, 'r') as pptx_zip:
        # Extract ALL embedded files from ppt/embeddings/
        for file_name in pptx_zip.namelist():
            if file_name.startswith("ppt/embeddings/"):  # Could be .xlsx, .csv, .bin
                extracted_path = os.path.join(output_dir, os.path.basename(file_name))
                with pptx_zip.open(file_name) as source, open(extracted_path, "wb") as target:
                    target.write(source.read())
                extracted_files.append(extracted_path.lower().strip())

        # Check slide-specific relationships for embedded files
        slide_rels_path = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
        slide_embedded_files = []

        if slide_rels_path in pptx_zip.namelist():
            with pptx_zip.open(slide_rels_path) as rels_file:
                rels_content = rels_file.read().decode("utf-8")

                # Find all embedded references (may be .xlsx, .bin, .csv)
                embedded_refs = re.findall(r'Target="(../embeddings/[^"]+)"', rels_content)
                for ref in embedded_refs:
                    embedded_filename = os.path.basename(ref)
                    matched_file = os.path.normpath(os.path.join(output_dir, embedded_filename)).lower().strip()  # ✅ Normalize path

                    # print(f"🔍 Checking Embedded File: {embedded_filename}")  
                    # print(f"➡ Matched Path: {matched_file}")  
                    # print(f"✅ Extracted Files: {extracted_files}")  

                    # Compare after ensuring lowercase + consistent path format
                    if matched_file in extracted_files:
                        # print("✅ Match Found! Adding to results.")  
                        slide_embedded_files.append(matched_file)

    # print(slide_embedded_files)
    return slide_embedded_files if slide_embedded_files else extracted_files

def get_total_slides(pptx_path):
    """Extracts the total number of slides from a PowerPoint file."""
    with zipfile.ZipFile(pptx_path, 'r') as pptx_zip:
        slide_files = [f for f in pptx_zip.namelist() if f.startswith("ppt/slides/slide") and f.endswith(".xml")]
        return len(slide_files)
    

# Function to get a slide's display name based on its extracted title
def get_slide_display_name(slide_number, slide_shapes):
    """Extract slide title and format slide name dynamically"""
    default_names = {1: "Title Page", 2: "Observations Slide"}  # Custom names for Slide 1 & 2
    extracted_title = slide_shapes.get("Title", "").strip()  # Extract the title text

    if slide_number in default_names:
        return f"Slide {slide_number} - {default_names[slide_number]}"
    elif extracted_title:
        return f"Slide {slide_number} - {extracted_title}"  # Use extracted title
    else:
        return f"Slide {slide_number}"  # Default fallback if no title
    
# Validate PowerPoint against selected row
def validate_ppt(zip_path, checklist_row):
    total_slides = get_total_slides(zip_path)
    results = {}

    # Extract named shapes from Slide 1
    # Slide 1 Validation
    slide1_shapes = extract_named_shapes(zip_path, 1)

    # Define expected fields and their corresponding expected values from checklist_row
    slide1_required_fields = {
        "Slide1ProjectName": "Project Name",
        "Slide1ProjectID": "Project ID",
        "Slide1AppID": "Application ID",
        "Slide1ApplicationName": "Application Name",
        "Slide1ReleaseName": "Release Name"
    }

    slide1_results = {}

    # Compare extracted vs. expected values
    for shape_name, field_name in slide1_required_fields.items():
        expected_value = checklist_row.get(field_name, "").strip()
        extracted_value = slide1_shapes.get(shape_name, "").split(":")[-1].strip() if shape_name in slide1_shapes else None
        
        if extracted_value is None:
            slide1_results[field_name] = f"🚫 Missing (Expected: {expected_value})"
        elif extracted_value.lower() == expected_value.lower():
            slide1_results[field_name] = "✅ Matched"
        else:
            slide1_results[field_name] = f"❌ Not Matched (Expected: {expected_value}, Found: {extracted_value})"

    # Store validation results
    results["Slide 1"] = slide1_results

    # Slide 2 Validation
    slide2_shapes = extract_named_shapes(zip_path, 2)
    slide2_tables = extract_tables_from_slide(zip_path, 2)
    embedded_files = extract_embedded_files(zip_path, 2)

    # Fetch Project ID & Release ID from checklist
    project_name = checklist_row.get("Project Name", "").strip().lower()
    release_id = checklist_row.get("Release ID", "").strip().lower()

    # ✅ Validate Slide2Title (Check if Project ID is present)
    # ✅ Extract Slide 2 Title & Convert to Lowercase
    slide2_title_text = slide2_shapes.get("Slide2Header", "").strip().lower()
    project_name_lower = project_name.lower().strip()  # Normalize for comparison

    # ✅ Use Regex to Find "Project Y" Anywhere in the Title
    match = re.search(rf"\b{re.escape(project_name_lower)}\b", slide2_title_text, re.IGNORECASE)

    # ✅ If Project Name is Found in the Title, It’s Valid
    title_missing = match is None  # If match is None, it means Project Name was NOT found

    # print("Extracted Project Name Found:", match.group(0) if match else "Not Found")
    # print("Expected Project Name:", project_name_lower)
    # print("Title Validation Result:", "✅ Valid" if not title_missing else "❌ Missing Project Name")


    # ✅ Validate Slide2Summary (Check for both Project ID & Release ID)
    # ✅ Validate Slide2Summary (Check for Project Name & Release ID in any order)
    slide2_summary_text = slide2_shapes.get("Slide2Summary", "").strip().lower()
    summary_missing = []

    
    # 🔹 Directly check for Release ID in text (from config)
    release_pattern = re.escape(release_id.lower())  # Escape special characters if any
    release_match = re.search(fr"\b{release_pattern}\b", slide2_summary_text)
    print(slide2_summary_text)

    # ✅ Validate Release ID presence
    if release_match:
        extracted_release_id = release_id  # Since it's an exact match
    else:
        summary_missing.append(f"Release ID '{release_id.upper()}' Not Found")

    # 🔹 Directly check for Project Name in text (from config)
    project_pattern = re.escape(project_name.lower())  # Escape special characters if any
    project_match = re.search(fr"\b{project_pattern}\b", slide2_summary_text)

    # ✅ Validate Project Name presence
    if project_match:
        extracted_project_name = project_name  # Since it's an exact match
    else:
        project_name = project_name.title();
        summary_missing.append(f"Project Name '{project_name}' Not Found")

    # 🔹 Print Debug Information (Optional)
    print("Extracted Release ID:", release_id if release_match else "Not Found")
    print("Extracted Project Name:", project_name if project_match else "Not Found")
    print("Validation Summary:", summary_missing if summary_missing else "✅ Valid")

    # ✅ Validate Table (Ensure at least one row contains "Load" or "Endurance" in first column)
    table_valid = False
    date_row_valid = False

    for table in slide2_tables:
        for row_index, row in enumerate(table):
            if row_index == 0:
                continue  # Skip header row

            first_column_text = row[0].strip().lower() if row and row[0] else ""
            second_column_text = str(row[1]).strip() if len(row) > 1 else ""
            third_column_text = str(row[2]).strip() if len(row) > 2 else ""

            # ✅ Condition 1: Check if first column contains "Load" or "Endurance"
            if first_column_text.lower() in ["load test", "endurance test", "load", "endurance"]:
                table_valid = True

            date_row_valid = False  # 🔹 Reset before validation
            # ✅ Condition 2: Ensure both second & third columns contain valid dates
            if len(second_column_text)>0 and len(third_column_text)>0:
                date_row_valid = True
                # try:
                #     datetime.strptime(second_column_text, "%d/%m/%Y")  # Adjust format as needed
                #     datetime.strptime(third_column_text, "%d/%m/%Y")
                #     date_row_valid = True
                #     print("Dates are present")
                # except ValueError:
                #     date_row_valid = False  # If parsing fails, mark it invalid
            # ✅ If both conditions met, exit loop early
            if table_valid and date_row_valid:
                break

        if table_valid and date_row_valid:
            break

    # ✅ Final Validation Result with Detailed Messages
    if table_valid and date_row_valid:
        table_validation_result = "✅ Valid"
    elif table_valid and not date_row_valid:
        table_validation_result = "❌ Found the Test Type, however, dates are missing."
    else:
        table_validation_result = "❌ Test Type is missing. Please validate and correct the Execution Details table."


    # ✅ Validate Embedded Excel File Presence
    has_embedded_excel = any(file.lower().endswith((".xlsm", ".xlsx", ".xls", ".csv")) for file in embedded_files)

    # ✅ Store validation results
    results["Slide 2"] = {
        "Title Validation": "✅ Valid" if not title_missing else "❌ Missing or Incorrect Project Name",
        "Summary Validation": "✅ Valid" if not summary_missing else f"❌  {', '.join(summary_missing)}",
        "Table Validation": table_validation_result,
        "Embedded Excel": "✅ Found" if has_embedded_excel else "❌ No Excel file found",
        # "Extracted Shapes": slide2_shapes
    }

    
    # Validate Slide 3 onwards for "Observations" shape
    for slide_number in range(3, total_slides+1):  # Check up to slide 10
        slide_shapes = extract_named_shapes(zip_path, slide_number)

        # Extract possible title and observation fields
        extracted_title = slide_shapes.get("Title", "").strip()
        extracted_observations = slide_shapes.get("Observations", "").strip()

        results[f"Slide {slide_number}"] = {
            "Title Found": "✅ Yes" if extracted_title else "❌ No",
            "Observations Found": "✅ Yes" if extracted_observations else "❌ No",
            # "Extracted Shapes": slide_shapes
        }

    return results

# Generate validation report in Excel
def generate_excel_report(validation_results):
    output = BytesIO()  # ✅ Create BytesIO buffer

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if validation_results and len(validation_results) > 0:  
            for slide, result in validation_results.items():
                df = pd.DataFrame.from_dict(result, orient='index', columns=["Validation Result"])
                df.to_excel(writer, sheet_name=slide)
        else:
            # ✅ Ensure at least one sheet is present
            df = pd.DataFrame([["No validation results found"]], columns=["Message"])
            df.to_excel(writer, sheet_name="Summary")

        writer.book.active = 0  # ✅ Ensure the first sheet is active

    writer.close()  # ✅ Explicitly close the writer

    output.seek(0)  # ✅ Reset buffer position
    return output

# Streamlit UI
st.title("📊 PPT Validation App with Sample Releases")

# Load Sample Releases
sample_releases_df = load_sample_releases()

# Display the table for selection
st.subheader("📋 Select a Release for Validation")
gb = GridOptionsBuilder.from_dataframe(sample_releases_df)
gb.configure_selection('single', use_checkbox=True)
grid_options = gb.build()

grid_response = AgGrid(
    sample_releases_df,
    gridOptions=grid_options,
    update_mode=GridUpdateMode.VALUE_CHANGED | GridUpdateMode.SELECTION_CHANGED,
    height=300,
    fit_columns_on_grid_load=True
)

# File Upload Section
st.subheader("📂 Upload PowerPoint File")
uploaded_ppt = st.file_uploader("Upload PPTX File", type=["pptx"])
# print(uploaded_ppt)

# Button to trigger validation
validation_results = None
selected_rows = grid_response.get('selected_rows', [])

# if isinstance(selected_rows, pd.DataFrame) and not selected_rows.empty:
if uploaded_ppt is not None and isinstance(selected_rows, pd.DataFrame) and not selected_rows.empty:
    selected_row_data = selected_rows.iloc[0]

    if st.button("✅ Validate PPT"):
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_ppt:
            tmp_ppt.write(uploaded_ppt.read())
            tmp_ppt_path = tmp_ppt.name

        # Run validation
        validation_results = validate_ppt(tmp_ppt_path, selected_row_data)

        # Clean up temp file
        os.remove(tmp_ppt_path)

        # Display results
        st.subheader("✅ Validation Results")
        for slide, result in validation_results.items():
            # Extract the slide title from validation results
            extracted_title = result.get("Extracted Shapes", {}).get("Title", "").strip()

            # Assign a custom name for Slide 1 and Slide 2
            default_names = {
                "Slide 1": "Title Page",
                "Slide 2": "Observations Slide"
            }

            # Determine the final display name
            if slide in default_names:
                slide_name = f"{slide} - {default_names[slide]}"
            elif extracted_title:
                slide_name = f"{slide} - {extracted_title}"
            else:
                slide_name = slide  # Fallback if no title is found

            # Display the updated slide name
            st.write(f"### {slide_name}")

            for key, value in result.items():
                st.write(f"**{key}:** {value}")

# Generate & Download Excel Report
if validation_results:
    excel_data = generate_excel_report(validation_results)
    st.download_button(
        label="📥 Download Validation Report",
        data=excel_data,
        file_name="PPT_Validation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
