import streamlit as st
import openpyxl
from docx import Document
import os
import shutil
import tempfile
import zipfile

# Streamlit Configuration
st.set_page_config(
    page_title="MSAmagi",
    page_icon=":globe_with_meridians:",
    layout="centered",
    initial_sidebar_state="expanded",
)

TEMPLATE_PATH_BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'mallar')
RESULTS_DIRECTORY = tempfile.mkdtemp()

def handle_uploaded_zip(uploaded_zip):
    # Extract the uploaded ZIP to a temporary directory
    with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
        temp_dir = tempfile.mkdtemp()
        zip_ref.extractall(temp_dir)

    # Process MSA files in the temporary directory
    process_msa_files(temp_dir)

    # Clean up (optional)
    shutil.rmtree(temp_dir)

def zip_results_directory():
    # Create a separate directory for the zip file
    zip_dir = tempfile.mkdtemp()
    zip_filename = os.path.join(zip_dir, "results.zip")
    
    # Archive the RESULTS_DIRECTORY
    shutil.make_archive(zip_filename[:-4], 'zip', RESULTS_DIRECTORY)
    return zip_filename

def process_msa_files(msa_directory):
    # Get the list of files in the directory
    files_in_directory = os.listdir(msa_directory)

    # Print the number of files in the directory
    print(f"Antal filer i {msa_directory}: {len(files_in_directory)}")
    
    for filename in os.listdir(msa_directory):
        if filename.endswith(".xlsx"):
            msa_path = os.path.join(msa_directory, filename)
            msa_data = read_msa_info(msa_path)
            customer = identify_customer(msa_data['customer_name'], msa_data['cell_d4_value'])
            print_info(msa_path, msa_data, customer)
            if customer:
                update_word_template(customer, msa_data)
                update_prisberakningsmall(customer, msa_data['landlord_desc'], msa_data['tenant_desc'], msa_data)
            else:
                print(f"Invalid customer name in MSA file: {msa_path}")

def process_single_msa_file(uploaded_file_content, file_name):
    with open("temp_msa_file.xlsx", "wb") as f:
        f.write(uploaded_file_content)
    
    msa_data = read_msa_info("temp_msa_file.xlsx")
    customer = identify_customer(msa_data['customer_name'], msa_data['cell_d4_value'])
    print_info(file_name, msa_data, customer)
    
    if customer:
        print("Processing Avtalsmall...")  # Debugging line
        update_word_template(customer, msa_data)
        
        print("Processing Prisberakningsmall...")  # Debugging line
        update_prisberakningsmall(customer, msa_data['landlord_desc'], msa_data['tenant_desc'], msa_data)
    else:
        print(f"Invalid customer name in uploaded MSA file: {file_name}")

    os.remove("temp_msa_file.xlsx")


def read_msa_info(msa_path):
    wb = openpyxl.load_workbook(msa_path, data_only=True)
    sheet = wb.active
    customer_name = sheet["D3"].value
    cell_d4_value = sheet["D4"].value
    data = {
        'customer_name': customer_name,
        'tenant_desc': get_tenant_description(sheet, customer_name, cell_d4_value),
        'landlord_desc': get_landlord_description(sheet),
        'x_coord': get_coordinate(sheet["D10"].value, sheet["D9"].value),
        'y_coord': get_coordinate(sheet["D11"].value, sheet["D9"].value),
        'mast_info': sheet["E12"].value,
        'cell_d4_value': cell_d4_value
    }
    wb.close()
    return data

def get_tenant_description(sheet, customer_name, cell_d4_value):
    customer_name = customer_name.lower() if isinstance(customer_name, str) else str(customer_name)
    cell_d4_value = cell_d4_value.lower() if isinstance(cell_d4_value, str) else str(cell_d4_value)
    if 'telia' in customer_name or 'telia sverige ab' in cell_d4_value:
        tenant_desc_d6 = sheet["D6"].value
        if tenant_desc_d6:
            print(f"For Telia customer, tenant description chosen: {tenant_desc_d6}")
            return tenant_desc_d6
    return sheet["D8"].value

def get_landlord_description(sheet):
    landlord_desc_d4 = sheet["E9"].value
    if landlord_desc_d4 != "Hyresvärds uppgifter" and landlord_desc_d4:
        return landlord_desc_d4
    landlord_desc_e7 = sheet["E7"].value
    if landlord_desc_e7 != "Hyresvärds uppgifter" and landlord_desc_e7:
        return landlord_desc_e7
    return sheet["E9"].value or ""

def get_coordinate(coord, alternative_coord):
    return coord or alternative_coord

def identify_customer(customer_name, cell_d4_value):
    customer_name = customer_name.lower() if isinstance(customer_name, str) else str(customer_name)
    cell_d4_value = cell_d4_value.lower() if isinstance(cell_d4_value, str) else str(cell_d4_value)
    if '3gis' in customer_name:
        return '3GIS'
    elif 'hi3g' in customer_name or 'hi3g access ab' in customer_name:
        return 'Hi3G'
    elif 'telia' in customer_name or 'telia sverige ab' in cell_d4_value:
        return 'Telia'
    return None

def print_info(msa_path, msa_data, customer):
    print(f"Processing file: {msa_path}")
    for key, value in msa_data.items():
        print(f"{key.capitalize()}: {value}")
    print(f"Customer: {customer}")
    print(f"____________________")

def sanitize_file_name(name):
    sanitized_name = '_'.join(name.split())
    sanitized_name = sanitized_name.replace("/", "_")
    return sanitized_name

def replace_placeholders_in_document(doc, data):
    placeholders = {
        "--::HYRESVARD::--": data['landlord_desc'] or "",
        "--::HYRESGAST::--": data['tenant_desc'] or "",
        "--::x_coordinate::--": str(data['x_coord']) if data['x_coord'] else "",
        "--::y_coordinate::--": str(data['y_coord']) if data['y_coord'] else "",
        "--::masttyp_hojd::--": str(data['mast_info']) if data['mast_info'] else "",
    }

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for key, value in placeholders.items():
                if key in run.text:
                    run.text = run.text.replace(key, value)

def update_word_template(customer, msa_data):
    template_path = os.path.join(TEMPLATE_PATH_BASE, customer, f"Avtalsmall_{customer}.docx")

    if os.path.exists(template_path):
        doc = Document(template_path)
        replace_placeholders_in_document(doc, msa_data)

        result_folder = os.path.join(RESULTS_DIRECTORY, customer)
        os.makedirs(result_folder, exist_ok=True)

        sanitized_landlord_desc = sanitize_file_name(msa_data['landlord_desc'])
        sanitized_tenant_desc = sanitize_file_name(msa_data['tenant_desc'])
        
        result_file_name = f"Avtalsmall_{sanitized_landlord_desc}_{sanitized_tenant_desc}_{customer}.docx"
        result_path = os.path.join(RESULTS_DIRECTORY, customer, result_file_name)
        
        doc.save(result_path)
    else:
        print(f"The file {template_path} does NOT exist.")


def update_prisberakningsmall(customer, landlord_desc, tenant_desc, msa_data):
    template_path = os.path.join(TEMPLATE_PATH_BASE, customer, f"Prisberakningsmall_{customer}.xlsx")
    print(f"Attempting to update Prisberakningsmall with template at {template_path}")  # Debugging line

    if os.path.exists(template_path):
        print("Template exists. Proceeding to update...")  # Debugging line

        # Load the Excel workbook and sheet
        wb = openpyxl.load_workbook(template_path)
        sheet = wb.active

        # Assuming you need to replace placeholders in the Excel sheet as you did in the Word document
        for row in sheet.iter_rows():
            for cell in row:
                for key, value in msa_data.items():
                    placeholder = f"--::{key.upper()}::--"
                    if placeholder in str(cell.value or ""):
                        cell.value = cell.value.replace(placeholder, str(value))

        # Save the modified Excel workbook
        result_folder = os.path.join(RESULTS_DIRECTORY, customer)
        os.makedirs(result_folder, exist_ok=True)
        sanitized_landlord_desc = sanitize_file_name(landlord_desc)
        sanitized_tenant_desc = sanitize_file_name(tenant_desc)
        result_file_name = f"Prisberakningsmall_{sanitized_landlord_desc}_{sanitized_tenant_desc}_{customer}.xlsx"
        result_path = os.path.join(result_folder, result_file_name)
        wb.save(result_path)
        wb.close()
    else:
        print(f"The file {template_path} does NOT exist.")  # This will show if the path is incorrect or if the template does not exist

results_directory = RESULTS_DIRECTORY  

choice = st.radio("Önskar du hantera en enskild MSA eller en hel folder?", ["Enskild MSA", "Hel folder"])

# Initialize and set session state variables for msa_directory, results_directory and templates_directory if they don't exist.
if 'msa_directory' not in st.session_state:
    st.session_state.msa_directory = RESULTS_DIRECTORY

if 'templates_directory' not in st.session_state:
    st.session_state.templates_directory = TEMPLATE_PATH_BASE

if choice == "Enskild MSA":
    uploaded_file = st.file_uploader("Ladda upp MSA nedan.", type=["xlsx"])
    if uploaded_file:
        process_single_msa_file(uploaded_file.getvalue(), uploaded_file.name)
        st.write("MSA hanterad.")

elif choice == "Hel folder":
    # Use session state variables as default values for the text inputs
    msa_directory = st.text_input("Ange filsökvägen till MSA-filer:", value=st.session_state.msa_directory)
    templates_directory = st.text_input("Ange filsökväg där mallar finns:", value=st.session_state.templates_directory)

    # Update session state values based on user input
    st.session_state.msa_directory = msa_directory
    st.session_state.templates_directory = templates_directory

    if st.button("Hantera filer"):
        if msa_directory and results_directory and templates_directory:
            process_msa_files(msa_directory)
            st.write("Alla MSA-filer i vald folder har nu blivit hanterade.")
            zip_file = zip_results_directory()

            with open(zip_file, "rb") as f:
                bytes_data = f.read()
                st.download_button(
                    label="Ladda hem hanterade avtalsmallar och prisberakningsmallar",
                    data=bytes_data,
                    file_name="results.zip",
                    mime="application/zip"
                )

st.stop()