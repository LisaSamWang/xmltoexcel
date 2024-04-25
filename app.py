import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO
import zipfile
import os
import shutil
from pyairtable import Table
# Set the configuration for the Streamlit app page
st.set_page_config(
    page_title="Convertique",  # Sets the browser tab title
    page_icon="📊",  # Sets the favicon to a bar chart emoji, you can use an image path instead
    layout="wide",  # Optional: Use the "wide" layout
    initial_sidebar_state="expanded",  # Optional: Start with an expanded sidebar
)

def parse_xml(xml_content):
    root = ET.fromstring(xml_content)
    def find_frequent_child(root):
        counts = {}
        for child in root:
            tag = child.tag
            if tag in counts:
                counts[tag] += 1
            else:
                counts[tag] = 1
        return max(counts, key=counts.get) if counts else None

    record_tag = find_frequent_child(root)
    records = []
    for elem in root.findall(f'.//{record_tag}'):
        record_data = {}
        for child in elem:
            record_data[child.tag] = child.text.strip() if child.text else ''
        records.append(record_data)
    return records

def to_excel(records, filename):
    df = pd.DataFrame(records)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue(), f"{filename}.xlsx"

def handle_files(files):
    processed_files = []
    for file in files:
        if file.name.endswith('.xml'):
            content = file.getvalue().decode('utf-8')
            records = parse_xml(content)
            excel_data, excel_name = to_excel(records, file.name.split('.')[0])
            processed_files.append((excel_name, excel_data))
    return processed_files

def handle_zip(file):
    with zipfile.ZipFile(file, 'r') as z:
        z.extractall("temp_xmls")
    files = [open(f"temp_xmls/{f}", 'rb') for f in os.listdir("temp_xmls") if f.endswith('.xml')]
    processed_files = handle_files(files)
    shutil.rmtree("temp_xmls")
    return processed_files

def create_zip(files):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as z:
        for file_name, data in files:
            z.writestr(file_name, data)
    return zip_buffer.getvalue()

def upload_to_airtable(api_key, base_id, table_name, files):
    table = Table(api_key, base_id, table_name)
    for file_name, data in files:
        df = pd.read_excel(BytesIO(data))
        records = df.to_dict('records')  # Convert DataFrame to list of dictionaries
        for record in records:
            table.create(record)
        st.success(f"Uploaded {file_name} to Airtable")

st.title('XML to Excel Converter with Airtable Publishing')

uploaded_files = st.file_uploader("Upload your XML files or a ZIP file containing XML files", accept_multiple_files=True, type=['xml', 'zip'])
api_key = st.text_input("Enter your Airtable API Key", type="password")
base_id = st.text_input("Enter your Airtable Base ID")
table_name = st.text_input("Enter your Airtable Table Name")

if uploaded_files:
    if len(uploaded_files) == 1 and uploaded_files[0].name.endswith('.zip'):
        processed_files = handle_zip(uploaded_files[0])
    else:
        processed_files = handle_files(uploaded_files)

    download_zip = st.checkbox("Download all as ZIP?")
    if download_zip:
        zip_data = create_zip(processed_files)
        st.download_button("Download All Excel Files as ZIP", zip_data, "output.zip", "application/zip")
    else:
        for file_name, data in processed_files:
            st.download_button(f"Download {file_name}", data, file_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if st.button("Publish to Airtable") and api_key and base_id and table_name:
        upload_to_airtable(api_key, base_id, table_name, processed_files)
