# Filename: app.py

import streamlit as st
import pandas as pd
from docx import Document
import os

def read_docx(file_path):
    """Read a .docx file from a path and return contents as a structured list."""
    doc = Document(file_path)
    data = []
    for table in doc.tables:
        headers = []
        row_data = {}
        for row in table.rows:
            for cell_index, cell in enumerate(row.cells):
                text = cell.text.strip()
                if text.startswith("“") and text.endswith("”"):  # Identify headers
                    headers.append(text.strip("“”"))
                elif headers and cell_index < len(headers):
                    row_data[headers[cell_index]] = text
            if row_data:
                data.append(row_data.copy())  # Append the copy of the row_data to data
                row_data.clear()  # Clear row_data for the next set of data
    return data

def save_to_csv(data, filename, folder='processed_files'):
    """Convert list of dictionaries to CSV and save to file."""
    if not os.path.exists(folder):
        os.makedirs(folder)
    df = pd.DataFrame(data)
    path = os.path.join(folder, filename)
    df.to_csv(path, index=False)
    return path

def process_template_and_download():
    file_path = 'path_to_your_ISR_template.docx'  # Path to your ISR template
    data = read_docx(file_path)
    if data:
        csv_full_name = "Test_Data_Register.csv"
        result_path = save_to_csv(data, csv_full_name)
        st.success(f'CSV file created successfully at {result_path}')
        return result_path
    else:
        st.error("No data extracted from the document.")
        return None

st.title('ISR Template Processor')

if st.button('Process ISR Template'):
    result_path = process_template_and_download()
    if result_path:
        with open(result_path, "rb") as file:
            st.download_button(label="Download CSV", data=file, file_name=result_path)
