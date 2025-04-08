# Filename: app.py

import streamlit as st
import pandas as pd
from docx import Document
import os

def read_docx(file):
    """Read a .docx file and return contents as a structured list."""
    doc = Document(file)
    data = []
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            text = (cell.text for cell in row.cells)
            if i == 0:
                keys = tuple(text)
                continue
            row_data = dict(zip(keys, text))
            data.append(row_data)
    return data

def save_to_csv(data, filename, folder='processed_files'):
    """Convert list of dictionaries to CSV and save to file."""
    if not os.path.exists(folder):
        os.makedirs(folder)
    df = pd.DataFrame(data)
    path = os.path.join(folder, filename)
    df.to_csv(path, index=False)
    return path

st.title('Word Document to CSV Processor')
uploaded_file = st.file_uploader("Upload a DOCX file", type="docx")

if uploaded_file is not None:
    data = read_docx(uploaded_file)
    if data:
        result_path = save_to_csv(data, f"processed_{uploaded_file.name}.csv")
        st.success(f'CSV file created successfully at {result_path}')
        st.download_button(label="Download CSV", data=open(result_path, "rb"), file_name=result_path)
    else:
        st.error("No data extracted from the document.")
