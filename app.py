import streamlit as st
from docx import Document
import pandas as pd
import io

def clean_text(text):
    """Utility function to clean text by removing unwanted characters."""
    return text.replace("â€œ", "").replace("â€", "").replace("\"", "").strip()

def read_docx(file):
    """Read a .docx file from a file-like object and return contents as a structured dictionary."""
    doc = Document(file)
    data = {}
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = clean_text(cell.text)
                # Split on newlines to separate multiple entries in a single cell, if applicable
                entries = text.split('\n')
                for entry in entries:
                    key, value = entry.split(':', 1) if ':' in entry else (entry, '')
                    key = key.strip()
                    value = value.strip()
                    if key in data:
                        data[key].append(value)
                    else:
                        data[key] = [value]
    return data

def save_to_excel(data, filename="output.xlsx"):
    """Convert dictionary of lists to an Excel file and save to a file-like object."""
    if data:
        df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in data.items()]))
        excel_file = io.BytesIO()
        df.to_excel(excel_file, index=False, engine='openpyxl')
        excel_file.seek(0)
        return excel_file
    return None

# Streamlit user interface
st.title('ISR Document to Excel Converter')
st.write('Upload your ISR DOCX file and convert its content to an Excel file, organized by columns.')

uploaded_file = st.file_uploader("Choose a DOCX file", type="docx")

if uploaded_file is not None:
    with st.spinner('Processing...'):
        try:
            file_data = read_docx(uploaded_file)
            if file_data:
                excel_file = save_to_excel(file_data)
                if excel_file:
                    st.success('Conversion successful! Download your Excel file below.')
                    st.download_button(label="Download Excel",
                                       data=excel_file,
                                       file_name="processed_data.xlsx",
                                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                else:
                    st.error('No data could be extracted and converted to Excel.')
            else:
                st.error('The document appears to be empty or the format is not recognized.')
        
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
