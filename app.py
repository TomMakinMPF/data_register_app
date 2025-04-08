import streamlit as st
from docx import Document
import pandas as pd
import io

def read_docx(file):
    """Read a .docx file from a file-like object and return contents as a structured list."""
    doc = Document(file)
    data = []
    headers_found = False
    headers = []

    for table in doc.tables:
        for i, row in enumerate(table.rows):
            text = (cell.text.strip() for cell in row.cells)
            if i == 0 or not headers_found:  # Assume the first row or the row after a header is found are headers
                headers = tuple(text)
                headers_found = True
            else:
                row_data = dict(zip(headers, text))
                if any(row_data.values()):  # Only add rows that have data
                    data.append(row_data)

    return data

def save_to_csv(data, filename="output.csv"):
    """Convert list of dictionaries to CSV and save to a file-like object."""
    if data:
        df = pd.DataFrame(data)
        csv = df.to_csv(index=False).encode('utf-8')
        return csv
    return None

# Streamlit user interface
st.title('ISR Document to CSV Converter')
st.write('Upload your ISR DOCX file and convert its content to a CSV file.')

uploaded_file = st.file_uploader("Choose a DOCX file", type="docx")

if uploaded_file is not None:
    with st.spinner('Processing...'):
        try:
            # Read the document
            file_data = read_docx(uploaded_file)
            
            # Generate CSV from data
            if file_data:
                csv_file = save_to_csv(file_data)
                if csv_file:
                    st.success('Conversion successful! Download your CSV below.')
                    st.download_button(label="Download CSV",
                                       data=csv_file,
                                       file_name="processed_data.csv",
                                       mime='text/csv')
                else:
                    st.error('No data could be extracted and converted to CSV.')
            else:
                st.error('The document appears to be empty or the format is not recognized.')
        
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
