import streamlit as st
from docx import Document
import pandas as pd
import io

def clean_text(text):
    """Utility function to clean text by removing unwanted characters and trimming."""
    return text.strip().replace("“", "").replace("”", "").replace("\"", "")

def read_docx(file):
    """Read a .docx file and extract data structured by sections or pages."""
    doc = Document(file)
    data = []
    current_section = {}
    section_headers = set()

    # Process each table in the document
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = clean_text(cell.text)
                # Splitting entries that appear to be key-value based
                if ':' in text:
                    key, value = map(clean_text, text.split(':', 1))
                else:
                    key, value = text, None

                if key in section_headers:
                    # If key is repeated, start a new section
                    if current_section:
                        data.append(current_section)
                    current_section = {}
                    section_headers.clear()

                current_section[key] = value if value else key
                section_headers.add(key)

        # Append the last section if not empty
        if current_section:
            data.append(current_section)
            current_section = {}  # Reset for the next table

    return data

def save_to_excel(data, filename="output.xlsx"):
    """Convert list of dictionaries to an Excel file and save to a file-like object."""
    if data:
        df = pd.DataFrame(data)
        excel_file = io.BytesIO()
        df.to_excel(excel_file, index=False, engine='openpyxl')
        excel_file.seek(0)
        return excel_file
    return None

# Streamlit user interface
st.title('ISR Document to Excel Converter')
st.write('Upload your ISR DOCX file and convert its content to an Excel file, organized by distinct sections.')

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
