import streamlit as st
from docx import Document
import pandas as pd
import io

def clean_text(text):
    """Utility function to clean text by removing unwanted characters and trimming."""
    return text.strip().replace("“", "").replace("”", "").replace("\"", "").replace("{", "").replace("}", "")

def extract_data_from_cell(cell):
    """Extracts text contained within curly braces, indicating placeholders."""
    import re
    pattern = r'\{([^}]*)\}'  # Pattern to find text within {}
    return re.findall(pattern, cell.text)

def read_docx(file):
    """Read a .docx file and extract data structured by tables, focusing on placeholders."""
    doc = Document(file)
    data = []
    headers_found = False
    headers = []
    
    for table in doc.tables:
        current_section = {}
        for row in table.rows:
            row_data = [extract_data_from_cell(cell) for cell in row.cells if extract_data_from_cell(cell)]
            if not headers_found and row_data:
                headers = [item for sublist in row_data for item in sublist]  # Flatten list
                headers_found = True
                continue
            if headers and row_data:
                for idx, cell_data in enumerate(row_data):
                    if idx < len(headers):
                        current_section[headers[idx]] = ' '.join(cell_data)  # Join data points into one string per cell
                
        if current_section:
            data.append(current_section)
        headers_found = False  # Reset headers for next table

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
st.write('Upload your ISR DOCX file and convert its content to an Excel file, organizing data by placeholders.')

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
