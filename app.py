import streamlit as st
from docx import Document
import pandas as pd
import io

def clean_text(text):
    """Utility function to clean text by removing unwanted characters and trimming."""
    return text.strip().replace("{", "").replace("}", "").replace("“", "").replace("”", "").replace("\"", "")

def read_docx(file):
    """Read a .docx file and extract data structured by tables, focusing on specific placeholders."""
    doc = Document(file)
    data = []
    
    # Process each table in the document
    for table in doc.tables:
        current_data = {}
        headers = [clean_text(cell.text) for cell in table.rows[0].cells]  # Assume first row is headers

        for row in table.rows[1:]:  # Skip header row
            for idx, cell in enumerate(row.cells):
                if idx >= len(headers):  # Avoid index error if there are more cells than headers
                    continue
                key = headers[idx]
                value_list = [clean_text(x) for x in cell.text.split('\n')]  # Split values by new lines
                # Append values or create a list of values under each header
                if key in current_data:
                    current_data[key].extend(value_list)
                else:
                    current_data[key] = value_list

        # Append the dictionary of the current table to the data list
        if current_data:
            data.append(current_data)

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
st.write('Upload your ISR DOCX file and convert its content to an Excel file, with each table as a new row.')

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
