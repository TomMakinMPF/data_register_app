import streamlit as st
from docx import Document
import pandas as pd
import io

def clean_text(text):
    """Utility function to clean text by removing unwanted characters."""
    return text.replace("â€œ", "").replace("â€", "").replace("\"", "").strip()

def read_docx(file):
    """Read a .docx file from a file-like object and return contents as a structured list."""
    doc = Document(file)
    data = []
    headers = []
    headers_found = False

    for table in doc.tables:
        for i, row in enumerate(table.rows):
            row_data = []
            for cell in row.cells:
                text = clean_text(cell.text)
                if i == 0 or not headers_found:  # Assuming first row is headers
                    # Check if text is enclosed in quotes indicating a header
                    if text.startswith("“") and text.endswith("”"):
                        headers.append(clean_text(text))  # Clean and add to headers
                    else:
                        headers.append(text)
                else:
                    row_data.append(text)
            if i == 0 or not headers_found:
                headers_found = True
            else:
                if row_data:  # Only append if row_data is not empty
                    data.append(dict(zip(headers, row_data)))

    return data

def save_to_excel(data, filename="output.xlsx"):
    """Convert list of dictionaries to an Excel file and save to a file-like object."""
    if data:
        df = pd.DataFrame(data)
        # Using an in-memory bytes buffer to store Excel file
        excel_file = io.BytesIO()
        df.to_excel(excel_file, index=False, engine='xlsxwriter')
        excel_file.seek(0)  # Go to the start of the stream
        return excel_file
    return None

# Streamlit user interface
st.title('ISR Document to Excel Converter')
st.write('Upload your ISR DOCX file and convert its content to an Excel file.')

uploaded_file = st.file_uploader("Choose a DOCX file", type="docx")

if uploaded_file is not None:
    with st.spinner('Processing...'):
        try:
            # Read the document
            file_data = read_docx(uploaded_file)
            
            # Generate Excel from data
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
