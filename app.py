import streamlit as st
from docx import Document
import pandas as pd
import io

def clean_text(text):
    """Utility function to clean text by removing unwanted characters and trimming."""
    return text.strip().replace("“", "").replace("”", "").replace("\"", "")

def placeholder_filter(text):
    """Check if the text contains placeholder and extract the content within '{}'."""
    if '{' in text and '}' in text:
        # Extracting text within the first pair of '{}'
        start = text.find('{') + 1
        end = text.find('}', start)
        return text[start:end].strip()
    return None

def read_docx(file):
    """Read a .docx file and extract data structured by tables, focusing on specific placeholders."""
    doc = Document(file)
    data = []
    
    # Process each table in the document
    for table in doc.tables:
        current_data = {}
        headers = []

        for row in table.rows:
            for idx, cell in enumerate(row.cells):
                extracted_text = placeholder_filter(cell.text)
                if extracted_text:
                    if idx >= len(headers):
                        headers.append(extracted_text)  # Treat extracted placeholders as headers if they fit
                    else:
                        # Assuming each row after headers is data corresponding to those headers
                        if headers[idx] in current_data:
                            current_data[headers[idx]].append(extracted_text)
                        else:
                            current_data[headers[idx]] = [extracted_text]

        if current_data:  # If data has been collected for current table
            flattened_data = {k: ' '.join(v) for k, v in current_data.items()}  # Flatten lists to strings
            data.append(flattened_data)

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
st.write('Upload your ISR DOCX file and convert its content to an Excel file, focusing on placeholders within {}.')

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
