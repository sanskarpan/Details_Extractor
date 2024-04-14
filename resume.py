import streamlit as st
import pandas as pd
import docx2txt
import fitz  # PyMuPDF
import re
from io import BytesIO
import base64
import tempfile

def extract_info(doc_file):
    if doc_file.name.endswith('.docx'):
        text = docx2txt.process(doc_file)
    elif doc_file.name.endswith('.pdf'):
        text = ''
        with fitz.open(stream=doc_file.read(), filetype="pdf") as pdf:
            for page in pdf:
                text += page.get_text()
    else:
        st.error("Unsupported file format. Please upload a DOCX or PDF file.")
        return None
    
    email = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
    phone = re.findall(r'\b(?:\+?\d{1,3}[-.\s]?)?\(?(?:\d{2,3})\)?[-.\s]?\d{3,4}[-.\s]?\d{4}\b', text)
    phone = phone if phone else ['-']  # If phone list is empty, replace with single dash "-"
    return {
        'Email': email,
        'Contact Number': phone,
        'Text': text
    }

def main():
    st.title("CV Information Extractor")

    uploaded_files = st.file_uploader("Upload CVs", type=['docx', 'pdf'], accept_multiple_files=True)
    if uploaded_files:
        all_info = []
        for uploaded_file in uploaded_files:
            info = extract_info(uploaded_file)
            if info is not None:
                all_info.append(info)
        
        if all_info:
            st.write("Extracted Information:")
            df = pd.DataFrame.from_records(all_info)
            st.write(df)

            if st.button('Export to Excel'):
                with tempfile.NamedTemporaryFile(delete=False) as temp:
                    temp_name = temp.name + ".xlsx"
                    with pd.ExcelWriter(temp_name, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='CV_Info')
                
                with open(temp_name, 'rb') as f:
                    data = f.read()
                    b64 = base64.b64encode(data).decode()
                    href = f'<a href="data:application/octet-stream;base64,{b64}" download="cv_info.xlsx">Download Excel File</a>'
                    st.markdown(href, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
