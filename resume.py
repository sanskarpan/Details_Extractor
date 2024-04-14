import streamlit as st
import pandas as pd
import docx2txt
from PyPDF2 import PdfReader
import re
from io import BytesIO

def extract_info(doc_file):
    if doc_file.name.endswith('.docx'):
        text = docx2txt.process(doc_file)
    elif doc_file.name.endswith('.pdf'):
        reader = PdfReader(doc_file)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    else:
        st.error("Unsupported file format. Please upload a DOCX or PDF file.")
        return None
    
    email = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
    phone = re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]', text)
    phone = phone if phone else ['-']  # If phone list is empty, replace with single dash "-"
    return {
        'Email': email,
        'Contact Number': phone,
        'Text': text
    }

def main():
    st.title("CV Information Extractor")

    uploaded_file = st.file_uploader("Upload CV", type=['docx', 'pdf'])
    if uploaded_file is not None:
        info = extract_info(uploaded_file)
        if info is not None:
            st.write("Extracted Information:")
            st.write(pd.DataFrame.from_dict([info]))

            if st.button('Export to Excel'):
                df = pd.DataFrame.from_dict([info])
                excel_file = df.to_excel(index=False)
                excel_bytes = excel_file.getvalue()
                b64 = base64.b64encode(excel_bytes).decode()
                href = f'<a href="data:file/xlsx;base64,{b64}" download="cv_info.xlsx">Download Excel File</a>'
                st.markdown(href, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
