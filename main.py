import os
import streamlit as st
import pandas as pd
from deep_translator import GoogleTranslator
from docx import Document
import base64
import io
from openpyxl import load_workbook
from openpyxl.styles import Font

def translate_text(text, target_language="id"):
    translated_text = GoogleTranslator(source='auto', target=target_language).translate(text)
    return translated_text

def save_translated_file(translated_text, file_name):
    if not os.path.exists("done"):
        os.makedirs("done")
    
    file_path = os.path.join("done", file_name)
    if os.path.exists(file_path):
        file_name, file_extension = os.path.splitext(file_name)
        count = 1
        while os.path.exists(os.path.join("done", f"{file_name}_{count}{file_extension}")):
            count += 1
        file_path = os.path.join("done", f"{file_name}_{count}{file_extension}")
    
    with open(file_path, "w", encoding="utf-8") as file:
        file.write(translated_text)
    return file_path

def save_excel_file(translated_df, file_name):
    if not os.path.exists("done"):
        os.makedirs("done")
    
    file_path = os.path.join("done", file_name)
    if os.path.exists(file_path):
        file_name, file_extension = os.path.splitext(file_name)
        count = 1
        while os.path.exists(os.path.join("done", f"{file_name}_{count}{file_extension}")):
            count += 1
        file_path = os.path.join("done", f"{file_name}_{count}{file_extension}")
    
    translated_df.to_excel(file_path, index=False)
    return file_path

def save_docx_file(translated_doc, file_name):
    if not os.path.exists("done"):
        os.makedirs("done")
    
    file_path = os.path.join("done", file_name)
    if os.path.exists(file_path):
        file_name, file_extension = os.path.splitext(file_name)
        count = 1
        while os.path.exists(os.path.join("done", f"{file_name}_{count}{file_extension}")):
            count += 1
        file_path = os.path.join("done", f"{file_name}_{count}{file_extension}")
    
    translated_doc.save(file_path)
    return file_path

def create_download_link(file_path):
    with open(file_path, "rb") as file:
        b64 = base64.b64encode(file.read()).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">Download {os.path.basename(file_path)}</a>'
    return href

def detect_foreign_language(text):
    non_ascii_count = sum(1 for char in text if ord(char) > 127)
    return non_ascii_count / len(text) > 0.8

def translate_excel(file_path, target_language="id"):
    df = pd.read_excel(file_path)
    num_rows, num_cols = df.shape
    translated_df = pd.DataFrame(columns=df.columns)

    for col in range(num_cols):
        translated_col = []
        for row in range(num_rows):
            cell_value = df.iloc[row, col]
            if pd.notnull(cell_value):
                if detect_foreign_language(cell_value):
                    translated_text = translate_text(cell_value, target_language)
                    translated_col.append(translated_text)
                else:
                    translated_col.append(cell_value)
            else:
                translated_col.append(None)
        translated_df[df.columns[col]] = translated_col

    return translated_df

def translate_docx_with_style(file_content, target_language="id"):
    translated_doc = Document()
    doc = Document(io.BytesIO(file_content))
    translator = GoogleTranslator(source='auto', target=target_language)

    for paragraph in doc.paragraphs:
        translated_paragraph = translated_doc.add_paragraph()
        for run in paragraph.runs:
            translated_text = translator.translate(run.text)
            new_run = translated_paragraph.add_run(translated_text)
            
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline
            new_run.font.size = run.font.size
            new_run.font.name = run.font.name
            
        translated_paragraph.style = paragraph.style

    return translated_doc

def main():
    st.title("YOP2 TR")
    st.write("Upload File nya cuy")
    
    uploaded_file = st.file_uploader("hm", type=["txt", "xlsx", "docx"])
    target_language_options = ["Arabic", "German", "Spanish", "French", "Hindi", "Indonesian", "Japanese", "Chinese"]
    target_language_codes = ["ar", "de", "es", "fr", "hi", "id", "ja", "zh-CN"]

    target_language = st.selectbox("Diterjemahin kemana?", target_language_options)
    target_language_code = target_language_codes[target_language_options.index(target_language)]

    if uploaded_file is not None:
        file_extension = uploaded_file.name.split(".")[-1]
        
        if 'translated_file_path' not in st.session_state:
            st.session_state.translated_file_path = None

        if st.button("Proses"):
            with st.spinner("Menerjemahkan..."):
                if file_extension == "txt":
                    # Handle TXT file translation (implement this part)
                    st.write("TXT translation not implemented yet.")
                elif file_extension == "xlsx":
                    # Handle XLSX file translation (implement this part)
                    st.write("XLSX translation not implemented yet.")
                elif file_extension == "docx":
                    translated_doc = translate_docx_with_style(uploaded_file.read(), target_language_code)
                    st.session_state.translated_file_path = save_docx_file(translated_doc, f"translated_{uploaded_file.name}")
                    st.success("Terjemahan selesai!")
                else:
                    st.write("Format file tidak didukung. Silakan unggah file TXT, XLSX, atau DOCX.")

        if st.session_state.translated_file_path:
            st.download_button(
                label="Download file terjemahan",
                data=open(st.session_state.translated_file_path, "rb").read(),
                file_name=os.path.basename(st.session_state.translated_file_path),
                mime="application/octet-stream"
            )

if __name__ == "__main__":
    main()
