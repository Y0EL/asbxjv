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
    # Create 'done' folder if not exists
    if not os.path.exists("done"):
        os.makedirs("done")
    
    # Save translated text to file in the 'done' folder
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
    # Create 'done' folder if not exists
    if not os.path.exists("done"):
        os.makedirs("done")
    
    # Save translated dataframe to excel file in the 'done' folder
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
    # Create 'done' folder if not exists
    if not os.path.exists("done"):
        os.makedirs("done")
    
    # Save translated document to docx file in the 'done' folder
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
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}"><button style="background-color: #008CBA; border: none; color: white; padding: 10px 20px; text-align: center; text-decoration: none; display: inline-block; font-size: 16px; margin: 4px 2px; cursor: pointer; border-radius: 12px;">Download {os.path.basename(file_path)}</button></a>'
    return href

def detect_foreign_language(text):
    # Your implementation to detect foreign language here
    # For simplicity, let's assume any text with more than 80% non-ASCII characters is considered a foreign language
    non_ascii_count = sum(1 for char in text if ord(char) > 127)
    return non_ascii_count / len(text) > 0.8

def translate_excel(file_path, target_language="id"):
    # Load Excel file
    df = pd.read_excel(file_path)

    # Get the number of rows and columns in the dataframe
    num_rows, num_cols = df.shape

    # Create an empty DataFrame to store the translated data
    translated_df = pd.DataFrame(columns=df.columns)

    # Translate foreign languages in each cell and add translated rows to the new DataFrame
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
    doc = Document(io.BytesIO(file_content))  # Gunakan io.BytesIO untuk membaca dari objek BytesIO
    translator = GoogleTranslator(source='auto', target=target_language)

    for paragraph in doc.paragraphs:
        translated_paragraph = translated_doc.add_paragraph()
        for run in paragraph.runs:
            translated_text = translator.translate(run.text)
            new_run = translated_paragraph.add_run(translated_text)
            
            # Apply styles from original run to the new run
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline
            new_run.font.size = run.font.size
            new_run.font.name = run.font.name
            # Add other style attributes as needed
            
        # Add paragraph style (if any)
        translated_paragraph.style = paragraph.style

    return translated_doc

def main():
    st.title("YOP2 TR")
    st.write("Upload File nya cuy")
    
    uploaded_file = st.file_uploader("hm", type=["txt", "xlsx", "docx"])
    target_language_options = ["Arabic", "German", "Spanish", "French", "Hindi", "Indonesian", "Japanese", "Chinese"]
    target_language_codes = ["ar", "de", "es", "fr", "hi", "id", "ja", "zh-CN"]  # Updated target language codes

    target_language = st.selectbox("Diterjemahin kemana?", target_language_options)

    # Map target language to its corresponding language code
    target_language_code = target_language_codes[target_language_options.index(target_language)]

    if uploaded_file is not None:
        file_extension = uploaded_file.name.split(".")[-1]
        if file_extension == "txt":
            # Handle TXT file translation
            pass
        elif file_extension == "xlsx":
            # Handle XLSX file translation
            pass
        elif file_extension == "docx":
            translated_doc = translate_docx_with_style(uploaded_file.read(), target_language_code)  # Pass the language code
            translated_file_path = save_docx_file(translated_doc, f"translated_{uploaded_file.name}")
            st.subheader("Translated Document:")
            st.write(translated_file_path)
            download_link = create_download_link(translated_file_path)
            st.markdown(download_link, unsafe_allow_html=True)
        else:
            st.write("Unsupported file format. Please upload a TXT, XLSX, or DOCX file.")

if __name__ == "__main__":
    main()
