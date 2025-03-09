import re
import pandas as pd
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pytesseract
import os
import streamlit as st

def clean_text(text):
    return re.sub(r"\s+", " ", text).strip()

def extract_text_from_pdf(pdf_path):
    with fitz.open(pdf_path) as doc:
        text = "\n".join(page.get_text("text") for page in doc)
    
    if not text.strip():
        images = convert_from_path(pdf_path)
        text = "\n".join(pytesseract.image_to_string(img) for img in images)
    
    return clean_text(text)

def extract_values(text):
    patterns = {
        "Nomor": r"PEMUNGUTAN\s+PPh\s+PEMUNGUTAN\s+([A-Z0-9]+)",
        "Masa Pajak": r"PEMUNGUTAN\s+PPh\s+PEMUNGUTAN\s+[A-Z0-9]+\s+([\d-]+)\s+TIDAK FINAL",
        "Kode Objek Pajak": r"B\.7\s+([\d-]+)\s+Jasa Perantara dan/atau Keagenan",
        "DPP": r"B\.7\s+[\d-]+\s+Jasa Perantara dan/atau Keagenan\s+([\d.]+)",
        "PPH": r"B\.7\s+[\d-]+\s+Jasa Perantara dan/atau Keagenan\s+[\d.]+\s+\d+\s+([\d.]+)",
        "NPWP": r"C\.1\s+NPWP / NIK\s*:\s*(\d+)\s+C\.2",
        "Nama Pemotong": r"C\.3\s+NAMA PEMOTONG DAN/ATAU PEMUNGUT\s+PPh\s*:\s*(.*?)\s+C\.4",
        "Tanggal": r"C\.4\s+TANGGAL\s*:\s*(.*?)\s+C\.5",
        "Tarif": r"Jasa Perantara dan/atau Keagenan\s+[\d.]+\s+(\d+)\s+[\d.]+\s+B\.8"
    }
    
    extracted_values = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.DOTALL)
        if match:
            value = match.group(1).strip()
            if key in ["DPP", "PPH", "Tarif"]:
                value = int(value.replace('.', ''))
            extracted_values[key] = value
    
    return extracted_values

st.title("Bukti Potong Data Extractor")

uploaded_files = st.file_uploader("Upload PDF Files", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    data = []
    for uploaded_file in uploaded_files:
        with open(uploaded_file.name, "wb") as f:
            f.write(uploaded_file.getbuffer())
        text = extract_text_from_pdf(uploaded_file.name)
        extracted_data = extract_values(text)
        extracted_data["File"] = uploaded_file.name
        data.append(extracted_data)
        os.remove(uploaded_file.name)
    
    df = pd.DataFrame(data)
    st.write("### Sample Extracted Data")
    st.dataframe(df.head(10))
    
    if not df.empty:
        file_name = st.text_input("Enter filename to save (without extension)")
        if st.button("Save to Excel"):
            excel_path = f"{file_name}.xlsx"
            df.to_excel(excel_path, index=False)
            st.success(f"Data saved to {excel_path}")
            st.download_button(label="Download Excel", data=open(excel_path, "rb").read(), file_name=excel_path, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
