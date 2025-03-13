import os
import re
import shutil
import zipfile
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import patoolib
import pytesseract
from pdf2image import convert_from_path
from tempfile import TemporaryDirectory

# Streamlit UI Header
st.title("Bukti Potong (PPh 23) Data Extractor")
st.write("### Upload files PDF, ZIP, atau RAR untuk diekstraksi !!")

# Folder sementara untuk ekstraksi
def extract_compressed_file(file_path, extract_dir):
    if zipfile.is_zipfile(file_path):
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            zip_ref.extractall(extract_dir)
    elif file_path.lower().endswith(".rar"):
        patoolib.extract_archive(file_path, outdir=extract_dir)
    return [os.path.join(extract_dir, f) for f in os.listdir(extract_dir) if f.lower().endswith(".pdf")]

def extract_text_from_pdf(pdf_path):
    with fitz.open(pdf_path) as doc:
        text = "\n".join(page.get_text("text") for page in doc).strip()
    if not text:  # Fallback ke OCR jika teks kosong
        images = convert_from_path(pdf_path)
        text = "\n".join(pytesseract.image_to_string(img) for img in images)
    return re.sub(r"\s+", " ", text).strip()


def extract_pph_dpp_tarif(text):
    """ Mengekstrak nilai PPH, DPP, dan Tarif dari teks """
    pattern = r"24-\d{3}-\d{2}.*?\n([\d.,]+)\s+(\d+)\s+([\d.,]+)"
    
    match = re.search(pattern, text, re.DOTALL)
    if match:
        try:
            dpp_value = int(match.group(1).replace(".", "").replace(",", ".")) if match.group(1) else None
            tarif_value = int(match.group(2)) if match.group(2) else None
            pph_value = int(match.group(3).replace(".", "").replace(",", ".")) if match.group(3) else None
        except ValueError:
            dpp_value, tarif_value, pph_value = None, None, None
    else:
        dpp_value, tarif_value, pph_value = None, None, None

    return {"PPH": pph_value, "DPP": dpp_value, "Tarif": tarif_value}

def extract_additional_values(text):
    patterns = {
        "Nomor Dokumen": r"B\.9\s+Nomor Dokumen\s*:\s*([\w/-]+)",
        "Nomor": r"PEMUNGUTAN\s+PPh\s+PEMUNGUTAN\s+([A-Z0-9]+)",
        "Masa Pajak": r"PEMUNGUTAN\s+PPh\s+PEMUNGUTAN\s+[A-Z0-9]+\s+([\d-]+)\s+TIDAK FINAL",
        "Kode Objek Pajak": r"B\.\d+\s+([\d-]+)\s+",
        "NPWP": r"C\.1\s+NPWP / NIK\s*:\s*(\d+)\s+C\.2",
        "Nama Pemotong": r"C\.3\s+NAMA PEMOTONG DAN/ATAU PEMUNGUT\s+PPh\s*:\s*(.*?)\s+C\.4",
        "Tanggal": r"C\.4\s+TANGGAL\s*:\s*(.*?)\s+C\.5"
    }
    
    extracted_values = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.DOTALL)
        if match:
            extracted_values[key] = match.group(1).strip()
    
    return extracted_values

def extract_all_values(text):
    extracted = extract_additional_values(text)
    extracted.update(extract_pph_dpp_tarif(text))
    return extracted

def process_files(uploaded_files):
    with TemporaryDirectory() as temp_dir:
        pdf_files = []
        count_zip, count_rar, count_pdf = 0, 0, 0
        for uploaded_file in uploaded_files:
            file_path = os.path.join(temp_dir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            if file_path.lower().endswith(".zip"):
                count_zip += 1
                pdf_files.extend(extract_compressed_file(file_path, temp_dir))
            elif file_path.lower().endswith(".rar"):
                count_rar += 1
                pdf_files.extend(extract_compressed_file(file_path, temp_dir))
            elif file_path.lower().endswith(".pdf"):
                count_pdf += 1
                pdf_files.append(file_path)
        
        data = []
        for pdf in pdf_files:
            try:
                text = extract_text_from_pdf(pdf)
                extracted_data = extract_all_values(text)
                extracted_data["File"] = os.path.basename(pdf)
                data.append(extracted_data)
            except Exception as e:
                st.error(f"❌ Error processing {pdf}: {e}")
        
        df = pd.DataFrame(data)
        duplicate_rows = df.duplicated().sum()
        df.drop_duplicates(inplace=True)
        df.reset_index(drop=True, inplace=True)
        unique_rows = len(df)
        
        return df, len(uploaded_files), count_zip, count_rar, count_pdf, duplicate_rows, unique_rows

# Streamlit UI - Upload File
uploaded_files = st.file_uploader("Upload file (PDF, ZIP, RAR)", accept_multiple_files=True, type=["pdf", "zip", "rar"])

if uploaded_files:
    df, total_files, count_zip, count_rar, count_pdf, duplicate_rows, unique_rows = process_files(uploaded_files)
    st.write(f"### Statistik File yang Diupload:")
    st.write(f"Total Files: {total_files}")
    st.write(f"PDF: {count_pdf}, ZIP: {count_zip}, RAR: {count_rar}")
    st.write(f"Baris Duplikat: {duplicate_rows}, Baris Unik: {unique_rows}")
    
    if not df.empty:
        st.write("### Hasil Ekstraksi (Contoh 10 Baris):")
        st.dataframe(df.head(10))
        
        st.write("### Simpan Hasil Ekstraksi ke Excel")
        file_name = st.text_input("Masukkan nama file untuk disimpan ke excel (tanpa ekstensi):", "hasil_ekstraksi")
        excel_file = f"{file_name}.xlsx"
        df.to_excel(excel_file, index=False)
        
        with open(excel_file, "rb") as f:
            st.download_button("📥 Download Excel", data=f, file_name=excel_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Tidak ada data yang dapat diekstrak.")
