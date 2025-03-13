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
EXTRACTED_FOLDER = "extracted_files"

def clean_text(text):
    """ Membersihkan teks dari spasi berlebih """
    return re.sub(r"\s+", " ", text).strip()

def extract_text_from_pdf(pdf_path, debug=False):
    """ Mengekstrak teks dari PDF menggunakan PyMuPDF, fallback ke OCR jika kosong """
    with fitz.open(pdf_path) as doc:
        text = "\n".join(page.get_text("text") for page in doc).strip()
    
    if not text:  # Jika teks kosong, gunakan OCR
        images = convert_from_path(pdf_path)
        text = "\n".join(pytesseract.image_to_string(img) for img in images)
    
    if debug:
        st.text(f"=== DEBUG TEXT FROM {pdf_path} ===\n{text}\n=======================")
    
    return clean_text(text)

def extract_pph_dpp_tarif(text):
    pattern_uu_pph = r"24-\d{3}-\d{2}.*?UU PPh\.\s*([\d.,]+)\s*(\d+)\s*([\d.,]+)\s*B\.8"
    pattern_jasa = r"24-\d{3}-\d{2}\s+Jasa Perantara dan/atau Keagenan\s+([\d.]+)\s+(\d+)\s+([\d.]+)\s+B\.8"
    
    match = re.search(pattern_uu_pph, text, re.DOTALL)
    if match:
        dpp_value = int(match.group(1).replace(".", "").replace(",", ".")) if match.group(1) else None
        tarif_value = int(match.group(2)) if match.group(2) else None
        pph_value = int(match.group(3).replace(".", "").replace(",", ".")) if match.group(3) else None
    else:
        dpp_value, tarif_value, pph_value = None, None, None
    
    if dpp_value is None or tarif_value is None or pph_value is None:
        match = re.search(pattern_jasa, text)
        if match:
            dpp_value = dpp_value or int(match.group(1).replace(".", ""))
            tarif_value = tarif_value or int(match.group(2))
            pph_value = pph_value or int(match.group(3).replace(".", ""))

    return {"PPH": pph_value, "DPP": dpp_value, "Tarif": tarif_value}

def extract_additional_values(text):
    patterns = {
        "Nomor Dokumen": r"B\.9\s+Nomor Dokumen\s*: \s*([\w/-]+)",
        "Nomor": r"PEMUNGUTAN\s+PPh\s+PEMUNGUTAN\s+([A-Z0-9]+)",
        "Masa Pajak": r"PEMUNGUTAN\s+PPh\s+PEMUNGUTAN\s+[A-Z0-9]+\s+([\d-]+)\s+TIDAK FINAL",
        "Status Bukti Pemotongan": r"TIDAK FINAL\s+([A-Z]+)\s+A\. IDENTITAS WAJIB PAJAK",
        "Kode Objek Pajak": r"B\.\d+\s+([\d-]+)\s+",
        "NPWP": r"C\.1\s+NPWP / NIK\s*: \s*(\d+)\s+C\.2",
        "Nama Pemotong": r"C\.3\s+NAMA PEMOTONG DAN/ATAU PEMUNGUT\s+PPh\s*: \s*(.*?)\s+C\.4",
        "Tanggal": r"C\.4\s+TANGGAL\s*: \s*(.*?)\s+C\.5"
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

def extract_compressed_file(file_path):
    """ Mengekstrak file ZIP atau RAR dan mengembalikan daftar file PDF """
    os.makedirs(EXTRACTED_FOLDER, exist_ok=True)
    
    if zipfile.is_zipfile(file_path):
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            zip_ref.extractall(EXTRACTED_FOLDER)
    elif file_path.lower().endswith(".rar"):
        patoolib.extract_archive(file_path, outdir=EXTRACTED_FOLDER)
    
    return [os.path.join(EXTRACTED_FOLDER, f) for f in os.listdir(EXTRACTED_FOLDER) if f.lower().endswith(".pdf")]

def process_files(file_paths):
    pdf_files = []
    count_zip = count_rar = count_pdf = 0
    for file in file_paths:
        if file.lower().endswith(".zip"):
            pdf_files.extend(extract_compressed_file(file))
            count_zip += 1
        elif file.lower().endswith(".rar"):
            pdf_files.extend(extract_compressed_file(file))
            count_rar += 1
        elif file.lower().endswith(".pdf"):
            pdf_files.append(file)
            count_pdf += 1
    
    data = []
    for pdf in pdf_files:
        try:
            text = extract_text_from_pdf(pdf)
            extracted_data = extract_all_values(text)
            extracted_data["File"] = os.path.basename(pdf)
            data.append(extracted_data)
            st.success(f"✅ Processed: {pdf}")
        except Exception as e:
            st.error(f"❌ Error processing {pdf}: {e}")
    
    df = pd.DataFrame(data)
    duplicate_rows = df.duplicated().sum()
    df.drop_duplicates(inplace=True)
    df.reset_index(drop=True, inplace=True)
    unique_rows = len(df) - duplicate_rows
    shutil.rmtree(EXTRACTED_FOLDER, ignore_errors=True)
    
    return df, len(file_paths), count_zip, count_rar, count_pdf, duplicate_rows, unique_rows

# Streamlit UI
uploaded_files = st.file_uploader("Upload file PDF atau ZIP/RAR", accept_multiple_files=True, type=["pdf", "zip", "rar"])

if uploaded_files:
    temp_files = []
    for uploaded_file in uploaded_files:
        temp_path = os.path.join(EXTRACTED_FOLDER, uploaded_file.name)
        os.makedirs(EXTRACTED_FOLDER, exist_ok=True)
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        temp_files.append(temp_path)
    
    df, total_files, count_zip, count_rar, count_pdf, duplicate_rows, unique_rows = process_files(temp_files)
    st.write(f"### Statistik File yang Diupload:")
    st.write(f"Total Files: {total_files}")
    st.write(f"PDF: {count_pdf}, ZIP: {count_zip}, RAR: {count_rar}")
    st.write(f"Baris Duplikat: {duplicate_rows}, Baris Unik: {unique_rows}")
    
    if not df.empty:
        st.write("### Hasil Ekstraksi:")
        st.dataframe(df)
        file_name = st.text_input("Masukkan nama file untuk disimpan ke excel:", "hasil_ekstraksi")
        excel_file = f"{file_name}.xlsx"
        df.to_excel(excel_file, index=False)
        with open(excel_file, "rb") as f:
            st.download_button("📥 Download Excel", data=f, file_name=excel_file)

    else:
        st.warning("Tidak ada data yang dapat diekstrak.")
