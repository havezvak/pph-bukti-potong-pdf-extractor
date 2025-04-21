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
import patoolib

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

def extract_compressed_file(file_path):
    """ Mengekstrak file ZIP atau RAR dan mengembalikan daftar file PDF """
    os.makedirs(EXTRACTED_FOLDER, exist_ok=True)
    extracted_pdfs = []
    
    if zipfile.is_zipfile(file_path):
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            zip_ref.extractall(EXTRACTED_FOLDER)
    elif file_path.lower().endswith(".rar"):
        patoolib.extract_archive(file_path, outdir=EXTRACTED_FOLDER)
    
    extracted_pdfs = [os.path.join(EXTRACTED_FOLDER, f) for f in os.listdir(EXTRACTED_FOLDER) if f.lower().endswith(".pdf")]
    return extracted_pdfs, len(extracted_pdfs)  # Kembalikan daftar PDF & jumlahnya

def extract_pph_dpp_tarif(text):
    """Ekstraksi PPH, DPP, dan Tarif dari berbagai pola teks."""
    patterns = [
        r"24-\d{3}-\d{2}.*?UU PPh\.\s*([\d.,]+)\s*(\d+)\s*([\d.,]+)",
        r"24-\d{3}-\d{2}\s+Jasa Perantara dan/atau Keagenan\s+([\d.,]+)\s+(\d+)\s+([\d.,]+)"
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.DOTALL)
        if match:
            dpp_value = int(match.group(1).replace(",", "").replace(".", ""))
            tarif_value = int(match.group(2))
            pph_value = int(match.group(3).replace(",", "").replace(".", ""))
            return {"PPH": pph_value, "DPP": dpp_value, "Tarif": tarif_value}

    return {"PPH": None, "DPP": None, "Tarif": None}

def extract_all_values(text):
    """Ekstraksi semua data dari pola teks yang berbeda, dengan fallback ke Code 1 jika perlu."""
    
    # Pola regex dari Code 2 (prioritas utama)
    patterns = {
        "Nomor Dokumen": [r"Nomor Dokumen\s*:\s*:?\s*([\w\s/\-]+?)(?=\s*\d{1,2}\s+[A-Za-z]+\s+\d{4}|$)", r"B\.9\s+Nomor Dokumen\s*:\s*([\w/-]+)"],
        "Nomor": [r"NOMOR\s*([\w\d]+)", r"PEMUNGUTAN\s+PPh\s+PEMUNGUTAN\s+([A-Z0-9]+)"],
        "Masa Pajak": [r"MASA PAJAK\s*([\d-]+)", r"PEMUNGUTAN\s+PPh\s+PEMUNGUTAN\s+[A-Z0-9]+\s+([\d-]+)\s+TIDAK FINAL"],
        "Kode Objek Pajak": [r"B\.7\s+(24-\d{3}-\d{2})", r"B\.\d+\s+([\d-]+)\s+"],
        "NPWP": [r"NPWP / NIK\s*:\s*(\d+)", r"C\.1\s+NPWP / NIK\s*:\s*(\d+)\s+C\.2"],
        "Nama Pemotong": [r"C\.3\s+NAMA PEMOTONG DAN/ATAU PEMUNGUT PPH\s*:\s*(.*?)\s*C\.4"],
        "Tanggal": [r"C\.4\s+TANGGAL\s*:\s*([\d]+\s+[A-Za-z]+\s+\d+)", r"C\.4\s+TANGGAL\s*:\s*(.*?)\s+C\.5"],
        "NITKU": [
            r"C\.1\s+NPWP / NIK\s*:\s*\d+\s+(\d{16,})\s+C\.2",  # Teks 1: setelah NPWP dan sebelum C.2
            r"C\.2\s+NOMOR IDENTITAS TEMPAT.*?USAHA \(NITKU\).*?:\s*(\d{16,} - .*?)\s+C\.3"  # Teks 2: langsung pada bagian C.2
        ]
    }

    extracted_values = {}
    for key, patterns_list in patterns.items():
        for pattern in patterns_list:
            match = re.search(pattern, text, re.DOTALL)
            if match:
                extracted_values[key] = match.group(1).strip()
                break  # Berhenti jika sudah menemukan hasil

    # **Fallback ke Code 1 jika "Nomor" < 9 karakter atau "Nama Pemotong" kosong**
    if len(extracted_values.get("Nomor", "")) != 9 or not extracted_values.get("Nama Pemotong"):
        fallback_patterns = {
            "Nomor": r"PEMUNGUTAN\s+PPh\s+PEMUNGUTAN\s+([A-Z0-9]+)",
            "Nama Pemotong": r"C\.3\s+NAMA PEMOTONG DAN/ATAU PEMUNGUT\s+PPh\s*:\s*(.*?)\s+C\.4"
        }

        for key, pattern in fallback_patterns.items():
            match = re.search(pattern, text, re.DOTALL)
            if match:
                extracted_values[key] = match.group(1).strip()

    # **Tambahkan hasil ekstraksi PPH, DPP, dan Tarif**
    extracted_values.update(extract_pph_dpp_tarif(text))

    return extracted_values

def process_files(file_paths):
    pdf_files = []
    count_zip = count_rar = count_pdf = 0
    extracted_from_zip = extracted_from_rar = 0
    success_logs = []
    error_logs = []

    for file in file_paths:
        if file.lower().endswith(".zip"):
            extracted_pdfs, extracted_count = extract_compressed_file(file)
            pdf_files.extend(extracted_pdfs)
            count_zip += 1
            extracted_from_zip += extracted_count
        elif file.lower().endswith(".rar"):
            extracted_pdfs, extracted_count = extract_compressed_file(file)
            pdf_files.extend(extracted_pdfs)
            count_rar += 1
            extracted_from_rar += extracted_count
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
            success_logs.append(f"✅ Processed: {pdf}")
        except Exception as e:
            error_logs.append(f"❌ Error processing {pdf}: {e}")
    
    df = pd.DataFrame(data)
    duplicate_rows = df.duplicated().sum()
    df.drop_duplicates(inplace=True)
    df.reset_index(drop=True, inplace=True)
    unique_rows = len(df)

    shutil.rmtree(EXTRACTED_FOLDER, ignore_errors=True)
    
    return df, len(file_paths), count_zip, count_rar, count_pdf, extracted_from_zip, extracted_from_rar, duplicate_rows, unique_rows, success_logs, error_logs

# Streamlit UI
uploaded_files = st.file_uploader("Upload file PDF atau ZIP/RAR", accept_multiple_files=True, type=["pdf", "zip", "rar"])

if uploaded_files:
    temp_files = [os.path.join(EXTRACTED_FOLDER, f.name) for f in uploaded_files]
    os.makedirs(EXTRACTED_FOLDER, exist_ok=True)
    for f in uploaded_files:
        with open(os.path.join(EXTRACTED_FOLDER, f.name), "wb") as out_file:
            out_file.write(f.getbuffer())
    
    df, total_files, count_zip, count_rar, count_pdf, extracted_from_zip, extracted_from_rar, duplicate_rows, unique_rows, success_logs, error_logs = process_files(temp_files)

    # Statistik File yang Diupload
    st.write(f"### Statistik File yang Diupload:")
    st.write(f"Total Files: {total_files}")
    st.write(f"PDF: {count_pdf}, ZIP: {count_zip}, RAR: {count_rar}")
    st.write(f"PDF dari ZIP: {extracted_from_zip}, PDF dari RAR: {extracted_from_rar}, PDF keseluruhan = {count_pdf + extracted_from_zip + extracted_from_rar}")
    st.write(f"Baris Duplikat: {duplicate_rows}, Baris Unik: {unique_rows}")

    # Log Processing dalam Expander
    with st.expander("✅ Log File yang Berhasil Diproses"):
        for log in success_logs:
            st.success(log)

    with st.expander("❌ Log File yang Gagal Diproses"):
        for log in error_logs:
            st.error(log)

    # Tampilkan Hasil Ekstraksi & Download
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
