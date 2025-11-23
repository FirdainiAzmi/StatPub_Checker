import streamlit as st
import re
import pdfplumber
from docx import Document
from pptx import Presentation

# ------------------------------
# EXTRACT TEXT FUNCTIONS
# ------------------------------
def extract_pdf(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

def extract_docx(file):
    doc = Document(file)
    full_text = [para.text for para in doc.paragraphs]
    return "\n".join(full_text)

def extract_pptx(file):
    prs = Presentation(file)
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + "\n"
    return text

# ------------------------------
# CHECK FUNCTIONS
# ------------------------------
def check_typo(text):
    dictionary = ["penduduk", "kemiskinan", "sidoarjo", "jumlah", "persentase", "rumah", "bps"]
    words = re.findall(r"\b\w+\b", text.lower())

    typo_list = [w for w in words if w not in dictionary]
    return typo_list

def check_comma_spacing(text):
    wrong = re.findall(r"\s,|,\S", text)
    return wrong

def check_number_format(text):
    wrong = re.findall(r"\b\d+,\d+%|\b\d{1,3}(\.\d{3})+,\d+\b", text)
    return wrong

def italicize_english(text):
    return re.sub(r"\b([A-Za-z]+)\b", r"*\1*", text)

# ------------------------------
# STREAMLIT UI
# ------------------------------
st.title("📝 Tools Pengecekan Typo & Konsistensi Angka BPS Sidoarjo")

uploaded = st.file_uploader("Upload dokumen (PDF, DOCX, PPTX)", type=["pdf", "docx", "pptx"])

if uploaded:
    st.success("File berhasil diupload!")

    if uploaded.type == "application/pdf":
        raw_text = extract_pdf(uploaded)
    elif uploaded.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        raw_text = extract_docx(uploaded)
    else:
        raw_text = extract_pptx(uploaded)

    st.subheader("📄 Teks Hasil Ekstraksi")
    st.text_area("", raw_text, height=200)

    st.subheader("🔍 Hasil Pengecekan")

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("Cek Typo"):
            typos = check_typo(raw_text)
            st.write(typos if typos else "Tidak ada typo!")

    with col2:
        if st.button("Cek Format Koma"):
            comma = check_comma_spacing(raw_text)
            st.write(comma if comma else "Format koma sudah benar!")

    with col3:
        if st.button("Cek Format Angka"):
            number = check_number_format(raw_text)
            st.write(number if number else "Format angka sudah benar!")

    st.subheader("✏️ Kata Bahasa Inggris (Italic)")
    if st.button("Convert English to Italic"):
        st.text_area("Hasil:", italicize_english(raw_text), height=200)
