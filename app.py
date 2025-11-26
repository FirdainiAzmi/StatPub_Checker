import streamlit as st
from docx import Document
import re
from spellchecker import SpellChecker
from langdetect import detect
import fitz  # PyMuPDF untuk PDF
from kbbi import KBBI
import tempfile


# =====================================================
# Ekstraksi teks dari DOCX
# =====================================================
def extract_text_docx(file):
    doc = Document(file)
    all_text = []

    for para in doc.paragraphs:
        all_text.append(para.text)

    return "\n".join(all_text), doc


# =====================================================
# Ekstraksi teks dari PDF
# =====================================================
def extract_text_pdf(file):
    pdf = fitz.open(stream=file.read(), filetype="pdf")
    text = ""

    for page in pdf:
        text += page.get_text()

    return text, None


# =====================================================
# Cek Typo
# =====================================================
def check_typo(text):
    spell_en = SpellChecker(language='en')   # hanya untuk English
    words = re.findall(r'\b[a-zA-Z]+\b', text)
    errors = []

    for word in words:
        w = word.lower()

        # 1. CEK BAHASA INDONESIA VIA KBBI
        try:
            _ = KBBI(w)
            continue  # kata Indonesia benar
        except:
            pass

        # 2. CEK APAKAH INI KATA INGGRIS?
        try:
            if detect(w) == "en":
                if w not in spell_en:  # jika salah
                    errors.append(word)
                continue
        except:
            pass

        # 3. TIDAK ADA DI KBBI & BUKAN ENGLISH = kemungkinan TYPO
        errors.append(word)

    return list(set(errors))



# =====================================================
# Cek Format Persen
# =====================================================
def check_percentage(text):
    found = re.findall(r'[\d\.,]+%', text)
    errors = []

    for p in found:
        if not re.match(r'^\d+\.\d+%$', p):
            errors.append(p)

    return errors


# =====================================================
# Deteksi apakah kata English
# =====================================================
def is_english(word):
    try:
        lang = detect(word)
        return lang == "en"
    except:
        return False


# =====================================================
# Cek Italic kata Inggris (DOCX only)
# =====================================================
def check_italic(doc):
    if doc is None:
        return []

    errors = []

    for i, para in enumerate(doc.paragraphs):
        for run in para.runs:
            words = run.text.split()
            for w in words:
                if is_english(w) and not run.italic:
                    errors.append(f"Kata '{w}' di paragraf {i+1} seharusnya *italic*")

    return errors


# =====================================================
# STREAMLIT UI
# =====================================================
st.set_page_config(page_title="BPS Publication Checker", layout="wide")

st.title("📘 BPS Publication Checker")
st.write("Aplikasi ini membantu mendeteksi typo, format persen, dan italic istilah bahasa Inggris pada publikasi BPS.")

uploaded_file = st.file_uploader("Upload dokumen (.pdf atau .docx)", type=["pdf", "docx"])

if uploaded_file is not None:
    st.success("File berhasil diupload!")

    if uploaded_file.name.endswith(".docx"):
        text, doc = extract_text_docx(uploaded_file)
    else:
        text, doc = extract_text_pdf(uploaded_file)

    st.subheader("🔍 Proses Pengecekan")
    with st.spinner("Sedang memproses dokumen..."):
        typo_errors = check_typo(text)
        percent_errors = check_percentage(text)
        italic_errors = check_italic(doc)

    st.subheader("📌 Hasil Pemeriksaan")

    # Typo
    st.write("### 1. Typo")
    if typo_errors:
        st.error(f"Ditemukan {len(typo_errors)} typo:")
        st.write(typo_errors)
    else:
        st.success("Tidak ada typo ditemukan ✔️")

    # Persen
    st.write("### 2. Format Persen (wajib 12.34%)")
    if percent_errors:
        st.error(f"Format persen salah ditemukan:")
        st.write(percent_errors)
    else:
        st.success("Format persen sudah benar ✔️")

    # Italic
    st.write("### 3. Kata Inggris Tidak Italic")
    if doc is None:
        st.info("Italic tidak dapat dicek untuk PDF.")
    else:
        if italic_errors:
            st.error("Ditemukan kata Inggris yang seharusnya italic:")
            st.write(italic_errors)
        else:
            st.success("Tidak ada kesalahan italic ✔️")

