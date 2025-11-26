import streamlit as st
import re
from spellchecker import SpellChecker
import PyPDF2
from docx import Document

# ======================
# LOAD WORDLIST
# ======================
@st.cache_resource
def load_wordlists():
    with open("wordlists/id_wordlist.txt", "r", encoding="utf-8") as f:
        indo_words = set(f.read().splitlines())

    with open("wordlist/englist.txt", "r", encoding="utf-8") as f:
         eng_words = set(f.read().splitlines())

    return indo_words, eng_words


indo_dict, eng_dict= load_wordlists()
spell_en = SpellChecker(language='en')


# ======================
# FILE READER
# ======================
def read_pdf(file):
    reader = PyPDF2.PdfReader(file)
    text = ""
    for p in reader.pages:
        text += p.extract_text() + " "
    return text


def read_docx(file):
    doc = Document(file)
    text = " ".join([p.text for p in doc.paragraphs])
    return text


# ======================
# DETEKSI TYPO
# ======================
def detect_typo(text):
    words = re.findall(r'\b[a-zA-Z]+\b', text)
    typos = set()

    for w in words:
        wl = w.lower()

        # Cek dictionary Indonesia
        if wl in indo_dict:
            continue

        # Cek dictionary English
        if wl in eng_dict:
            continue

        # Cek SpellChecker EN
        if wl in spell_en:  
            continue

        typos.add(w)

    return sorted(list(typos))


# ======================
# DETEKSI KATA INGGRIS (untuk italic)
# ======================
def detect_english_words(text):
    words = re.findall(r'\b[a-zA-Z]+\b', text)
    english_found = []

    for w in words:
        wl = w.lower()
        if wl in eng_dict or wl in spell_en:
            english_found.append(w)

    return sorted(set(english_found))


# ======================
# DETEKSI FORMAT PERSENTASE BENAR
# ======================
def detect_wrong_percentage(text):
    wrong = []

    # Format benar: 12.34%
    pattern_correct = r'\b\d+\.\d+%\b'
    pattern_any = r'\b\d+[%]\b|\b\d+[,]\d+[%]\b'  # menangkap angka tanpa titik

    all_found = re.findall(pattern_any, text)
    correct = re.findall(pattern_correct, text)

    for item in all_found:
        if item not in correct:
            wrong.append(item)

    return sorted(set(wrong))


# ======================
# STREAMLIT UI
# ======================
st.title("📘 Checker Publikasi BPS – Typo, English Italic, & Persentase")
st.write("Unggah PDF atau Word untuk dianalisis.")

uploaded = st.file_uploader("Upload file", type=["pdf", "docx"])

if uploaded:
    # Baca file
    if uploaded.type == "application/pdf":
        text = read_pdf(uploaded)
    else:
        text = read_docx(uploaded)

    st.subheader("📄 Hasil Analisis")
    
    # TYPO
    typos = detect_typo(text)
    st.write("### ❌ Kemungkinan Typo:")
    if typos:
        st.error(typos)
    else:
        st.success("Tidak ditemukan typo.")

    # English words (for italic)
    eng_words = detect_english_words(text)
    st.write("### 🔤 Kata Bahasa Inggris (seharusnya italic):")
    if eng_words:
        st.warning(eng_words)
    else:
        st.success("Tidak ada kata Bahasa Inggris.")

    # Percentage wrong format
    wrong_percent = detect_wrong_percentage(text)
    st.write("### % Format Persentase Salah:")
    if wrong_percent:
        st.warning(wrong_percent)
    else:
        st.success("Format persentase sudah benar (contoh benar: 12.45%).")



