import streamlit as st
import re
from spellchecker import SpellChecker
import PyPDF2
from docx import Document
from docx.enum.text import WD_COLOR_INDEX

# ==========================
# LOAD WORDLIST
# ==========================
@st.cache_resource
def load_wordlists():
    with open("wordlists/id_wordlist.txt", "r", encoding="utf-8") as f:
        indo_words = set(f.read().splitlines())

    with open("wordlists/eng_wordlist.txt", "r", encoding="utf-8") as f:
        eng_words = set(f.read().splitlines())

    return indo_words, eng_words


indo_dict, eng_dict = load_wordlists()
spell_en = SpellChecker(language="en")


# ==========================
# FILE READER
# ==========================
def read_pdf(file):
    reader = PyPDF2.PdfReader(file)
    text = ""
    for p in reader.pages:
        try:
            text += p.extract_text() + " "
        except:
            pass
    return text


def read_docx(file):
    doc = Document(file)
    text = " ".join([p.text for p in doc.paragraphs])
    return text


# ==========================
# DETEKSI TYPO
# ==========================
def detect_typo(text):
    words = re.findall(r"\b[a-zA-Z]+\b", text)
    typos = set()

    for w in words:
        wl = w.lower()

        if wl in indo_dict:
            continue
        if wl in eng_dict:
            continue
        if wl in spell_en:
            continue

        typos.add(w)

    return typos


# ==========================
# DETEKSI KATA INGGRIS
# ==========================
def detect_english_words(text):
    words = re.findall(r"\b[a-zA-Z]+\b", text)
    english_found = []

    for w in words:
        wl = w.lower()
        if wl in eng_dict or wl in spell_en:
            english_found.append(w)

    return set(english_found)


# ==========================
# DETEKSI FORMAT PERSENTASE SALAH
# ==========================
def detect_wrong_percentage(text):
    wrong = []
    pattern_correct = r"\b\d+\.\d+%\b"
    pattern_any = r"\b\d+[%]\b|\b\d+[.,]\d+[%]\b"

    all_found = re.findall(pattern_any, text)
    correct = re.findall(pattern_correct, text)

    for item in all_found:
        if item not in correct:
            wrong.append(item)

    return set(wrong)


# ==========================
# HIGHLIGHT RUN DALAM DOCX
# ==========================
def highlight_docx(input_file, typos, english_words, wrong_percent):
    doc = Document(input_file)

    problem_words = typos.union(english_words).union(wrong_percent)

    for para in doc.paragraphs:
        for run in para.runs:
            text = run.text
            words_in_run = re.findall(r"\b[a-zA-Z0-9.,%]+\b", text)

            # Jika sebuah run mengandung 1 kata salah → highlight seluruh run
            if any(w in problem_words for w in words_in_run):
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    return doc


# ==========================
# STREAMLIT UI
# ==========================
st.title("📘 Checker Publikasi BPS – Typo, English, & Format Persentase")
st.write("Upload PDF atau Word, hasil akhir berupa DOCX dengan highlight tanpa mengubah struktur dokumen.")

uploaded = st.file_uploader("Upload file publikasi", type=["pdf", "docx"])

if uploaded:
    # Ambil teks untuk dianalisis
    if uploaded.type == "application/pdf":
        text = read_pdf(uploaded)
    else:
        text = read_docx(uploaded)

    # Analisis
    typos = detect_typo(text)
    eng_words = detect_english_words(text)
    wrong_percent = detect_wrong_percentage(text)

    st.subheader("🔍 Hasil Analisis")

    st.write("### ❌ Typo ditemukan:")
    st.write(typos if typos else "Tidak ada.")

    st.write("### 🔤 Kata Inggris:")
    st.write(eng_words if eng_words else "Tidak ada.")

    st.write("### % Format persentase salah:")
    st.write(wrong_percent if wrong_percent else "Tidak ada.")

    # Hasil DOCX
    st.subheader("📄 Download Dokumen Hasil Highlight")

    # Untuk PDF → convert ke DOCX dulu
    if uploaded.type == "application/pdf":
        temp_doc = Document()
        for line in text.split("\n"):
            temp_doc.add_paragraph(line)
        temp_doc.save("converted_temp.docx")
        highlighted = highlight_docx("converted_temp.docx", typos, eng_words, wrong_percent)
    else:
        highlighted = highlight_docx(uploaded, typos, eng_words, wrong_percent)

    output_path = "hasil_checker.docx"
    highlighted.save(output_path)

    with open(output_path, "rb") as f:
        st.download_button("⬇ Download hasil highlight (DOCX)", f, file_name="hasil_checker.docx")
