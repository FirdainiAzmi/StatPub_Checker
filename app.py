import streamlit as st
import re
import pdfplumber
from docx import Document
from pptx import Presentation
from docx.shared import RGBColor


# =========================================================
# 1. Fungsi Extraction
# =========================================================
def extract_pdf(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                text += txt + "\n"
    return text

def extract_docx(file):
    doc = Document(file)
    return "\n".join(p.text for p in doc.paragraphs)

def extract_pptx(file):
    prs = Presentation(file)
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + "\n"
    return text


# =========================================================
# 2. Dictionaries
# =========================================================
indonesian_words = {
    "jumlah","penduduk","tahun","data","statistik","bps","kabupaten","kota",
    "menurut","persen","rata","rata-rata","nilai","tingkat","kemiskinan",
    "sidoarjo","kecamatan","metode","publikasi","indonesia","luas","wilayah"
}

english_dict = {
    "trend","growth","poverty","income","population",
    "analysis","rate","average","value","standard","index"
}


# =========================================================
# 3. Functions Check
# =========================================================
def check_typo(text):
    words = re.findall(r"\b\w+\b", text)
    typo = []
    for w in words:
        lw = w.lower()
        if lw not in indonesian_words and lw not in english_dict and not lw.isnumeric():
            typo.append(w)
    return typo

def check_punctuation(text):
    issues = []

    if re.search(r"\s,", text):
        issues.append("Spasi sebelum koma salah.")
    if re.search(r",\S", text):
        issues.append("Tidak ada spasi setelah koma.")
    if re.search(r"\s\.", text):
        issues.append("Spasi sebelum titik salah.")
    if re.search(r":\S", text):
        issues.append("Tidak ada spasi setelah titik dua.")

    return issues

def check_percentage(text):
    wrong = re.findall(r"\d+,\d+%", text)
    return wrong

def check_number_format(text):
    wrong = re.findall(r"\b\d{4,}\b", text)
    return wrong


# =========================================================
# 4. Build corrected + highlighted text
# =========================================================
def highlight_text(text, typo, pct, numbers, punct):
    result = text

    # highlight semua kesalahan → warna **kuning** di Streamlit
    for w in typo:
        result = re.sub(fr"\b{re.escape(w)}\b",
                        f"**🟡{w}**", result)

    for w in pct:
        result = result.replace(w, f"**🟡{w}**")

    for w in numbers:
        result = result.replace(w, f"**🟡{w}**")

    # italic english
    for w in english_dict:
        result = re.sub(fr"\b{w}\b", f"*{w}*", result, flags=re.I)

    return result


# =========================================================
# 5. Generate Word DOCX with highlight
# =========================================================
def generate_docx(text, typo, pct, numbers):
    doc = Document()

    words = text.split()

    p = doc.add_paragraph()

    for w in words:
        run = p.add_run(w + " ")

        lw = w.lower()

        # highlight kesalahan
        if w in typo or w in pct or w in numbers:
            font = run.font
            font.highlight_color = 7   # Yellow

        # italic English
        if lw in english_dict:
            run.italic = True

    return doc


# =========================================================
# 6. Streamlit UI
# =========================================================
st.title("📝 Tools Pengecekan Typo & Konsistensi Angka BPS Sidoarjo")

uploaded = st.file_uploader("Upload dokumen:", type=["pdf", "docx", "pptx"])

if uploaded:
    if uploaded.type == "application/pdf":
        text = extract_pdf(uploaded)
    elif uploaded.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        text = extract_docx(uploaded)
    else:
        text = extract_pptx(uploaded)

    st.subheader("📄 Teks Hasil Ekstraksi")
    st.text_area("", text, height=200)

    if st.button("🔍 Jalankan Semua Pemeriksaan"):
        typo = check_typo(text)
        punct = check_punctuation(text)
        pct = check_percentage(text)
        numbers = check_number_format(text)

        highlighted = highlight_text(text, typo, pct, numbers, punct)

        st.subheader("✨ Hasil Anotasi (Highlight Kesalahan)")
        st.markdown(highlighted)

        # Generate DOCX
        doc = generate_docx(text, typo, pct, numbers)
        doc.save("hasil_pemeriksaan.docx")

        with open("hasil_pemeriksaan.docx", "rb") as f:
            st.download_button(
                "⬇️ Download DOCX (ada highlight)",
                f,
                file_name="hasil_pemeriksaan.docx"
            )
