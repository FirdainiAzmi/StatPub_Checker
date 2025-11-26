import streamlit as st
import docx
import re
from io import BytesIO
from PyPDF2 import PdfReader

# -----------------------------
# LOAD WORDLIST
# -----------------------------
def load_wordlist(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return set([w.strip().lower() for w in f.readlines() if w.strip()])
    except:
        return set()

ENG_LIST = load_wordlist("wordlist/eng_list.txt")
IND_LIST = load_wordlist("wordlist/ind_list.txt")

# -----------------------------
# EXTRACT TEXT
# -----------------------------
def extract_from_pdf(upload):
    reader = PdfReader(upload)
    text = ""
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text += extracted + "\n"
    return text

def extract_from_docx(upload):
    doc = docx.Document(upload)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

# -----------------------------
# HIGHLIGHT WORDS NOT IN WORDLIST
# -----------------------------
def highlight_text(text, lang):

    if lang == "English":
        wordlist = ENG_LIST
    else:
        wordlist = IND_LIST

    def mark(match):
        word = match.group(0)
        if word.lower() not in wordlist:
            return f"🔶**{word}**"     # highlight
        return word

    # only highlight words, not remove or reorder anything
    processed = re.sub(r"\b[A-Za-zÀ-ÿ']+\b", mark, text)
    return processed

# -----------------------------
# STREAMLIT APP
# -----------------------------
st.set_page_config(page_title="Wordlist Checker", layout="wide")

st.title("📘 Wordlist Checker (Indonesia / English)")
st.write("Unggah dokumen dan sistem akan menandai kata yang tidak ada di wordlist Anda. Struktur dokumen tidak akan diubah.")

uploaded = st.file_uploader("Upload PDF atau DOCX", type=["pdf", "docx"])
language = st.selectbox("Pilih Bahasa Dokumen", ["Indonesia", "English"])

if uploaded:
    # Extract text depending on file type
    if uploaded.type == "application/pdf":
        raw_text = extract_from_pdf(uploaded)
    else:
        raw_text = extract_from_docx(uploaded)

    st.subheader("📄 Teks Asli (Tanpa Diubah)")
    st.code(raw_text, language="text")

    # Process with highlight
    highlighted = highlight_text(raw_text, language)

    st.subheader("✨ Teks dengan Highlight (Struktur 100% Sama)")
    st.markdown(highlighted)

    # Optional: Download result
    output = highlighted
    buffer = BytesIO()
    buffer.write(output.encode("utf-8"))
    buffer.seek(0)

    st.download_button(
        "⬇️ Download Hasil (TXT)",
        buffer,
        file_name="highlighted_result.txt",
        mime="text/plain"
    )
