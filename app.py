import streamlit as st
from utils.extract import extract_docx, extract_pptx, extract_pdf
from utils.checks import check_typo, check_comma_spacing, check_number_format, italic_english

st.title("📝 Tools Pengecekan Typo & Konsistensi Angka BPS Sidoarjo")

uploaded = st.file_uploader("Upload dokumen (.docx, .pptx, .pdf)")

if uploaded:
    ext = uploaded.name.split('.')[-1].lower()

    if ext == "docx":
        text = extract_docx(uploaded)
    elif ext == "pptx":
        text = extract_pptx(uploaded)
    elif ext == "pdf":
        text = extract_pdf(uploaded)
    else:
        st.error("Format tidak didukung")
        st.stop()

    st.subheader("📄 Isi Dokumen")
    st.text_area("Teks:", text, height=300)

    # Jalankan pengecekan
    typos = check_typo(text)
    comma_errors = check_comma_spacing(text)
    number_errors = check_number_format(text)
    italicized = italic_english(text)

    st.subheader("🔍 Hasil Pengecekan")

    st.write("### Typo:")
    st.write(typos if typos else "Tidak ada typo.")

    st.write("### Salah Koma/Spasi:")
    st.write(comma_errors if comma_errors else "Tidak ada kesalahan koma.")

    st.write("### Format Angka:")
    st.write(number_errors if number_errors else "Format angka sudah benar.")

    st.write("### Teks dengan kata Inggris dicetak miring:")
    st.markdown(italicized)
