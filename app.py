# app.py
import streamlit as st
import re
import io
import os
import pandas as pd
import pdfplumber
from docx import Document
from pptx import Presentation
from docx.shared import RGBColor
from spellchecker import SpellChecker
from datetime import datetime

# ---------------------------
# Config / Paths
# ---------------------------
BASE_DIR = os.path.dirname(__file__)
WORDLIST_DIR = os.path.join(BASE_DIR, "wordlists")
ID_WORDLIST = os.path.join(WORDLIST_DIR, "id_wordlist.txt")
EN_WORDLIST = os.path.join(WORDLIST_DIR, "en_wordlist.txt")
ALLOWED_WORDS = os.path.join(WORDLIST_DIR, "allowed_words.txt")

# ---------------------------
# Utility: load wordlists
# ---------------------------
def load_wordset(path):
    s = set()
    if not os.path.exists(path):
        return s
    with open(path, encoding="utf-8") as f:
        for line in f:
            w = line.strip()
            if w:
                s.add(w.lower())
    return s

ID_WORDS = load_wordset(ID_WORDLIST)     # besar: bahasa Indonesia
EN_WORDS = load_wordset(EN_WORDLIST)     # besar: english
ALLOWED = load_wordset(ALLOWED_WORDS)    # pengecualian BPS (kecamatan, istilah)

# fallback spellchecker (uses english by default; we set language=None and load id words)
spell_id = SpellChecker(language=None)
# add id words to spellchecker frequency so unknown() works better
for w in ID_WORDS:
    spell_id.word_frequency.add(w)

# ---------------------------
# Extraction functions
# ---------------------------
def extract_docx(file):
    doc = Document(file)
    paragraphs = [p.text for p in doc.paragraphs]
    return paragraphs

def extract_pdf(file):
    paragraphs = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                # split to paragraphs by double newline OR keep line by line depending on layout
                ps = [p.strip() for p in txt.split("\n\n") if p.strip()]
                if not ps:
                    # fallback: line by line
                    ps = [ln.strip() for ln in txt.split("\n") if ln.strip()]
                paragraphs.extend(ps)
    return paragraphs

def extract_pptx(file):
    prs = Presentation(file)
    texts = []
    for slide in prs.slides:
        slide_text = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                # split shape text to lines
                slide_text.extend([ln.strip() for ln in shape.text.split("\n") if ln.strip()])
        if slide_text:
            texts.append(" ".join(slide_text))
    return texts

def extract_txt(file):
    txt = file.read().decode("utf-8", errors="ignore")
    # split by double newline or single
    paras = [p.strip() for p in txt.split("\n\n") if p.strip()]
    if not paras:
        paras = [p.strip() for p in txt.split("\n") if p.strip()]
    return paras

# ---------------------------
# Normalization helpers
# ---------------------------
def normalize_whitespace(s):
    return re.sub(r"\s+", " ", s).strip()

# identify numeric patterns
re_percent_comma = re.compile(r"\b\d+,\d+%")      # e.g. 12,5%
re_percent_dot = re.compile(r"\b\d+\.\d+%")      # e.g. 12.5%
re_thousands_comma = re.compile(r"\b\d{1,3},\d{3}\b")
re_thousands_dot = re.compile(r"\b\d{1,3}\.\d{3}\b")
re_large_number = re.compile(r"\b\d{4,}\b")

# ---------------------------
# Checks (per paragraph)
# ---------------------------
def check_spelling(paragraph):
    words = re.findall(r"\b[\w'-]+\b", paragraph, flags=re.UNICODE)
    unknown = set()
    for w in words:
        lw = w.lower()
        if lw in ALLOWED or lw in EN_WORDS:
            continue
        # simple numbers skip
        if re.fullmatch(r"[\d\.,%]+", lw):
            continue
        if lw not in ID_WORDS:
            # check via spellchecker unknown
            if lw not in spell_id:
                unknown.add(w)
    return sorted(unknown)

def check_punctuation_spacing(paragraph):
    issues = []
    if re.search(r"\s,", paragraph):
        issues.append("Spasi sebelum koma (contoh: 'kata ,').")
    if re.search(r",\S", paragraph):
        # but if comma followed immediately by ) or punctuation allow? still flag
        issues.append("Tidak ada spasi setelah koma (contoh: 'kata, kata').")
    if re.search(r"\s\.", paragraph):
        issues.append("Spasi sebelum titik.")
    if re.search(r":\S", paragraph):
        issues.append("Tidak ada spasi setelah titik dua.")
    if re.search(r"\(\s|\s\)", paragraph):
        issues.append("Spasi di dalam/telah tanda kurung.")
    return issues

def check_number_rules(paragraph):
    issues = []
    # percent: check if comma used for decimal, follow BPS guideline (choose one)
    # Suppose BPS style: desimal gunakan koma (12,5%), ribuan gunakan titik (1.234)
    # We'll flag obvious mismatches: usage of comma in thousands (1,234) or dot in decimal (12.5%)
    if re_thousands_comma.search(paragraph):
        issues.append("Kemungkinan pemisah ribuan memakai koma (seharusnya titik).")
    if re_percent_dot.search(paragraph):
        issues.append("Persen menggunakan titik desimal (pertimbangkan menggunakan koma sesuai pedoman).")
    # check large number without separators
    for m in re_large_number.finditer(paragraph):
        num = m.group()
        # if num not part of longer token e.g. 2023 is valid; we won't flag years (1900-2100)
        try:
            n = int(num)
            if 1900 <= n <= 2100:
                continue
        except:
            pass
        # if number length >=5 then it's likely needs thousand separator
        if len(num) >= 5:
            issues.append(f"Angka besar tanpa pemisah ribuan: {num}")
    return issues

def detect_english_words(paragraph):
    words = re.findall(r"\b[A-Za-z]{2,}\b", paragraph)
    # return those English-looking words that match EN_WORDS and are not all-caps acronyms
    found = []
    for w in words:
        lw = w.lower()
        if lw in EN_WORDS and lw not in ALLOWED:
            found.append(w)
    return sorted(set(found))

# ---------------------------
# Annotate paragraph with markdown highlights
# ---------------------------
def annotate_paragraph(paragraph, typos, pct_issues, number_issues, punct_issues, english_words):
    s = paragraph
    # highlight typos
    for w in sorted(set(typos), key=lambda x: -len(x)):
        # word-boundary replace, case-insensitive approximate by original casing search
        s = re.sub(rf"\b{re.escape(w)}\b", f"**🟡{w}**", s, flags=re.IGNORECASE)
    # highlight percent forms
    for w in pct_issues:
        s = s.replace(w, f"**🟡{w}**")
    for w in number_issues:
        s = s.replace(w, f"**🟡{w}**")
    # punctuation issues are general; we'll note them at paragraph end
    if punct_issues:
        s = s + "  \n" + "  \n" + "**[Punctuation Warnings: " + '; '.join(punct_issues) + "]**"
    # italicize english words
    for w in sorted(set(english_words), key=lambda x: -len(x)):
        s = re.sub(rf"\b{re.escape(w)}\b", f"*{w}*", s)
    return s

# ---------------------------
# Generate Word docx with highlight (yellow) & italic English
# ---------------------------
def docx_from_paragraphs(paragraphs, results):
    doc = Document()
    for i, p in enumerate(paragraphs):
        para = doc.add_paragraph()
        tokens = re.split(r"(\s+)", p)  # keep whitespace
        # tokens are words and spaces; we'll check token content for matches
        for t in tokens:
            run = para.add_run(t)
            # check if this token contains a typo exactly (ignore space tokens)
            txt_clean = re.sub(r"\s+", "", t)
            if not txt_clean:
                continue
            lw = re.sub(r"[^\w%.,]", "", txt_clean).lower()
            # highlight if flagged in results for this paragraph
            res = results.get(i, {})
            if txt_clean in res.get("typos", []) or txt_clean in res.get("pct", []) or txt_clean in res.get("numbers", []):
                # highlight yellow (using highlight_color property if available)
                try:
                    run.font.highlight_color = 7  # 7 == yellow in python-docx
                except Exception:
                    # as fallback, set font color red
                    run.font.color.rgb = RGBColor(255, 0, 0)
            # italicize english words
            if re.fullmatch(r"[A-Za-z]{2,}", txt_clean) and txt_clean.lower() in EN_WORDS:
                run.italic = True
    return doc

# ---------------------------
# Main Streamlit app
# ---------------------------
st.set_page_config(page_title="StatPub Checker (BPS Sidoarjo)", layout="wide")
st.title("StatPub Checker — Tools Pengecekan Publikasi (BPS Sidoarjo)")

st.markdown("""
Aplikasi ini melakukan pemeriksaan kualitas teks publikasi: ejaan Bahasa Indonesia, tanda baca, format angka, dan deteksi istilah bahasa Inggris (italic).  
Upload file (PDF/DOCX/PPTX/TXT), klik **Run all checks**, lalu unduh laporan / Word hasil anotasi.
""")

uploaded = st.file_uploader("Upload dokumen (pdf/docx/pptx/txt)", type=["pdf", "docx", "pptx", "txt"])

if uploaded:
    # extract paragraphs
    ext = uploaded.name.lower().split(".")[-1]
    if ext == "pdf":
        paragraphs = extract_pdf(uploaded)
    elif ext == "docx":
        paragraphs = extract_docx(uploaded)
    elif ext == "pptx":
        paragraphs = extract_pptx(uploaded)
    elif ext == "txt":
        paragraphs = extract_txt(uploaded)
    else:
        st.error("Format file tidak didukung.")
        st.stop()

    # show preview
    st.subheader("Preview (beberapa paragraf pertama)")
    for i, p in enumerate(paragraphs[:6]):
        st.markdown(f"**Paragraf {i+1}:** {normalize_whitespace(p[:500])}")

    if st.button("🔍 Run all checks"):
        st.info("Sedang memproses...")

        results = {}   # index -> dict{typos, pct, numbers, punct, english}
        annotated_md = []

        for i, p in enumerate(paragraphs):
            p_norm = normalize_whitespace(p)
            typos = check_spelling(p_norm)
            punct = check_punctuation_spacing(p_norm)
            pct = re_percent_comma.findall(p_norm) + re_percent_dot.findall(p_norm)
            numbers = re_thousands_comma.findall(p_norm) + re_thousands_dot.findall(p_norm) + re_large_number.findall(p_norm)
            eng = detect_english_words(p_norm)
            results[i] = {"typos": typos, "pct": pct, "numbers": numbers, "punct": punct, "english": eng}
            annotated = annotate_paragraph(p_norm, typos, pct, numbers, punct, eng)
            annotated_md.append((i, annotated))

        # show summary counts
        total_typos = sum(len(v["typos"]) for v in results.values())
        total_pct = sum(len(v["pct"]) for v in results.values())
        total_numbers = sum(len(v["numbers"]) for v in results.values())
        st.success(f"Proses selesai — {total_typos} typo, {total_pct} percent-format issues, {total_numbers} numeric issues ditemukan.")

        # show annotated preview
        st.subheader("Annotated Preview")
        for i, ann in annotated_md[:8]:
            st.markdown(f"**Paragraf {i+1}:**")
            st.markdown(ann)

        # prepare dataframe report
        rows = []
        for i, v in results.items():
            if v["typos"] or v["pct"] or v["numbers"] or v["punct"] or v["english"]:
                rows.append({
                    "paragraph_index": i,
                    "text_preview": normalize_whitespace(paragraphs[i])[:200],
                    "typos": "; ".join(v["typos"]),
                    "pct_issues": "; ".join(v["pct"]),
                    "number_issues": "; ".join(v["numbers"]),
                    "punctuation_issues": "; ".join(v["punct"]),
                    "english_terms": "; ".join(v["english"])
                })
        df_report = pd.DataFrame(rows)
        st.subheader("Detailed Report")
        st.dataframe(df_report)

        # download report csv/xlsx
        csv = df_report.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download CSV Report", csv, file_name=f"report_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv")
        # xlsx
        to_xlsx = io.BytesIO()
        with pd.ExcelWriter(to_xlsx, engine="openpyxl") as writer:
            df_report.to_excel(writer, index=False, sheet_name="report")
        to_xlsx.seek(0)
        st.download_button("⬇️ Download Excel Report", to_xlsx, file_name=f"report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

        # generate annotated docx
        docx_doc = docx_from_paragraphs(paragraphs, results)
        to_docx = io.BytesIO()
        docx_doc.save(to_docx)
        to_docx.seek(0)
        st.download_button("⬇️ Download Annotated DOCX", to_docx, file_name=f"annotated_{os.path.splitext(uploaded.name)[0]}_{datetime.now().strftime('%Y%m%d')}.docx")
