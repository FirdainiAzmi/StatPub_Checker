# app.py
import streamlit as st
import os, io, re, csv
from datetime import datetime
import pandas as pd
import pdfplumber
from docx import Document
from pptx import Presentation
from spellchecker import SpellChecker

# ------------------ Config paths ------------------
BASE_DIR = os.path.dirname(__file__)
WORDLIST_DIR = os.path.join(BASE_DIR, "wordlists")
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(WORDLIST_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)

ID_WORDLIST = os.path.join(WORDLIST_DIR, "id_wordlist.txt")
EN_WORDLIST = os.path.join(WORDLIST_DIR, "en_wordlist.txt")
ALLOWED_WORDS = os.path.join(WORDLIST_DIR, "allowed_words.txt")

UNKNOWN_CSV = os.path.join(DATA_DIR, "unknown_words.csv")
TYPO_LOG_CSV = os.path.join(DATA_DIR, "typo_log.csv")

# ------------------ Helpers: load wordlists ------------------
def load_wordset(path):
    s = set()
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            for line in f:
                w = line.strip()
                if w:
                    s.add(w.lower())
    return s

ID_WORDS = load_wordset(ID_WORDLIST)
EN_WORDS = load_wordset(EN_WORDLIST)
ALLOWED = load_wordset(ALLOWED_WORDS)

# fallback spellchecker
spell = SpellChecker(language=None)
for w in ID_WORDS:
    spell.word_frequency.add(w)
# (optional) add allowed words so they are not marked misspelled
for w in ALLOWED:
    spell.word_frequency.add(w)

# ------------------ Extraction functions ------------------
def extract_docx(file):
    doc = Document(file)
    return [p.text for p in doc.paragraphs if p.text.strip()]

def extract_pdf(file):
    paras = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                # split by double newline to approximate paragraphs
                parts = [p.strip() for p in txt.split("\n\n") if p.strip()]
                if parts:
                    paras.extend(parts)
                else:
                    # fallback line by line
                    paras.extend([ln.strip() for ln in txt.split("\n") if ln.strip()])
    return paras

def extract_pptx(file):
    prs = Presentation(file)
    paras = []
    for slide in prs.slides:
        lines = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                lines.extend([ln.strip() for ln in shape.text.split("\n") if ln.strip()])
        if lines:
            paras.append(" ".join(lines))
    return paras

def extract_txt(file):
    txt = file.read().decode("utf-8", errors="ignore")
    parts = [p.strip() for p in txt.split("\n\n") if p.strip()]
    return parts if parts else [p.strip() for p in txt.split("\n") if p.strip()]

# ------------------ Normalization & tokenization ------------------
def normalize(s):
    s = s.replace("\u200b"," ").replace("\xa0"," ").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def tokenize_words(text):
    tokens = re.findall(r"\b[\w'-]+\b", text, flags=re.UNICODE)
    return tokens

# ------------------ Checkers ------------------
re_percent_comma = re.compile(r"\b\d+,\d+%")
re_percent_dot = re.compile(r"\b\d+\.\d+%")
re_thousands_comma = re.compile(r"\b\d{1,3},\d{3}\b")
re_thousands_dot = re.compile(r"\b\d{1,3}\.\d{3}\b")
re_large_number = re.compile(r"\b\d{4,}\b")

def check_spelling(paragraph):
    tokens = tokenize_words(paragraph)
    typos = []
    unknowns = []
    for t in tokens:
        lw = t.lower()
        if lw in ALLOWED or lw in ID_WORDS or lw in EN_WORDS:
            continue
        # skip pure numbers with punctuation
        if re.fullmatch(r"[\d\.,%]+", lw):
            continue
        # if recognized by spell object -> not typo
        if lw in spell:
            # recognized by fallback, but not in ID_WORDS -> candidate unknown
            if lw not in ID_WORDS:
                unknowns.append(lw)
            continue
        # not recognized => likely typo
        correction = spell.correction(lw) or ""
        typos.append({"word": t, "suggest": correction})
    # dedupe
    typos_unique = {t["word"]:t for t in typos}.values()
    unknowns = sorted(set(unknowns))
    return list(typos_unique), unknowns

def check_punctuation(paragraph):
    issues = []
    if re.search(r"\s,", paragraph):
        issues.append("Spasi sebelum koma.")
    if re.search(r",\S", paragraph):
        # if comma directly followed by letter without space
        issues.append("Tidak ada spasi setelah koma.")
    if re.search(r"\s\.", paragraph):
        issues.append("Spasi sebelum titik.")
    if re.search(r":\S", paragraph):
        issues.append("Tidak ada spasi setelah titik dua.")
    return issues

def check_numbers(paragraph):
    issues = []
    if re_thousands_comma.search(paragraph):
        issues.append("Pemisah ribuan menggunakan koma (mungkin salah).")
    if re_percent_dot.search(paragraph):
        issues.append("Persen menggunakan titik desimal (cek pedoman).")
    for m in re_large_number.finditer(paragraph):
        num = m.group()
        try:
            n = int(num)
            if 1900 <= n <= 2100:
                continue
        except:
            pass
        if len(num) >= 5:
            issues.append(f"Angka besar tanpa pemisah ribuan: {num}")
    return issues

def detect_english(paragraph):
    words = re.findall(r"\b[A-Za-z]{2,}\b", paragraph)
    found = [w for w in set(words) if w.lower() in EN_WORDS and w.lower() not in ALLOWED]
    return sorted(found)

# ------------------ CSV helpers ------------------
def append_unknowns(csv_path, rows):
    # rows: list of dict {word, frequency, first_seen_doc, context}
    if not rows:
        return
    # load existing
    existing = {}
    if os.path.exists(csv_path):
        df = pd.read_csv(csv_path)
        for _, r in df.iterrows():
            existing[r['word']] = r.to_dict()
    # update counts
    for r in rows:
        w = r['word']
        if w in existing:
            existing[w]['frequency'] = int(existing[w]['frequency']) + int(r.get('frequency',1))
        else:
            existing[w] = {'word': w, 'frequency': int(r.get('frequency',1)), 'first_seen_doc': r.get('first_seen_doc',''), 'context': r.get('context','')}
    # write back
    df_out = pd.DataFrame(list(existing.values()))
    df_out.to_csv(csv_path, index=False, encoding='utf-8')

def log_typo(csv_path, docname, para_index, typos):
    # typos: list of dict {word,suggest}
    if not typos:
        return
    rows = []
    if os.path.exists(csv_path):
        df_old = pd.read_csv(csv_path)
        rows = df_old.to_dict('records')
    for t in typos:
        rows.append({"doc":docname, "paragraph_index":para_index, "word":t['word'], "suggest":t.get('suggest','')})
    pd.DataFrame(rows).to_csv(csv_path, index=False, encoding='utf-8')

# ------------------ DOCX annotate generator ------------------
from docx import Document
from docx.shared import RGBColor

def generate_annotated_docx(paragraphs, results):
    doc = Document()
    for i, p in enumerate(paragraphs):
        para = doc.add_paragraph()
        tokens = re.split(r"(\s+)", p)  # preserve spaces
        res = results.get(i, {})
        typos = set([t['word'] for t in res.get('typos',[])])
        unknowns = set(res.get('unknowns',[]))
        pct = set(res.get('pct',[]))
        numbers = set(res.get('numbers',[]))
        eng = set(res.get('english',[]))

        for t in tokens:
            run = para.add_run(t)
            clean = re.sub(r"[^\w%.,]", "", t).lower()
            if t.strip() == "":
                continue
            if t.strip() in typos or clean in typos or t.strip() in pct or clean in numbers:
                try:
                    run.font.highlight_color = 7  # yellow highlight
                except:
                    run.font.color.rgb = RGBColor(255,0,0)
            if clean in eng:
                run.italic = True
    return doc

# ------------------ Streamlit UI ------------------
st.set_page_config(page_title="StatPub Checker (BPS Sidoarjo)", layout="wide")
st.title("StatPub Checker — (Final B)")

st.write("Upload dokumen (.pdf/.docx/.pptx/.txt). Aplikasi akan mengekstrak teks, melakukan spell-check, cek angka & tanda baca, menyimpan kata baru ke `data/unknown_words.csv` (bukan typo).")

uploaded = st.file_uploader("Upload dokumen", type=["pdf","docx","pptx","txt"])
if uploaded:
    st.info(f"File: {uploaded.name}")
    ext = uploaded.name.lower().split(".")[-1]
    if ext == "pdf":
        paras = extract_pdf(uploaded)
    elif ext == "docx":
        paras = extract_docx(uploaded)
    elif ext == "pptx":
        paras = extract_pptx(uploaded)
    elif ext == "txt":
        paras = extract_txt(uploaded)
    else:
        st.error("Format tidak didukung."); st.stop()

    st.subheader("Preview (paragraf pertama)")
    for i,p in enumerate(paras[:5]):
        st.markdown(f"**Par {i+1}:** {normalize(p)[:400]}")

    if st.button("🔍 Run all checks"):
        st.info("Memproses dokumen...")
        results = {}
        unknown_rows = []
        for i,p in enumerate(paras):
            p_norm = normalize(p)
            typos, unknowns = check_spelling(p_norm)
            punct = check_punctuation(p_norm)
            pct = re_percent_comma.findall(p_norm) + re_percent_dot.findall(p_norm)
            numbers = re_thousands_comma.findall(p_norm) + re_thousands_dot.findall(p_norm) + re_large_number.findall(p_norm)
            eng = detect_english(p_norm)
            results[i] = {"typos": typos, "unknowns": unknowns, "punct": punct, "pct": pct, "numbers": numbers, "english": eng}
            # append unknowns to rows for CSV (word-level frequency)
            for u in unknowns:
                unknown_rows.append({"word":u, "frequency":1, "first_seen_doc":uploaded.name, "context": p_norm[:200]})
            # log typos
            if typos:
                log_typo(TYPO_LOG_CSV, uploaded.name, i, typos)

        # aggregate unknown_rows frequencies before append
        df_unknown = pd.DataFrame(unknown_rows)
        if not df_unknown.empty:
            agg = df_unknown.groupby('word', as_index=False).agg({'frequency':'sum','first_seen_doc':'first','context':'first'})
            agg_rows = agg.to_dict('records')
            append_unknowns(UNKNOWN_CSV, agg_rows)

        # Prepare report dataframe
        rows = []
        for i,v in results.items():
            if v['typos'] or v['unknowns'] or v['pct'] or v['numbers'] or v['punct'] or v['english']:
                rows.append({
                    "paragraph_index": i,
                    "preview": normalize(paras[i])[:200],
                    "typos": "; ".join([t['word'] for t in v['typos']]) if v['typos'] else "",
                    "unknowns": "; ".join(v['unknowns']) if v['unknowns'] else "",
                    "pct": "; ".join(v['pct']) if v['pct'] else "",
                    "numbers": "; ".join(v['numbers']) if v['numbers'] else "",
                    "punct": "; ".join(v['punct']) if v['punct'] else "",
                    "english": "; ".join(v['english']) if v['english'] else ""
                })
        df_report = pd.DataFrame(rows)

        st.success("Selesai! Lihat ringkasan di bawah.")
        st.subheader("Annotated Preview (paragraf sample)")
        for i in list(results.keys())[:6]:
            ann = paras[i]
            st.markdown(f"**Par {i+1}:** {normalize(ann)[:500]}")
            if results[i]['typos']:
                st.write("Typos: ", [t['word'] for t in results[i]['typos']])
            if results[i]['unknowns']:
                st.write("Kata baru: ", results[i]['unknowns'])
            if results[i]['english']:
                st.write("English terms: ", results[i]['english'])
            if results[i]['pct'] or results[i]['numbers'] or results[i]['punct']:
                st.write("Numeric/Punct issues: ", results[i]['pct'] + results[i]['numbers'] + results[i]['punct'])

        # downloads
        csv_bytes = df_report.to_csv(index=False).encode('utf-8')
        st.download_button("⬇️ Download CSV Report", csv_bytes, file_name=f"report_{uploaded.name}_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv")
        # excel
        to_xlsx = io.BytesIO()
        with pd.ExcelWriter(to_xlsx, engine="openpyxl") as writer:
            df_report.to_excel(writer, index=False, sheet_name="report")
        to_xlsx.seek(0)
        st.download_button("⬇️ Download Excel Report", to_xlsx, file_name=f"report_{uploaded.name}_{datetime.now().strftime('%Y%m%d')}.xlsx")

        # docx annotated
        docx_obj = generate_annotated_docx(paras, results)
        buf = io.BytesIO()
        docx_obj.save(buf)
        buf.seek(0)
        st.download_button("⬇️ Download Annotated DOCX", buf, file_name=f"annotated_{os.path.splitext(uploaded.name)[0]}_{datetime.now().strftime('%Y%m%d')}.docx")
