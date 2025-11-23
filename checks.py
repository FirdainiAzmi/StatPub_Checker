import re
from spellchecker import SpellChecker

spell = SpellChecker(language='id')

ENGLISH_WORDS = {"survey", "data", "analysis", "trend", "economic", "population"}  # nanti bisa diperbesar

def check_typo(text):
    words = re.findall(r"\b\w+\b", text)
    typos = [w for w in words if w.lower() in spell.unknown([w.lower()])]
    return typos

def check_comma_spacing(text):
    errors = []
    pattern = r",\S"  # koma langsung ketemu huruf tanpa spasi
    for m in re.finditer(pattern, text):
        errors.append((m.group(), m.start()))
    return errors

def check_number_format(text):
    errors = []
    # persen harus "20%" bukan "20 %"
    if re.search(r"\d+\s%", text):
        errors.append("Persen salah format (tidak boleh ada spasi sebelum %)")
    
    # ribuan harus titik: 1.000 bukan 1,000
    if re.search(r"\d{1,3},\d{3}", text):
        errors.append("Gunakan titik (.) sebagai pemisah ribuan, bukan koma (,)")
    
    return errors

def italic_english(text):
    def repl(match):
        word = match.group()
        return f"*{word}*"
    pattern = r"\b(" + "|".join(ENGLISH_WORDS) + r")\b"
    return re.sub(pattern, repl, text, flags=re.IGNORECASE)
