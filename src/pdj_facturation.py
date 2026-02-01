import pandas as pd
import pytesseract
import fitz
from PIL import Image
import io
import re

pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"

SITE_COLS = {
    "FAM TL": "D", "Bussière": "E", "Bruyères": "F",
    "FO": "G", "MAS": "H", "ESAT": "I",
    "FM": "J", "René Lalouette": "K", "Internat": "L"
}

def detect_site(text, filename=""):
    t = (text + filename).lower()
    if "lautrec" in t or "mas" in t: return "MAS"
    if "24" in t or "internat" in t: return "Internat"
    return None

def pdf_text(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    txt = ""
    for p in doc:
        pix = p.get_pixmap()
        img = Image.open(io.BytesIO(pix.tobytes()))
        txt += pytesseract.image_to_string(img, lang="fra")
    return txt

def parse_pdf(file):
    text = pdf_text(file)
    site = detect_site(text, file.name)
    rows = re.findall(r"([A-Za-zéèàùêôîç ]+)\s+(\d+)\s+\d*", text)
    return site, rows

def parse_excel(file):
    df = pd.read_excel(file)
    site = detect_site(" ".join(df.columns), file.name)
    rows = []
    for _, r in df.iterrows():
        for c in df.columns:
            if isinstance(r[c], (int, float)) and r[c] > 0:
                rows.append((c, int(r[c])))
    return site, rows

def aggregate(files):
    data = {}
    for f in files:
        site, rows = parse_pdf(f) if f.name.endswith(".pdf") else parse_excel(f)
        if not site: continue
        data.setdefault(site, {})
        for prod, q in rows:
            data[site][prod] = data[site].get(prod, 0) + q
    return data

def update_facturation_economat(model_file, data):
    wb = pd.ExcelWriter("Facturation_economat_PDJ.xlsx", engine="openpyxl")
    df = pd.read_excel(model_file)
    for site, produits in data.items():
        col = SITE_COLS.get(site)
        if not col: continue
        for p, q in produits.items():
            idx = df["Produit"].str.contains(p, case=False, na=False)
            if idx.any():
                df.loc[idx, site] = df.loc[idx, site].fillna(0) + q
    df.to_excel(wb, index=False)
    wb.close()
    return "Facturation_economat_PDJ.xlsx"
