from __future__ import annotations

"""Facturation PDJ (petit-déjeuner / goûter).

L'app dispose déjà d'une facturation mensuelle (repas + mixé/lissé) basée sur
le planning. Ici, on ajoute une facturation "économat" alimentée par des bons
de commande (PDJ / goûter).

Entrées supportées
------------------
- PDF scanné type "Bon de commande petit déjeuner ..." : on extrait le site via
  l'en-tête (ex: "MAS TOULOUSE LAUTREC") et les quantités via OCR.
- Excel .xls/.xlsx type "Détail Déj-gouter ..." : tableau simple 2 colonnes
  (ingrédient / quantité) avec une ligne de titre contenant le site.

Sortie
------
Un classeur Excel (basé sur le modèle "Facturation économat") complété :
- une ligne par produit (col A)
- une colonne par site (en-têtes ligne 4)
- les quantités sont ajoutées (cumulées) par produit + site.

Notes
-----
Le modèle peut évoluer (produits/prix). Le code :
- crée la ligne si le produit n'existe pas
- ne touche pas aux formules existantes
- met à jour le prix unitaire seulement si un prix est fourni (optionnel).
"""

import datetime as dt
import difflib
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import openpyxl


# --- OCR / PDF helpers (imports optionnels au runtime) ---


def _lazy_import_ocr():
    """Importe les libs OCR uniquement si nécessaires.

    Les PDF *scannés* (avec quantités manuscrites) nécessitent un moteur OCR.
    Ici on s'appuie sur **Tesseract** via `pytesseract`.

    Note : on n'utilise pas pdf2image/Poppler. Les pages PDF sont rendues
    en images via PyMuPDF (fitz), dépendance Python pure.
    """
    import pytesseract  # type: ignore

    # Vérifie que le binaire `tesseract` est bien installé sur la machine.
    try:
        _ = pytesseract.get_tesseract_version()
    except Exception as e:  # pragma: no cover
        raise RuntimeError(
            "Tesseract n'est pas installé (ou pas dans le PATH).\n\n"
            "➡️ Les bons PDF scannés nécessitent Tesseract pour lire les quantités manuscrites.\n"
            "Solutions :\n"
            "1) Installer Tesseract :\n"
            "   - Windows : installeur UB Mannheim, puis ajouter le dossier au PATH\n"
            "   - macOS : `brew install tesseract`\n"
            "   - Linux : `sudo apt-get install tesseract-ocr`\n"
            "2) Ou fournir des bons au format Excel (XLS/XLSX) à la place.\n"
        ) from e

    return pytesseract


def _pdf_to_images(path: str, dpi: int = 220):
    """Rend un PDF en liste d'images PIL via PyMuPDF (sans Poppler)."""
    import fitz  # type: ignore
    from PIL import Image  # type: ignore

    doc = fitz.open(path)
    images = []
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    for page in doc:
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    return images


# -----------------------------
# Data model
# -----------------------------


@dataclass
class PDJLine:
    site: str
    product: str
    qty: float
    unit_price: Optional[float] = None


# -----------------------------
# Normalisation & mapping
# -----------------------------


def _strip_accents(s: str) -> str:
    s = str(s or "")
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))


def norm_text(s: str) -> str:
    return re.sub(r"\s+", " ", _strip_accents(str(s or "")).strip().lower())


def normalize_site_for_economat(site: str) -> str:
    """Aligne des variantes de site avec les en-têtes du modèle économat."""
    s = norm_text(site)
    # Règle existante dans billing.py : 24 ter / 24 simple -> Internat
    if re.search(r"\b24\s*(ter|simple)\b", s):
        return "Internat"
    if "internat" in s:
        return "Internat"

    # Variantes courantes (PDF scanné)
    if "toulouse" in s and "lautrec" in s:
        return "MAS"
    if s in {"mas", "mas toulouse lautrec"}:
        return "MAS"

    # On retourne l'original si non reconnu (l'app tentera de trouver la colonne)
    return site.strip()


def canonical_product_name(raw: str) -> str:
    """Nettoyage léger (on conserve la casse/accents d'origine si possible)."""
    s = str(raw or "").strip()
    s = re.sub(r"\s+", " ", s)
    return s


def explode_composite_quantities(product: str, qty_cell: str) -> List[Tuple[str, float]]:
    """Gère les cellules du type: "5 pommes - 3 oranges - 2 raisin"."""
    txt = norm_text(qty_cell)
    if not txt:
        return []

    out: List[Tuple[str, float]] = []

    # Jus de fruit : "5 pommes - 3 oranges - 2 raisin"
    if "jus" in norm_text(product) or "fruit" in norm_text(product):
        # capture "<n> <mot>" répété
        for m in re.finditer(r"(\d+(?:[\.,]\d+)?)\s*(pomme|pommes|orange|oranges|raisin|raisins)", txt):
            n = float(m.group(1).replace(",", "."))
            fruit = m.group(2)
            if "pomm" in fruit:
                out.append(("Jus Pomme", n))
            elif "orang" in fruit:
                out.append(("Jus Orange", n))
            elif "raisin" in fruit:
                out.append(("Jus Raisin", n))
        if out:
            return out

    # Sinon : tentative simple (un nombre unique)
    m = re.search(r"(\d+(?:[\.,]\d+)?)", txt)
    if m:
        out.append((canonical_product_name(product), float(m.group(1).replace(",", "."))))
    return out


# -----------------------------
# Excel parsing (.xls/.xlsx)
# -----------------------------


def parse_pdj_excel(path: str) -> Tuple[str, List[PDJLine], Optional[str]]:
    """Parse un bon PDJ Excel (souvent .xls legacy).

    Retourne (site, lignes, mois_yyyy_mm optionnel)
    """
    p = Path(path)
    engine = None
    if p.suffix.lower() == ".xls":
        engine = "xlrd"

    df = pd.read_excel(path, sheet_name=0, header=None, engine=engine)

    # Site: chercher une ligne de titre
    site = ""
    for v in df.iloc[:, 0].dropna().astype(str).tolist():
        if "COMMANDE" in v.upper() and "PDJ" in v.upper() or "PETITS-DEJEUNERS" in v.upper():
            site = v
            break
    if not site:
        # fallback : 1ère cellule texte
        for v in df.iloc[:, 0].dropna().astype(str).tolist():
            if len(v.strip()) > 3:
                site = v
                break

    # Extraction du site depuis parenthèses ou mots-clés
    site_clean = site
    m = re.search(r"\(([^\)]+)\)", site)
    if m:
        site_clean = m.group(1)
    site_clean = site_clean.replace('"', "").strip()
    site_clean = normalize_site_for_economat(site_clean)

    # Mois : tenter d'extraire une date "janvier 2026" ou "2026".
    month = None
    for v in df.iloc[:, 0].dropna().astype(str).tolist():
        m2 = re.search(r"\b(\d{1,2})\s*(janvier|fevrier|février|mars|avril|mai|juin|juillet|aout|août|septembre|octobre|novembre|decembre|décembre)\s*(\d{4})\b", norm_text(v))
        if m2:
            y = int(m2.group(3))
            mo_map = {
                "janvier": 1,
                "fevrier": 2,
                "février": 2,
                "mars": 3,
                "avril": 4,
                "mai": 5,
                "juin": 6,
                "juillet": 7,
                "aout": 8,
                "août": 8,
                "septembre": 9,
                "octobre": 10,
                "novembre": 11,
                "decembre": 12,
                "décembre": 12,
            }
            month = f"{y:04d}-{mo_map[m2.group(2)]:02d}"
            break
        m3 = re.search(r"\b(\d{4})\b", v)
        if m3 and not month:
            # on ne devine pas le mois, juste l'année
            pass

    # Données : repérer la ligne d'en-tête "Ingrédients" puis lire 2 colonnes
    start_row = None
    for i in range(len(df)):
        if norm_text(df.iloc[i, 0]).startswith("ingredients"):
            start_row = i + 1
            break
    if start_row is None:
        # fallback : première ligne où col0 est texte et col1 est non vide
        start_row = 0

    lines: List[PDJLine] = []
    for i in range(start_row, len(df)):
        prod = df.iloc[i, 0]
        qty = df.iloc[i, 1] if df.shape[1] > 1 else None
        if pd.isna(prod) or str(prod).strip() == "":
            continue
        prod_s = str(prod).strip()
        if prod_s.lower().startswith("commande"):
            continue
        if pd.isna(qty) or str(qty).strip() == "":
            continue

        for pname, q in explode_composite_quantities(prod_s, str(qty)):
            if q <= 0:
                continue
            lines.append(PDJLine(site=site_clean, product=canonical_product_name(pname), qty=float(q)))

    # Petite règle: "Lait" -> "Lait demi - écrémé" par défaut
    for ln in lines:
        if norm_text(ln.product) == "lait":
            ln.product = "Lait demi - écrémé"

    return site_clean, lines, month


# -----------------------------
# PDF parsing (OCR)
# -----------------------------


_PDF_PRODUCTS_CANON = [
    "Lait demi - écrémé",
    "Lait entier",
    "Céréales",
    "Biscotte",
    "Sucre en sachet",
    "Sucre en morceau",
    "Beurre, plaquettes de 250g",
    "Chocolat en poudre",
    "Brioche",
    "Bledine arome chocolat",
    "Bledine arome vanille",
    "Confiture en carton",
    "Thé en boîte",
    "Café en carton",
    "Jus d'orange",
    "Jus de pomme",
    "Jus de raisin",
    "Fromage blanc pot de 5kg",
    "Mayonnaise",
    "Ketchup",
    "Sel",
    "Poivre",
    "Yaourt Nature",
]


def _parse_date_ddmmyy(text: str) -> Optional[dt.date]:
    """Tente d'extraire une date du type 26.1.26 ou 26/01/2026."""
    t = norm_text(text)
    m = re.search(r"\b(\d{1,2})[\./-](\d{1,2})[\./-](\d{2,4})\b", t)
    if not m:
        return None
    d = int(m.group(1))
    mo = int(m.group(2))
    y = int(m.group(3))
    if y < 100:
        y += 2000
    try:
        return dt.date(y, mo, d)
    except Exception:
        return None


def parse_pdj_pdf(path: str) -> Tuple[str, List[PDJLine], Optional[str]]:
    """Parse un bon PDJ scanné (1 ou plusieurs pages)."""
    pytesseract = _lazy_import_ocr()

    # OpenCV (cv2) est pratique pour isoler l'encre (souvent bleue) sur des formulaires scannés.
    # MAIS on ne doit pas dépendre d'OpenCV : certaines installations (notamment Streamlit Cloud,
    # ou des postes sans OpenCV) n'ont pas le module.
    # => On tente d'importer cv2 et on bascule sur un fallback PIL/numpy si indisponible.
    try:
        import cv2  # type: ignore
    except Exception:  # pragma: no cover
        cv2 = None  # type: ignore
    import numpy as np  # type: ignore

    imgs = _pdf_to_images(path, dpi=220)
    all_lines: List[PDJLine] = []
    month: Optional[str] = None
    site_final = ""

    for img in imgs:
        # 1) Site (plus fiable en OCR texte sur un crop du haut)
        top_crop = img.crop((0, 0, img.size[0], 260))
        top_txt = pytesseract.image_to_string(top_crop, lang="fra", config="--psm 6")
        m_site = re.search(r"petit\s*dejeuner\s*(.+)$", norm_text(top_txt))
        site = site_final or ""
        if m_site:
            site = m_site.group(1)
        site = normalize_site_for_economat(site)
        if site:
            site_final = site

        # 2) Date (zone proche du haut)
        dd = _parse_date_ddmmyy(top_txt)
        if dd:
            month = f"{dd.year:04d}-{dd.month:02d}"

        # 3) OCR en mode "data" pour repérer les lignes produits imprimées
        d = pytesseract.image_to_data(img, lang="fra", output_type=pytesseract.Output.DATAFRAME)
        d = d.dropna(subset=["text"]).copy()
        d["text"] = d["text"].astype(str)
        d = d[d["text"].str.strip() != ""]

        # Zone quantité (colonne centrale) : on crop large et on lit un chiffre manuscrit.
        width, _height = img.size
        qty_x1 = int(width * 0.42)
        qty_x2 = int(width * 0.74)

        def _ocr_hand_qty(y1: int, y2: int) -> Optional[float]:
            """OCR robuste pour chiffres manuscrits (souvent en bleu)."""
            y1 = max(0, y1)
            y2 = min(_height, y2)
            crop = img.crop((qty_x1, y1, qty_x2, y2))
            arr = np.array(crop)

            if cv2 is not None:
                # isolate blue-ish ink in HSV (meilleur sur manuscrit bleu)
                bgr = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
                hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
                lower = np.array([90, 25, 25])
                upper = np.array([145, 255, 255])
                mask = cv2.inRange(hsv, lower, upper)
                kernel = np.ones((3, 3), np.uint8)
                mask = cv2.dilate(mask, kernel, iterations=2)
                img_for_ocr = 255 - mask
            else:
                # Fallback sans OpenCV : on accentue le contraste et on binarise.
                # Moins performant sur encre bleue, mais évite le crash 'No module named cv2'.
                gray = arr.mean(axis=2) if arr.ndim == 3 else arr
                # auto-threshold simple
                thr = gray.mean() * 0.9
                bin_img = (gray < thr).astype("uint8") * 255
                img_for_ocr = 255 - bin_img

            txt = pytesseract.image_to_string(
                img_for_ocr,
                lang="eng",
                config="--psm 7 -c tessedit_char_whitelist=0123456789",
            ).strip()
            m = re.search(r"\d+", txt)
            if not m:
                return None
            try:
                return float(m.group(0))
            except Exception:
                return None

        # 4) Repérer toutes les lignes produits (position y), puis OCR quantité par quantité
        found: List[Tuple[str, int]] = []

        for (_b, _p, _l), g in d.groupby(["block_num", "par_num", "line_num"]):
            words = g.sort_values("left")
            line_txt = " ".join(words["text"].tolist()).strip()
            if not line_txt:
                continue

            line_norm = norm_text(line_txt)
            matched_product = None
            for prod in _PDF_PRODUCTS_CANON:
                pn = norm_text(prod)
                if pn and pn in line_norm:
                    matched_product = prod
                    break
            if not matched_product:
                continue

            y0 = int(words["top"].min())
            found.append((matched_product, y0))

        found = sorted(found, key=lambda t: t[1])
        for i, (prod, y0) in enumerate(found):
            # borne basse/haute : évite de capter le chiffre de la ligne du dessous
            y1 = y0 - 25
            if i + 1 < len(found):
                y_next = found[i + 1][1]
                y2 = int((y0 + y_next) / 2) + 10
            else:
                y2 = y0 + 75

            qty_val = _ocr_hand_qty(y1, y2)
            if qty_val is None or qty_val <= 0:
                continue
            all_lines.append(PDJLine(site=site_final or site or "", product=prod, qty=qty_val))

    if not site_final:
        site_final = ""
    return site_final, all_lines, month


# -----------------------------
# Dispatch parse
# -----------------------------


def parse_pdj_order_file(path: str) -> Tuple[str, List[PDJLine], Optional[str]]:
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        return parse_pdj_pdf(path)
    if ext in {".xls", ".xlsx", ".xlsm"}:
        return parse_pdj_excel(path)
    raise ValueError(f"Format non supporté: {ext}")


# -----------------------------
# Economat workbook update
# -----------------------------


def _find_header_row(ws) -> int:
    # Par défaut le modèle fourni a l'en-tête à la ligne 4
    for r in range(1, 20):
        v = ws.cell(r, 1).value
        if norm_text(v) == "produit":
            return r
    return 4


def _build_site_col_map(ws, header_row: int) -> Dict[str, int]:
    m: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        name = ws.cell(header_row, c).value
        if not name:
            continue
        n = str(name).strip()
        if n in {"Produit", "Unité", "Prix unitaire (€)", "Total produit (€)"}:
            continue
        m[norm_text(n)] = c
    return m


def _build_product_row_map(ws, header_row: int) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is None or str(v).strip() == "":
            continue
        out[norm_text(str(v))] = r
    return out


def _append_product_row(ws, header_row: int, product: str) -> int:
    """Ajoute une ligne produit en bas, en copiant le style de la dernière ligne."""
    # Trouver la dernière ligne produit (col A non vide)
    last = None
    for r in range(ws.max_row, header_row, -1):
        if ws.cell(r, 1).value not in (None, ""):
            last = r
            break
    if last is None:
        last = header_row + 1

    new_r = last + 1
    ws.insert_rows(new_r)

    # Copier styles/formules de la ligne précédente
    from copy import copy
    for c in range(1, ws.max_column + 1):
        src = ws.cell(last, c)
        dst = ws.cell(new_r, c)
        dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.font = copy(src.font)
        dst.alignment = copy(src.alignment)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        # Copie de la formule Total produit si présente (col "Total produit (€)")
        if c == 13 and isinstance(src.value, str) and src.value.startswith("="):
            dst.value = src.value.replace(str(last), str(new_r))
        else:
            dst.value = None

    ws.cell(new_r, 1).value = product
    return new_r


# -----------------------------
# Mapping produit & correction totaux
# -----------------------------

def _normalize_product_key(s: str) -> str:
    return norm_text(s)


# Alias "entrée bon" -> libellé du modèle économat
# (facile à ajuster si les libellés changent)
_PRODUCT_ALIASES = {
    _normalize_product_key("Yaourt Nature"): _normalize_product_key("Yaourt"),
    _normalize_product_key("Lait demi - écrémé"): _normalize_product_key("Lait 1/2 écrémé"),
    _normalize_product_key("Lait demi-écrémé"): _normalize_product_key("Lait 1/2 écrémé"),
    _normalize_product_key("Jus de pomme"): _normalize_product_key("Jus Pomme"),
    _normalize_product_key("Jus d'orange"): _normalize_product_key("Jus Orange"),
    _normalize_product_key("Jus de raisin"): _normalize_product_key("Jus Raisin"),
    _normalize_product_key("Sucre en poudre (sachets)"): _normalize_product_key("Sucre"),
    _normalize_product_key("Compotes pruneaux"): _normalize_product_key("Compotes"),
    _normalize_product_key("Tablettes chocolat"): _normalize_product_key("Chocolat"),
    _normalize_product_key("Nesquick"): _normalize_product_key("Chocolat"),
    _normalize_product_key("Beurre"): _normalize_product_key("Micro-beurres"),
    _normalize_product_key("Biscotte"): _normalize_product_key("Pain de mie complet"),
}


def _best_fuzzy_match(key: str, candidates: List[str]) -> Optional[str]:
    if not candidates:
        return None
    best = None
    best_score = 0.0
    for c in candidates:
        score = difflib.SequenceMatcher(None, key, c).ratio()
        if score > best_score:
            best_score = score
            best = c
    return best if best_score >= 0.78 else None


def _resolve_product_row(prod_key: str, prod_rows: Dict[str, int]) -> Optional[int]:
    # 1) match exact
    if prod_key in prod_rows:
        return prod_rows[prod_key]
    # 2) alias
    alias = _PRODUCT_ALIASES.get(prod_key)
    if alias and alias in prod_rows:
        return prod_rows[alias]
    # 3) fuzzy
    fuzzy = _best_fuzzy_match(prod_key, list(prod_rows.keys()))
    if fuzzy and fuzzy in prod_rows:
        return prod_rows[fuzzy]
    return None


def _ensure_totals_row_formulas(ws, header_row: int, site_cols: Dict[str, int], *, first_product_row: int, last_product_row: int) -> None:
    """Corrige les formules de la ligne 'TOTAL SITE (mensuel)' (le modèle fourni contient des #REF!)."""
    total_row = None
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v and "total site" in norm_text(v):
            total_row = r
            break
    if total_row is None:
        return

    price_col = 3  # C
    price_letter = openpyxl.utils.get_column_letter(price_col)

    for _, col in site_cols.items():
        col_letter = openpyxl.utils.get_column_letter(col)
        ws.cell(total_row, col).value = (
            f"=SUMPRODUCT(${price_letter}${first_product_row}:${price_letter}${last_product_row},"
            f"{col_letter}{first_product_row}:{col_letter}{last_product_row})"
        )

    ws.cell(total_row, 13).value = f"=SUM(M{first_product_row}:M{last_product_row})"

def update_facturation_economat(
    economat_xlsx_path: str,
    pdj_lines: Iterable[PDJLine],
    *,
    force_month: Optional[str] = None,
) -> str:
    """Applique les lignes PDJ au fichier économat (sans casser le modèle).

    - Remplit uniquement les cellules "quantité" (colonnes sites)
    - Ne crée pas de nouvelles lignes dans le tableau principal : si un produit n'est
      pas reconnu, il est listé dans un onglet "Non reconnus".
    - Corrige la ligne 'TOTAL SITE (mensuel)' si le modèle contient des formules #REF!.
    """
    wb = openpyxl.load_workbook(economat_xlsx_path)
    ws = wb[wb.sheetnames[0]]

    header_row = _find_header_row(ws)
    site_cols = _build_site_col_map(ws, header_row)
    prod_rows = _build_product_row_map(ws, header_row)

    # Mois en B2 si présent
    if force_month:
        try:
            ws["B2"].value = force_month
        except Exception:
            pass

    # Détermine le bloc produit (entre l'en-tête et la ligne TOTAL SITE)
    first_product_row = header_row + 1
    last_product_row = ws.max_row
    for r in range(ws.max_row, header_row, -1):
        v = ws.cell(r, 1).value
        if v and "total site" in norm_text(v):
            last_product_row = r - 1
            break

    unknown: List[PDJLine] = []

    for ln in pdj_lines:
        if not ln.product or ln.qty is None:
            continue

        prod_key = norm_text(ln.product)
        row = _resolve_product_row(prod_key, prod_rows)
        if row is None or row < first_product_row or row > last_product_row:
            unknown.append(ln)
            continue

        site = normalize_site_for_economat(ln.site)
        col = site_cols.get(norm_text(site))
        if col is None:
            unknown.append(ln)
            continue

        cell = ws.cell(row, col)
        try:
            old_num = float(cell.value) if cell.value not in (None, "") else 0.0
        except Exception:
            old_num = 0.0
        cell.value = old_num + float(ln.qty)

        # Prix unitaire (col C) si fourni
        if ln.unit_price is not None:
            ws.cell(row, 3).value = float(ln.unit_price)

    _ensure_totals_row_formulas(
        ws,
        header_row,
        site_cols,
        first_product_row=first_product_row,
        last_product_row=last_product_row,
    )

    if unknown:
        if "Non reconnus" in wb.sheetnames:
            unk_ws = wb["Non reconnus"]
            wb.remove(unk_ws)
        unk_ws = wb.create_sheet("Non reconnus")
        unk_ws.append(["Site", "Produit", "Quantité", "Prix unitaire"])
        for u in unknown:
            unk_ws.append([u.site, u.product, u.qty, u.unit_price])

    out_path = str(Path(economat_xlsx_path).with_name("Facturation_economat_PDJ.xlsx"))
    wb.save(out_path)
    return out_path
