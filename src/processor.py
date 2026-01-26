# =========================
# processor.py — PARTIE 1/2
# =========================
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple, Union
import re
import unicodedata
import datetime as dt
from io import BytesIO

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# PDF (bons de livraison)
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


DAY_NAMES = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]


# -------------------------
# Utils
# -------------------------
def clean_text(x) -> str:
    """Normalize cell content into a clean string."""
    if x is None:
        return ""
    s = str(x)
    s = s.replace("\u2026", "...")  # ellipsis
    s = s.replace("*", "")
    s = re.sub(r"\s+", " ", s).strip()
    s = s.strip(" -;,:")
    return s


def _norm(s: str) -> str:
    s = (s or "").lower()
    s = (
        s.replace("é", "e")
        .replace("è", "e")
        .replace("ê", "e")
        .replace("î", "i")
        .replace("ï", "i")
        .replace("ô", "o")
        .replace("à", "a")
        .replace("ç", "c")
    )
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_regime_label(regime: str) -> str:
    """Normalise les libellés de régimes pour éviter les confusions (casse/accents ignorés)."""
    raw = clean_text(regime)
    if not raw:
        return ""
    s = raw.strip().lower()
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    if re.search(r"\b(ss|sans)\b", s):
        return "SANS"
    if re.search(r"\b(hypo|hypocal)\b", s):
        return "Hypocaloriques"
    if "special" in s or re.search(r"\b(speciaux|special|speciale|speciales)\b", s):
        return "SPÉCIAL"

    return raw


def _parse_date(x, default_year: Optional[int] = None) -> Optional[dt.date]:
    if x is None:
        return None

    if isinstance(x, dt.datetime):
        return x.date()
    if isinstance(x, dt.date):
        return x

    # Excel serial date (int/float) -> uniquement si plausible (évite 18278 = 1950)
    if isinstance(x, (int, float)):
        v = float(x)
        if not (41000 <= v <= 51000):  # ~2012..2039
            return None
        try:
            base = dt.date(1899, 12, 30)
            return base + dt.timedelta(days=int(v))
        except Exception:
            return None

    if isinstance(x, str):
        s = x.strip()
        if not s:
            return None
        s = s.replace("-", "/").replace(".", "/")
        m = re.match(r"^(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?$", s)
        if m:
            d = int(m.group(1))
            mo = int(m.group(2))
            y = m.group(3)
            if y is None:
                y = default_year if default_year is not None else dt.date.today().year
            else:
                y = int(y)
                if y < 100:
                    y += 2000
            try:
                return dt.date(int(y), mo, d)
            except Exception:
                return None

    return None


def is_date_cell(x) -> bool:
    return _parse_date(x) is not None


def _to_number(x) -> float:
    """Robust numeric conversion."""
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if not s:
        return 0.0
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def _read_bytes(source: Union[str, BytesIO, object]) -> Optional[bytes]:
    """
    Return bytes if source is an UploadedFile / file-like, else None.
    - If `source` is a path string => None (we will open by path)
    - If it has getbuffer() (Streamlit UploadedFile) => bytes
    - If it has read() => bytes
    """
    if isinstance(source, str):
        return None
    if hasattr(source, "getbuffer"):
        return bytes(source.getbuffer())
    if hasattr(source, "read"):
        try:
            pos = None
            if hasattr(source, "tell") and hasattr(source, "seek"):
                pos = source.tell()
                source.seek(0)
            b = source.read()
            if pos is not None:
                source.seek(pos)
            return b
        except Exception:
            return None
    return None


def _load_workbook_pair(source, data_only_1: bool, data_only_2: bool):
    """
    Charge 2 workbooks depuis:
    - path str  => openpyxl.load_workbook(path,...)
    - uploaded/file-like => on lit les bytes et on charge depuis BytesIO 2 fois
    """
    b = _read_bytes(source)
    if b is None:
        wb1 = openpyxl.load_workbook(source, data_only=data_only_1)
        wb2 = openpyxl.load_workbook(source, data_only=data_only_2)
        return wb1, wb2

    wb1 = openpyxl.load_workbook(BytesIO(b), data_only=data_only_1)
    wb2 = openpyxl.load_workbook(BytesIO(b), data_only=data_only_2)
    return wb1, wb2


# -------------------------
# Planning fabrication
# -------------------------
def parse_planning_fabrication(
    path,
    sheet_name: str = "PLANNING FAB",
) -> Dict[str, pd.DataFrame]:
    """
    Planning fabrication -> {"dejeuner": df, "diner": df}

    Compatible avec:
    - ancien format + format avec ligne d'en-tête des jours
    - effectifs en valeurs ou formules Excel.

    openpyxl ne calcule pas les formules. On récupère donc :
    1) la valeur "figée" si elle existe (data_only=True)
    2) sinon, on évalue un petit sous-ensemble de formules très fréquentes :
       - référence simple : =Feuil!A1 ou ='Feuil'!$A$1
       - somme : =SOMME(...) ou =SUM(...)
       - additions/soustractions : =ref+ref-ref...
    """
    wb_val, wb_fx = _load_workbook_pair(path, data_only_1=True, data_only_2=False)

    if sheet_name not in wb_fx.sheetnames:
        raise ValueError(f"Feuille '{sheet_name}' introuvable. Feuilles dispo: {wb_fx.sheetnames}")
    ws_val = wb_val[sheet_name]
    ws_fx = wb_fx[sheet_name]

    targets = {_norm(d) for d in DAY_NAMES}

    _ref_re = re.compile(
        r"^=\s*(?:\+)?\s*(?:'([^']+)'\s*|([^'!]+))!(\$?[A-Z]{1,3}\$?\d{1,7})\s*$"
    )
    _ref_token_re = re.compile(
        r"(?:'([^']+)'\s*|([^'!+\-*/(),;]+))!(\$?[A-Z]{1,3}\$?\d{1,7})"
    )

    def _norm_cell(x) -> str:
        return _norm(clean_text(x))

    def _strip_dollars(addr: str) -> str:
        return addr.replace("$", "")

    def _resolve_ref(sheet: str, addr: str, depth: int = 0):
        if depth > 25:
            return None
        if sheet not in wb_fx.sheetnames:
            return None
        addr0 = _strip_dollars(addr)
        wsV = wb_val[sheet] if sheet in wb_val.sheetnames else None
        wsF = wb_fx[sheet]

        if wsV is not None:
            v = wsV[addr0].value
            if v is not None:
                return v

        v2 = wsF[addr0].value
        if isinstance(v2, str) and v2.startswith("="):
            return _eval_formula(v2, depth + 1)
        return v2

    def _eval_formula(formula: str, depth: int = 0):
        if depth > 25:
            return None
        if not isinstance(formula, str) or not formula.startswith("="):
            return formula

        f = formula.strip()

        m = _ref_re.match(f)
        if m:
            sheet = m.group(1) or m.group(2)
            addr = m.group(3)
            return _resolve_ref(sheet.strip(), addr, depth + 1)

        f_up = f.upper().replace(" ", "")
        if f_up.startswith("=SOMME(") or f_up.startswith("=SUM("):
            inside = f[f.find("(") + 1 : f.rfind(")")]
            parts = re.split(r"[;,]", inside)
            total = 0.0
            for part in parts:
                part = part.strip()
                if not part:
                    continue
                mt = _ref_token_re.search(part)
                if mt:
                    sh = mt.group(1) or mt.group(2)
                    ad = mt.group(3)
                    total += _to_number(_resolve_ref(sh.strip(), ad, depth + 1))
                else:
                    total += _to_number(part)
            return total

        if any(op in f for op in ["+", "-"]) and "(" not in f and ")" not in f:
            expr = f[1:]
            tokens = re.split(r"(\+|\-)", expr)
            total = 0.0
            sign = +1.0
            first = True
            for t in tokens:
                t = t.strip()
                if not t:
                    continue
                if t == "+":
                    sign = +1.0
                    continue
                if t == "-":
                    sign = -1.0
                    continue

                mt = _ref_token_re.search(t)
                if mt:
                    sh = mt.group(1) or mt.group(2)
                    ad = mt.group(3)
                    val = _to_number(_resolve_ref(sh.strip(), ad, depth + 1))
                else:
                    val = _to_number(t)

                if first:
                    total = val
                    first = False
                else:
                    total = total + sign * val
            return total

        return None

    def get_value(r: int, c: int):
        if not c:
            return None
        v = ws_val.cell(r, c).value
        if v is not None:
            return v
        fx = ws_fx.cell(r, c).value
        if isinstance(fx, str) and fx.startswith("="):
            return _eval_formula(fx, 0)
        return fx

    dejeuner_row = None
    diner_row = None
    for r in range(1, ws_fx.max_row + 1):
        v = ws_fx.cell(r, 1).value
        if isinstance(v, str):
            vv = v.upper()
            if "PLANNING" in vv and "FABRICATION" in vv and "DEJEUN" in vv:
                dejeuner_row = r
            if "PLANNING" in vv and "FABRICATION" in vv and ("DINER" in vv or "DÎNER" in vv):
                diner_row = r
    if dejeuner_row is None or diner_row is None:
        raise ValueError("Impossible de localiser les sections DÉJEUNER / DÎNER (colonne A).")

    def _looks_like_days_header(r: int) -> bool:
        hits = 0
        for c in range(1, ws_fx.max_column + 1):
            if _norm_cell(ws_fx.cell(r, c).value) in targets:
                hits += 1
        return hits >= 3

    def _find_header_row_near(title_row: int) -> Optional[int]:
        for r in range(title_row + 1, min(ws_fx.max_row, title_row + 10) + 1):
            if _looks_like_days_header(r):
                return r
        return None

    def _get_day_columns(header_row: int) -> Dict[str, int]:
        cols: Dict[str, int] = {}
        for c in range(1, ws_fx.max_column + 1):
            v = _norm_cell(ws_fx.cell(header_row, c).value)
            for d in DAY_NAMES:
                if v == _norm(d):
                    cols[d] = c
        return cols

    def _find_total_col(header_row: int) -> Optional[int]:
        for c in range(1, ws_fx.max_column + 1):
            v = _norm_cell(ws_fx.cell(header_row, c).value)
            if v == "total" or "total" in v:
                return c
        return None

    def read_block(title_row: int, end_row: int) -> pd.DataFrame:
        header_r = _find_header_row_near(title_row)

        if header_r is None:
            day_cols = {d: 3 + i for i, d in enumerate(DAY_NAMES)}  # C..I
            total_col = 10  # J
            data_start = title_row + 1
        else:
            day_cols = _get_day_columns(header_r)
            total_col = _find_total_col(header_r)
            data_start = header_r + 1

        rows = []
        current_site = None

        for r in range(data_start, end_row + 1):
            rowA = ws_fx.cell(r, 1).value

            if isinstance(rowA, str):
                up = rowA.upper()
                if "PLANNING" in up and "FABRICATION" in up:
                    break

            row_vals = [ws_fx.cell(r, c).value for c in range(1, 12)]
            if all(v is None or (isinstance(v, str) and not v.strip()) for v in row_vals):
                continue

            if isinstance(rowA, str) and rowA.strip().upper() == "TOTAL":
                break
            if isinstance(rowA, str) and _norm(rowA.strip()) == "site":
                continue

            site = ws_fx.cell(r, 1).value
            regime = ws_fx.cell(r, 2).value

            if site is not None and clean_text(site):
                current_site = clean_text(site)

            regime_txt = normalize_regime_label(regime)
            if not regime_txt:
                continue

            vals = []
            for d in DAY_NAMES:
                raw = get_value(r, day_cols.get(d, 0))
                vals.append(int(_to_number(raw)))

            if total_col:
                total = _to_number(get_value(r, total_col))
            else:
                total = float(sum(vals))

            rows.append([current_site or "", regime_txt, *vals, total])

        cols = ["Site", "Regime", *DAY_NAMES, "Total"]
        df = pd.DataFrame(rows, columns=cols)
        if not df.empty:
            df["Regime"] = df["Regime"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
        return df

    df_dej = read_block(dejeuner_row, diner_row - 1)
    df_din = read_block(diner_row, ws_fx.max_row)
    return {"dejeuner": df_dej, "diner": df_din}


# -------------------------
# Planning mixé/lissé
# -------------------------
def parse_planning_mixe_lisse(path, sheet_name: str = "Planning mixe lisse ") -> Dict[str, pd.DataFrame]:
    """
    Lit la feuille "Planning mixe lisse" (déjeuner + dîner) et retourne:
    {"dejeuner": df, "diner": df} avec colonnes Site, Regime (Mixé/Lissé), Lundi..Dimanche.
    Gère les formules simples: références, SUM(), expressions + / -.
    """
    wb_val, wb_fx = _load_workbook_pair(path, data_only_1=True, data_only_2=False)

    if sheet_name not in wb_fx.sheetnames:
        return {
            "dejeuner": pd.DataFrame(columns=["Site", "Regime"] + DAY_NAMES),
            "diner": pd.DataFrame(columns=["Site", "Regime"] + DAY_NAMES),
        }
    ws_val = wb_val[sheet_name]
    ws_fx = wb_fx[sheet_name]

    ref_re = re.compile(r"^(?:'([^']+)'|([^'!]+))!?\$?([A-Z]{1,3})\$?(\d{1,7})$")
    sheetref_re = re.compile(r"^=\s*\+?\s*(?:'([^']+)'\s*|([^'!]+))!?\$?([A-Z]{1,3})\$?(\d{1,7})\s*$")
    cell_re = re.compile(r"^\$?([A-Z]{1,3})\$?(\d{1,7})$")

    def addr(col: str, row: int) -> str:
        return f"{col}{row}"

    def _get(sheet: str, a: str):
        if sheet in wb_val.sheetnames:
            v = wb_val[sheet][a].value
            if v is not None:
                return v
        return wb_fx[sheet][a].value

    def _eval(sheet: str, a: str, depth: int = 0):
        if depth > 20:
            return None
        v = _get(sheet, a)

        if v is None or isinstance(v, (int, float, dt.date, dt.datetime)):
            return v

        if isinstance(v, str) and v.startswith("="):
            s = v.strip()

            m = sheetref_re.match(s)
            if m:
                sh = m.group(1) or m.group(2)
                col = m.group(3)
                row = int(m.group(4))
                return _eval(sh, addr(col, row), depth + 1)

            sm = re.match(r"^=\s*SUM\(([^)]+)\)\s*$", s, re.I)
            if sm:
                rng = sm.group(1).strip()
                if ":" in rng:
                    a1, a2 = [x.strip().replace("$", "") for x in rng.split(":", 1)]
                    m1 = cell_re.match(a1)
                    m2 = cell_re.match(a2)
                    if m1 and m2:
                        c1, r1 = m1.group(1), int(m1.group(2))
                        c2, r2 = m2.group(1), int(m2.group(2))
                        if c1 == c2:
                            total = 0.0
                            for rr in range(min(r1, r2), max(r1, r2) + 1):
                                total += _to_number(_eval(sheet, addr(c1, rr), depth + 1))
                            return total
                parts = [p.strip() for p in re.split(r"[;,]", rng) if p.strip()]
                total = 0.0
                for p in parts:
                    p2 = p.replace("$", "")
                    if ":" in p2:
                        continue
                    m1 = cell_re.match(p2)
                    if m1:
                        total += _to_number(_eval(sheet, addr(m1.group(1), int(m1.group(2))), depth + 1))
                return total

            expr = s.lstrip("=").strip().replace(";", "+")
            tokens = re.split(r"(\+|\-)", expr)
            total = 0.0
            sign = +1.0
            for t in tokens:
                t = t.strip()
                if not t:
                    continue
                if t == "+":
                    sign = +1.0
                    continue
                if t == "-":
                    sign = -1.0
                    continue

                t_clean = t.replace("$", "")
                mm = ref_re.match(t_clean)
                if mm:
                    sh = mm.group(1) or mm.group(2) or sheet
                    col = mm.group(3)
                    row = int(mm.group(4))
                    total += sign * _to_number(_eval(sh, addr(col, row), depth + 1))
                    continue
                mc = cell_re.match(t_clean)
                if mc:
                    total += sign * _to_number(_eval(sheet, addr(mc.group(1), int(mc.group(2))), depth + 1))
                    continue
                total += sign * _to_number(t_clean)
            return total

        return v

    def _read_block(start_row: int, header_row: int) -> pd.DataFrame:
        day_cols = {DAY_NAMES[i]: 3 + i for i in range(7)}  # C..I
        rows = []
        cur_site = None
        r = start_row
        while r <= ws_fx.max_row:
            a = clean_text(ws_fx.cell(r, 1).value)
            if a.upper() == "TOTAL":
                break
            if not a and not clean_text(ws_fx.cell(r, 2).value):
                r += 1
                continue

            if a:
                cur_site = a
            reg = clean_text(ws_fx.cell(r, 2).value)
            if not reg or reg.lower() not in {"mixe", "mixé", "lisse", "lissé"}:
                r += 1
                continue

            reg_out = "Mixé" if "mix" in _norm(reg) else "Lissé"
            vals = []
            for d, c in day_cols.items():
                col_letter = ws_fx.cell(header_row, c).column_letter
                v = _eval(sheet_name, addr(col_letter, r))
                vals.append(int(_to_number(v)))

            rows.append([cur_site or "", reg_out, *vals])
            r += 1

        return pd.DataFrame(rows, columns=["Site", "Regime"] + DAY_NAMES)

    df_dej = _read_block(start_row=3, header_row=2)
    df_din = _read_block(start_row=19, header_row=18)
    return {"dejeuner": df_dej, "diner": df_din}


# -------------------------
# Menu parsing
# -------------------------
@dataclass
class MenuItem:
    date: dt.date
    repas: str  # "Déjeuner" or "Dîner"
    categorie: str  # Entrée, Plat, Laitage, Dessert
    regime: str
    produit: str


def parse_menu(path: str, sheet_name: str = "Feuil2") -> List[MenuItem]:
    """
    Parse menu excel.
    - on scanne la colonne A pour trouver toutes les dates
    - on lit chaque bloc repas en détectant 5 ou 6 lignes (selon présence 2 ou 3 lignes de plat)
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Feuille '{sheet_name}' introuvable. Feuilles dispo: {wb.sheetnames}")
    ws = wb[sheet_name]

    group = {c: clean_text(ws.cell(2, c).value) for c in range(2, ws.max_column + 1)}
    header = {c: clean_text(ws.cell(3, c).value) for c in range(2, ws.max_column + 1)}

    regime_by_col: Dict[int, str] = {}
    for c in range(2, ws.max_column + 1):
        h = header.get(c, "")
        g = group.get(c, "")
        if not h and not g:
            continue
        if not h:
            label = g
        elif not g:
            label = h
        elif g and g.upper() not in h.upper():
            label = f"{g} - {h}"
        else:
            label = h

        label = label.replace("STANDART", "STANDARD")
        label = re.sub(r"\s+", " ", label).strip()
        regime_by_col[c] = label

    items: List[MenuItem] = []

    def _norm_kw(s: str) -> str:
        s = clean_text(s).lower()
        s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
        return s

    _DESSERT_KW = (
        "compote", "fruit", "gateau", "gâteau", "tarte", "flan", "creme", "crème",
        "mousse", "riz au lait", "ile flottante", "île flottante"
    )
    _DAIRY_KW = (
        "fromage", "yaourt", "yogourt", "fromage blanc", "petit suisse",
        "camembert", "emmental", "kiri", "tartare", "babybel", "gouda", "boursin"
    )

    def _row_score(row: int, kws: Tuple[str, ...]) -> int:
        score = 0
        for c in regime_by_col.keys():
            v = clean_text(ws.cell(row, c).value)
            if not v:
                continue
            t = _norm_kw(v)
            if any(k in t for k in kws):
                score += 1
        return score

    def detect_block_height(start_row: int) -> int:
        candidates = [(6, 4, 5), (5, 3, 4)]
        best_h = 6
        best_score = -1
        for h, off_lait, off_des in candidates:
            laitage_row = start_row + off_lait
            dessert_row = start_row + off_des
            if dessert_row > ws.max_row:
                continue
            score = _row_score(laitage_row, _DAIRY_KW) + _row_score(dessert_row, _DESSERT_KW)
            if _row_score(dessert_row, _DESSERT_KW) > 0:
                score += 2
            if score > best_score:
                best_score = score
                best_h = h
        return best_h

    def read_block(start_row: int, date_val: dt.date, repas: str, block_height: int):
        entree_row = start_row
        if block_height == 5:
            plat_rows = [start_row + 1, start_row + 2]
            laitage_row = start_row + 3
            dessert_row = start_row + 4
        else:
            plat_rows = [start_row + 1, start_row + 2, start_row + 3]
            laitage_row = start_row + 4
            dessert_row = start_row + 5

        for c, regime in regime_by_col.items():
            entree = clean_text(ws.cell(entree_row, c).value)
            plat_parts = [clean_text(ws.cell(rr, c).value) for rr in plat_rows]
            plat_parts = [p for p in plat_parts if p]
            laitage = clean_text(ws.cell(laitage_row, c).value)
            dessert = clean_text(ws.cell(dessert_row, c).value)

            if entree:
                items.append(MenuItem(date_val, repas, "Entrée", regime, entree))
            for p in plat_parts:
                items.append(MenuItem(date_val, repas, "Plat", regime, p))
            if laitage:
                items.append(MenuItem(date_val, repas, "Laitage", regime, laitage))
            if dessert:
                items.append(MenuItem(date_val, repas, "Dessert", regime, dessert))

    date_rows: List[Tuple[int, dt.date]] = []
    anchor_year: Optional[int] = None
    for r in range(4, ws.max_row + 1):
        raw = ws.cell(r, 1).value
        d = _parse_date(raw, default_year=anchor_year)
        if d is not None:
            anchor_year = d.year
            date_rows.append((r, d))

    seen = set()
    uniq: List[Tuple[int, dt.date]] = []
    for rr, dd in date_rows:
        if dd in seen:
            continue
        seen.add(dd)
        uniq.append((rr, dd))

    for rr, dd in uniq:
        h_dej = detect_block_height(rr)
        read_block(rr, dd, "Déjeuner", h_dej)

        rr_din = rr + h_dej
        if rr_din <= ws.max_row:
            h_din = detect_block_height(rr_din)
            read_block(rr_din, dd, "Dîner", h_din)

    return items


# -------------------------
# Production tables
# -------------------------
def make_production_summary(df_planning: pd.DataFrame) -> pd.DataFrame:
    """Long format: Jour, Regime, Nb"""
    if df_planning is None or df_planning.empty:
        return pd.DataFrame(columns=["Jour", "Regime", "Nb"])

    long = df_planning.melt(
        id_vars=["Site", "Regime"],
        value_vars=DAY_NAMES,
        var_name="Jour",
        value_name="Nb",
    )
    long["Nb"] = pd.to_numeric(long["Nb"], errors="coerce").fillna(0).astype(int)
    long["Regime"] = long["Regime"].apply(normalize_regime_label)

    out = long.groupby(["Jour", "Regime"], as_index=False)["Nb"].sum()
    out["Nb"] = out["Nb"].astype(int)
    return out


def make_production_pivot(df_planning: pd.DataFrame) -> pd.DataFrame:
    """Pivot: lignes = régimes, colonnes = jours + Total"""
    if df_planning is None or df_planning.empty:
        cols = ["Regime", *DAY_NAMES, "Total"]
        return pd.DataFrame(columns=cols)

    long = make_production_summary(df_planning)
    if long.empty:
        cols = ["Regime", *DAY_NAMES, "Total"]
        return pd.DataFrame(columns=cols)

    pivot = (
        long.pivot_table(index="Regime", columns="Jour", values="Nb", aggfunc="sum", fill_value=0)
        .reindex(columns=DAY_NAMES, fill_value=0)
        .reset_index()
    )
    pivot["Total"] = pivot[DAY_NAMES].sum(axis=1).astype(int)

    total_row = {"Regime": "TOTAL"}
    for j in DAY_NAMES:
        total_row[j] = int(pivot[j].sum())
    total_row["Total"] = int(pivot["Total"].sum())
    pivot = pd.concat([pivot, pd.DataFrame([total_row])], ignore_index=True)

    for j in DAY_NAMES + ["Total"]:
        pivot[j] = pd.to_numeric(pivot[j], errors="coerce").fillna(0).astype(int)

    return pivot


# -------------------------
# Bons de livraison PDF
# -------------------------
def clean_text_delivery(x) -> str:
    """Clean text for delivery notes (keep asterisks)."""
    if x is None:
        return ""
    s = str(x)
    s = s.replace("\u2026", "...")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_menu_delivery(path: str, sheet_name: str = "Feuil2") -> Dict[tuple, List[str]]:
    """
    Output: dict[(date, repas, regime)] -> 5 lines (Entrée, Plat1, Plat2, Laitage, Dessert)
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Feuille '{sheet_name}' introuvable. Feuilles dispo: {wb.sheetnames}")
    ws = wb[sheet_name]

    group = {c: clean_text_delivery(ws.cell(2, c).value) for c in range(2, ws.max_column + 1)}
    header = {c: clean_text_delivery(ws.cell(3, c).value) for c in range(2, ws.max_column + 1)}

    regime_by_col: Dict[int, str] = {}
    for c in range(2, ws.max_column + 1):
        h = header.get(c, "")
        if not h:
            continue
        g = group.get(c, "")
        if g and g.upper() not in h.upper():
            label = f"{g} - {h}"
        else:
            label = h
        label = label.replace("STANDART", "STANDARD")
        label = re.sub(r"\s+", " ", label).strip()
        regime_by_col[c] = label

    out: Dict[tuple, List[str]] = {}

    def read_block(start_row: int, date_val: dt.date, repas: str):
        for c, regime in regime_by_col.items():
            entree = clean_text_delivery(ws.cell(start_row + 0, c).value)
            plat1 = clean_text_delivery(ws.cell(start_row + 1, c).value)
            plat2a = clean_text_delivery(ws.cell(start_row + 2, c).value)
            plat2b = clean_text_delivery(ws.cell(start_row + 3, c).value)
            plat2_parts = [p for p in [plat2a, plat2b] if p]
            plat2 = " / ".join(plat2_parts)
            laitage = clean_text_delivery(ws.cell(start_row + 4, c).value)
            dessert = clean_text_delivery(ws.cell(start_row + 5, c).value)

            lines = [entree, plat1, plat2, laitage, dessert]
            lines = [ln if ln else "—" for ln in lines]
            out[(date_val, repas, regime)] = lines

    date_rows: List[Tuple[int, dt.date]] = []
    anchor_year: Optional[int] = None
    for r in range(4, ws.max_row + 1):
        raw = ws.cell(r, 1).value
        d = _parse_date(raw, default_year=anchor_year)
        if d is not None:
            anchor_year = d.year
            date_rows.append((r, d))

    seen = set()
    uniq: List[Tuple[int, dt.date]] = []
    for rr, dd in date_rows:
        if dd in seen:
            continue
        seen.add(dd)
        uniq.append((rr, dd))

    for rr, dd in uniq:
        read_block(rr, dd, "Déjeuner")
        read_block(rr + 6, dd, "Dîner")

    return out


def export_bons_livraison_pdf(
    planning: Dict[str, pd.DataFrame],
    menu_path: str,
    out_pdf_path: str,
    planning_path: Optional[str] = None,
    sheet_menu: str = "Feuil2",
    sites_exclus: Optional[List[str]] = None,
) -> None:
    """Generate delivery notes PDF."""
    sites_exclus = sites_exclus or ["24 ter", "24 simple", "IME TL"]
    excl_norm = {_norm(s) for s in sites_exclus}

    menu_lines = parse_menu_delivery(menu_path, sheet_name=sheet_menu)

    def _dayname(d: dt.date) -> str:
        return DAY_NAMES[d.weekday()]

    def _site_day_total(df: pd.DataFrame, site: str, day_name: str) -> int:
        if df is None or df.empty:
            return 0
        tmp = df.copy()
        tmp["Site"] = tmp["Site"].astype(str).str.strip().str.lower()
        s = (site or "").strip().lower()
        sub = tmp[tmp["Site"] == s]
        if sub.empty or day_name not in sub.columns:
            return 0
        return int(pd.to_numeric(sub[day_name], errors="coerce").fillna(0).sum())

    all_dates = sorted({d for d in (_parse_date(k[0]) for k in menu_lines.keys()) if d is not None})

    sites = set()
    for key in ["dejeuner", "diner"]:
        df = planning.get(key)
        if df is not None and not df.empty:
            sites |= set(df["Site"].astype(str).tolist())
    sites_list = [s for s in sites if _norm(s) not in excl_norm]

    preferred = ["FAM", "Bussière", "Bruyère", "FO", "FDP", "MAS", "ESAT", "FM", "René"]
    pref_norm = [_norm(x) for x in preferred]

    def site_rank(s: str) -> tuple:
        ns = _norm(s)
        for idx, pn in enumerate(pref_norm):
            if pn in {"fo", "fdp"}:
                if ns == "fo" or ns == "fdp" or ns.startswith("fo ") or ns.startswith("fdp ") or " fo" in ns or " fdp" in ns:
                    return (idx, s)
            if ns == pn or ns.startswith(pn + " ") or (" " + pn) in ns:
                return (idx, s)
        return (999, s)

    sites_sorted = sorted(sites_list, key=site_rank)

    SITE_LABELS = {"fo": "Foyer Du Près", "fdp": "Foyer Du Près", "fm": "Foyer Fernand Marlier"}

    def display_site_name(s: str) -> str:
        ns = _norm(s)
        for k, v in SITE_LABELS.items():
            if ns == k or ns.startswith(k + " ") or (" " + k) in ns:
                return v
        return s

    order = ["hypocalorique", "speciaux", "sans", "lisse", "mixe", "standard", "vegetarien"]

    def regime_sort_key(reg: str):
        n = _norm(reg)
        for i, tok in enumerate(order):
            if tok in n:
                return (i, n)
        return (999, n)

    def is_mixe_lisse(reg: str) -> bool:
        n = _norm(reg)
        if "mixe" in n or "lisse" in n:
            return True
        if n in {"ml", "m l"} or " ml " in f" {n} ":
            return True
        return False

    def jour_col(date_val: dt.date) -> str:
        return DAY_NAMES[date_val.weekday()]

    def count_for(df: pd.DataFrame, site: str, regime: str, date_val: dt.date) -> int:
        if df is None or df.empty:
            return 0
        col = jour_col(date_val)
        sub = df[(df["Site"] == site) & (df["Regime"] == regime)]
        if sub.empty:
            return 0
        return int(pd.to_numeric(sub[col], errors="coerce").fillna(0).sum())

    def regimes_for_site(date_val: dt.date, site: str) -> List[str]:
        regs = set()
        for key in ["dejeuner", "diner"]:
            df = planning.get(key)
            if df is None or df.empty:
                continue
            col = jour_col(date_val)
            sub = df[df["Site"] == site]
            if sub.empty:
                continue
            sub = sub[pd.to_numeric(sub[col], errors="coerce").fillna(0) > 0]
            regs |= set(sub["Regime"].astype(str).tolist())
        return sorted(regs, key=regime_sort_key)

    def fmt_date_fr(d: dt.date) -> str:
        jours = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
        mois = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
        return f"{jours[d.weekday()]} {d.day} {mois[d.month-1]} {d.year}"

    W, H = A4
    x0 = 34
    y_top = H - 18
    line_h = 12

    mix_planning = None
    if planning_path:
        try:
            mix_planning = parse_planning_mixe_lisse(planning_path)
        except Exception:
            mix_planning = None

    c = canvas.Canvas(out_pdf_path, pagesize=A4)

    def _wrap_text(text: str, font_name: str, font_size: int, max_width: float) -> List[str]:
        words = (text or "").split()
        if not words:
            return [""]
        out_lines: List[str] = []
        cur = words[0]
        for w in words[1:]:
            test = cur + " " + w
            if c.stringWidth(test, font_name, font_size) <= max_width:
                cur = test
            else:
                out_lines.append(cur)
                cur = w
        out_lines.append(cur)
        return out_lines

    def _draw_bullet(x: float, y: float, text: str, max_width: float, font_size: int = 9, leading: float = 11) -> float:
        bullet = "• "
        font_name = "Helvetica"
        c.setFont(font_name, font_size)
        first_indent = c.stringWidth(bullet, font_name, font_size)
        lines = _wrap_text(text, font_name, font_size, max_width - first_indent) or [""]
        c.drawString(x, y, bullet + lines[0])
        y -= leading
        for ln in lines[1:]:
            c.drawString(x + first_indent, y, ln)
            y -= leading
        return y

    def draw_page(site: str, date_val: dt.date):
        c.setFont("Helvetica-Bold", 15)
        c.drawString(x0, y_top, "BON DE LIVRAISON")

        y = y_top - 24
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x0, y, "Site : ")
        c.setFont("Helvetica", 9)
        c.drawString(x0 + 27.5, y, site)

        y_min = 60
        max_width = (W - x0) - (x0 + 30)

        def _redraw_header_suite():
            c.setFont("Helvetica-Bold", 15)
            c.drawString(x0, y_top, "BON DE LIVRAISON (suite)")
            yy = y_top - 24
            c.setFont("Helvetica-Bold", 9)
            c.drawString(x0, yy, "Site : ")
            c.setFont("Helvetica", 9)
            c.drawString(x0 + 27.5, yy, site)
            yy -= 12
            c.setFont("Helvetica-Bold", 9)
            c.drawString(x0, yy, "Date : ")
            c.setFont("Helvetica", 9)
            c.drawString(x0 + 27.5, yy, fmt_date_fr(date_val))
            return yy - 8

        def ensure_space(needed_height: float):
            nonlocal y
            if y - needed_height < y_min:
                c.showPage()
                y = _redraw_header_suite()

        y -= line_h
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x0, y, "Date : ")
        c.setFont("Helvetica", 9)
        c.drawString(x0 + 27.5, y, fmt_date_fr(date_val))

        tournee = "Barquette" if _norm(site) == _norm("MAS") else "Camion"
        y -= line_h
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x0, y, "Tournée : ")
        c.setFont("Helvetica", 9)
        c.drawString(x0 + 44, y, tournee)

        df_dej = planning.get("dejeuner")
        df_din = planning.get("diner")
        col = jour_col(date_val)

        tot_dej = int(pd.to_numeric(df_dej[df_dej["Site"] == site][col], errors="coerce").fillna(0).sum()) if df_dej is not None and not df_dej.empty else 0
        tot_din = int(pd.to_numeric(df_din[df_din["Site"] == site][col], errors="coerce").fillna(0).sum()) if df_din is not None and not df_din.empty else 0

        mix_dej = lisse_dej = mix_din = lisse_din = 0
        if mix_planning is not None:
            try:
                mdej = mix_planning.get("dejeuner")
                mdin = mix_planning.get("diner")
                if mdej is not None and not mdej.empty:
                    sub = mdej[mdej["Site"].astype(str).str.strip().str.lower() == str(site).strip().lower()]
                    if not sub.empty:
                        v_l = sub[sub["Regime"].astype(str).str.lower().str.contains("liss")][col].sum()
                        v_m = sub[sub["Regime"].astype(str).str.lower().str.contains("mix")][col].sum()
                        lisse_dej = int(_to_number(v_l)); mix_dej = int(_to_number(v_m))
                if mdin is not None and not mdin.empty:
                    sub = mdin[mdin["Site"].astype(str).str.strip().str.lower() == str(site).strip().lower()]
                    if not sub.empty:
                        v_l = sub[sub["Regime"].astype(str).str.lower().str.contains("liss")][col].sum()
                        v_m = sub[sub["Regime"].astype(str).str.lower().str.contains("mix")][col].sum()
                        lisse_din = int(_to_number(v_l)); mix_din = int(_to_number(v_m))
            except Exception:
                pass

        tot_dej_all = tot_dej + mix_dej + lisse_dej
        tot_din_all = tot_din + mix_din + lisse_din

        y -= line_h * 1.6
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x0, y, f"Total Déjeuner : {tot_dej_all}    Total Dîner : {tot_din_all}")

        if (mix_dej + lisse_dej + mix_din + lisse_din) > 0:
            y -= 12
            c.setFont("Helvetica-Bold", 9)
            c.drawString(
                x0, y,
                f"Détail Mixé/Lissé (inclus dans le total) — Déj: Mixé {mix_dej} / Lissé {lisse_dej}   |   Dîn: Mixé {mix_din} / Lissé {lisse_din}"
            )
            c.setFont("Helvetica", 9)

        y -= 24
        box_h = 28
        gap = 20
        box_w = (W - 2 * x0 - gap) / 2
        box_top = y

        c.rect(x0, box_top - box_h, box_w, box_h, stroke=1, fill=0)
        c.rect(x0 + box_w + gap, box_top - box_h, box_w, box_h, stroke=1, fill=0)

        c.setFont("Helvetica-Bold", 9)
        c.drawString(x0 + 6, box_top - 12, "Température départ : ____ °C")
        c.drawString(x0 + box_w + gap + 6, box_top - 12, "Température réception : ____ °C")

        y = box_top - box_h - 10

        regs = regimes_for_site(date_val, site)
        for reg in regs:
            dej_n = count_for(df_dej, site, reg, date_val)
            din_n = count_for(df_din, site, reg, date_val)
            if dej_n <= 0 and din_n <= 0:
                continue

            y -= 14
            c.setFont("Helvetica-Bold", 9)
            c.drawString(x0, y, f"{reg} —  Déj {dej_n} / Dîn {din_n}")

            y -= 12
            c.setFont("Helvetica-Bold", 9)
            c.drawString(x0 + 12, y, "Déjeuner")
            y -= 12
            c.setFont("Helvetica", 9)

            reg_norm = _norm(reg)
            if is_mixe_lisse(reg):
                c.drawString(x0 + 30, y, f"• Quantité mixé/lissé à livrer : {dej_n}")
                y -= 12
            else:
                lines = menu_lines.get((date_val, "Déjeuner", reg))
                if not lines:
                    target = set(reg_norm.split())
                    best = None
                    best_score = -1
                    for (d, repas, rlabel), lns in menu_lines.items():
                        if d != date_val or repas != "Déjeuner":
                            continue
                        score = len(target & set(_norm(rlabel).split()))
                        if score > best_score:
                            best_score = score
                            best = lns
                    lines = best
                lines = lines or ["Menu non détaillé"]
                for ln in lines:
                    ensure_space(14)
                    y = _draw_bullet(x0 + 30, y, ln, max_width, font_size=9, leading=11)

            y -= 6
            c.setFont("Helvetica-Bold", 9)
            c.drawString(x0 + 12, y, "Dîner")
            y -= 12
            c.setFont("Helvetica", 9)

            if is_mixe_lisse(reg):
                c.drawString(x0 + 30, y, f"• Quantité mixé/lissé à livrer : {din_n}")
                y -= 12
            else:
                lines = menu_lines.get((date_val, "Dîner", reg))
                if not lines:
                    target = set(reg_norm.split())
                    best = None
                    best_score = -1
                    for (d, repas, rlabel), lns in menu_lines.items():
                        if d != date_val or repas != "Dîner":
                            continue
                        score = len(target & set(_norm(rlabel).split()))
                        if score > best_score:
                            best_score = score
                            best = lns
                    lines = best
                lines = lines or ["Menu non détaillé"]
                for ln in lines:
                    ensure_space(14)
                    y = _draw_bullet(x0 + 30, y, ln, max_width, font_size=9, leading=11)

            y -= 6

        c.setLineWidth(1)
        c.line(x0, 46, W - x0, 46)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x0, 30, "Chauffeur (signature) : ____________________")
        c.drawString(W / 2 + 20, 30, "Réception (signature) : ____________________")

    for site in sites_sorted:
        for d in all_dates:
            day = _dayname(d)
            tot_dej = _site_day_total(planning.get("dejeuner"), site, day)
            tot_din = _site_day_total(planning.get("diner"), site, day)
            if (tot_dej + tot_din) <= 0:
                continue
            draw_page(display_site_name(site), d)
            c.showPage()

    c.save()
