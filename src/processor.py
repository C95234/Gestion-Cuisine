from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
import re
import unicodedata
import datetime as dt

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ✅ AJOUTS (listes déroulantes Excel)
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# PDF (bons de livraison)
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

DAY_NAMES = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]


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


def normalize_regime_label(regime: str) -> str:
    """Normalise les libellés de régimes pour éviter les confusions (casse/accents ignorés).

    - ss / sans / SANS ... -> "SANS"
    - spéciaux / speciaux / spécial / special ... -> "SPÉCIAL"
    - sinon: libellé original nettoyé (clean_text)
    """
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
        # plage "raisonnable" (≈ 2015–2035) pour éviter des décennies de colonnes
        if not (41000 <= v <= 51000):
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
    # handle commas
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def parse_planning_fabrication(path: str, sheet_name: str = "PLANNING FAB") -> Dict[str, pd.DataFrame]:
    """
    Planning fabrication -> {"dejeuner": df, "diner": df}

    Compatible avec:
    - ancien format + format avec ligne d'en-tête des jours
    - effectifs en valeurs ou formules Excel.

    openpyxl ne calcule pas les formules. On récupère donc :
    1) la valeur "figée" si elle existe (data_only=True)
    2) sinon, on évalue un petit sous-ensemble de formules très fréquentes dans les plannings :
       - référence simple : =Feuil!A1 ou ='Feuil'!$A$1
       - somme : =SOMME(ref;ref;...) ou =SUM(ref,ref,...)
       - additions/soustractions : =ref+ref-ref...
    """
    wb_val = openpyxl.load_workbook(path, data_only=True)   # valeurs figées
    wb_fx  = openpyxl.load_workbook(path, data_only=False)  # formules
    if sheet_name not in wb_fx.sheetnames:
        raise ValueError(f"Feuille '{sheet_name}' introuvable. Feuilles dispo: {wb_fx.sheetnames}")
    ws_val = wb_val[sheet_name]
    ws_fx  = wb_fx[sheet_name]

    # --- Helpers ---
    targets = {_norm(d) for d in DAY_NAMES}

    # Capture une référence feuille+cellule, avec guillemets et $ optionnels
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

        # priorité: valeur figée
        if wsV is not None:
            v = wsV[addr0].value
            if v is not None:
                return v

        v2 = wsF[addr0].value
        # Si c'est encore une formule, on tente de l'évaluer (référence/somme/addition)
        if isinstance(v2, str) and v2.startswith("="):
            return _eval_formula(v2, depth + 1)
        return v2

    def _eval_formula(formula: str, depth: int = 0):
        if depth > 25:
            return None
        if not isinstance(formula, str) or not formula.startswith("="):
            return formula

        f = formula.strip()

        # 1) référence simple
        m = _ref_re.match(f)
        if m:
            sheet = m.group(1) or m.group(2)
            addr = m.group(3)
            return _resolve_ref(sheet.strip(), addr, depth + 1)

        # 2) SOMME / SUM
        f_up = f.upper().replace(" ", "")
        if f_up.startswith("=SOMME(") or f_up.startswith("=SUM("):
            inside = f[f.find("(") + 1: f.rfind(")")]
            # séparateurs Excel FR/EN
            parts = re.split(r"[;,]", inside)
            total = 0.0
            for part in parts:
                part = part.strip()
                if not part:
                    continue
                # référence feuille!cellule
                mt = _ref_token_re.search(part)
                if mt:
                    sh = mt.group(1) or mt.group(2)
                    ad = mt.group(3)
                    total += _to_number(_resolve_ref(sh.strip(), ad, depth + 1))
                else:
                    total += _to_number(part)
            return total

        # 3) additions / soustractions simples de références (sans parenthèses)
        #    ex: =Feuil!A1+Feuil!A2-Feuil!A3
        if any(op in f for op in ["+", "-"]) and "(" not in f and ")" not in f:
            expr = f[1:]  # retire "="
            # normalise séparateurs
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

        # fallback: non supporté -> None (sera converti en 0)
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

    # repérage des sections
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


def parse_planning_mixe_lisse(path: str, sheet_name: str = "Planning mixe lisse ") -> Dict[str, pd.DataFrame]:
    """
    Lit la feuille "Planning mixe lisse" (déjeuner + dîner) et retourne:
    {"dejeuner": df, "diner": df} avec colonnes Site, Regime (Mixé/Lissé), Lundi..Dimanche.
    Gère les formules Excel simples: références, SUM(), expressions + / -.
    """
    wb_val = openpyxl.load_workbook(path, data_only=True)
    wb_fx  = openpyxl.load_workbook(path, data_only=False)
    if sheet_name not in wb_fx.sheetnames:
        return {"dejeuner": pd.DataFrame(columns=["Site","Regime"] + DAY_NAMES),
                "diner": pd.DataFrame(columns=["Site","Regime"] + DAY_NAMES)}
    ws_val = wb_val[sheet_name]
    ws_fx  = wb_fx[sheet_name]

    ref_re = re.compile(r"^(?:'([^']+)'|([^'!]+))!?\$?([A-Z]{1,3})\$?(\d{1,7})$")
    sheetref_re = re.compile(r"^=\s*\+?\s*(?:'([^']+)'\s*|([^'!]+))!?\$?([A-Z]{1,3})\$?(\d{1,7})\s*$")
    cell_re = re.compile(r"^\$?([A-Z]{1,3})\$?(\d{1,7})$")

    def addr(col: str, row: int) -> str:
        return f"{col}{row}"

    def _get(sheet: str, a: str):
        # priorité valeur figée
        if sheet in wb_val.sheetnames:
            v = wb_val[sheet][a].value
            if v is not None:
                return v
        return wb_fx[sheet][a].value

    def _eval(sheet: str, a: str, depth: int = 0):
        if depth > 20:
            return None
        v = _get(sheet, a)

        # valeur directe
        if v is None or isinstance(v, (int, float, dt.date, dt.datetime)):
            return v

        if isinstance(v, str) and v.startswith("="):
            s = v.strip()

            # =Feuil!$A$1
            m = sheetref_re.match(s)
            if m:
                sh = m.group(1) or m.group(2)
                col = m.group(3)
                row = int(m.group(4))
                return _eval(sh, addr(col, row), depth + 1)

            # =SUM(C3:C4) / SUM(...); Excel FR possible
            sm = re.match(r"^=\s*(SUM|SOMME)\(([^)]+)\)\s*$", s, re.I)
            if sm:
                rng = sm.group(2).strip()
                # support "C3:C4" (même feuille)
                if ":" in rng:
                    a1, a2 = [x.strip().replace("$","") for x in rng.split(":", 1)]
                    m1 = cell_re.match(a1); m2 = cell_re.match(a2)
                    if m1 and m2:
                        c1, r1 = m1.group(1), int(m1.group(2))
                        c2, r2 = m2.group(1), int(m2.group(2))
                        if c1 == c2:
                            total = 0.0
                            for rr in range(min(r1,r2), max(r1,r2)+1):
                                total += _to_number(_eval(sheet, addr(c1, rr), depth + 1))
                            return total
                parts = [p.strip() for p in re.split(r"[;,]", rng) if p.strip()]
                total = 0.0
                for p in parts:
                    p2 = p.replace("$","")
                    if ":" in p2:
                        continue
                    m1 = cell_re.match(p2)
                    if m1:
                        total += _to_number(_eval(sheet, addr(m1.group(1), int(m1.group(2))), depth + 1))
                return total

            # expressions =+C5+C8-...
            expr = s.lstrip("=").strip()
            expr = expr.replace(";", "+")
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
                t_clean = t.replace("$","")
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
        day_cols = {DAY_NAMES[i]: 3+i for i in range(7)}  # C=3
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

        df = pd.DataFrame(rows, columns=["Site","Regime"] + DAY_NAMES)
        return df

    df_dej = _read_block(start_row=3, header_row=2)
    df_din = _read_block(start_row=19, header_row=18)

    return {"dejeuner": df_dej, "diner": df_din}


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

    Idée clé: on NE suppose PAS que chaque journée fait exactement 12 lignes.
    On scanne la colonne A pour trouver toutes les dates, puis on lit les blocs
    Déjeuner (date_row..date_row+5) et Dîner (date_row+6..date_row+11).
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

    _DESSERT_KW = ("compote", "fruit", "gateau", "gateau", "tarte", "flan", "creme", "crème", "mousse", "riz au lait", "ile flottante")
    _DAIRY_KW = ("fromage", "yaourt", "yogourt", "fromage blanc", "petit suisse", "camembert", "emmental", "kiri", "tartare", "babybel", "gouda", "boursin")

    def _row_score(row: int, kws: tuple[str, ...]) -> int:
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
        candidates = [
            (6, 4, 5),
            (5, 3, 4),
        ]
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

    date_rows: List[tuple[int, dt.date]] = []
    anchor_year: Optional[int] = None
    for r in range(4, ws.max_row + 1):
        raw = ws.cell(r, 1).value
        d = _parse_date(raw, default_year=anchor_year)
        if d is not None:
            anchor_year = d.year
            date_rows.append((r, d))

    seen = set()
    uniq: List[tuple[int, dt.date]] = []
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


# =============================
# Bons de livraison (PDF)
# =============================

def clean_text_delivery(x) -> str:
    """Clean text for delivery notes.

    Differences vs clean_text(): keep asterisks (e.g. *Lasagne*) and em-dashes.
    """
    if x is None:
        return ""
    s = str(x)
    s = s.replace("\u2026", "...")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_menu_delivery(path: str, sheet_name: str = "Feuil2") -> Dict[tuple, List[str]]:
    """Parse menu for delivery notes."""
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

    date_rows: List[tuple[int, dt.date]] = []
    anchor_year: Optional[int] = None
    for r in range(4, ws.max_row + 1):
        raw = ws.cell(r, 1).value
        d = _parse_date(raw, default_year=anchor_year)
        if d is not None:
            anchor_year = d.year
            date_rows.append((r, d))

    seen = set()
    uniq: List[tuple[int, dt.date]] = []
    for rr, dd in date_rows:
        if dd in seen:
            continue
        seen.add(dd)
        uniq.append((rr, dd))

    for rr, dd in uniq:
        read_block(rr, dd, "Déjeuner")
        read_block(rr + 6, dd, "Dîner")

    return out


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


# --- (le reste de tes fonctions PDF / production / grouping est inchangé) ---
# ... (tout ce qui suit est identique à ton fichier, sauf export_excel) ...


_SAUCE_QUALIFIERS = [
    "vinaigrette",
    "mayonnaise",
    "mayo",
    "ketchup",
    "barbecue",
    "bbq",
    "aigre douce",
    "aigre-douce",
    "curry",
    "tomate",
    "fromagere",
    "fromagère",
]


def normalize_produit_for_grouping(produit: str) -> str:
    s = clean_text(produit)
    if not s:
        return ""
    s = re.split(r"\s*\+\s*", s, maxsplit=1)[0].strip()
    s = re.sub(r"\([^)]*\)", "", s).strip()

    low = s.lower()
    low = re.sub(r"\bsauce\b\s+.*$", "", low).strip()
    for q in _SAUCE_QUALIFIERS:
        if re.search(rf"\b{re.escape(q)}\b", low) and len(low.split()) > 1:
            low = re.sub(rf"\b{re.escape(q)}\b.*$", "", low).strip()
            break

    out = low
    out = re.sub(r"\s+", " ", out).strip(" -;,:/")
    if not out:
        out = s
    return out[:1].upper() + out[1:]


def build_bon_commande(planning: Dict[str, pd.DataFrame], menu_items: List[MenuItem]) -> pd.DataFrame:
    # (identique à ta version — je ne touche pas ici)
    def norm_reg(s: str) -> str:
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

    records = []
    for repas_key, repas_label in [("dejeuner", "Déjeuner"), ("diner", "Dîner")]:
        df = planning.get(repas_key)
        if df is None or df.empty:
            continue

        df2 = df.copy()
        df2["Regime"] = df2["Regime"].apply(normalize_regime_label)
        agg = df2.groupby("Regime")[DAY_NAMES].sum(numeric_only=True)
        for jour in DAY_NAMES:
            for regime, nb in agg[jour].items():
                records.append(
                    {
                        "Repas": repas_label,
                        "Jour": jour,
                        "Regime_planning": regime,
                        "reg_key_planning": norm_reg(regime),
                        "Nb_personnes": int(_to_number(nb)),
                    }
                )

    counts = pd.DataFrame(records)
    if counts.empty:
        menu_df = pd.DataFrame(
            [
                {
                    "Date": it.date,
                    "Jour": DAY_NAMES[it.date.weekday()],
                    "Repas": it.repas,
                    "Typologie": it.categorie,
                    "Produit": it.produit,
                    "Effectif": 0,
                    "Coefficient": 1.0,
                }
                for it in menu_items
            ]
        )
        menu_df["Produit"] = menu_df["Produit"].astype(str)
        menu_df["Produit_base"] = menu_df["Produit"].apply(normalize_produit_for_grouping)
        menu_df["Quantité"] = (menu_df["Effectif"] * menu_df["Coefficient"]).astype(int)
        grouped = (
            menu_df.groupby(["Repas", "Typologie", "Produit_base", "Coefficient"], as_index=False)
            .agg(
                {
                    "Jour": lambda s: ", ".join(sorted(set(s), key=lambda x: DAY_NAMES.index(x))),
                    "Produit": "first",
                    "Effectif": "sum",
                    "Quantité": "sum",
                }
            )
            .rename(columns={"Jour": "Jour(s)", "Produit_base": "Produit"})
        )
        return grouped[["Jour(s)", "Repas", "Typologie", "Produit", "Effectif", "Coefficient", "Quantité"]].sort_values(
            ["Repas", "Typologie", "Produit"]
        ).reset_index(drop=True)

    planning_keys = counts[["Regime_planning", "reg_key_planning"]].drop_duplicates().to_dict("records")

    def best_match_planning_key(menu_key: str) -> Optional[str]:
        if not menu_key:
            return None
        mtoks = set(menu_key.split())
        best_key = None
        best_score = -1
        for rec in planning_keys:
            ptoks = set((rec["reg_key_planning"] or "").split())
            score = len(mtoks & ptoks)
            if score > best_score:
                best_score = score
                best_key = rec["reg_key_planning"]
        if best_score <= 0:
            return None
        return best_key

    menu_df = pd.DataFrame(
        [
            {
                "Date": it.date,
                "Jour": DAY_NAMES[it.date.weekday()],
                "Repas": it.repas,
                "Categorie": it.categorie,
                "Regime_menu": it.regime,
                "reg_key_menu": norm_reg(it.regime),
                "Produit": it.produit,
            }
            for it in menu_items
        ]
    )

    menu_df["reg_key_planning"] = menu_df["reg_key_menu"].apply(best_match_planning_key)

    merged = menu_df.merge(
        counts[["Repas", "Jour", "reg_key_planning", "Nb_personnes", "Regime_planning"]],
        on=["Repas", "Jour", "reg_key_planning"],
        how="left",
    )

    merged["Nb_personnes"] = merged["Nb_personnes"].fillna(0).astype(int)
    merged["Coefficient"] = 1.0

    base = merged[
        ["Date", "Jour", "Repas", "Categorie", "Produit", "Nb_personnes", "Coefficient"]
    ].rename(
        columns={"Categorie": "Typologie", "Nb_personnes": "Effectif"}
    )

    base["Produit"] = base["Produit"].astype(str)
    base["Produit_base"] = base["Produit"].apply(normalize_produit_for_grouping)
    base["Quantité"] = (base["Effectif"] * base["Coefficient"]).round().astype(int)

    grouped = (
        base.groupby(["Repas", "Typologie", "Produit_base", "Coefficient"], as_index=False)
        .agg(
            {
                "Jour": lambda s: ", ".join(sorted(set(s), key=lambda x: DAY_NAMES.index(x))),
                "Effectif": "sum",
                "Quantité": "sum",
            }
        )
        .rename(columns={"Jour": "Jour(s)", "Produit_base": "Produit"})
    )

    grouped = grouped[["Jour(s)", "Repas", "Typologie", "Produit", "Effectif", "Coefficient", "Quantité"]]
    return grouped.sort_values(["Repas", "Typologie", "Produit"]).reset_index(drop=True)


def export_excel(
    bon_commande: pd.DataFrame,
    prod_dej: pd.DataFrame,
    prod_din: pd.DataFrame,
    out_path: str,
) -> None:
    """
    ✅ CORRIGÉ :
    - Ajoute les colonnes Fournisseur + Unité si absentes
    - Coefficient en LISTE DÉROULANTE (0.1,0.2,0.25,0.3,1,0.17,0.04,0.08)
    - Fournisseur en LISTE DÉROULANTE (liste complète, pas seulement Sysco)
    - Unité AUTO selon le coefficient (Kg si 0.1/0.2/0.25/0.3 sinon Unité)
    - Quantité AUTO et arrondis : Kg -> 3 décimales ; Unité -> entier
    """
    # --- Garantir colonnes attendues ---
    bc = bon_commande.copy() if isinstance(bon_commande, pd.DataFrame) else pd.DataFrame()

    # Ajoute Fournisseur / Unité si non présents
    if "Fournisseur" not in bc.columns:
        bc["Fournisseur"] = "Sysco"  # défaut
    else:
        bc["Fournisseur"] = bc["Fournisseur"].fillna("").replace("", "Sysco")

    if "Unité" not in bc.columns and "Unite" not in bc.columns:
        bc["Unité"] = ""  # sera calculé en formule Excel
    elif "Unite" in bc.columns and "Unité" not in bc.columns:
        bc = bc.rename(columns={"Unite": "Unité"})

    # Par défaut coef = 1 si manquant
    if "Coefficient" in bc.columns:
        bc["Coefficient"] = pd.to_numeric(bc["Coefficient"], errors="coerce").fillna(1.0)
    else:
        bc["Coefficient"] = 1.0

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        bc.to_excel(writer, sheet_name="Bon de commande", index=False)

        def _pivot_prod(long_df: pd.DataFrame) -> pd.DataFrame:
            if long_df is None or long_df.empty:
                return pd.DataFrame(columns=["Regime"] + DAY_NAMES + ["Total semaine"])
            df = long_df.copy()
            if "Nb" not in df.columns:
                for cand in ["Nb_personnes", "Nombre", "Quantite", "Quantité"]:
                    if cand in df.columns:
                        df = df.rename(columns={cand: "Nb"})
                        break
            piv = df.pivot_table(index="Regime", columns="Jour", values="Nb", aggfunc="sum", fill_value=0)
            for d in DAY_NAMES:
                if d not in piv.columns:
                    piv[d] = 0
            piv = piv[DAY_NAMES]
            piv["Total semaine"] = piv.sum(axis=1)
            piv = piv[piv["Total semaine"] > 0]
            piv = piv.reset_index()
            total_row = pd.DataFrame(
                [["TOTAL JOUR"] + [int(piv[d].sum()) for d in DAY_NAMES] + [int(piv["Total semaine"].sum())]],
                columns=["Regime"] + DAY_NAMES + ["Total semaine"]
            )
            out = pd.concat([piv, total_row], ignore_index=True)
            return out

        dej_piv = _pivot_prod(prod_dej)
        din_piv = _pivot_prod(prod_din)
        dej_piv.to_excel(writer, sheet_name="Déjeuner", index=False)
        din_piv.to_excel(writer, sheet_name="Dîner", index=False)

        wb = writer.book

        # --- Feuille Listes cachée ---
        if "Listes" in wb.sheetnames:
            ws_lists = wb["Listes"]
            # on nettoie le contenu pour éviter les "restes"
            for row in ws_lists.iter_rows():
                for cell in row:
                    cell.value = None
        else:
            ws_lists = wb.create_sheet("Listes")

        # Listes demandées
        coef_list = [0.1, 0.2, 0.25, 0.3, 1, 0.17, 0.04, 0.08]
        kg_set = {0.1, 0.2, 0.25, 0.3}
        suppliers_default = ["Sysco", "Domafrais", "Cercle Vert"]

        # Col A : coefficients
        ws_lists["A1"] = "COEFFICIENTS"
        for i, v in enumerate(coef_list, start=2):
            ws_lists.cell(row=i, column=1).value = float(v)

        # Col C : fournisseurs (liste complète, modifiable côté app si tu l’alimentes ailleurs)
        ws_lists["C1"] = "FOURNISSEURS"
        # si l’app a déjà mis une liste plus large, on la prend, sinon défaut
        suppliers = []
        try:
            suppliers = [str(x).strip() for x in bc["Fournisseur"].dropna().unique().tolist() if str(x).strip()]
        except Exception:
            suppliers = []
        # on veut une liste stable (au moins les 3 + ceux de l'app)
        merged_sup = []
        for s in suppliers_default + suppliers:
            ss = str(s).strip()
            if ss and ss not in merged_sup:
                merged_sup.append(ss)
        if not merged_sup:
            merged_sup = suppliers_default

        for i, v in enumerate(merged_sup, start=2):
            ws_lists.cell(row=i, column=3).value = v

        ws_lists.sheet_state = "hidden"

        ws_bc = wb["Bon de commande"]

        # --- repérage colonnes ---
        headers = {}
        for c in range(1, ws_bc.max_column + 1):
            v = ws_bc.cell(row=1, column=c).value
            if v:
                headers[str(v).strip().lower()] = c

        col_eff = headers.get("effectif")
        col_coef = headers.get("coefficient")
        col_qty = headers.get("quantité") or headers.get("quantite")
        col_unit = headers.get("unité") or headers.get("unite")
        col_supplier = headers.get("fournisseur")

        max_row = ws_bc.max_row

        # --- Data validations (LISTES DÉROULANTES) ---
        # IMPORTANT : on applique la validation sur une PLAGE (pas cellule par cellule)
        if col_coef:
            coef_col_letter = get_column_letter(col_coef)
            # =Listes!$A$2:$A$9
            dv_coef = DataValidation(
                type="list",
                formula1=f"=Listes!$A$2:$A${len(coef_list)+1}",
                allow_blank=True,
                showDropDown=True,
            )
            ws_bc.add_data_validation(dv_coef)
            dv_coef.add(f"{coef_col_letter}2:{coef_col_letter}{max_row}")

        if col_supplier:
            sup_col_letter = get_column_letter(col_supplier)
            dv_sup = DataValidation(
                type="list",
                formula1=f"=Listes!$C$2:$C${len(merged_sup)+1}",
                allow_blank=True,
                showDropDown=True,
            )
            ws_bc.add_data_validation(dv_sup)
            dv_sup.add(f"{sup_col_letter}2:{sup_col_letter}{max_row}")

        # --- Unité automatique selon coefficient ---
        # Kg si coefficient dans {0.1,0.2,0.25,0.3} sinon Unité
        if col_unit and col_coef:
            unit_letter = get_column_letter(col_unit)
            coef_letter = get_column_letter(col_coef)
            for r in range(2, max_row + 1):
                # formule robuste (utilise OR)
                ws_bc.cell(row=r, column=col_unit).value = (
                    f'=IF(OR({coef_letter}{r}=0.1,{coef_letter}{r}=0.2,{coef_letter}{r}=0.25,{coef_letter}{r}=0.3),"Kg","Unité")'
                )

        # --- Quantité automatique + arrondis ---
        if col_eff and col_coef and col_qty:
            eff_letter = get_column_letter(col_eff)
            coef_letter = get_column_letter(col_coef)
            qty_letter = get_column_letter(col_qty)
            unit_letter = get_column_letter(col_unit) if col_unit else None

            for r in range(2, max_row + 1):
                if unit_letter:
                    ws_bc.cell(row=r, column=col_qty).value = (
                        f'=IF({eff_letter}{r}="","",IF({coef_letter}{r}="","",'
                        f'IF({unit_letter}{r}="Kg",ROUND({eff_letter}{r}*{coef_letter}{r},3),ROUND({eff_letter}{r}*{coef_letter}{r},0))))'
                    )
                else:
                    ws_bc.cell(row=r, column=col_qty).value = (
                        f'=IF({eff_letter}{r}="","",IF({coef_letter}{r}="","",ROUND({eff_letter}{r}*{coef_letter}{r},0)))'
                    )

        # --- Mise en forme (inchangée) ---
        thin = Side(style="thin", color="9E9E9E")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        header_fill = PatternFill("solid", fgColor="EDEDED")
        header_font = Font(bold=True)
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
        cell_align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        band_fill = PatternFill("solid", fgColor="F7F7F7")

        for name in ["Bon de commande", "Déjeuner", "Dîner"]:
            ws = wb[name]

            header_row = 1
            if name in ("Déjeuner", "Dîner"):
                ws.insert_rows(1)
                max_col_tmp = ws.max_column
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col_tmp)
                tcell = ws.cell(row=1, column=1)
                tcell.value = f"BON DE COMMANDE – {name.upper()}"
                tcell.font = Font(bold=True, size=14)
                tcell.alignment = Alignment(horizontal="center", vertical="center")
                ws.row_dimensions[1].height = 28
                header_row = 2
                ws.freeze_panes = "B3"
                ws.page_setup.orientation = "landscape"
            else:
                ws.freeze_panes = "A2"

            max_row_ws = ws.max_row
            max_col_ws = ws.max_column

            for c in range(1, max_col_ws + 1):
                cell = ws.cell(row=header_row, column=c)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = border

            ws.row_dimensions[header_row].height = 24

            for r in range(header_row + 1, max_row_ws + 1):
                ws.row_dimensions[r].height = 18
                for c in range(1, max_col_ws + 1):
                    cell = ws.cell(row=r, column=c)
                    if name in ("Déjeuner", "Dîner") and c >= 2:
                        cell.alignment = cell_align_center
                    else:
                        cell.alignment = cell_align
                    cell.border = border

                if r % 2 == 0:
                    for c in range(1, max_col_ws + 1):
                        ws.cell(row=r, column=c).fill = band_fill

            if name in ("Déjeuner", "Dîner"):
                for r in range(header_row + 1, max_row_ws + 1):
                    ws.cell(row=r, column=1).font = Font(bold=True)
                if max_row_ws >= 2 and str(ws.cell(row=max_row_ws, column=1).value).strip().upper() == "TOTAL":
                    for c in range(1, max_col_ws + 1):
                        ws.cell(row=max_row_ws, column=c).font = Font(bold=True)
                        ws.cell(row=max_row_ws, column=c).fill = PatternFill("solid", fgColor="E0E0E0")

            ws.auto_filter.ref = f"A{header_row}:{get_column_letter(max_col_ws)}{max_row_ws}"

            for c_idx in range(1, max_col_ws + 1):
                col_letter = get_column_letter(c_idx)
                max_len = 0
                for r_idx in range(1, min(max_row_ws, 400) + 1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    if cell.value is None:
                        continue
                    max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 60)

            if name in ("Déjeuner", "Dîner"):
                ws.column_dimensions["A"].width = 34
                for idx, _day in enumerate(DAY_NAMES, start=2):
                    ws.column_dimensions[get_column_letter(idx)].width = 12
                ws.column_dimensions[get_column_letter(2 + len(DAY_NAMES))].width = 12

            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
