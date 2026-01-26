from __future__ import annotations

from pathlib import Path
from typing import Optional, Dict, List, Tuple

import pandas as pd
from openpyxl import load_workbook

from .config import ALLERGEN_COLUMNS
from .utils import normalize_key, normalize_space

ALLERGEN_LETTERS = list("CDEFGHIJKLMNOPQR")  # C -> R


def _truthy(v) -> bool:
    s = normalize_space(str(v or "")).strip().lower()
    return s in ("x", "1", "oui", "vrai", "true", "yes")


# ----------------------------------------------------------------------
# EXTRACTION depuis un classeur rempli
# ----------------------------------------------------------------------
def extract_reference_from_filled_allergen_workbook(filled_path: str) -> pd.DataFrame:
    """Extrait un référentiel (Plat + colonnes allergènes) depuis un classeur 'Allergène ...' rempli."""
    wb = load_workbook(filled_path, data_only=True)
    rows: List[Dict[str, object]] = []

    for ws in wb.worksheets:
        for r in range(4, 28):
            dish = normalize_space(str(ws[f"B{r}"].value or "")).strip()
            if not dish or dish == "—":
                continue

            rec = {"Plat": dish}
            for i, a in enumerate(ALLERGEN_COLUMNS):
                col = ALLERGEN_LETTERS[i]
                rec[a] = "X" if _truthy(ws[f"{col}{r}"].value) else ""
            rows.append(rec)

    if not rows:
        return pd.DataFrame(columns=["Plat"] + ALLERGEN_COLUMNS)

    df = pd.DataFrame(rows)
    df["__k"] = df["Plat"].apply(normalize_key)

    out = (
        df.groupby("__k", as_index=False)
        .agg({a: "max" for a in ALLERGEN_COLUMNS})
    )
    out["Plat"] = df.groupby("__k")["Plat"].first().values

    return out[["Plat"] + ALLERGEN_COLUMNS]


# ----------------------------------------------------------------------
# LECTURE ROBUSTE DU RÉFÉRENTIEL MAÎTRE (TON FORMAT)
# ----------------------------------------------------------------------
def _read_master_reference_xlsx(path: str) -> pd.DataFrame:
    """
    Lit le référentiel maître au format généré par write_reference_workbook():
    - titre en ligne 1
    - en-têtes en ligne 3
    - plats en colonne B (B3 parfois vide dans l'ancien writer)
    - allergènes en colonnes C..R
    """
    wb = load_workbook(path, data_only=True)

    # On cherche la feuille "Référence" sinon la 1ère
    ws = wb["Référence"] if "Référence" in wb.sheetnames else wb.worksheets[0]

    # En-tête en ligne 3
    header_row = 3

    # Colonnes: Plat en B, allergènes en C.. selon ALLERGEN_LETTERS
    # On lit jusqu'à un maximum raisonnable (évite d’embarquer 5000 lignes vides)
    rows: List[Dict[str, object]] = []
    max_scan = min(ws.max_row, 2000)

    for r in range(header_row + 1, max_scan + 1):
        dish = normalize_space(str(ws[f"B{r}"].value or "")).strip()
        if not dish:
            continue

        rec: Dict[str, object] = {"Plat": dish}
        for i, a in enumerate(ALLERGEN_COLUMNS):
            col = ALLERGEN_LETTERS[i]  # C..R
            rec[a] = "X" if _truthy(ws[f"{col}{r}"].value) else ""
        rows.append(rec)

    if not rows:
        return pd.DataFrame(columns=["Plat"] + ALLERGEN_COLUMNS)

    df = pd.DataFrame(rows)
    return _sanitize_reference_df(df)


def _sanitize_reference_df(df: pd.DataFrame) -> pd.DataFrame:
    """Assure la présence de Plat + colonnes allergènes et normalise les colonnes."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Plat"] + ALLERGEN_COLUMNS)

    df2 = df.copy()
    df2.columns = [normalize_space(str(c or "")).strip() for c in df2.columns]

    # Tolérances : certains fichiers peuvent utiliser "Produit"
    if "Plat" not in df2.columns and "Produit" in df2.columns:
        df2 = df2.rename(columns={"Produit": "Plat"})

    if "Plat" not in df2.columns:
        # tente de retrouver une colonne "Unnamed: 1" etc qui contient le plat
        # heuristique: 1ère colonne texte non-allergène avec le plus de valeurs non vides
        candidates = []
        for c in df2.columns:
            if c in ALLERGEN_COLUMNS:
                continue
            non_empty = df2[c].astype(str).map(lambda x: normalize_space(x).strip()).astype(bool).sum()
            candidates.append((non_empty, c))
        candidates.sort(reverse=True)
        if candidates and candidates[0][0] > 0:
            df2 = df2.rename(columns={candidates[0][1]: "Plat"})

    if "Plat" not in df2.columns:
        return pd.DataFrame(columns=["Plat"] + ALLERGEN_COLUMNS)

    for a in ALLERGEN_COLUMNS:
        if a not in df2.columns:
            df2[a] = ""

    df2["Plat"] = df2["Plat"].astype(str).map(lambda x: normalize_space(x).strip())
    df2 = df2[df2["Plat"].astype(bool)]

    return df2[["Plat"] + ALLERGEN_COLUMNS]


def _load_master_df(master_path: str) -> pd.DataFrame:
    """
    Charge un master quel que soit son format:
    1) TON format "Référence" (openpyxl)
    2) un excel plat+colonnes allergènes classique (pandas)
    3) un classeur allergènes rempli (fallback)
    """
    # 1) Essai: ton format "Référence"
    try:
        df = _read_master_reference_xlsx(master_path)
        if df is not None and not df.empty and "Plat" in df.columns:
            return df
    except Exception:
        pass

    # 2) Essai: excel plat+colonnes allergènes "classique"
    try:
        m = pd.read_excel(master_path)
        m = _sanitize_reference_df(m)
        if m is not None and not m.empty and "Plat" in m.columns:
            return m
    except Exception:
        pass

    # 3) Fallback: classeur allergènes rempli
    try:
        return extract_reference_from_filled_allergen_workbook(master_path)
    except Exception:
        return pd.DataFrame(columns=["Plat"] + ALLERGEN_COLUMNS)


# ----------------------------------------------------------------------
# ÉCRITURE DU RÉFÉRENTIEL (SANS DÉTRUIRE L’APPRENTISSAGE)
# ----------------------------------------------------------------------
def write_reference_workbook(
    df: pd.DataFrame,
    out_path: str,
    title: str = "Référentiel allergènes"
) -> str:
    """Écrit un référentiel allergènes au format classeur."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Référence"

    ws["A1"] = title

    # ✅ IMPORTANT: on met un vrai header "Plat"
    ws["A3"] = "Référence"
    ws["B3"] = "Plat"
    for i, a in enumerate(ALLERGEN_COLUMNS):
        ws[f"{ALLERGEN_LETTERS[i]}3"] = a

    if df is None or df.empty:
        wb.save(out_path)
        return out_path

    df2 = _sanitize_reference_df(df)
    if df2.empty:
        wb.save(out_path)
        return out_path

    df2["__k"] = df2["Plat"].apply(normalize_key)
    df2 = df2[df2["__k"].astype(bool)]

    # OR FINAL
    merged = (
        df2.groupby("__k", as_index=False)
        .agg({a: "max" for a in ALLERGEN_COLUMNS})
    )
    merged["Plat"] = df2.groupby("__k")["Plat"].first().values

    r = 4
    for _, row in merged.iterrows():
        ws[f"B{r}"] = row["Plat"]
        for i, a in enumerate(ALLERGEN_COLUMNS):
            ws[f"{ALLERGEN_LETTERS[i]}{r}"] = "X" if _truthy(row[a]) else ""
        r += 1

    wb.save(out_path)
    return out_path


# ----------------------------------------------------------------------
# FUSION MAÎTRE + NOUVEL APPRENTISSAGE
# ----------------------------------------------------------------------
def merge_reference(
    master_path: Optional[str],
    new_df: pd.DataFrame,
    out_path: str
) -> str:
    """Fusionne new_df dans master (OR logique)."""

    master = pd.DataFrame(columns=["Plat"] + ALLERGEN_COLUMNS)
    if master_path and Path(master_path).exists():
        master = _load_master_df(master_path)

    new_df = _sanitize_reference_df(new_df)
    master = _sanitize_reference_df(master)

    df = pd.concat([master, new_df], ignore_index=True)

    if df.empty or "Plat" not in df.columns:
        return write_reference_workbook(pd.DataFrame(columns=["Plat"] + ALLERGEN_COLUMNS), out_path)

    df["__k"] = df["Plat"].apply(normalize_key)
    df = df[df["__k"].astype(bool)]

    merged = (
        df.groupby("__k", as_index=False)
        .agg({a: "max" for a in ALLERGEN_COLUMNS})
    )
    merged["Plat"] = df.groupby("__k")["Plat"].first().values

    return write_reference_workbook(
        merged[["Plat"] + ALLERGEN_COLUMNS],
        out_path
    )


# ----------------------------------------------------------------------
# POINT D’ENTRÉE
# ----------------------------------------------------------------------
def learn_from_filled_allergen_workbook(
    filled_path: str,
    master_path: Optional[str],
    out_master_path: str,
) -> str:
    """Lit un classeur rempli et met à jour le référentiel maître."""
    new_df = extract_reference_from_filled_allergen_workbook(filled_path)
    return merge_reference(master_path, new_df, out_master_path)
