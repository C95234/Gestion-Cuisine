from __future__ import annotations

from pathlib import Path
from typing import Optional, Dict, List

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

    # OR logique par plat normalisé
    out = (
        df
        .groupby("__k", as_index=False)
        .agg({a: "max" for a in ALLERGEN_COLUMNS})
    )

    out["Plat"] = (
        df.groupby("__k")["Plat"]
        .first()
        .values
    )

    return out[["Plat"] + ALLERGEN_COLUMNS]


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

    ws["A3"] = "Référence"
    ws["B3"] = None
    for i, a in enumerate(ALLERGEN_COLUMNS):
        ws[f"{ALLERGEN_LETTERS[i]}3"] = a

    if df is None or df.empty:
        wb.save(out_path)
        return out_path

    df2 = df.copy()

    if "Plat" not in df2.columns:
        raise ValueError("Le DataFrame doit contenir une colonne 'Plat'.")

    for a in ALLERGEN_COLUMNS:
        if a not in df2.columns:
            df2[a] = ""

    df2["__k"] = df2["Plat"].apply(normalize_key)
    df2 = df2[df2["__k"].astype(bool)]

    # ⚠️ OR FINAL — NE JAMAIS DÉDOUBLER AVANT
    df2 = (
        df2
        .groupby("__k", as_index=False)
        .agg({a: "max" for a in ALLERGEN_COLUMNS})
    )

    df2["Plat"] = (
        df.groupby(df["Plat"].apply(normalize_key))["Plat"]
        .first()
        .values
    )

    r = 4
    for _, row in df2.iterrows():
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

    def _load_master_df(p: str) -> pd.DataFrame:
        try:
            m = pd.read_excel(p)
            m.columns = [normalize_space(str(c or "")).strip() for c in m.columns]
            if "Plat" in m.columns:
                return m[["Plat"] + ALLERGEN_COLUMNS]
        except Exception:
            pass
        return extract_reference_from_filled_allergen_workbook(p)

    if master_path and Path(master_path).exists():
        master = _load_master_df(master_path)
    else:
        master = pd.DataFrame(columns=["Plat"] + ALLERGEN_COLUMNS)

    df = pd.concat([master, new_df], ignore_index=True)
    df["__k"] = df["Plat"].apply(normalize_key)
    df = df[df["__k"].astype(bool)]

    merged = (
        df
        .groupby("__k", as_index=False)
        .agg({a: "max" for a in ALLERGEN_COLUMNS})
    )

    merged["Plat"] = (
        df.groupby("__k")["Plat"]
        .first()
        .values
    )

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
