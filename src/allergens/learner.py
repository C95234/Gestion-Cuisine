
from __future__ import annotations

from pathlib import Path
from typing import Optional, Dict, Set, List

import pandas as pd
from openpyxl import load_workbook

from .config import ALLERGEN_COLUMNS
from .utils import normalize_key, normalize_space

ALLERGEN_LETTERS = list("CDEFGHIJKLMNOPQR")  # C -> R

def _truthy(v) -> bool:
    s = normalize_space(str(v or "")).strip().lower()
    return s in ("x", "1", "oui", "vrai", "true", "yes")

def extract_reference_from_filled_allergen_workbook(filled_path: str) -> pd.DataFrame:
    """Extrait un référentiel (Plat + colonnes allergènes) depuis un classeur 'Allergène ...' rempli (avec des X)."""
    wb = load_workbook(filled_path, data_only=True)
    rows: List[Dict[str, object]] = []

    for ws in wb.worksheets:
        # zone plats : B4:B27, allergènes : C4:R27
        for r in range(4, 28):
            dish = normalize_space(str(ws[f"B{r}"].value or "")).strip()
            if not dish or dish == "—":
                continue
            rec: Dict[str, object] = {"Plat": dish}
            for i, a in enumerate(ALLERGEN_COLUMNS):
                col = ALLERGEN_LETTERS[i]
                rec[a] = "X" if _truthy(ws[f"{col}{r}"].value) else ""
            rows.append(rec)

    if not rows:
        return pd.DataFrame(columns=["Plat"] + ALLERGEN_COLUMNS)

    df = pd.DataFrame(rows)
    # dédoublonne par plat normalisé, OR sur allergènes
    df["__k"] = df["Plat"].apply(normalize_key)
    agg = []
    for k, g in df.groupby("__k", dropna=False):
        if not k:
            continue
        base = {"Plat": g.iloc[0]["Plat"]}
        for a in ALLERGEN_COLUMNS:
            base[a] = "X" if any(_truthy(x) for x in g[a].tolist()) else ""
        agg.append(base)
    out = pd.DataFrame(agg)
    return out[["Plat"] + ALLERGEN_COLUMNS]


def write_reference_workbook(df: pd.DataFrame, out_path: str, title: str = "Référentiel allergènes") -> str:
    """Écrit un référentiel (Plat + colonnes allergènes) dans un classeur au format 'Allergène ...' (plats en colonne B, X en C:R)."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Référence"

    # Titre
    ws["A1"] = title

    # En-têtes (ligne 3) : A3 libellé, B3 vide, C3..R3 allergènes
    ws["A3"] = "Référence"
    ws["B3"] = None
    for i, a in enumerate(ALLERGEN_COLUMNS):
        col = ALLERGEN_LETTERS[i]
        ws[f"{col}3"] = a

    # Lignes de données à partir de 4
    if df is None or df.empty:
        wb.save(out_path)
        return out_path

    df2 = df.copy()
    # assure colonnes attendues
    if "Plat" not in df2.columns:
        raise ValueError("Le DataFrame de référence doit contenir une colonne 'Plat'.")
    for a in ALLERGEN_COLUMNS:
        if a not in df2.columns:
            df2[a] = ""

    # Dé-doublonnage/tri stable
    df2["__k"] = df2["Plat"].apply(normalize_key)
    df2 = df2[df2["__k"].astype(bool)].copy()
    df2 = df2.sort_values("__k").drop_duplicates("__k", keep="first")

    r = 4
    for _, row in df2.iterrows():
        dish = normalize_space(str(row["Plat"] or "")).strip()
        if not dish:
            continue
        ws[f"B{r}"] = dish
        for i, a in enumerate(ALLERGEN_COLUMNS):
            col = ALLERGEN_LETTERS[i]
            ws[f"{col}{r}"] = "X" if _truthy(row.get(a, "")) else ""
        r += 1

    wb.save(out_path)
    return out_path


def merge_reference(master_path: Optional[str], new_df: pd.DataFrame, out_path: str) -> str:
    """Fusionne new_df dans master (OR) et réécrit le référentiel au format classeur (plats en colonne B, X en C:R)."""

    def _load_master_df(p: str) -> pd.DataFrame:
        # 1) Essai format 'plat en colonnes' (ancien)
        try:
            m = pd.read_excel(p)
            m.columns = [normalize_space(str(c or "")).strip() for c in m.columns]
            if "Plat" in m.columns:
                return m[["Plat"] + [c for c in ALLERGEN_COLUMNS if c in m.columns]].copy()
        except Exception:
            pass
        # 2) Sinon, on considère que c'est un classeur au format 'Allergène ...'
        return extract_reference_from_filled_allergen_workbook(p)

    if master_path and Path(master_path).exists():
        master = _load_master_df(master_path)
    else:
        master = pd.DataFrame(columns=["Plat"] + ALLERGEN_COLUMNS)

    if master is None or master.empty:
        merged = new_df.copy()
    else:
        master = master.copy()
        new_df = new_df.copy()

        # assure colonnes
        if "Plat" not in master.columns:
            master["Plat"] = ""
        for a in ALLERGEN_COLUMNS:
            if a not in master.columns:
                master[a] = ""
            if a not in new_df.columns:
                new_df[a] = ""

        master["__k"] = master["Plat"].apply(normalize_key)
        new_df["__k"] = new_df["Plat"].apply(normalize_key)

        # index master
        master_idx = {k: i for i, k in enumerate(master["__k"].tolist()) if k}

        for _, row in new_df.iterrows():
            k = row.get("__k", "")
            if not k:
                continue
            if k in master_idx:
                i = master_idx[k]
                for a in ALLERGEN_COLUMNS:
                    if _truthy(master.at[i, a]) or _truthy(row.get(a, "")):
                        master.at[i, a] = "X"
            else:
                master_idx[k] = len(master)
                master.loc[len(master)] = {**{c: row.get(c, "") for c in ["Plat"] + ALLERGEN_COLUMNS}, "__k": k}

        merged = master.drop(columns=["__k"], errors="ignore")

    # Réécrit au format attendu (comme ton fichier 'Allergène semaine.xlsx')
    return write_reference_workbook(merged[["Plat"] + ALLERGEN_COLUMNS], out_path)


def learn_from_filled_allergen_workbook(
    filled_path: str,
    master_path: Optional[str],
    out_master_path: str,
) -> str:
    """Point d'entrée: lit le classeur rempli et met à jour le référentiel maître."""
    new_df = extract_reference_from_filled_allergen_workbook(filled_path)
    return merge_reference(master_path, new_df, out_master_path)
