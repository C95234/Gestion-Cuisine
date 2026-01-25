
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

def merge_reference(master_path: Optional[str], new_df: pd.DataFrame, out_path: str) -> str:
    """Fusionne new_df dans master (OR)."""
    if master_path and Path(master_path).exists():
        master = pd.read_excel(master_path)
    else:
        master = pd.DataFrame(columns=["Plat"] + ALLERGEN_COLUMNS)

    if master.empty:
        merged = new_df.copy()
        merged.to_excel(out_path, index=False)
        return out_path

    # normalisation
    master = master.copy()
    master["__k"] = master["Plat"].apply(normalize_key)
    new_df = new_df.copy()
    new_df["__k"] = new_df["Plat"].apply(normalize_key)

    # index
    master_idx = {k: i for i, k in enumerate(master["__k"].tolist()) if k}

    for _, row in new_df.iterrows():
        k = row["__k"]
        if not k:
            continue
        if k in master_idx:
            i = master_idx[k]
            for a in ALLERGEN_COLUMNS:
                if _truthy(row.get(a)):
                    master.at[i, a] = "X"
        else:
            master = pd.concat([master, pd.DataFrame([row])], ignore_index=True)
            master_idx[k] = len(master) - 1

    master = master.drop(columns=["__k"], errors="ignore")
    # ordonne colonnes
    cols = ["Plat"] + ALLERGEN_COLUMNS
    master = master[cols]
    master.to_excel(out_path, index=False)
    return out_path

def learn_from_filled_allergen_workbook(
    filled_path: str,
    master_path: Optional[str],
    out_master_path: str,
) -> str:
    """Point d'entrée: lit le classeur rempli et met à jour le référentiel maître."""
    new_df = extract_reference_from_filled_allergen_workbook(filled_path)
    return merge_reference(master_path, new_df, out_master_path)
