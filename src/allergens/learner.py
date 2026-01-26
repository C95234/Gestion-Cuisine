
from __future__ import annotations

from pathlib import Path
from typing import Optional, Dict, Set, List, Tuple

import pandas as pd
from openpyxl import load_workbook

from .config import ALLERGEN_COLUMNS
from .utils import normalize_key, normalize_space

ALLERGEN_LETTERS = list("CDEFGHIJKLMNOPQR")  # C -> R
MEAT_ROWS = [33, 35, 37]  # B33/B35/B37 + origines en C/H/N


def _truthy(v) -> bool:
    s = normalize_space(str(v or "")).strip().lower()
    return s in ("x", "1", "oui", "vrai", "true", "yes")

def extract_reference_from_filled_allergen_workbook(filled_path: str) -> pd.DataFrame:
    """Extrait un référentiel (Plat + colonnes allergènes) depuis un classeur 'Allergène ...' rempli (avec des X)."""
    wb = load_workbook(filled_path, data_only=True)
    rows: List[Dict[str, object]] = []

    for ws in wb.worksheets:
        # Le classeur maître peut contenir une feuille "Origines" qu'il ne faut pas interpréter
        if ws.title.strip().lower() in ("origines",):
            continue
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


def extract_meat_origins_from_filled_allergen_workbook(filled_path: str) -> pd.DataFrame:
    """Extrait un référentiel "origines des viandes" depuis un classeur allergènes rempli.

    Dans le template, la zone "Origine des viandes" se trouve en bas de la feuille :
      - plats (avec type de viande) en B33/B35/B37
      - Naissance en C33/C35/C37
      - Élevage en H33/H35/H37
      - Abattage en N33/N35/N37

    Retourne un DataFrame : ["Viande -plat", "Naissance", "Élevage", "Abattage"].
    """
    wb = load_workbook(filled_path, data_only=True)
    rows: List[Dict[str, object]] = []

    for ws in wb.worksheets:
        for r in MEAT_ROWS:
            dish = normalize_space(str(ws[f"B{r}"].value or "")).strip()
            if not dish or dish == "—":
                continue
            naissance = normalize_space(str(ws[f"C{r}"].value or "")).strip()
            elevage = normalize_space(str(ws[f"H{r}"].value or "")).strip()
            abattage = normalize_space(str(ws[f"N{r}"].value or "")).strip()
            if not any([naissance, elevage, abattage]):
                # Rien de renseigné, on ignore pour éviter d'écraser le maître
                continue
            rows.append(
                {
                    "Viande -plat": dish,
                    "Naissance": naissance,
                    "Élevage": elevage,
                    "Abattage": abattage,
                }
            )

    if not rows:
        return pd.DataFrame(columns=["Viande -plat", "Naissance", "Élevage", "Abattage"])

    df = pd.DataFrame(rows)
    df["__k"] = df["Viande -plat"].apply(normalize_key)
    df = df[df["__k"].astype(bool)].copy()

    # dédoublonne (dernier non-vide par champ)
    agg: List[Dict[str, object]] = []
    for k, g in df.groupby("__k", dropna=False):
        if not k:
            continue
        base = {
            "Viande -plat": g.iloc[0]["Viande -plat"],
            "Naissance": "",
            "Élevage": "",
            "Abattage": "",
        }
        for col in ("Naissance", "Élevage", "Abattage"):
            vals = [normalize_space(str(x or "")).strip() for x in g[col].tolist()]
            vals = [v for v in vals if v]
            if vals:
                base[col] = vals[-1]
        agg.append(base)

    out = pd.DataFrame(agg)
    return out[["Viande -plat", "Naissance", "Élevage", "Abattage"]]


def write_master_workbook(
    allergen_df: pd.DataFrame,
    origins_df: Optional[pd.DataFrame],
    out_path: str,
    title: str = "Référentiel allergènes",
) -> str:
    """Écrit un classeur maître avec 2 feuilles :

    - **Référence** : plats en colonne B, X en C:R (format compatible template).
    - **Origines** (optionnel) : viandes en colonne B, Naissance/Élevage/Abattage en C/D/E.
    """
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
    if allergen_df is not None and not allergen_df.empty:
        df2 = allergen_df.copy()
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

    # --- Feuille Origines ---
    if origins_df is not None:
        wso = wb.create_sheet("Origines")
        wso["A1"] = "Référentiel origines viandes"
        wso["A3"] = "Origines"
        wso["B3"] = "Viande -plat"
        wso["C3"] = "Naissance"
        wso["D3"] = "Élevage"
        wso["E3"] = "Abattage"

        if origins_df is not None and not origins_df.empty:
            odf = origins_df.copy()
            if "Viande -plat" not in odf.columns:
                raise ValueError("Le DataFrame d'origines doit contenir une colonne 'Viande -plat'.")
            for c in ["Naissance", "Élevage", "Abattage"]:
                if c not in odf.columns:
                    odf[c] = ""

            odf["__k"] = odf["Viande -plat"].apply(normalize_key)
            odf = odf[odf["__k"].astype(bool)].copy()
            odf = odf.sort_values("__k").drop_duplicates("__k", keep="first")

            r = 4
            for _, row in odf.iterrows():
                dish = normalize_space(str(row["Viande -plat"] or "")).strip()
                if not dish:
                    continue
                wso[f"B{r}"] = dish
                wso[f"C{r}"] = normalize_space(str(row.get("Naissance", "") or "")).strip() or ""
                wso[f"D{r}"] = normalize_space(str(row.get("Élevage", "") or "")).strip() or ""
                wso[f"E{r}"] = normalize_space(str(row.get("Abattage", "") or "")).strip() or ""
                r += 1

    wb.save(out_path)
    return out_path


def merge_reference(master_path: Optional[str], new_df: pd.DataFrame, out_path: str) -> str:
    """Fusionne new_df dans master (OR) et réécrit le classeur maître (Référence + Origines)."""

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

    # Ré-écrit au format attendu (comme ton fichier 'Allergène semaine.xlsx')
    # (origines gérées à part ci-dessous)

    # --- Origines ---
    new_orig = extract_meat_origins_from_filled_allergen_workbook(master_path) if False else None

    # On charge l'existant (si possible) + on merge ensuite avec les origines extraites du classeur rempli.
    master_orig = pd.DataFrame(columns=["Viande -plat", "Naissance", "Élevage", "Abattage"])
    if master_path and Path(master_path).exists():
        try:
            wb = load_workbook(master_path, data_only=True)
            if "Origines" in wb.sheetnames:
                wso = wb["Origines"]
                rows = []
                for r in range(4, 5000):
                    dish = normalize_space(str(wso[f"B{r}"].value or "")).strip()
                    if not dish:
                        if r > 80:
                            break
                        continue
                    rows.append(
                        {
                            "Viande -plat": dish,
                            "Naissance": normalize_space(str(wso[f"C{r}"].value or "")).strip(),
                            "Élevage": normalize_space(str(wso[f"D{r}"].value or "")).strip(),
                            "Abattage": normalize_space(str(wso[f"E{r}"].value or "")).strip(),
                        }
                    )
                if rows:
                    master_orig = pd.DataFrame(rows)
        except Exception:
            pass

    filled_orig = None
    # new_df vient d'un classeur rempli : on essaye d'en extraire les origines au même moment
    try:
        # Heuristique : si out_path est un merge depuis un classeur rempli, l'appelant peut fournir ce df.
        filled_orig = None
    except Exception:
        filled_orig = None

    # NB : merge_reference ne sait pas quel classeur rempli a servi ; la fusion des origines est faite dans
    # learn_from_filled_allergen_workbook (point d'entrée). Ici on conserve l'existant.

    return write_master_workbook(merged[["Plat"] + ALLERGEN_COLUMNS], master_orig, out_path)


def learn_from_filled_allergen_workbook(
    filled_path: str,
    master_path: Optional[str],
    out_master_path: str,
) -> str:
    """Point d'entrée: lit le classeur rempli et met à jour le référentiel maître."""
    new_df = extract_reference_from_filled_allergen_workbook(filled_path)
    new_orig = extract_meat_origins_from_filled_allergen_workbook(filled_path)

    # 1) merge allergènes
    tmp_out = out_master_path
    merged_path = merge_reference(master_path, new_df, tmp_out)

    # 2) merge origines (lecture du merged_path puis ré-écriture)
    master_orig = pd.DataFrame(columns=["Viande -plat", "Naissance", "Élevage", "Abattage"])
    try:
        wb = load_workbook(merged_path, data_only=True)
        if "Origines" in wb.sheetnames:
            wso = wb["Origines"]
            rows = []
            for r in range(4, 5000):
                dish = normalize_space(str(wso[f"B{r}"].value or "")).strip()
                if not dish:
                    if r > 80:
                        break
                    continue
                rows.append(
                    {
                        "Viande -plat": dish,
                        "Naissance": normalize_space(str(wso[f"C{r}"].value or "")).strip(),
                        "Élevage": normalize_space(str(wso[f"D{r}"].value or "")).strip(),
                        "Abattage": normalize_space(str(wso[f"E{r}"].value or "")).strip(),
                    }
                )
            if rows:
                master_orig = pd.DataFrame(rows)
    except Exception:
        pass

    if new_orig is not None and not new_orig.empty:
        # merge: pour chaque champ, si new non vide -> overwrite
        if master_orig is None or master_orig.empty:
            merged_orig = new_orig.copy()
        else:
            master_orig = master_orig.copy()
            new_orig = new_orig.copy()
            master_orig["__k"] = master_orig["Viande -plat"].apply(normalize_key)
            new_orig["__k"] = new_orig["Viande -plat"].apply(normalize_key)
            idx = {k: i for i, k in enumerate(master_orig["__k"].tolist()) if k}
            for _, row in new_orig.iterrows():
                k = row.get("__k", "")
                if not k:
                    continue
                if k not in idx:
                    idx[k] = len(master_orig)
                    master_orig.loc[len(master_orig)] = {
                        "Viande -plat": row.get("Viande -plat", ""),
                        "Naissance": row.get("Naissance", ""),
                        "Élevage": row.get("Élevage", ""),
                        "Abattage": row.get("Abattage", ""),
                        "__k": k,
                    }
                else:
                    i = idx[k]
                    for c in ["Naissance", "Élevage", "Abattage"]:
                        v = normalize_space(str(row.get(c, "") or "")).strip()
                        if v:
                            master_orig.at[i, c] = v
            merged_orig = master_orig.drop(columns=["__k"], errors="ignore")

        # Ré-écrit le classeur maître avec les 2 feuilles
        # On recharge aussi la feuille Référence sous forme DataFrame pour ré-écriture propre
        merged_all = extract_reference_from_filled_allergen_workbook(merged_path)
        return write_master_workbook(merged_all, merged_orig, out_master_path)

    return merged_path
