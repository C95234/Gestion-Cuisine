from __future__ import annotations

from typing import List, Tuple
import os

import pandas as pd

from .menu_reader import read_menus
from .allergen_reference import load_allergen_reference
from .template_filler import fill_allergen_workbook
from .learner import extract_reference_from_filled_allergen_workbook
from .utils import normalize_key


def _merge_learned_into_reference(allergen_ref_path: str, learned_df: pd.DataFrame) -> None:
    """
    "Apprentissage" simple :
    - On extrait Plat + colonnes allergènes depuis un tableau Allergènes corrigé (avec des X)
    - On fusionne dans le référentiel (ajout d'information : OR sur les X)
    Objectif : pré-remplissage automatique les semaines suivantes.

    NB : on n'efface pas d'allergènes automatiquement (pour éviter de supprimer par erreur).
    """
    if learned_df is None or learned_df.empty:
        return

    # Charger référentiel existant si possible (format tableau classique)
    ref_df = None
    if os.path.exists(allergen_ref_path):
        try:
            ref_df = pd.read_excel(allergen_ref_path)
        except Exception:
            ref_df = None

    # Normaliser
    learned = learned_df.copy()
    if "Plat" not in learned.columns:
        return
    learned["__k"] = learned["Plat"].astype(str).apply(normalize_key)
    learned = learned[learned["__k"].astype(bool)].copy()

    if ref_df is None or ref_df.empty or "Plat" not in ref_df.columns:
        # créer un nouveau référentiel
        out = learned.drop(columns=["__k"]).copy()
        out.to_excel(allergen_ref_path, index=False)
        return

    ref = ref_df.copy()
    ref["__k"] = ref["Plat"].astype(str).apply(normalize_key)
    ref = ref[ref["__k"].astype(bool)].copy()

    # Colonnes allergènes = intersection (on conserve les colonnes déjà présentes + celles apprises)
    allergen_cols = [c for c in learned.columns if c not in ("Plat", "__k")]
    for c in allergen_cols:
        if c not in ref.columns:
            ref[c] = ""

    ref_map = {k: i for i, k in enumerate(ref["__k"].tolist())}

    for _, row in learned.iterrows():
        k = row["__k"]
        if k in ref_map:
            i = ref_map[k]
            # OR sur les X
            for c in allergen_cols:
                lv = str(row.get(c, "") or "").strip().upper()
                if lv == "X":
                    ref.at[i, c] = "X"
        else:
            # nouvelle ligne
            new_row = {"Plat": row["Plat"]}
            for c in allergen_cols:
                new_row[c] = "X" if str(row.get(c, "") or "").strip().upper() == "X" else ""
            new_row["__k"] = k
            ref.loc[len(ref)] = new_row
            ref_map[k] = len(ref) - 1

    ref = ref.drop(columns=["__k"], errors="ignore")
    ref.to_excel(allergen_ref_path, index=False)


def generate_allergen_workbook(
    menu_excel_path: str,
    allergen_ref_path: str,
    out_xlsx_path: str,
    template_dir: str,
) -> Tuple[str, List[str]]:
    """
    Génère le classeur "Allergènes" au format UNIQUE (celui du nouveau template),
    1 feuille par service.

    Apprentissage :
    - Si out_xlsx_path existe déjà (tableau corrigé de la semaine précédente),
      le logiciel en extrait les X et les fusionne dans allergen_ref_path
      afin de pré-remplir automatiquement les semaines suivantes.
    """
    # 1) Apprentissage depuis le dernier tableau corrigé (si présent)
    if out_xlsx_path and os.path.exists(out_xlsx_path):
        try:
            learned_df = extract_reference_from_filled_allergen_workbook(out_xlsx_path)
            _merge_learned_into_reference(allergen_ref_path, learned_df)
        except Exception:
            # On n'empêche jamais la génération si l'apprentissage échoue
            pass

    # 2) Lecture du menu (format nouveau)
    menus_by_day = read_menus(menu_excel_path)

    # 3) Chargement référentiel + génération
    ref = load_allergen_reference(allergen_ref_path)

    out_path, missing = fill_allergen_workbook(
        menus_by_day=menus_by_day,
        allergen_ref_key_to_allergens=ref.map_key_to_allergens,
        template_dir=template_dir,
        out_path=out_xlsx_path,
    )
    return out_path, missing
