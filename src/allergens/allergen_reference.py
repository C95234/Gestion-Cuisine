from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Set, Tuple, Optional
import difflib

import pandas as pd
from openpyxl import load_workbook

from .config import ALLERGEN_COLUMNS
from .utils import normalize_key, normalize_space

@dataclass
class AllergenRef:
    # key normalisée du plat -> set d'allergènes (libellés exacts ALLERGEN_COLUMNS)
    map_key_to_allergens: Dict[str, Set[str]]
    # key normalisée du plat -> (naissance, élevage, abattage)
    map_key_to_meat_origins: Dict[str, Tuple[str, str, str]]

    def lookup(self, dish_name: str) -> Tuple[Set[str], Optional[str]]:
        """Retourne (allergènes, clé_trouvée)."""
        k = normalize_key(dish_name)
        if not k:
            return set(), None
        if k in self.map_key_to_allergens:
            return set(self.map_key_to_allergens[k]), k

        # fuzzy match (tolérant aux petites différences)
        keys = list(self.map_key_to_allergens.keys())
        cand = difflib.get_close_matches(k, keys, n=1, cutoff=0.88)
        if cand:
            ck = cand[0]
            return set(self.map_key_to_allergens[ck]), ck
        return set(), None


def load_allergen_reference(path: str) -> AllergenRef:
    """Lit un référentiel allergènes **Excel** (.xlsx/.xlsm/.xls).

    Attendu :
      - une colonne Plat (ou plat/denrée)
      - 16 colonnes allergènes (noms identiques ou proches de ALLERGEN_COLUMNS),
        valeurs : X / 1 / vrai / oui -> allergène présent.
    """
    # NOTE: on désactive volontairement la lecture .ods pour éviter la dépendance
    # optionnelle "odfpy" (souvent absente sur Windows). Convertis le .ods en .xlsx.
    if path.lower().endswith('.ods'):
        raise ValueError(
            "Le référentiel allergènes doit être un fichier Excel (.xlsx). "
            "Merci de convertir ton .ods en .xlsx (Fichier > Enregistrer sous)."
        )
    # 1) Cas "classeur maître" au format templates (sheet Référence)
    #    -> on lit via openpyxl pour être robuste (pandas lit mal le header décalé).
    try:
        wb = load_workbook(path, data_only=True)
        if "Référence" in wb.sheetnames:
            ws = wb["Référence"]
            # Heuristique : C3 doit correspondre au 1er allergène du template
            if normalize_key(str(ws["C3"].value or "")) == normalize_key(ALLERGEN_COLUMNS[0]):
                map_key_to_allergens: Dict[str, Set[str]] = {}
                for r in range(4, 5000):
                    dish = normalize_space(str(ws[f"B{r}"].value or "")).strip()
                    if not dish:
                        # stop si on rencontre une longue zone vide
                        if r > 80:
                            break
                        continue
                    k = normalize_key(dish)
                    if not k:
                        continue
                    als: Set[str] = set()
                    for i, a in enumerate(ALLERGEN_COLUMNS):
                        col_letter = chr(ord("C") + i)  # C..R
                        v = normalize_space(str(ws[f"{col_letter}{r}"].value or "")).strip().lower()
                        if v in ("x", "1", "oui", "vrai", "true", "y", "yes"):
                            als.add(a)
                    map_key_to_allergens[k] = map_key_to_allergens.get(k, set()) | als

                # Origines viandes (sheet optionnelle)
                map_key_to_meat_origins: Dict[str, Tuple[str, str, str]] = {}
                if "Origines" in wb.sheetnames:
                    wso = wb["Origines"]
                    # Attendu : en-tête en ligne 3, données à partir de 4 (col B = plat)
                    for r in range(4, 5000):
                        dish = normalize_space(str(wso[f"B{r}"].value or "")).strip()
                        if not dish:
                            if r > 80:
                                break
                            continue
                        k = normalize_key(dish)
                        if not k:
                            continue
                        naissance = normalize_space(str(wso[f"C{r}"].value or "")).strip()
                        elevage = normalize_space(str(wso[f"D{r}"].value or "")).strip()
                        abattage = normalize_space(str(wso[f"E{r}"].value or "")).strip()
                        map_key_to_meat_origins[k] = (naissance, elevage, abattage)

                return AllergenRef(
                    map_key_to_allergens=map_key_to_allergens,
                    map_key_to_meat_origins=map_key_to_meat_origins,
                )
    except Exception:
        # fallback pandas
        pass

    # 2) Cas "table" standard (1 ligne = 1 plat)
    df = pd.read_excel(path)

    # trouver colonne plat
    cols = {c: normalize_key(str(c)) for c in df.columns}
    plat_col = None
    for c, nk in cols.items():
        if nk in ('plat', 'plats', 'denree', 'denree produit', 'produit', 'libelle', 'intitule'):
            plat_col = c
            break
    if plat_col is None:
        # fallback: première colonne
        plat_col = df.columns[0]

    # mapper colonnes allergènes
    norm_to_real = {normalize_key(x): x for x in ALLERGEN_COLUMNS}
    col_map = {}
    for c in df.columns:
        nk = normalize_key(str(c))
        if nk in norm_to_real:
            col_map[c] = norm_to_real[nk]

    # si en-têtes légèrement différents, tenter inclusion
    if len(col_map) < len(ALLERGEN_COLUMNS):
        for c in df.columns:
            nk = normalize_key(str(c))
            for tgt in ALLERGEN_COLUMNS:
                tnk = normalize_key(tgt)
                if tnk and (tnk in nk or nk in tnk) and c not in col_map:
                    col_map[c] = tgt

    # colonnes origines (optionnelles)
    origin_cols = {}
    for c in df.columns:
        nk = normalize_key(str(c))
        if nk in ("naissance", "lieux de naissance", "lieu de naissance"):
            origin_cols["naissance"] = c
        elif nk in ("elevage", "élevage", "lieux d elevage", "lieu d elevage"):
            origin_cols["elevage"] = c
        elif nk in ("abattage", "abatage", "lieux d abattage", "lieu d abattage"):
            origin_cols["abattage"] = c

    map_key_to_allergens: Dict[str, Set[str]] = {}
    map_key_to_meat_origins: Dict[str, Tuple[str, str, str]] = {}
    for _, row in df.iterrows():
        dish = normalize_space(str(row.get(plat_col, '') or ''))
        if not dish or dish.lower() == 'nan':
            continue
        dish_key = normalize_key(dish)
        if not dish_key:
            continue
        allergens: Set[str] = set()
        for c, allergen_name in col_map.items():
            v = row.get(c, '')
            if v is None:
                continue
            s = str(v).strip().lower()
            if s in ('x', '1', 'vrai', 'true', 'oui', 'y'):
                allergens.add(allergen_name)
        if dish_key in map_key_to_allergens:
            map_key_to_allergens[dish_key] |= allergens
        else:
            map_key_to_allergens[dish_key] = allergens

        if origin_cols:
            naissance = normalize_space(str(row.get(origin_cols.get("naissance"), "") or "")).strip()
            elevage = normalize_space(str(row.get(origin_cols.get("elevage"), "") or "")).strip()
            abattage = normalize_space(str(row.get(origin_cols.get("abattage"), "") or "")).strip()
            if any([naissance, elevage, abattage]):
                map_key_to_meat_origins[dish_key] = (naissance, elevage, abattage)

    return AllergenRef(
        map_key_to_allergens=map_key_to_allergens,
        map_key_to_meat_origins=map_key_to_meat_origins,
    )
