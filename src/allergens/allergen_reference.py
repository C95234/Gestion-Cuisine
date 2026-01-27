from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Set, Tuple, Optional, Iterable
import difflib

import pandas as pd

from .config import ALLERGEN_COLUMNS
from .utils import normalize_key, normalize_space

@dataclass
class AllergenRef:
    # key normalisée du plat -> set d'allergènes (libellés exacts ALLERGEN_COLUMNS)
    map_key_to_allergens: Dict[str, Set[str]]

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

    map_key_to_allergens: Dict[str, Set[str]] = {}
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

    return AllergenRef(map_key_to_allergens=map_key_to_allergens)
