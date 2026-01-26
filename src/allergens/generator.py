
from __future__ import annotations

from typing import List, Tuple
import os

from .menu_reader import read_menus
from .allergen_reference import load_allergen_reference
from .template_filler import fill_allergen_workbook

def generate_allergen_workbook(
    menu_excel_path: str,
    allergen_ref_path: str,
    out_xlsx_path: str,
    template_dir: str,
) -> Tuple[str, List[str]]:
    """Génère le classeur "Allergènes" au format EXACT du template, 1 feuille par service.
    Retourne (out_xlsx_path, plats_non_trouves).
    """
    menus_by_day = read_menus(menu_excel_path)
    ref = load_allergen_reference(allergen_ref_path)

    # ref.map_key_to_allergens: normalized dish key -> set(allergen labels)
    out_path, missing = fill_allergen_workbook(
        menus_by_day=menus_by_day,
        allergen_ref_key_to_allergens=ref.map_key_to_allergens,
        meat_origin_ref=ref.map_key_to_meat_origins,
        template_dir=template_dir,
        out_path=out_xlsx_path,
    )
    return out_path, missing
