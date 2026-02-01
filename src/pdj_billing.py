"""Facturation PDJ (économat).

But
---
Importer plusieurs bons de commande PDJ (Excel ou PDF) et remplir un classeur
de facturation économat (PU unique).

Contraintes
-----------
- Ne pas dépendre de la position des lignes (les produits peuvent changer).
- Ne pas écraser les prix : on utilise le prix du classeur de facturation.
- Si un produit est nouveau, on l'ajoute en bas avec une formule de total.

Remarque PDF
------------
Les PDF PDJ fournis sont souvent des scans (texte non extractible). Sans OCR,
on ne peut pas lire automatiquement les quantités manuscrites. Le module
fournit donc un "format items" générique ; la saisie des quantités pour les
PDF est faite côté UI (st.data_editor) dans app.py.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, List, Tuple

import datetime as dt
import re

import openpyxl


# Liste "de base" des produits PDJ (pour aide saisie, surtout côté UI).
DEFAULT_PDJ_PRODUCTS: List[str] = [
    "Lait demi - écrémé",
    "Lait entier",
    "Céréales",
    "Biscotte",
    "Sucre en sachet",
    "Sucre en morceau",
    "Beurre, plaquettes de 250g",
    "Chocolat en poudre",
    "Brioche",
    "Blédine arome chocolat",
    "Blédine arome vanille",
    "Confiture en carton",
    "Thé en boite",
    "Café en carton",
    "Jus d'orange",
    "Jus de pomme",
    "Jus de raisin",
    "Mayonnaise",
    "Ketchup",
    "Sel",
    "Poivre",
    "Fromage blanc pot de 5kg",
    "Yaourt Nature",
]


def _norm(s: str) -> str:
    """Normalise un libellé produit/site pour matching robuste."""
    if s is None:
        return ""
    s = str(s)
    s = s.strip().lower()
    # enlève accents de façon simple (sans dépendances)
    s = (
        s.replace("é", "e")
        .replace("è", "e")
        .replace("ê", "e")
        .replace("à", "a")
        .replace("â", "a")
        .replace("î", "i")
        .replace("ï", "i")
        .replace("ô", "o")
        .replace("ù", "u")
        .replace("û", "u")
        .replace("ç", "c")
    )
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-z0-9 ',-]", "", s)
    return s.strip()


@dataclass
class PDJItem:
    product: str
    qty: float
    unit: str | None = None


@dataclass
class PDJOrder:
    """Une commande PDJ pour un site donné."""

    site: str
    date: dt.date
    items: List[PDJItem]


def read_sites_and_products(template_path: str) -> Tuple[List[str], List[str]]:
    """Retourne (sites, produits) à partir du classeur économat."""
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    header_row = _find_header_row(ws)
    if header_row is None:
        raise ValueError("Impossible de trouver la ligne d'en-tête (Produit/Unité/Prix...).")

    sites = []
    # D..L = 4..12
    for col in range(4, 13):
        v = ws.cell(header_row, col).value
        if v is None:
            continue
        s = str(v).strip()
        if s:
            sites.append(s)

    # produits = colonne A, sous l'en-tête
    products = []
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is None:
            continue
        t = str(v).strip()
        if t:
            products.append(t)
    return sites, products


def _find_header_row(ws) -> int | None:
    """Trouve la ligne contenant 'Produit' 'Unité' 'Prix unitaire'."""
    for r in range(1, min(ws.max_row, 60) + 1):
        a = _norm(ws.cell(r, 1).value or "")
        b = _norm(ws.cell(r, 2).value or "")
        c = _norm(ws.cell(r, 3).value or "")
        if a == "produit" and b.startswith("unite") and c.startswith("prix"):
            return r
    return None


def apply_pdj_orders_to_economat_workbook(
    template_path: str,
    orders: Iterable[PDJOrder],
    out_path: str,
) -> None:
    """Applique des commandes PDJ dans un classeur économat.

    - Ajoute les quantités (somme) dans la colonne du site.
    - Crée la ligne du produit si absente.
    - Met à jour la cellule "Mois (YYYY-MM)" (B2) avec le mois des commandes.
      Si plusieurs mois, prend le mois du premier ordre.
    """

    orders = list(orders)
    if not orders:
        raise ValueError("Aucune commande PDJ fournie.")

    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    header_row = _find_header_row(ws)
    if header_row is None:
        raise ValueError("Format du classeur économat inconnu (ligne d'en-tête introuvable).")

    # mapping sites -> colonne
    site_to_col: Dict[str, int] = {}
    for col in range(4, 13):
        name = ws.cell(header_row, col).value
        if name is None:
            continue
        site_to_col[_norm(str(name))] = col

    # mapping produits -> ligne
    prod_to_row: Dict[str, int] = {}
    last_row = header_row
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is None:
            continue
        prod = str(v).strip()
        if prod:
            prod_to_row[_norm(prod)] = r
            last_row = max(last_row, r)

    # met le mois
    month = orders[0].date.strftime("%Y-%m")
    ws["B2"].value = month

    def _ensure_product_row(product: str, unit: str | None = None) -> int:
        nonlocal last_row
        key = _norm(product)
        if key in prod_to_row:
            return prod_to_row[key]

        # Nouvelle ligne
        last_row += 1
        r = last_row
        ws.cell(r, 1).value = product
        if unit:
            ws.cell(r, 2).value = unit
        # Prix unitaire (col 3) laissé vide (ou 0) => utilisateur gère
        # Formule total (col 13)
        ws.cell(r, 13).value = f"=IF(C{r}=\"\",0,C{r}*SUM(D{r}:L{r}))"
        prod_to_row[key] = r
        return r

    # applique quantités
    for o in orders:
        site_key = _norm(o.site)
        if site_key not in site_to_col:
            raise ValueError(
                f"Site '{o.site}' introuvable dans le classeur économat. "
                f"Sites attendus: {', '.join(sorted(set(site_to_col.keys())))}"
            )
        col = site_to_col[site_key]

        for it in o.items:
            if it.qty is None:
                continue
            try:
                qty = float(it.qty)
            except Exception:
                continue
            if qty == 0:
                continue
            r = _ensure_product_row(it.product, unit=it.unit)
            cell = ws.cell(r, col)
            prev = cell.value
            try:
                prev_f = float(prev) if prev not in (None, "") else 0.0
            except Exception:
                prev_f = 0.0
            cell.value = prev_f + qty

    wb.save(out_path)
