"""Génération du Bon de Commande (BC).

✅ Fichier À MODIFIER si tu veux changer la logique/mise en forme du bon de commande.
"""
from __future__ import annotations

from typing import Dict, List
import re
import pandas as pd

from .processor import (
    DAY_NAMES,
    MenuItem,
    _to_number,
    normalize_produit_for_grouping,
    normalize_regime_label,
)

# ==============================
# NOUVELLE FONCTION
# Déduit l’unité depuis le coefficient
# ==============================

def unite_depuis_coefficient(c):
    c = str(c).lower()

    if "kg" in c:
        return "kg"

    if "g" in c:
        return "g"

    if "ml" in c:
        return "ml"

    if "l" in c:
        return "L"

    return "unité"


def build_bon_commande(planning: Dict[str, pd.DataFrame], menu_items: List[MenuItem]) -> pd.DataFrame:

    def norm_reg(s: str) -> str:
        s = (s or "").lower()
        s = (
            s.replace("é", "e")
            .replace("è", "e")
            .replace("ê", "e")
            .replace("î", "i")
            .replace("ï", "i")
            .replace("ô", "o")
            .replace("à", "a")
            .replace("ç", "c")
        )
        s = re.sub(r"[^a-z0-9 ]+", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    records = []

    for repas_key, repas_label in [("dejeuner", "Déjeuner"), ("diner", "Dîner")]:

        df = planning.get(repas_key)

        if df is None or df.empty:
            continue

        df2 = df.copy()

        df2["Regime"] = df2["Regime"].apply(normalize_regime_label)

        agg = df2.groupby("Regime")[DAY_NAMES].sum(numeric_only=True)

        for jour in DAY_NAMES:
            for regime, nb in agg[jour].items():

                records.append(
                    {
                        "Repas": repas_label,
                        "Jour": jour,
                        "Regime_planning": regime,
                        "reg_key_planning": norm_reg(regime),
                        "Nb_personnes": int(_to_number(nb)),
                    }
                )

    counts = pd.DataFrame(records)

    planning_keys = counts[["Regime_planning", "reg_key_planning"]].drop_duplicates().to_dict("records")

    def best_match_planning_key(menu_key: str):

        if not menu_key:
            return None

        mtoks = set(menu_key.split())

        best_key = None
        best_score = -1

        for rec in planning_keys:

            ptoks = set((rec["reg_key_planning"] or "").split())

            score = len(mtoks & ptoks)

            if score > best_score:
                best_score = score
                best_key = rec["reg_key_planning"]

        if best_score <= 0:
            return None

        return best_key

    menu_df = pd.DataFrame(
        [
            {
                "Date": it.date,
                "Jour": DAY_NAMES[it.date.weekday()],
                "Repas": it.repas,
                "Categorie": it.categorie,
                "Regime_menu": it.regime,
                "reg_key_menu": norm_reg(it.regime),
                "Produit": it.produit,
            }
            for it in menu_items
        ]
    )

    menu_df["reg_key_planning"] = menu_df["reg_key_menu"].apply(best_match_planning_key)

    merged = menu_df.merge(
        counts[["Repas", "Jour", "reg_key_planning", "Nb_personnes"]],
        on=["Repas", "Jour", "reg_key_planning"],
        how="left",
    )

    merged["Nb_personnes"] = merged["Nb_personnes"].fillna(0).astype(int)

    merged["Coefficient"] = "1"
    merged["Fournisseur"] = ""

    merged["Unité"] = merged["Coefficient"].apply(unite_depuis_coefficient)

    base = merged[
        ["Date", "Jour", "Repas", "Categorie", "Produit", "Nb_personnes", "Coefficient", "Unité", "Fournisseur"]
    ].rename(columns={"Categorie": "Typologie", "Nb_personnes": "Effectif"})

    base["Produit"] = base["Produit"].astype(str)

    base["Produit_base"] = base["Produit"].apply(normalize_produit_for_grouping)

    base["Quantité"] = (base["Effectif"] * 1.0).round().astype(int)

    grouped = (
        base.groupby(
            ["Repas", "Typologie", "Produit_base", "Coefficient", "Fournisseur"],
            as_index=False,
        )
        .agg(
            {
                "Jour": lambda s: ", ".join(sorted(set(s), key=lambda x: DAY_NAMES.index(x))),
                "Effectif": "sum",
                "Quantité": "sum",
            }
        )
        .rename(columns={"Jour": "Jour(s)", "Produit_base": "Produit"})
    )

    grouped["Unité"] = grouped["Coefficient"].apply(unite_depuis_coefficient)

    grouped["Prix cible unitaire"] = ""
    grouped["Prix cible total"] = ""
    grouped["Poids unitaire (kg)"] = ""
    grouped["Poids total (kg)"] = ""

    grouped = grouped[
        [
            "Jour(s)",
            "Repas",
            "Typologie",
            "Produit",
            "Effectif",
            "Coefficient",
            "Unité",
            "Fournisseur",
            "Quantité",
            "Prix cible unitaire",
            "Prix cible total",
            "Poids unitaire (kg)",
            "Poids total (kg)",
        ]
    ]

    return grouped.sort_values(["Repas", "Typologie", "Produit"]).reset_index(drop=True)


# ==============================
# RÈGLE AJOUTÉE : VIANDE FRAÎCHE
# ==============================

from datetime import timedelta

def est_viande_fraiche(ligne):
    try:
        coef = float(ligne.get("coefficient", 0))
        return coef == 0.15
    except:
        return False

def appliquer_regle_viande_si_applicable(ligne):

    if est_viande_fraiche(ligne) and "dates_conso" in ligne:

        date_premiere = min(ligne["dates_conso"])

        ligne["date_livraison"] = date_premiere - timedelta(days=1)

        ligne["couverture"] = 8

    return ligne