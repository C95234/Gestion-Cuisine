"""Génération du Bon de Commande (BC).

Fichier modifié pour associer automatiquement une unité au coefficient.
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

# association coefficient → unité
COEFF_UNITE = {
    "1": "unité",
    "0.15": "kg",
    "0.2": "kg",
    "0.05": "L",
    "0.01": "g",
}


def unite_depuis_coefficient(c):
    return COEFF_UNITE.get(str(c), "unité")


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

    # CORRECTION : on additionne tous les régimes pour obtenir l'effectif total
    counts = (
        counts.groupby(["Repas","Jour"], as_index=False)["Nb_personnes"]
        .sum()
    )

    menu_df = pd.DataFrame(
        [
            {
                "Date": it.date,
                "Jour": DAY_NAMES[it.date.weekday()],
                "Repas": it.repas,
                "Categorie": it.categorie,
                "Regime_menu": normalize_regime_label(it.regime),
                "Produit": it.produit,
            }
            for it in menu_items
        ]
    )

    counts["Regime_planning"] = counts["Regime_planning"].apply(normalize_regime_label)

    merged = menu_df.merge(
        counts[["Repas", "Jour", "Regime_planning", "Nb_personnes"]],
        left_on=["Repas","Jour","Regime_menu"],
        right_on=["Repas","Jour","Regime_planning"],
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