"""Génération du Bon de Commande (BC)."""

from __future__ import annotations

from typing import Dict, List, Optional
import re
import unicodedata
import pandas as pd

from .processor import (
    DAY_NAMES,
    MenuItem,
    _to_number,
    normalize_produit_for_grouping,
    normalize_regime_label,
)

COEFF_UNITE = {
    "1": "unité",
    "0.15": "kg",
    "0.2": "kg",
    "0.05": "L",
    "0.01": "g",
}


def unite_depuis_coefficient(c):
    return COEFF_UNITE.get(str(c), "unité")


def _norm_text(s: str) -> str:
    s = (s or "").strip().lower()
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _canon_regime_key(regime: str) -> str:
    s = _norm_text(normalize_regime_label(regime))
    if not s:
        return ""
    if "vegetalien" in s or "vegan" in s:
        return "vegetalien"
    if "vegetar" in s:
        return "vegetarien"
    if "hypo" in s or "hypocal" in s:
        return "hypocalorique"
    if "diab" in s:
        return "diabetique"
    if re.search(r"(ss|sans)", s):
        return "sans"
    if "special" in s:
        return "special"
    if any(tok in s.split() for tok in ["standard", "normal", "base", "classique", "ordinaire"]):
        return "standard"
    return s


def _best_match_planning_key(menu_key: str, planning_keys: list[str]) -> Optional[str]:
    if not menu_key:
        return None
    if menu_key in planning_keys:
        return menu_key
    mtoks = set(menu_key.split())
    best_key = None
    best_score = -1
    for pkey in planning_keys:
        ptoks = set((pkey or "").split())
        score = len(mtoks & ptoks)
        if score > best_score:
            best_score = score
            best_key = pkey
    return best_key if best_score > 0 else None



def build_bon_commande(planning: Dict[str, pd.DataFrame], menu_items: List[MenuItem]) -> pd.DataFrame:
    records = []

    # Effectifs par repas + jour + régime (surtout pas les totaux de journée sur chaque ligne menu)
    for repas_key, repas_label in [("dejeuner", "Déjeuner"), ("diner", "Dîner")]:
        df = planning.get(repas_key)
        if df is None or df.empty:
            continue

        df2 = df.copy()
        if "Regime" not in df2.columns:
            continue
        df2["Regime"] = df2["Regime"].apply(normalize_regime_label)
        agg = df2.groupby("Regime")[DAY_NAMES].sum(numeric_only=True)

        for jour in DAY_NAMES:
            for regime, nb in agg[jour].items():
                records.append(
                    {
                        "Repas": repas_label,
                        "Jour": jour,
                        "Regime_planning": regime,
                        "canon_regime": _canon_regime_key(regime),
                        "Nb_personnes": int(_to_number(nb)),
                    }
                )

    counts = pd.DataFrame(records)

    menu_df = pd.DataFrame(
        [
            {
                "Date": it.date,
                "Jour": DAY_NAMES[it.date.weekday()],
                "Repas": it.repas,
                "Categorie": it.categorie,
                "Regime_menu": normalize_regime_label(it.regime),
                "canon_regime": _canon_regime_key(it.regime),
                "Produit": it.produit,
            }
            for it in menu_items
        ]
    )

    if counts.empty or menu_df.empty:
        merged = menu_df.copy()
        merged["Nb_personnes"] = 0
    else:
        planning_keys = sorted(k for k in counts["canon_regime"].dropna().astype(str).unique().tolist() if k)
        menu_df["canon_regime"] = menu_df["canon_regime"].apply(lambda k: _best_match_planning_key(k, planning_keys) or k)

        merged = menu_df.merge(
            counts[["Repas", "Jour", "canon_regime", "Nb_personnes"]],
            on=["Repas", "Jour", "canon_regime"],
            how="left",
        )

        # Fallback très limité : uniquement si le produit n'existe qu'une seule fois ce jour-là pour ce repas.
        # Cela évite de coller l'effectif total à chaque régime et de doubler/tripler les quantités.
        total_counts = counts.groupby(["Repas", "Jour"], as_index=False)["Nb_personnes"].sum().rename(columns={"Nb_personnes": "Nb_total_jour"})
        merged = merged.merge(total_counts, on=["Repas", "Jour"], how="left")
        merged["Produit_base"] = merged["Produit"].astype(str).apply(normalize_produit_for_grouping)
        key_cols = ["Date", "Jour", "Repas", "Categorie", "Produit_base"]
        merged["_nb_lignes_produit_jour"] = merged.groupby(key_cols)["Produit"].transform("size")
        merged["Nb_personnes"] = merged["Nb_personnes"].where(
            merged["Nb_personnes"].notna(),
            merged["Nb_total_jour"].where(merged["_nb_lignes_produit_jour"] == 1, 0),
        )

    merged["Nb_personnes"] = merged["Nb_personnes"].fillna(0).astype(int)
    merged["Coefficient"] = "1"
    merged["Fournisseur"] = ""
    merged["Unité"] = merged["Coefficient"].apply(unite_depuis_coefficient)

    base = merged[["Date", "Jour", "Repas", "Categorie", "Produit", "Nb_personnes", "Coefficient", "Unité", "Fournisseur"]].rename(
        columns={"Categorie": "Typologie", "Nb_personnes": "Effectif"}
    )
    base["Produit"] = base["Produit"].astype(str)
    base["Produit_base"] = base["Produit"].apply(normalize_produit_for_grouping)
    base["Quantité"] = (base["Effectif"] * 1.0).round().astype(int)

    grouped = (
        base.groupby(["Repas", "Typologie", "Produit_base", "Coefficient", "Fournisseur"], as_index=False)
        .agg({
            "Jour": lambda s: ", ".join(sorted(set(s), key=lambda x: DAY_NAMES.index(x))),
            "Effectif": "sum",
            "Quantité": "sum",
        })
        .rename(columns={"Jour": "Jour(s)", "Produit_base": "Produit"})
    )

    grouped["Unité"] = grouped["Coefficient"].apply(unite_depuis_coefficient)
    grouped["Prix cible unitaire"] = ""
    grouped["Prix cible total"] = ""
    grouped["Poids unitaire (kg)"] = ""
    grouped["Poids total (kg)"] = ""

    grouped = grouped[[
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
    ]]

    return grouped.sort_values(["Repas", "Typologie", "Produit"]).reset_index(drop=True)
