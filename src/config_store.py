"""Petite couche de persistance (JSON) pour des listes éditables dans l'app.

Objectif : permettre de modifier facilement des listes (ex: fournisseurs)
sans toucher aux fonctions cœur (parsing/planning).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import List


DEFAULT_SUPPLIERS = ["Sysco", "Domafrais", "Cercle Vert"]


def _data_dir() -> Path:
    # dossier persistant (dans le repo) ; en cloud, Streamlit garde en général l'espace
    # de travail, sinon on retombe sur les valeurs par défaut.
    d = Path(__file__).resolve().parent / "data"
    d.mkdir(parents=True, exist_ok=True)
    return d


def suppliers_path() -> Path:
    return _data_dir() / "suppliers.json"


def load_suppliers() -> List[str]:
    p = suppliers_path()
    if not p.exists():
        return DEFAULT_SUPPLIERS.copy()
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, list):
            out = [str(x).strip() for x in data if str(x).strip()]
            return out or DEFAULT_SUPPLIERS.copy()
    except Exception:
        pass
    return DEFAULT_SUPPLIERS.copy()


def save_suppliers(values: List[str]) -> None:
    cleaned = [str(v).strip() for v in (values or []) if str(v).strip()]
    if not cleaned:
        cleaned = DEFAULT_SUPPLIERS.copy()
    suppliers_path().write_text(json.dumps(cleaned, ensure_ascii=False, indent=2), encoding="utf-8")
