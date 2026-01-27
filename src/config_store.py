"""Petite couche de persistance (JSON) pour des listes éditables dans l'app.

Objectif : permettre de modifier facilement des listes (ex: fournisseurs)
sans toucher aux fonctions cœur (parsing/planning).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List


# =====================
# Valeurs par défaut
# =====================

# Coefficients proposés dans les bons (ex: 1 = 1 part ; 0.5 = 1/2 part, etc.)
DEFAULT_COEFFICIENTS = [1.0]

# Unités proposées dans les bons
DEFAULT_UNITS = ["kg", "unité", "L", "barquette"]

# Fournisseurs : fiches (nom + code client + 1/2 coordonnées libres)
DEFAULT_SUPPLIERS: List[Dict[str, str]] = [
    {"name": "Sysco", "code_client": "", "coord1": "", "coord2": ""},
    {"name": "Domafrais", "code_client": "", "coord1": "", "coord2": ""},
    {"name": "Cercle Vert", "code_client": "", "coord1": "", "coord2": ""},
]


def _data_dir() -> Path:
    # dossier persistant (dans le repo) ; en cloud, Streamlit garde en général l'espace
    # de travail, sinon on retombe sur les valeurs par défaut.
    d = Path(__file__).resolve().parent / "data"
    d.mkdir(parents=True, exist_ok=True)
    return d


def suppliers_path() -> Path:
    return _data_dir() / "suppliers.json"


def units_path() -> Path:
    return _data_dir() / "units.json"


def coeffs_path() -> Path:
    return _data_dir() / "coefficients.json"


def _clean_text(x: Any) -> str:
    s = "" if x is None else str(x)
    s = " ".join(s.split()).strip()
    return s


def load_suppliers() -> List[Dict[str, str]]:
    """Charge la liste de fournisseurs (fiches).

    Format attendu: [{"name":..., "code_client":..., "coord1":..., "coord2":...}, ...]
    """
    p = suppliers_path()
    if not p.exists():
        return [dict(x) for x in DEFAULT_SUPPLIERS]
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, list):
            out: List[Dict[str, str]] = []
            for it in data:
                if not isinstance(it, dict):
                    # compat ancien format: liste de strings
                    name = _clean_text(it)
                    if name:
                        out.append({"name": name, "code_client": "", "coord1": "", "coord2": ""})
                    continue
                name = _clean_text(it.get("name"))
                if not name:
                    continue
                out.append(
                    {
                        "name": name,
                        "code_client": _clean_text(it.get("code_client")),
                        "coord1": _clean_text(it.get("coord1")),
                        "coord2": _clean_text(it.get("coord2")),
                    }
                )
            return out or [dict(x) for x in DEFAULT_SUPPLIERS]
    except Exception:
        pass
    return [dict(x) for x in DEFAULT_SUPPLIERS]


def save_suppliers(values: List[Dict[str, str]]) -> None:
    cleaned: List[Dict[str, str]] = []
    for it in (values or []):
        if not isinstance(it, dict):
            continue
        name = _clean_text(it.get("name"))
        if not name:
            continue
        cleaned.append(
            {
                "name": name,
                "code_client": _clean_text(it.get("code_client")),
                "coord1": _clean_text(it.get("coord1")),
                "coord2": _clean_text(it.get("coord2")),
            }
        )
    if not cleaned:
        cleaned = [dict(x) for x in DEFAULT_SUPPLIERS]
    suppliers_path().write_text(json.dumps(cleaned, ensure_ascii=False, indent=2), encoding="utf-8")


def load_units() -> List[str]:
    p = units_path()
    if not p.exists():
        return DEFAULT_UNITS.copy()
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, list):
            out = [_clean_text(x) for x in data]
            out = [x for x in out if x]
            return out or DEFAULT_UNITS.copy()
    except Exception:
        pass
    return DEFAULT_UNITS.copy()


def save_units(values: List[str]) -> None:
    cleaned = [_clean_text(v) for v in (values or [])]
    cleaned = [x for x in cleaned if x]
    if not cleaned:
        cleaned = DEFAULT_UNITS.copy()
    units_path().write_text(json.dumps(cleaned, ensure_ascii=False, indent=2), encoding="utf-8")


def load_coefficients() -> List[float]:
    p = coeffs_path()
    if not p.exists():
        return DEFAULT_COEFFICIENTS.copy()
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, list):
            out: List[float] = []
            for x in data:
                try:
                    v = float(str(x).replace(",", ".").strip())
                    if v > 0:
                        out.append(v)
                except Exception:
                    continue
            return out or DEFAULT_COEFFICIENTS.copy()
    except Exception:
        pass
    return DEFAULT_COEFFICIENTS.copy()


def save_coefficients(values: List[float]) -> None:
    cleaned: List[float] = []
    for x in (values or []):
        try:
            v = float(str(x).replace(",", ".").strip())
            if v > 0:
                cleaned.append(v)
        except Exception:
            continue
    if not cleaned:
        cleaned = DEFAULT_COEFFICIENTS.copy()
    coeffs_path().write_text(json.dumps(cleaned, ensure_ascii=False, indent=2), encoding="utf-8")
