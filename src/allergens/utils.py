from __future__ import annotations

import re
import unicodedata

def normalize_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def strip_asterisks(s: str) -> str:
    return (s or "").replace("*", "").strip()

def normalize_key(s: str) -> str:
    """Clé de comparaison pour retrouver un plat dans le référentiel."""
    s = normalize_space(s).lower()
    # enlever les accents
    s = "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")
    # unifier
    s = s.replace("œ", "oe")
    s = re.sub(r"[’'`]", " ", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = normalize_space(s)
    return s
