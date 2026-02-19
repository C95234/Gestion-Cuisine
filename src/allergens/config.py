# =========================
# CONFIGURATION ALLERGENES
# =========================

# --- Constantes régimes (compatibilité interne obligatoire) ---
REG_STANDARD = "standard"
REG_VEGETARIAN = "vegetarian"
REG_VEGAN = "vegan"

# --- Colonnes allergènes (alcool supprimé proprement) ---
ALLERGEN_COLUMNS = [
    "gluten",
    "lait",
    "oeuf",
    "arachide",
    "soja",
    "fruit à coque",
    "moutarde",
    "céleri",
    "sésame",
    "sulfites",
    "poisson",
    "crustacé",
    "mollusque"
]

def normalize(value):
    if not isinstance(value, str):
        return ""
    return value.strip().lower()

def is_known_allergen(name):
    return normalize(name) in [normalize(a) for a in ALLERGEN_COLUMNS]
