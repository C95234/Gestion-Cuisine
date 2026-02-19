# =========================
# CONFIGURATION ALLERGENES
# =========================

# --- Constantes régimes (compatibilité interne obligatoire) ---
REG_STANDARD = "standard"
<<<<<<< HEAD
REG_VEGETARIAN = "vegetarian"
REG_VEGAN = "vegan"

# --- Colonnes allergènes (alcool supprimé proprement) ---
=======

REG_VEGETARIEN = "vegetarien"
REG_VEGETARIAN = REG_VEGETARIEN

REG_VEGAN = "vegan"

# --- Colonnes allergènes ---
>>>>>>> 8b62f1e46a8c9f156a26ece7762a402ab9de7361
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
<<<<<<< HEAD
=======

>>>>>>> 8b62f1e46a8c9f156a26ece7762a402ab9de7361
