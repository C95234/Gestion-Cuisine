# =========================
# CONFIGURATION ALLERGENES
# =========================

# --- Constantes régimes ---
REG_STANDARD = "standard"
REG_VEGETARIEN = "vegetarien"
REG_VEGETARIAN = REG_VEGETARIEN
REG_VEGAN = "vegan"

# --- Colonnes allergènes ---
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
