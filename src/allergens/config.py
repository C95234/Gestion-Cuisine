# =========================
# CONFIGURATION ALLERGENES
# =========================

# --- Constantes régimes (compatibilité totale) ---

REG_STANDARD = "standard"

# Variantes végétarien
REG_VEGETARIEN = "vegetarien"
REG_VEGETARIAN = REG_VEGETARIEN

# Variantes végétalien / vegan
REG_VEGETALIEN = "vegetalien"
REG_VEGAN = REG_VEGETALIEN

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

