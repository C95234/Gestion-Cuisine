
 # =========================
# CONFIGURATION ALLERGENES
# =========================

# =========================
# CONSTANTES REGIMES
# (compatibilité complète projet)
# =========================

REG_STANDARD = "standard"

# Végétarien
REG_VEGETARIEN = "vegetarien"
REG_VEGETARIAN = REG_VEGETARIEN

# Végétalien / Vegan
REG_VEGETALIEN = "vegetalien"
REG_VEGAN = REG_VEGETALIEN

# Hypocalorique / Hypo
REG_HYPO = "hypo"
REG_HYPOCALORIQUE = REG_HYPO

# Sans lactose
REG_SANS_LACTOSE = "sans_lactose"
REG_LACTOSE_FREE = REG_SANS_LACTOSE

# Sécurité générique (si d'autres imports existent)
REG_AUTRE = "autre"

# =========================
# ALLERGENES
# =========================

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
