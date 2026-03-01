from __future__ import annotations

# Régimes (libellés visibles dans les tableaux Allergènes)
REG_STANDARD = "Standards"
REG_VEGETARIEN = "Végétariens"
REG_HYPO = "Hypocaloriques"
REG_SPEC_SANS = "Sans lactose"
REG_SPEC_AVEC = "Spéciaux av lactose"

# Ordre d'affichage (format unique)
REGIMES_ORDER = [
    REG_STANDARD,
    REG_VEGETARIEN,
    REG_HYPO,
    REG_SPEC_SANS,
    REG_SPEC_AVEC,
]

SERVICE_DEJ = "Déjeuner"
SERVICE_DIN = "Dîner"

# Colonnes allergènes (doivent correspondre aux en-têtes du nouveau template)
ALLERGEN_COLUMNS = [
    "Céréales/gluten",
    "Crustacés",
    "Mollusques",
    "Poisson",
    "Œuf",
    "Arachide",
    "Soja",
    "Lactose",
    "Fruit à coques",
    "Céleri",
    "Moutarde",
    "Sésame",
    "Lupin",
    "G6PD*",
    "Sulfites",
]
