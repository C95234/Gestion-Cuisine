from __future__ import annotations

# Régimes (libellés visibles dans les templates)
REG_STANDARD = "Standards"
REG_VEGETARIEN = "Végétariens"
REG_HYPO = "Hypocaloriques "

# Pour compatibilité avec ton parseur de menus existant :
# - "Sans lactose" correspond à la colonne "sans lactose"
# - "Spéciaux avec lactose" correspond à la colonne "avec lactose" (anciennement "régimes spéciaux")
REG_SPEC_SANS = "Sans lactose"
REG_SPEC_AVEC = "Spéciaux avec lactose"

REGIMES_ORDER = [
    REG_STANDARD,
    REG_VEGETARIEN,
    REG_HYPO,
    REG_SPEC_SANS,
    REG_SPEC_AVEC,
]

SERVICE_DEJ = "Déjeuner"
SERVICE_DIN = "Diner"

# Colonnes allergènes (doivent correspondre aux en-têtes du template)
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
