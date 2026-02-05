from __future__ import annotations

from typing import Set
from .config import ALLERGEN_COLUMNS
from .utils import normalize_key

def heuristic_allergens(dish_name: str) -> Set[str]:
    """Heuristique simple (fallback) si un plat n'est pas dans le référentiel.

    Objectif : éviter des tableaux vides ; à ajuster selon ta cuisine.
    """
    k = normalize_key(dish_name)
    out: Set[str] = set()

    def has(*words):
        return any(w in k for w in words)

    if has('saumon','poisson','cabillaud','colin','thon','merlu','truite','sardine'):
        out.add("Poisson")
    if has('oeuf','omelette','quiche'):
        out.add("Œuf")
    if has('fromage','emmental','cheddar','mozzarella','yaourt','lait','beurre','creme','crème','mornay'):
        out.add("Lactose")
    if has('moutarde'):
        out.add("Moutarde")
    if has('celeri','céleri'):
        out.add("Céleri")
    if has('sesame','sésame'):
        out.add("Sésame")
    if has('soja'):
        out.add("Soja")
    if has('arachide','cacahuete','cacahuète'):
        out.add("Arachide")
    if has('lupin'):
        out.add("Lupin")
    if has('noix','amande','noisette','pistache','cajou'):
        out.add("Fruit à coques")
    if has('biere','bière','vin','alcool','cognac','rhum','armagnac'):
        out.add("alcool dans sauce")
    if has('sulfite','sulfites'):
        out.add("Sulfites")
    if has('ble','blé','farine','pates','pâtes','semoule','pain','chapelure','pizza','biscuit','gateau','gâteau'):
        out.add("Céréales/gluten")

    return out
