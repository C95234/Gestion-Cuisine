PATCH MINIMAL (zéro conflit)

Ce patch NE remplace PAS tout le projet (et surtout PAS le dossier .git du zip).
Il modifie uniquement app.py pour:
- charger src.pdj_facturation si disponible
- pré-remplir le tableau PDJ à l'import d'un bon (Excel/PDF)
- conserver la saisie manuelle possible (indispensable pour les scans MAS)

Option A (recommandé si tu utilises git):
1) Dézippe Gestion-Cuisine.zip (ton projet)
2) Ouvre un terminal dans le dossier Gestion-Cuisine/
3) Applique le patch:
   git apply PATCH_app_pdj_import.patch

Option B (sans git):
1) Ouvre PATCH_app_pdj_import.patch
2) Reporte les modifications dans ton app.py (elles sont toutes dans la section Facturation PDJ)

IMPORTANT:
- Ne copie PAS le dossier .git contenu dans les zip que je t'ai donnés auparavant; ça crée des conflits.
- Si src/pdj_facturation.py ne s'importe pas (dépendances OCR manquantes), l'app reste en mode manuel.
