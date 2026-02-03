PDJ Add-on (sans détricoter le reste)

Objectif
- Ajouter une fonctionnalité "Facturation PDJ (économat)".
- Ne pas toucher au graphisme ni aux autres fonctions existantes (seulement un ajout en bas de l'onglet "Facturation mensuelle").

Contenu du paquet
- src/pdj_facturation.py
- src/pdj_billing.py
- app.py.patch (diff unifié: ajouts MINIMAUX dans app.py)
- requirements.txt.patch (ajout de dépendances)

IMPORTANT (pour éviter les conflits)
- Ne merge pas un zip contenant .git.
- Ne copie pas des dossiers .git ou __pycache__.

Installation (recommandée)
1) Copie les 2 fichiers suivants dans TON projet, au même endroit:
   - <ton_projet>/src/pdj_facturation.py
   - <ton_projet>/src/pdj_billing.py

2) Applique les patches.
   Si tu es sous git:
     git apply app.py.patch
     git apply requirements.txt.patch

   Sinon, ouvre app.py et requirements.txt et copie/colle les ajouts visibles dans les patches.

3) Installe les dépendances:
   pip install -r requirements.txt

4) Redémarre l'application.

Détection du site par bon de commande
- Excel: l'app lit le nom du site/établissement dans le titre (ex: "(24 TER)") et le mappe vers "Internat".
- PDF: l'app OCR l'entête du document pour récupérer le site.
- Si non détecté: fallback via le nom de fichier (lautrec/mas -> MAS ; 24T/24 TER/internat -> Internat).

Dépannage
- Si OCR PDF: la lecture manuscrite peut être imparfaite. Préférer les fichiers Excel quand possible.
- Si erreur "poppler": sur Windows, pdf2image peut demander poppler. Dans ce cas, utilise soit les fichiers Excel, soit installe poppler.
