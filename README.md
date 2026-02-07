# Gestion-Cuisine — Dev local (pour modifier le Bon de Commande)

## Le fichier à modifier (Bon de Commande)
➡️ **`src/bon_commande.py`**

C’est le SEUL fichier à toucher si tu veux changer :
- les colonnes du bon de commande,
- le regroupement (fournisseur / produit / repas / typologie),
- les règles de quantités / unités / prix cibles,
- l’ordre de tri, etc.

L’application affiche dans la **sidebar** :
- le chemin du fichier `src/bon_commande.py` réellement chargé,
- son hash SHA1,
- sa date de dernière modification,
pour que tu sois sûr que tu modifies bien le bon fichier.

## Lancer l’application (recommandé)
### Windows
Double-clique **`Lancer-Gestion-Cuisine.bat`**

### macOS / Linux
```bash
./lancer.sh
```

## Si tu ne vois pas tes modifications
1. **Vérifie la sidebar** : le chemin affiché doit pointer vers ton dossier de travail.
2. Ferme totalement l’app (CTRL+C dans le terminal), puis relance.
3. Assure-toi de modifier **`src/bon_commande.py`** (et pas un autre dossier/copier-coller du projet).
