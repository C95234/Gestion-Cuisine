# Gestion-Cuisine

Application Streamlit de gestion de cuisine centrale : production, bon de commande, bons de livraison, facturation mensuelle, tableaux allergènes.

En ligne : https://gestioncuisine.streamlit.app/

## Lancer l'application en local

### Windows
Double-clique **`Lancer-Gestion-Cuisine.bat`**

### macOS / Linux
```bash
./lancer.sh
```

Les deux scripts créent un environnement virtuel `.venv`, installent les dépendances (`requirements.txt`) puis lancent Streamlit.

## Structure du code

- `app.py` — point d'entrée Streamlit (mise en page, onglets, appels aux modules `src/`)
- `src/processor.py` — parsing du planning/menu, production, bons de livraison
- `src/bon_commande.py` — construction du bon de commande (colonnes, regroupement, quantités/unités/prix cibles)
- `src/order_forms.py` — export des bons par fournisseur (Excel/PDF)
- `src/billing.py` — mémorisation des semaines et facturation mensuelle
- `src/config_store.py` — coefficients / unités / fournisseurs (listes mémorisées en JSON)
- `src/allergens/` — génération et apprentissage des tableaux allergènes
- `data/` — fichiers persistés par l'app (référentiel maître allergènes, config)
- `templates/allergen/` — gabarit Excel des tableaux allergènes

Pour changer les colonnes, le regroupement ou les règles de quantités/unités/prix du bon de commande, le seul fichier à modifier est `src/bon_commande.py`.

## Si tu ne vois pas tes modifications

1. Ferme totalement l'app (CTRL+C dans le terminal), puis relance.
2. Vérifie que tu modifies bien le dossier de travail que le script `Lancer-Gestion-Cuisine.bat` / `lancer.sh` utilise (et pas une autre copie du projet).
