# Gestion-Cuisine — Démarrage rapide

## Windows (sans ligne de commande)
Double-clique sur **Lancer-Gestion-Cuisine.bat**.

Le script :
1) crée un environnement virtuel `.venv`
2) installe les dépendances depuis `requirements.txt`
3) lance l'application Streamlit

> Si Windows affiche que Python est introuvable : installe Python 3 depuis python.org en cochant **Add Python to PATH**, puis relance le `.bat`.

## Mac / Linux
Dans un terminal, depuis le dossier :
```bash
chmod +x lancer.sh
./lancer.sh
```

## Lancement manuel (toutes plateformes)
```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```
