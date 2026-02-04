PATCH PDJ (remplissage automatique + avoirs)

But: activer le pré-remplissage des quantités et du site à l'import d'un bon PDJ (Excel/PDF),
SANS changer l'UI, ni les autres fonctions.

Fichiers inclus:
- Gestion-Cuisine/app.py
- Gestion-Cuisine/src/pdj_billing.py

⚠️ IMPORTANT (pour éviter les conflits)
Si ton projet a déjà des marqueurs de merge (<<<<<<< ======= >>>>>>>) dans app.py ou src/pdj_billing.py,
remplace ces fichiers en écrasant (copier/coller) avec ceux fournis ici.

Étapes recommandées:
1) Fais une sauvegarde de tes fichiers actuels:
   - app.py -> app.py.bak
   - src/pdj_billing.py -> src/pdj_billing.py.bak
2) Copie les fichiers du patch aux mêmes emplacements dans ton projet.
3) Relance l'app.

Résultat attendu:
- Quand tu importes un bon PDJ (Excel/PDF), la table se pré-remplit (best-effort).
- Le champ 'Site' se pré-remplit si détecté et si le champ était vide.
- Rien n'empêche la correction manuelle (cas MAS).
- Les enregistrements de type 'avoir_qty' sont toujours enregistrés en quantités négatives.
