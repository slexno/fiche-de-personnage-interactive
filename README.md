# fiche-de-personnage-interactive

Application Python + interface HTML pour gérer une fiche de personnage inspirée de Donjons & Dragons.

## Lancer l'application

```bash
python app.py
```

Puis ouvrir `http://localhost:8000` (ou `http://localhost:8080` si le port 8000 est déjà utilisé).

## Fonctionnalités implémentées

- Onglets **Statistiques**, **Inventaire**, **Magasin**.
- Modification des stats avec plafond à 20, recalcul du bonus D&D et de l'Armor Class.
- Gestion du sac à dos et du coffre (ajout manuel, tri, transfert, détails d'objets).
- Assignation d'objets en armes/équipements, avec équipement/déséquipement (limites 4 armes, 3 équipements).
- Magasin multi-sous-onglets basé sur toutes les feuilles de `magasin.xlsx` (achat et vente via crédits).
- Synchronisation des changements vers les fichiers Excel (`caracteristique.xlsx`, `inventaire.xlsx`, `magasin.xlsx`).


> Si vous tapez seulement `http://localhost` sans port, le navigateur essaie le port 80 et vous aurez "connexion refusée".
