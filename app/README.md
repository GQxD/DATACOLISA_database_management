# DATACOLISA Importer (MVP)

CLI Python pour le flux Excel -> Excel décrit dans `context master/master_context.md`.

## Fonctions couvertes
- Lecture source `.xls` sur l'onglet `Travail4avril2012`.
- Filtre par plage `REF` (ex: `CA961` à `CA989`).
- Génération d'un CSV de sélection manuelle (`include=1/0`).
- Validation des champs essentiels.
- Import partiel vers la cible `.xlsx` onglet `Feuil1`.
- Préservation des colonnes calculées (pas d'écriture sur `Code échantillon` / `LE02...`).
- Gestion des doublons par clé (`Numéro individu` + `Code type échantillon`) avec mode `alert|ignore|replace`.
- Historique des statuts: `importe`, `non_importe_manuel`, `a_reimporter`.

## Installation
```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
```

## Usage
```bash
python datacolisa_importer.py extract \
  --source "PacFinalTL14novembrel2012.xls" \
  --start-ref CA961 \
  --end-ref CA989 \
  --out-csv selection_import.csv \
  --default-type-echantillon "EC MONTEE"
```

Éditer `selection_import.csv`:
- `include=1` importe la ligne.
- `include=0` conserve la ligne en `non_importe_manuel`.

```bash
python datacolisa_importer.py import \
  --selection-csv selection_import.csv \
  --target "COLISA 89463-.xlsx" \
  --out-target COLISA_imported.xlsx \
  --history import_history.json \
  --on-duplicate alert
```

Lister les lignes à réimporter:
```bash
python datacolisa_importer.py reimport --history import_history.json
```

Filtrer sur des références:
```bash
python datacolisa_importer.py reimport --history import_history.json --refs CA971 CA972
```

## Interface graphique (recommandé)
Lancer l'interface Streamlit:

```bash
streamlit run ui_streamlit.py
```

Fonctions UI:
- Chargement de plage REF avec résumé des codes manquants.
- Tableau éditable avec cases à cocher par ligne.
- Actions en masse (cocher/décocher, type échantillon).
- Export CSV de sélection.
- Import direct vers le fichier cible.
- Visualisation de l'historique de traitement.

## Notes
- Le mapping actuel suit les positions source fixées dans le contexte (C2, C6, C7, etc.).
- Les colonnes non mappées restent vides.
- Les menus déroulants Excel ciblés par colonne ne sont pas encore injectés dans ce MVP CLI.
