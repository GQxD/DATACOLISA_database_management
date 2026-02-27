# DATACOLISA - Guide logiciel

## Objectif
DATACOLISA est un outil d'import métier pour transférer des données depuis un fichier Excel source vers un template cible, avec contrôle utilisateur avant écriture.

## Ce que fait le logiciel
- Charge une source Excel `.xls`.
- Filtre les lignes sur une plage de références (`REF`).
- Permet une sélection manuelle des lignes à importer.
- Vérifie les champs métier obligatoires.
- Importe vers le template cible en préservant les colonnes calculées.
- Gère les doublons selon une stratégie (`alert`, `ignore`, `replace`).
- Produit un historique d'import pour suivi et réimport.

## Format de base supporté
Le logiciel fonctionne pour un seul format de base de données:
- Source attendue: fichier `.xls` avec la structure métier DATACOLISA.
- Onglet source attendu: `Travail4avril2012`.
- Cible attendue: `COLISA_template_interne.xlsx`.
- Onglet cible attendu: `Feuil1`.
- Mapping fixe: positions/en-têtes métier définies dans le code (`app/datacolisa_importer.py`) et le contexte (`context master/master_context.md`).

Si la structure source change (colonnes, onglet, position), le logiciel doit être adapté dans le mapping.

## Utilisation (interface PySide6)
1. Lancer l'application `app/ui_pyside6_poc.py` (ou l'exécutable si build déjà fait).
2. Sélectionner le fichier source `.xls`.
3. Vérifier la plage REF à traiter.
4. Charger les lignes et contrôler la sélection dans le tableau.
5. Lancer l'import vers le fichier de sortie.
6. Consulter le fichier d'historique pour les statuts (`importé`, `non_importe_manuel`, `a_reimporter`).

## Fichiers utilisés
- Entrée métier locale: fichier `.xls` utilisateur.
- Template: `COLISA_template_interne.xlsx`.
- Sorties locales: `COLISA_imported*.xlsx`, `import_history*.json`, `selection_import.csv`.

## Données métier et Git
- Le dépôt contient le code, les assets et le template, pas les jeux de données de production.

## Assistance au développement
Le code de ce dépôt a bénéficié du support d'outils d'intelligence artificielle pour l'optimisation syntaxique et la documentation.
