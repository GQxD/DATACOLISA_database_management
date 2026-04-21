"""
Column mappings for source and target Excel files.

This module contains all hard-coded column positions and header names
used by the source reader and the integrated COLISA base.
"""

from typing import Dict, List

# Mapping V1: source positional columns (1-based indexes in source data row)
# Based on current project context.
SOURCE_POSITIONS: Dict[str, int] = {
    "num_individu_primary": 2,   # col 2  = référence (XT301, XT302...)
    "pecheur": 4,                 # col 4  = pêcheur (FIPALTL09thonon...)
    "contexte": 5,                # col 5  = contexte (PATLL09FR...)
    "code_espece": 6,             # col 6  = espèce (TL, OBL...)
    "date_capture": 7,            # col 7  = date capture
    "lac_riviere": 10,            # col 10 = lieu (L=Léman, R=Rivière)
    "engin_source": 14,           # col 14 = engin (T=traîne, F=filet...)
    "longueur_mm": 16,            # col 16 = LT longueur totale (mm)
    "poids_g": 17,                # col 17 = PoidsTOT (g)
    "maturite": 20,               # col 20 = maturité sexuelle
    "sexe": 21,                   # col 21 = sexe
    "num_individu_fallback": 26,  # col 26 = référence secondaire
    "age_total": 36,              # col 36 = COH (cohorte/âge)
}

# Target Excel header names with possible variations.
TARGET_HEADERS: Dict[str, List[str]] = {
    "code_unite_gestionnaire": ["Code unit? gestionnaire", "Code unite gestionnaire", "Code unité gestionnaire"],
    "site_atelier": ["Site Atelier", "Site atelier"],
    "numero_correspondant": ["Numero du correspondant", "Num?ro du correspondant", "Numéro du correspondant"],
    "code_type_echantillon": ["Code type echantillon", "Code type échantillon"],
    "code_echantillon": ["Code echantillon", "Code échantillon"],
    "code_espece": ["Code espece", "Code espèce"],
    "organisme": ["Organisme preleveur", "Organisme préleveur"],
    "pays": ["Pays capture", "Pays capture "],
    "date_capture": ["Date capture", "Date de capture", "Date de capture (JJ/MM/AAAA)"],
    "lac_riviere": ["Lac/riviere", "Lac/rivière"],
    "categorie": ["Categorie pecheur", "Catégorie pêcheur "],
    "type_peche": ["Type peche/engin", "Type pêche/engin "],
    "num_individu": ["Numero individu", "Numero individu (numero de capture)", "Numéro individu (numéro de capture)"],
    "longueur_mm": ["Longueur totale (mm)"],
    "poids_g": ["Poids (g)"],
    "maturite": ["Code maturite sexuelle", "Code maturité sexuelle"],
    "sexe": ["Code sexe"],
    "age_total": ["Age total"],
    "ecailles_brutes": ["Ecailles brutes", "Écailles brutes"],
    "montees": ["Montees", "Montées"],
    "empreintes": ["Empreintes", "Empreinte"],
    "otolithes": ["Otolithes"],
    "autre_oss": [
        "Autre ?chantillon osseuses collect?e sur l'individu OUI/NON",
        "Autre echantillon osseuses collectee sur l'individu OUI/NON",
        "Autre échantillon osseuses collectée sur l'individu OUI/NON",
    ],
    "observation_disponibilite": ["Observation disponibilit?", "Observation disponibilite", "Observation disponibilité"],
    "numero_identification": ["CODE IDENTIFICATION", "Numero identification", "Numéro identification", "Numero d'identification", "Numéro d'identification"],
    "sous_espece": ["Sous-espece", "Sous-espèce", "Sous-esp?ce", "Sous-esp\u00e8ce "],
    "nom_operateur": ["Nom de l'operateur", "Nom de l'op\u00e9rateur", "Nom de l'op?rateur"],
    "lieu_capture": ["Lieu de capture / debarquement", "Lieu de capture"],
    "maille_mm": ["Maille (mm)"],
    "code_stade": ["Code stade"],
    "presence_otolithe_gauche": ["Presence de l'otolithe gauche (0 si non, 1 si oui)", "Pr\u00e9sence de l'otolithe gauche (0 si non, 1 si oui)"],
    "presence_otolithe_droite": ["Presence de l'otolithe droite (0 si non, 1 si oui)", "Pr\u00e9sence de l'otolithe droite (0 si non, 1 si oui)"],
    "nb_opercules": ["Nombre d'opercules en etat", "Nombre d' opercules en \u00e9tat"],
    "information_stockage": ["Information stockage", "Information stockage "],
    "age_riviere": ["Age riviere", "Age rivi\u00e8re", "Age rivi?re", "Age rivi\u00e8re "],
    "age_lac": ["Age lac"],
    "nb_fraie": ["Nombre de fraie"],
    "ecailles_regenerees": ["Ecailles regenerees ? (0 si non, 1 si oui)", "Ecailles regen\u00e9r\u00e9es ? (0 si non, 1 si oui)"],
    "observations": ["Observations"],
}

TYPE_SHEET_CANDIDATES: List[str] = [
    "Type echantillon",
    "Type ?chantillon",
    "Type_echantillon",
    "Types d'échantillon",
]
