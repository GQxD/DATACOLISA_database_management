"""
Application constants and enumerations.

This module contains all constants used throughout the application,
including default values, option lists, and enumerated types.
"""

from enum import Enum
from typing import List

# Default sheet names
DEFAULT_SOURCE_SHEET: str = "Travail4avril2012"
DEFAULT_TARGET_SHEET: str = "Feuil1 "  # Note: espace à la fin (nom réel dans les fichiers COLISA)

# Default values for import
DEFAULT_ORGANISME: str = "INRAE"
DEFAULT_COUNTRY: str = "France"
DEFAULT_TYPE_ECHANTILLON: str = "EC"

# Application metadata
APP_VERSION: str = "1.3.0"
APP_ORGANISATION: str = "INRAE"
APP_AUTHOR: str = "Quentin Godeaux"


# Enumerations

class DuplicatePolicy(str, Enum):
    """Policy for handling duplicate rows during import."""
    ALERT = "alert"      # Alert user but don't import
    IGNORE = "ignore"    # Skip duplicates silently
    REPLACE = "replace"  # Replace existing row with new data


class ImportStatus(str, Enum):
    """Status of a row in the import process."""
    PRET = "pret"                          # Ready to import
    A_REIMPORTER = "a_reimporter"          # Needs to be re-imported (validation error)
    IMPORTE = "importe"                    # Successfully imported
    IMPORTE_REMPLACE = "importe_remplace"  # Imported, replaced duplicate
    NON_IMPORTE_MANUEL = "non_importe_manuel"  # Manually excluded by user
    IGNORE_DOUBLON = "ignore_doublon"      # Ignored duplicate


# UI Option Lists (from ui_pyside6_poc.py)

TYPE_PECHE_OPTIONS: List[str] = [
    "",
    "LIGNE",
    "FILET",
    "TRAINE",
    "SONDE",
]

CATEGORIE_OPTIONS: List[str] = [
    "",
    "PRO",
    "AMATEUR",
    "SCIENTIFIQUE",
]

OUI_NON_OPTIONS: List[str] = [
    "",
    "OUI",
    "NON",
]

OBSERVATION_OPTIONS: List[str] = [
    "",
    "+",
    "++",
    "+++",
]

# Numeric options for UI dropdowns
NUMERIC_OPTIONS: List[str] = [""] + [str(i) for i in range(1, 11)]
ECAILLES_BRUTES_OPTIONS: List[str] = [""] + [str(i) for i in range(1, 21)]

# Table column names (from ui_pyside6_poc.py)
COLS: List[str] = [
    "selected",
    "include",
    "ref",
    "code_type_echantillon",
    "categorie",
    "type_peche",
    "autre_oss",
    "ecailles_brutes",
    "montees",
    "empreintes",
    "otolithes",
    "observation_disponibilite",
    "source_row",
    "num_individu",
    "date_capture",
    "code_espece",
    "lac_riviere",
    "pays_capture",
    "pecheur",
    "longueur_mm",
    "poids_g",
    "maturite",
    "sexe",
    "age_total",
    "status",
    "errors",
]

# Date formats for parsing
DATE_FORMATS: List[str] = [
    "%d/%m/%Y",
    "%d/%m/%y",
    "%Y-%m-%d",
]

# Excel constants
DEFAULT_CODE_WIDTH: int = 5  # Width for auto-generated codes (e.g., T00001)
MAX_HEADER_SEARCH_ROWS: int = 50  # Maximum rows to search for headers
MAX_HEADER_SEARCH_COLS: int = 300  # Maximum columns to search for headers

# File paths (relative defaults)
DEFAULT_SELECTION_CSV: str = "selection_import.csv"
DEFAULT_OUTPUT_FILE: str = "COLISA_imported.xlsx"
DEFAULT_HISTORY_FILE: str = "import_history.json"
