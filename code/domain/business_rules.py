"""
Business rules and domain logic.

This module contains pure business logic without dependencies on
infrastructure or external systems. All rules are stateless and testable.
"""

from __future__ import annotations

import re
from typing import List, Optional, Tuple


class TypePecheDeriver:
    """
    Derives type of fishing and fisher category from source data.

    This encapsulates the business rules for determining fishing type
    (TRAINE, FILET, SONDE, etc.) and fisher category (PRO, AMATEUR,
    SCIENTIFIQUE) from raw source values.
    """

    @staticmethod
    def derive_type_and_categorie(raw_engin: str) -> Tuple[str, str]:
        """
        Derive fishing type and fisher category from raw equipment string.

        Business rules:
        - Single letter codes: T=TRAINE, F=FILET, S=SONDE, _=empty
        - Text containing keywords: "traine", "filet", "sonde", "ligne"
        - Category inferred from context: "pic"=PRO, "sci"=SCIENTIFIQUE
        - TRAINE defaults to AMATEUR if no category specified

        Args:
            raw_engin: Raw equipment/method string from source

        Returns:
            Tuple of (type_peche, categorie)
        """
        if not raw_engin:
            return "", ""

        txt = str(raw_engin).strip()
        if not txt:
            return "", ""

        low = txt.lower()
        up = txt.upper()

        # Derive type_peche
        type_peche = ""
        if low == "t":
            type_peche = "TRAINE"
        elif low == "f":
            type_peche = "FILET"
        elif low == "s":
            type_peche = "SONDE"
        elif low == "_":
            type_peche = ""
        elif "traine" in low:
            type_peche = "TRAINE"
        elif "filet" in low:
            type_peche = "FILET"
        elif "sonde" in low:
            type_peche = "SONDE"
        elif "ligne" in low:
            type_peche = "LIGNE"
        else:
            type_peche = up  # Keep uppercase if not recognized

        # Derive categorie
        categorie = ""
        if "pic" in low:
            categorie = "PRO"
        elif "sci" in low:
            categorie = "SCIENTIFIQUE"
        elif type_peche == "TRAINE" or "traine" in low:
            categorie = "AMATEUR"
        elif "pro" in low:
            categorie = "PRO"
        elif "amat" in low:
            categorie = "AMATEUR"

        return type_peche, categorie


class PaysDeriver:
    """
    Derives country from context code.

    Extracts country from codes ending with FR (France) or CH (Suisse).
    """

    @staticmethod
    def derive_country(raw_contexte: str) -> str:
        """
        Extract country from context string.

        Args:
            raw_contexte: Context code (e.g., "SOMETHING_FR", "OTHER_CH")

        Returns:
            "France", "Suisse", or empty string
        """
        if not raw_contexte:
            return ""

        txt = str(raw_contexte).strip().upper()
        if not txt:
            return ""

        # Look for FR or CH at end of string
        m = re.search(r"(CH|FR)\s*$", txt)
        if not m:
            return ""

        code = m.group(1)
        if code == "CH":
            return "Suisse"
        elif code == "FR":
            return "France"

        return ""


class ValidationRules:
    """
    Validation rules for source rows.

    Contains business validation logic to ensure data quality
    before import.
    """

    @staticmethod
    def validate_source_row(
        row,
        default_type_echantillon: str
    ) -> List[str]:
        """
        Validate a source row against business rules.

        Required fields:
        - date_capture (T11)
        - code_espece (T6)
        - num_individu (T19)
        - code_type_echantillon (T4)

        Args:
            row: SourceRow object to validate
            default_type_echantillon: Default type if not specified

        Returns:
            List of error messages (empty if valid)
        """
        missing = []

        # Check required fields
        if not ValidationRules._normalize(row.date_capture):
            missing.append("Date de capture (entrée manquante pour T11)")

        if not ValidationRules._normalize(row.code_espece):
            missing.append("Taxon / Code espèce (entrée manquante pour T6)")

        if not ValidationRules._normalize(row.num_individu):
            missing.append("Numéro individu (entrée manquante pour T19)")

        if not ValidationRules._normalize(default_type_echantillon):
            missing.append("Code type échantillon (T4)")

        # Calculated fields require source data
        if not ValidationRules._normalize(row.num_individu):
            missing.append("Entrée requise pour génération Code échantillon")

        if not ValidationRules._normalize(row.code_espece):
            missing.append("Entrée requise pour génération ligne LE02")

        return missing

    @staticmethod
    def _normalize(value) -> str:
        """Normalize value to string."""
        if value is None:
            return ""
        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return format(value, "f").rstrip("0").rstrip(".")
        return str(value).strip()


class DataTransformations:
    """
    Data transformation rules.

    Contains business logic for transforming data during import.
    """

    @staticmethod
    def normalize_placeholder(value: str, placeholder: str = "_") -> str:
        """
        Replace placeholder value with empty string.

        Args:
            value: Value to check
            placeholder: Placeholder string (default "_")

        Returns:
            Empty string if value equals placeholder, otherwise value
        """
        if not value:
            return ""
        if str(value).strip() == placeholder:
            return ""
        return str(value).strip()

    @staticmethod
    def normalize_poids_sexe(
        poids_g: Optional[str],
        sexe: Optional[str]
    ) -> Tuple[str, str]:
        """
        Normalize weight and sex fields.

        Business rule: "_" placeholder becomes empty or "N"

        Args:
            poids_g: Weight in grams
            sexe: Sex code

        Returns:
            Tuple of (normalized_poids, normalized_sexe)
        """
        poids_norm = DataTransformations.normalize_placeholder(poids_g or "")
        sexe_norm = DataTransformations.normalize_placeholder(sexe or "")

        # If empty after normalization, use "N"
        if not poids_norm:
            poids_norm = "N"
        if not sexe_norm:
            sexe_norm = "N"

        return poids_norm, sexe_norm

    @staticmethod
    def determine_autre_oss(otolithes_value: Optional[str]) -> str:
        """
        Determine "autre ossements" based on otolithes value.

        Business rule: If otolithes present and not "0", autre_oss = "OUI"

        Args:
            otolithes_value: Otolithes count or indicator

        Returns:
            "OUI" or "NON"
        """
        if not otolithes_value:
            return "NON"

        val = str(otolithes_value).strip()
        if val and val != "0":
            return "OUI"

        return "NON"

    @staticmethod
    def normalize_lac(lac_riviere: Optional[str]) -> str:
        """
        Normalize lake/river name.

        Business rule: "LEMAN" becomes "L"

        Args:
            lac_riviere: Lake or river name

        Returns:
            Normalized name
        """
        if not lac_riviere:
            return ""

        val = str(lac_riviere).strip().upper()
        if val == "LEMAN":
            return "L"

        return val


class ReferenceCodeGenerator:
    """
    Generates sequential reference codes.

    Not yet implemented - will be used in Phase 4.
    """

    @staticmethod
    def init_sequence_from_workbook(ws, header_row: int, code_col: Optional[int]) -> dict:
        """
        Initialize code sequence from existing workbook.

        Scans workbook to find highest existing code number.

        Args:
            ws: Worksheet object
            header_row: Header row index
            code_col: Code column index

        Returns:
            Dict with 'prefix', 'num', 'width' keys
        """
        if not code_col:
            return {"prefix": "T", "num": 0, "width": 5}

        pattern = re.compile(r"^([A-Za-z]*)(\d+)$")
        best_prefix = "T"
        best_num = 0
        best_width = 0

        for r in range(header_row + 1, ws.max_row + 1):
            raw = str(ws.cell(r, code_col).value or "").strip()
            if not raw:
                continue

            m = pattern.match(raw)
            if not m:
                continue

            prefix, num_s = m.group(1), m.group(2)
            num = int(num_s)
            width = len(num_s)

            if num > best_num:
                best_num = num
                best_prefix = prefix or best_prefix
                best_width = width

        if best_width == 0:
            best_width = 5

        return {
            "prefix": best_prefix or "T",
            "num": best_num,
            "width": best_width
        }
